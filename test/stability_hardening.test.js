const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const assert = require('node:assert/strict');

const rootDir = path.join(__dirname, '..');
const createAppSource = fs.readFileSync(path.join(rootDir, 'create_app.js'), 'utf8');
const bootstrapSource = fs.readFileSync(path.join(rootDir, 'bootstrap_windows.ps1'), 'utf8');
const userReadmeSource = fs.readFileSync(path.join(rootDir, 'user_ops', 'README.md'), 'utf8');
const packageReleasePs1Source = fs.readFileSync(path.join(rootDir, 'package_release.ps1'), 'utf8');
const packageReleaseShSource = fs.readFileSync(path.join(rootDir, 'package_release.sh'), 'utf8');

test('runtime startup logs version from VERSION file without requiring git metadata', () => {
    assert.match(createAppSource, /const VERSION_FILE_NAME\s*=\s*'VERSION'/);
    assert.match(createAppSource, /function loadRuntimeVersionInfo\(\)/);
    assert.match(createAppSource, /console\.log\(`\[VERSION\]/);
    assert.doesNotMatch(createAppSource, /execFileSync\(\s*['"]git['"]/);
});

test('Chrome CDP profile is derived from contact email and guarded by a stale-safe run lock', () => {
    assert.match(createAppSource, /function getCdpBrowserProfileKey\(\)/);
    assert.match(createAppSource, /function sanitizeProfileKey/);
    assert.match(createAppSource, /path\.join\(baseDir,\s*profileKey\)/);
    assert.match(createAppSource, /function acquireRunLock\(\)/);
    assert.match(createAppSource, /function releaseRunLock/);
    assert.match(createAppSource, /stale/i);
    assert.match(createAppSource, /APP_CREATE_CONTACT_EMAIL/);
});

test('CDP restart waits for stale browser processes and profile lock release before relaunch', () => {
    assert.match(createAppSource, /async function waitForDebugBrowserProcessesToExit/);
    assert.match(createAppSource, /async function waitForChromeProfileReady/);
    assert.match(createAppSource, /SingletonLock|SingletonCookie|SingletonSocket/);
    assert.match(createAppSource, /await waitForDebugBrowserProcessesToExit/);
    assert.match(createAppSource, /await waitForChromeProfileReady/);
});

test('Google Sites editor wait captures diagnostic state and screenshot on timeout', () => {
    assert.match(createAppSource, /async function captureGoogleSitesDiagnostics/);
    assert.match(createAppSource, /classifyGoogleSitesPage/);
    assert.match(createAppSource, /google-sites-/);
    assert.match(createAppSource, /sitesPage\.screenshot/);
    assert.match(createAppSource, /Google Sites editor did not become ready/);
});

test('batch loop performs controlled CDP health checks between packages only', () => {
    assert.match(createAppSource, /const CDP_HEALTH_CHECK_INTERVAL_ROWS\s*=\s*3/);
    assert.match(createAppSource, /async function runInterIterationCdpHealthCheck/);
    assert.match(createAppSource, /if \(shouldRunInterIterationCdpHealthCheck\(i \+ 1, selectedTasks\.length\)\)/);
    assert.match(createAppSource, /after completed row/);
});

test('10-item branch still stops before store listing release and review upload calls', () => {
    assert.match(createAppSource, /Executing post-country item 1\/3: Privacy policy/);
    assert.doesNotMatch(createAppSource, /await runStoreSettingsStep\(\);/);
    assert.doesNotMatch(createAppSource, /await runStoreListingStep\(\);/);
    assert.doesNotMatch(createAppSource, /await runReleaseStep\(\);/);
    assert.doesNotMatch(createAppSource, /await runReviewScreenshotUploadStep\(\);/);
});

test('Windows bootstrap passes account-scoped CDP profile directory from contact email', () => {
    assert.match(bootstrapSource, /APP_CREATE_BROWSER_USER_DATA_DIR/);
    assert.match(bootstrapSource, /APP_CREATE_CONTACT_EMAIL/);
    assert.match(bootstrapSource, /Get-SafeProfileName/);
    assert.match(bootstrapSource, /chrome-cdp-app-create/);
});

test('user docs describe status tracking without promising breakpoint resume', () => {
    assert.doesNotMatch(userReadmeSource, /断点续跑/);
    assert.match(userReadmeSource, /运行状态|排查状态/);
});

test('release packages write VERSION metadata into staging archives', () => {
    assert.match(packageReleasePs1Source, /VERSION/);
    assert.match(packageReleasePs1Source, /git\s+rev-parse/);
    assert.match(packageReleasePs1Source, /Set-Content[^\n]+VERSION/);
    assert.match(packageReleaseShSource, /VERSION/);
    assert.match(packageReleaseShSource, /git rev-parse/);
    assert.match(packageReleaseShSource, />\s*"\$STAGING_DIR\/VERSION"/);
});
