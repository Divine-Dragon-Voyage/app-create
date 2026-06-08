const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const assert = require('node:assert/strict');

const script = fs.readFileSync(path.join(__dirname, '..', 'create_app.js'), 'utf8');

test('create flow keeps copied seven-items and country fallbacks only', () => {
    assert.match(script, /ensureCreateAppDefaultLanguageEnUs/);
    assert.match(script, /Executing declaration 2\/7: Sign-in details/);
    assert.match(script, /clickFinancialFeaturesSave/);
    assert.match(script, /Countries \/ regions tab not ready, using direct URL fallback/);
    assert.match(script, /Production page loaded but Countries\/regions controls are not visible yet; continuing with fallback navigation/);
});

test('seven-items branch stops after country plus three post-country items', () => {
    assert.match(
        script,
        /await runPrivacyPolicyStep\(\);\s*await runContentRatingsStep\(\);\s*await runDataSafetyStep\(\);/
    );
    assert.match(script, /Executing post-country item 1\/3: Privacy policy/);
    assert.match(script, /Executing post-country item 2\/3: Content ratings/);
    assert.match(script, /Executing post-country item 3\/3: Data safety/);
    assert.doesNotMatch(script, /await runStoreSettingsStep\(\);/);
    assert.doesNotMatch(script, /await runStoreListingStep\(\);/);
    assert.doesNotMatch(script, /await runReleaseStep\(\);/);
    assert.doesNotMatch(script, /await runReviewScreenshotUploadStep\(\);/);
});
