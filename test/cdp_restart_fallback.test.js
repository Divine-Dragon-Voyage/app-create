const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const assert = require('node:assert/strict');

const createAppSource = fs.readFileSync(path.join(__dirname, '..', 'create_app.js'), 'utf8');

test('CDP connection fallback restarts debug browser before final failure', () => {
    assert.match(createAppSource, /async function restartChromeForCdpFallback\(\)/);
    assert.match(createAppSource, /await restartChromeForCdpFallback\(\)/);
    assert.match(createAppSource, /Retrying Chrome CDP connection after browser restart/);
    assert.match(createAppSource, /--remote-debugging-port=9222/);
    assert.match(createAppSource, /--user-data-dir="\$\{getCdpBrowserUserDataDir\(\)\}"/);
    assert.match(createAppSource, /await waitForCdpEndpointHealth\(getPrimaryCdpJsonVersionUrl\(\)/);
});

test('CDP restart fallback targets only Chrome or Edge debug browser processes', () => {
    assert.match(createAppSource, /wmic/);
    assert.match(createAppSource, /remote-debugging-port=9222/);
    assert.match(createAppSource, /processid/);
    assert.match(createAppSource, /taskkill/);
});
