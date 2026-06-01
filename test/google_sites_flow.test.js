const test = require('node:test');
const assert = require('node:assert/strict');
const {
    GOOGLE_SITES_DELETE_HEADER_SELECTORS
} = require('../google_sites_flow');

test('google sites delete header fallback includes current aria label and jsname', () => {
    assert.ok(GOOGLE_SITES_DELETE_HEADER_SELECTORS.includes('button[aria-label="Delete the header"]'));
    assert.ok(GOOGLE_SITES_DELETE_HEADER_SELECTORS.includes('button[jsname="D0Nom"]'));
    assert.ok(GOOGLE_SITES_DELETE_HEADER_SELECTORS.includes('button[aria-label="Delete header"]'));
});
