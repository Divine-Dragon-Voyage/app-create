const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const assert = require('node:assert/strict');

const script = fs.readFileSync(path.join(__dirname, '..', 'create_app.js'), 'utf8');

test('create flow keeps copied seven-items and country fallbacks only', () => {
    assert.match(script, /ensureCreateAppDefaultLanguageEnUs/);
    assert.match(script, /Executing declaration 2\/7: Sign in details/);
    assert.match(script, /clickFinancialFeaturesSave/);
    assert.match(script, /Countries \/ regions tab not ready, using direct URL fallback/);
    assert.match(script, /Production page loaded but Countries\/regions controls are not visible yet; continuing with fallback navigation/);
});

test('seven-items branch includes copied post-country ten-item flow', () => {
    assert.match(script, /runPrivacyPolicyStep/);
    assert.match(script, /runContentRatingsStep/);
    assert.match(script, /runDataSafetyStep/);
    assert.match(script, /runStoreSettingsStep/);
    assert.match(script, /runStoreListingStep/);
    assert.match(script, /runReleaseStep/);
    assert.match(script, /runReviewScreenshotUploadStep/);
});
