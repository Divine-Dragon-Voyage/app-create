const assert = require('assert');
const test = require('node:test');

const {
    AD_ID_ACK_CHECKBOX_SELECTOR,
    CREATE_APP_EN_US_OPTION_SELECTORS,
    CREATE_APP_LANGUAGE_DROPDOWN_SELECTOR,
    PRODUCTION_EDIT_RELEASE_SELECTORS,
    PRODUCTION_RELEASES_TAB_SELECTORS,
    RELEASE_UPDATE_DECLARATION_SELECTORS,
    hasBlockingReleaseErrorText,
    isAdIdReleasePermissionError,
    isEnUsLanguageText
} = require('../release_error_flow');

test('detects release review errors caused by missing AD_ID permission', () => {
    const text = [
        'Your advertising ID declaration in Play Console says that your app uses advertising ID.',
        "A manifest file in one of your active artifacts doesn't include the",
        'com.google.android.gms.permission.AD_ID permission.'
    ].join(' ');

    assert.equal(isAdIdReleasePermissionError(text), true);
});

test('does not treat generic release review problems as AD_ID permission errors', () => {
    assert.equal(isAdIdReleasePermissionError('We found some problems with your release'), false);
});

test('distinguishes blocking release errors from warnings', () => {
    assert.equal(hasBlockingReleaseErrorText('1 Error To save, fix errors'), true);
    assert.equal(hasBlockingReleaseErrorText('2 Errors'), true);
    assert.equal(hasBlockingReleaseErrorText('1 Warning We found some problems with your release'), false);
});

test('recognizes English United States language labels', () => {
    assert.equal(isEnUsLanguageText('English (United States) - en-US'), true);
    assert.equal(isEnUsLanguageText('Default language: English (United States) \u2013 en-US'), true);
    assert.equal(isEnUsLanguageText('English (United Kingdom) \u2013 en-GB'), false);
});

test('release error flow selectors cover language and AD_ID recovery controls', () => {
    assert.equal(CREATE_APP_LANGUAGE_DROPDOWN_SELECTOR, 'material-dropdown-select[debug-id="language-dropdown"]');
    assert.ok(CREATE_APP_EN_US_OPTION_SELECTORS.some(selector => /en-US|English \(United States\)/.test(selector)));
    assert.ok(RELEASE_UPDATE_DECLARATION_SELECTORS.some(selector => selector.includes('Update declaration')));
    assert.equal(AD_ID_ACK_CHECKBOX_SELECTOR, 'material-checkbox[debug-id="ack-checkbox"]');
    assert.ok(PRODUCTION_RELEASES_TAB_SELECTORS.some(selector => selector.includes('aria-label="Releases"')));
    assert.ok(PRODUCTION_EDIT_RELEASE_SELECTORS.some(selector => selector.includes('edit-draft-release-button')));
});
