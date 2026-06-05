const assert = require('assert');
const test = require('node:test');

const {
    CREATE_APP_EN_US_OPTION_SELECTORS,
    CREATE_APP_LANGUAGE_DROPDOWN_SELECTOR,
    isEnUsLanguageText
} = require('../release_error_flow');

test('recognizes English United States language labels', () => {
    assert.equal(isEnUsLanguageText('English (United States) - en-US'), true);
    assert.equal(isEnUsLanguageText('Default language: English (United States) \u2013 en-US'), true);
    assert.equal(isEnUsLanguageText('English (United Kingdom) \u2013 en-GB'), false);
});

test('release selectors cover create app default language controls', () => {
    assert.equal(CREATE_APP_LANGUAGE_DROPDOWN_SELECTOR, 'material-dropdown-select[debug-id="language-dropdown"]');
    assert.ok(CREATE_APP_EN_US_OPTION_SELECTORS.some(selector => /en-US|English \(United States\)/.test(selector)));
});
