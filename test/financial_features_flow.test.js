const test = require('node:test');
const assert = require('node:assert/strict');
const {
    FINANCIAL_FEATURES_MORE_OPTIONS_SELECTORS,
    FINANCIAL_FEATURES_SAVE_MENU_SELECTORS
} = require('../financial_features_flow');

test('financial features fallback can locate more options button', () => {
    assert.ok(FINANCIAL_FEATURES_MORE_OPTIONS_SELECTORS.includes('button[aria-label="More options"]'));
    assert.ok(FINANCIAL_FEATURES_MORE_OPTIONS_SELECTORS.includes('button[debug-id="icon-button"][aria-haspopup="listbox"]'));
});

test('financial features fallback can locate save menu item', () => {
    assert.ok(FINANCIAL_FEATURES_SAVE_MENU_SELECTORS.includes('[role="menuitem"][aria-label="Save"]'));
    assert.ok(FINANCIAL_FEATURES_SAVE_MENU_SELECTORS.includes('material-select-item[aria-label="Save"]'));
});
