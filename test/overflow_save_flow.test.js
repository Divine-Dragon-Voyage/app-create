const test = require('node:test');
const assert = require('node:assert/strict');
const {
    OVERFLOW_MORE_OPTIONS_SELECTORS,
    OVERFLOW_SAVE_MENU_SELECTORS
} = require('../overflow_save_flow');

test('overflow save fallback can locate more options button', () => {
    assert.ok(OVERFLOW_MORE_OPTIONS_SELECTORS.includes('button[aria-label="More options"]'));
    assert.ok(OVERFLOW_MORE_OPTIONS_SELECTORS.includes('button[debug-id="icon-button"][aria-haspopup="listbox"]'));
    assert.ok(OVERFLOW_MORE_OPTIONS_SELECTORS.includes('[role="button"][aria-label="More options"]'));
});

test('overflow save fallback can locate save menu item', () => {
    assert.ok(OVERFLOW_SAVE_MENU_SELECTORS.includes('[role="menuitem"][aria-label="Save"]'));
    assert.ok(OVERFLOW_SAVE_MENU_SELECTORS.includes('material-select-item[aria-label="Save"]'));
    assert.ok(OVERFLOW_SAVE_MENU_SELECTORS.includes('[role="menuitem"]:has-text("Save")'));
    assert.ok(OVERFLOW_SAVE_MENU_SELECTORS.includes('material-select-item:has-text("Save")'));
});
