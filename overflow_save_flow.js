const OVERFLOW_MORE_OPTIONS_SELECTORS = [
    'button[aria-label="More options"]',
    'button[debug-id="icon-button"][aria-haspopup="listbox"]',
    '[role="button"][aria-label="More options"]'
];

const OVERFLOW_SAVE_MENU_SELECTORS = [
    '[role="menuitem"][aria-label="Save"]',
    'material-select-item[aria-label="Save"]',
    '[role="menuitem"]:has-text("Save")',
    'material-select-item:has-text("Save")'
];

module.exports = {
    OVERFLOW_MORE_OPTIONS_SELECTORS,
    OVERFLOW_SAVE_MENU_SELECTORS
};
