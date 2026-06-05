const CREATE_APP_LANGUAGE_DROPDOWN_SELECTOR = 'material-dropdown-select[debug-id="language-dropdown"]';

const CREATE_APP_LANGUAGE_BUTTON_SELECTORS = [
    'material-dropdown-select[debug-id="language-dropdown"] [role="button"][aria-label*="Default language"]',
    'material-dropdown-select[debug-id="language-dropdown"] dropdown-button [role="button"]',
    'material-dropdown-select[debug-id="language-dropdown"] .button'
];

const CREATE_APP_EN_US_OPTION_SELECTORS = [
    'material-select-item[role="option"]:has-text("English (United States)")',
    '[role="option"]:has-text("English (United States)")',
    'material-select-item:has-text("English (United States)")',
    '[role="option"]:has-text("en-US")',
    'material-select-item:has-text("en-US")'
];

function normalizeDash(value) {
    return String(value || '').replace(/[\u2010-\u2015]/g, '-');
}

function isEnUsLanguageText(text) {
    const normalized = normalizeDash(text).replace(/\s+/g, ' ').trim();
    return /English\s*\(United States\)\s*-\s*en-US/i.test(normalized);
}

module.exports = {
    CREATE_APP_EN_US_OPTION_SELECTORS,
    CREATE_APP_LANGUAGE_BUTTON_SELECTORS,
    CREATE_APP_LANGUAGE_DROPDOWN_SELECTOR,
    isEnUsLanguageText,
    normalizeDash
};
