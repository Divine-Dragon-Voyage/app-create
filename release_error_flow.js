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

const RELEASE_UPDATE_DECLARATION_SELECTORS = [
    'button[type="submit"]:has-text("Update declaration")',
    'button:has-text("Update declaration")',
    '[role="button"]:has-text("Update declaration")'
];

const AD_ID_ACK_CHECKBOX_SELECTOR = 'material-checkbox[debug-id="ack-checkbox"]';

const PRODUCTION_RELEASES_TAB_SELECTORS = [
    'tab-button[aria-label="Releases"]',
    '[role="tab"][aria-label="Releases"]',
    '[role="tab"]:has-text("Releases")',
    'tab-button:has-text("Releases")'
];

const PRODUCTION_EDIT_RELEASE_SELECTORS = [
    'button[debug-id="edit-draft-release-button"]',
    'button:has-text("Edit release")',
    '[role="button"]:has-text("Edit release")'
];

const AD_ID_RELEASE_PERMISSION_ERROR_REGEX =
    /advertising\s+ID\s+declaration[\s\S]*com\.google\.android\.gms\.permission\.AD_ID|com\.google\.android\.gms\.permission\.AD_ID[\s\S]*advertising\s+ID/i;

const BLOCKING_RELEASE_ERROR_REGEX = /\b[1-9]\d*\s+Errors?\b|To\s+save,\s*fix\s+errors/i;

function normalizeDash(value) {
    return String(value || '').replace(/[\u2010-\u2015]/g, '-');
}

function isEnUsLanguageText(text) {
    const normalized = normalizeDash(text).replace(/\s+/g, ' ').trim();
    return /English\s*\(United States\)\s*-\s*en-US/i.test(normalized);
}

function isAdIdReleasePermissionError(text) {
    return AD_ID_RELEASE_PERMISSION_ERROR_REGEX.test(String(text || ''));
}

function hasBlockingReleaseErrorText(text) {
    return BLOCKING_RELEASE_ERROR_REGEX.test(String(text || ''));
}

module.exports = {
    AD_ID_ACK_CHECKBOX_SELECTOR,
    CREATE_APP_EN_US_OPTION_SELECTORS,
    CREATE_APP_LANGUAGE_BUTTON_SELECTORS,
    CREATE_APP_LANGUAGE_DROPDOWN_SELECTOR,
    PRODUCTION_EDIT_RELEASE_SELECTORS,
    PRODUCTION_RELEASES_TAB_SELECTORS,
    RELEASE_UPDATE_DECLARATION_SELECTORS,
    hasBlockingReleaseErrorText,
    isAdIdReleasePermissionError,
    isEnUsLanguageText,
    normalizeDash
};
