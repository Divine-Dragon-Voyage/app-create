const path = require('path');

const PLAY_PROTECTED_WITH_PLAY_PATH = '/protect-with-play';
const PLAY_RELEASES_OVERVIEW_PATH = '/releases/overview';
const LATEST_RELEASES_SCREENSHOT_NAME = 'latest-releases-and-bundles.png';

const APP_SIGNING_PLAY_STORE_PROTECTION_CARD_SELECTORS = [
    'protection-category-card[debug-id="play-store-card"]',
    '[debug-id="play-store-card-container"] protection-category-card',
    '[debug-id="play-store-card-container"]',
    'protection-category-card:has-text("Play Store protection")'
];

const APP_SIGNING_PLAY_STORE_PROTECTION_EXPAND_SELECTORS = [
    'protection-category-card[debug-id="play-store-card"] button[debug-id="expansion-button"][aria-label="Show details"]',
    'protection-category-card[debug-id="play-store-card"] button[debug-id="expansion-button"]',
    '[debug-id="play-store-card-container"] button[debug-id="expansion-button"][aria-label="Show details"]',
    '[debug-id="play-store-card-container"] button[debug-id="expansion-button"]',
    'protection-category-card[debug-id="play-store-card"] button[aria-label="Show details"]'
];

const APP_SIGNING_MANAGE_BUTTON_SELECTORS = [
    'button[aria-label="Manage Play app signing"]',
    'button[debug-id="cta-button"]:has-text("Manage Play app signing")',
    'button:has-text("Manage Play app signing")',
    '[role="button"]:has-text("Manage Play app signing")'
];

const APP_SIGNING_MANAGE_BUTTON_SCOPED_SELECTORS = [
    'protection-category-card[debug-id="play-store-card"] .feature-row:has-text("Protect app signing key") button[aria-label="Manage Play app signing"]',
    'protection-category-card[debug-id="play-store-card"] .feature-row:has-text("Protect app signing key") button[debug-id="cta-button"]:has-text("Manage Play app signing")',
    '[debug-id="play-store-card-container"] .feature-row:has-text("Protect app signing key") button[aria-label="Manage Play app signing"]',
    '[debug-id="play-store-card-container"] .feature-row:has-text("Protect app signing key") button[debug-id="cta-button"]:has-text("Manage Play app signing")',
    'protection-category-card:has-text("Play Store protection") .feature-row:has-text("Protect app signing key") button[aria-label="Manage Play app signing"]',
    'protection-category-card:has-text("Play Store protection") .feature-row:has-text("Protect app signing key") button[debug-id="cta-button"]:has-text("Manage Play app signing")'
];

const APP_SIGNING_SHA1_REGEX = /(?<![:A-Fa-f0-9])(?:[A-Fa-f0-9]{2}:){19}[A-Fa-f0-9]{2}(?![:A-Fa-f0-9])/;

const APPGENIE_REVIEW_FILE_INPUT_SELECTOR = '.ant-modal input[type="file"][accept*="image"]';
const APPGENIE_REVIEW_SHA1_INPUT_SELECTOR = '.ant-modal input[placeholder*="SHA1"], .ant-modal input[placeholder*="SHA-1"]';
const APPGENIE_REVIEW_SUBMIT_BUTTON_SELECTORS = [
    '.ant-modal-footer button.ant-btn-primary',
    '.ant-modal button.ant-btn-primary',
    'button.ant-btn-primary'
];
const PLAY_RELEASE_SEND_FOR_REVIEW_BUTTON_SELECTOR = 'button[debug-id="send-for-review-button"], [role="button"][debug-id="send-for-review-button"]';
const PLAY_RELEASE_SEND_FOR_REVIEW_BUTTON_TEXT_REGEX = /(?:Send|Submit)\s+\d+\s+changes?\s+for\s+review/i;
const PLAY_RELEASE_SEND_FOR_REVIEW_DIALOG_TEXT_REGEX = /(?:Send|Submit)\s+\d+\s+changes?\s+for\s+review|These changes will be sent/i;
const PLAY_RELEASE_CONFIRM_SEND_FOR_REVIEW_BUTTON_TEXT_REGEX = /^(?:Send|Submit)\s+changes?\s+for\s+review$/i;

function normalizeReviewFileToken(value) {
    return String(value || 'app')
        .replace(/[^A-Za-z0-9._-]+/g, '_')
        .replace(/^_+|_+$/g, '')
        .slice(0, 80) || 'app';
}

function extractSha1Fingerprint(text) {
    const match = String(text || '').match(APP_SIGNING_SHA1_REGEX);
    return match ? match[0].toUpperCase() : '';
}

function buildReviewScreenshotPath(rootDir, appName) {
    return path.join(rootDir, `${normalizeReviewFileToken(appName)}-${LATEST_RELEASES_SCREENSHOT_NAME}`);
}

module.exports = {
    APP_SIGNING_PLAY_STORE_PROTECTION_CARD_SELECTORS,
    APP_SIGNING_PLAY_STORE_PROTECTION_EXPAND_SELECTORS,
    APP_SIGNING_MANAGE_BUTTON_SCOPED_SELECTORS,
    APP_SIGNING_MANAGE_BUTTON_SELECTORS,
    APP_SIGNING_SHA1_REGEX,
    APPGENIE_REVIEW_FILE_INPUT_SELECTOR,
    APPGENIE_REVIEW_SHA1_INPUT_SELECTOR,
    APPGENIE_REVIEW_SUBMIT_BUTTON_SELECTORS,
    PLAY_RELEASE_CONFIRM_SEND_FOR_REVIEW_BUTTON_TEXT_REGEX,
    PLAY_RELEASE_SEND_FOR_REVIEW_BUTTON_SELECTOR,
    PLAY_RELEASE_SEND_FOR_REVIEW_BUTTON_TEXT_REGEX,
    PLAY_RELEASE_SEND_FOR_REVIEW_DIALOG_TEXT_REGEX,
    LATEST_RELEASES_SCREENSHOT_NAME,
    PLAY_PROTECTED_WITH_PLAY_PATH,
    PLAY_RELEASES_OVERVIEW_PATH,
    buildReviewScreenshotPath,
    extractSha1Fingerprint,
    normalizeReviewFileToken
};
