const path = require('path');

const PLAY_PROTECTED_WITH_PLAY_PATH = '/protect-with-play';
const PLAY_RELEASES_OVERVIEW_PATH = '/releases/overview';
const LATEST_RELEASES_SCREENSHOT_NAME = 'latest-releases-and-bundles.png';

const APP_SIGNING_MANAGE_BUTTON_SELECTORS = [
    'button[aria-label="Manage Play app signing"]',
    'button[debug-id="cta-button"]:has-text("Manage Play app signing")',
    'button:has-text("Manage Play app signing")',
    '[role="button"]:has-text("Manage Play app signing")'
];

const APP_SIGNING_SHA1_REGEX = /(?<![:A-Fa-f0-9])(?:[A-Fa-f0-9]{2}:){19}[A-Fa-f0-9]{2}(?![:A-Fa-f0-9])/;

const APPGENIE_REVIEW_FILE_INPUT_SELECTOR = '.ant-modal input[type="file"][accept*="image"]';
const APPGENIE_REVIEW_SHA1_INPUT_SELECTOR = '.ant-modal input[placeholder*="SHA1"], .ant-modal input[placeholder*="SHA-1"]';
const APPGENIE_REVIEW_SUBMIT_BUTTON_SELECTORS = [
    '.ant-modal-footer button.ant-btn-primary',
    '.ant-modal button.ant-btn-primary',
    'button.ant-btn-primary'
];

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
    APP_SIGNING_MANAGE_BUTTON_SELECTORS,
    APP_SIGNING_SHA1_REGEX,
    APPGENIE_REVIEW_FILE_INPUT_SELECTOR,
    APPGENIE_REVIEW_SHA1_INPUT_SELECTOR,
    APPGENIE_REVIEW_SUBMIT_BUTTON_SELECTORS,
    LATEST_RELEASES_SCREENSHOT_NAME,
    PLAY_PROTECTED_WITH_PLAY_PATH,
    PLAY_RELEASES_OVERVIEW_PATH,
    buildReviewScreenshotPath,
    extractSha1Fingerprint,
    normalizeReviewFileToken
};
