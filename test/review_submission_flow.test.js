const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');
const {
    APP_SIGNING_PLAY_STORE_PROTECTION_EXPAND_SELECTORS,
    APP_SIGNING_PLAY_STORE_PROTECTION_CARD_SELECTORS,
    APP_SIGNING_MANAGE_BUTTON_SCOPED_SELECTORS,
    APP_SIGNING_MANAGE_BUTTON_SELECTORS,
    APP_SIGNING_SHA1_REGEX,
    APPGENIE_REVIEW_FILE_INPUT_SELECTOR,
    APPGENIE_REVIEW_SHA1_INPUT_SELECTOR,
    APPGENIE_REVIEW_SUBMIT_BUTTON_SELECTORS,
    LATEST_RELEASES_SCREENSHOT_NAME,
    PLAY_PROTECTED_WITH_PLAY_PATH,
    PLAY_RELEASES_OVERVIEW_PATH,
    extractSha1Fingerprint,
    buildReviewScreenshotPath
} = require('../review_submission_flow');

test('extracts SHA-1 fingerprint from visible Play Console text', () => {
    const text = [
        'MD5 certificate fingerprint 87:3B:DE:A7:6A:CA:D5:F2:49:9B:0D:9C:98:7A:89:C0',
        'SHA-1 certificate fingerprint FB:CE:4D:5E:62:BF:F8:90:95:2A:2A:0A:5E:1D:70:74:BB:D8:FD:B0',
        'SHA-256 certificate fingerprint 55:C0:A6:6D:5D:86:B5:2E:F5:1C:1F:96:6E:45:DD:87:8A:BD:A4:D6:2A:58'
    ].join('\n');

    assert.equal(
        extractSha1Fingerprint(text),
        'FB:CE:4D:5E:62:BF:F8:90:95:2A:2A:0A:5E:1D:70:74:BB:D8:FD:B0'
    );
});

test('extracts first SHA-1 shaped value when page labels are localized', () => {
    const text = 'certificate: aa:bb:cc:dd:ee:ff:00:11:22:33:44:55:66:77:88:99:aa:bb:cc:dd';

    assert.equal(
        extractSha1Fingerprint(text),
        'AA:BB:CC:DD:EE:FF:00:11:22:33:44:55:66:77:88:99:AA:BB:CC:DD'
    );
});

test('does not extract a SHA-256 prefix as SHA-1', () => {
    const text = 'SHA-256 certificate fingerprint 55:C0:A6:6D:5D:86:B5:2E:F5:1C:1F:96:6E:45:DD:87:8A:BD:A4:D6:2A:58';

    assert.equal(extractSha1Fingerprint(text), '');
});

test('review submission selectors cover Play and AppGenie controls', () => {
    assert.equal(PLAY_PROTECTED_WITH_PLAY_PATH, '/protect-with-play');
    assert.equal(PLAY_RELEASES_OVERVIEW_PATH, '/releases/overview');
    assert.ok(APP_SIGNING_PLAY_STORE_PROTECTION_CARD_SELECTORS.some(selector => selector.includes('play-store-card')));
    assert.ok(APP_SIGNING_PLAY_STORE_PROTECTION_EXPAND_SELECTORS.some(selector => selector.includes('expansion-button')));
    assert.ok(APP_SIGNING_MANAGE_BUTTON_SELECTORS.includes('button[aria-label="Manage Play app signing"]'));
    assert.equal(APP_SIGNING_SHA1_REGEX.source, '(?<![:A-Fa-f0-9])(?:[A-Fa-f0-9]{2}:){19}[A-Fa-f0-9]{2}(?![:A-Fa-f0-9])');
    assert.equal(APPGENIE_REVIEW_FILE_INPUT_SELECTOR, '.ant-modal input[type="file"][accept*="image"]');
    assert.ok(APPGENIE_REVIEW_SHA1_INPUT_SELECTOR.includes('SHA1'));
    assert.ok(APPGENIE_REVIEW_SUBMIT_BUTTON_SELECTORS.some(selector => selector.includes('ant-btn-primary')));
});

test('app signing manage selectors are scoped to Play Store protection app signing row', () => {
    assert.ok(APP_SIGNING_MANAGE_BUTTON_SCOPED_SELECTORS.length >= 4);

    for (const selector of APP_SIGNING_MANAGE_BUTTON_SCOPED_SELECTORS) {
        assert.match(selector, /play-store-card|play-store-card-container|Play Store protection/);
        assert.match(selector, /Protect app signing key/);
        assert.match(selector, /Manage Play app signing/);
        assert.doesNotMatch(selector, /^button\[debug-id="cta-button"\]/);
    }
});

test('builds review screenshot path under task temp directory', () => {
    assert.equal(
        buildReviewScreenshotPath('C:\\Temp\\run', 'Code Cipher/Slide'),
        path.join('C:\\Temp\\run', `Code_Cipher_Slide-${LATEST_RELEASES_SCREENSHOT_NAME}`)
    );
});
