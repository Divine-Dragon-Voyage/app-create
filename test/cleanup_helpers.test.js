const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');
const {
    TEMP_DOWNLOAD_DIR_NAME,
    buildTempDownloadRoot,
    isPlayConsoleUrl,
    shouldCloseAuxiliaryPage
} = require('../cleanup_helpers');

test('builds temp download root under system temp directory', () => {
    assert.equal(
        buildTempDownloadRoot('C:\\Temp'),
        path.join('C:\\Temp', TEMP_DOWNLOAD_DIR_NAME)
    );
});

test('keeps play console pages open during cleanup', () => {
    assert.equal(
        isPlayConsoleUrl('https://play.google.com/console/u/0/developers/123/app-list'),
        true
    );
    assert.equal(
        shouldCloseAuxiliaryPage('https://play.google.com/console/u/0/developers/123/app/456'),
        false
    );
});

test('closes appgenie sites and blank auxiliary pages', () => {
    assert.equal(shouldCloseAuxiliaryPage('https://appgenie-ai.com/login'), true);
    assert.equal(shouldCloseAuxiliaryPage('https://sites.google.com/new'), true);
    assert.equal(shouldCloseAuxiliaryPage('about:blank'), true);
});

test('does not close the explicitly kept page', () => {
    const keepUrl = 'https://sites.google.com/new';
    assert.equal(shouldCloseAuxiliaryPage(keepUrl, keepUrl), false);
});
