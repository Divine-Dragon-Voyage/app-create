const test = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');
const {
    TEMP_DOWNLOAD_DIR_NAME,
    FALLBACK_AAB_CLEANUP_AFTER_ROWS,
    buildTempDownloadRoot,
    cleanupTrackedFallbackAabFiles,
    shouldCleanupFallbackAabAfterRow,
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

test('fallback AAB cleanup is fixed to the fourth processed row', () => {
    assert.equal(FALLBACK_AAB_CLEANUP_AFTER_ROWS, 4);
    assert.equal(shouldCleanupFallbackAabAfterRow(1), false);
    assert.equal(shouldCleanupFallbackAabAfterRow(3), false);
    assert.equal(shouldCleanupFallbackAabAfterRow(4), true);
    assert.equal(shouldCleanupFallbackAabAfterRow(5), false);
    assert.equal(shouldCleanupFallbackAabAfterRow(8), false);
});

test('fallback AAB cleanup removes only tracked aab files once', () => {
    const removed = [];
    const existing = new Set([
        path.normalize('C:\\Downloads\\tracked.aab'),
        path.normalize('C:\\Downloads\\tracked-two.aab'),
        path.normalize('C:\\Downloads\\notes.txt')
    ]);
    const tracker = [
        'C:\\Downloads\\tracked.aab',
        'C:\\Downloads\\notes.txt',
        'C:\\Downloads\\missing.aab',
        'C:\\Downloads\\tracked.aab',
        'C:\\Downloads\\tracked-two.aab'
    ];

    const result = cleanupTrackedFallbackAabFiles(tracker, {
        existsSync(filePath) {
            return existing.has(path.normalize(filePath));
        },
        unlinkSync(filePath) {
            removed.push(path.normalize(filePath));
            existing.delete(path.normalize(filePath));
        },
        log() { }
    });

    assert.deepEqual(removed, [
        path.normalize('C:\\Downloads\\tracked.aab'),
        path.normalize('C:\\Downloads\\tracked-two.aab')
    ]);
    assert.deepEqual(result, { removed: 2, failed: 0, skipped: 2 });
});
