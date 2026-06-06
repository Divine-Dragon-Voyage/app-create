const path = require('path');

const TEMP_DOWNLOAD_DIR_NAME = 'app-create-downloads';
const FALLBACK_AAB_CLEANUP_AFTER_ROWS = 4;

function buildTempDownloadRoot(tmpDir) {
    return path.join(tmpDir, TEMP_DOWNLOAD_DIR_NAME);
}

function isPlayConsoleUrl(url) {
    return /^https:\/\/play\.google\.com\/console/i.test(String(url || ''));
}

function isAuxiliaryAutomationUrl(url) {
    const text = String(url || '');
    return /appgenie-ai\.com|sites\.google\.com/i.test(text);
}

function shouldCloseAuxiliaryPage(url, keepUrl = '') {
    const text = String(url || '');
    if (keepUrl && text === String(keepUrl || '')) {
        return false;
    }
    if (isPlayConsoleUrl(text)) {
        return false;
    }
    return isAuxiliaryAutomationUrl(text) || text === 'about:blank';
}

function shouldCleanupFallbackAabAfterRow(processedRows) {
    return Number(processedRows) === FALLBACK_AAB_CLEANUP_AFTER_ROWS;
}

function cleanupTrackedFallbackAabFiles(trackedPaths, options = {}) {
    const nodeFs = require('fs');
    const fsApi = options.fs || {
        existsSync: options.existsSync || nodeFs.existsSync,
        unlinkSync: options.unlinkSync || nodeFs.unlinkSync
    };
    const log = typeof options.log === 'function' ? options.log : console.log;
    const seen = new Set();
    const result = { removed: 0, failed: 0, skipped: 0 };

    for (const filePath of trackedPaths || []) {
        const resolved = path.resolve(String(filePath || ''));
        if (!resolved || seen.has(resolved)) {
            continue;
        }
        seen.add(resolved);

        // 只清理脚本记录过的 AAB，避免误删 Downloads 里的其他用户文件。
        if (path.extname(resolved).toLowerCase() !== '.aab' || !fsApi.existsSync(resolved)) {
            result.skipped += 1;
            continue;
        }

        try {
            fsApi.unlinkSync(resolved);
            result.removed += 1;
            log(`[CLEANUP] Removed fallback AAB: ${resolved}`);
        } catch (err) {
            result.failed += 1;
            log(`[CLEANUP] Could not remove fallback AAB (${resolved}): ${err.message}`);
        }
    }

    return result;
}

module.exports = {
    TEMP_DOWNLOAD_DIR_NAME,
    FALLBACK_AAB_CLEANUP_AFTER_ROWS,
    buildTempDownloadRoot,
    cleanupTrackedFallbackAabFiles,
    isAuxiliaryAutomationUrl,
    isPlayConsoleUrl,
    shouldCleanupFallbackAabAfterRow,
    shouldCloseAuxiliaryPage
};
