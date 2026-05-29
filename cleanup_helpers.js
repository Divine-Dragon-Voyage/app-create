const path = require('path');

const TEMP_DOWNLOAD_DIR_NAME = 'app-create-downloads';

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

module.exports = {
    TEMP_DOWNLOAD_DIR_NAME,
    buildTempDownloadRoot,
    isAuxiliaryAutomationUrl,
    isPlayConsoleUrl,
    shouldCloseAuxiliaryPage
};
