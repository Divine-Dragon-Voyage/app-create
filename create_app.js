const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const { execFileSync } = require('child_process');
const XLSX = require('xlsx');
const { chromium } = require('playwright');

// Default delay (ms). Increase if VPS is slow.
const DELAY = 3000;

const APP_NAME_HEADER_CANDIDATES = new Set([
    '应用名称',
    '应用名',
    'appname',
    'applicationname'
]);

const PACKAGE_NAME_HEADER_CANDIDATES = new Set([
    '应用包名',
    '包名',
    'apppackagename',
    'packagename',
    'applicationid'
]);

const PACKAGE_NAME_REGEX = /^[a-z][a-z0-9_]*(\.[a-z][a-z0-9_]*)+$/;
const CONFIG_DIR_ENV = 'APP_CREATE_CONFIG_DIR';
const CDP_ENDPOINT_ENV = 'APP_CREATE_CDP_ENDPOINT';
const DEVELOPER_URL_ENV = 'APP_CREATE_DEVELOPER_URL';
const DEVELOPER_ID_ENV = 'APP_CREATE_DEVELOPER_ID';
const STATUS_HEADER_CANDIDATES = new Set(['status', '\u72B6\u6001']);
const PROGRESS_HEADER_CANDIDATES = new Set(['progressstep', 'progress', '\u8fdb\u5ea6', '\u6b65\u9aa4']);
const STATUS_PARTIAL = 'PARTIAL';
const STATUS_DONE = 'DONE';
const STATUS_FAILED = 'FAILED';
const PROGRESS_STEP_APP_CREATED = 'APP_CREATED';
const PROGRESS_STEP_ADS_DONE = 'ADS_DONE';
const PROGRESS_STEP_APP_ACCESS_DONE = 'APP_ACCESS_DONE';
const PROGRESS_STEP_AUDIENCE_DONE = 'AUDIENCE_DONE';
const PROGRESS_STEP_AD_ID_DONE = 'AD_ID_DONE';
const PROGRESS_STEP_GOV_DONE = 'GOV_DONE';
const PROGRESS_STEP_FINANCE_DONE = 'FINANCE_DONE';
const PROGRESS_STEP_HEALTH_DONE = 'HEALTH_DONE';
const PROGRESS_STEP_COUNTRY_DONE = 'COUNTRY_DONE';
const PROGRESS_STEP_DONE = 'DONE';
const PROGRESS_STEP_ORDER = [
    PROGRESS_STEP_APP_CREATED,
    PROGRESS_STEP_ADS_DONE,
    PROGRESS_STEP_APP_ACCESS_DONE,
    PROGRESS_STEP_AUDIENCE_DONE,
    PROGRESS_STEP_AD_ID_DONE,
    PROGRESS_STEP_GOV_DONE,
    PROGRESS_STEP_FINANCE_DONE,
    PROGRESS_STEP_HEALTH_DONE,
    PROGRESS_STEP_COUNTRY_DONE,
    PROGRESS_STEP_DONE
];
const PROGRESS_STEP_SET = new Set(PROGRESS_STEP_ORDER);
const DEVELOPER_URL_CONFIG_FILE = 'developer_url.txt';
const SUPPORTED_INPUT_EXTENSIONS = new Set(['.xlsx', '.xls', '.csv']);
const CSV_STATE_DIR_NAME = 'csv-state';
const DEVELOPER_URL_TEMPLATE = [
    '# Paste your Play Console developer URL below (single line).',
    '# Example:',
    '# https://play.google.com/console/u/0/developers/1234567890123456789/app-list',
    '',
    'https://play.google.com/console/u/0/developers/REPLACE_WITH_YOUR_DEVELOPER_ID/app-list'
    
].join('\n');

function normalizeHeader(value) {
    return String(value || '')
        .trim()
        .toLowerCase()
        .replace(/[\s_-]/g, '');
}

function parseCliArgs() {
    const args = process.argv.slice(2);
    let inputFileArg;

    if (args.length > 1) {
        throw new Error('Usage: node create_app.js [data_file_path]');
    }

    if (args.length === 1) {
        inputFileArg = args[0];
    }

    return { inputFileArg };
}

function resolveConfigDirectory() {
    const configuredDir = String(process.env[CONFIG_DIR_ENV] || '').trim();
    if (!configuredDir) {
        return process.cwd();
    }
    return path.resolve(configuredDir);
}

function ensureDeveloperUrlTemplate(configPath) {
    if (fs.existsSync(configPath)) {
        return;
    }

    fs.mkdirSync(path.dirname(configPath), { recursive: true });
    try {
        fs.writeFileSync(configPath, DEVELOPER_URL_TEMPLATE + '\n', { encoding: 'utf8', flag: 'wx' });
    } catch (_) {
        // Ignore race conditions when multiple processes create the file at once.
    }
}

function parseDeveloperBaseUrl(rawUrl) {
    const text = String(rawUrl || '').trim();
    if (!text) {
        throw new Error('Developer URL is empty.');
    }

    let parsed;
    try {
        parsed = new URL(text);
    } catch (_) {
        throw new Error(`Developer URL is invalid: "${text}"`);
    }

    if (parsed.hostname !== 'play.google.com') {
        throw new Error(`Developer URL host must be play.google.com, got: ${parsed.hostname}`);
    }

    const normalized = `${parsed.origin}${parsed.pathname}`;
    const match = normalized.match(
        /^(https?:\/\/play\.google\.com\/(?:console(?:\/u\/\d+)?\/)?developers\/\d+)/
    );

    if (!match) {
        throw new Error(
            'Developer URL must contain "/developers/<developer_id>". ' +
            `Current URL: "${text}"`
        );
    }

    return match[1];
}

function loadDeveloperConsoleAppListUrl() {
    const envUrl = String(process.env[DEVELOPER_URL_ENV] || '').trim();
    if (envUrl) {
        const baseUrl = parseDeveloperBaseUrl(envUrl);
        return { appListUrl: `${baseUrl}/app-list`, configPath: `env:${DEVELOPER_URL_ENV}` };
    }

    const envDeveloperId = String(process.env[DEVELOPER_ID_ENV] || '').trim();
    if (envDeveloperId) {
        if (!/^\d+$/.test(envDeveloperId)) {
            throw new Error(
                `Environment ${DEVELOPER_ID_ENV} must be numeric, got: "${envDeveloperId}".`
            );
        }
        const baseUrl = parseDeveloperBaseUrl(
            `https://play.google.com/console/u/0/developers/${envDeveloperId}/app-list`
        );
        return { appListUrl: `${baseUrl}/app-list`, configPath: `env:${DEVELOPER_ID_ENV}` };
    }

    const configDir = resolveConfigDirectory();
    const configPath = path.resolve(configDir, DEVELOPER_URL_CONFIG_FILE);
    ensureDeveloperUrlTemplate(configPath);

    const fileContent = fs.readFileSync(configPath, 'utf8');
    const rawUrl = fileContent
        .split(/\r?\n/)
        .map(line => line.trim())
        .find(line => line && !line.startsWith('#'));

    if (!rawUrl) {
        throw new Error(
            `Missing developer URL in config file: ${configPath}. ` +
            'Please paste your Play Console URL into that file.'
        );
    }

    const baseUrl = parseDeveloperBaseUrl(rawUrl);
    return { appListUrl: `${baseUrl}/app-list`, configPath };
}

function resolveInputExcelFile(inputFileArg) {
    if (inputFileArg) {
        const explicitPath = path.resolve(process.cwd(), inputFileArg);
        if (!fs.existsSync(explicitPath)) {
            throw new Error(`Input file not found: ${explicitPath}`);
        }
        const explicitExt = path.extname(explicitPath).toLowerCase();
        if (!SUPPORTED_INPUT_EXTENSIONS.has(explicitExt)) {
            throw new Error(
                `Unsupported file extension "${explicitExt}". ` +
                'Only .xlsx/.xls/.csv are supported.'
            );
        }
        return explicitPath;
    }

    const dataFiles = fs.readdirSync(process.cwd())
        .filter(name => SUPPORTED_INPUT_EXTENSIONS.has(path.extname(name).toLowerCase()))
        .sort();

    if (!dataFiles.length) {
        throw new Error(
            'No input file found in project root. Put one .xlsx/.xls/.csv file here, ' +
            'or pass path: node create_app.js ./apps.xlsx'
        );
    }

    return path.resolve(process.cwd(), dataFiles[0]);
}

function buildCsvStateFilePath(csvPath) {
    const configDir = resolveConfigDirectory();
    const stateDir = path.resolve(configDir, CSV_STATE_DIR_NAME);
    fs.mkdirSync(stateDir, { recursive: true });

    const key = path.resolve(csvPath).toLowerCase();
    const hash = crypto.createHash('sha1').update(key).digest('hex');
    return path.join(stateDir, `${hash}.json`);
}

function resolveRuntimeInput(inputFilePath) {
    const ext = path.extname(inputFilePath).toLowerCase();
    if (ext !== '.csv') {
        return {
            sourceFilePath: inputFilePath,
            runtimeWorkbookPath: inputFilePath,
            isCsvInput: false,
            stateFilePath: ''
        };
    }

    return {
        sourceFilePath: inputFilePath,
        runtimeWorkbookPath: inputFilePath,
        isCsvInput: true,
        stateFilePath: buildCsvStateFilePath(inputFilePath)
    };
}

function loadCsvStateRows(stateFilePath) {
    if (!stateFilePath || !fs.existsSync(stateFilePath)) {
        return {};
    }

    try {
        const raw = fs.readFileSync(stateFilePath, 'utf8');
        const parsed = JSON.parse(raw || '{}');
        if (!parsed || typeof parsed !== 'object') {
            return {};
        }
        if (!parsed.rows || typeof parsed.rows !== 'object') {
            return {};
        }
        return parsed.rows;
    } catch (_) {
        return {};
    }
}

function saveCsvStateRows(stateFilePath, sourceFilePath, rows) {
    if (!stateFilePath) return;

    const payload = {
        sourceFilePath,
        updatedAt: new Date().toISOString(),
        rows
    };
    fs.writeFileSync(stateFilePath, JSON.stringify(payload, null, 2), 'utf8');
}

function getCsvStateKey(packageName) {
    return String(packageName || '').trim().toLowerCase();
}

function pickHeader(headers, candidates) {
    for (const header of headers) {
        if (candidates.has(normalizeHeader(header))) {
            return header;
        }
    }
    return null;
}

function normalizeStatusValue(value) {
    const text = String(value || '').trim().toUpperCase();
    if (text === STATUS_DONE) return STATUS_DONE;
    if (text === STATUS_PARTIAL) return STATUS_PARTIAL;
    if (text === STATUS_FAILED) return STATUS_FAILED;
    return '';
}

function normalizeProgressStep(value) {
    const text = String(value || '').trim().toUpperCase();
    if (PROGRESS_STEP_SET.has(text)) return text;
    return '';
}

function progressRank(step) {
    return PROGRESS_STEP_ORDER.indexOf(normalizeProgressStep(step));
}

function showWindowsSummaryPopup(summaryText) {
    if (process.platform !== 'win32') {
        return;
    }
    try {
        const ps = [
            'Add-Type -AssemblyName System.Windows.Forms;',
            '[void][System.Windows.Forms.MessageBox]::Show(',
            '  $args[0],',
            '  "App Create 运行结果",',
            '  [System.Windows.Forms.MessageBoxButtons]::OK,',
            '  [System.Windows.Forms.MessageBoxIcon]::Information',
            ')'
        ].join(' ');
        execFileSync(
            'powershell.exe',
            ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-Command', ps, summaryText],
            { stdio: 'ignore', windowsHide: true }
        );
    } catch (err) {
        console.log(`[WARN] Cannot show summary popup. ${err.message}`);
    }
}

function getCellText(sheet, rowIndex, colIndex) {
    const address = XLSX.utils.encode_cell({ r: rowIndex, c: colIndex });
    const cell = sheet[address];
    if (!cell || cell.v === undefined || cell.v === null) {
        return '';
    }
    return String(cell.v).trim();
}

function setCellText(sheet, rowIndex, colIndex, text) {
    const address = XLSX.utils.encode_cell({ r: rowIndex, c: colIndex });
    sheet[address] = { t: 's', v: String(text || '') };
}

function createExcelStatusManager({
    filePath,
    workbook,
    sheetName,
    sheet,
    statusColumnIndex,
    progressColumnIndex
}) {
    function saveWorkbook() {
        XLSX.writeFile(workbook, filePath);
    }

    return {
        saveWorkbook,
        updateTaskStatus(task, status) {
            const nextStatus = normalizeStatusValue(status);
            if (!nextStatus) {
                throw new Error(`Invalid status "${status}" for row ${task.rowNumber}.`);
            }

            const currentStatus = normalizeStatusValue(getCellText(sheet, task.sheetRowIndex, statusColumnIndex));
            if (currentStatus === nextStatus) {
                task.status = nextStatus;
                return;
            }

            setCellText(sheet, task.sheetRowIndex, statusColumnIndex, nextStatus);
            saveWorkbook();
            task.status = nextStatus;

            console.log(
                `[STATUS] Row ${task.rowNumber} (${task.appName}/${task.packageName}): ` +
                `${currentStatus || 'EMPTY'} -> ${nextStatus}`
            );
        },
        getTaskStatus(task) {
            return normalizeStatusValue(getCellText(sheet, task.sheetRowIndex, statusColumnIndex));
        },
        updateTaskProgress(task, step) {
            const nextStep = normalizeProgressStep(step);
            if (!nextStep) {
                throw new Error(`Invalid progress step "${step}" for row ${task.rowNumber}.`);
            }

            const currentStep = normalizeProgressStep(getCellText(sheet, task.sheetRowIndex, progressColumnIndex));
            if (currentStep === nextStep) {
                task.progressStep = nextStep;
                return;
            }

            setCellText(sheet, task.sheetRowIndex, progressColumnIndex, nextStep);
            saveWorkbook();
            task.progressStep = nextStep;

            console.log(
                `[PROGRESS] Row ${task.rowNumber} (${task.appName}/${task.packageName}): ` +
                `${currentStep || 'EMPTY'} -> ${nextStep}`
            );
        },
        ensureTaskProgressAtLeast(task, step) {
            const nextStep = normalizeProgressStep(step);
            if (!nextStep) {
                throw new Error(`Invalid progress step "${step}" for row ${task.rowNumber}.`);
            }

            const currentStep = normalizeProgressStep(getCellText(sheet, task.sheetRowIndex, progressColumnIndex));
            if (progressRank(currentStep) >= progressRank(nextStep)) {
                task.progressStep = currentStep || task.progressStep;
                return;
            }

            setCellText(sheet, task.sheetRowIndex, progressColumnIndex, nextStep);
            saveWorkbook();
            task.progressStep = nextStep;

            console.log(
                `[PROGRESS] Row ${task.rowNumber} (${task.appName}/${task.packageName}): ` +
                `${currentStep || 'EMPTY'} -> ${nextStep}`
            );
        },
        getTaskProgress(task) {
            return normalizeProgressStep(getCellText(sheet, task.sheetRowIndex, progressColumnIndex));
        },
        statusColumnIndex,
        progressColumnIndex,
        sheetName
    };
}

function createCsvStatusManager({
    stateFilePath,
    sourceFilePath,
    rows
}) {
    const persist = () => saveCsvStateRows(stateFilePath, sourceFilePath, rows);

    return {
        saveWorkbook() {
            persist();
        },
        updateTaskStatus(task, status) {
            const nextStatus = normalizeStatusValue(status);
            if (!nextStatus) {
                throw new Error(`Invalid status "${status}" for row ${task.rowNumber}.`);
            }

            const key = getCsvStateKey(task.packageName);
            const currentRow = rows[key] || {};
            const currentStatus = normalizeStatusValue(currentRow.status);
            if (currentStatus === nextStatus) {
                task.status = nextStatus;
                return;
            }

            rows[key] = {
                ...currentRow,
                rowNumber: task.rowNumber,
                appName: task.appName,
                packageName: task.packageName,
                status: nextStatus,
                progressStep: normalizeProgressStep(currentRow.progressStep)
            };
            persist();
            task.status = nextStatus;

            console.log(
                `[STATUS] Row ${task.rowNumber} (${task.appName}/${task.packageName}): ` +
                `${currentStatus || 'EMPTY'} -> ${nextStatus}`
            );
        },
        getTaskStatus(task) {
            const key = getCsvStateKey(task.packageName);
            return normalizeStatusValue((rows[key] || {}).status);
        },
        updateTaskProgress(task, step) {
            const nextStep = normalizeProgressStep(step);
            if (!nextStep) {
                throw new Error(`Invalid progress step "${step}" for row ${task.rowNumber}.`);
            }

            const key = getCsvStateKey(task.packageName);
            const currentRow = rows[key] || {};
            const currentStep = normalizeProgressStep(currentRow.progressStep);
            if (currentStep === nextStep) {
                task.progressStep = nextStep;
                return;
            }

            rows[key] = {
                ...currentRow,
                rowNumber: task.rowNumber,
                appName: task.appName,
                packageName: task.packageName,
                status: normalizeStatusValue(currentRow.status),
                progressStep: nextStep
            };
            persist();
            task.progressStep = nextStep;

            console.log(
                `[PROGRESS] Row ${task.rowNumber} (${task.appName}/${task.packageName}): ` +
                `${currentStep || 'EMPTY'} -> ${nextStep}`
            );
        },
        ensureTaskProgressAtLeast(task, step) {
            const nextStep = normalizeProgressStep(step);
            if (!nextStep) {
                throw new Error(`Invalid progress step "${step}" for row ${task.rowNumber}.`);
            }

            const key = getCsvStateKey(task.packageName);
            const currentRow = rows[key] || {};
            const currentStep = normalizeProgressStep(currentRow.progressStep);
            if (progressRank(currentStep) >= progressRank(nextStep)) {
                task.progressStep = currentStep || task.progressStep;
                return;
            }

            rows[key] = {
                ...currentRow,
                rowNumber: task.rowNumber,
                appName: task.appName,
                packageName: task.packageName,
                status: normalizeStatusValue(currentRow.status),
                progressStep: nextStep
            };
            persist();
            task.progressStep = nextStep;

            console.log(
                `[PROGRESS] Row ${task.rowNumber} (${task.appName}/${task.packageName}): ` +
                `${currentStep || 'EMPTY'} -> ${nextStep}`
            );
        },
        getTaskProgress(task) {
            const key = getCsvStateKey(task.packageName);
            return normalizeProgressStep((rows[key] || {}).progressStep);
        },
        statusColumnIndex: -1,
        progressColumnIndex: -1,
        sheetName: 'Sheet1'
    };
}

function loadTasksFromExcel(filePath, options = {}) {
    const persistMode = options.persistMode || 'workbook';
    const csvStateFilePath = options.csvStateFilePath || '';
    const sourceFilePath = options.sourceFilePath || filePath;
    const csvStateRows = persistMode === 'csv-state' ? loadCsvStateRows(csvStateFilePath) : {};

    const workbook = XLSX.readFile(filePath);
    const firstSheetName = workbook.SheetNames[0];
    if (!firstSheetName) {
        throw new Error(`No sheet found in input file: ${filePath}`);
    }

    const sheet = workbook.Sheets[firstSheetName];
    const ref = sheet['!ref'];
    if (!ref) {
        throw new Error(`Sheet "${firstSheetName}" is empty.`);
    }
    const range = XLSX.utils.decode_range(ref);
    const headerRowIndex = range.s.r;

    const headers = [];
    let appNameColumnIndex = -1;
    let packageNameColumnIndex = -1;
    let statusColumnIndex = -1;
    let progressColumnIndex = -1;

    for (let c = range.s.c; c <= range.e.c; c++) {
        const headerValue = getCellText(sheet, headerRowIndex, c);
        headers.push(headerValue);
        const normalized = normalizeHeader(headerValue);
        if (appNameColumnIndex < 0 && APP_NAME_HEADER_CANDIDATES.has(normalized)) {
            appNameColumnIndex = c;
        }
        if (packageNameColumnIndex < 0 && PACKAGE_NAME_HEADER_CANDIDATES.has(normalized)) {
            packageNameColumnIndex = c;
        }
        if (statusColumnIndex < 0 && STATUS_HEADER_CANDIDATES.has(normalized)) {
            statusColumnIndex = c;
        }
        if (progressColumnIndex < 0 && PROGRESS_HEADER_CANDIDATES.has(normalized)) {
            progressColumnIndex = c;
        }
    }

    if (appNameColumnIndex < 0 || packageNameColumnIndex < 0) {
        throw new Error(
            `Missing required columns. Found headers: ${headers.join(', ')}. ` +
            'Need columns like: 应用名称 / 应用包名 (or App Name / App Package Name).'
        );
    }

    let statusColumnAdded = false;
    if (statusColumnIndex < 0 && persistMode === 'workbook') {
        statusColumnIndex = range.e.c + 1;
        setCellText(sheet, headerRowIndex, statusColumnIndex, 'status');
        range.e.c = statusColumnIndex;
        sheet['!ref'] = XLSX.utils.encode_range(range);
        statusColumnAdded = true;
        console.log('[STATUS] Added "status" column automatically.');
    }

    let progressColumnAdded = false;
    if (progressColumnIndex < 0 && persistMode === 'workbook') {
        progressColumnIndex = range.e.c + 1;
        setCellText(sheet, headerRowIndex, progressColumnIndex, 'progress_step');
        range.e.c = progressColumnIndex;
        sheet['!ref'] = XLSX.utils.encode_range(range);
        progressColumnAdded = true;
        console.log('[PROGRESS] Added "progress_step" column automatically.');
    }

    const statusManager = persistMode === 'csv-state'
        ? createCsvStatusManager({
            stateFilePath: csvStateFilePath,
            sourceFilePath,
            rows: csvStateRows
        })
        : createExcelStatusManager({
            filePath,
            workbook,
            sheetName: firstSheetName,
            sheet,
            statusColumnIndex,
            progressColumnIndex
        });

    if (statusColumnAdded || progressColumnAdded) {
        statusManager.saveWorkbook();
    }

    const tasks = [];

    for (let r = headerRowIndex + 1; r <= range.e.r; r++) {
        const excelRowNumber = r + 1;
        const appName = getCellText(sheet, r, appNameColumnIndex);
        const rawPackageName = getCellText(sheet, r, packageNameColumnIndex);
        const packageName = rawPackageName.toLowerCase();
        const rowState = persistMode === 'csv-state'
            ? (csvStateRows[getCsvStateKey(packageName)] || {})
            : null;

        let status = persistMode === 'csv-state'
            ? normalizeStatusValue(rowState.status)
            : normalizeStatusValue(getCellText(sheet, r, statusColumnIndex));
        let progressStep = persistMode === 'csv-state'
            ? normalizeProgressStep(rowState.progressStep)
            : normalizeProgressStep(getCellText(sheet, r, progressColumnIndex));

        if (status === STATUS_DONE && !progressStep) {
            progressStep = PROGRESS_STEP_DONE;
        }
        if (status !== STATUS_DONE && progressStep) {
            status = STATUS_PARTIAL;
        }

        if (!appName && !packageName) {
            continue;
        }

        if (!appName || !packageName) {
            throw new Error(`Row ${excelRowNumber}: both app name and package name are required.`);
        }

        if (appName.length > 30) {
            throw new Error(`Row ${excelRowNumber}: app name exceeds 30 chars (${appName.length}).`);
        }

        if (!PACKAGE_NAME_REGEX.test(packageName)) {
            throw new Error(
                `Row ${excelRowNumber}: invalid package name "${rawPackageName}". ` +
                'Expected format like "com.example.appname" (lowercase letters/numbers/underscore).'
            );
        }

        tasks.push({
            appName,
            packageName,
            rowNumber: excelRowNumber,
            sheetRowIndex: r,
            status,
            progressStep
        });
    }

    if (!tasks.length) {
        throw new Error(`No valid rows found in sheet "${firstSheetName}".`);
    }

    return { tasks, sheetName: firstSheetName, statusManager };
}

function formatDuration(totalSeconds) {
    const hours = Math.floor(totalSeconds / 3600);
    const minutes = Math.floor((totalSeconds % 3600) / 60);
    const seconds = totalSeconds % 60;
    return [
        hours > 0 ? `${hours}h ` : '',
        minutes > 0 ? `${minutes}m ` : '',
        `${seconds}s`
    ].join('').trim() || '0s';
}

async function delay(page, ms = DELAY) {
    await page.waitForTimeout(ms);
}

async function waitForEnabled(locator, timeoutMs = 15000) {
    const started = Date.now();
    while (Date.now() - started < timeoutMs) {
        if (await locator.isEnabled().catch(() => false)) {
            return true;
        }
        await new Promise(r => setTimeout(r, 500));
    }
    return false;
}

// Global retry wrapper for UI interactions.
function getCdpEndpoints() {
    const configured = String(process.env[CDP_ENDPOINT_ENV] || '').trim();
    if (configured) {
        return [configured];
    }
    // Prefer IPv4 first to avoid localhost -> ::1 connection failures.
    return ['http://127.0.0.1:9222', 'http://localhost:9222'];
}

function isCdpConnectionError(err) {
    const message = String((err && err.message) || '');
    return /connectOverCDP|ECONNREFUSED|9222|CDP/i.test(message);
}

async function connectBrowserOverCdp() {
    const endpoints = getCdpEndpoints();
    let lastError;

    for (const endpoint of endpoints) {
        try {
            console.log(`Connecting to Chrome CDP endpoint: ${endpoint}`);
            const browser = await chromium.connectOverCDP(endpoint, { timeout: 120000 });
            return { browser, endpoint };
        } catch (err) {
            lastError = err;
            console.log(`CDP connection failed at ${endpoint}: ${err.message}`);
        }
    }

    throw new Error(
        `Could not connect to Chrome CDP. Tried: ${endpoints.join(', ')}. ` +
        'Please start Chrome with --remote-debugging-port=9222 and verify: ' +
        'http://127.0.0.1:9222/json/version',
        { cause: lastError }
    );
}

async function ensureBrowserWindowMaximized(page) {
    try {
        const cdpSession = await page.context().newCDPSession(page);
        const { windowId } = await cdpSession.send('Browser.getWindowForTarget');
        if (windowId === undefined || windowId === null) {
            console.log('Warning: Could not resolve browser window id for maximize check.');
            return;
        }

        const boundsResult = await cdpSession.send('Browser.getWindowBounds', { windowId });
        const currentState = boundsResult && boundsResult.bounds ? boundsResult.bounds.windowState : '';
        if (currentState === 'maximized' || currentState === 'fullscreen') {
            return;
        }

        await cdpSession.send('Browser.setWindowBounds', {
            windowId,
            bounds: { windowState: 'maximized' }
        });
        await page.waitForTimeout(500);
        console.log('Browser window was not maximized; switched to maximized state.');
    } catch (err) {
        console.log(`Warning: Failed to enforce maximized browser window: ${err.message}`);
    }
}

async function retryAction(action, label = 'action', retries = 3) {
    for (let i = 0; i < retries; i++) {
        try {
            return await action();
        } catch (e) {
            console.log(`[Retry] ${label} failed (attempt ${i + 1}/${retries}): ${e.message}`);
            if (i === retries - 1) throw e;
            await new Promise(r => setTimeout(r, 5000 + Math.random() * 5000));
        }
    }
}

async function waitSaved(page) {
    console.log('Waiting for save confirmation...');
    // Loosely match various "saved" messages (case-insensitive).
    try {
        await page.locator('text=/saved/i').first().waitFor({ state: 'visible', timeout: 30000 });
        console.log('Save detected.');
    } catch (e) {
        console.log('Save confirmation timeout, continuing...');
    }
    await delay(page, 5000); // Extra wait after save to allow backend processing.
}

async function runOnce(task, appListUrl, statusManager) {
    const startTime = Date.now();
    const appName = task.appName;
    const packageName = task.packageName;
    console.log(
        `Running task. Name="${appName}", Package="${packageName}", ` +
        `Status="${task.status || 'EMPTY'}", Progress="${task.progressStep || 'EMPTY'}"`
    );

    let browser;
    let page;
    try {
        const cdpConnection = await connectBrowserOverCdp();
        browser = cdpConnection.browser;
        console.log(`Connected via CDP: ${cdpConnection.endpoint}`);
        const context = browser.contexts()[0];
        page = await context.newPage();

        await page.bringToFront();
        await ensureBrowserWindowMaximized(page);
        page.setDefaultTimeout(60000); // Set page-wide timeout to 60s.

        let appBasePath = '';
        const extractAppBasePathFromCurrentUrl = () => {
            const currentUrl = page.url();
            const baseMatch = currentUrl.match(/(.*\/app\/\d+)/);
            if (!baseMatch) {
                throw new Error('Could not extract app base path from URL: ' + currentUrl);
            }
            return baseMatch[1];
        };

        const openAppListPage = async () => {
            const createBtn = page.locator('[debug-id="create-app-button"], a:has-text("Create app"), button:has-text("Create app")').first();

            const ensureDeveloperSelectedIfNeeded = async () => {
                const devItem = page.locator('developer-item, [debug-id="all-developers"]').first();
                if (await devItem.isVisible().catch(() => false)) {
                    console.log('Developer picker detected, clicking first developer item...');
                    await devItem.click();
                    await delay(page, 7000);
                }
            };

            console.log('Opening app list page...');
            await page.goto(appListUrl, { timeout: 90000, waitUntil: 'domcontentloaded' });
            await delay(page, 4000);
            await ensureDeveloperSelectedIfNeeded();

            // Some accounts redirect through picker/home once; enforce app-list one more time.
            if (!(await createBtn.isVisible().catch(() => false))) {
                console.log('App list not ready yet, retrying app-list navigation...');
                await page.goto(appListUrl, { timeout: 90000, waitUntil: 'domcontentloaded' });
                await delay(page, 4000);
                await ensureDeveloperSelectedIfNeeded();
            }
        };

        const openExistingAppForPartial = async () => {
            console.log(`[RESUME] status=PARTIAL, locating existing app: ${appName} / ${packageName}`);
            const searchInput = page.locator(
                'input[placeholder*="Search by app or package"], input[aria-label*="Search by app or package"], input[type="search"]'
            ).first();
            if (await searchInput.isVisible().catch(() => false)) {
                await searchInput.fill(packageName);
                await delay(page, 1200);
            }

            const appTextContainer = page.locator('div.text-container')
                .filter({ hasText: appName })
                .filter({ hasText: packageName })
                .first();

            await retryAction(async () => {
                await appTextContainer.waitFor({ state: 'visible', timeout: 20000 });
            }, 'Find PARTIAL app row', 2);

            let viewAppBtn = appTextContainer.locator(
                'xpath=ancestor::tr[1]//*[self::a or self::button or @role="button"][contains(normalize-space(.), "View app")]'
            ).first();

            if (!(await viewAppBtn.isVisible().catch(() => false))) {
                viewAppBtn = page.locator('a:has-text("View app"), button:has-text("View app"), [role="button"]:has-text("View app")').first();
            }

            await retryAction(async () => {
                await viewAppBtn.scrollIntoViewIfNeeded({ timeout: 3000 }).catch(() => { });
                await viewAppBtn.click({ timeout: 10000 });
            }, 'Click View app for PARTIAL row', 2);

            try {
                await page.waitForURL(/\/app\/\d+/, { timeout: 60000 });
            } catch (_) {
                console.log('Warning: URL did not switch to app details in time, continuing...');
            }
            await delay(page, 4000);
            appBasePath = extractAppBasePathFromCurrentUrl();
            console.log('[RESUME] App path:', appBasePath);
        };

        await openAppListPage();

        if (task.status === STATUS_PARTIAL) {
            await openExistingAppForPartial();
            statusManager.ensureTaskProgressAtLeast(task, PROGRESS_STEP_APP_CREATED);
            statusManager.updateTaskStatus(task, STATUS_PARTIAL);
        } else {
            console.log('Clicking "Create app" button on list page...');
            const createBtn = page.locator('[debug-id="create-app-button"], a:has-text("Create app"), button:has-text("Create app")').first();
            const appNameInput = page.locator(
                '[debug-id="app-name-input"] input, input[aria-label="App name"]'
            ).first();
            const packageNameInput = page.locator(
                '[debug-id="app-package-name-input"] input, input[aria-label="App package name"]'
            ).first();

            const openCreateFormWithRetry = async () => {
                const maxAttempts = 4;
                for (let attempt = 1; attempt <= maxAttempts; attempt++) {
                    console.log(`[CREATE] Open form attempt ${attempt}/${maxAttempts}...`);

                    if (!(await createBtn.isVisible().catch(() => false))) {
                        console.log('[CREATE] Create app button not visible, reopening app list...');
                        await openAppListPage();
                    }

                    await retryAction(async () => {
                        await createBtn.waitFor({ state: 'visible', timeout: 20000 });
                        await createBtn.scrollIntoViewIfNeeded({ timeout: 5000 }).catch(() => { });
                        await createBtn.click({ timeout: 10000 });
                    }, `Click Create App button (attempt ${attempt})`, 2);

                    const formReady = await appNameInput
                        .waitFor({ state: 'visible', timeout: 15000 })
                        .then(() => true)
                        .catch(() => false);

                    if (formReady) {
                        await delay(page, 1500);
                        return;
                    }

                    console.log('[CREATE] Form still not visible after click.');
                    if (attempt < maxAttempts) {
                        console.log('[CREATE] Reopening app list and retrying...');
                        await openAppListPage();
                    }
                }

                throw new Error(
                    'Create app form did not appear after multiple click attempts. ' +
                    'Likely page not fully ready or click not accepted by UI.'
                );
            };

            await openCreateFormWithRetry();

            // Fill app name and package name.
            console.log('Filling App name...');
            await retryAction(async () => {
                await appNameInput.fill(appName);
            }, 'Fill App name');
            await delay(page, 1500);

            console.log('Filling App package name...');
            await retryAction(async () => {
                await packageNameInput.waitFor({ state: 'visible', timeout: 15000 });
                await packageNameInput.fill(packageName);
            }, 'Fill App package name');
            await delay(page, 1500);

            const checkPackageBtn = page.locator('[debug-id="check-package-name-availability-button"]').first();
            if (await checkPackageBtn.isVisible().catch(() => false)) {
                console.log('Checking package name availability...');
                const enabled = await waitForEnabled(checkPackageBtn, 15000);
                if (enabled) {
                    await retryAction(async () => {
                        await checkPackageBtn.click({ timeout: 10000 });
                    }, 'Click Check availability');
                    await delay(page, 5000);
                } else {
                    console.log('Warning: "Check availability" button did not become enabled in time.');
                }
            } else {
                console.log('Warning: Check availability button not found, continuing...');
            }

            // Randomly choose App or Game.
            const isApp = Math.random() < 0.5;
            const typeLabel = isApp ? 'App' : 'Game';
            const typeId = isApp ? 'app-radio' : 'game-radio';
            console.log(`Selecting type: ${typeLabel}`);
            await retryAction(async () => {
                await page.locator(`[debug-id="${typeId}"]`).first().click({ timeout: 5000 });
            }, `Select type: ${typeLabel}`);
            await delay(page, 2000);

            // Select Free.
            console.log('Selecting mode: Free');
            await retryAction(async () => {
                await page.getByText('Free', { exact: true }).first().click({ timeout: 5000 });
            }, 'Select Free mode');
            await delay(page, 2000);

            // Tick required declarations.
            console.log('Checking policy declarations...');
            const declarationCheckboxes = [
                '[debug-id="guidelines-checkbox"]',
                '[debug-id="export-laws-checkbox"]'
            ];
            for (const sel of declarationCheckboxes) {
                const checkbox = page.locator(sel);
                if (await checkbox.isVisible().catch(() => false)) {
                    const input = checkbox.locator('input[type="checkbox"]');
                    if (!(await input.isChecked().catch(() => false))) {
                        await checkbox.click().catch(() => { });
                        await delay(page, 1000);
                    }
                }
            }

            const otherBoxes = await page.locator('input[type="checkbox"]').all();
            for (const b of otherBoxes) {
                if (!(await b.isChecked().catch(() => false))) {
                    await b.check().catch(() => b.click().catch(() => { }));
                    await delay(page, 1000);
                }
            }

            await delay(page);

            // Click Create.
            console.log('Clicking "Create" submit button...');
            const submitBtn = page.locator('material-button[debug-id="create-app-button"] button').first();
            await retryAction(async () => {
                await submitBtn.scrollIntoViewIfNeeded({ timeout: 5000 });
                await submitBtn.click({ timeout: 10000 });
            }, 'Final Create button');

            const createFailedToastCandidates = [
                page.locator('text=/Your app couldn.?t be created/i').first(),
                page.locator('span[aria-live="polite"]').filter({
                    hasText: /Your app couldn.?t be created/i
                }).first(),
                page.locator('[role="alert"]').filter({
                    hasText: /Your app couldn.?t be created/i
                }).first()
            ];

            let createOutcome = '';
            const createOutcomeDeadline = Date.now() + 120000;
            while (Date.now() < createOutcomeDeadline) {
                const currentUrl = page.url();
                if (/\/app\/\d+/.test(currentUrl)) {
                    createOutcome = 'success';
                    break;
                }

                let createFailedVisible = false;
                for (const candidate of createFailedToastCandidates) {
                    if (await candidate.isVisible().catch(() => false)) {
                        createFailedVisible = true;
                        break;
                    }
                }
                if (createFailedVisible) {
                    createOutcome = 'failed';
                    break;
                }

                await page.waitForTimeout(500);
            }

            if (createOutcome === 'failed') {
                console.log('[CREATE] Detected failure message: "Your app couldn\'t be created".');
                console.log('[CREATE] Waiting 5 seconds then marking this row as FAILED...');
                await delay(page, 5000);
                statusManager.updateTaskStatus(task, STATUS_FAILED);

                const createFailedError = new Error(
                    'Create blocked by Play Console: Your app couldn\'t be created.'
                );
                createFailedError.code = 'CREATE_FAILED_TOAST';
                throw createFailedError;
            }

            if (createOutcome !== 'success') {
                const currentUrl = page.url();
                const visibleErrors = await page.locator(
                    '[role="alert"], .error, .errors, .warning, .mdc-snackbar, text=/error|invalid|already|unable|failed/i'
                ).allInnerTexts().catch(() => []);
                const compactErrors = visibleErrors
                    .map(text => String(text || '').trim())
                    .filter(Boolean)
                    .slice(0, 3)
                    .join(' | ');
                throw new Error(
                    'Create app did not finish in time. ' +
                    `Current URL: ${currentUrl}. ` +
                    `Visible errors: ${compactErrors || 'none'}. ` +
                    'Likely still blocked on page.'
                );
            }

            await page.waitForLoadState('domcontentloaded', { timeout: 30000 }).catch(() => { });
            console.log('App created successfully, navigating to dashboard...');
            await delay(page, 8000);

            // Extract app base path.
            appBasePath = extractAppBasePathFromCurrentUrl();
            console.log('App path:', appBasePath);

            // Package exists in Play Console now; persist PARTIAL for safe resume.
            statusManager.ensureTaskProgressAtLeast(task, PROGRESS_STEP_APP_CREATED);
            statusManager.updateTaskStatus(task, STATUS_PARTIAL);
        }

        // --- Helper functions inside this run ---
        async function goToAppContent() {
            await page.goto(appBasePath + '/app-content/overview', { timeout: 90000, waitUntil: 'domcontentloaded' });
            await delay(page, 8000);
            const tab = page.locator('div[role="tab"]:has-text("Need attention")');
            if (await tab.isVisible().catch(() => false)) {
                await tab.click().catch(() => { });
                await delay(page, 3000);
            }
        }

        async function clickStartDeclaration(sectionTitle) {
            const buttonByAria = page.locator(
                `button[aria-label="Start ${sectionTitle} declaration"], button[aria-label*="Start ${sectionTitle} declaration"]`
            ).first();
            const buttonByCard = page.locator(
                `xpath=//*[normalize-space(text())="${sectionTitle}"]/ancestor::*[.//button[contains(normalize-space(.), "Start declaration")]][1]//button[contains(normalize-space(.), "Start declaration")]`
            ).first();
            await retryAction(async () => {
                if (await buttonByAria.isVisible().catch(() => false)) {
                    await buttonByAria.scrollIntoViewIfNeeded({ timeout: 5000 });
                    await buttonByAria.click({ timeout: 10000 });
                    return;
                }

                await buttonByCard.waitFor({ state: 'visible', timeout: 12000 });
                await buttonByCard.scrollIntoViewIfNeeded({ timeout: 5000 });
                await buttonByCard.click({ timeout: 10000 });
            }, `Start ${sectionTitle} declaration`);
            await delay(page, 8000);
        }

        async function selectRadio(textRegex) {
            await retryAction(async () => {
                const radio = page.locator('material-radio, [role="radio"]').filter({ hasText: textRegex }).first();
                await radio.scrollIntoViewIfNeeded({ timeout: 10000 });
                await radio.click({ timeout: 10000 });
                await delay(page, 1500);
            }, `Select radio containing "${textRegex}"`);
        }

        async function selectCheckbox(textRegex) {
            await retryAction(async () => {
                const checkbox = page.locator('material-checkbox, [role="checkbox"]').filter({ hasText: textRegex }).first();
                await checkbox.scrollIntoViewIfNeeded({ timeout: 10000 });
                const input = checkbox.locator('input[type="checkbox"]');
                const isChecked = await input.isChecked().catch(() => false);
                if (!isChecked) {
                    await checkbox.click({ timeout: 10000 });
                    await delay(page, 1000);
                }
            }, `Select checkbox containing "${textRegex}"`);
        }

        async function clickMainButton(text) {
            console.log(`Looking for and clicking button: ${text}...`);
            await retryAction(async () => {
                const selectors = [
                    '[debug-id="main-button"]',
                    `button:has-text("${text}")`,
                    `div[role="button"]:has-text("${text}")`,
                    `material-button:has-text("${text}")`
                ];
                let foundBtn = null;
                for (const sel of selectors) {
                    const loc = page.locator(sel).first();
                    if (await loc.isVisible().catch(() => false)) {
                        if (sel === '[debug-id="main-button"]') {
                            const content = await loc.innerText().catch(() => '');
                            if (!content.toLowerCase().includes(text.toLowerCase())) continue;
                        }
                        foundBtn = loc;
                        break;
                    }
                }
                if (!foundBtn) throw new Error(`Main button with text "${text}" not found`);

                const isDisabled = await foundBtn.evaluate(el =>
                    el.hasAttribute('disabled') || el.classList.contains('mdc-button--disabled') || el.getAttribute('aria-disabled') === 'true'
                ).catch(() => false);
                if (isDisabled) throw new Error(`Button "${text}" is disabled`);

                await foundBtn.scrollIntoViewIfNeeded({ timeout: 5000 });
                await foundBtn.click({ timeout: 10000 });
                await delay(page, 3000);
            }, `Click "${text}" button`);
        }

        async function getDeclarationCardState(sectionTitle) {
            const startByAria = await page.locator(
                `button[aria-label="Start ${sectionTitle} declaration"], button[aria-label*="Start ${sectionTitle} declaration"]`
            ).count().catch(() => 0);
            if (startByAria > 0) {
                return 'pending';
            }

            const startByCard = await page.locator(
                `xpath=//*[normalize-space(text())="${sectionTitle}"]/ancestor::*[.//button[contains(normalize-space(.), "Start declaration")]][1]//button[contains(normalize-space(.), "Start declaration")]`
            ).count().catch(() => 0);
            if (startByCard > 0) {
                return 'pending';
            }

            const titleCount = await page.locator(
                `xpath=//*[normalize-space(text())="${sectionTitle}"]`
            ).count().catch(() => 0);
            if (titleCount > 0) {
                return 'done';
            }

            // Not shown on current page: unknown, do NOT auto-advance progress.
            return 'unknown';
        }

        async function resolveDeclarationCardState(sectionTitle) {
            let state = await getDeclarationCardState(sectionTitle);
            if (state !== 'unknown') {
                return state;
            }

            // Retry once after a lightweight refresh to avoid false "unknown" due to UI lag.
            await delay(page, 1500);
            await goToAppContent();
            state = await getDeclarationCardState(sectionTitle);
            return state;
        }

        async function syncProgressStepFromUiIfPossible() {
            if (task.status !== STATUS_PARTIAL) {
                return;
            }

            await goToAppContent();

            const checkpoints = [
                { title: 'Ads', step: PROGRESS_STEP_ADS_DONE },
                { title: 'App access', step: PROGRESS_STEP_APP_ACCESS_DONE },
                { title: 'Target audience and content', step: PROGRESS_STEP_AUDIENCE_DONE },
                { title: 'Advertising ID', step: PROGRESS_STEP_AD_ID_DONE },
                { title: 'Government apps', step: PROGRESS_STEP_GOV_DONE },
                { title: 'Financial features', step: PROGRESS_STEP_FINANCE_DONE },
                { title: 'Health apps', step: PROGRESS_STEP_HEALTH_DONE }
            ];

            let detectedStep = '';
            for (const checkpoint of checkpoints) {
                const state = await resolveDeclarationCardState(checkpoint.title);
                if (state === 'pending') {
                    break;
                }
                if (state === 'done') {
                    detectedStep = checkpoint.step;
                    continue;
                }
                break;
            }

            if (detectedStep && progressRank(detectedStep) > progressRank(task.progressStep)) {
                statusManager.ensureTaskProgressAtLeast(task, detectedStep);
                statusManager.updateTaskStatus(task, STATUS_PARTIAL);
                console.log(`[RESUME] UI checkpoint detected, progress corrected to ${detectedStep}.`);
            }
        }

        // --- Execute declarations flow with fine-grained checkpoint resume ---
        const shouldRunStep = (doneStep) => progressRank(task.progressStep) < progressRank(doneStep);
        const markStepDone = (doneStep) => {
            statusManager.updateTaskProgress(task, doneStep);
            statusManager.updateTaskStatus(task, STATUS_PARTIAL);
        };
        const ensureKnownDeclarationState = async (sectionTitle) => {
            const state = await resolveDeclarationCardState(sectionTitle);
            if (state === 'unknown') {
                throw new Error(
                    `[DECLARATION] "${sectionTitle}" state is unknown. ` +
                    'Stop this row to avoid false progress advancement.'
                );
            }
            return state;
        };

        await syncProgressStepFromUiIfPossible();

        if (shouldRunStep(PROGRESS_STEP_ADS_DONE)) {
            await goToAppContent();
            const adsState = await ensureKnownDeclarationState('Ads');
            if (adsState === 'pending') {
                console.log('Executing declaration 1/7: Ads...');
                await clickStartDeclaration('Ads');
                await selectRadio(/^No/);
                await clickMainButton('Save');
                await waitSaved(page);
            } else {
                console.log('[RESUME] Ads already done on UI, skip filling.');
            }
            markStepDone(PROGRESS_STEP_ADS_DONE);
        } else {
            console.log(`[RESUME] Skip Ads (already >= ${task.progressStep}).`);
        }

        if (shouldRunStep(PROGRESS_STEP_APP_ACCESS_DONE)) {
            await goToAppContent();
            const appAccessState = await ensureKnownDeclarationState('App access');
            if (appAccessState === 'pending') {
                console.log('Executing declaration 2/7: App access...');
                await clickStartDeclaration('App access');
                await selectRadio('All functionality in my app is available without any access restrictions');
                await clickMainButton('Save');
                await waitSaved(page);
            } else {
                console.log('[RESUME] App access already done on UI, skip filling.');
            }
            markStepDone(PROGRESS_STEP_APP_ACCESS_DONE);
        } else {
            console.log(`[RESUME] Skip App access (already >= ${task.progressStep}).`);
        }

        if (shouldRunStep(PROGRESS_STEP_AUDIENCE_DONE)) {
            await goToAppContent();
            const audienceState = await ensureKnownDeclarationState('Target audience and content');
            if (audienceState === 'pending') {
                console.log('Executing declaration 3/7: Target audience and content...');
                await clickStartDeclaration('Target audience and content');
                await selectCheckbox('13-15');
                await selectCheckbox('16-17');
                await selectCheckbox('18 and over');
                await clickMainButton('Next');

                let stepsLeft = 5;
                while (stepsLeft-- > 0) {
                    const appealText = "Could your store listing unintentionally appeal to children?";
                    if (await page.locator(`text=${appealText}`).first().isVisible().catch(() => false)) {
                        console.log('Handling child appeal question...');
                        await selectRadio(/^No/);
                    }
                    const isSaveVisible = await page.locator('button:has-text("Save"), [debug-id="main-button"]:has-text("Save")').first().isVisible().catch(() => false);
                    if (isSaveVisible) {
                        await clickMainButton('Save');
                        break;
                    }
                    const isNextVisible = await page.locator('button:has-text("Next"), [debug-id="main-button"]:has-text("Next")').first().isVisible().catch(() => false);
                    if (isNextVisible) {
                        await clickMainButton('Next');
                    } else {
                        console.log('Warning: Target Audience path reached end or got stuck.');
                        break;
                    }
                }
                await waitSaved(page);
            } else {
                console.log('[RESUME] Target audience and content already done on UI, skip filling.');
            }
            markStepDone(PROGRESS_STEP_AUDIENCE_DONE);
        } else {
            console.log(`[RESUME] Skip Target audience and content (already >= ${task.progressStep}).`);
        }

        if (shouldRunStep(PROGRESS_STEP_AD_ID_DONE)) {
            await goToAppContent();
            const adIdState = await ensureKnownDeclarationState('Advertising ID');
            if (adIdState === 'pending') {
                console.log('Executing declaration 4/7: Advertising ID...');
                await clickStartDeclaration('Advertising ID');
                await selectRadio(/^Yes/);
                await selectCheckbox('Developer communications');
                await clickMainButton('Save');
                await waitSaved(page);
                await page.evaluate(() => window.scrollTo(0, 0));
                await delay(page, 2000);
            } else {
                console.log('[RESUME] Advertising ID already done on UI, skip filling.');
            }
            markStepDone(PROGRESS_STEP_AD_ID_DONE);
        } else {
            console.log(`[RESUME] Skip Advertising ID (already >= ${task.progressStep}).`);
        }

        if (shouldRunStep(PROGRESS_STEP_GOV_DONE)) {
            await goToAppContent();
            const govState = await ensureKnownDeclarationState('Government apps');
            if (govState === 'pending') {
                console.log('Executing declaration 5/7: Government apps...');
                await clickStartDeclaration('Government apps');
                await selectRadio(/^No/);
                await clickMainButton('Save');
                await waitSaved(page);
            } else {
                console.log('[RESUME] Government apps already done on UI, skip filling.');
            }
            markStepDone(PROGRESS_STEP_GOV_DONE);
        } else {
            console.log(`[RESUME] Skip Government apps (already >= ${task.progressStep}).`);
        }

        if (shouldRunStep(PROGRESS_STEP_FINANCE_DONE)) {
            await goToAppContent();
            const financeState = await ensureKnownDeclarationState('Financial features');
            if (financeState === 'pending') {
                console.log('Executing declaration 6/7: Financial features...');
                await clickStartDeclaration('Financial features');
                await selectCheckbox("My app doesn't provide any financial features");
                await clickMainButton('Next');
                await clickMainButton('Save');
                await waitSaved(page);
            } else {
                console.log('[RESUME] Financial features already done on UI, skip filling.');
            }
            markStepDone(PROGRESS_STEP_FINANCE_DONE);
        } else {
            console.log(`[RESUME] Skip Financial features (already >= ${task.progressStep}).`);
        }

        if (shouldRunStep(PROGRESS_STEP_HEALTH_DONE)) {
            await goToAppContent();
            const healthState = await ensureKnownDeclarationState('Health apps');
            if (healthState === 'pending') {
                console.log('Executing declaration 7/7: Health apps...');
                await clickStartDeclaration('Health apps');
                await selectCheckbox('My app does not have any health features');
                await clickMainButton('Next');
                await clickMainButton('Save');
                await waitSaved(page);
            } else {
                console.log('[RESUME] Health apps already done on UI, skip filling.');
            }
            markStepDone(PROGRESS_STEP_HEALTH_DONE);
        } else {
            console.log(`[RESUME] Skip Health apps (already >= ${task.progressStep}).`);
        }

        if (shouldRunStep(PROGRESS_STEP_COUNTRY_DONE)) {
            // 8. Test and release
            console.log('Navigating to "Test and release"...');
            const testAndReleaseLink = page.locator(
                'a[href*="/test-and-release"], a.item-link:has(.item-label:has-text("Test and release"))'
            ).first();
            await retryAction(async () => {
                await testAndReleaseLink.waitFor({ state: 'visible', timeout: 20000 });
                await testAndReleaseLink.scrollIntoViewIfNeeded({ timeout: 5000 });
                await testAndReleaseLink.click({ timeout: 10000 });
            }, 'Click Test and release');
            try {
                await page.waitForURL(/\/test-and-release(?:\/|$)/, { timeout: 60000 });
            } catch (_) {
                console.log('Warning: URL did not switch to /test-and-release in time, continuing...');
            }
            await delay(page, 5000);

            // 9. Production
            console.log('Navigating to "Production"...');
            const productionLink = page.locator(
                'a[href*="/tracks/production"], a.item-link:has(.item-label:has-text("Production"))'
            ).first();
            await retryAction(async () => {
                await productionLink.waitFor({ state: 'visible', timeout: 20000 });
                await productionLink.scrollIntoViewIfNeeded({ timeout: 5000 });
                await productionLink.click({ timeout: 10000 });
            }, 'Click Production');
            try {
                await page.waitForURL(/\/tracks\/production(?:\/|$)/, { timeout: 60000 });
            } catch (_) {
                console.log('Warning: URL did not switch to /tracks/production in time, continuing...');
            }
            await delay(page, 5000);

            // 10. Countries / regions tab
            console.log('Opening "Countries / regions" tab...');
            const countriesRegionsTab = page.locator(
                '[role="tab"]:has-text("Countries / regions"), a:has-text("Countries / regions"), button:has-text("Countries / regions")'
            ).first();
            await retryAction(async () => {
                await countriesRegionsTab.waitFor({ state: 'visible', timeout: 20000 });
                await countriesRegionsTab.scrollIntoViewIfNeeded({ timeout: 5000 });
                await countriesRegionsTab.click({ timeout: 10000 });
            }, 'Click Countries / regions tab');
            await delay(page, 3000);

            // 11. Add countries / regions
            console.log('Clicking "Add countries / regions"...');
            const addCountriesBtn = page.locator(
                'button:has-text("Add countries / regions"), [role="button"]:has-text("Add countries / regions"), a:has-text("Add countries / regions"), .mdc-button:has-text("Add countries / regions")'
            ).first();
            await retryAction(async () => {
                await addCountriesBtn.waitFor({ state: 'visible', timeout: 20000 });
                await addCountriesBtn.scrollIntoViewIfNeeded({ timeout: 5000 });
                await addCountriesBtn.click({ timeout: 10000 });
            }, 'Click Add countries / regions');
            await delay(page, 3000);

            // 12. Select all countries/regions (prefer direct "Select all rows" checkbox)
            console.log('Selecting "Select all rows" checkbox...');
            const selectAllRowsCheckbox = page.locator(
                'mat-checkbox[aria-label="Select all rows"][role="checkbox"], [role="checkbox"][aria-label="Select all rows"]'
            ).first();
            const countryRegionHeaderRow = page.locator(
                'tr:has-text("Country / region"), [role="row"]:has-text("Country / region"), div:has-text("Country / region")'
            ).first();
            const countryRegionCheckboxInput = countryRegionHeaderRow.locator('input[type="checkbox"]').first();
            const countryRegionCheckboxTarget = countryRegionHeaderRow.locator('[role="checkbox"], .mdc-checkbox').first();

            await retryAction(async () => {
                const directVisible = await selectAllRowsCheckbox.isVisible().catch(() => false);
                if (directVisible) {
                    await selectAllRowsCheckbox.scrollIntoViewIfNeeded({ timeout: 3000 }).catch(() => { });
                    const directCheckedBefore = await selectAllRowsCheckbox.getAttribute('aria-checked').catch(() => null);
                    if (directCheckedBefore !== 'true') {
                        await selectAllRowsCheckbox.click({ timeout: 5000, force: true });
                    }

                    const directCheckedAfter = await selectAllRowsCheckbox.getAttribute('aria-checked').catch(() => null);
                    if (directCheckedAfter === 'true') {
                        return;
                    }
                }

                // Fallback: old row-based selectors if "Select all rows" is unavailable.
                await countryRegionHeaderRow.waitFor({ state: 'visible', timeout: 10000 });
                await countryRegionHeaderRow.scrollIntoViewIfNeeded({ timeout: 3000 }).catch(() => { });

                const checkedBefore = await countryRegionCheckboxInput.isChecked().catch(() => false);
                if (!checkedBefore) {
                    await countryRegionCheckboxInput.check({ force: true, timeout: 5000 }).catch(async () => {
                        if (await countryRegionCheckboxTarget.isVisible().catch(() => false)) {
                            await countryRegionCheckboxTarget.click({ timeout: 5000 });
                        } else {
                            await countryRegionHeaderRow.click({ timeout: 5000 });
                        }
                    });
                }

                const checkedAfter = await countryRegionCheckboxInput.isChecked().catch(() => false);
                const ariaCheckedAfter = await countryRegionCheckboxTarget.getAttribute('aria-checked').catch(() => null);
                if (!checkedAfter && ariaCheckedAfter !== 'true') {
                    throw new Error('Country / region checkbox is still not checked.');
                }
            }, 'Ensure Select all rows checkbox checked', 2);
            await delay(page, 1000);

            // 13. Save countries/regions selection
            console.log('Saving countries/regions selection...');
            await clickMainButton('Save');

            // 14. Verify success marker: Targeted (N)
            console.log('Verifying "Targeted (N)" appears...');
            const targetedBadge = page.locator('span.button-text, button, [role="button"]').filter({
                hasText: /Targeted\s*\(\d+\)/i
            }).first();
            await retryAction(async () => {
                await targetedBadge.waitFor({ state: 'visible', timeout: 90000 });
            }, 'Wait for Targeted (N)');
            await delay(page, 3000);

            markStepDone(PROGRESS_STEP_COUNTRY_DONE);
        } else {
            console.log(`[RESUME] Skip countries/regions publish step (already >= ${task.progressStep}).`);
        }

        statusManager.ensureTaskProgressAtLeast(task, PROGRESS_STEP_DONE);
        statusManager.updateTaskStatus(task, STATUS_DONE);

        const durationSeconds = Math.round((Date.now() - startTime) / 1000);
        console.log(`Finished: ${appName} (${packageName}), Duration: ${formatDuration(durationSeconds)}`);
        if (page && !page.isClosed()) {
            await page.close().catch(() => { });
        }
        await browser.close();
    } catch (err) {
        const currentUrl = (page && typeof page.url === 'function') ? page.url() : '';
        if (currentUrl && /\/app\/\d+/.test(currentUrl) && task.status !== STATUS_DONE) {
            try {
                statusManager.ensureTaskProgressAtLeast(task, PROGRESS_STEP_APP_CREATED);
                statusManager.updateTaskStatus(task, STATUS_PARTIAL);
            } catch (_) {
                // Ignore status write failures in error path; keep original error as primary signal.
            }
        }

        console.error(`[FATAL ERROR] In iteration: ${err.message}`);
        if (page && !page.isClosed()) {
            await page.close().catch(() => { });
        }
        if (browser) await browser.close().catch(() => { });
        throw err;
    }
}

(async () => {
    const { inputFileArg } = parseCliArgs();
    const { appListUrl, configPath } = loadDeveloperConsoleAppListUrl();
    const inputFilePath = resolveInputExcelFile(inputFileArg);
    const runtimeInput = resolveRuntimeInput(inputFilePath);
    const { tasks, sheetName, statusManager } = loadTasksFromExcel(
        runtimeInput.runtimeWorkbookPath,
        runtimeInput.isCsvInput
            ? {
                persistMode: 'csv-state',
                csvStateFilePath: runtimeInput.stateFilePath,
                sourceFilePath: runtimeInput.sourceFilePath
            }
            : { persistMode: 'workbook' }
    );
    const selectedTasks = tasks.filter(task => task.status !== STATUS_DONE);
    const skippedDone = tasks.length - selectedTasks.length;

    console.log(`Developer console URL loaded from: ${configPath}`);
    console.log(`Loaded ${tasks.length} rows from input file: ${path.basename(inputFilePath)} (sheet: ${sheetName})`);
    if (runtimeInput.isCsvInput) {
        console.log(`[CSV] Source file: ${runtimeInput.sourceFilePath}`);
        console.log(`[CSV] Runtime state file: ${runtimeInput.stateFilePath}`);
        console.log('[CSV] No Excel file will be generated in source directory.');
    }
    if (skippedDone > 0) {
        console.log(`[STATUS] Skipping ${skippedDone} row(s) already marked DONE.`);
    }
    console.log(`Will execute ${selectedTasks.length} iteration(s), excluding DONE rows.`);

    if (!selectedTasks.length) {
        console.log('No pending rows to execute. All rows are DONE.');
        return;
    }

    const runStats = {
        totalLoaded: tasks.length,
        planned: selectedTasks.length,
        success: 0,
        failed: []
    };

    for (let i = 0; i < selectedTasks.length; i++) {
        const task = selectedTasks[i];
        console.log(
            `Iteration ${i + 1}/${selectedTasks.length} | Excel row ${task.rowNumber} | ` +
            `App="${task.appName}" | Package="${task.packageName}" | ` +
            `Status="${task.status || 'EMPTY'}" | Progress="${task.progressStep || 'EMPTY'}"`
        );
        try {
            await runOnce(task, appListUrl, statusManager);
            runStats.success += 1;
        } catch (e) {
            const isCreateFailedToast = String(e && e.code || '') === 'CREATE_FAILED_TOAST';
            if (isCreateFailedToast && task.status !== STATUS_FAILED) {
                try {
                    statusManager.updateTaskStatus(task, STATUS_FAILED);
                } catch (_) {
                    // Keep original error flow.
                }
            }

            runStats.failed.push({
                appName: task.appName,
                packageName: task.packageName,
                reason: isCreateFailedToast
                    ? "Your app couldn't be created"
                    : String((e && e.message) || 'Unknown error')
            });

            console.error(`Iteration ${i + 1} failed but continuing...`);
            if (isCdpConnectionError(e)) {
                throw e;
            }
        }

        if (i < selectedTasks.length - 1) {
            const wait = 30000 + Math.random() * 40000;
            console.log('Wait until next iteration:', formatDuration(Math.round(wait / 1000)));
            await new Promise(r => setTimeout(r, wait));
        }
    }

    const failedCount = runStats.failed.length;
    const failedNames = runStats.failed.map(item => `${item.appName} (${item.packageName})`);
    const summaryLines = [
        `本次总读取: ${runStats.totalLoaded} 条`,
        `本次计划执行: ${runStats.planned} 条`,
        `成功: ${runStats.success} 条`,
        `失败: ${failedCount} 条`
    ];
    if (failedCount > 0) {
        summaryLines.push('失败应用:');
        for (const name of failedNames) {
            summaryLines.push(`- ${name}`);
        }
    }
    const summaryText = summaryLines.join('\n');

    console.log('================ Run Summary ================');
    console.log(summaryText);
    console.log('=============================================');
    showWindowsSummaryPopup(summaryText);
})().catch(err => {
    console.error(`[INIT ERROR] ${err.message}`);
    process.exit(1);
});

