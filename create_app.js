const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const os = require('os');
const { execFileSync } = require('child_process');
const XLSX = require('xlsx');
const { chromium } = require('playwright');

// Default delay (ms). Increase if VPS is slow.
const DELAY = 3000;

const APP_NAME_HEADER_CANDIDATES = new Set([
    '\u5e94\u7528\u540d\u79f0',
    '\u5e94\u7528\u540d',
    'appname',
    'applicationname'
]);

const PACKAGE_NAME_HEADER_CANDIDATES = new Set([
    '\u5e94\u7528\u5305\u540d',
    '\u5305\u540d',
    'apppackagename',
    'packagename',
    'applicationid'
]);

const PACKAGE_NAME_REGEX = /^[a-z][a-z0-9_]*(\.[a-z][a-z0-9_]*)+$/;
const CONFIG_DIR_ENV = 'APP_CREATE_CONFIG_DIR';
const CDP_ENDPOINT_ENV = 'APP_CREATE_CDP_ENDPOINT';
const LOGIN_WAIT_SECONDS_ENV = 'APP_CREATE_LOGIN_WAIT_SECONDS';
const DEVELOPER_URL_ENV = 'APP_CREATE_DEVELOPER_URL';
const DEVELOPER_ID_ENV = 'APP_CREATE_DEVELOPER_ID';
const RUN_SUMMARY_PATH_ENV = 'APP_CREATE_RUN_SUMMARY_PATH';
const WEB_USERNAME_ENV = 'APP_CREATE_WEB_USERNAME';
const WEB_PASSWORD_ENV = 'APP_CREATE_WEB_PASSWORD';
const CONTACT_EMAIL_ENV = 'APP_CREATE_CONTACT_EMAIL';
const DEFAULT_LOGIN_WAIT_SECONDS = 900;
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
const PROGRESS_STEP_PRIVACY_DONE = 'PRIVACY_DONE';
const PROGRESS_STEP_RATING_DONE = 'RATING_DONE';
const PROGRESS_STEP_SAFETY_DONE = 'SAFETY_DONE';
const PROGRESS_STEP_STORE_SETTINGS_DONE = 'STORE_SETTINGS_DONE';
const PROGRESS_STEP_STORE_LISTING_DONE = 'STORE_LISTING_DONE';
const PROGRESS_STEP_RELEASE_DONE = 'RELEASE_DONE';
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
    PROGRESS_STEP_PRIVACY_DONE,
    PROGRESS_STEP_RATING_DONE,
    PROGRESS_STEP_SAFETY_DONE,
    PROGRESS_STEP_STORE_SETTINGS_DONE,
    PROGRESS_STEP_STORE_LISTING_DONE,
    PROGRESS_STEP_RELEASE_DONE,
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
    return {
        sourceFilePath: inputFilePath,
        runtimeWorkbookPath: inputFilePath,
        isCsvInput: ext === '.csv'
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

function writeRunSummaryFile(summaryPayload) {
    const configuredPath = String(process.env[RUN_SUMMARY_PATH_ENV] || '').trim();
    if (!configuredPath) {
        return;
    }
    try {
        const resolvedPath = path.resolve(configuredPath);
        fs.mkdirSync(path.dirname(resolvedPath), { recursive: true });
        fs.writeFileSync(resolvedPath, JSON.stringify(summaryPayload, null, 2), 'utf8');
        console.log(`[SUMMARY] Run summary saved to: ${resolvedPath}`);
    } catch (err) {
        console.log(`[WARN] Cannot save run summary file. ${err.message}`);
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

function createNoopStatusManager() {
    return {
        saveWorkbook() { },
        updateTaskStatus(task, status) {
            task.status = normalizeStatusValue(status);
        },
        getTaskStatus(task) {
            return normalizeStatusValue(task.status);
        },
        updateTaskProgress(task, step) {
            task.progressStep = normalizeProgressStep(step);
        },
        ensureTaskProgressAtLeast(task, step) {
            const current = normalizeProgressStep(task.progressStep);
            const next = normalizeProgressStep(step);
            if (progressRank(next) > progressRank(current)) {
                task.progressStep = next;
            }
        },
        getTaskProgress(task) {
            return normalizeProgressStep(task.progressStep);
        },
        statusColumnIndex: -1,
        progressColumnIndex: -1,
        sheetName: ''
    };
}

function loadTasksFromExcel(filePath, sourceFilePath = filePath) {
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
    const sourceExt = path.extname(String(sourceFilePath || filePath)).toLowerCase();
    const isCsvSource = sourceExt === '.csv';

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
            'Need columns like: App Name / App Package Name (or 应用名称 / 应用包名).'
        );
    }

    let statusManager;
    let csvRows = null;
    if (isCsvSource) {
        const stateFilePath = buildCsvStateFilePath(sourceFilePath);
        const rows = loadCsvStateRows(stateFilePath);
        csvRows = rows;
        statusManager = createCsvStatusManager({
            stateFilePath,
            sourceFilePath,
            rows
        });
    } else {
        const refRange = XLSX.utils.decode_range(sheet['!ref']);
        let nextColumnIndex = refRange.e.c + 1;
        let changed = false;

        if (statusColumnIndex < 0) {
            statusColumnIndex = nextColumnIndex++;
            setCellText(sheet, headerRowIndex, statusColumnIndex, 'status');
            changed = true;
            console.log('[STATUS] Added "status" column automatically.');
        }
        if (progressColumnIndex < 0) {
            progressColumnIndex = nextColumnIndex++;
            setCellText(sheet, headerRowIndex, progressColumnIndex, 'progress_step');
            changed = true;
            console.log('[PROGRESS] Added "progress_step" column automatically.');
        }
        if (changed) {
            const nextRange = {
                s: { r: refRange.s.r, c: refRange.s.c },
                e: { r: Math.max(refRange.e.r, headerRowIndex), c: nextColumnIndex - 1 }
            };
            sheet['!ref'] = XLSX.utils.encode_range(nextRange);
        }

        statusManager = createExcelStatusManager({
            filePath,
            workbook,
            sheetName: firstSheetName,
            sheet,
            statusColumnIndex,
            progressColumnIndex
        });
        if (changed) {
            statusManager.saveWorkbook();
        }
    }
    statusManager.sheetName = firstSheetName;

    const tasks = [];

    for (let r = headerRowIndex + 1; r <= range.e.r; r++) {
        const excelRowNumber = r + 1;
        const appName = getCellText(sheet, r, appNameColumnIndex);
        const rawPackageName = getCellText(sheet, r, packageNameColumnIndex);
        const packageName = rawPackageName.toLowerCase();

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

        let status = '';
        let progressStep = '';
        if (isCsvSource) {
            const key = getCsvStateKey(packageName);
            const stateRow = (csvRows || {})[key] || {};
            status = normalizeStatusValue(stateRow.status);
            progressStep = normalizeProgressStep(stateRow.progressStep);
        } else {
            status = normalizeStatusValue(getCellText(sheet, r, statusColumnIndex));
            progressStep = normalizeProgressStep(getCellText(sheet, r, progressColumnIndex));
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

function getManualLoginWaitMs() {
    const raw = String(process.env[LOGIN_WAIT_SECONDS_ENV] || '').trim();
    if (!raw) {
        return DEFAULT_LOGIN_WAIT_SECONDS * 1000;
    }

    const parsed = Number(raw);
    if (!Number.isFinite(parsed) || parsed < 1) {
        console.log(
            `[LOGIN] Invalid ${LOGIN_WAIT_SECONDS_ENV}="${raw}", ` +
            `fallback to default ${DEFAULT_LOGIN_WAIT_SECONDS}s.`
        );
        return DEFAULT_LOGIN_WAIT_SECONDS * 1000;
    }

    return Math.round(parsed * 1000);
}

async function delay(page, ms = DELAY) {
    await page.waitForTimeout(ms);
}

function randomInt(min, max) {
    const lo = Math.min(min, max);
    const hi = Math.max(min, max);
    return Math.floor(Math.random() * (hi - lo + 1)) + lo;
}

async function randomDelay(page, minMs, maxMs) {
    await delay(page, randomInt(minMs, maxMs));
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

async function clickLocatorRobust(locator, label = 'element', timeoutMs = 10000) {
    await locator.waitFor({ state: 'visible', timeout: timeoutMs });
    await locator.scrollIntoViewIfNeeded({ timeout: 5000 }).catch(() => { });
    try {
        await locator.click({ timeout: timeoutMs });
        return;
    } catch (normalClickError) {
        try {
            await locator.click({ timeout: timeoutMs, force: true });
            return;
        } catch (forceClickError) {
            try {
                await locator.evaluate((el) => {
                    if (typeof el.click === 'function') {
                        el.click();
                    } else {
                        el.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true }));
                    }
                });
                return;
            } catch (evaluateClickError) {
                throw new Error(
                    `${label} click failed. ` +
                    `normal=${normalClickError.message}; ` +
                    `force=${forceClickError.message}; ` +
                    `evaluate=${evaluateClickError.message}`
                );
            }
        }
    }
}

async function isLocatorDisabled(locator) {
    const visible = await locator.isVisible().catch(() => false);
    if (!visible) return true;
    return await locator.evaluate((el) => {
        return el.hasAttribute('disabled') ||
            el.classList.contains('mdc-button--disabled') ||
            el.getAttribute('aria-disabled') === 'true';
    }).catch(() => true);
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

function getRequiredEnvValue(envName, displayLabel) {
    const value = String(process.env[envName] || '').trim();
    if (!value) {
        throw new Error(
            `Missing required input: ${displayLabel}. ` +
            `Please fill "${displayLabel}" in launcher form.`
        );
    }
    return value;
}

function getRuntimeOptions() {
    const webUsername = getRequiredEnvValue(WEB_USERNAME_ENV, 'web_username');
    const webPassword = getRequiredEnvValue(WEB_PASSWORD_ENV, 'web_password');
    const contactEmail = getRequiredEnvValue(CONTACT_EMAIL_ENV, 'contact_email');

    if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(contactEmail)) {
        throw new Error(`Invalid contact_email format: "${contactEmail}"`);
    }

    return {
        webUsername,
        webPassword,
        contactEmail
    };
}

function scoreConsolePageForReuse(url, appListUrl) {
    const currentUrl = String(url || '');
    if (!currentUrl) return -1;

    if (currentUrl.startsWith(appListUrl)) return 100;
    if (/^https:\/\/play\.google\.com\/console\/u\/\d+\/developers\/\d+\/app-list(?:[/?#]|$)/i.test(currentUrl)) {
        return 95;
    }
    if (/^https:\/\/play\.google\.com\/console\/u\/\d+\/developers\/\d+\/create-new-app(?:[/?#]|$)/i.test(currentUrl)) {
        return 85;
    }
    if (/^https:\/\/play\.google\.com\/console\/u\/\d+\/developers\/\d+\/app\/\d+/i.test(currentUrl)) {
        return 75;
    }
    if (/^https:\/\/play\.google\.com\/console/i.test(currentUrl)) {
        return 60;
    }

    return -1;
}

async function acquireWorkingConsolePage(context, appListUrl) {
    const pages = context.pages().filter(p => !p.isClosed());
    let bestPage = null;
    let bestScore = -1;

    for (const candidate of pages) {
        const score = scoreConsolePageForReuse(candidate.url(), appListUrl);
        if (score > bestScore) {
            bestScore = score;
            bestPage = candidate;
        }
    }

    if (bestPage) {
        return { page: bestPage, source: 'reused-existing' };
    }

    const newPage = await context.newPage();
    return { page: newPage, source: 'created-new' };
}

function buildSiteSlug(appName) {
    const base = String(appName || '')
        .replace(/[^A-Za-z0-9]+/g, '')
        .slice(0, 28);
    const safeBase = base || 'appcreate';
    const digitsLength = 2 + Math.floor(Math.random() * 2); // 2-3 digits
    let suffix = '';
    for (let i = 0; i < digitsLength; i += 1) {
        suffix += String(1 + Math.floor(Math.random() * 9)); // 1-9 only
    }

    const candidate = `${safeBase}${suffix}`;
    return candidate.slice(0, 31);
}

function escapeRegExp(value) {
    return String(value || '').replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function safeFileToken(value) {
    return String(value || 'app')
        .replace(/[^A-Za-z0-9._-]+/g, '_')
        .replace(/^_+|_+$/g, '')
        .slice(0, 80) || 'app';
}

function escapePowerShellSingleQuotedString(value) {
    return String(value || '').replace(/'/g, "''");
}

function extractZipToDirectory(zipPath, outputDir) {
    fs.rmSync(outputDir, { recursive: true, force: true });
    fs.mkdirSync(outputDir, { recursive: true });
    const command = [
        `$zip='${escapePowerShellSingleQuotedString(zipPath)}'`,
        `$out='${escapePowerShellSingleQuotedString(outputDir)}'`,
        'Expand-Archive -LiteralPath $zip -DestinationPath $out -Force'
    ].join('; ');
    execFileSync(
        'powershell.exe',
        ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-Command', command],
        { stdio: 'pipe' }
    );
}

function collectFilesRecursive(rootDir) {
    const files = [];
    const walk = (dir) => {
        for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
            const fullPath = path.join(dir, entry.name);
            if (entry.isDirectory()) {
                walk(fullPath);
            } else if (entry.isFile()) {
                files.push(fullPath);
            }
        }
    };

    if (fs.existsSync(rootDir)) {
        walk(rootDir);
    }
    return files;
}

function findImagesNamed(rootDir, names = ['1', '2', '3']) {
    const files = collectFilesRecursive(rootDir);
    const imageExts = new Set(['.png', '.jpg', '.jpeg', '.webp']);
    return names.map((name) => {
        const match = files.find((file) => {
            const parsed = path.parse(file);
            return parsed.name === name && imageExts.has(parsed.ext.toLowerCase());
        });
        if (!match) {
            throw new Error(`Image file named "${name}" not found in extracted ZIP.`);
        }
        return match;
    });
}

function normalizeAppGenieCardText(value) {
    return String(value || '')
        .replace(/\r\n/g, '\n')
        .replace(/[ \t]+\n/g, '\n')
        .trim();
}

function ensureTrailingSpace(value) {
    const text = normalizeAppGenieCardText(value);
    return text.endsWith(' ') ? text : `${text} `;
}

function ensureTrailingSpaceWithinLimit(value, maxLength) {
    let text = normalizeAppGenieCardText(value).replace(/\s+/g, ' ');
    const hardLimit = Math.max(1, Number(maxLength) || text.length + 1);
    if (text.length >= hardLimit) {
        text = text.slice(0, hardLimit - 1).replace(/\s+\S*$/, '').trim();
    }
    return ensureTrailingSpace(text).slice(0, hardLimit);
}

function shortenShortDescription(raw, maxLength = 80) {
    let text = normalizeAppGenieCardText(raw).replace(/\s+/g, ' ');
    if (text.length <= maxLength - 1) {
        return ensureTrailingSpaceWithinLimit(text, maxLength);
    }

    const replacements = [
        [/\bbrain-teasing\b/ig, 'clever'],
        [/\blogic puzzle game\b/ig, 'logic puzzle'],
        [/\bwhere you\b/ig, 'to'],
        [/\bcontinuous paths\b/ig, 'paths'],
        [/\busing\b/ig, 'with'],
        [/\bexciting\b/ig, 'fun'],
        [/\bchallenging\b/ig, 'smart']
    ];
    for (const [pattern, replacement] of replacements) {
        text = text.replace(pattern, replacement).replace(/\s+/g, ' ');
        if (text.length <= maxLength - 1) {
            return ensureTrailingSpaceWithinLimit(text, maxLength);
        }
    }

    const hardLimit = maxLength - 1;
    let cut = text.slice(0, hardLimit);
    const lastSpace = cut.lastIndexOf(' ');
    if (lastSpace >= 45) {
        cut = cut.slice(0, lastSpace);
    }
    cut = cut.replace(/[,\-:;]+$/g, '').trim();
    if (!/[.!?]$/.test(cut)) {
        cut += '.';
    }
    if (cut.length > hardLimit) {
        cut = cut
            .slice(0, hardLimit)
            .replace(/\s+\S*$/, '')
            .replace(/[,\-:;]+$/g, '')
            .trim();
    }
    return ensureTrailingSpaceWithinLimit(cut, maxLength);
}

async function acquireOrCreateAuxPage(context, matcher, urlToOpen, label) {
    const pages = context.pages().filter(p => !p.isClosed());
    const reused = pages.find(p => matcher.test(String(p.url() || '')));
    if (reused) {
        await reused.bringToFront().catch(() => { });
        return { page: reused, source: `reused-${label}` };
    }

    const created = await context.newPage();
    await created.bringToFront().catch(() => { });
    await created.goto(urlToOpen, { timeout: 120000, waitUntil: 'domcontentloaded' });
    await delay(created, 2500);
    return { page: created, source: `created-${label}` };
}

async function ensureAppGenieLoggedIn(appGeniePage, runtimeOptions) {
    const isLoggedIn = async () => {
        const currentUrl = String(appGeniePage.url() || '');
        if (/\/(my-submit-tasks|submit-tasks)(?:[/?#]|$)/i.test(currentUrl)) {
            return true;
        }

        const loggedInIndicators = [
            'li[data-menu-id*="/my-submit-tasks"]',
            'li.ant-menu-item:has-text("我的任务")',
            'button:has-text("退出"), a:has-text("退出")'
        ];

        for (const selector of loggedInIndicators) {
            const visible = await appGeniePage.locator(selector).first().isVisible().catch(() => false);
            if (visible) {
                return true;
            }
        }
        return false;
    };

    if (await isLoggedIn()) {
        return;
    }

    const accountInput = appGeniePage.locator(
        '#account, input[name="account"], input[autocomplete="username"], input[placeholder*="账号"], input[placeholder*="邮箱"]'
    ).first();
    const passwordInput = appGeniePage.locator(
        '#password, input[name="password"], input[type="password"]'
    ).first();
    const submitBtn = appGeniePage.locator(
        'button[type="submit"], button:has-text("登录"), .ant-btn-primary:has-text("登录")'
    ).first();

    const needLogin = await passwordInput.isVisible().catch(() => false);
    if (!needLogin) {
        return;
    }

    console.log('[STEP] AppGenie login required, filling web_username/web_password...');
    await retryAction(async () => {
        await accountInput.waitFor({ state: 'visible', timeout: 30000 });
        await accountInput.click({ timeout: 10000 });
        await accountInput.fill('');
        await accountInput.type(runtimeOptions.webUsername, { delay: 60 });

        await randomDelay(appGeniePage, 1500, 3000);

        await passwordInput.waitFor({ state: 'visible', timeout: 30000 });
        await passwordInput.click({ timeout: 10000 });
        await passwordInput.fill('');
        await passwordInput.type(runtimeOptions.webPassword, { delay: 55 });

        await randomDelay(appGeniePage, 2000, 4000);

        await submitBtn.waitFor({ state: 'visible', timeout: 30000 });
        await submitBtn.click({ timeout: 10000 });
    }, 'AppGenie login submit', 3);

    // Use fast polling here to avoid long global retry backoff (5-10s) on slow redirects.
    const loginWaitStart = Date.now();
    const loginWaitTimeoutMs = 45000;
    while (Date.now() - loginWaitStart < loginWaitTimeoutMs) {
        if (await isLoggedIn()) {
            return;
        }
        await appGeniePage.waitForTimeout(1200 + Math.floor(Math.random() * 1200));
    }
    throw new Error('AppGenie did not reach logged-in state yet.');
}

async function ensureAppGenieOnMyTasks(appGeniePage) {
    const currentUrl = String(appGeniePage.url() || '');
    if (/\/my-submit-tasks(?:[/?#]|$)/i.test(currentUrl)) {
        return;
    }

    const myTasksMenu = appGeniePage.locator(
        'li[data-menu-id*="/my-submit-tasks"], li.ant-menu-item:has-text("我的任务"), a:has-text("我的任务")'
    ).first();

    await retryAction(async () => {
        await myTasksMenu.waitFor({ state: 'visible', timeout: 30000 });
        await myTasksMenu.scrollIntoViewIfNeeded({ timeout: 5000 }).catch(() => { });
        await myTasksMenu.click({ timeout: 10000 });
    }, 'Open AppGenie 我的任务', 3);

    await retryAction(async () => {
        const onMyTasks = /\/my-submit-tasks(?:[/?#]|$)/i.test(String(appGeniePage.url() || ''));
        const selected = await appGeniePage
            .locator('li.ant-menu-item-selected:has-text("我的任务"), li[data-menu-id*="/my-submit-tasks"].ant-menu-item-selected')
            .first()
            .isVisible()
            .catch(() => false);
        if (!onMyTasks && !selected) {
            throw new Error('Still not on AppGenie 我的任务 page.');
        }
    }, 'Wait AppGenie 我的任务 page', 4);
}

function normalizeAppGenieTypeKey(typeLabel) {
    const raw = String(typeLabel || '').trim();
    if (!raw) {
        return '';
    }
    if (/游戏|game/i.test(raw)) {
        return 'game';
    }
    if (/应用|app|工具/i.test(raw)) {
        return 'app';
    }
    return '';
}

async function readAppGenieTypeFromCard(appCard) {
    const typeLabel = await appCard.evaluate((card) => {
        const normalizeText = (value) => String(value || '').replace(/\s+/g, ' ').trim();

        // Primary: AppGenie type tag (e.g. <span class="ant-tag ant-tag-filled ant-tag-orange">游戏类</span>)
        const primarySelectors = [
            'span.ant-tag.ant-tag-filled.ant-tag-orange',
            'span.ant-tag.ant-tag-orange'
        ];
        for (const selector of primarySelectors) {
            const orangeTag = card.querySelector(selector);
            if (orangeTag) {
                const text = normalizeText(orangeTag.textContent);
                if (text) {
                    return text;
                }
            }
        }

        // Fallback: pick a type-like tag text from this card.
        const tags = Array.from(card.querySelectorAll('span.ant-tag'));
        for (const tag of tags) {
            const text = normalizeText(tag.textContent);
            if (/(类$)|游戏|工具|应用|game|app/i.test(text)) {
                return text;
            }
        }

        return '';
    }).catch(() => '');

    return String(typeLabel || '').trim();
}

async function openAppGenieDetailsAndReadPrivacyText(context, task, runtimeOptions) {
    const { page: appGeniePage, source } = await acquireOrCreateAuxPage(
        context,
        /appgenie-ai\.com/i,
        'https://appgenie-ai.com/login',
        'appgenie'
    );
    console.log(`[PAGE] AppGenie tab: ${source} | ${appGeniePage.url() || 'about:blank'}`);
    await appGeniePage.bringToFront();
    await delay(appGeniePage, 1000);

    await ensureAppGenieLoggedIn(appGeniePage, runtimeOptions);
    await ensureAppGenieOnMyTasks(appGeniePage);
    // Slow down intentionally on VPS so list state is fully ready before actions.
    await randomDelay(appGeniePage, 5000, 8000);

    let emailRow = appGeniePage.locator('tr.ant-table-row').filter({ hasText: runtimeOptions.contactEmail }).first();
    if (!(await emailRow.isVisible().catch(() => false))) {
        emailRow = appGeniePage.locator('tr.ant-table-row').first();
    }
    await emailRow.waitFor({ state: 'visible', timeout: 60000 });

    await retryAction(async () => {
        let openBtn = emailRow.locator('button').filter({ hasText: /查看应用|View\s*app/i }).first();
        if (!(await openBtn.isVisible().catch(() => false))) {
            openBtn = emailRow.locator('button').last();
        }
        if (await openBtn.isVisible().catch(() => false)) {
            await clickLocatorRobust(openBtn, 'AppGenie View app button', 10000);
        } else {
            await clickLocatorRobust(emailRow, 'AppGenie task row', 10000);
        }
    }, 'Open AppGenie pending app drawer', 3);

    // Give drawer animation/data some time before locating target card.
    await randomDelay(appGeniePage, 4000, 7000);

    const drawerBody = appGeniePage.locator('.ant-drawer-content .ant-drawer-body, .ant-drawer-body').first();
    await drawerBody.waitFor({ state: 'visible', timeout: 60000 }).catch(() => { });
    const cardScope = (await drawerBody.isVisible().catch(() => false)) ? drawerBody : appGeniePage;

    let appCard = cardScope.locator('div.ant-card').filter({ hasText: task.packageName }).first();
    if (!(await appCard.isVisible().catch(() => false))) {
        appCard = cardScope.locator('div.ant-card').filter({ hasText: task.appName }).first();
    }
    await appCard.waitFor({ state: 'visible', timeout: 60000 });

    // Capture current row type (e.g. 游戏类/应用类) before entering details page.
    const appGenieTypeLabel = await readAppGenieTypeFromCard(appCard);
    task.appGenieTypeLabel = appGenieTypeLabel;
    task.appGenieTypeKey = normalizeAppGenieTypeKey(appGenieTypeLabel);
    if (task.appGenieTypeLabel) {
        const keyText = task.appGenieTypeKey ? ` -> ${task.appGenieTypeKey}` : '';
        console.log(`[APPGENIE] Type detected: ${task.appGenieTypeLabel}${keyText}`);
    } else {
        console.log('[APPGENIE] Type tag not found on current app card.');
    }

    const detailBtnByText = appCard.locator('button').filter({ hasText: /详\s*情|Detail/i }).first();
    const detailBtnFallback = appCard.locator('div[style*="border-top"] button.ant-btn').first();

    let detailsPage = appGeniePage;
    const popupPromise = context.waitForEvent('page', { timeout: 30000 }).catch(() => null);
    await retryAction(async () => {
        let detailBtn = detailBtnByText;
        if (!(await detailBtn.isVisible().catch(() => false))) {
            detailBtn = detailBtnFallback;
        }
        await detailBtn.waitFor({ state: 'visible', timeout: 30000 });
        await detailBtn.scrollIntoViewIfNeeded({ timeout: 5000 }).catch(() => { });
        // Slow down between View App and Details click.
        await randomDelay(appGeniePage, 3000, 6000);
        await clickLocatorRobust(detailBtn, 'AppGenie Details button', 10000);
    }, 'Click AppGenie details button', 3);

    const popup = await popupPromise;
    if (popup) {
        detailsPage = popup;
        await detailsPage.bringToFront().catch(() => { });
        await detailsPage.waitForLoadState('domcontentloaded', { timeout: 60000 }).catch(() => { });
        await detailsPage.waitForLoadState('load', { timeout: 60000 }).catch(() => { });
        await detailsPage.waitForLoadState('networkidle', { timeout: 60000 }).catch(() => { });
    }
    await randomDelay(detailsPage, 5000, 9000);
    await delay(detailsPage, 3000);

    const privacyCard = detailsPage.locator('div.ant-card').filter({
        has: detailsPage.locator('.ant-card-head-title').filter({ hasText: /隐私|Privacy/i }).first()
    }).first();
    const privacyBody = privacyCard.locator('.ant-card-body').first();

    let privacyText = '';
    await retryAction(async () => {
        await privacyCard.waitFor({ state: 'visible', timeout: 45000 });
        await privacyBody.waitFor({ state: 'visible', timeout: 45000 });
        const text = String(await privacyBody.innerText().catch(() => '')).trim();
        // Prevent premature handoff to Sites with half-loaded/empty content.
        if (text.length < 40) {
            throw new Error(`Privacy text not ready yet (length=${text.length}).`);
        }
        privacyText = text;
    }, 'Wait AppGenie privacy content ready', 4);
    console.log(`[COPY] Privacy text found (length=${privacyText.length}).`);

    let copied = false;
    const copyBtn = privacyCard.locator('button:has-text("复制"), button:has-text("Copy"), [role="button"]:has-text("复制"), [role="button"]:has-text("Copy")').first();
    if (await copyBtn.isVisible().catch(() => false)) {
        copied = await clickLocatorRobust(copyBtn, 'AppGenie privacy Copy button', 10000).then(() => true).catch(() => false);
    }
    if (!copied) {
        copied = await detailsPage.evaluate(async (text) => {
            try {
                await navigator.clipboard.writeText(text);
                return true;
            } catch {
                return false;
            }
        }, privacyText).catch(() => false);
    }
    if (copied) {
        console.log('[COPY] Privacy text copy success.');
    } else {
        console.log('[COPY] Privacy text copy skipped/blocked, using captured text in memory.');
    }

    // Keep pages open for later reuse/download steps.
    task.appGeniePage = appGeniePage;
    task.appGenieDetailPage = detailsPage;
    return privacyText;
}

async function createAndPublishGoogleSite(context, task, privacyText) {
    const { page: sitesPage, source } = await acquireOrCreateAuxPage(
        context,
        /sites\.google\.com/i,
        'https://sites.google.com/new',
        'sites'
    );
    console.log(`[PAGE] Sites tab: ${source} | ${sitesPage.url() || 'about:blank'}`);
    await sitesPage.bringToFront();
    await sitesPage.goto('https://sites.google.com/new', { timeout: 120000, waitUntil: 'domcontentloaded' });
    await randomDelay(sitesPage, 800, 1800);

    const waitSitesEditorReady = async (timeoutMs = 90000) => {
        await sitesPage.waitForFunction(() => {
            const urlReady = /\/edit(?:[/?#]|$)/i.test(String(window.location.pathname || ''));
            const hasDocNameInput = !!document.querySelector(
                '#i3, input#i3, input[aria-labelledby*="Loading name"], input.VfPpkd-fmcmS-wGMbrd'
            );
            const hasEditableText = !!document.querySelector('div[contenteditable="true"][role="textbox"]');
            const hasPublish = Array.from(document.querySelectorAll('button, div[role="button"]')).some(el =>
                /\bPublish\b/i.test((el.textContent || '').trim())
            );
            return urlReady || hasDocNameInput || hasEditableText || hasPublish;
        }, { timeout: timeoutMs });
    };

    const titleInput = sitesPage.locator(
        '#i3, input#i3, input[aria-labelledby*="Loading name"], input.VfPpkd-fmcmS-wGMbrd, input[aria-label*="site name" i], input[placeholder*="site name" i]'
    ).first();

    // sites.google.com/new often lands on template chooser. Enter editor by clicking Blank site first.
    if (!(await titleInput.isVisible().catch(() => false))) {
        console.log('[STEP] Sites template chooser detected, clicking Blank site...');
        const blankSiteCard = sitesPage
            .locator('div[role="option"], div.docs-homescreen-templates-templateview')
            .filter({ hasText: /Blank site|空白站点|空白网站|空白/i })
            .first();

        await retryAction(async () => {
            await blankSiteCard.waitFor({ state: 'visible', timeout: 60000 });
            await blankSiteCard.scrollIntoViewIfNeeded({ timeout: 10000 }).catch(() => { });
            await blankSiteCard.click({ timeout: 10000 });
        }, 'Click Sites Blank site', 3);
        await waitSitesEditorReady(120000);
    }
    await waitSitesEditorReady(120000);

    await titleInput.waitFor({ state: 'visible', timeout: 120000 });
    console.log('[STEP] Sites editor ready, filling site name...');
    await titleInput.click({ clickCount: 2, timeout: 10000 });
    await titleInput.fill(task.appName);
    await delay(sitesPage, 500);

    const headerArea = sitesPage.locator('div[jsname="Zz0G1b"], section[data-header="true"]').first();
    const deleteHeaderBtn = sitesPage.locator(
        'button[aria-label="Delete header"], button[aria-label*="Delete header"], button[aria-label*="删除"]'
    ).first();

    console.log('[STEP] Clearing Sites header block by trash button...');
    const headerVisible = await headerArea.isVisible().catch(() => false);
    if (headerVisible) {
        await retryAction(async () => {
            await headerArea.waitFor({ state: 'visible', timeout: 15000 });
            await headerArea.click({ timeout: 10000 });
            await delay(sitesPage, 320);
            await deleteHeaderBtn.waitFor({ state: 'visible', timeout: 15000 });
            await deleteHeaderBtn.click({ timeout: 10000 });
        }, 'Delete Sites header by trash button', 4);
        await delay(sitesPage, 900);
    } else {
        console.log('[STEP] Sites header area not visible, skip header cleanup.');
    }

    const bodyCanvas = sitesPage.locator(
        'div.YyZdKd, article[guidedhelpid="at-canvas"], article.UynGwb, article[guidedhelpid]'
    ).first();
    await bodyCanvas.waitFor({ state: 'visible', timeout: 120000 });
    console.log('[STEP] Activating white canvas by single click...');
    await retryAction(async () => {
        await bodyCanvas.click({ timeout: 10000, position: { x: 300, y: 300 } });
    }, 'Activate Sites white canvas by click', 3);
    await delay(sitesPage, 380);

    const contentTarget = sitesPage.locator(
        'xpath=(//div[@contenteditable="true" and @role="textbox" and not(ancestor::section[@data-header="true"])])[1]'
    ).first();
    if (await contentTarget.isVisible().catch(() => false)) {
        await contentTarget.click({ timeout: 10000 }).catch(() => { });
    }
    await delay(sitesPage, 220);

    const clipboardSeeded = await sitesPage.evaluate(async (text) => {
        try {
            await navigator.clipboard.writeText(String(text || ''));
            return true;
        } catch (_) {
            return false;
        }
    }, privacyText).catch(() => false);
    if (clipboardSeeded) {
        console.log('[COPY] Sites clipboard prepared.');
    } else {
        console.log('[COPY] Sites clipboard write blocked, trying Ctrl+V with existing clipboard.');
    }

    await sitesPage.keyboard.press('Control+V');
    await delay(sitesPage, 3000);
    console.log('[COPY] Privacy text pasted into Sites editor.');

    const topPublishBtn = sitesPage.locator(
        'div[role="button"][data-tooltip="Publish"], div[role="button"]:has-text("Publish"), button:has-text("Publish")'
    ).first();
    await retryAction(async () => {
        await topPublishBtn.waitFor({ state: 'visible', timeout: 60000 });
        await clickLocatorRobust(topPublishBtn, 'Sites top Publish button', 10000);
    }, 'Click Sites Publish(top)', 3);
    await delay(sitesPage, 3000);

    const publishDialog = sitesPage.locator('div[role="dialog"], div[aria-modal="true"]')
        .filter({ hasText: /Publish to the web|Web address|发布/i })
        .first();
    await publishDialog.waitFor({ state: 'visible', timeout: 60000 });

    const addressInput = publishDialog.locator(
        'div[jsname="YEWROd"] input.poFWNe[jsname="YPqjbf"], input.poFWNe[jsname="YPqjbf"][maxlength="31"], input.poFWNe[maxlength="31"]'
    ).first();
    await addressInput.waitFor({ state: 'visible', timeout: 60000 });

    let selectedSlug = '';
    for (let attempt = 1; attempt <= 5; attempt++) {
        selectedSlug = buildSiteSlug(task.appName);
        await addressInput.evaluate((el, slug) => {
            el.focus();
            el.value = '';
            el.dispatchEvent(new Event('input', { bubbles: true }));
            el.value = slug;
            el.dispatchEvent(new Event('input', { bubbles: true }));
            el.dispatchEvent(new Event('change', { bubbles: true }));
        }, selectedSlug);
        const actualSlug = await addressInput.inputValue().catch(() => '');
        if (actualSlug !== selectedSlug) {
            await addressInput.fill(selectedSlug, { timeout: 10000 }).catch(() => { });
        }
        await delay(sitesPage, 3000);

        const unavailableHint = publishDialog.locator("text=/unavailable|already exists|can't|invalid|已被使用|不可用|无效/i").first();
        if (await unavailableHint.isVisible().catch(() => false)) {
            continue;
        }
        break;
    }
    if (!selectedSlug) {
        throw new Error('Failed to generate a valid Google Sites web address.');
    }

    const finalPublishBtn = publishDialog.locator(
        'div[role="button"][data-id="j6LnYe"], div[role="button"]:has-text("Publish"), button:has-text("Publish"), div[role="button"]:has-text("发布"), button:has-text("发布")'
    ).last();
    await retryAction(async () => {
        await finalPublishBtn.waitFor({ state: 'visible', timeout: 60000 });
        await clickLocatorRobust(finalPublishBtn, 'Sites final Publish button', 10000);
    }, 'Click Sites Publish(final)', 3);
    await delay(sitesPage, 5000);

    // Try to copy the link from publish dialog. If not available, fallback to deterministic URL.
    const copyLinkBtn = sitesPage.locator('div[role="button"]:has-text("Copy link"), button:has-text("Copy link")').first();
    if (await copyLinkBtn.isVisible().catch(() => false)) {
        await copyLinkBtn.click({ timeout: 10000 }).catch(() => { });
    }

    return `https://sites.google.com/view/${selectedSlug}/home`;
}

async function runOnce(task, appListUrl, statusManager, runtimeOptions) {
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
        const pageSelection = await acquireWorkingConsolePage(context, appListUrl);
        page = pageSelection.page;
        console.log(`[PAGE] Working tab: ${pageSelection.source} | ${page.url() || 'about:blank'}`);

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

        const manualLoginWaitMs = getManualLoginWaitMs();
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

            const isLikelyGoogleLoginPage = async () => {
                const currentUrl = page.url();
                if (/accounts\.google\.com|ServiceLogin|signin\/v2|identifier/i.test(currentUrl)) {
                    return true;
                }

                const loginInputs = page.locator(
                    'input[type="email"], input[type="password"], input[name="identifier"], input[name="Passwd"], #identifierId'
                ).first();
                if (await loginInputs.isVisible().catch(() => false)) {
                    return true;
                }

                const loginTexts = page.locator(
                    'text=/Sign in|鐧诲綍|Use your Google Account|閫夋嫨璐﹀彿|Choose an account/i'
                ).first();
                return await loginTexts.isVisible().catch(() => false);
            };

            const waitForManualLoginUntilAppListReady = async () => {
                const startMs = Date.now();
                let lastProgressLogAt = 0;
                let lastAppListRedirectAt = 0;
                console.log(
                    `[LOGIN] Waiting for manual login/account selection. Timeout: ${Math.round(manualLoginWaitMs / 1000)}s`
                );

                while (Date.now() - startMs < manualLoginWaitMs) {
                    if (await createBtn.isVisible().catch(() => false)) {
                        await ensureDeveloperSelectedIfNeeded();
                        if (await createBtn.isVisible().catch(() => false)) {
                            console.log('[LOGIN] Login completed. "Create app" is visible.');
                            return true;
                        }
                    }

                    await ensureDeveloperSelectedIfNeeded();

                    const currentUrl = page.url();
                    const now = Date.now();
                    const elapsedMs = now - startMs;

                    if (now - lastProgressLogAt >= 15000) {
                        const remainingSec = Math.max(
                            0,
                            Math.ceil((manualLoginWaitMs - elapsedMs) / 1000)
                        );
                        console.log(`[LOGIN] Still waiting for login... ${remainingSec}s left.`);
                        lastProgressLogAt = now;
                    }

                    const onGoogleLoginPage = await isLikelyGoogleLoginPage();
                    if (!onGoogleLoginPage &&
                        /play\.google\.com\/console/i.test(currentUrl) &&
                        !/\/app-list(?:\?|$)/.test(currentUrl) &&
                        now - lastAppListRedirectAt >= 15000) {
                        console.log('[LOGIN] Returned to Play Console but not app list, redirecting to app-list...');
                        await page.goto(appListUrl, { timeout: 90000, waitUntil: 'domcontentloaded' }).catch(() => { });
                        await delay(page, 2500);
                        lastAppListRedirectAt = now;
                    }

                    await delay(page, 2500);
                }

                return false;
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

            if (!(await createBtn.isVisible().catch(() => false))) {
                const loginReady = await waitForManualLoginUntilAppListReady();
                if (!loginReady) {
                    throw new Error(
                        'Manual login wait timed out. ' +
                        `Please finish Google login within ${Math.round(manualLoginWaitMs / 1000)} seconds and rerun.`
                    );
                }
            }
        };

        await openAppListPage();
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

            // Randomly choose App or Game, and verify the selected state.
            const isApp = Math.random() < 0.5;
            const typeLabel = isApp ? 'App' : 'Game';
            const typeId = isApp ? 'app-radio' : 'game-radio';
            console.log(`Selecting type: ${typeLabel}`);

            const isTypeSelected = async (targetTypeId) => {
                const root = page.locator(`[debug-id="${targetTypeId}"]`).first();
                if (!(await root.isVisible().catch(() => false))) {
                    return false;
                }

                const inputRadio = root.locator('input[type="radio"]').first();
                if (await inputRadio.count().catch(() => 0)) {
                    if (await inputRadio.isChecked().catch(() => false)) {
                        return true;
                    }
                }

                const roleRadio = root.locator('[role="radio"]').first();
                if (await roleRadio.count().catch(() => 0)) {
                    const roleChecked = await roleRadio.getAttribute('aria-checked').catch(() => null);
                    if (roleChecked === 'true') {
                        return true;
                    }
                }

                const selfChecked = await root.getAttribute('aria-checked').catch(() => null);
                return selfChecked === 'true';
            };

            const selectTypeWithVerify = async () => {
                for (let attempt = 1; attempt <= 4; attempt++) {
                    const typeRoot = page.locator(`[debug-id="${typeId}"]`).first();
                    await typeRoot.waitFor({ state: 'visible', timeout: 10000 });
                    await typeRoot.scrollIntoViewIfNeeded({ timeout: 5000 }).catch(() => { });
                    await typeRoot.click({ timeout: 10000 });
                    await delay(page, 900);

                    if (await isTypeSelected(typeId)) {
                        return;
                    }

                    // Fallback: click text label if debug-id click didn't stick.
                    const textTarget = page.getByText(typeLabel, { exact: true }).first();
                    if (await textTarget.isVisible().catch(() => false)) {
                        await textTarget.click({ timeout: 8000 }).catch(() => { });
                        await delay(page, 900);
                    }

                    if (await isTypeSelected(typeId)) {
                        return;
                    }

                    console.log(`[CREATE] Type selection not confirmed (${typeLabel}), retrying (${attempt}/4)...`);
                }

                throw new Error(`Failed to confirm selected type: ${typeLabel}`);
            };

            await retryAction(selectTypeWithVerify, `Select type: ${typeLabel}`);
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
                const recoverDelayMs = 4000 + Math.floor(Math.random() * 3001);
                console.log(`[CREATE] Waiting ${Math.round(recoverDelayMs / 1000)}s before skipping this row...`);
                await delay(page, recoverDelayMs);

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

                const recoverDelayMs = 4000 + Math.floor(Math.random() * 3001);
                console.log(
                    `[CREATE] Create result unresolved, waiting ${Math.round(recoverDelayMs / 1000)}s before skipping this row...`
                );
                await delay(page, recoverDelayMs);

                const unresolvedError = new Error(
                    'Create app did not finish in time. ' +
                    `Current URL: ${currentUrl}. ` +
                    `Visible errors: ${compactErrors || 'none'}. ` +
                    'Skipping this row and continuing with the next one.'
                );
                unresolvedError.code = 'CREATE_FAILED_TIMEOUT';
                throw unresolvedError;
            }

            await page.waitForLoadState('domcontentloaded', { timeout: 30000 }).catch(() => { });
            console.log('App created successfully, navigating to dashboard...');
            await delay(page, 8000);

            // Extract app base path.
            appBasePath = extractAppBasePathFromCurrentUrl();
            console.log('App path:', appBasePath);

            statusManager.ensureTaskProgressAtLeast(task, PROGRESS_STEP_APP_CREATED);
            statusManager.updateTaskStatus(task, STATUS_PARTIAL);

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

        async function goToAppContentViaMonitorPolicyMenu() {
            const monitorAndImproveLink = page.locator(
                'a[href*="/monitor"], a.item-link:has(.item-label:has-text("Monitor and improve")), [role="button"]:has-text("Monitor and improve")'
            ).first();
            const policyAndProgramsLink = page.locator(
                'a.item-link:has(.item-label:has-text("Policy and programs")), [role="button"]:has-text("Policy and programs")'
            ).first();
            const appContentLink = page.locator(
                'a[href*="/app-content/overview"], a[href*="/app-content"], a.item-link:has(.item-label:has-text("App content")), [role="button"]:has-text("App content")'
            ).first();

            console.log('Navigating to "Monitor and improve"...');
            await retryAction(async () => {
                await monitorAndImproveLink.waitFor({ state: 'visible', timeout: 60000 });
                await monitorAndImproveLink.scrollIntoViewIfNeeded({ timeout: 5000 }).catch(() => { });
                await monitorAndImproveLink.click({ timeout: 10000 });
                await delay(page, 2500);
            }, 'Open Monitor and improve', 3);
            await delay(page, 2500);

            console.log('Opening "Policy and programs"...');
            await retryAction(async () => {
                const appContentVisibleBefore = await appContentLink.isVisible().catch(() => false);
                if (!appContentVisibleBefore) {
                    await policyAndProgramsLink.waitFor({ state: 'visible', timeout: 60000 });
                    await policyAndProgramsLink.scrollIntoViewIfNeeded({ timeout: 5000 }).catch(() => { });
                    await policyAndProgramsLink.click({ timeout: 10000 });
                    await delay(page, 1500);
                }
                await appContentLink.waitFor({ state: 'visible', timeout: 45000 });
            }, 'Open Policy and programs', 3);
            await delay(page, 2000);

            console.log('Opening "App content"...');
            await retryAction(async () => {
                await appContentLink.waitFor({ state: 'visible', timeout: 60000 });
                await appContentLink.scrollIntoViewIfNeeded({ timeout: 5000 }).catch(() => { });
                await appContentLink.click({ timeout: 10000 });
                await page.waitForURL(/\/app-content\/overview(?:[/?#]|$)|\/app-content(?:[/?#]|$)/, { timeout: 120000 });
            }, 'Open App content from menu', 3);
            await delay(page, 4000);

            const tab = page.locator('div[role="tab"]:has-text("Need attention")');
            if (await tab.isVisible().catch(() => false)) {
                await tab.click().catch(() => { });
                await delay(page, 2000);
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

                await buttonByCard.waitFor({ state: 'visible', timeout: 20000 });
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

                await clickLocatorRobust(foundBtn, `"${text}" button`, 10000);
                await delay(page, 3000);
            }, `Click "${text}" button`);
        }

        async function fillFirstVisibleInput(selectors, value, label) {
            await retryAction(async () => {
                let target = null;
                for (const sel of selectors) {
                    const loc = page.locator(sel).first();
                    if (await loc.isVisible().catch(() => false)) {
                        target = loc;
                        break;
                    }
                }
                if (!target) {
                    throw new Error(`${label} input not found.`);
                }
                await target.scrollIntoViewIfNeeded({ timeout: 5000 }).catch(() => { });
                await target.fill(value);
                await delay(page, 1200);
            }, `Fill ${label}`);
        }

        async function afterItem10Wait(ms = 3000) {
            await delay(page, ms);
        }

        async function gotoAppSubPage(subPath, label, waitMs = 3000) {
            console.log(`Navigating to "${label}"...`);
            await page.bringToFront().catch(() => { });
            await page.goto(appBasePath + subPath, { timeout: 120000, waitUntil: 'domcontentloaded' });
            await page.waitForLoadState('load', { timeout: 60000 }).catch(() => { });
            await delay(page, waitMs);
        }

        async function fillFirstVisibleInputOn(targetPage, selectors, value, label, waitMs = 3000) {
            await retryAction(async () => {
                let target = null;
                for (const sel of selectors) {
                    const loc = targetPage.locator(sel).first();
                    if (await loc.isVisible().catch(() => false)) {
                        target = loc;
                        break;
                    }
                }
                if (!target) {
                    throw new Error(`${label} input not found.`);
                }

                await target.scrollIntoViewIfNeeded({ timeout: 5000 }).catch(() => { });
                await target.fill(value, { timeout: 10000 }).catch(async () => {
                    await target.evaluate((el, nextValue) => {
                        el.focus();
                        el.value = '';
                        el.dispatchEvent(new Event('input', { bubbles: true }));
                        el.value = nextValue;
                        el.dispatchEvent(new Event('input', { bubbles: true }));
                        el.dispatchEvent(new Event('change', { bubbles: true }));
                    }, value);
                });
                await target.evaluate((el, nextValue) => {
                    el.focus();
                    if (el.value !== nextValue) {
                        el.value = nextValue;
                    }
                    el.dispatchEvent(new Event('input', { bubbles: true }));
                    el.dispatchEvent(new Event('change', { bubbles: true }));
                    el.blur();
                }, value).catch(() => { });
                await delay(targetPage, waitMs);
            }, `Fill ${label}`, 3);
        }

        async function clickFirstVisibleOn(targetPage, locators, label, waitMs = 3000) {
            await retryAction(async () => {
                let target = null;
                for (const loc of locators) {
                    if (await loc.isVisible().catch(() => false)) {
                        target = loc;
                        break;
                    }
                }
                if (!target) {
                    throw new Error(`${label} not found.`);
                }
                await clickLocatorRobust(target, label, 10000);
                await delay(targetPage, waitMs);
            }, `Click ${label}`, 3);
        }

        async function ensureAppGenieDetailPageReady() {
            if (task.appGenieDetailPage && !task.appGenieDetailPage.isClosed()) {
                await task.appGenieDetailPage.bringToFront().catch(() => { });
                await delay(task.appGenieDetailPage, 3000);
                return task.appGenieDetailPage;
            }

            // Re-open details if this run started after privacy step or Chrome recycled tabs.
            await openAppGenieDetailsAndReadPrivacyText(context, task, runtimeOptions);
            if (!task.appGenieDetailPage || task.appGenieDetailPage.isClosed()) {
                throw new Error('AppGenie detail page is not available for later steps.');
            }
            await task.appGenieDetailPage.bringToFront().catch(() => { });
            await delay(task.appGenieDetailPage, 3000);
            return task.appGenieDetailPage;
        }

        async function readAppGenieDetailCardText(titleRegex, label) {
            const detailPage = await ensureAppGenieDetailPageReady();
            let card = detailPage.locator('div.ant-card').filter({
                has: detailPage.locator('.ant-card-head-title').filter({ hasText: titleRegex }).first()
            }).first();

            let text = '';
            if (await card.isVisible().catch(() => false)) {
                text = await card.locator('.ant-card-body').first().innerText().catch(() => '');
            }

            if (!String(text || '').trim()) {
                text = await detailPage.evaluate(({ source, flags }) => {
                    const titlePattern = new RegExp(source, flags);
                    const normalize = (value) => String(value || '').replace(/\s+/g, ' ').trim();
                    const cards = Array.from(document.querySelectorAll('div.ant-card'));
                    for (const candidate of cards) {
                        const title = normalize(candidate.querySelector('.ant-card-head-title')?.textContent || '');
                        if (!titlePattern.test(title)) continue;
                        const body = candidate.querySelector('.ant-card-body');
                        return String(body?.innerText || body?.textContent || '').trim();
                    }
                    return '';
                }, { source: titleRegex.source, flags: titleRegex.flags || 'i' }).catch(() => '');
            }

            text = normalizeAppGenieCardText(text);
            if (!text) {
                throw new Error(`${label} not found on AppGenie details page.`);
            }
            return text;
        }

        async function saveDownloadToTemp(download, folderLabel, fallbackName) {
            const root = path.join(
                os.tmpdir(),
                'app-create-downloads',
                safeFileToken(`${task.appName}-${task.packageName}`),
                `${safeFileToken(folderLabel)}-${Date.now()}`
            );
            fs.mkdirSync(root, { recursive: true });
            const suggested = safeFileToken(download.suggestedFilename() || fallbackName);
            const filePath = path.join(root, suggested || fallbackName);
            await download.saveAs(filePath);
            return { filePath, root };
        }

        async function downloadFromAppGenieDetail(buttonRegex, label, skipRegex = null) {
            const detailPage = await ensureAppGenieDetailPageReady();
            const candidates = detailPage.locator('button, a, [role="button"]').filter({ hasText: buttonRegex });
            const count = await candidates.count().catch(() => 0);
            let target = null;
            for (let i = 0; i < count; i += 1) {
                const loc = candidates.nth(i);
                if (!(await loc.isVisible().catch(() => false))) continue;
                const text = await loc.innerText().catch(() => '');
                if (skipRegex && skipRegex.test(text)) continue;
                target = loc;
                break;
            }
            if (!target) {
                throw new Error(`${label} download button not found on AppGenie details page.`);
            }

            console.log(`[APPGENIE] Downloading ${label}...`);
            const downloadPromise = detailPage.waitForEvent('download', { timeout: 90000 });
            await clickLocatorRobust(target, `${label} download button`, 10000);
            return await downloadPromise;
        }

        async function downloadAndExtractStoreImages() {
            const download = await downloadFromAppGenieDetail(/全部下载|Download\s*all|All\s*download/i, 'store images ZIP');
            const { filePath, root } = await saveDownloadToTemp(download, 'images', `${safeFileToken(task.appName)}-images.zip`);
            const extractDir = path.join(root, 'extract');
            extractZipToDirectory(filePath, extractDir);
            console.log(`[APPGENIE] Images ZIP extracted: ${extractDir}`);
            await randomDelay(page, 10000, 15000);
            const images = findImagesNamed(extractDir, ['1', '2', '3']);
            console.log(`[APPGENIE] Store images ready: ${images.map(p => path.basename(p)).join(', ')}`);
            return images;
        }

        async function downloadAabFromAppGenie() {
            const detailPage = await ensureAppGenieDetailPageReady();
            let download = null;

            const aabCard = detailPage.locator('div.ant-card').filter({
                has: detailPage.locator('.ant-card-head-title').filter({ hasText: /AAB|App\s*Bundle|应用包|安装包|包体/i }).first()
            }).first();
            if (await aabCard.isVisible().catch(() => false)) {
                const btn = aabCard.locator('button, a, [role="button"]').filter({ hasText: /下载|Download|AAB|aab/i }).first();
                if (await btn.isVisible().catch(() => false)) {
                    console.log('[APPGENIE] Downloading AAB from AAB card...');
                    const downloadPromise = detailPage.waitForEvent('download', { timeout: 90000 });
                    await clickLocatorRobust(btn, 'AAB download button', 10000);
                    download = await downloadPromise;
                }
            }

            if (!download) {
                download = await downloadFromAppGenieDetail(
                    /AAB|aab|App\s*Bundle|应用包|安装包|包体|下载|Download/i,
                    'AAB',
                    /全部下载|Download\s*all|图片|image|zip/i
                );
            }

            const saved = await saveDownloadToTemp(download, 'aab', `${safeFileToken(task.appName)}.aab`);
            console.log(`[APPGENIE] AAB saved: ${saved.filePath}`);
            return saved.filePath;
        }

        async function uploadFilesByUploadButton(targetPage, uploadButton, files, label) {
            await retryAction(async () => {
                const chooserPromise = targetPage.waitForEvent('filechooser', { timeout: 30000 }).catch(() => null);
                await clickLocatorRobust(uploadButton, `${label} Upload button`, 10000);
                const chooser = await chooserPromise;
                if (chooser) {
                    await chooser.setFiles(files);
                    return;
                }

                const input = targetPage.locator('input[type="file"]').last();
                await input.setInputFiles(files, { timeout: 30000 });
            }, `Upload ${label}`, 3);
            await delay(targetPage, 3000);
        }

        async function addSelectedStoreAssetsToListing(label) {
            console.log(`[STORE LISTING] Adding selected ${label} assets to listing...`);
            await retryAction(async () => {
                const addButton = page.locator(
                    'button[debug-id="add-to-content-button"], button[aria-label="Add"], [role="button"][debug-id="add-to-content-button"], [role="button"][aria-label="Add"]'
                ).last();
                await addButton.waitFor({ state: 'visible', timeout: 30000 });
                if (await isLocatorDisabled(addButton)) {
                    throw new Error(`${label} Add button is disabled.`);
                }
                await clickLocatorRobust(addButton, `${label} Add button`, 15000);
            }, `Click ${label} Add button`, 3);
            await randomDelay(page, 5000, 7000);
            console.log(`[STORE LISTING] Selected ${label} assets added.`);
        }

        async function clickStoreSectionEdit(sectionTitleRegex, label, fallbackMode = 'first') {
            const sectionEdit = page.locator('console-section, console-block-1-column, console-form-row, section').filter({
                hasText: sectionTitleRegex
            }).locator('material-button[debug-id="edit-store-listing-section-button"], button:has-text("Edit")').first();
            const allEditButtons = page.locator(
                'material-button[debug-id="edit-store-listing-section-button"], button:has-text("Edit"), [role="button"]:has-text("Edit")'
            );
            const fallback = fallbackMode === 'last' ? allEditButtons.last() : allEditButtons.first();
            await clickFirstVisibleOn(page, [sectionEdit, fallback], `${label} Edit`, 3000);
        }

        function storeListingContactDialogLocator() {
            return page.locator(
                'div[role="dialog"], material-dialog, .mdc-dialog, .mat-mdc-dialog-container, .cdk-overlay-pane'
            ).filter({ hasText: /Store\s*listing\s*contact\s*details/i }).first();
        }

        async function ensureStoreListingContactDialogOpen() {
            const dialog = storeListingContactDialogLocator();
            if (await dialog.isVisible().catch(() => false)) {
                return dialog;
            }

            await clickStoreSectionEdit(/Store\s*listing\s*contact\s*details/i, 'Store listing contact details');
            await dialog.waitFor({ state: 'visible', timeout: 30000 });
            await delay(page, 3000);
            return dialog;
        }

        async function fillStoreListingContactEmail(email) {
            await retryAction(async () => {
                const dialog = await ensureStoreListingContactDialogOpen();
                const target = dialog.locator(
                    'material-input[debug-id="email-input"] input, input[debug-id="email-input"], input.mdc-text-field__input, input[type="email"], input[type="text"]'
                ).first();
                await target.waitFor({ state: 'visible', timeout: 15000 });
                await target.scrollIntoViewIfNeeded({ timeout: 5000 }).catch(() => { });

                await target.evaluate((el, nextValue) => {
                    el.focus();
                    el.value = '';
                    el.dispatchEvent(new Event('input', { bubbles: true }));
                    el.value = nextValue;
                    el.dispatchEvent(new Event('input', { bubbles: true }));
                    el.dispatchEvent(new Event('change', { bubbles: true }));
                    el.blur();
                }, email);

                const currentValue = await target.inputValue().catch(() => '');
                if (currentValue !== email) {
                    throw new Error(`Store listing contact email was not filled correctly. Current="${currentValue}"`);
                }
                await delay(page, 3000);
            }, 'Fill Store listing contact email', 3);
        }

        async function clickStoreListingContactSave() {
            await retryAction(async () => {
                const dialog = await ensureStoreListingContactDialogOpen();
                const candidates = dialog.locator(
                    'button[debug-id="main-button"], button:has-text("Save"), material-button:has-text("Save"), [role="button"]:has-text("Save")'
                );
                const count = await candidates.count().catch(() => 0);
                let target = null;
                for (let i = count - 1; i >= 0; i -= 1) {
                    const candidate = candidates.nth(i);
                    const text = (await candidate.innerText().catch(() => '')).replace(/\s+/g, ' ').trim();
                    if (!/^Save$/i.test(text)) continue;
                    if (!(await candidate.isVisible().catch(() => false))) continue;
                    const disabled = await candidate.evaluate((el) =>
                        el.hasAttribute('disabled') ||
                        el.classList.contains('mdc-button--disabled') ||
                        el.getAttribute('aria-disabled') === 'true'
                    ).catch(() => false);
                    if (disabled) continue;
                    target = candidate;
                    break;
                }

                if (!target) {
                    throw new Error('Store listing contact details Save button not found.');
                }

                await clickLocatorRobust(target, 'Store listing contact details Save button', 15000);
                await delay(page, 3000);
            }, 'Click Store listing contact details Save', 3);
        }

        async function closeStoreListingContactDialogIfVisible() {
            const dialog = storeListingContactDialogLocator();
            if (!(await dialog.isVisible().catch(() => false))) {
                return;
            }

            console.log('[STORE SETTINGS] Closing Store listing contact details dialog...');
            const closeBtn = dialog.locator(
                'button[aria-label="Close"], button.close-icon-button, .close-icon-button, material-button[aria-label="Close"]'
            ).first();

            if (await closeBtn.isVisible().catch(() => false)) {
                await clickLocatorRobust(closeBtn, 'Store listing contact details Close button', 15000);
            } else {
                await page.keyboard.press('Escape').catch(() => { });
            }

            await dialog.waitFor({ state: 'hidden', timeout: 15000 }).catch(async () => {
                await page.keyboard.press('Escape').catch(() => { });
                await delay(page, 2000);
            });
            await randomDelay(page, 5000, 7000);
            console.log('[STORE SETTINGS] Store listing contact details dialog closed.');
        }

        async function closePanelIfVisible() {
            const closeBtn = page.locator(
                'button[aria-label="Close"], material-button[aria-label="Close"], .close-icon-button, button:has-text("Close")'
            ).first();
            if (await closeBtn.isVisible().catch(() => false)) {
                await clickLocatorRobust(closeBtn, 'Close panel button', 10000);
                await randomDelay(page, 2000, 3000);
            }
        }

        async function selectDropdownByDebugId(debugId, optionRegex, label) {
            const dropdown = page.locator(
                `material-dropdown-select[debug-id="${debugId}"] dropdown-button, ` +
                `material-dropdown-select[debug-id="${debugId}"] [role="button"]`
            ).first();
            await clickLocatorRobust(dropdown, `${label} dropdown`, 10000);
            await delay(page, 1200);

            const options = page.locator(
                'material-select-dropdown-item, material-option, material-list-item, [role="option"], .mdc-list-item'
            ).filter({ hasText: optionRegex });
            const optionCount = await options.count().catch(() => 0);
            let option = null;
            for (let i = 0; i < optionCount; i += 1) {
                const candidate = options.nth(i);
                if (await candidate.isVisible().catch(() => false)) {
                    option = candidate;
                    break;
                }
            }
            if (!option) {
                throw new Error(`${label} option not found: ${optionRegex}`);
            }
            await clickLocatorRobust(option, `${label} option`, 10000);
            await delay(page, 3000);
        }

        async function setStoreCategoryFromAppGenieType() {
            const typeKey = task.appGenieTypeKey || normalizeAppGenieTypeKey(task.appGenieTypeLabel);
            const isGame = typeKey === 'game';
            const appOrGame = isGame ? 'Game' : 'App';
            const category = isGame ? 'Casual' : 'Tools';
            console.log(`[STORE SETTINGS] Setting App category from AppGenie type: ${task.appGenieTypeLabel || 'unknown'} -> ${appOrGame}/${category}`);

            await gotoAppSubPage('/store-settings', 'Store settings for app category');
            await clickStoreSectionEdit(/App\s*category|App or game|Category/i, 'App category', 'last');
            await selectDropdownByDebugId('type-dropdown', new RegExp(`^\\s*${escapeRegExp(appOrGame)}\\s*$`, 'i'), 'App or game');
            await selectDropdownByDebugId('category-dropdown', new RegExp(`^\\s*${escapeRegExp(category)}\\s*$`, 'i'), 'Category');
            await clickMainButton('Save');
            await waitSaved(page);
            await closePanelIfVisible();
        }

        async function runStoreSettingsStep() {
            console.log('Executing item 11/13: Store settings...');
            await gotoAppSubPage('/store-settings', 'Store settings');
            await ensureStoreListingContactDialogOpen();
            await fillStoreListingContactEmail(runtimeOptions.contactEmail);
            await randomDelay(page, 2000, 3000);
            await clickStoreListingContactSave();
            await waitSaved(page);
            await closeStoreListingContactDialogIfVisible();
            await randomDelay(page, 5000, 7000);
            markStepDone(PROGRESS_STEP_STORE_SETTINGS_DONE);
        }

        async function runStoreListingStep() {
            console.log('Executing item 12/13: Store listing...');
            await gotoAppSubPage('/store-listings', 'Store listings');

            const createDefault = page.locator(
                'button[debug-id="get-started-create-default-listing-button"], button:has-text("Create default store listing")'
            ).first();
            if (await createDefault.isVisible().catch(() => false)) {
                await clickLocatorRobust(createDefault, 'Create default store listing button', 10000);
                await delay(page, 3000);
            } else {
                await gotoAppSubPage('/main-store-listing', 'Default store listing');
            }

            const shortDescription = ensureTrailingSpaceWithinLimit(
                shortenShortDescription(
                    await readAppGenieDetailCardText(/应用简介|Short\s*description|App\s*summary/i, 'AppGenie short description')
                ),
                80
            );
            const fullDescription = ensureTrailingSpace(
                await readAppGenieDetailCardText(/应用描述|Full\s*description|App\s*description/i, 'AppGenie full description')
            );

            await page.bringToFront().catch(() => { });
            await delay(page, 5000);
            console.log('[STORE LISTING] Filling short description...');
            await fillFirstVisibleInputOn(page, [
                'input[aria-label="Short description of the app"]',
                'input[aria-label*="Short description" i]',
                'input[maxlength="80"]',
                'material-input input.mdc-text-field__input'
            ], shortDescription, 'Short description', 5000);
            console.log('[STORE LISTING] Short description filled.');
            await delay(page, 5000);

            console.log('[STORE LISTING] Filling full description...');
            await fillFirstVisibleInputOn(page, [
                'textarea[aria-label="Full description of the app"]',
                'textarea[aria-label*="Full description" i]',
                'textarea[maxlength="4000"]',
                'textarea.mdc-text-field__input'
            ], fullDescription, 'Full description', 5000);
            console.log('[STORE LISTING] Full description filled.');
            await delay(page, 5000);

            const imageFiles = await downloadAndExtractStoreImages();
            await page.bringToFront().catch(() => { });
            await delay(page, 5000);
            console.log('[STORE LISTING] Moving to graphics section...');
            await page.locator('text=/Graphics/i').first().scrollIntoViewIfNeeded({ timeout: 10000 }).catch(async () => {
                await page.mouse.wheel(0, 1200).catch(() => { });
            });
            await delay(page, 5000);

            const phoneAddAssets = page.locator(
                'xpath=(//*[contains(normalize-space(.), "Phone screenshots")]/following::button[contains(normalize-space(.), "Add assets")])[1]'
            ).first();
            const addAssetsFallback = page.locator(
                'button:has-text("Add assets"), [role="button"]:has-text("Add assets"), material-button:has-text("Add assets")'
            ).nth(1);
            console.log('[STORE LISTING] Opening phone screenshots asset panel...');
            await clickFirstVisibleOn(page, [phoneAddAssets, addAssetsFallback], 'Phone screenshots Add assets', 3000);
            await delay(page, 5000);

            const uploadBtn = page.locator(
                'button:has-text("Upload"), [role="button"]:has-text("Upload"), material-button:has-text("Upload")'
            ).last();
            console.log('[STORE LISTING] Uploading phone screenshots...');
            await uploadFilesByUploadButton(page, uploadBtn, imageFiles, 'Phone screenshots');
            console.log('[STORE LISTING] Phone screenshots uploaded, waiting for processing...');
            await randomDelay(page, 10000, 15000);
            await addSelectedStoreAssetsToListing('Phone screenshots');
            await randomDelay(page, 5000, 7000);

            const saveDraft = page.locator(
                'button:has-text("Save as draft"), [debug-id="main-button"]:has-text("Save as draft"), button:has-text("Save")'
            ).last();
            if (await saveDraft.isVisible().catch(() => false) && !(await isLocatorDisabled(saveDraft))) {
                console.log('[STORE LISTING] Saving store listing...');
                await randomDelay(page, 5000, 7000);
                await clickLocatorRobust(saveDraft, 'Store listing Save button', 10000);
                await waitSaved(page);
                await randomDelay(page, 5000, 7000);
            } else {
                console.log('[STORE LISTING] Save button not enabled after upload, continuing with existing autosave state.');
                await randomDelay(page, 5000, 7000);
            }
            markStepDone(PROGRESS_STEP_STORE_LISTING_DONE);
        }

        async function waitForAabUploadReady(timeoutMs = 300000) {
            const deadline = Date.now() + timeoutMs;
            while (Date.now() < deadline) {
                const ready = await page.locator(
                    'text=/App bundle|Version code|Target SDK|View details for App bundle|Details/i'
                ).first().isVisible().catch(() => false);
                if (ready) {
                    console.log('[RELEASE] AAB upload detected.');
                    return;
                }
                await delay(page, 5000);
            }
            const error = new Error('AAB upload did not finish within 5 minutes. Pausing script.');
            error.stopRun = true;
            throw error;
        }

        async function openProductionReleaseEditor() {
            const uploadButton = page.locator(
                'button:has-text("Upload"), [role="button"]:has-text("Upload"), material-button:has-text("Upload")'
            ).first();

            await retryAction(async () => {
                if (await uploadButton.isVisible().catch(() => false)) {
                    return;
                }

                const actionButtons = [
                    page.locator('button[debug-id="create-release-button"], [role="button"][debug-id="create-release-button"]').first(),
                    page.locator('button:has-text("Create new release"), [role="button"]:has-text("Create new release")').first(),
                    page.locator('button:has-text("Create release"), [role="button"]:has-text("Create release")').first(),
                    page.locator('button:has-text("Edit release"), [role="button"]:has-text("Edit release")').first(),
                    page.locator('button:has-text("Continue editing"), [role="button"]:has-text("Continue editing")').first()
                ];

                let clicked = false;
                for (const actionButton of actionButtons) {
                    if (await actionButton.isVisible().catch(() => false)) {
                        await clickLocatorRobust(actionButton, 'Production release create/edit button', 10000);
                        clicked = true;
                        break;
                    }
                }

                if (!clicked) {
                    throw new Error('Production release create/edit button not found.');
                }

                await delay(page, 3000);
                await uploadButton.waitFor({ state: 'visible', timeout: 30000 });
            }, 'Open production release editor', 5);
            await delay(page, 3000);
            return uploadButton;
        }

        async function runReleaseStep() {
            console.log('Executing item 13/13: Production release...');
            console.log('[RELEASE] Downloading AAB package...');
            const aabPath = await downloadAabFromAppGenie();
            await randomDelay(page, 5000, 7000);

            console.log('[RELEASE] Opening Production page...');
            await gotoAppSubPage('/tracks/production', 'Production');
            await randomDelay(page, 5000, 7000);
            const uploadButton = await openProductionReleaseEditor();
            const releaseUrl = page.url();
            await randomDelay(page, 5000, 7000);

            console.log('[RELEASE] Uploading AAB package...');
            await uploadFilesByUploadButton(page, uploadButton, [aabPath], 'AAB');
            await randomDelay(page, 5000, 7000);
            await waitForAabUploadReady(300000);
            await randomDelay(page, 5000, 7000);

            console.log('[RELEASE] Updating Store settings category...');
            await setStoreCategoryFromAppGenieType();
            await randomDelay(page, 5000, 7000);

            console.log('[RELEASE] Returning to Production release editor...');
            await page.goto(releaseUrl, { timeout: 120000, waitUntil: 'domcontentloaded' });
            await randomDelay(page, 5000, 7000);
            console.log('[RELEASE] Clicking Next...');
            await clickMainButton('Next');
            await randomDelay(page, 5000, 7000);
            console.log('[RELEASE] Saving Production release...');
            await clickMainButton('Save');
            await randomDelay(page, 5000, 7000);

            const goOverview = page.locator(
                'button[debug-id="yes-button"], button:has-text("Go to overview"), [role="button"]:has-text("Go to overview")'
            ).first();
            if (await goOverview.isVisible().catch(() => false)) {
                console.log('[RELEASE] Going back to overview...');
                await clickLocatorRobust(goOverview, 'Go to overview button', 10000);
                await randomDelay(page, 5000, 7000);
            }

            markStepDone(PROGRESS_STEP_RELEASE_DONE);
        }

        async function runPrivacyPolicyStep() {
            await goToAppContentViaMonitorPolicyMenu();
            console.log('Executing declaration 8/13: Privacy policy...');
            await clickStartDeclaration('Privacy policy');

            const privacyText = await openAppGenieDetailsAndReadPrivacyText(context, task, runtimeOptions);
            const privacyUrl = await createAndPublishGoogleSite(context, task, privacyText);

            await page.bringToFront();
            await delay(page, 2000);
            await fillFirstVisibleInput(
                [
                    'input[aria-label="Privacy policy URL"]',
                    'input[aria-label*="Privacy policy"]',
                    'input[type="url"]',
                    'input[type="text"]'
                ],
                privacyUrl,
                'Privacy policy URL'
            );
            await clickMainButton('Save');
            await waitSaved(page);
            markStepDone(PROGRESS_STEP_PRIVACY_DONE);
        }

        async function selectNoInFirstUncheckedRadioGroup() {
            const result = await page.evaluate(() => {
                const groups = Array.from(document.querySelectorAll('material-radio-group'));

                const hasChecked = (group) => {
                    return !!group.querySelector(
                        'input[type="radio"]:checked, [role="radio"][aria-checked="true"], input[aria-checked="true"]'
                    );
                };

                for (let i = 0; i < groups.length; i++) {
                    const group = groups[i];
                    if (hasChecked(group)) continue;

                    const radios = Array.from(group.querySelectorAll('material-radio, [role="radio"]'));
                    if (!radios.length) continue;

                    let target = radios.find((r) => /\bNo\b/i.test((r.textContent || '').trim()));
                    if (!target && radios.length >= 2) {
                        target = radios[1];
                    }
                    if (!target) continue;

                    const label = target.querySelector('label');
                    const input = target.querySelector('input[type="radio"]');
                    const clickable = label || input || target.querySelector('.mdc-radio, [role="radio"]') || target;
                    clickable.click();

                    return {
                        clicked: true,
                        groupIndex: i + 1,
                        totalGroups: groups.length
                    };
                }

                return {
                    clicked: false,
                    groupIndex: 0,
                    totalGroups: groups.length
                };
            });

            return result || { clicked: false, groupIndex: 0, totalGroups: 0 };
        }

        async function getUncheckedRadioGroupCount() {
            const count = await page.evaluate(() => {
                const groups = Array.from(document.querySelectorAll('material-radio-group'));
                let remaining = 0;
                for (const group of groups) {
                    const hasChecked = !!group.querySelector(
                        'input[type="radio"]:checked, [role="radio"][aria-checked="true"], input[aria-checked="true"]'
                    );
                    if (!hasChecked) remaining += 1;
                }
                return remaining;
            }).catch(() => 0);
            return Number(count || 0);
        }

        async function runContentRatingsStep() {
            await goToAppContent();
            console.log('Executing declaration 9/13: Content ratings...');
            await clickStartDeclaration('Content ratings');

            const startQuestionnaireBtn = page.locator(
                '[debug-id="get-started-start-button"], button:has-text("Start questionnaire")'
            ).first();
            if (await startQuestionnaireBtn.isVisible().catch(() => false)) {
                await retryAction(async () => {
                    await startQuestionnaireBtn.click({ timeout: 10000 });
                }, 'Click Start questionnaire', 3);
                await delay(page, 4000);
            }

            await fillFirstVisibleInput(
                [
                    'input[type="email"]',
                    'input[aria-label*="mail" i]',
                    'input[aria-label*="email" i]'
                ],
                runtimeOptions.contactEmail,
                'Content ratings email'
            );

            console.log('Selecting category: All Other App Types...');
            await retryAction(async () => {
                const appTypeRadio = page.locator('material-radio').filter({
                    hasText: /All\s*Other\s*App\s*Types/i
                }).first();
                if (!(await appTypeRadio.isVisible().catch(() => false))) {
                    throw new Error('All Other App Types radio not visible.');
                }

                const isChecked = async () => {
                    const byInput = await appTypeRadio.locator('input[type="radio"]').first().isChecked().catch(() => false);
                    if (byInput) return true;
                    const byAria = await appTypeRadio.evaluate((el) => {
                        const input = el.querySelector('input[type="radio"]');
                        if (input && input.checked) return true;
                        if (el.getAttribute('aria-checked') === 'true') return true;
                        return !!el.querySelector('[role="radio"][aria-checked="true"], input[aria-checked="true"]');
                    }).catch(() => false);
                    return !!byAria;
                };

                await appTypeRadio.scrollIntoViewIfNeeded({ timeout: 5000 }).catch(() => { });
                if (!(await isChecked())) {
                    const label = appTypeRadio.locator('label:has-text("All Other App Types")').first();
                    if (await label.isVisible().catch(() => false)) {
                        await label.click({ timeout: 10000 });
                    } else {
                        await appTypeRadio.click({ timeout: 10000 });
                    }
                    await delay(page, 1000);
                }

                if (!(await isChecked())) {
                    const control = appTypeRadio.locator('input[type="radio"], .mdc-radio, [role="radio"]').first();
                    if (await control.isVisible().catch(() => false)) {
                        await control.click({ timeout: 10000 });
                        await delay(page, 1000);
                    }
                }

                if (!(await isChecked())) {
                    throw new Error('All Other App Types radio still unchecked.');
                }
            }, 'Select All other App Types', 3);

            console.log('Checking terms and conditions...');
            await retryAction(async () => {
                const termsCheckbox = page.locator('material-checkbox, [role="checkbox"]').filter({
                    hasText: /Terms of Use|terms|conditions/i
                }).first();
                if (!(await termsCheckbox.isVisible().catch(() => false))) {
                    throw new Error('Terms/conditions checkbox not visible.');
                }

                await termsCheckbox.scrollIntoViewIfNeeded({ timeout: 10000 });
                const input = termsCheckbox.locator('input[type="checkbox"]').first();
                let checked = await input.isChecked().catch(() => false);
                if (!checked) {
                    await termsCheckbox.click({ timeout: 10000 });
                    await delay(page, 1000);
                    checked = await input.isChecked().catch(() => false);
                }

                if (!checked) {
                    const ariaChecked = await termsCheckbox.evaluate((el) => {
                        const checkbox = el.querySelector('[role="checkbox"]');
                        if (checkbox && checkbox.getAttribute('aria-checked') === 'true') return true;
                        const inputEl = el.querySelector('input[type="checkbox"]');
                        return !!(inputEl && inputEl.checked);
                    }).catch(() => false);
                    checked = !!ariaChecked;
                }

                if (!checked) {
                    throw new Error('Terms/conditions checkbox still unchecked.');
                }
            }, 'Accept terms and conditions', 3);
            await clickMainButton('Next');

            let questionnaireSaved = false;
            let noSelectedCount = 0;
            for (let i = 0; i < 60; i++) {
                const clickedResult = await selectNoInFirstUncheckedRadioGroup();
                if (clickedResult.clicked) {
                    noSelectedCount += 1;
                    console.log(`[QUESTIONNAIRE] Selected "No" (${noSelectedCount}) for group ${clickedResult.groupIndex}/${clickedResult.totalGroups}.`);
                    await randomDelay(page, 2000, 3000);
                    continue;
                }

                const remainingGroups = await getUncheckedRadioGroupCount();
                if (remainingGroups > 0) {
                    await page.mouse.wheel(0, 800).catch(() => { });
                    await delay(page, 2000);
                    continue;
                }

                const saveBtn = page.locator(
                    'button[debug-id="save-button"], button:has-text("Save"), [debug-id="main-button"]:has-text("Save")'
                ).first();
                if (await saveBtn.isVisible().catch(() => false)) {
                    console.log(`[QUESTIONNAIRE] Selected "No" for ${noSelectedCount} group(s).`);
                    await clickMainButton('Save');
                    await waitSaved(page);
                    questionnaireSaved = true;
                    break;
                }

                await page.mouse.wheel(0, 600).catch(() => { });
                await delay(page, 2000);
            }

            if (!questionnaireSaved) {
                throw new Error('Content ratings questionnaire did not complete all No selections / Save.');
            }

            const nextAfterSave = page.locator(
                'button[debug-id="next-button"], button:has-text("Next"), [debug-id="main-button"]:has-text("Next")'
            ).first();
            if (await nextAfterSave.isVisible().catch(() => false)) {
                await delay(page, 3000);
                await clickMainButton('Next');
                const finalSave = page.locator(
                    'button[debug-id="save-button"], button:has-text("Save"), [debug-id="main-button"]:has-text("Save")'
                ).first();
                if (await finalSave.isVisible().catch(() => false)) {
                    await delay(page, 3000);
                    await clickMainButton('Save');
                    await waitSaved(page);
                }
            }

            markStepDone(PROGRESS_STEP_RATING_DONE);
        }

        async function runDataSafetyStep() {
            await goToAppContent();
            console.log('Executing declaration 10/13: Data safety...');
            await clickStartDeclaration('Data safety');

            const isButtonEnabled = async (locator) => {
                const visible = await locator.isVisible().catch(() => false);
                if (!visible) return false;
                const disabled = await locator.evaluate((el) => {
                    return el.hasAttribute('disabled') ||
                        el.classList.contains('mdc-button--disabled') ||
                        el.getAttribute('aria-disabled') === 'true';
                }).catch(() => true);
                return !disabled;
            };

            const clickScopedButton = async (locator, label) => {
                await locator.scrollIntoViewIfNeeded({ timeout: 5000 }).catch(() => { });
                await locator.click({ timeout: 10000 });
                console.log(`[DATA SAFETY] Clicked ${label}.`);
                await delay(page, 3000);
            };

            const getMainSaveButton = () => page
                .locator('button[debug-id="main-button"]')
                .filter({ hasText: /^\s*Save\s*$/ })
                .last();

            const clickMainSaveButton = async (label) => {
                const saveButton = getMainSaveButton();
                await saveButton.waitFor({ state: 'visible', timeout: 15000 });
                await clickLocatorRobust(saveButton, `Data safety ${label}`, 15000);
                console.log(`[DATA SAFETY] Clicked ${label}.`);
                await delay(page, 3000);
            };

            const getNoLocators = () => {
                const dataCollectionStep = page.locator('step-data-collection').first();
                const noGroup = dataCollectionStep
                    .locator('material-radio-group[debug-id="personal-data-no"]')
                    .first();
                const noRadio = noGroup.locator('material-radio').first();
                const noLabel = noGroup.locator('label').filter({ hasText: /^\s*No\s*$/i }).first();
                const noBackground = noGroup.locator('.mdc-radio__background').first();
                const noInput = noGroup.locator('input[type="radio"][role="radio"]').first();
                return { dataCollectionStep, noGroup, noRadio, noLabel, noBackground, noInput };
            };

            const isNoSelected = async () => {
                const { noGroup, noInput } = getNoLocators();
                const groupVisible = await noGroup.isVisible().catch(() => false);
                if (!groupVisible) return false;
                const byInputChecked = await noInput.isChecked().catch(() => false);
                if (byInputChecked) return true;
                const byInputAria = await noInput.evaluate((el) => el.getAttribute('aria-checked') === 'true').catch(() => false);
                if (byInputAria) return true;
                return await noGroup.evaluate((el) => {
                    const input = el.querySelector('input[type="radio"]');
                    if (input && input.checked) return true;
                    return !!el.querySelector('.mdc-radio--checked, [role="radio"][aria-checked="true"], input[aria-checked="true"], input[checked]');
                }).catch(() => false);
            };

            for (let i = 0; i < 30; i++) {
                const nextBtn = page.locator(
                    'button[debug-id="next-button"], button:has-text("Next"), [debug-id="main-button"]:has-text("Next"), button[debug-id="button-next"]'
                ).first();
                const saveBtn = getMainSaveButton();

                const { dataCollectionStep, noGroup, noLabel, noBackground, noInput, noRadio } = getNoLocators();
                const stepVisible = await dataCollectionStep.isVisible().catch(() => false);
                const noVisible = stepVisible && await noGroup.isVisible().catch(() => false);
                if (noVisible) {
                    if (!(await isNoSelected())) {
                        await noGroup.scrollIntoViewIfNeeded({ timeout: 10000 }).catch(() => { });
                        await delay(page, 1000);

                        if (await noInput.isVisible().catch(() => false)) {
                            await noInput.click({ timeout: 10000, force: true });
                        } else if (await noLabel.isVisible().catch(() => false)) {
                            await noLabel.click({ timeout: 10000, force: true });
                        } else if (await noBackground.isVisible().catch(() => false)) {
                            await noBackground.click({ timeout: 10000, force: true });
                        } else {
                            await noRadio.click({ timeout: 10000, force: true });
                        }

                        // DOM fallback for dynamic wrappers.
                        if (!(await isNoSelected())) {
                            await page.evaluate(() => {
                                const root = document.querySelector('step-data-collection material-radio-group[debug-id="personal-data-no"]');
                                if (!root) return;
                                const no = Array.from(root.querySelectorAll('label'))
                                    .find((l) => (l.textContent || '').trim().toLowerCase() === 'no');
                                if (!no) return;
                                const targetId = no.getAttribute('for');
                                const input = targetId ? document.getElementById(targetId) : root.querySelector('input[type="radio"]');
                                if (input && !input.checked) {
                                    input.click();
                                    input.dispatchEvent(new Event('input', { bubbles: true }));
                                    input.dispatchEvent(new Event('change', { bubbles: true }));
                                }
                                no.click();
                            }).catch(() => { });
                        }

                        await delay(page, 1500);
                        if (!(await isNoSelected())) {
                            throw new Error('Data safety "No" radio click did not take effect.');
                        }
                        console.log('[DATA SAFETY] "No" selected.');
                    }
                }

                // Follow document flow first: Next -> answer No if asked -> Next ... -> Save at end.
                if (await isButtonEnabled(nextBtn)) {
                    await clickScopedButton(nextBtn, 'Next');
                    continue;
                }

                if (await isButtonEnabled(saveBtn)) {
                    await clickMainSaveButton('Save (1/2)');
                    await waitSaved(page);
                    await delay(page, 3000);

                    const saveAfterFirst = getMainSaveButton();
                    if (!(await isButtonEnabled(saveAfterFirst))) {
                        console.log('[DATA SAFETY] Save button is disabled after first save; continuing to next step.');
                        markStepDone(PROGRESS_STEP_SAFETY_DONE);
                        return;
                    }

                    await retryAction(async () => {
                        await clickMainSaveButton('Save (2/2)');
                    }, 'Click Data safety second Save', 3);
                    await waitSaved(page).catch((secondSaveWaitError) => {
                        console.log(`[DATA SAFETY] Second Save confirmation not detected after click; continuing: ${secondSaveWaitError.message}`);
                    });
                    await delay(page, 3000);

                    markStepDone(PROGRESS_STEP_SAFETY_DONE);
                    return;
                }

                await page.mouse.wheel(0, 800).catch(() => { });
                await delay(page, 1500);
            }

            throw new Error('Data safety flow did not reach Save state.');
        }

        // --- Execute declarations flow in strict one-pass mode ---
        const markStepDone = (doneStep) => {
            statusManager.updateTaskProgress(task, doneStep);
        };

        await goToAppContent();
        console.log('Executing declaration 1/7: Ads...');
        await clickStartDeclaration('Ads');
        await selectRadio(/^No/);
        await clickMainButton('Save');
        await waitSaved(page);
        markStepDone(PROGRESS_STEP_ADS_DONE);

        await goToAppContent();
        console.log('Executing declaration 2/7: App access...');
        await clickStartDeclaration('App access');
        await selectRadio('All functionality in my app is available without any access restrictions');
        await clickMainButton('Save');
        await waitSaved(page);
        markStepDone(PROGRESS_STEP_APP_ACCESS_DONE);

        await goToAppContent();
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
        markStepDone(PROGRESS_STEP_AUDIENCE_DONE);

        await goToAppContent();
        console.log('Executing declaration 4/7: Advertising ID...');
        await clickStartDeclaration('Advertising ID');
        await selectRadio(/^Yes/);
        await selectCheckbox('Developer communications');
        await clickMainButton('Save');
        await waitSaved(page);
        await page.evaluate(() => window.scrollTo(0, 0));
        await delay(page, 2000);
        markStepDone(PROGRESS_STEP_AD_ID_DONE);

        await goToAppContent();
        console.log('Executing declaration 5/7: Government apps...');
        await clickStartDeclaration('Government apps');
        await selectRadio(/^No/);
        await clickMainButton('Save');
        await waitSaved(page);
        markStepDone(PROGRESS_STEP_GOV_DONE);

        await goToAppContent();
        console.log('Executing declaration 6/7: Financial features...');
        await clickStartDeclaration('Financial features');
        await selectCheckbox("My app doesn't provide any financial features");
        await clickMainButton('Next');
        await clickMainButton('Save');
        await waitSaved(page);
        markStepDone(PROGRESS_STEP_FINANCE_DONE);

        await goToAppContent();
        console.log('Executing declaration 7/7: Health apps...');
        await clickStartDeclaration('Health apps');
        await selectCheckbox('My app does not have any health features');
        await clickMainButton('Next');
        await clickMainButton('Save');
        await waitSaved(page);
        markStepDone(PROGRESS_STEP_HEALTH_DONE);

            // 8-11. Release navigation (strict: must really enter target pages before continuing)
            const testAndReleaseUrl = `${appBasePath}/test-and-release`;
            const productionUrl = `${appBasePath}/tracks/production`;
            const productionCountryAvailabilityUrl = `${productionUrl}?tab=countryAvailability`;
            const testAndReleaseLink = page.locator(
                'a[href*="/test-and-release"], a.item-link:has(.item-label:has-text("Test and release"))'
            ).first();
            const productionLink = page.locator(
                'a[href*="/tracks/production"], a.item-link:has(.item-label:has-text("Production"))'
            ).first();
            const countriesRegionsTab = page.locator(
                '[role="tab"]:has-text("Countries / regions"), a:has-text("Countries / regions"), button:has-text("Countries / regions")'
            ).first();
            const addCountriesBtn = page.locator(
                'button:has-text("Add countries / regions"), [role="button"]:has-text("Add countries / regions"), a:has-text("Add countries / regions"), .mdc-button:has-text("Add countries / regions")'
            ).first();

            // 8. Test and release
            console.log('Navigating to "Test and release"...');
            await retryAction(async () => {
                const linkVisible = await testAndReleaseLink.isVisible().catch(() => false);
                if (linkVisible) {
                    await testAndReleaseLink.scrollIntoViewIfNeeded({ timeout: 5000 }).catch(() => { });
                    await testAndReleaseLink.click({ timeout: 10000 });
                } else {
                    console.log('[RELEASE] Test and release link not visible, using direct URL fallback.');
                    await page.goto(testAndReleaseUrl, { timeout: 120000, waitUntil: 'domcontentloaded' });
                }

                await page.waitForURL(/\/test-and-release(?:\/|$)/, { timeout: 120000 });
                await productionLink.waitFor({ state: 'visible', timeout: 60000 });
            }, 'Open Test and release', 3);
            await delay(page, 5000);

            // 9. Production
            console.log('Navigating to "Production"...');
            await retryAction(async () => {
                const linkVisible = await productionLink.isVisible().catch(() => false);
                if (linkVisible) {
                    await productionLink.scrollIntoViewIfNeeded({ timeout: 5000 }).catch(() => { });
                    await productionLink.click({ timeout: 10000 });
                } else {
                    console.log('[RELEASE] Production link not visible, using direct URL fallback.');
                    await page.goto(productionUrl, { timeout: 120000, waitUntil: 'domcontentloaded' });
                }

                await page.waitForURL(/\/tracks\/production(?:\/|$)/, { timeout: 120000 });
                const countriesTabVisible = await countriesRegionsTab.isVisible().catch(() => false);
                const addButtonVisible = await addCountriesBtn.isVisible().catch(() => false);
                if (!countriesTabVisible && !addButtonVisible) {
                    throw new Error('Production page loaded but Countries/regions controls are not ready yet.');
                }
            }, 'Open Production', 3);
            await delay(page, 3000);

            // 10. Countries / regions tab (with direct URL fallback)
            console.log('Opening "Countries / regions" tab...');
            await retryAction(async () => {
                if (await addCountriesBtn.isVisible().catch(() => false)) {
                    return;
                }

                const tabVisible = await countriesRegionsTab.isVisible().catch(() => false);
                if (tabVisible) {
                    await countriesRegionsTab.scrollIntoViewIfNeeded({ timeout: 5000 }).catch(() => { });
                    await countriesRegionsTab.click({ timeout: 10000 });
                    await delay(page, 2000);
                }

                if (await addCountriesBtn.isVisible().catch(() => false)) {
                    return;
                }

                console.log('[RELEASE] Countries / regions tab not ready, using direct URL fallback...');
                await page.goto(productionCountryAvailabilityUrl, { timeout: 120000, waitUntil: 'domcontentloaded' });
                await delay(page, 2500);
                await addCountriesBtn.waitFor({ state: 'visible', timeout: 60000 });
            }, 'Open Countries / regions tab', 3);
            await delay(page, 2000);

            // 11. Add countries / regions
            console.log('Clicking "Add countries / regions"...');
            await retryAction(async () => {
                await addCountriesBtn.waitFor({ state: 'visible', timeout: 60000 });
                await addCountriesBtn.scrollIntoViewIfNeeded({ timeout: 5000 });
                await addCountriesBtn.click({ timeout: 10000 });
            }, 'Click Add countries / regions', 3);
            await delay(page, 5000);

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
            await delay(page, 3000);

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
            await runPrivacyPolicyStep();
            await runContentRatingsStep();
            await runDataSafetyStep();
            await runStoreSettingsStep();
            await runStoreListingStep();
            await runReleaseStep();

            statusManager.ensureTaskProgressAtLeast(task, PROGRESS_STEP_DONE);
            statusManager.updateTaskStatus(task, STATUS_DONE);
            task.status = STATUS_DONE;

        const durationSeconds = Math.round((Date.now() - startTime) / 1000);
        console.log(`Finished: ${appName} (${packageName}), Duration: ${formatDuration(durationSeconds)}`);
        // Keep the working tab for reuse in next iteration; only disconnect CDP client.
        await browser.close();
    } catch (err) {
        console.error(`[FATAL ERROR] In iteration: ${err.message}`);
        // Keep the working tab for reuse/debugging; only disconnect CDP client.
        if (browser) await browser.close().catch(() => { });
        throw err;
    }
}

(async () => {
    const { inputFileArg } = parseCliArgs();
    const runtimeOptions = getRuntimeOptions();
    const { appListUrl, configPath } = loadDeveloperConsoleAppListUrl();
    const inputFilePath = resolveInputExcelFile(inputFileArg);
    const runtimeInput = resolveRuntimeInput(inputFilePath);
    const { tasks, sheetName, statusManager } = loadTasksFromExcel(
        runtimeInput.runtimeWorkbookPath,
        runtimeInput.sourceFilePath
    );
    const selectedTasks = tasks;

    console.log(`Developer console URL loaded from: ${configPath}`);
    console.log(`Loaded ${tasks.length} rows from input file: ${path.basename(inputFilePath)} (sheet: ${sheetName})`);
    if (runtimeInput.isCsvInput) {
        console.log(`[CSV] Source file: ${runtimeInput.sourceFilePath}`);
        console.log('[CSV] Direct read mode enabled (no state file, no workbook output).');
    }
    console.log(`Will execute ${selectedTasks.length} iteration(s).`);

    if (!selectedTasks.length) {
        console.log('No pending rows to execute.');
        return;
    }

    const runStats = {
        totalLoaded: tasks.length,
        planned: selectedTasks.length,
        success: 0,
        successItems: [],
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
            await runOnce(task, appListUrl, statusManager, runtimeOptions);
            runStats.success += 1;
            runStats.successItems.push({
                appName: task.appName,
                packageName: task.packageName
            });
        } catch (e) {
            const errorCode = String((e && e.code) || '');
            const isCreateFailedToast = errorCode === 'CREATE_FAILED_TOAST';
            const isCreateFailedTimeout = errorCode === 'CREATE_FAILED_TIMEOUT';
            task.status = STATUS_FAILED;
            statusManager.updateTaskStatus(task, STATUS_FAILED);

            runStats.failed.push({
                appName: task.appName,
                packageName: task.packageName,
                reason: isCreateFailedToast
                    ? "Your app couldn't be created"
                    : isCreateFailedTimeout
                        ? 'Create app timed out (skipped this row).'
                        : String((e && e.message) || 'Unknown error')
            });

            console.error(`Iteration ${i + 1} failed but continuing...`);
            if (e && e.stopRun) {
                console.error(`[STOP] ${e.message}`);
                process.exitCode = 1;
                break;
            }
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
    const successNames = runStats.successItems.map(item => `${item.appName} (${item.packageName})`);
    const summaryLines = [
        `Total loaded: ${runStats.totalLoaded}`,
        `Planned: ${runStats.planned}`,
        `Success: ${runStats.success}`,
        `Failed: ${failedCount}`
    ];
    if (successNames.length > 0) {
        summaryLines.push('Successful apps:');
        for (const name of successNames) {
            summaryLines.push(`[OK] ${name}`);
        }
    }
    if (failedCount > 0) {
        summaryLines.push('Failed apps:');
        for (const name of failedNames) {
            summaryLines.push(`[FAIL] ${name}`);
        }
    }
    const summaryText = summaryLines.join('\n');
    const summaryPayload = {
        generatedAt: new Date().toISOString(),
        totalLoaded: runStats.totalLoaded,
        planned: runStats.planned,
        success: runStats.success,
        successItems: runStats.successItems,
        failedCount,
        failed: runStats.failed,
        summaryText
    };

    console.log('================ Run Summary ================');
    console.log(summaryText);
    console.log('=============================================');
    writeRunSummaryFile(summaryPayload);
})().catch(err => {
    console.error(`[INIT ERROR] ${err.message}`);
    process.exit(1);
});




