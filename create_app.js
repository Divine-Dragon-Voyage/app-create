const fs = require('fs');
const path = require('path');
/*
 * 主脚本总览：
 * 1. 从 Excel/CSV 读取应用任务，并维护 status / progress_step。
 * 2. 通过 Chrome CDP 接管已登录浏览器，进入 Google Play Console 创建应用。
 * 3. 依次完成 App content、国家地区、隐私政策、内容分级、Data safety、商店资料、Production release。
 * 4. 发布后读取 SHA-1、截取 Latest releases 页面，并回传到 AppGenie 提交审核。
 *
 * 维护提示：
 * - 本文件仍是主流程编排文件，页面选择器和复杂兜底逻辑尽量下沉到独立 flow 文件。
 * - status/progress_step 目前主要用于记录和排查，不是严格的断点续跑控制器。
 */
const crypto = require('crypto');
const os = require('os');
const { execFileSync } = require('child_process');
const XLSX = require('xlsx');
const { chromium } = require('playwright');
// Data safety 页面答案和稳定选择器集中在独立模块，避免主流程里散落大量常量。
const {
    DATA_SAFETY_COLLECTION_SECURITY_ACTIONS,
    DATA_SAFETY_COLLECTION_SECURITY_STEP_SELECTOR,
    DATA_SAFETY_DATA_TYPES_ACTIONS,
    DATA_SAFETY_DEVICE_IDS_CHECKBOX_TEXT,
    DATA_SAFETY_DEVICE_IDS_SYNC_WAIT_MS,
    DATA_SAFETY_PRIMARY_NEXT_BUTTON_SELECTOR,
    DATA_SAFETY_NEXT_BUTTON_SELECTORS,
    DATA_SAFETY_USAGE_ACTIONS,
    DATA_SAFETY_SECTION_SELECTORS,
    DATA_SAFETY_OUTSIDE_APP_LOGIN_GROUP_SELECTORS,
    DATA_SAFETY_DATA_DELETION_GROUP_SELECTORS,
    DATA_SAFETY_OUTSIDE_APP_LOGIN_NO_RADIO_INDEX,
    DATA_SAFETY_DATA_DELETION_NO_RADIO_INDEX,
    DATA_SAFETY_DIRECT_MATERIAL_RADIO_SELECTOR,
    DATA_SAFETY_OUTSIDE_APP_LOGIN_NO_ANSWER_TEXT,
    DATA_SAFETY_DATA_DELETION_NO_ANSWER_TEXT,
    pickLastVisibleDataSafetyNextButton
} = require('./data_safety_flow');
// 清理相关逻辑：临时下载根目录和辅助标签页识别。
const {
    buildTempDownloadRoot,
    cleanupTrackedFallbackAabFiles,
    shouldCleanupFallbackAabAfterRow,
    shouldCloseAuxiliaryPage
} = require('./cleanup_helpers');
// Google Sites 页面结构经常变化，删除 header 的选择器单独维护。
const {
    GOOGLE_SITES_DELETE_HEADER_SELECTORS
} = require('./google_sites_flow');
// Play Console 部分页面 Save 会藏在右下角三点菜单里。
const {
    OVERFLOW_MORE_OPTIONS_SELECTORS,
    OVERFLOW_SAVE_MENU_SELECTORS
} = require('./overflow_save_flow');
// 发布后提审回传流程：Play App signing、截图路径、AppGenie 提审弹窗选择器。
const {
    APP_SIGNING_PLAY_STORE_PROTECTION_CARD_SELECTORS,
    APP_SIGNING_PLAY_STORE_PROTECTION_EXPAND_SELECTORS,
    APP_SIGNING_MANAGE_BUTTON_SCOPED_SELECTORS,
    APPGENIE_REVIEW_FILE_INPUT_SELECTOR,
    APPGENIE_REVIEW_SHA1_INPUT_SELECTOR,
    APPGENIE_REVIEW_SUBMIT_BUTTON_SELECTORS,
    PLAY_PROTECTED_WITH_PLAY_PATH,
    PLAY_RELEASES_OVERVIEW_PATH,
    buildReviewScreenshotPath,
    extractSha1Fingerprint
} = require('./review_submission_flow');
// Release 错误处理：默认语言、AD_ID release error、Production Releases tab 等选择器。
const {
    AD_ID_ACK_CHECKBOX_SELECTOR,
    CREATE_APP_EN_US_OPTION_SELECTORS,
    CREATE_APP_LANGUAGE_BUTTON_SELECTORS,
    CREATE_APP_LANGUAGE_DROPDOWN_SELECTOR,
    PRODUCTION_EDIT_RELEASE_SELECTORS,
    PRODUCTION_RELEASES_TAB_SELECTORS,
    RELEASE_UPDATE_DECLARATION_SELECTORS,
    hasBlockingReleaseErrorText,
    isAdIdReleasePermissionError,
    isEnUsLanguageText
} = require('./release_error_flow');

// 默认等待时间。VPS 慢或页面加载不稳定时，优先调具体步骤，不建议盲目调大这个全局值。
const DELAY = 3000;

// 输入表头兼容：支持中文列名和英文列名。
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

// Google Play 包名格式校验和运行时环境变量。
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

// 任务状态：PARTIAL 表示任务已开始但未完成；NEED_FIX 表示需要人工介入。
const STATUS_HEADER_CANDIDATES = new Set(['status', '\u72B6\u6001']);
const PROGRESS_HEADER_CANDIDATES = new Set(['progressstep', 'progress', '\u8fdb\u5ea6', '\u6b65\u9aa4']);
const STATUS_PARTIAL = 'PARTIAL';
const STATUS_DONE = 'DONE';
const STATUS_FAILED = 'FAILED';
const STATUS_NEED_FIX = 'NEED_FIX';

// 进度步骤按实际执行顺序排列，用于日志记录和 ensureTaskProgressAtLeast 比较。
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
const PROGRESS_STEP_REVIEW_UPLOAD_DONE = 'REVIEW_UPLOAD_DONE';
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
    PROGRESS_STEP_REVIEW_UPLOAD_DONE,
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

// 解析命令行输入文件路径：启动器通常传入显式路径，命令行也可省略后自动找根目录数据文件。
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

// 配置目录由启动器通过环境变量指定；本地直接运行时默认使用当前项目目录。
function resolveConfigDirectory() {
    const configuredDir = String(process.env[CONFIG_DIR_ENV] || '').trim();
    if (!configuredDir) {
        return process.cwd();
    }
    return path.resolve(configuredDir);
}

// 首次运行时自动生成 developer_url.txt 模板，避免用户不知道该填什么。
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

// 将用户粘贴的开发者 URL 归一化成 developers/<id> 基础路径，后续拼 app-list 和 app 子页面。
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

// 开发者入口优先级：完整 URL 环境变量 > developer_id 环境变量 > developer_url.txt。
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

// 输入文件可显式指定；未指定时从项目根目录挑第一个 xlsx/xls/csv。
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

// CSV 不直接回写原文件，状态独立写入配置目录下的 csv-state JSON。
function buildCsvStateFilePath(csvPath) {
    const configDir = resolveConfigDirectory();
    const stateDir = path.resolve(configDir, CSV_STATE_DIR_NAME);
    fs.mkdirSync(stateDir, { recursive: true });

    const key = path.resolve(csvPath).toLowerCase();
    const hash = crypto.createHash('sha1').update(key).digest('hex');
    return path.join(stateDir, `${hash}.json`);
}

// 运行期输入描述。当前 CSV 仍直接读取源文件，状态通过 JSON 旁路维护。
function resolveRuntimeInput(inputFilePath) {
    const ext = path.extname(inputFilePath).toLowerCase();
    return {
        sourceFilePath: inputFilePath,
        runtimeWorkbookPath: inputFilePath,
        isCsvInput: ext === '.csv'
    };
}

// 读取 CSV 状态文件；文件不存在或损坏时返回空状态，避免阻断主流程。
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

// 保存 CSV 状态：以包名为 key，记录 status/progress_step 和基础展示信息。
function saveCsvStateRows(stateFilePath, sourceFilePath, rows) {
    if (!stateFilePath) return;

    const payload = {
        sourceFilePath,
        updatedAt: new Date().toISOString(),
        rows
    };
    fs.writeFileSync(stateFilePath, JSON.stringify(payload, null, 2), 'utf8');
}

// 包名大小写统一归一化，避免 CSV 状态匹配漂移。
function getCsvStateKey(packageName) {
    return String(packageName || '').trim().toLowerCase();
}

// 从表头列表中挑出候选列，目前保留给后续扩展使用。
function pickHeader(headers, candidates) {
    for (const header of headers) {
        if (candidates.has(normalizeHeader(header))) {
            return header;
        }
    }
    return null;
}

// 状态值严格归一化，只接受脚本认识的状态，防止 Excel 手工填写污染流程。
function normalizeStatusValue(value) {
    const text = String(value || '').trim().toUpperCase();
    if (text === STATUS_DONE) return STATUS_DONE;
    if (text === STATUS_PARTIAL) return STATUS_PARTIAL;
    if (text === STATUS_FAILED) return STATUS_FAILED;
    if (text === STATUS_NEED_FIX) return STATUS_NEED_FIX;
    return '';
}

// 进度值严格按 PROGRESS_STEP_ORDER 识别。
function normalizeProgressStep(value) {
    const text = String(value || '').trim().toUpperCase();
    if (PROGRESS_STEP_SET.has(text)) return text;
    return '';
}

// 比较两个进度步骤的先后顺序。
function progressRank(step) {
    return PROGRESS_STEP_ORDER.indexOf(normalizeProgressStep(step));
}

// 可选输出机器可读的运行汇总，供启动器或外部监控读取。
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

// xlsx 单元格统一以 trim 后字符串读取。
function getCellText(sheet, rowIndex, colIndex) {
    const address = XLSX.utils.encode_cell({ r: rowIndex, c: colIndex });
    const cell = sheet[address];
    if (!cell || cell.v === undefined || cell.v === null) {
        return '';
    }
    return String(cell.v).trim();
}

// xlsx 单元格统一以字符串写入，避免数字/公式类型影响后续读取。
function setCellText(sheet, rowIndex, colIndex, text) {
    const address = XLSX.utils.encode_cell({ r: rowIndex, c: colIndex });
    sheet[address] = { t: 's', v: String(text || '') };
}

// Excel 状态管理器：直接把 status/progress_step 回写到工作簿。
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

// CSV 状态管理器：原始 CSV 不回写，状态落在 csv-state JSON。
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

// 空状态管理器主要用于测试或临时运行，不做持久化。
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

// 加载任务表并补齐状态列；同时完成应用名、包名的基础校验。
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

// 手动登录等待时间：给 Google 账号选择、二次验证或 Play Console 跳转留出窗口。
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

// Playwright 页面等待统一入口，便于后续替换为条件式等待。
async function delay(page, ms = DELAY) {
    await page.waitForTimeout(ms);
}

// 为批量运行加入轻微随机性，降低固定节奏导致页面状态未稳定的问题。
function randomInt(min, max) {
    const lo = Math.min(min, max);
    const hi = Math.max(min, max);
    return Math.floor(Math.random() * (hi - lo + 1)) + lo;
}

async function randomDelay(page, minMs, maxMs) {
    await delay(page, randomInt(minMs, maxMs));
}

// 等待按钮从 disabled 变可用，常用于 Play Console 保存/下一步按钮。
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

// CDP 端点可由环境变量覆盖；默认优先 IPv4，避免 localhost 解析到 ::1 连接失败。
function getCdpEndpoints() {
    const configured = String(process.env[CDP_ENDPOINT_ENV] || '').trim();
    if (configured) {
        return [configured];
    }
    // Prefer IPv4 first to avoid localhost -> ::1 connection failures.
    return ['http://127.0.0.1:9222', 'http://localhost:9222'];
}

// CDP 连接失败属于整批运行环境问题，遇到后不应继续盲跑下一条。
function isCdpConnectionError(err) {
    const message = String((err && err.message) || '');
    return /connectOverCDP|ECONNREFUSED|9222|CDP/i.test(message);
}

// 连接用户已经打开的 Chrome，而不是启动新的无状态浏览器。
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

// 最大化窗口可减少响应式布局导致的按钮不可见或左侧导航折叠问题。
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

// UI 操作通用重试：页面偶发慢、点击被吞时先重试，再把最终错误交给上层处理。
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

// 点击兜底顺序：普通点击 -> force 点击 -> DOM click/MouseEvent。
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

// Material/Ant/原生按钮的 disabled 状态不完全统一，统一用这个函数判断。
async function isLocatorDisabled(locator) {
    const visible = await locator.isVisible().catch(() => false);
    if (!visible) return true;
    return await locator.evaluate((el) => {
        return el.hasAttribute('disabled') ||
            el.classList.contains('mdc-button--disabled') ||
            el.getAttribute('aria-disabled') === 'true';
    }).catch(() => true);
}

// 保存后只宽松等待 saved 文案；某些页面保存成功但提示不出现，所以这里不抛错。
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

// 启动器传入的必要字段统一校验，缺失时给用户可理解的字段名。
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

// AppGenie 登录和联系邮箱来自启动器表单。
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

// 多标签环境下挑一个最像 Play Console 工作页的标签复用。
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

// 获取 Play Console 工作页；没有可复用标签时创建新标签。
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

// 生成 Google Sites 地址 slug：只保留英文数字并加随机数字，降低重名概率。
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

// 组装正则前转义用户/页面文本。
function escapeRegExp(value) {
    return String(value || '').replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

// 文件名安全化，避免 Windows 路径中出现非法字符。
function safeFileToken(value) {
    return String(value || 'app')
        .replace(/[^A-Za-z0-9._-]+/g, '_')
        .replace(/^_+|_+$/g, '')
        .slice(0, 80) || 'app';
}

// PowerShell 单引号字符串转义，用于 Expand-Archive 命令。
function escapePowerShellSingleQuotedString(value) {
    return String(value || '').replace(/'/g, "''");
}

// Windows 下用系统 PowerShell 解压 AppGenie 下载的素材 ZIP。
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

// 递归收集素材文件，后续按固定文件名挑图。
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

// AAB 下载兜底扫描浏览器默认下载目录。
function getDefaultDownloadDirectories() {
    const candidates = [
        process.env.USERPROFILE ? path.join(process.env.USERPROFILE, 'Downloads') : '',
        path.join(os.homedir(), 'Downloads')
    ];

    const seen = new Set();
    return candidates
        .map(dir => path.resolve(dir))
        .filter(dir => {
            if (!dir || seen.has(dir)) return false;
            seen.add(dir);
            return fs.existsSync(dir);
        });
}

// 从下载目录找最近完成的 AAB 文件，并优先匹配包名 token。
function findRecentAabDownloads(downloadDirs, startedAtMs, expectedToken) {
    const token = String(expectedToken || '').toLowerCase();
    const earliestMtimeMs = Number(startedAtMs || 0) - 5000;
    const candidates = [];

    for (const dir of downloadDirs) {
        let entries = [];
        try {
            entries = fs.readdirSync(dir, { withFileTypes: true });
        } catch (_) {
            continue;
        }

        for (const entry of entries) {
            if (!entry.isFile()) continue;
            const filePath = path.join(dir, entry.name);
            const lowerName = entry.name.toLowerCase();
            if (!lowerName.endsWith('.aab')) continue;

            let stat;
            try {
                stat = fs.statSync(filePath);
            } catch (_) {
                continue;
            }

            if (stat.size <= 0 || stat.mtimeMs < earliestMtimeMs) continue;
            candidates.push({
                filePath,
                size: stat.size,
                mtimeMs: stat.mtimeMs,
                score: token && lowerName.includes(token) ? 1 : 0
            });
        }
    }

    candidates.sort((a, b) => {
        if (b.score !== a.score) return b.score - a.score;
        return b.mtimeMs - a.mtimeMs;
    });
    return candidates;
}

// 检测 Chrome 临时下载文件，避免把未下载完的 AAB 当成完成文件。
function hasActiveChromePartialDownload(downloadDirs, startedAtMs, expectedToken) {
    const token = String(expectedToken || '').toLowerCase();
    const earliestMtimeMs = Number(startedAtMs || 0) - 5000;

    for (const dir of downloadDirs) {
        let entries = [];
        try {
            entries = fs.readdirSync(dir, { withFileTypes: true });
        } catch (_) {
            continue;
        }

        for (const entry of entries) {
            if (!entry.isFile()) continue;
            const lowerName = entry.name.toLowerCase();
            if (!lowerName.endsWith('.crdownload')) continue;
            if (token && !lowerName.includes(token) && !lowerName.includes('.aab')) continue;

            try {
                const stat = fs.statSync(path.join(dir, entry.name));
                if (stat.mtimeMs >= earliestMtimeMs) {
                    return true;
                }
            } catch (_) {
                // Ignore files that disappear while Chrome finishes the download.
            }
        }
    }

    return false;
}

// 如果 Playwright download 事件没捕获到，就从磁盘下载目录等 AAB 稳定完成。
async function waitForRecentAabDownloadFromDisk({
    startedAtMs,
    expectedPackageName,
    timeoutMs = 300000,
    stableMs = 5000
}) {
    const downloadDirs = getDefaultDownloadDirectories();
    if (!downloadDirs.length) {
        const error = new Error('No existing browser download directory found for AAB fallback scan.');
        error.stopRun = true;
        throw error;
    }

    console.log(`[APPGENIE] Watching download folder(s) for AAB: ${downloadDirs.join(', ')}`);
    const deadline = Date.now() + timeoutMs;
    let lastPath = '';
    let lastSize = -1;
    let stableSinceMs = 0;

    while (Date.now() < deadline) {
        const activePartial = hasActiveChromePartialDownload(downloadDirs, startedAtMs, expectedPackageName);
        const [candidate] = findRecentAabDownloads(downloadDirs, startedAtMs, expectedPackageName);

        if (candidate) {
            if (candidate.filePath !== lastPath || candidate.size !== lastSize) {
                lastPath = candidate.filePath;
                lastSize = candidate.size;
                stableSinceMs = Date.now();
                console.log(`[APPGENIE] AAB fallback candidate found: ${candidate.filePath}`);
            } else if (!activePartial && Date.now() - stableSinceMs >= stableMs) {
                console.log(`[APPGENIE] AAB fallback download completed: ${candidate.filePath}`);
                return candidate.filePath;
            }
        }

        await new Promise(resolve => setTimeout(resolve, 2500));
    }

    const error = new Error(
        `AAB download was not detected in browser download folder within ${Math.round(timeoutMs / 1000)}s.`
    );
    error.stopRun = true;
    throw error;
}

// 从素材 ZIP 里按约定文件名取截图：1/2/3 等。
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

// AppGenie 卡片文本归一化，用于描述和类型读取。
function normalizeAppGenieCardText(value) {
    return String(value || '')
        .replace(/\r\n/g, '\n')
        .replace(/[ \t]+\n/g, '\n')
        .trim();
}

// Play Store 短描述保存时尾部需要保留空格，避免某些输入框最后一个词未触发变更。
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

// 短描述超长时做保守压缩，优先替换冗词，再按词边界截断。
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

// 获取或创建辅助页，例如 AppGenie；可复用已有标签减少重复登录。
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

// 创建全新的辅助页，例如每次生成 Google Sites 都需要干净页面。
async function createFreshAuxPage(context, urlToOpen, label) {
    const created = await context.newPage();
    await created.bringToFront().catch(() => { });
    await created.goto(urlToOpen, { timeout: 120000, waitUntil: 'domcontentloaded' });
    await delay(created, 2500);
    return { page: created, source: `created-fresh-${label}` };
}

// 每条任务前后清理 AppGenie、Sites 等辅助标签，保留 Play Console 主标签。
async function closeAutomationAuxiliaryPages(context, keepPage = null, label = 'cleanup') {
    if (!context) {
        return;
    }

    const keepUrl = keepPage && !keepPage.isClosed() ? String(keepPage.url() || '') : '';
    const pagesToClose = context.pages()
        .filter(candidate => !candidate.isClosed())
        .filter(candidate => candidate !== keepPage)
        .filter(candidate => shouldCloseAuxiliaryPage(candidate.url(), keepUrl));

    if (!pagesToClose.length) {
        console.log(`[CLEANUP] No auxiliary automation tabs to close (${label}).`);
        return;
    }

    console.log(`[CLEANUP] Closing ${pagesToClose.length} auxiliary automation tab(s) (${label}).`);
    for (const candidate of pagesToClose) {
        const candidateUrl = candidate.url();
        try {
            await candidate.close({ runBeforeUnload: false });
            console.log(`[CLEANUP] Closed auxiliary tab: ${candidateUrl || 'about:blank'}`);
        } catch (closeError) {
            console.log(`[CLEANUP] Could not close auxiliary tab (${candidateUrl || 'about:blank'}): ${closeError.message}`);
        }
    }

    if (keepPage && !keepPage.isClosed()) {
        await keepPage.bringToFront().catch(() => { });
    }
}

// 清理本批任务下载的临时素材，避免长批次磁盘和文件名污染。
function cleanupTempDownloadRoot() {
    const tempDownloadRoot = buildTempDownloadRoot(os.tmpdir());
    if (!fs.existsSync(tempDownloadRoot)) {
        return;
    }

    try {
        fs.rmSync(tempDownloadRoot, { recursive: true, force: true });
        console.log(`[CLEANUP] Removed temp download directory: ${tempDownloadRoot}`);
    } catch (err) {
        console.log(`[CLEANUP] Could not remove temp download directory (${tempDownloadRoot}): ${err.message}`);
    }
}

// AppGenie 登录态检查与自动登录；登录失败会阻断当前任务。
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

// 确保 AppGenie 位于“我的任务”页面，后续按邮箱和包名找任务。
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

// AppGenie 类型标签归一化为 Play Console 的 App/Game 分支。
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

// 从 AppGenie 应用卡片读取类型标签，供 Store settings 设置分类使用。
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

// 打开 AppGenie 详情页，读取隐私政策文本，并记录素材下载要复用的详情页。
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
    if (await emailRow.isVisible().catch(() => false)) {
        console.log(`[APPGENIE] Task account found on current page: ${runtimeOptions.contactEmail}`);
    } else {
        const searchInput = appGeniePage.locator(
            'input[placeholder*="搜索邮箱"], input[placeholder*="邮箱"], input[placeholder*="email" i], input.ant-input'
        ).first();
        await retryAction(async () => {
            await searchInput.waitFor({ state: 'visible', timeout: 30000 });
            await searchInput.click({ timeout: 10000 });
            await searchInput.fill('');
            await searchInput.type(runtimeOptions.contactEmail, { delay: 40 });
            await appGeniePage.keyboard.press('Enter').catch(() => { });
        }, 'Search AppGenie tasks by contact email', 3);
        console.log(`[APPGENIE] Searching task account by email: ${runtimeOptions.contactEmail}`);
        await randomDelay(appGeniePage, 8000, 12000);

        emailRow = appGeniePage.locator('tr.ant-table-row').filter({ hasText: runtimeOptions.contactEmail }).first();
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

// 在 Google Sites 创建隐私政策站点，发布后返回可预测的公开 URL。
async function createAndPublishGoogleSite(context, task, privacyText) {
    const { page: sitesPage, source } = await createFreshAuxPage(
        context,
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
    const deleteHeaderBtn = sitesPage.locator(GOOGLE_SITES_DELETE_HEADER_SELECTORS.join(', ')).first();

    console.log('[STEP] Clearing Sites header block by trash button...');
    const headerVisible = await headerArea.isVisible().catch(() => false);
    if (headerVisible) {
        await retryAction(async () => {
            await headerArea.waitFor({ state: 'visible', timeout: 15000 });
            await headerArea.click({ timeout: 10000 });
            await delay(sitesPage, 320);
            await deleteHeaderBtn.scrollIntoViewIfNeeded({ timeout: 5000 }).catch(() => { });
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
    // Google Sites 发布弹窗偶发加载较慢，这里放宽到 90 秒再判定超时。
    await publishDialog.waitFor({ state: 'visible', timeout: 90000 });

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
        await closeAutomationAuxiliaryPages(context, page, 'before task');
        cleanupTempDownloadRoot();

        let appBasePath = '';
        const extractAppBasePathFromCurrentUrl = () => {
            const currentUrl = page.url();
            const baseMatch = currentUrl.match(/(.*\/app\/\d+)/);
            if (!baseMatch) {
                throw new Error('Could not extract app base path from URL: ' + currentUrl);
            }
            return baseMatch[1];
        };

        // 当前任务收尾或异常时清理辅助标签页，保留主 Play Console 页用于下一条任务。
        async function closeOtherTabsAfterTaskDone() {
            const contextToClean = page && !page.isClosed()
                ? page.context()
                : (browser && browser.contexts()[0]);
            if (!contextToClean) {
                return;
            }
            await closeAutomationAuxiliaryPages(contextToClean, page, 'after task');
        }

        const manualLoginWaitMs = getManualLoginWaitMs();
        // 打开 Play Console app-list；如果遇到登录页，就等待用户手动完成登录/选账号。
        const openAppListPage = async () => {
            const createBtn = page.locator('[debug-id="create-app-button"], a:has-text("Create app"), button:has-text("Create app")').first();

            // 某些账号会先进入开发者选择器，检测到后点第一个开发者项。
            const ensureDeveloperSelectedIfNeeded = async () => {
                const devItem = page.locator('developer-item, [debug-id="all-developers"]').first();
                if (await devItem.isVisible().catch(() => false)) {
                    console.log('Developer picker detected, clicking first developer item...');
                    await devItem.click();
                    await delay(page, 7000);
                }
            };

            // 用 URL、输入框和登录文案综合判断当前是否还在 Google 登录流程。
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

            // 登录等待循环：定期打印剩余时间，并在回到 Console 非 app-list 时重定向。
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

        // Create app 按钮偶发点击无响应，这里重新打开 app-list 后多次尝试。
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

        // 读取创建应用表单中的默认语言，用于判断是否需要从 en-GB 切回 en-US。
        const readCreateAppLanguageText = async () => {
            const dropdown = page.locator(CREATE_APP_LANGUAGE_DROPDOWN_SELECTOR).first();
            const buttonText = await dropdown.locator('.button-text').first().innerText().catch(() => '');
            const ariaText = await dropdown.locator('[role="button"]').first().getAttribute('aria-label').catch(() => '');
            return String(`${buttonText} ${ariaText}`).replace(/\s+/g, ' ').trim();
        };

        // 创建阶段强制默认语言为 en-US，避免商店列表默认语言被账号环境带偏。
        const ensureCreateAppDefaultLanguageEnUs = async () => {
            await retryAction(async () => {
                const dropdown = page.locator(CREATE_APP_LANGUAGE_DROPDOWN_SELECTOR).first();
                await dropdown.waitFor({ state: 'visible', timeout: 20000 });

                const currentLanguage = await readCreateAppLanguageText();
                if (isEnUsLanguageText(currentLanguage)) {
                    console.log('[CREATE] Default language is already en-US.');
                    return;
                }

                console.log(`[CREATE] Default language is "${currentLanguage || 'unknown'}"; selecting en-US...`);
                let languageButton = null;
                for (const selector of CREATE_APP_LANGUAGE_BUTTON_SELECTORS) {
                    const candidate = page.locator(selector).first();
                    if (await candidate.isVisible().catch(() => false)) {
                        languageButton = candidate;
                        break;
                    }
                }
                if (!languageButton) {
                    throw new Error('Default language dropdown button not found.');
                }

                await clickLocatorRobust(languageButton, 'Default language dropdown', 10000);
                await delay(page, 1000);

                let enUsOption = null;
                for (const selector of CREATE_APP_EN_US_OPTION_SELECTORS) {
                    const candidate = page.locator(selector).first();
                    if (await candidate.isVisible().catch(() => false)) {
                        enUsOption = candidate;
                        break;
                    }
                }
                if (!enUsOption) {
                    throw new Error('English (United States) - en-US option not found.');
                }

                await clickLocatorRobust(enUsOption, 'English (United States) - en-US option', 10000);
                await delay(page, 1500);

                const updatedLanguage = await readCreateAppLanguageText();
                if (!isEnUsLanguageText(updatedLanguage)) {
                    throw new Error(`Default language is still not en-US: ${updatedLanguage || 'unknown'}`);
                }
            }, 'Ensure default language en-US', 3);
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

            await ensureCreateAppDefaultLanguageEnUs();
            await delay(page, 1500);

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
        // 直接进入 App content overview，并优先切到 Need attention 标签。
        async function goToAppContent() {
            await page.goto(appBasePath + '/app-content/overview', { timeout: 90000, waitUntil: 'domcontentloaded' });
            await delay(page, 8000);
            const tab = page.locator('div[role="tab"]:has-text("Need attention")');
            if (await tab.isVisible().catch(() => false)) {
                await tab.click().catch(() => { });
                await delay(page, 3000);
            }
        }

        // 通过左侧 Monitor and improve -> Policy and programs -> App content 进入，失败时回落到直达 URL。
        async function goToAppContentViaMonitorPolicyMenu() {
            const policyAndProgramsPattern = /Policy and program(?:mes|s)/i;
            const monitorAndImproveLink = page.locator(
                'a[href*="/monitor"], a.item-link:has(.item-label:has-text("Monitor and improve")), [role="button"]:has-text("Monitor and improve")'
            ).first();
            const policyAndProgramsLink = page.locator(
                'div[role="listitem"], a.item-link, [role="button"]'
            ).filter({
                hasText: policyAndProgramsPattern
            }).first();
            const policyAndProgramsDirectLink = page.locator(
                'a.item-link, [role="button"]'
            ).filter({
                hasText: policyAndProgramsPattern
            }).first();
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
                    const policyVisible = await policyAndProgramsLink.isVisible().catch(() => false) ||
                        await policyAndProgramsDirectLink.isVisible().catch(() => false);
                    if (policyVisible) {
                        const policyLink = await policyAndProgramsDirectLink.isVisible().catch(() => false)
                            ? policyAndProgramsDirectLink
                            : policyAndProgramsLink;
                        await policyLink.scrollIntoViewIfNeeded({ timeout: 5000 }).catch(() => { });
                        await policyLink.click({ timeout: 10000 });
                        await delay(page, 1500);
                    } else {
                        console.log('[RELEASE] Policy and programs link not visible, using direct App content URL fallback.');
                        await page.goto(appBasePath + '/app-content/overview', { timeout: 90000, waitUntil: 'domcontentloaded' });
                        await delay(page, 8000);
                    }
                }
                const appContentVisibleAfter = await appContentLink.isVisible().catch(() => false);
                if (!appContentVisibleAfter) {
                    console.log('[RELEASE] App content link still not visible, using direct App content URL fallback.');
                    await page.goto(appBasePath + '/app-content/overview', { timeout: 90000, waitUntil: 'domcontentloaded' });
                    await delay(page, 8000);
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

        function getDeclarationTitleVariants(sectionTitle) {
            const variants = [sectionTitle];
            const normalizedTitle = String(sectionTitle || '').trim().toLowerCase();
            if (normalizedTitle === 'sign-in details' || normalizedTitle === 'sign in details') {
                variants.push('Sign-in details', 'Sign in details');
            }
            return [...new Set(variants)];
        }

        // 在 App content 列表中打开指定声明项。
        async function clickStartDeclaration(sectionTitle) {
            const sectionTitleVariants = getDeclarationTitleVariants(sectionTitle);
            const buttonByAria = page.locator(
                sectionTitleVariants
                    .map(title => `button[aria-label="Start ${title} declaration"], button[aria-label*="Start ${title} declaration"]`)
                    .join(', ')
            ).first();
            const buttonByCard = page.locator(
                `xpath=${sectionTitleVariants
                    .map(title => `//*[normalize-space(text())="${title}"]/ancestor::*[.//button[contains(normalize-space(.), "Start declaration")]][1]//button[contains(normalize-space(.), "Start declaration")]`)
                    .join(' | ')}`
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

        // 按文案选择 radio，适用于 Play Console Material 组件。
        async function selectRadio(textRegex) {
            await retryAction(async () => {
                const radio = page.locator('material-radio, [role="radio"]').filter({ hasText: textRegex }).first();
                await radio.scrollIntoViewIfNeeded({ timeout: 10000 });
                await radio.click({ timeout: 10000 });
                await delay(page, 1500);
            }, `Select radio containing "${textRegex}"`);
        }

        // 按文案勾选 checkbox，已勾选时不重复点击。
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

        // 点击当前页面主按钮，兼容 debug-id、button、role=button、material-button。
        async function clickMainButtonOn(targetPage, text) {
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
                    const loc = targetPage.locator(sel).first();
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
                await delay(targetPage, 3000);
            }, `Click "${text}" button`);
        }

        async function clickMainButton(text) {
            await clickMainButtonOn(page, text);
        }

        // Save 主按钮找不到时，尝试右下角 More options -> Save。
        async function clickOverflowSaveFallback(contextLabel = 'Save') {
            console.log(`[SAVE] ${contextLabel}: trying More options -> Save fallback...`);
            await retryAction(async () => {
                let moreOptionsButton = null;
                for (const selector of OVERFLOW_MORE_OPTIONS_SELECTORS) {
                    const candidate = page.locator(selector).last();
                    if (await candidate.isVisible().catch(() => false)) {
                        moreOptionsButton = candidate;
                        break;
                    }
                }
                if (!moreOptionsButton) {
                    throw new Error(`${contextLabel} More options button not found.`);
                }

                await clickLocatorRobust(moreOptionsButton, `${contextLabel} More options button`, 10000);
                await delay(page, 1000);

                let saveMenuItem = null;
                for (const selector of OVERFLOW_SAVE_MENU_SELECTORS) {
                    const candidate = page.locator(selector).last();
                    if (await candidate.isVisible().catch(() => false)) {
                        saveMenuItem = candidate;
                        break;
                    }
                }
                if (!saveMenuItem) {
                    throw new Error(`${contextLabel} Save menu item not found.`);
                }

                await clickLocatorRobust(saveMenuItem, `${contextLabel} Save menu item`, 10000);
                await delay(page, 3000);
            }, `${contextLabel} More options Save fallback`, 3);
        }

        // 优先点主 Save，只有明确找不到 Save 主按钮时才走三点菜单兜底。
        async function clickSaveWithOverflowFallback(contextLabel = 'Save') {
            try {
                await clickMainButton('Save');
            } catch (err) {
                const message = String((err && err.message) || '');
                if (!/Main button with text "Save" not found/i.test(message)) {
                    throw err;
                }
                await clickOverflowSaveFallback(contextLabel);
            }
        }

        async function clickFinancialFeaturesSave() {
            await clickSaveWithOverflowFallback('Financial features');
        }

        // 在多个候选选择器中找第一个可见输入框并填值。
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

        // 在指定页面对象上打开当前 app 的子路径，供主标签和临时标签共用。
        async function gotoAppSubPageOn(targetPage, subPath, label, waitMs = 3000) {
            console.log(`Navigating to "${label}"...`);
            await targetPage.bringToFront().catch(() => { });
            await targetPage.goto(appBasePath + subPath, { timeout: 120000, waitUntil: 'domcontentloaded' });
            await targetPage.waitForLoadState('load', { timeout: 60000 }).catch(() => { });
            await delay(targetPage, waitMs);
        }

        async function gotoAppSubPage(subPath, label, waitMs = 3000) {
            await gotoAppSubPageOn(page, subPath, label, waitMs);
        }

        // 兼容 fill 失败的输入框：fill 失败后用 DOM 赋值并派发 input/change。
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

        // 从一组 locator 中点击第一个可见元素。
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

        // 后续素材、描述、AAB 都依赖 AppGenie 详情页；丢失时重新打开详情。
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

        // 从 AppGenie 详情卡片读取指定字段，例如短描述、长描述。
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

        // 将 Playwright 捕获到的下载保存到当前任务临时目录。
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

        // 从 AppGenie 详情页点击指定下载按钮，并返回 download 对象。
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

        // 下载并解压商店素材 ZIP，按约定取 icon、feature graphic、手机截图。
        async function downloadAndExtractStoreImages() {
            const download = await downloadFromAppGenieDetail(/全部下载|Download\s*all|All\s*download/i, 'store images ZIP');
            const { filePath, root } = await saveDownloadToTemp(download, 'images', `${safeFileToken(task.appName)}-images.zip`);
            const extractDir = path.join(root, 'extract');
            extractZipToDirectory(filePath, extractDir);
            console.log(`[APPGENIE] Images ZIP extracted: ${extractDir}`);
            await randomDelay(page, 10000, 15000);
            const [appIcon] = findImagesNamed(extractDir, ['0']);
            const [featureGraphic] = findImagesNamed(extractDir, ['1024']);
            const phoneScreenshots = findImagesNamed(extractDir, ['1', '2', '3']);
            console.log(
                `[APPGENIE] Store images ready: ` +
                `icon=${path.basename(appIcon)}, ` +
                `feature=${path.basename(featureGraphic)}, ` +
                `phone=${phoneScreenshots.map(p => path.basename(p)).join(', ')}`
            );
            return { appIcon, featureGraphic, phoneScreenshots };
        }

        // 下载 AAB；如果 download 事件漏掉，则扫描浏览器下载目录兜底。
        async function downloadAabFromAppGenie() {
            const detailPage = await ensureAppGenieDetailPageReady();
            const candidates = detailPage.locator('button, a, [role="button"]').filter({
                hasText: /AAB\s*下载|下载\s*AAB|AAB\s*Download|Download\s*AAB/i
            });
            const count = await candidates.count().catch(() => 0);
            let target = null;
            for (let i = 0; i < count; i += 1) {
                const loc = candidates.nth(i);
                if (!(await loc.isVisible().catch(() => false))) continue;
                const text = await loc.innerText().catch(() => '');
                if (!/AAB/i.test(text) || !/(下载|Download)/i.test(text)) continue;
                if (/APK|全部下载|Download\s*all|图片|image|zip/i.test(text)) continue;
                target = loc;
                break;
            }

            if (!target) {
                throw new Error('AAB download button not found on AppGenie details page.');
            }

            console.log('[APPGENIE] Downloading AAB package from AppGenie details...');
            const downloadStartedAtMs = Date.now();
            const downloadPromise = detailPage.waitForEvent('download', { timeout: 90000 })
                .then(download => ({ download }))
                .catch(error => ({ error }));
            await clickLocatorRobust(target, 'AAB download button', 10000);
            const downloadResult = await downloadPromise;

            if (downloadResult.download) {
                const saved = await saveDownloadToTemp(downloadResult.download, 'aab', `${safeFileToken(task.appName)}.aab`);
                if (path.extname(saved.filePath).toLowerCase() !== '.aab') {
                    throw new Error(`Downloaded package is not an AAB file: ${saved.filePath}`);
                }
                console.log(`[APPGENIE] AAB saved: ${saved.filePath}`);
                return saved.filePath;
            }

            console.log(
                `[APPGENIE] AAB download event was not captured; scanning browser download folder. ` +
                `Reason: ${downloadResult.error.message}`
            );
            const fallbackPath = await waitForRecentAabDownloadFromDisk({
                startedAtMs: downloadStartedAtMs,
                expectedPackageName: task.packageName,
                timeoutMs: 300000
            });
            if (Array.isArray(runtimeOptions.fallbackAabCleanupPaths)) {
                runtimeOptions.fallbackAabCleanupPaths.push(fallbackPath);
            }
            return fallbackPath;
        }

        // 通用上传：优先用 filechooser，拿不到时寻找隐藏 input[type=file]。
        async function uploadFilesByUploadButton(targetPage, uploadButton, files, label) {
            await retryAction(async () => {
                const chooserPromise = targetPage.waitForEvent('filechooser', { timeout: 30000 }).catch(() => null);
                await clickLocatorRobust(uploadButton, `${label} Upload button`, 10000);
                const chooser = await chooserPromise;
                if (chooser) {
                    await chooser.setFiles(files);
                    return;
                }

                const inputCandidates = [
                    uploadButton.locator('xpath=ancestor::*[contains(@class, "upload") or contains(@class, "file")][1]//input[@type="file"]').first(),
                    targetPage.locator('button.upload-button input[type="file"], .upload-button input[type="file"]').first(),
                    targetPage.locator('input[type="file"][accept*=".aab"], input[type="file"][accept*="application"]').last(),
                    targetPage.locator('input[type="file"]').last()
                ];
                let input = null;
                for (const candidate of inputCandidates) {
                    if (await candidate.count().catch(() => 0)) {
                        input = candidate;
                        break;
                    }
                }
                if (!input) {
                    throw new Error(`${label} file input not found after Upload click.`);
                }
                await input.setInputFiles(files, { timeout: 30000 });
            }, `Upload ${label}`, 3);
            await delay(targetPage, 3000);
        }

        // 从选择器数组中点击第一个可见元素。
        async function clickFirstVisibleSelectorOn(targetPage, selectors, label, waitMs = 3000) {
            await retryAction(async () => {
                let target = null;
                for (const selector of selectors) {
                    const candidate = targetPage.locator(selector).first();
                    if (await candidate.isVisible().catch(() => false)) {
                        target = candidate;
                        break;
                    }
                }
                if (!target) {
                    throw new Error(`${label} not found.`);
                }
                await clickLocatorRobust(target, label, 15000);
                await delay(targetPage, waitMs);
            }, `Click ${label}`, 3);
        }

        // 只查找不点击，供更复杂的条件判断复用。
        async function firstVisibleLocatorFromSelectors(targetPage, selectors) {
            for (const selector of selectors) {
                const candidate = targetPage.locator(selector).first();
                if (await candidate.isVisible().catch(() => false)) {
                    return candidate;
                }
            }
            return null;
        }

        // 在 Protected with Play 页面定位 Play Store protection 卡片。
        async function findVisiblePlayStoreProtectionCard() {
            for (const selector of APP_SIGNING_PLAY_STORE_PROTECTION_CARD_SELECTORS) {
                const candidates = page.locator(selector);
                const count = await candidates.count().catch(() => 0);
                for (let i = 0; i < count; i += 1) {
                    const candidate = candidates.nth(i);
                    if (!(await candidate.isVisible().catch(() => false))) {
                        continue;
                    }
                    const text = await candidate.innerText({ timeout: 1000 }).catch(() => '');
                    if (/Play Store protection/i.test(text)) {
                        return candidate;
                    }
                }
            }
            return null;
        }

        // 只在 Protect app signing key 行里找 Manage Play app signing，避免误点其他 Manage。
        async function findManagePlayAppSigningButton() {
            for (const selector of APP_SIGNING_MANAGE_BUTTON_SCOPED_SELECTORS) {
                const candidate = page.locator(selector).first();
                if (await candidate.isVisible().catch(() => false)) {
                    return candidate;
                }
            }

            const card = await findVisiblePlayStoreProtectionCard();
            if (!card) {
                return null;
            }

            const rows = card.locator(
                'protection-feature-list[debug-id="expanded-feature-list"] .feature-row, ' +
                'protection-feature-list[debug-id="expanded-feature-list"] .pc-card, ' +
                '.feature-row, .pc-card'
            );
            const count = await rows.count().catch(() => 0);
            for (let i = 0; i < count; i += 1) {
                const row = rows.nth(i);
                if (!(await row.isVisible().catch(() => false))) {
                    continue;
                }
                const text = await row.innerText({ timeout: 1000 }).catch(() => '');
                if (!/Protect app signing key/i.test(text)) {
                    continue;
                }
                const button = row.locator(
                    'button[aria-label="Manage Play app signing"], ' +
                    'button[debug-id="cta-button"]:has-text("Manage Play app signing"), ' +
                    'button:has-text("Manage Play app signing")'
                ).first();
                if (await button.isVisible().catch(() => false)) {
                    return button;
                }
            }

            return null;
        }

        // 判断 Manage 是否真的进入了 App signing 页面。
        async function isAppSigningPageVisible() {
            if (/app-signing/i.test(page.url())) {
                return true;
            }
            const markers = [
                page.getByText(/^App signing$/i).first(),
                page.getByText(/App signing key certificate/i).first(),
                page.getByText(/SHA-1 certificate fingerprint/i).first()
            ];
            for (const marker of markers) {
                if (await marker.isVisible().catch(() => false)) {
                    return true;
                }
            }
            return false;
        }

        // Play Store protection 默认可能折叠，需要展开后才能看到 app signing 管理按钮。
        async function ensurePlayStoreProtectionDetailsExpanded() {
            const manageButton = await findManagePlayAppSigningButton();
            if (manageButton) return;

            console.log('[REVIEW_UPLOAD] Expanding Play Store protection details...');
            await retryAction(async () => {
                const visibleManage = await findManagePlayAppSigningButton();
                if (visibleManage) return;

                const card = await findVisiblePlayStoreProtectionCard();
                if (!card) {
                    throw new Error('Play Store protection card not found.');
                }

                let expandButton = card.locator(
                    'button[debug-id="expansion-button"], button[aria-label="Show details"], button[aria-label="Hide details"]'
                ).first();
                if (!(await expandButton.isVisible().catch(() => false))) {
                    expandButton = await firstVisibleLocatorFromSelectors(page, APP_SIGNING_PLAY_STORE_PROTECTION_EXPAND_SELECTORS);
                }

                if (!expandButton || !(await expandButton.isVisible().catch(() => false))) {
                    throw new Error('Play Store protection expand button not found.');
                }

                const expanded = await expandButton.evaluate(el => el.getAttribute('aria-expanded') === 'true').catch(() => false);
                if (!expanded) {
                    await clickLocatorRobust(expandButton, 'Play Store protection expand button', 15000);
                    await delay(page, 2000);
                }

                const expandedManageButton = await findManagePlayAppSigningButton();
                if (!expandedManageButton) {
                    throw new Error('Manage Play app signing button not visible in Protect app signing key row after expanding Play Store protection.');
                }
            }, 'Expand Play Store protection details', 3);
        }

        // 输入框快速填值并校验；fill 不生效时依次尝试键盘和逐字输入。
        async function fillInputFastWithFallback(targetPage, inputLocator, value, label) {
            await inputLocator.waitFor({ state: 'visible', timeout: 30000 });
            await inputLocator.scrollIntoViewIfNeeded({ timeout: 5000 }).catch(() => { });

            const verifyValue = async () => {
                const current = await inputLocator.inputValue({ timeout: 5000 }).catch(() => '');
                return String(current || '').trim() === String(value || '').trim();
            };

            await inputLocator.fill(value, { timeout: 10000 }).catch(() => { });
            await inputLocator.evaluate((el) => {
                el.dispatchEvent(new Event('input', { bubbles: true }));
                el.dispatchEvent(new Event('change', { bubbles: true }));
            }).catch(() => { });
            await delay(targetPage, 500);
            if (await verifyValue()) return;

            await inputLocator.click({ timeout: 10000 }).catch(() => { });
            await targetPage.keyboard.press('Control+A').catch(() => { });
            await targetPage.keyboard.insertText(value).catch(() => { });
            await delay(targetPage, 500);
            if (await verifyValue()) return;

            await inputLocator.fill('', { timeout: 10000 }).catch(() => { });
            await inputLocator.type(value, { delay: 5 });
            await delay(targetPage, 500);
            if (!(await verifyValue())) {
                throw new Error(`${label} was not filled correctly.`);
            }
        }

        // 从 Play App signing 页面读取 SHA-1，优先解析页面文本，失败时尝试复制按钮/剪贴板。
        async function readAppSigningSha1Fingerprint() {
            console.log('[REVIEW_UPLOAD] Opening Protected with Play...');
            await gotoAppSubPage(PLAY_PROTECTED_WITH_PLAY_PATH, 'Protected with Play', 5000);

            await ensurePlayStoreProtectionDetailsExpanded();

            await retryAction(async () => {
                await ensurePlayStoreProtectionDetailsExpanded();

                const manageButton = await findManagePlayAppSigningButton();
                if (!manageButton) {
                    throw new Error('Manage Play app signing button not found in Protect app signing key row.');
                }

                await clickLocatorRobust(manageButton, 'Manage Play app signing button', 15000);
                await page.waitForLoadState('domcontentloaded', { timeout: 60000 }).catch(() => { });
                await page.waitForLoadState('load', { timeout: 60000 }).catch(() => { });
                await delay(page, 5000);

                if (!(await isAppSigningPageVisible())) {
                    const wrongUrl = page.url();
                    await gotoAppSubPage(PLAY_PROTECTED_WITH_PLAY_PATH, 'Protected with Play', 3000).catch(() => { });
                    throw new Error(`Manage Play app signing did not open App signing page (url: ${wrongUrl}).`);
                }
            }, 'Open Play app signing', 3);

            console.log('[REVIEW_UPLOAD] Reading SHA-1 certificate fingerprint...');
            let sha1 = '';
            await retryAction(async () => {
                const candidateText = await page.evaluate(() => {
                    const normalize = value => String(value || '').replace(/\s+/g, ' ').trim();
                    const labels = Array.from(document.querySelectorAll('body *'))
                        .filter(el => /SHA-1\s+certificate\s+fingerprint/i.test(normalize(el.textContent)));
                    const chunks = [];
                    for (const labelEl of labels) {
                        let current = labelEl;
                        for (let depth = 0; current && depth < 6; depth += 1, current = current.parentElement) {
                            chunks.push(normalize(current.innerText || current.textContent));
                        }
                    }
                    chunks.push(normalize(document.body?.innerText || ''));
                    return chunks.join('\n');
                }).catch(() => '');
                sha1 = extractSha1Fingerprint(candidateText);

                if (!sha1) {
                    await page.context().grantPermissions(['clipboard-read', 'clipboard-write'], {
                        origin: 'https://play.google.com'
                    }).catch(() => { });
                    const sha1Row = page.locator('div, console-form-row, section').filter({
                        hasText: /SHA-1\s+certificate\s+fingerprint/i
                    }).last();
                    const copyButton = sha1Row.locator(
                        'button, [role="button"], material-icon[role="button"], material-icon[aria-label*="Copy"], material-icon[aria-label*="copy"]'
                    ).last();
                    if (await copyButton.isVisible().catch(() => false)) {
                        await clickLocatorRobust(copyButton, 'SHA-1 copy button', 10000);
                        await delay(page, 1000);
                        const clipboardText = await page.evaluate(async () => {
                            try {
                                return await navigator.clipboard.readText();
                            } catch {
                                return '';
                            }
                        }).catch(() => '');
                        sha1 = extractSha1Fingerprint(clipboardText);
                    }
                }

                if (!sha1) {
                    throw new Error('SHA-1 fingerprint text not found.');
                }
            }, 'Read SHA-1 fingerprint from page', 3);

            console.log(`[REVIEW_UPLOAD] SHA-1 fingerprint captured: ${sha1}`);
            return sha1;
        }

        // 打开 Latest releases and bundles 并保存整页截图，作为 AppGenie 提审材料。
        async function captureLatestReleasesScreenshot() {
            console.log('[REVIEW_UPLOAD] Opening Latest releases and bundles for screenshot...');
            await gotoAppSubPage(PLAY_RELEASES_OVERVIEW_PATH, 'Latest releases and bundles', 5000);
            const heading = page.locator('text=/Latest releases and bundles/i').first();
            await heading.waitFor({ state: 'visible', timeout: 90000 });
            await page.waitForLoadState('networkidle', { timeout: 60000 }).catch(() => { });
            await delay(page, 3000);

            const screenshotRoot = path.join(
                os.tmpdir(),
                'app-create-downloads',
                safeFileToken(`${task.appName}-${task.packageName}`),
                'review-upload'
            );
            fs.mkdirSync(screenshotRoot, { recursive: true });
            const screenshotPath = buildReviewScreenshotPath(screenshotRoot, task.appName);
            await page.screenshot({ path: screenshotPath, fullPage: true });
            console.log(`[REVIEW_UPLOAD] Latest releases screenshot saved: ${screenshotPath}`);
            return screenshotPath;
        }

        // 打开 AppGenie 待提审任务页，并展开包含当前包名的任务列表。
        async function ensureAppGenieReviewTaskListPage() {
            const { page: appGeniePage, source } = await acquireOrCreateAuxPage(
                context,
                /appgenie-ai\.com/i,
                'https://appgenie-ai.com/login',
                'appgenie'
            );
            console.log(`[REVIEW_UPLOAD] AppGenie tab: ${source} | ${appGeniePage.url() || 'about:blank'}`);
            await appGeniePage.bringToFront().catch(() => { });
            await delay(appGeniePage, 1000);
            await ensureAppGenieLoggedIn(appGeniePage, runtimeOptions);
            if (!/\/my-submit-tasks(?:[/?#]|$)/i.test(String(appGeniePage.url() || ''))) {
                await appGeniePage.goto('https://appgenie-ai.com/my-submit-tasks', {
                    timeout: 120000,
                    waitUntil: 'domcontentloaded'
                });
                await appGeniePage.waitForLoadState('load', { timeout: 60000 }).catch(() => { });
                await delay(appGeniePage, 3000);
            }
            await ensureAppGenieOnMyTasks(appGeniePage);
            await randomDelay(appGeniePage, 3000, 5000);
            await openAppGeniePendingDrawerIfNeeded(appGeniePage);
            return appGeniePage;
        }

        // 如果当前页面还没显示包名，就按邮箱找到任务行并打开抽屉。
        async function openAppGeniePendingDrawerIfNeeded(appGeniePage) {
            const packageVisible = await appGeniePage.locator(`text=${task.packageName}`).first().isVisible().catch(() => false);
            if (packageVisible) {
                return;
            }

            let emailRow = appGeniePage.locator('tr.ant-table-row').filter({ hasText: runtimeOptions.contactEmail }).first();
            if (!(await emailRow.isVisible().catch(() => false))) {
                const searchInput = appGeniePage.locator(
                    'input[placeholder*="搜索邮箱"], input[placeholder*="邮箱"], input[placeholder*="email" i], input.ant-input'
                ).first();
                await retryAction(async () => {
                    await searchInput.waitFor({ state: 'visible', timeout: 30000 });
                    await searchInput.click({ timeout: 10000 });
                    await searchInput.fill('');
                    await searchInput.type(runtimeOptions.contactEmail, { delay: 40 });
                    await appGeniePage.keyboard.press('Enter').catch(() => { });
                }, 'Search AppGenie tasks for review upload', 3);
                await randomDelay(appGeniePage, 5000, 8000);
                emailRow = appGeniePage.locator('tr.ant-table-row').filter({ hasText: runtimeOptions.contactEmail }).first();
            }

            await retryAction(async () => {
                await emailRow.waitFor({ state: 'visible', timeout: 60000 });
                let openBtn = emailRow.locator('button').filter({ hasText: /查看应用|View\s*app/i }).first();
                if (!(await openBtn.isVisible().catch(() => false))) {
                    openBtn = emailRow.locator('button').last();
                }
                if (await openBtn.isVisible().catch(() => false)) {
                    await clickLocatorRobust(openBtn, 'AppGenie View app button for review upload', 10000);
                } else {
                    await clickLocatorRobust(emailRow, 'AppGenie task row for review upload', 10000);
                }
            }, 'Open AppGenie pending drawer for review upload', 3);
            await randomDelay(appGeniePage, 3000, 6000);
        }

        // 在 AppGenie 待提审列表中按包名定位具体应用卡片。
        async function findAppGenieReviewCardByPackage(appGeniePage) {
            const packageText = task.packageName;
            await retryAction(async () => {
                const bodyText = await appGeniePage.locator('body').innerText({ timeout: 10000 }).catch(() => '');
                if (!bodyText.includes(packageText)) {
                    throw new Error(`Package "${packageText}" not visible in AppGenie task list.`);
                }
            }, 'Wait AppGenie package visible', 4);

            const packageTextNode = appGeniePage.locator(`text=${packageText}`).first();
            const card = packageTextNode.locator(
                'xpath=ancestor::*[.//button[contains(normalize-space(.), "提交审核")]][1]'
            );
            await card.waitFor({ state: 'visible', timeout: 60000 });
            return card;
        }

        // 打开 AppGenie 的“提交审核”弹窗。
        async function openAppGenieReviewSubmitDialog(appGeniePage) {
            console.log('[REVIEW_UPLOAD] Opening AppGenie submit review dialog...');
            const appCard = await findAppGenieReviewCardByPackage(appGeniePage);
            const submitButton = appCard.locator('button').filter({ hasText: /提交审核/ }).last();
            await retryAction(async () => {
                await submitButton.waitFor({ state: 'visible', timeout: 30000 });
                await clickLocatorRobust(submitButton, 'AppGenie submit review button', 15000);
            }, 'Click AppGenie submit review button', 3);

            const modal = appGeniePage.locator('.ant-modal').filter({
                hasText: /提交审核|SHA1|上传/
            }).last();
            await modal.waitFor({ state: 'visible', timeout: 60000 });
            await delay(appGeniePage, 1500);
            return modal;
        }

        // 上传审核截图、填写 SHA-1 并提交 AppGenie 审核材料。
        async function submitAppGenieReviewUpload(screenshotPath, sha1Fingerprint) {
            const appGeniePage = await ensureAppGenieReviewTaskListPage();
            const modal = await openAppGenieReviewSubmitDialog(appGeniePage);

            console.log('[REVIEW_UPLOAD] Uploading review screenshot in AppGenie...');
            await retryAction(async () => {
                const fileInput = appGeniePage.locator(APPGENIE_REVIEW_FILE_INPUT_SELECTOR).last();
                await fileInput.waitFor({ state: 'attached', timeout: 30000 });
                await fileInput.setInputFiles(screenshotPath, { timeout: 30000 });
            }, 'Upload AppGenie review screenshot', 3);
            // AppGenie 截图上传后给页面更多处理时间，降低未处理完就提交的概率。
            await delay(appGeniePage, 5000);

            console.log('[REVIEW_UPLOAD] Filling SHA-1 in AppGenie...');
            await retryAction(async () => {
                let sha1Input = appGeniePage.locator(APPGENIE_REVIEW_SHA1_INPUT_SELECTOR).last();
                if (!(await sha1Input.count().catch(() => 0))) {
                    sha1Input = modal.locator('input.ant-input, input[type="text"]').first();
                }
                await fillInputFastWithFallback(appGeniePage, sha1Input, sha1Fingerprint, 'AppGenie SHA1 input');
            }, 'Fill AppGenie SHA1 input', 3);
            await delay(appGeniePage, 1500);

            console.log('[REVIEW_UPLOAD] Submitting AppGenie review upload...');
            await retryAction(async () => {
                let submitButton = null;
                for (const selector of APPGENIE_REVIEW_SUBMIT_BUTTON_SELECTORS) {
                    const candidate = appGeniePage.locator(selector).last();
                    if (await candidate.isVisible().catch(() => false)) {
                        submitButton = candidate;
                        break;
                    }
                }
                if (!submitButton) {
                    throw new Error('AppGenie modal submit button not found.');
                }
                if (await isLocatorDisabled(submitButton)) {
                    throw new Error('AppGenie modal submit button is disabled.');
                }
                await clickLocatorRobust(submitButton, 'AppGenie modal submit button', 15000);
            }, 'Submit AppGenie review upload', 3);

            await retryAction(async () => {
                const stillVisible = await modal.isVisible().catch(() => false);
                const successToast = await appGeniePage.locator('.ant-message-notice, .ant-notification-notice, text=/成功|提交成功/i')
                    .first()
                    .isVisible()
                    .catch(() => false);
                if (stillVisible && !successToast) {
                    throw new Error('AppGenie review submit confirmation not detected yet.');
                }
            }, 'Wait AppGenie review submit confirmation', 4);
            await delay(appGeniePage, 3000);
        }

        // 发布后的后置步骤：SHA-1 + Latest releases 截图 + AppGenie 回传。
        async function runReviewScreenshotUploadStep() {
            console.log('Executing post-release step: Review screenshot upload...');
            const sha1Fingerprint = await readAppSigningSha1Fingerprint();
            const screenshotPath = await captureLatestReleasesScreenshot();
            await submitAppGenieReviewUpload(screenshotPath, sha1Fingerprint);
            markStepDone(PROGRESS_STEP_REVIEW_UPLOAD_DONE);
        }

        // 上传面板中素材默认已选中时，直接点 Add 加入商店列表。
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

        // Add 按钮置灰时，按文件名手动选中刚上传的素材。
        async function selectStoreAssetsByFileNames(files, label) {
            for (const file of files) {
                const fileName = path.basename(file);
                const fileNameText = page.locator(`text="${fileName}"`).last();
                if (!(await fileNameText.isVisible().catch(() => false))) {
                    console.log(`[STORE LISTING] ${label} asset ${fileName} not visible for manual selection.`);
                    continue;
                }

                console.log(`[STORE LISTING] Selecting ${label} asset: ${fileName}`);
                await clickLocatorRobust(fileNameText, `${label} asset ${fileName}`, 10000);
                await randomDelay(page, 2000, 3000);
            }
        }

        // 加入素材的加强版：必要时先按文件名选择，再点 Add。
        async function addSelectedStoreAssetsToListingWithFallback(label, files) {
            console.log(`[STORE LISTING] Adding selected ${label} assets to listing...`);
            await retryAction(async () => {
                const addButton = page.locator(
                    'button[debug-id="add-to-content-button"], button[aria-label="Add"], [role="button"][debug-id="add-to-content-button"], [role="button"][aria-label="Add"]'
                ).last();
                await addButton.waitFor({ state: 'visible', timeout: 30000 });
                if (await isLocatorDisabled(addButton)) {
                    console.log(`[STORE LISTING] ${label} Add button disabled; selecting uploaded file(s) first...`);
                    await selectStoreAssetsByFileNames(files, label);
                }
                if (await isLocatorDisabled(addButton)) {
                    throw new Error(`${label} Add button is disabled.`);
                }
                await clickLocatorRobust(addButton, `${label} Add button`, 15000);
            }, `Click ${label} Add button`, 3);
            await randomDelay(page, 5000, 7000);
            console.log(`[STORE LISTING] Selected ${label} assets added.`);
        }

        // 打开商店列表某个图片区块的 Add assets 面板。
        async function openStoreListingAssetPanel(sectionLabel, uploaderDebugId, label) {
            const sectionByAria = page.locator(`div[role="group"][aria-label="${sectionLabel}"]`).first();
            const sectionByText = page.locator(
                `xpath=(//*[contains(normalize-space(.), "${sectionLabel}")]/ancestor::*[@role="group"][1])[1]`
            ).first();
            const uploader = uploaderDebugId
                ? page.locator(`localized-image-uploader[debug-id="${uploaderDebugId}"]`).first()
                : page.locator(`div[role="group"][aria-label="${sectionLabel}"] localized-image-uploader`).first();
            const scrollTarget = (await sectionByAria.isVisible().catch(() => false))
                ? sectionByAria
                : ((await sectionByText.isVisible().catch(() => false)) ? sectionByText : uploader);

            await scrollTarget.scrollIntoViewIfNeeded({ timeout: 15000 }).catch(async () => {
                await page.mouse.wheel(0, 900).catch(() => { });
            });
            await randomDelay(page, 4000, 6000);

            const addByUploader = uploaderDebugId
                ? page.locator(`localized-image-uploader[debug-id="${uploaderDebugId}"] button[debug-id="add-button"]`).first()
                : page.locator(`div[role="group"][aria-label="${sectionLabel}"] button[debug-id="add-button"]`).first();
            const addByAria = page.locator(`div[role="group"][aria-label="${sectionLabel}"] button:has-text("Add assets")`).first();
            const addByXpath = page.locator(
                `xpath=(//*[@role="group" and @aria-label="${sectionLabel}"]//button[contains(normalize-space(.), "Add assets")])[1]`
            ).first();
            const addByText = page.locator(
                `xpath=(//*[contains(normalize-space(.), "${sectionLabel}")]/following::button[contains(normalize-space(.), "Add assets")])[1]`
            ).first();

            console.log(`[STORE LISTING] Opening ${label} asset panel...`);
            await clickFirstVisibleOn(page, [addByUploader, addByAria, addByXpath, addByText], `${label} Add assets`, 5000);
            await randomDelay(page, 5000, 7000);
        }

        // 上传图标、Feature graphic 或手机截图，并加入对应素材区。
        async function uploadStoreListingGraphicAsset(sectionLabel, uploaderDebugId, files, label) {
            const fileList = Array.isArray(files) ? files : [files];
            if (!fileList.length || fileList.some(file => !file || !fs.existsSync(file))) {
                throw new Error(`${label} upload file is missing.`);
            }

            await openStoreListingAssetPanel(sectionLabel, uploaderDebugId, label);

            const uploadBtn = page.locator(
                'button:has-text("Upload"), [role="button"]:has-text("Upload"), material-button:has-text("Upload")'
            ).last();
            console.log(`[STORE LISTING] Uploading ${label}: ${fileList.map(file => path.basename(file)).join(', ')}`);
            await uploadFilesByUploadButton(page, uploadBtn, fileList, label);
            console.log(`[STORE LISTING] ${label} uploaded, waiting for processing...`);
            await randomDelay(page, 10000, 15000);
            await addSelectedStoreAssetsToListingWithFallback(label, fileList);
            await closePanelIfVisible();
            await randomDelay(page, 5000, 7000);
        }

        // Store settings/listing 里按区块标题点击 Edit，找不到时用全局 Edit 按钮兜底。
        async function clickStoreSectionEdit(sectionTitleRegex, label, fallbackMode = 'first', targetPage = page) {
            const sectionEdit = targetPage.locator('console-section, console-block-1-column, console-form-row, section').filter({
                hasText: sectionTitleRegex
            }).locator('material-button[debug-id="edit-store-listing-section-button"], button:has-text("Edit")').first();
            const allEditButtons = targetPage.locator(
                'material-button[debug-id="edit-store-listing-section-button"], button:has-text("Edit"), [role="button"]:has-text("Edit")'
            );
            const fallback = fallbackMode === 'last' ? allEditButtons.last() : allEditButtons.first();
            await clickFirstVisibleOn(targetPage, [sectionEdit, fallback], `${label} Edit`, 3000);
        }

        // Store listing contact details 弹窗定位器。
        function storeListingContactDialogLocator() {
            return page.locator(
                'div[role="dialog"], material-dialog, .mdc-dialog, .mat-mdc-dialog-container, .cdk-overlay-pane'
            ).filter({ hasText: /Store\s*listing\s*contact\s*details/i }).first();
        }

        // 确保联系信息弹窗已经打开。
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

        // 填写商店联系邮箱，并校验输入框实际值。
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

        // 点击联系信息弹窗内的 Save，避免误点页面外层按钮。
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

        // 保存后如果弹窗还在，尝试关闭，避免遮挡后续 Store listing 操作。
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

            await dialog.waitFor({ state: 'hidden', timeout: 5000 }).catch(async () => {
                await page.keyboard.press('Escape').catch(() => { });
                await delay(page, 2000);
            });
            await randomDelay(page, 2000, 3000);
            console.log('[STORE SETTINGS] Store listing contact details dialog closed.');
        }

        // 关闭右侧抽屉/编辑面板，给下一步留下干净页面状态。
        async function closePanelIfVisible(targetPage = page) {
            const closeBtn = targetPage.locator(
                'button[aria-label="Close"], material-button[aria-label="Close"], .close-icon-button, button:has-text("Close")'
            ).first();
            if (await closeBtn.isVisible().catch(() => false)) {
                await clickLocatorRobust(closeBtn, 'Close panel button', 10000);
                await randomDelay(targetPage, 2000, 3000);
            }
        }

        // 按 debug-id 打开下拉并选择匹配项，主要用于 App/Game 和分类。
        async function selectDropdownByDebugId(debugId, optionRegex, label, targetPage = page) {
            const dropdown = targetPage.locator(
                `material-dropdown-select[debug-id="${debugId}"] dropdown-button, ` +
                `material-dropdown-select[debug-id="${debugId}"] [role="button"]`
            ).first();
            await clickLocatorRobust(dropdown, `${label} dropdown`, 10000);
            await delay(targetPage, 1200);

            const options = targetPage.locator(
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
            await delay(targetPage, 3000);
        }

        // 根据 AppGenie 类型设置 Play Console 的 App/Game 和分类。
        async function setStoreCategoryFromAppGenieType(preOpenedPage = null) {
            const typeKey = task.appGenieTypeKey || normalizeAppGenieTypeKey(task.appGenieTypeLabel);
            const isGame = typeKey === 'game';
            const appOrGame = isGame ? 'Game' : 'App';
            const category = isGame ? 'Casual' : 'Tools';
            console.log(`[STORE SETTINGS] Setting App category in temporary tab: ${task.appGenieTypeLabel || 'unknown'} -> ${appOrGame}/${category}`);

            const categoryPage = preOpenedPage || await context.newPage();
            try {
                if (preOpenedPage) {
                    console.log('[STORE SETTINGS] Using pre-opened Production release tab, then navigating it to Store settings...');
                }
                await gotoAppSubPageOn(categoryPage, '/store-settings', 'Store settings for app category', 5000);
                await clickStoreSectionEdit(/App\s*category|App or game|Category/i, 'App category', 'last', categoryPage);
                await randomDelay(categoryPage, 3000, 5000);

                await selectDropdownByDebugId(
                    'type-dropdown',
                    new RegExp(`^\\s*${escapeRegExp(appOrGame)}\\s*$`, 'i'),
                    'App or game',
                    categoryPage
                );
                await randomDelay(categoryPage, 3000, 5000);

                await selectDropdownByDebugId(
                    'category-dropdown',
                    new RegExp(`^\\s*${escapeRegExp(category)}\\s*$`, 'i'),
                    'Category',
                    categoryPage
                );
                await randomDelay(categoryPage, 3000, 5000);

                await clickMainButtonOn(categoryPage, 'Save');
                await waitSaved(categoryPage);
                await closePanelIfVisible(categoryPage);
                await randomDelay(categoryPage, 3000, 5000);
                console.log('[STORE SETTINGS] App category updated in temporary tab.');
            } finally {
                console.log('[STORE SETTINGS] Closing temporary Store settings tab...');
                if (!categoryPage.isClosed()) {
                    await categoryPage.close().catch(() => { });
                }
                await page.bringToFront().catch(() => { });
                await randomDelay(page, 3000, 5000);
            }
        }

        // Store settings 大步骤：目前主要维护商店联系邮箱。
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

        // Store listing 大步骤：填写描述，上传图标、Feature graphic 和截图。
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

            const storeImages = await downloadAndExtractStoreImages();
            await page.bringToFront().catch(() => { });
            await delay(page, 5000);
            console.log('[STORE LISTING] Moving to graphics section...');
            await page.locator('text=/Graphics/i').first().scrollIntoViewIfNeeded({ timeout: 10000 }).catch(async () => {
                await page.mouse.wheel(0, 1200).catch(() => { });
            });
            await delay(page, 5000);

            await uploadStoreListingGraphicAsset('App icon', 'icon-uploader', storeImages.appIcon, 'App icon');
            await uploadStoreListingGraphicAsset('Feature graphic', 'feature-graphic-uploader', storeImages.featureGraphic, 'Feature graphic');
            await uploadStoreListingGraphicAsset('Phone screenshots', 'phone-screenshots-uploader', storeImages.phoneScreenshots, 'Phone screenshots');

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

        // 等 AAB 上传列表出现目标文件并显示可移除按钮，代表上传处理基本完成。
        async function waitForAabUploadReady(aabPath, timeoutMs = 300000) {
            const aabFileName = path.basename(aabPath);
            console.log(`[RELEASE] Waiting for uploaded AAB file item: ${aabFileName}`);
            const deadline = Date.now() + timeoutMs;
            while (Date.now() < deadline) {
                const ready = await page.evaluate((fileName) => {
                    const items = Array.from(document.querySelectorAll('.file-list-item'));
                    return items.some((item) => {
                        const name = item.querySelector('.multiple-file-name')?.textContent?.trim();
                        const hasRemoveButton = Array.from(item.querySelectorAll('button[aria-label]')).some((button) => {
                            return (button.getAttribute('aria-label') || '').includes(fileName);
                        });
                        return name === fileName && hasRemoveButton;
                    });
                }, aabFileName).catch(() => false);
                if (ready) {
                    console.log(`[RELEASE] Uploaded AAB file item detected: ${aabFileName}`);
                    return;
                }
                await delay(page, 5000);
            }
            const error = new Error(`AAB upload file item did not appear within 5 minutes: ${aabFileName}. Pausing script.`);
            error.stopRun = true;
            throw error;
        }

        // 打开 Production release 创建/编辑页，返回 AAB 上传按钮。
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

        // AAB 上传期间另开同 release URL 标签，后续用于并行设置 Store settings 分类。
        async function openSameProductionReleaseTab(releaseUrl) {
            console.log('[RELEASE] Opening same Production release page in a temporary tab before AAB upload...');
            const releaseMirrorPage = await context.newPage();
            try {
                await releaseMirrorPage.goto(releaseUrl, { timeout: 120000, waitUntil: 'domcontentloaded' });
                await releaseMirrorPage.waitForLoadState('load', { timeout: 60000 }).catch(() => { });
                await randomDelay(releaseMirrorPage, 5000, 7000);
                console.log('[RELEASE] Temporary Production release tab is ready.');
                console.log('[RELEASE] Switching focus back to original Production release page for AAB upload...');
                await page.bringToFront().catch(() => { });
                await randomDelay(page, 3000, 5000);
                return releaseMirrorPage;
            } catch (error) {
                await releaseMirrorPage.close().catch(() => { });
                throw error;
            }
        }

        // 在 release review 页识别阻塞错误；AD_ID 错误会交给自动修复流程。
        async function detectReleaseNeedFixReason() {
            const problemStatus = page.locator('status-text').filter({
                hasText: /We found some problems with your release/i
            }).first();
            const problemText = page.locator('text=/We found some problems with your release/i').first();
            const errorsHeader = page.locator(
                'simple-html[debug-id="header-text"], [debug-id="header-text"], [role="heading"]'
            ).filter({
                hasText: /Errors,\s*warnings\s*and\s*messages/i
            }).first();

            const hasProblemStatus = await problemStatus.isVisible().catch(() => false);
            const hasProblemText = await problemText.isVisible().catch(() => false);
            const hasErrorsHeader = await errorsHeader.isVisible().catch(() => false);
            if (!hasProblemStatus && !hasProblemText && !hasErrorsHeader) {
                return '';
            }

            const errorDetails = await page.locator('text=/Your advertising ID declaration/i')
                .first()
                .innerText()
                .catch(() => '');
            const pageText = await page.locator('body').innerText().catch(() => '');
            const combinedText = `${errorDetails}\n${pageText}`;
            if (isAdIdReleasePermissionError(combinedText)) {
                return combinedText;
            }
            if (!hasBlockingReleaseErrorText(combinedText)) {
                return '';
            }
            if (errorDetails) {
                return errorDetails;
            }
            return 'We found some problems with your release / Errors, warnings and messages';
        }

        // 把需要人工处理的 release 问题标记成 NEED_FIX 异常。
        function createReleaseNeedFixError(reason) {
            const error = new Error(reason || 'Release needs manual fixes before it can be saved.');
            error.code = 'RELEASE_NEED_FIX';
            error.needFix = true;
            return error;
        }

        // AD_ID 声明页里的“Turn off release errors”确认框。
        async function ensureAdIdReleaseErrorsAcknowledged() {
            await retryAction(async () => {
                const ackCheckbox = page.locator(AD_ID_ACK_CHECKBOX_SELECTOR).first();
                await ackCheckbox.waitFor({ state: 'visible', timeout: 60000 });

                const isSelected = async () => {
                    return await ackCheckbox.evaluate((root) => {
                        const input = root.querySelector('input[type="checkbox"]');
                        return Boolean(
                            (input && input.checked) ||
                            (input && input.getAttribute('aria-checked') === 'true') ||
                            root.querySelector('.mdc-checkbox--selected, input[type="checkbox"]:checked, [aria-checked="true"]')
                        );
                    }).catch(() => false);
                };

                if (!(await isSelected())) {
                    const clickTarget = ackCheckbox.locator('.mdc-checkbox, .checkbox-content').first();
                    await clickLocatorRobust(clickTarget, 'AD_ID turn off release errors checkbox', 15000);
                    await delay(page, 1200);
                }

                if (!(await isSelected())) {
                    throw new Error('AD_ID turn off release errors checkbox is not selected.');
                }
            }, 'Check AD_ID turn off release errors acknowledgement', 4);
        }

        // AD_ID 修复后回到同一个 Production draft，不重新上传 AAB。
        async function reopenProductionDraftReleaseForReview() {
            console.log('[RELEASE] Returning to Production draft release after AD_ID fix...');
            await gotoAppSubPage('/tracks/production', 'Production', 5000);
            await randomDelay(page, 3000, 5000);

            console.log('[RELEASE] Opening Production Releases tab...');
            await clickFirstVisibleSelectorOn(page, PRODUCTION_RELEASES_TAB_SELECTORS, 'Production Releases tab', 5000);
            await randomDelay(page, 3000, 5000);

            await clickFirstVisibleSelectorOn(page, PRODUCTION_EDIT_RELEASE_SELECTORS, 'Edit release button', 5000);
            await randomDelay(page, 5000, 7000);

            console.log('[RELEASE] Reopening review page without re-uploading AAB...');
            await clickMainButton('Next');
            await randomDelay(page, 5000, 7000);
        }

        // 自动处理 AAB 无 AD_ID 权限但声明使用 Advertising ID 的 release error。
        async function resolveAdIdReleasePermissionError() {
            console.log('[RELEASE] AD_ID permission release error detected; updating declaration...');
            await clickFirstVisibleSelectorOn(page, RELEASE_UPDATE_DECLARATION_SELECTORS, 'Update declaration button', 5000);
            await page.waitForLoadState('load', { timeout: 60000 }).catch(() => { });
            await randomDelay(page, 5000, 7000);

            const advertisingIdTitle = page.locator('h1, [role="heading"], simple-html[debug-id="header-text"], [debug-id="header-text"]').filter({
                hasText: /Advertising ID/i
            }).first();
            await advertisingIdTitle.waitFor({ state: 'visible', timeout: 90000 }).catch(() => { });

            console.log('[RELEASE] Enabling AD_ID release error acknowledgement...');
            await ensureAdIdReleaseErrorsAcknowledged();

            console.log('[RELEASE] Saving AD_ID declaration update...');
            await clickSaveWithOverflowFallback('Advertising ID release error acknowledgement');
            await waitSaved(page);
            await randomDelay(page, 5000, 7000);

            await reopenProductionDraftReleaseForReview();

            const remainingNeedFixReason = await detectReleaseNeedFixReason();
            if (remainingNeedFixReason) {
                throw createReleaseNeedFixError(remainingNeedFixReason);
            }
        }

        // 可自动修复的 release error 在这里处理；其他阻塞错误转 NEED_FIX。
        async function handleReleaseNeedFixOrThrow(reason, sourceLabel) {
            console.log(`[RELEASE] Need manual fix detected${sourceLabel ? ` ${sourceLabel}` : ''}: ${reason}`);
            if (isAdIdReleasePermissionError(reason)) {
                await resolveAdIdReleasePermissionError();
                return;
            }

            await closeOtherTabsAfterTaskDone();
            throw createReleaseNeedFixError(reason);
        }

        // Publishing overview 中发送所有待发布变更进入审核。
        async function sendChangesForReview() {
            console.log('[RELEASE] Opening Publishing overview...');
            await page.bringToFront().catch(() => { });
            if (!/\/publishing(?:[/?#]|$)/.test(String(page.url() || ''))) {
                await page.goto(`${appBasePath}/publishing`, { timeout: 120000, waitUntil: 'domcontentloaded' });
            }
            await page.waitForLoadState('load', { timeout: 60000 }).catch(() => { });
            await randomDelay(page, 5000, 7000);

            const sendForReviewButton = page.locator(
                'button[debug-id="send-for-review-button"], [role="button"][debug-id="send-for-review-button"]'
            ).filter({
                hasText: /Send\s+\d+\s+changes?\s+for\s+review/i
            }).first();

            console.log('[RELEASE] Sending changes for review...');
            await retryAction(async () => {
                await sendForReviewButton.waitFor({ state: 'visible', timeout: 90000 });
                if (await isLocatorDisabled(sendForReviewButton)) {
                    throw new Error('Send changes for review button is disabled.');
                }
                await clickLocatorRobust(sendForReviewButton, 'Send changes for review button', 15000);
            }, 'Click Send changes for review', 5);
            await randomDelay(page, 3000, 5000);

            const sendDialog = page.locator('div[role="dialog"], div[aria-modal="true"]').filter({
                hasText: /Send\s+\d+\s+changes?\s+for\s+review|These changes will be sent/i
            }).first();
            const confirmButton = sendDialog.locator(
                'button[debug-id="yes-button"], [role="button"][debug-id="yes-button"]'
            ).filter({
                hasText: /^Send changes for review$/i
            }).first();

            console.log('[RELEASE] Confirming send changes for review dialog...');
            await retryAction(async () => {
                await sendDialog.waitFor({ state: 'visible', timeout: 60000 });
                await confirmButton.waitFor({ state: 'visible', timeout: 60000 });
                if (await isLocatorDisabled(confirmButton)) {
                    throw new Error('Confirm send changes for review button is disabled.');
                }
                await clickLocatorRobust(confirmButton, 'Confirm send changes for review button', 15000);
            }, 'Confirm Send changes for review', 3);

            console.log('[RELEASE] Waiting for changes to enter review...');
            await page.waitForLoadState('networkidle', { timeout: 120000 }).catch(() => { });
            const changesInReview = page.locator('text=/Changes\\s+in\\s+review/i').first();
            const reviewConfirmation = page.locator('text=/Your changes are now in review/i').first();
            await retryAction(async () => {
                const hasChangesInReview = await changesInReview.isVisible().catch(() => false);
                const hasReviewConfirmation = await reviewConfirmation.isVisible().catch(() => false);
                if (!hasChangesInReview && !hasReviewConfirmation) {
                    throw new Error('Changes in review confirmation not visible yet.');
                }
            }, 'Wait for Changes in review confirmation', 12);
            await randomDelay(page, 5000, 7000);
        }

        // Production release 大步骤：下载 AAB、上传、保存、发送审核。
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
            console.log('[RELEASE] Create release page ready; opening same URL tab before AAB upload...');
            await randomDelay(page, 5000, 7000);

            let releaseMirrorPage = null;
            try {
                releaseMirrorPage = await openSameProductionReleaseTab(releaseUrl);

                console.log('[RELEASE] Uploading AAB package on original Production release page...');
                await uploadFilesByUploadButton(page, uploadButton, [aabPath], 'AAB');
                await randomDelay(page, 5000, 7000);
                await waitForAabUploadReady(aabPath, 300000);
                await randomDelay(page, 5000, 7000);

                console.log('[RELEASE] Switching to temporary tab for Store settings category...');
                await setStoreCategoryFromAppGenieType(releaseMirrorPage);
                releaseMirrorPage = null;
                await randomDelay(page, 5000, 7000);
            } catch (error) {
                if (releaseMirrorPage && !releaseMirrorPage.isClosed()) {
                    console.log('[RELEASE] Closing temporary Production release tab after error...');
                    await releaseMirrorPage.close().catch(() => { });
                    await page.bringToFront().catch(() => { });
                }
                throw error;
            }

            console.log('[RELEASE] Returning to Production release editor...');
            await page.bringToFront().catch(() => { });
            if (page.url() !== releaseUrl) {
                await page.goto(releaseUrl, { timeout: 120000, waitUntil: 'domcontentloaded' });
            }
            await randomDelay(page, 5000, 7000);
            console.log('[RELEASE] Clicking Next...');
            await clickMainButton('Next');
            await randomDelay(page, 5000, 7000);
            const releaseNeedFixReason = await detectReleaseNeedFixReason();
            if (releaseNeedFixReason) {
                await handleReleaseNeedFixOrThrow(releaseNeedFixReason, '');
            }
            console.log('[RELEASE] Saving Production release...');
            try {
                await clickMainButton('Save');
            } catch (saveError) {
                const lateNeedFixReason = await detectReleaseNeedFixReason();
                if (lateNeedFixReason) {
                    await handleReleaseNeedFixOrThrow(lateNeedFixReason, 'after Save failed');
                    console.log('[RELEASE] Saving Production release after AD_ID fix...');
                    await clickMainButton('Save');
                } else {
                    throw saveError;
                }
            }
            await randomDelay(page, 5000, 7000);

            const goOverview = page.locator(
                'button[debug-id="yes-button"], button:has-text("Go to overview"), [role="button"]:has-text("Go to overview")'
            ).first();
            if (await goOverview.isVisible().catch(() => false)) {
                console.log('[RELEASE] Going back to overview...');
                await clickLocatorRobust(goOverview, 'Go to overview button', 10000);
                await randomDelay(page, 5000, 7000);
            }

            await sendChangesForReview();
            markStepDone(PROGRESS_STEP_RELEASE_DONE);
        }

        // Privacy policy 大步骤：AppGenie 取文本，Google Sites 发布，回填 Play Console URL。
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

        // 内容分级问卷里，每次选择第一个未回答组的 No。
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

        // 统计内容分级问卷还剩多少未回答 radio group。
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

        // Content ratings 大步骤：选择 All Other App Types，并把问卷统一回答 No。
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

        // Data safety 大步骤：按当前合规策略声明 Device or other IDs 的收集和用途。
        async function runDataSafetyStep() {
            await goToAppContent();
            console.log('Executing declaration 10/13: Data safety...');
            await clickStartDeclaration('Data safety');

            // Data safety 按钮可用性判断，过滤掉可见但 disabled 的按钮。
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

            // 点击已经定位好的 Data safety 按钮，并统一打印日志。
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

            // Next 按钮在页面上可能有多个候选，优先取底部可见的真实按钮。
            const getDataSafetyNextButton = async () => {
                const primary = page.locator(DATA_SAFETY_PRIMARY_NEXT_BUTTON_SELECTOR)
                    .filter({ hasText: /^\s*Next\s*$/ })
                    .last();
                if (await primary.isVisible().catch(() => false)) {
                    return primary;
                }

                const candidates = page.locator(DATA_SAFETY_NEXT_BUTTON_SELECTORS.join(', '));
                const count = await candidates.count().catch(() => 0);
                const states = [];
                for (let index = 0; index < count; index++) {
                    const candidate = candidates.nth(index);
                    const state = await candidate.evaluate((el) => {
                        const text = String(el.textContent || '').replace(/\s+/g, ' ').trim();
                        const style = window.getComputedStyle(el);
                        const rect = el.getBoundingClientRect();
                        const visible = style.display !== 'none' &&
                            style.visibility !== 'hidden' &&
                            rect.width > 0 &&
                            rect.height > 0;
                        const disabled = el.hasAttribute('disabled') ||
                            el.classList.contains('mdc-button--disabled') ||
                            el.getAttribute('aria-disabled') === 'true';
                        return { text, visible, disabled };
                    }).catch(() => ({ text: '', visible: false, disabled: true }));
                    states.push(state);
                }
                const selectedIndex = pickLastVisibleDataSafetyNextButton(states);
                if (selectedIndex < 0) {
                    return null;
                }
                return candidates.nth(selectedIndex);
            };

            // Data safety 页面主 Save 按钮。
            const clickMainSaveButton = async (label) => {
                const saveButton = getMainSaveButton();
                await saveButton.waitFor({ state: 'visible', timeout: 15000 });
                await clickLocatorRobust(saveButton, `Data safety ${label}`, 15000);
                console.log(`[DATA SAFETY] Clicked ${label}.`);
                await delay(page, 3000);
            };

            // 点击 Data safety 的 Next，并确保按钮可用。
            const clickDataSafetyNextButton = async (label = 'Next') => {
                const nextButton = await getDataSafetyNextButton();
                if (!nextButton) {
                    throw new Error('Data safety Next button not found.');
                }
                await nextButton.waitFor({ state: 'visible', timeout: 15000 });
                if (!(await waitForEnabled(nextButton, 15000)) || await isLocatorDisabled(nextButton)) {
                    throw new Error('Data safety Next button is disabled.');
                }
                await clickLocatorRobust(nextButton, `Data safety ${label}`, 15000);
                console.log(`[DATA SAFETY] Clicked ${label}.`);
                await delay(page, 3000);
            };

            // 选择数据类型后 Next 可能延迟解锁，这里轮询等待。
            const waitForDataSafetyNextEnabled = async (timeoutMs = 5000) => {
                const started = Date.now();
                while (Date.now() - started < timeoutMs) {
                    const nextButton = await getDataSafetyNextButton();
                    if (nextButton && await isButtonEnabled(nextButton)) {
                        return true;
                    }
                    await delay(page, 500);
                }
                return false;
            };

            const regexPayload = (regex) => ({
                source: regex.source,
                flags: regex.flags
            });

            // 以下几个判断用于识别当前 Data safety 走到哪一页。
            const isDataCollectionSecurityStepVisible = async () => {
                return await page.locator(DATA_SAFETY_COLLECTION_SECURITY_STEP_SELECTOR).first()
                    .isVisible()
                    .catch(() => false);
            };

            const isDataTypesStepVisible = async () => {
                return await page.locator('text=/^Data types$/i').first().isVisible().catch(() => false) &&
                    await page.locator('text=/Device or other IDs/i').first().isVisible().catch(() => false);
            };

            const isDataUsageStepVisible = async () => {
                return await page.locator('text=/^Data usage and handling$/i').first().isVisible().catch(() => false) &&
                    await page.locator('text=/Device or other IDs/i').first().isVisible().catch(() => false);
            };

            const isDataSafetyPreviewStepVisible = async () => {
                return await page.locator('text=/Store listing preview/i').first().isVisible().catch(() => false);
            };

            // 按“问题文案 + 答案文案”选择 radio，适配 Material DOM 层级变化。
            const clickDataSafetyRadioAnswer = async (questionRegex, answerRegex, label) => {
                await retryAction(async () => {
                    const result = await page.evaluate(({ question, answer }) => {
                        const questionPattern = new RegExp(question.source, question.flags);
                        const answerPattern = new RegExp(answer.source, answer.flags);
                        const normalize = (value) => String(value || '').replace(/\s+/g, ' ').trim();
                        const isVisible = (el) => {
                            if (!el || !(el instanceof HTMLElement)) return false;
                            const style = window.getComputedStyle(el);
                            const rect = el.getBoundingClientRect();
                            return style.display !== 'none' &&
                                style.visibility !== 'hidden' &&
                                rect.width > 0 &&
                                rect.height > 0;
                        };
                        const elementTop = (el) => el.getBoundingClientRect().top + window.scrollY;
                        const getInputLabelText = (input) => {
                            const labelledBy = String(input.getAttribute('aria-labelledby') || '').trim();
                            if (labelledBy) {
                                const text = labelledBy
                                    .split(/\s+/)
                                    .map(id => document.getElementById(id))
                                    .filter(Boolean)
                                    .map(el => normalize(el.textContent))
                                    .filter(Boolean)
                                    .join(' ');
                                if (text) return text;
                            }

                            if (input.id) {
                                const label = Array.from(document.querySelectorAll('label'))
                                    .find(candidate => candidate.htmlFor === input.id);
                                if (label) return normalize(label.textContent);
                            }

                            return '';
                        };
                        const getOptionText = (optionRoot) => {
                            const input = optionRoot.matches('input[type="radio"]')
                                ? optionRoot
                                : optionRoot.querySelector('input[type="radio"]');
                            const texts = [
                                normalize(optionRoot.textContent),
                                input ? getInputLabelText(input) : '',
                                input ? normalize((input.closest('material-radio, label, .mdc-form-field, [role="radio"]') || input.parentElement || input).textContent) : ''
                            ].filter(Boolean);
                            return texts[0] || '';
                        };
                        const isChecked = (optionRoot, input) => {
                            return Boolean(
                                (input && input.checked) ||
                                (input && input.getAttribute('aria-checked') === 'true') ||
                                optionRoot.getAttribute('aria-checked') === 'true' ||
                                optionRoot.querySelector('.mdc-radio--checked, input[type="radio"]:checked, [aria-checked="true"]')
                            );
                        };
                        const clickOption = (optionRoot, input) => {
                            optionRoot.scrollIntoView({ block: 'center', inline: 'nearest' });
                            const labelEl = input && input.id
                                ? Array.from(document.querySelectorAll('label')).find(candidate => candidate.htmlFor === input.id)
                                : null;
                            const target = labelEl || optionRoot.closest('label') || optionRoot;
                            target.click();
                            if (input && !input.checked && input.getAttribute('aria-checked') !== 'true') {
                                input.click();
                            }
                            if (input) {
                                input.dispatchEvent(new Event('input', { bubbles: true }));
                                input.dispatchEvent(new Event('change', { bubbles: true }));
                            }
                        };

                        const questionEl = Array.from(document.querySelectorAll('body *'))
                            .filter(isVisible)
                            .filter(el => questionPattern.test(normalize(el.textContent)))
                            .sort((a, b) => normalize(a.textContent).length - normalize(b.textContent).length)[0];
                        if (!questionEl) {
                            return { ok: false, reason: 'question not found' };
                        }

                        const questionY = elementTop(questionEl);
                        const groups = Array.from(document.querySelectorAll('material-radio-group, [role="radiogroup"]'))
                            .filter(isVisible)
                            .filter(group => elementTop(group) >= questionY - 5)
                            .sort((a, b) => elementTop(a) - elementTop(b));

                        for (const group of groups) {
                            const options = Array.from(group.querySelectorAll('material-radio, label, .mdc-form-field, [role="radio"], .mdc-radio'))
                                .filter(isVisible);
                            for (const optionRoot of options) {
                                const optionText = getOptionText(optionRoot);
                                if (!answerPattern.test(optionText)) continue;

                                const input = optionRoot.matches('input[type="radio"]')
                                    ? optionRoot
                                    : optionRoot.querySelector('input[type="radio"]');
                                if (!isChecked(optionRoot, input)) {
                                    clickOption(optionRoot, input);
                                }
                                return {
                                    ok: isChecked(optionRoot, input),
                                    optionText
                                };
                            }
                        }

                        return { ok: false, reason: 'answer option not found' };
                    }, {
                        question: regexPayload(questionRegex),
                        answer: regexPayload(answerRegex)
                    });

                    if (!result || !result.ok) {
                        throw new Error(`${label} not selected: ${(result && result.reason) || 'unknown reason'}`);
                    }
                }, `Data safety ${label}`, 4);
                console.log(`[DATA SAFETY] Selected ${label}.`);
                await delay(page, 1500);
            };

            // 按答案文案勾选 checkbox，优先关联 label/input，找不到再按最近位置匹配。
            const clickDataSafetyCheckbox = async (answerRegex, label) => {
                await retryAction(async () => {
                    const result = await page.evaluate(({ answer }) => {
                        const answerPattern = new RegExp(answer.source, answer.flags);
                        const normalize = (value) => String(value || '').replace(/\s+/g, ' ').trim();
                        const isVisible = (el) => {
                            if (!el || !(el instanceof HTMLElement)) return false;
                            const style = window.getComputedStyle(el);
                            const rect = el.getBoundingClientRect();
                            return style.display !== 'none' &&
                                style.visibility !== 'hidden' &&
                                rect.width > 0 &&
                                rect.height > 0;
                        };
                        const isChecked = (root, input) => Boolean(
                            (input && input.checked) ||
                            (input && input.getAttribute('aria-checked') === 'true') ||
                            (root && root.getAttribute('aria-checked') === 'true') ||
                            (root && root.querySelector('input[type="checkbox"]:checked, .mdc-checkbox--selected, [aria-checked="true"]'))
                        );

                        const labelEl = Array.from(document.querySelectorAll('body *'))
                            .filter(isVisible)
                            .filter(el => answerPattern.test(normalize(el.textContent)))
                            .sort((a, b) => normalize(a.textContent).length - normalize(b.textContent).length)[0];
                        if (!labelEl) {
                            return { ok: false, reason: 'checkbox label not found' };
                        }

                        let input = null;
                        if (labelEl.id) {
                            input = Array.from(document.querySelectorAll('input[type="checkbox"]'))
                                .find(candidate => String(candidate.getAttribute('aria-labelledby') || '')
                                    .split(/\s+/)
                                    .includes(labelEl.id));
                        }
                        if (!input && labelEl.htmlFor) {
                            input = document.getElementById(labelEl.htmlFor);
                        }

                        const root = labelEl.closest('material-checkbox, label, .mdc-form-field, [role="checkbox"]') ||
                            (input && input.closest('material-checkbox, label, .mdc-form-field, [role="checkbox"]')) ||
                            labelEl.parentElement;
                        if (!input && root) {
                            input = root.querySelector('input[type="checkbox"]');
                        }
                        if (!input) {
                            const labelY = labelEl.getBoundingClientRect().top + window.scrollY;
                            input = Array.from(document.querySelectorAll('input[type="checkbox"]'))
                                .filter(isVisible)
                                .sort((a, b) => {
                                    const aDistance = Math.abs((a.getBoundingClientRect().top + window.scrollY) - labelY);
                                    const bDistance = Math.abs((b.getBoundingClientRect().top + window.scrollY) - labelY);
                                    return aDistance - bDistance;
                                })[0];
                        }
                        if (!input) {
                            return { ok: false, reason: 'checkbox input not found' };
                        }

                        const clickRoot = root || input.closest('material-checkbox, label, .mdc-form-field, [role="checkbox"]') || input;
                        clickRoot.scrollIntoView({ block: 'center', inline: 'nearest' });
                        if (!isChecked(clickRoot, input)) {
                            const target = labelEl.closest('label') || clickRoot;
                            target.click();
                            if (!input.checked && input.getAttribute('aria-checked') !== 'true') {
                                input.click();
                            }
                            input.dispatchEvent(new Event('input', { bubbles: true }));
                            input.dispatchEvent(new Event('change', { bubbles: true }));
                        }

                        return { ok: isChecked(clickRoot, input) };
                    }, {
                        answer: regexPayload(answerRegex)
                    });

                    if (!result || !result.ok) {
                        throw new Error(`${label} not checked: ${(result && result.reason) || 'unknown reason'}`);
                    }
                }, `Data safety ${label}`, 4);
                console.log(`[DATA SAFETY] Checked ${label}.`);
                await delay(page, 1500);
            };

            // “不允许用户创建账号”选项会触发后续外部账号问题，必要时做二次点击。
            const selectNoInAppAccountCreation = async () => {
                await retryAction(async () => {
                    const accountCreationGroup = page
                        .locator('div[role="group"][aria-label*="Which of the following methods of account creation does your app support"]')
                        .first();
                    const checkboxLabel = accountCreationGroup
                        .locator('label')
                        .filter({ hasText: /^My app does not allow users to create an account$/i })
                        .first();
                    const outsideGroup = page
                        .locator('material-radio-group[debug-id="has-outside-app-accounts"]')
                        .first();

                    await accountCreationGroup.waitFor({ state: 'visible', timeout: 10000 });
                    await checkboxLabel.waitFor({ state: 'visible', timeout: 10000 });
                    await clickLocatorRobust(checkboxLabel, 'My app does not allow users to create an account label', 10000);

                    const groupVisibleAfterFirstClick = await outsideGroup
                        .waitFor({ state: 'visible', timeout: 3000 })
                        .then(() => true)
                        .catch(() => false);
                    if (!groupVisibleAfterFirstClick) {
                        const checkedAfterFirstClick = await checkboxLabel.evaluate((label) => {
                            const checkbox = label.closest('material-checkbox');
                            const input = checkbox && checkbox.querySelector('input[type="checkbox"]');
                            return Boolean(
                                (input && input.checked) ||
                                (input && input.getAttribute('aria-checked') === 'true') ||
                                (checkbox && checkbox.querySelector('.mdc-checkbox--selected, input[type="checkbox"]:checked, [aria-checked="true"]'))
                            );
                        }).catch(() => false);
                        console.log('[DATA SAFETY] Outside app login group not visible after first click; retrying label toggle.');
                        if (checkedAfterFirstClick) {
                            await clickLocatorRobust(checkboxLabel, 'My app does not allow users to create an account label', 10000);
                            await delay(page, 700);
                        }
                        await clickLocatorRobust(checkboxLabel, 'My app does not allow users to create an account label', 10000);
                    }

                    const groupVisibleAfterRetry = await outsideGroup
                        .waitFor({ state: 'visible', timeout: 3000 })
                        .then(() => true)
                        .catch(() => false);
                    if (!groupVisibleAfterRetry) {
                        throw new Error('outside app login question group not found');
                    }
                }, 'Data safety no in-app account creation', 4);
                console.log('[DATA SAFETY] Checked no in-app account creation.');
                await delay(page, 1500);
            };

            // 在指定 radio group 里选择答案，用于稳定 debug-id 的问题组。
            const clickDataSafetyRadioGroupAnswer = async (groupSelectors, answerRegex, label) => {
                await retryAction(async () => {
                    const target = await page.evaluate((config) => {
                        const answerPattern = new RegExp(config.answer.source, config.answer.flags);
                        const normalize = (value) => String(value || '').replace(/\s+/g, ' ').trim();
                        const visibleRect = (el) => {
                            if (!el || !(el instanceof HTMLElement)) return false;
                            const style = window.getComputedStyle(el);
                            const rect = el.getBoundingClientRect();
                            if (style.display !== 'none' &&
                                style.visibility !== 'hidden' &&
                                rect.width > 0 &&
                                rect.height > 0) {
                                return rect;
                            }
                            return null;
                        };
                        const isSelected = (radio) => {
                            const input = radio.querySelector('input[type="radio"]');
                            return Boolean(
                                (input && input.checked) ||
                                (input && input.getAttribute('aria-checked') === 'true') ||
                                radio.querySelector('.mdc-radio--checked, input[type="radio"]:checked, [aria-checked="true"]')
                            );
                        };
                        const groups = config.groupSelectors
                            .map(selector => document.querySelector(selector))
                            .filter(Boolean);
                        const group = groups.find((candidate) => {
                            if (visibleRect(candidate)) return true;
                            return Array.from(candidate.querySelectorAll(`${config.radioSelector} .mdc-radio, ${config.radioSelector} label`))
                                .some(visibleRect);
                        });
                        if (!group) {
                            return { ok: false, reason: `${config.label} group not visible` };
                        }
                        const directRadios = Array.from(group.children)
                            .filter(child => child.matches && child.matches(config.radioSelector));
                        const options = directRadios.map(radio => normalize(radio.querySelector('label') && radio.querySelector('label').textContent));
                        const targetRadio = directRadios.find((radio) => {
                            const optionText = normalize(radio.querySelector('label') && radio.querySelector('label').textContent);
                            return answerPattern.test(optionText);
                        }) || directRadios[config.fallbackIndex];
                        if (!targetRadio) {
                            return {
                                ok: false,
                                reason: `${config.label} target radio not found; direct radio count=${directRadios.length}; options=${options.join(' | ')}`
                            };
                        }
                        const input = targetRadio.querySelector('input[type="radio"]');
                        const labelEl = targetRadio.querySelector('label');
                        const mdcRadio = targetRadio.querySelector('.mdc-radio');
                        const optionText = normalize(labelEl && labelEl.textContent);
                        const clickTarget = mdcRadio || labelEl || input || targetRadio;
                        clickTarget.scrollIntoView({ block: 'center', inline: 'nearest' });
                        const rect = visibleRect(clickTarget) || visibleRect(labelEl) || visibleRect(mdcRadio) || visibleRect(input);
                        if (!rect) {
                            return {
                                ok: false,
                                reason: `${config.label} target has no visible click rect; option=${optionText}; options=${options.join(' | ')}`
                            };
                        }
                        return {
                            ok: true,
                            alreadySelected: isSelected(targetRadio),
                            x: rect.left + rect.width / 2,
                            y: rect.top + rect.height / 2,
                            optionText,
                            options,
                            directRadioCount: directRadios.length
                        };
                    }, {
                        groupSelectors,
                        radioSelector: DATA_SAFETY_DIRECT_MATERIAL_RADIO_SELECTOR,
                        answer: regexPayload(answerRegex),
                        fallbackIndex: DATA_SAFETY_OUTSIDE_APP_LOGIN_NO_RADIO_INDEX,
                        label
                    });

                    if (!target || !target.ok || target.optionText !== DATA_SAFETY_OUTSIDE_APP_LOGIN_NO_ANSWER_TEXT) {
                        const reason = target && target.reason
                            ? target.reason
                            : `target=${target && target.optionText}, directRadioCount=${target && target.directRadioCount}, options=${target && target.options}`;
                        throw new Error(`${label} target not found: ${reason}`);
                    }

                    if (target.alreadySelected) {
                        console.log(`[DATA SAFETY] ${label} already selected (${target.optionText}).`);
                    } else {
                        console.log(`[DATA SAFETY] ${label} options: ${(target.options || []).join(' | ')}; clicking "${target.optionText}" at ${Math.round(target.x)},${Math.round(target.y)}.`);
                        await page.mouse.click(target.x, target.y);
                        await delay(page, 1000);
                    }

                    const selected = await page.evaluate((config) => {
                        const answerPattern = new RegExp(config.answer.source, config.answer.flags);
                        const normalize = (value) => String(value || '').replace(/\s+/g, ' ').trim();
                        const group = config.groupSelectors
                            .map(selector => document.querySelector(selector))
                            .find(Boolean);
                        if (!group) return { ok: false, reason: `${config.label} group not found after click` };
                        const directRadios = Array.from(group.children)
                            .filter(child => child.matches && child.matches(config.radioSelector));
                        const targetRadio = directRadios.find((radio) => {
                            const optionText = normalize(radio.querySelector('label') && radio.querySelector('label').textContent);
                            return answerPattern.test(optionText);
                        }) || directRadios[config.fallbackIndex];
                        if (!targetRadio) return { ok: false, reason: `${config.label} target radio not found after click` };
                        const input = targetRadio.querySelector('input[type="radio"]');
                        const optionText = normalize(targetRadio.querySelector('label') && targetRadio.querySelector('label').textContent);
                        return {
                            ok: Boolean(
                                (input && input.checked) ||
                                (input && input.getAttribute('aria-checked') === 'true') ||
                                targetRadio.querySelector('.mdc-radio--checked, input[type="radio"]:checked, [aria-checked="true"]')
                            ),
                            optionText,
                            ariaChecked: input && input.getAttribute('aria-checked'),
                            checked: input && input.checked
                        };
                    }, {
                        groupSelectors,
                        radioSelector: DATA_SAFETY_DIRECT_MATERIAL_RADIO_SELECTOR,
                        answer: regexPayload(answerRegex),
                        fallbackIndex: DATA_SAFETY_OUTSIDE_APP_LOGIN_NO_RADIO_INDEX,
                        label
                    });
                    if (!selected || !selected.ok || selected.optionText !== DATA_SAFETY_OUTSIDE_APP_LOGIN_NO_ANSWER_TEXT) {
                        const reason = selected && selected.reason
                            ? selected.reason
                            : `target=${selected && selected.optionText}, checked=${selected && selected.checked}, ariaChecked=${selected && selected.ariaChecked}`;
                        throw new Error(`${label} not selected after mouse click: ${reason}`);
                    }
                }, `Data safety ${label}`, 4);
                console.log(`[DATA SAFETY] Selected ${label}.`);
                await delay(page, 1500);
            };

            // 在弹窗/局部范围内勾选 checkbox，并可等待它触发的后续问题出现。
            const clickDataSafetyCheckboxInScope = async (scopeLocator, answerRegex, label, selector = '', revealsSelector = '') => {
                await retryAction(async () => {
                    await scopeLocator.waitFor({ state: 'visible', timeout: 30000 });
                    const findTarget = async () => scopeLocator.evaluate((scope, { answer, selector }) => {
                        const answerPattern = new RegExp(answer.source, answer.flags);
                        const normalize = (value) => String(value || '').replace(/\s+/g, ' ').trim();
                        const isVisible = (el) => {
                            if (!el || !(el instanceof HTMLElement)) return false;
                            const style = window.getComputedStyle(el);
                            const rect = el.getBoundingClientRect();
                            return style.display !== 'none' &&
                                style.visibility !== 'hidden' &&
                                rect.width > 0 &&
                                rect.height > 0;
                        };
                        const isChecked = (root, input) => Boolean(
                            (input && input.checked) ||
                            (input && input.getAttribute('aria-checked') === 'true') ||
                            (root && root.getAttribute('aria-checked') === 'true') ||
                            (root && root.querySelector('input[type="checkbox"]:checked, .mdc-checkbox--selected, [aria-checked="true"]'))
                        );
                        const resolveCheckboxRoot = (candidate) => {
                            if (!candidate) return null;
                            return candidate.matches && candidate.matches('material-checkbox')
                                ? candidate
                                : candidate.closest('material-checkbox') || candidate;
                        };
                        const visibleRect = (el) => {
                            if (!isVisible(el)) return null;
                            const rect = el.getBoundingClientRect();
                            return {
                                x: rect.left,
                                y: rect.top,
                                width: rect.width,
                                height: rect.height
                            };
                        };

                        let targetRoot = selector
                            ? resolveCheckboxRoot(scope.querySelector(selector))
                            : null;
                        let input = null;
                        if (targetRoot) {
                            input = targetRoot.matches && targetRoot.matches('input[type="checkbox"]')
                                ? targetRoot
                                : targetRoot.querySelector('input[type="checkbox"]');
                        }

                        if (!targetRoot) {
                            const checkboxRoots = Array.from(scope.querySelectorAll('material-checkbox'))
                                .filter(isVisible);
                            targetRoot = checkboxRoots.find(root => answerPattern.test(normalize(root.textContent)));
                            if (targetRoot) {
                                input = targetRoot.querySelector('input[type="checkbox"]');
                            }
                        }

                        if (!targetRoot || !input) {
                            const labelEl = Array.from(scope.querySelectorAll('*'))
                                .filter(isVisible)
                                .filter(el => answerPattern.test(normalize(el.textContent)))
                                .sort((a, b) => normalize(a.textContent).length - normalize(b.textContent).length)[0];
                            if (!labelEl) {
                                return { ok: false, reason: 'checkbox label not found' };
                            }
                            if (labelEl.id) {
                                input = Array.from(scope.querySelectorAll('input[type="checkbox"]'))
                                    .find(candidate => String(candidate.getAttribute('aria-labelledby') || '')
                                        .split(/\s+/)
                                        .includes(labelEl.id));
                            }
                            if (!input && labelEl.htmlFor) {
                                input = document.getElementById(labelEl.htmlFor);
                            }
                            targetRoot = labelEl.closest('material-checkbox') ||
                                (input && input.closest('material-checkbox')) ||
                                labelEl.closest('label, .mdc-form-field, [role="checkbox"]') ||
                                (input && input.closest('label, .mdc-form-field, [role="checkbox"]')) ||
                                labelEl.parentElement;
                        }

                        if (!input && targetRoot) {
                            input = targetRoot.querySelector('input[type="checkbox"]');
                        }
                        if (!input) {
                            return { ok: false, reason: 'checkbox input not found' };
                        }

                        const clickRoot = targetRoot || input.closest('material-checkbox, label, .mdc-form-field, [role="checkbox"]') || input;
                        clickRoot.scrollIntoView({ block: 'center', inline: 'nearest' });
                        const mdcCheckbox = clickRoot.querySelector('.mdc-checkbox') || input || clickRoot;
                        const rect = visibleRect(mdcCheckbox) || visibleRect(clickRoot) || visibleRect(input);
                        if (!rect) {
                            return { ok: false, reason: 'checkbox click target not visible' };
                        }
                        return {
                            ok: true,
                            selected: isChecked(clickRoot, input),
                            x: rect.x + rect.width / 2,
                            y: rect.y + rect.height / 2,
                            inputChecked: Boolean(input.checked),
                            ariaChecked: String(input.getAttribute('aria-checked') || ''),
                            visualSelected: Boolean(clickRoot.querySelector('.mdc-checkbox--selected, input[type="checkbox"]:checked, [aria-checked="true"]'))
                        };
                    }, {
                        answer: regexPayload(answerRegex),
                        selector
                    });

                    const isRevealVisible = async (timeoutMs = 3000) => {
                        if (!revealsSelector) return true;
                        return await scopeLocator
                            .locator(revealsSelector)
                            .first()
                            .waitFor({ state: 'visible', timeout: timeoutMs })
                            .then(() => true)
                            .catch(() => false);
                    };

                    let result = await findTarget();
                    if (!result || !result.ok) {
                        throw new Error(`${label} not checked in scope: ${(result && result.reason) || 'unknown reason'}`);
                    }
                    if (!result.selected) {
                        console.log(`[DATA SAFETY] Clicking ${label} checkbox at ${Math.round(result.x)},${Math.round(result.y)}.`);
                        await page.mouse.click(result.x, result.y);
                        await delay(page, 1000);
                    }

                    result = await findTarget();
                    if (!result || !result.ok || !result.selected) {
                        throw new Error(
                            `${label} not checked in scope: ` +
                            `checked=${result && result.inputChecked}, aria=${(result && result.ariaChecked) || 'none'}, visual=${result && result.visualSelected}`
                        );
                    }

                    if (!(await isRevealVisible(3000))) {
                        console.log(`[DATA SAFETY] ${label} is checked but dependent controls are hidden; toggling checkbox off/on once.`);
                        await page.mouse.click(result.x, result.y);
                        await delay(page, 500);
                        const toggleBack = await findTarget();
                        if (!toggleBack || !toggleBack.ok) {
                            throw new Error(`${label} checkbox not found while toggling`);
                        }
                        await page.mouse.click(toggleBack.x, toggleBack.y);
                        await delay(page, 1000);
                        result = await findTarget();
                        if (!result || !result.ok || !result.selected) {
                            throw new Error(`${label} not checked after toggle`);
                        }
                        if (!(await isRevealVisible(5000))) {
                            throw new Error(`${label} dependent controls not visible after toggle`);
                        }
                    }
                }, `Data safety ${label}`, 4);
                console.log(`[DATA SAFETY] Checked ${label}.`);
                await delay(page, 1500);
            };

            // 在弹窗/局部范围内选择 radio，优先按 groupSelector 限定问题组。
            const clickDataSafetyRadioInScope = async (scopeLocator, answerRegex, label, groupSelector = '') => {
                await retryAction(async () => {
                    await scopeLocator.waitFor({ state: 'visible', timeout: 30000 });
                    const result = await scopeLocator.evaluate((scope, { answer, groupSelector }) => {
                        const answerPattern = new RegExp(answer.source, answer.flags);
                        const normalize = (value) => String(value || '').replace(/\s+/g, ' ').trim();
                        const isVisible = (el) => {
                            if (!el || !(el instanceof HTMLElement)) return false;
                            const style = window.getComputedStyle(el);
                            const rect = el.getBoundingClientRect();
                            return style.display !== 'none' &&
                                style.visibility !== 'hidden' &&
                                rect.width > 0 &&
                                rect.height > 0;
                        };
                        const getInputLabelText = (input) => {
                            const labelledBy = String(input.getAttribute('aria-labelledby') || '').trim();
                            if (labelledBy) {
                                const text = labelledBy
                                    .split(/\s+/)
                                    .map(id => document.getElementById(id))
                                    .filter(Boolean)
                                    .map(el => normalize(el.textContent))
                                    .filter(Boolean)
                                    .join(' ');
                                if (text) return text;
                            }
                            if (input.id) {
                                const labelEl = Array.from(document.querySelectorAll('label'))
                                    .find(candidate => candidate.htmlFor === input.id);
                            if (labelEl) return normalize(labelEl.textContent);
                            }
                            return '';
                        };
                        const searchRoot = groupSelector && scope.querySelector(groupSelector)
                            ? scope.querySelector(groupSelector)
                            : scope;
                        const optionRoots = Array.from(searchRoot.querySelectorAll('material-radio'))
                            .filter(isVisible);

                        for (const optionRoot of optionRoots) {
                            const input = optionRoot.matches('input[type="radio"]')
                                ? optionRoot
                                : optionRoot.querySelector('input[type="radio"]');
                            const optionText = [
                                normalize(optionRoot.textContent),
                                input ? getInputLabelText(input) : ''
                            ].filter(Boolean).join(' ');
                            if (!answerPattern.test(optionText)) continue;

                            const checked = Boolean(
                                (input && input.checked) ||
                                (input && input.getAttribute('aria-checked') === 'true') ||
                                optionRoot.getAttribute('aria-checked') === 'true' ||
                                optionRoot.querySelector('.mdc-radio--checked, input[type="radio"]:checked, [aria-checked="true"]')
                            );
                            optionRoot.scrollIntoView({ block: 'center', inline: 'nearest' });
                            if (!checked) {
                                const target = optionRoot.querySelector('.mdc-radio') || optionRoot.querySelector('label') || optionRoot;
                                target.click();
                                if (input && !input.checked && input.getAttribute('aria-checked') !== 'true') {
                                    input.click();
                                }
                                if (input) {
                                    input.dispatchEvent(new Event('input', { bubbles: true }));
                                    input.dispatchEvent(new Event('change', { bubbles: true }));
                                }
                            }
                            const selected = Boolean(
                                (input && input.checked) ||
                                (input && input.getAttribute('aria-checked') === 'true') ||
                                optionRoot.getAttribute('aria-checked') === 'true' ||
                                optionRoot.querySelector('.mdc-radio--checked, input[type="radio"]:checked, [aria-checked="true"]')
                            );
                            return { ok: selected, optionText };
                        }

                        return { ok: false, reason: 'radio option not found' };
                    }, {
                        answer: regexPayload(answerRegex),
                        groupSelector
                    });

                    if (!result || !result.ok) {
                        throw new Error(`${label} not selected in scope: ${(result && result.reason) || 'unknown reason'}`);
                    }
                }, `Data safety ${label}`, 4);
                console.log(`[DATA SAFETY] Selected ${label}.`);
                await delay(page, 1500);
            };

            // Data collection and security 页：Yes/Yes/无账号/外部登录 No/删除请求 No。
            const answerDataCollectionSecurityStep = async () => {
                console.log('[DATA SAFETY] Answering Data collection and security questions...');
                for (const action of DATA_SAFETY_COLLECTION_SECURITY_ACTIONS) {
                    if (action.type === 'radio') {
                        if (action.label === 'outside app login = No') {
                            await clickDataSafetyRadioGroupAnswer(
                                DATA_SAFETY_OUTSIDE_APP_LOGIN_GROUP_SELECTORS,
                                action.answer,
                                action.label
                            );
                        } else if (action.label === 'data deletion request = No') {
                            await retryAction(async () => {
                                const target = await page.evaluate((config) => {
                                    const normalize = (value) => String(value || '').replace(/\s+/g, ' ').trim();
                                    const visibleRect = (el) => {
                                        if (!el || !(el instanceof HTMLElement)) return false;
                                        const style = window.getComputedStyle(el);
                                        const rect = el.getBoundingClientRect();
                                        if (style.display !== 'none' &&
                                            style.visibility !== 'hidden' &&
                                            rect.width > 0 &&
                                            rect.height > 0) {
                                            return rect;
                                        }
                                        return null;
                                    };
                                    const groups = config.groupSelectors
                                        .map(selector => document.querySelector(selector))
                                        .filter(Boolean);
                                    const group = groups.find((candidate) => {
                                        if (visibleRect(candidate)) return true;
                                        return Array.from(candidate.querySelectorAll(`${config.radioSelector} .mdc-radio, ${config.radioSelector} label`))
                                            .some(visibleRect);
                                    });
                                    if (!group) {
                                        return { ok: false, reason: 'data-deletion group not visible' };
                                    }
                                    const directRadios = Array.from(group.children)
                                        .filter(child => child.matches && child.matches(config.radioSelector));
                                    const targetRadio = directRadios.find((radio) => {
                                        const label = radio.querySelector('label');
                                        return normalize(label && label.textContent) === config.answerText;
                                    }) || directRadios[config.fallbackIndex];
                                    if (!targetRadio) {
                                        return {
                                            ok: false,
                                            reason: `target radio not found; direct radio count=${directRadios.length}`
                                        };
                                    }
                                    const input = targetRadio.querySelector('input[type="radio"]');
                                    const label = targetRadio.querySelector('label');
                                    const mdcRadio = targetRadio.querySelector('.mdc-radio');
                                    const clickTarget = mdcRadio || label || input || targetRadio;
                                    clickTarget.scrollIntoView({ block: 'center', inline: 'nearest' });
                                    const rect = visibleRect(clickTarget) || visibleRect(label) || visibleRect(mdcRadio) || visibleRect(input);
                                    if (!rect) {
                                        return {
                                            ok: false,
                                            reason: `target radio has no visible click rect; label=${normalize(label && label.textContent)}; direct radio count=${directRadios.length}`
                                        };
                                    }
                                    return {
                                        ok: true,
                                        x: rect.left + rect.width / 2,
                                        y: rect.top + rect.height / 2,
                                        label: normalize(label && label.textContent),
                                        directRadioCount: directRadios.length,
                                        options: directRadios.map(radio => normalize(radio.querySelector('label') && radio.querySelector('label').textContent))
                                    };
                                }, {
                                    groupSelectors: DATA_SAFETY_DATA_DELETION_GROUP_SELECTORS,
                                    radioSelector: DATA_SAFETY_DIRECT_MATERIAL_RADIO_SELECTOR,
                                    answerText: DATA_SAFETY_DATA_DELETION_NO_ANSWER_TEXT,
                                    fallbackIndex: DATA_SAFETY_DATA_DELETION_NO_RADIO_INDEX
                                });

                                if (!target || !target.ok || target.label !== DATA_SAFETY_DATA_DELETION_NO_ANSWER_TEXT) {
                                    const reason = target && target.reason
                                        ? target.reason
                                        : `target=${target && target.label}, directRadioCount=${target && target.directRadioCount}, options=${target && target.options}`;
                                    throw new Error(`Data deletion request = No target not found: ${reason}`);
                                }

                                console.log(`[DATA SAFETY] Data deletion options: ${(target.options || []).join(' | ')}; clicking "${target.label}" at ${Math.round(target.x)},${Math.round(target.y)}.`);
                                await page.mouse.click(target.x, target.y);
                                await delay(page, 1000);

                                const selected = await page.evaluate((config) => {
                                    const normalize = (value) => String(value || '').replace(/\s+/g, ' ').trim();
                                    const group = config.groupSelectors
                                        .map(selector => document.querySelector(selector))
                                        .find(Boolean);
                                    if (!group) return { ok: false, reason: 'data-deletion group not found after click' };
                                    const directRadios = Array.from(group.children)
                                        .filter(child => child.matches && child.matches(config.radioSelector));
                                    const targetRadio = directRadios.find((radio) => {
                                        const label = radio.querySelector('label');
                                        return normalize(label && label.textContent) === config.answerText;
                                    });
                                    if (!targetRadio) return { ok: false, reason: 'target radio not found after click' };
                                    const input = targetRadio.querySelector('input[type="radio"]');
                                    return {
                                        ok: Boolean(
                                            (input && input.checked) ||
                                            (input && input.getAttribute('aria-checked') === 'true') ||
                                            targetRadio.querySelector('.mdc-radio--checked, input[type="radio"]:checked, [aria-checked="true"]')
                                        ),
                                        ariaChecked: input && input.getAttribute('aria-checked'),
                                        checked: input && input.checked
                                    };
                                }, {
                                    groupSelectors: DATA_SAFETY_DATA_DELETION_GROUP_SELECTORS,
                                    radioSelector: DATA_SAFETY_DIRECT_MATERIAL_RADIO_SELECTOR,
                                    answerText: DATA_SAFETY_DATA_DELETION_NO_ANSWER_TEXT
                                });
                                if (!selected || !selected.ok) {
                                    const reason = selected && selected.reason
                                        ? selected.reason
                                        : `checked=${selected && selected.checked}, ariaChecked=${selected && selected.ariaChecked}`;
                                    throw new Error(`Data deletion request = No not selected after mouse click: ${reason}`);
                                }
                            }, `Data safety ${action.label}`, 4);
                        } else {
                            await clickDataSafetyRadioAnswer(action.question, action.answer, action.label);
                        }
                    } else if (action.type === 'checkbox') {
                        if (action.label === 'no in-app account creation') {
                            await selectNoInAppAccountCreation();
                        } else {
                            await clickDataSafetyCheckbox(action.answer, action.label);
                        }
                    }
                }
                await delay(page, 2000);
            };

            // Data types 页：只展开并勾选 Device or other IDs。
            const selectDataSafetyDeviceIdsType = async () => {
                console.log('[DATA SAFETY] Selecting Device or other IDs on Data types page...');
                const action = DATA_SAFETY_DATA_TYPES_ACTIONS[0];
                await retryAction(async () => {
                    const result = await page.evaluate(({ section, answer }) => {
                        const sectionPattern = new RegExp(section.source, section.flags);
                        const answerPattern = new RegExp(answer.source, answer.flags);
                        const normalize = (value) => String(value || '').replace(/\s+/g, ' ').trim();
                        const isVisible = (el) => {
                            if (!el || !(el instanceof HTMLElement)) return false;
                            const style = window.getComputedStyle(el);
                            const rect = el.getBoundingClientRect();
                            return style.display !== 'none' &&
                                style.visibility !== 'hidden' &&
                                rect.width > 0 &&
                                rect.height > 0;
                        };
                        const sectionTitle = Array.from(document.querySelectorAll('body *'))
                            .filter(isVisible)
                            .filter(el => sectionPattern.test(normalize(el.textContent)))
                            .sort((a, b) => normalize(a.textContent).length - normalize(b.textContent).length)[0];
                        if (!sectionTitle) {
                            return { ok: false, reason: 'section title not found' };
                        }

                        const sectionRoot = sectionTitle.closest('console-section, section, div[role="group"], ess-section, .ess-table-wrapper') ||
                            sectionTitle.parentElement;
                        const searchRoot = sectionRoot || document.body;
                        const showButton = Array.from(searchRoot.querySelectorAll('button, [role="button"]'))
                            .filter(isVisible)
                            .find(button => /Show/i.test(normalize(button.textContent)) ||
                                sectionPattern.test(button.getAttribute('aria-label') || '')) ||
                            Array.from(document.querySelectorAll('button, [role="button"]'))
                                .filter(isVisible)
                                .find(button => sectionPattern.test(button.getAttribute('aria-label') || '') &&
                                    /Show content/i.test(button.getAttribute('aria-label') || ''));
                        if (showButton) {
                            showButton.scrollIntoView({ block: 'center', inline: 'nearest' });
                            showButton.click();
                        }
                        return { ok: true };
                    }, {
                        section: regexPayload(action.section),
                        answer: regexPayload(action.answer)
                    });

                    if (!result || !result.ok) {
                        throw new Error(`Device or other IDs Show not clicked: ${(result && result.reason) || 'unknown reason'}`);
                    }
                }, 'Open Device or other IDs data type', 4);
                await delay(page, 2000);

                // 找 Device or other IDs checkbox 的可点击坐标和当前选中状态。
                const findDeviceIdsCheckboxTarget = async () => {
                    return await page.evaluate(({ answerText }) => {
                        const normalize = (value) => String(value || '').replace(/\s+/g, ' ').trim();
                        const targetText = normalize(answerText).toLowerCase();
                        const isVisible = (el) => {
                            if (!el || !(el instanceof HTMLElement)) return false;
                            const style = window.getComputedStyle(el);
                            const rect = el.getBoundingClientRect();
                            return style.display !== 'none' &&
                                style.visibility !== 'hidden' &&
                                rect.width > 0 &&
                                rect.height > 0;
                        };
                        const getVisibleRect = (el) => {
                            if (!isVisible(el)) return null;
                            const rect = el.getBoundingClientRect();
                            return {
                                x: rect.left,
                                y: rect.top,
                                width: rect.width,
                                height: rect.height
                            };
                        };
                        const getLabelledText = (input) => {
                            const labelledBy = input ? String(input.getAttribute('aria-labelledby') || '') : '';
                            return labelledBy
                                .split(/\s+/)
                                .map(id => document.getElementById(id))
                                .filter(Boolean)
                                .map(el => normalize(el.textContent))
                                .join(' ');
                        };
                        const checkboxRoots = Array.from(document.querySelectorAll('material-checkbox'))
                            .filter(isVisible);

                        for (const root of checkboxRoots) {
                            const input = root.querySelector('input[type="checkbox"]');
                            const optionText = normalize(`${root.textContent || ''} ${getLabelledText(input)}`);
                            if (!optionText.toLowerCase().includes(targetText)) continue;
                            const disabled = root.getAttribute('aria-disabled') === 'true' ||
                                root.classList.contains('disabled') ||
                                Boolean(input && input.disabled);
                            if (disabled) continue;

                            const checkbox = root.querySelector('.mdc-checkbox') || input || root;
                            checkbox.scrollIntoView({ block: 'center', inline: 'nearest' });
                            const checkboxRect = getVisibleRect(checkbox);
                            const labelRect = getVisibleRect(root.querySelector('.checkbox-content, label')) || checkboxRect;
                            const clickRect = checkboxRect || labelRect || getVisibleRect(root);
                            if (!clickRect) continue;

                            const inputChecked = Boolean(input && input.checked);
                            const ariaChecked = input
                                ? String(input.getAttribute('aria-checked') || '')
                                : String(root.getAttribute('aria-checked') || '');
                            const visualSelected = Boolean(
                                root.querySelector('.mdc-checkbox--selected, input[type="checkbox"]:checked, [aria-checked="true"]')
                            );
                            return {
                                ok: true,
                                optionText,
                                x: clickRect.x + clickRect.width / 2,
                                y: clickRect.y + clickRect.height / 2,
                                labelX: labelRect ? labelRect.x + labelRect.width / 2 : clickRect.x + clickRect.width / 2,
                                labelY: labelRect ? labelRect.y + labelRect.height / 2 : clickRect.y + clickRect.height / 2,
                                inputChecked,
                                ariaChecked,
                                visualSelected,
                                modelSelected: inputChecked || ariaChecked === 'true'
                            };
                        }

                        return { ok: false, reason: 'Device or other IDs checkbox not found' };
                    }, { answerText: DATA_SAFETY_DEVICE_IDS_CHECKBOX_TEXT });
                };

                // UI 已显示选中但 Next 未解锁时，手动派发 input/change 事件同步状态。
                const dispatchDeviceIdsCheckboxEvents = async () => {
                    return await page.evaluate(({ answerText }) => {
                        const normalize = (value) => String(value || '').replace(/\s+/g, ' ').trim();
                        const targetText = normalize(answerText).toLowerCase();
                        const roots = Array.from(document.querySelectorAll('material-checkbox'));
                        for (const root of roots) {
                            const input = root.querySelector('input[type="checkbox"]');
                            const text = normalize(root.textContent).toLowerCase();
                            if (!input || !text.includes(targetText)) continue;
                            input.focus();
                            input.dispatchEvent(new Event('input', { bubbles: true }));
                            input.dispatchEvent(new Event('change', { bubbles: true }));
                            input.blur();
                            return {
                                ok: true,
                                inputChecked: Boolean(input.checked),
                                ariaChecked: String(input.getAttribute('aria-checked') || '')
                            };
                        }
                        return { ok: false, reason: 'Device or other IDs checkbox input not found' };
                    }, { answerText: DATA_SAFETY_DEVICE_IDS_CHECKBOX_TEXT });
                };

                await retryAction(async () => {
                    let target = await findDeviceIdsCheckboxTarget();
                    if (!target || !target.ok) {
                        throw new Error(`Device or other IDs checkbox not selected: ${(target && target.reason) || 'unknown reason'}`);
                    }

                    if (!target.modelSelected) {
                        console.log(
                            `[DATA SAFETY] Clicking Device or other IDs checkbox at ` +
                            `${Math.round(target.x)},${Math.round(target.y)} ` +
                            `(visual=${target.visualSelected}, checked=${target.inputChecked}, aria=${target.ariaChecked || 'none'}).`
                        );
                        await page.mouse.click(target.x, target.y);
                        await delay(page, 1000);
                    }

                    target = await findDeviceIdsCheckboxTarget();
                    if (!target || !target.ok || (!target.modelSelected && !target.visualSelected)) {
                        throw new Error(`Device or other IDs checkbox not selected: ${(target && target.reason) || 'unknown reason'}`);
                    }

                    if (!(await waitForDataSafetyNextEnabled(DATA_SAFETY_DEVICE_IDS_SYNC_WAIT_MS))) {
                        console.log('[DATA SAFETY] Device or other IDs selected but Next is still disabled; syncing checkbox state...');
                        target = await findDeviceIdsCheckboxTarget();
                        if (!target || !target.ok) {
                            throw new Error(`Device or other IDs checkbox not found for sync: ${(target && target.reason) || 'unknown reason'}`);
                        }
                        if (!target.modelSelected) {
                            await page.mouse.click(target.x, target.y);
                        } else {
                            const syncResult = await dispatchDeviceIdsCheckboxEvents();
                            if (!syncResult || !syncResult.ok) {
                                throw new Error(`Device or other IDs sync failed: ${(syncResult && syncResult.reason) || 'unknown reason'}`);
                            }
                        }
                        await delay(page, 1000);
                        if (!(await waitForDataSafetyNextEnabled(DATA_SAFETY_DEVICE_IDS_SYNC_WAIT_MS))) {
                            const beforeToggle = await findDeviceIdsCheckboxTarget();
                            if (beforeToggle && beforeToggle.ok) {
                                console.log('[DATA SAFETY] Device or other IDs Next is still disabled; toggling checkbox off/on once.');
                                await page.mouse.click(beforeToggle.x, beforeToggle.y);
                                await delay(page, 500);
                                const toggleBack = await findDeviceIdsCheckboxTarget();
                                if (toggleBack && toggleBack.ok) {
                                    await page.mouse.click(toggleBack.x, toggleBack.y);
                                    await delay(page, 1000);
                                }
                            }
                        }
                        if (!(await waitForDataSafetyNextEnabled(DATA_SAFETY_DEVICE_IDS_SYNC_WAIT_MS))) {
                            const afterSync = await findDeviceIdsCheckboxTarget();
                            throw new Error(
                                'Data safety Next button still disabled after Device or other IDs sync ' +
                                `(visual=${afterSync && afterSync.visualSelected}, ` +
                                `checked=${afterSync && afterSync.inputChecked}, ` +
                                `aria=${(afterSync && afterSync.ariaChecked) || 'none'}).`
                            );
                        }
                    }
                }, 'Select Device or other IDs data type', 4);
                console.log('[DATA SAFETY] Device or other IDs data type selected.');
                await delay(page, 2000);
            };

            // Device or other IDs 弹窗：Collected、ephemeral、required、App functionality。
            const answerDataSafetyDeviceIdsUsage = async () => {
                console.log('[DATA SAFETY] Answering Device or other IDs usage questions...');
                const startAction = DATA_SAFETY_USAGE_ACTIONS.find(action => action.type === 'start');
                await retryAction(async () => {
                    const startButton = page.locator(
                        'button[aria-label="Open Device or other IDs questions"], [role="button"][aria-label="Open Device or other IDs questions"]'
                    ).first();
                    if (await startButton.isVisible().catch(() => false)) {
                        await clickLocatorRobust(startButton, 'Device or other IDs Start button', 10000);
                        return;
                    }
                    const fallback = page.locator('button, [role="button"]').filter({
                        hasText: /^Start$/i
                    }).first();
                    await clickLocatorRobust(fallback, 'Device or other IDs Start button', 10000);
                }, startAction.label, 4);
                await delay(page, 3000);

                const dialog = page.locator('div[role="dialog"], material-dialog, .mdc-dialog, .mat-mdc-dialog-container, .cdk-overlay-pane')
                    .filter({ hasText: /Device or other IDs/i })
                    .first();
                await dialog.waitFor({ state: 'visible', timeout: 60000 });

                for (const action of DATA_SAFETY_USAGE_ACTIONS) {
                    if (action.type === 'checkbox') {
                        await clickDataSafetyCheckboxInScope(
                            dialog,
                            action.answer,
                            action.label,
                            action.selector || '',
                            action.revealsSelector || ''
                        );
                    } else if (action.type === 'radio') {
                        await clickDataSafetyRadioInScope(dialog, action.answer, action.label, action.groupSelector || '');
                    } else if (action.type === 'save') {
                        const saveButton = dialog.locator(
                            'button[debug-id="save-button"], button:has-text("Save"), [role="button"]:has-text("Save")'
                        ).last();
                        await retryAction(async () => {
                            await saveButton.waitFor({ state: 'visible', timeout: 30000 });
                            if (await isLocatorDisabled(saveButton)) {
                                throw new Error('Device or other IDs Save button is disabled.');
                            }
                            await clickLocatorRobust(saveButton, 'Device or other IDs Save button', 10000);
                        }, 'Save Device or other IDs usage answers', 4);
                        await dialog.waitFor({ state: 'hidden', timeout: 60000 }).catch(() => { });
                        await waitSaved(page);
                    }
                }
                await delay(page, 2000);
            };

            for (let i = 0; i < 30; i++) {
                const nextBtn = await getDataSafetyNextButton();
                const saveBtn = getMainSaveButton();

                if (await isDataCollectionSecurityStepVisible()) {
                    await answerDataCollectionSecurityStep();
                    await retryAction(async () => {
                        await clickDataSafetyNextButton('Next');
                    }, 'Click Data safety Next after collection/security answers', 3);
                    continue;
                }
                if (await isDataTypesStepVisible()) {
                    await selectDataSafetyDeviceIdsType();
                    await retryAction(async () => {
                        await clickDataSafetyNextButton('Next');
                    }, 'Click Data safety Next after Device or other IDs selection', 3);
                    continue;
                }
                if (await isDataUsageStepVisible()) {
                    await answerDataSafetyDeviceIdsUsage();
                }

                if (await isDataSafetyPreviewStepVisible()) {
                    if (await isButtonEnabled(saveBtn)) {
                        await clickMainSaveButton('Save (preview)');
                    } else {
                        await clickOverflowSaveFallback('Data safety preview');
                    }
                    await waitSaved(page);
                    await delay(page, 3000);
                    markStepDone(PROGRESS_STEP_SAFETY_DONE);
                    return;
                }

                // Follow document flow: answer current step, Next through pages, Save at end.
                if (nextBtn && await isButtonEnabled(nextBtn)) {
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

        // --- 主执行顺序：当前仍是严格单向流程，不根据 progress_step 做真正断点跳转。 ---
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
        console.log('Executing declaration 2/7: Sign-in details...');
        await clickStartDeclaration('Sign-in details');
        await selectRadio(/^No/);
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
        await clickFinancialFeaturesSave();
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

            // 国家地区设置：必须确认真的进入目标页面后再继续，避免左侧导航点击被吞。
            const testAndReleaseUrl = `${appBasePath}/test-and-release`;
            const productionUrl = `${appBasePath}/tracks/production`;
            const productionCountryAvailabilityUrl = `${productionUrl}?tab=countryAvailability`;
            const countriesRegionsPattern = /Countries\s*\/\s*regions/i;
            const addCountriesRegionsPattern = /Add countries\s*\/\s*regions/i;
            const testAndReleaseLink = page.locator(
                'a[href*="/test-and-release"], a.item-link:has(.item-label:has-text("Test and release"))'
            ).first();
            const productionLink = page.locator(
                'a[href*="/tracks/production"], a.item-link:has(.item-label:has-text("Production"))'
            ).first();
            const countriesRegionsTab = page.locator(
                '[role="tab"], tab-button, a, button'
            ).filter({ hasText: countriesRegionsPattern }).first();
            const addCountriesBtn = page.locator(
                'button, [role="button"], a'
            ).filter({ hasText: addCountriesRegionsPattern }).first();

            // 进入 Test and release。
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

            // 进入 Production。
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
                    console.log('[RELEASE] Production page loaded but Countries/regions controls are not visible yet; continuing with fallback navigation.');
                }
            }, 'Open Production', 3);
            await delay(page, 3000);

            // 打开 Countries / regions；标签页不可见时用直达 URL 兜底。
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

            // 打开添加国家/地区弹窗。
            console.log('Clicking "Add countries / regions"...');
            await retryAction(async () => {
                await addCountriesBtn.waitFor({ state: 'visible', timeout: 60000 });
                await addCountriesBtn.scrollIntoViewIfNeeded({ timeout: 5000 });
                await addCountriesBtn.click({ timeout: 10000 });
            }, 'Click Add countries / regions', 3);
            await delay(page, 5000);

            // 勾选全部国家/地区，优先使用新页面的 Select all rows checkbox。
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

                // 旧页面兜底：从表头行里寻找全选 checkbox。
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

            // 保存国家/地区选择。
            console.log('Saving countries/regions selection...');
            await clickMainButton('Save');

            // 验证 Targeted (N) 出现，确认国家地区保存完成。
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
            await runReviewScreenshotUploadStep();

            statusManager.ensureTaskProgressAtLeast(task, PROGRESS_STEP_DONE);
            statusManager.updateTaskStatus(task, STATUS_DONE);
            task.status = STATUS_DONE;
            await closeOtherTabsAfterTaskDone();
            cleanupTempDownloadRoot();

        const durationSeconds = Math.round((Date.now() - startTime) / 1000);
        console.log(`Finished: ${appName} (${packageName}), Duration: ${formatDuration(durationSeconds)}`);
        // Keep the working tab for reuse in next iteration; only disconnect CDP client.
        await browser.close();
    } catch (err) {
        console.error(`[FATAL ERROR] In iteration: ${err.message}`);
        // Keep the working tab for reuse/debugging; only disconnect CDP client.
        if (browser) {
            const contextToClean = page && !page.isClosed()
                ? page.context()
                : browser.contexts()[0];
            if (contextToClean) {
                await closeAutomationAuxiliaryPages(contextToClean, page, 'after error').catch(() => { });
            }
            cleanupTempDownloadRoot();
            await browser.close().catch(() => { });
        }
        throw err;
    }
}

// 程序入口：准备配置、读取任务、逐条执行，并在最后输出批次汇总。
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
    // 当前版本会执行读取到的全部任务；后续如需跳过 DONE，应在这里筛选。
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

    // 批次统计只记录本次运行结果，不影响 Excel/CSV 中的持久状态。
    const runStats = {
        totalLoaded: tasks.length,
        planned: selectedTasks.length,
        success: 0,
        successItems: [],
        needFix: [],
        failed: []
    };
    const fallbackAabCleanupPaths = [];
    runtimeOptions.fallbackAabCleanupPaths = fallbackAabCleanupPaths;

    // 单条失败不会阻断整批；只有 CDP 等环境级错误才会继续向外抛出。
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
            const isReleaseNeedFix = errorCode === 'RELEASE_NEED_FIX' || Boolean(e && e.needFix);
            if (isReleaseNeedFix) {
                // Release 阶段识别到必须人工处理的问题时，标记 NEED_FIX 后继续下一条。
                task.status = STATUS_NEED_FIX;
                statusManager.updateTaskStatus(task, STATUS_NEED_FIX);
                runStats.needFix.push({
                    appName: task.appName,
                    packageName: task.packageName,
                    reason: String((e && e.message) || 'Release needs manual fixes.')
                });
                console.warn(`[NEED_FIX] Row ${task.rowNumber} (${task.appName}/${task.packageName}): ${e.message}`);
                console.warn(`Iteration ${i + 1} needs manual fix; continuing to next row...`);
            } else {
                // 普通异常标记 FAILED，并尽量继续后续任务。
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
                    // CDP 连接错误通常说明浏览器/端口环境失效，继续跑下一条也会失败。
                    throw e;
                }
            }
        }

        // 每包固定 9 行数据；第 4 行结束后清理一次 Chrome Downloads 兜底 AAB，避免中途占用磁盘。
        if (shouldCleanupFallbackAabAfterRow(i + 1)) {
            console.log(`[CLEANUP] Running fallback AAB cleanup after ${i + 1} processed row(s)...`);
            const cleanupResult = cleanupTrackedFallbackAabFiles(fallbackAabCleanupPaths);
            fallbackAabCleanupPaths.length = 0;
            console.log(
                `[CLEANUP] Fallback AAB cleanup complete: ` +
                `removed=${cleanupResult.removed}, failed=${cleanupResult.failed}, skipped=${cleanupResult.skipped}.`
            );
        }

        if (i < selectedTasks.length - 1) {
            const wait = 15000 + Math.random() * 10000;
            console.log('Wait until next iteration:', formatDuration(Math.round(wait / 1000)));
            await new Promise(r => setTimeout(r, wait));
        }
    }

    // 控制台和可选 summary JSON 使用同一份汇总内容。
    const failedCount = runStats.failed.length;
    const needFixCount = runStats.needFix.length;
    const failedNames = runStats.failed.map(item => `${item.appName} (${item.packageName})`);
    const needFixNames = runStats.needFix.map(item => `${item.appName} (${item.packageName})`);
    const successNames = runStats.successItems.map(item => `${item.appName} (${item.packageName})`);
    const summaryLines = [
        `Total loaded: ${runStats.totalLoaded}`,
        `Planned: ${runStats.planned}`,
        `Success: ${runStats.success}`,
        `Need fix: ${needFixCount}`,
        `Failed: ${failedCount}`
    ];
    if (successNames.length > 0) {
        summaryLines.push('Successful apps:');
        for (const name of successNames) {
            summaryLines.push(`[OK] ${name}`);
        }
    }
    if (needFixCount > 0) {
        summaryLines.push('Need fix apps:');
        for (const name of needFixNames) {
            summaryLines.push(`[NEED_FIX] ${name}`);
        }
    }
    if (failedCount > 0) {
        summaryLines.push('Failed apps:');
        for (const name of failedNames) {
            summaryLines.push(`[FAIL] ${name}`);
        }
    }
    const summaryText = summaryLines.join('\n');
    const coloredSummaryText = summaryLines
        .map(line => line.startsWith('Need fix') || line.startsWith('[NEED_FIX]')
            ? `\x1b[33m${line}\x1b[0m`
            : line)
        .join('\n');
    const summaryPayload = {
        generatedAt: new Date().toISOString(),
        totalLoaded: runStats.totalLoaded,
        planned: runStats.planned,
        success: runStats.success,
        successItems: runStats.successItems,
        needFixCount,
        needFix: runStats.needFix,
        failedCount,
        failed: runStats.failed,
        summaryText
    };

    console.log('================ Run Summary ================');
    console.log(coloredSummaryText);
    console.log('=============================================');
    writeRunSummaryFile(summaryPayload);
})().catch(err => {
    console.error(`[INIT ERROR] ${err.message}`);
    process.exit(1);
});




