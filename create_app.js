const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');
const { chromium } = require('playwright');

// 默认延迟（毫秒），如果 VPS 较慢可适当增大
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
const DEVELOPER_URL_CONFIG_FILE = 'developer_url.txt';
const DEVELOPER_URL_TEMPLATE = [
    '# Paste your Play Console developer URL below (single line).',
    '# Example:',
    '# https://play.google.com/console/u/0/developers/1234567890123456789/app-list',
    '',
    //'https://play.google.com/console/u/0/developers/REPLACE_WITH_YOUR_DEVELOPER_ID/app-list'
    'https://play.google.com/console/u/0/developers/5719511147760424406/app-list'
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
        throw new Error('Usage: node create_app.js [excel_file_path]');
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
            throw new Error(`Excel file not found: ${explicitPath}`);
        }
        return explicitPath;
    }

    const excelFiles = fs.readdirSync(process.cwd())
        .filter(name => ['.xlsx', '.xls'].includes(path.extname(name).toLowerCase()))
        .sort();

    if (!excelFiles.length) {
        throw new Error(
            'No Excel file found in project root. Put one .xlsx/.xls file here, or pass path: node create_app.js ./apps.xlsx'
        );
    }

    return path.resolve(process.cwd(), excelFiles[0]);
}

function pickHeader(headers, candidates) {
    for (const header of headers) {
        if (candidates.has(normalizeHeader(header))) {
            return header;
        }
    }
    return null;
}

function loadTasksFromExcel(filePath) {
    const workbook = XLSX.readFile(filePath);
    const firstSheetName = workbook.SheetNames[0];
    if (!firstSheetName) {
        throw new Error(`No sheet found in Excel file: ${filePath}`);
    }

    const sheet = workbook.Sheets[firstSheetName];
    const rows = XLSX.utils.sheet_to_json(sheet, { defval: '', raw: false });
    if (!rows.length) {
        throw new Error(`Excel sheet "${firstSheetName}" has no data rows.`);
    }

    const headers = Object.keys(rows[0]);
    const appNameHeader = pickHeader(headers, APP_NAME_HEADER_CANDIDATES);
    const packageNameHeader = pickHeader(headers, PACKAGE_NAME_HEADER_CANDIDATES);

    if (!appNameHeader || !packageNameHeader) {
        throw new Error(
            `Missing required columns. Found headers: ${headers.join(', ')}. ` +
            'Need columns like: 应用名称 / 应用包名 (or App Name / App Package Name).'
        );
    }

    const tasks = [];

    rows.forEach((row, idx) => {
        const excelRowNumber = idx + 2; // Excel 行号（从 1 开始，包含表头）
        const appName = String(row[appNameHeader] || '').trim();
        const rawPackageName = String(row[packageNameHeader] || '').trim();
        const packageName = rawPackageName.toLowerCase();

        if (!appName && !packageName) {
            return;
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
            rowNumber: excelRowNumber
        });
    });

    if (!tasks.length) {
        throw new Error(`No valid rows found in sheet "${firstSheetName}".`);
    }

    return { tasks, sheetName: firstSheetName };
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

// 全局重试包装器，用于 UI 交互
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
    // 宽松匹配各种 “saved” 提示（不区分大小写）
    try {
        await page.locator('text=/saved/i').first().waitFor({ state: 'visible', timeout: 30000 });
        console.log('Save detected.');
    } catch (e) {
        console.log('Save confirmation timeout, continuing...');
    }
    await delay(page, 5000); // 保存后额外等待，给后端处理留时间
}

async function runOnce(task, appListUrl) {
    const startTime = Date.now();
    const appName = task.appName;
    const packageName = task.packageName;
    console.log(`Creating app. Name="${appName}", Package="${packageName}"`);

    let browser;
    try {
        const cdpConnection = await connectBrowserOverCdp();
        browser = cdpConnection.browser;
        console.log(`Connected via CDP: ${cdpConnection.endpoint}`);
        const context = browser.contexts()[0];
        let page = context.pages().find(p => p.url().includes('play.google.com'));

        if (!page) {
            page = await context.newPage();
        }

        await page.bringToFront();
        page.setDefaultTimeout(60000); // 设置页面全局超时为 60 秒

        // 先进入开发者账户列表页
        console.log('Navigating to developer picker page...');
        await page.goto('https://play.google.com/console/u/0/developers', { timeout: 90000, waitUntil: 'domcontentloaded' });
        await delay(page, 8000);

        // 检查是否在开发者选择页面，如果是，则点击第一个开发者项
        const devItem = page.locator('developer-item, [debug-id="all-developers"]').first();
        if (await devItem.isVisible().catch(() => false)) {
            console.log('Developer picker detected, clicking first developer item...');
            await devItem.click();
            await delay(page, 10000);
        } else {
            console.log('No picker detected or already redirected, moving to app list...');
            await page.goto(appListUrl, { timeout: 60000, waitUntil: 'domcontentloaded' });
            await delay(page, 5000);
        }

        console.log('Clicking "Create app" button on list page...');
        const createBtn = page.locator('[debug-id="create-app-button"], a:has-text("Create app"), button:has-text("Create app")').first();
        await retryAction(async () => {
            await createBtn.click({ timeout: 10000 });
        }, 'Click Create App button');
        await delay(page, 10000);

        // 填写应用名称和包名
        console.log('Filling App name...');
        const appNameInput = page.locator(
            '[debug-id="app-name-input"] input, input[aria-label="App name"]'
        ).first();
        await retryAction(async () => {
            await appNameInput.waitFor({ state: 'visible', timeout: 15000 });
            await appNameInput.fill(appName);
        }, 'Fill App name');
        await delay(page, 1500);

        console.log('Filling App package name...');
        const packageNameInput = page.locator(
            '[debug-id="app-package-name-input"] input, input[aria-label="App package name"]'
        ).first();
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

        // 随机选择 App 或 Game
        const isApp = Math.random() < 0.5;
        const typeLabel = isApp ? 'App' : 'Game';
        const typeId = isApp ? 'app-radio' : 'game-radio';
        console.log(`Selecting type: ${typeLabel}`);
        await retryAction(async () => {
            await page.locator(`[debug-id="${typeId}"]`).first().click({ timeout: 5000 });
        }, `Select type: ${typeLabel}`);
        await delay(page, 2000);

        // 选择 Free
        console.log('Selecting mode: Free');
        await retryAction(async () => {
            await page.getByText('Free', { exact: true }).first().click({ timeout: 5000 });
        }, 'Select Free mode');
        await delay(page, 2000);

        // 勾选声明项
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

        // 点击 Create
        console.log('Clicking "Create" submit button...');
        const submitBtn = page.locator('material-button[debug-id="create-app-button"] button').first();
        await retryAction(async () => {
            await submitBtn.scrollIntoViewIfNeeded({ timeout: 5000 });
            await submitBtn.click({ timeout: 10000 });
        }, 'Final Create button');
        await page.waitForURL(/app-dashboard|app\/setup|apps\/publish/, { timeout: 90000 });
        console.log('App created successfully, navigating to dashboard...');
        await delay(page, 8000);

        // 提取应用基础路径
        const currentUrl = page.url();
        const baseMatch = currentUrl.match(/(.*\/app\/\d+)/);
        if (!baseMatch) {
            throw new Error('Could not extract app base path from URL: ' + currentUrl);
        }
        const appBasePath = baseMatch[1];
        console.log('App path:', appBasePath);

        // --- 闭包内辅助函数 ---
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
            const button = page.locator(`button[aria-label="Start ${sectionTitle} declaration"]`);
            await retryAction(async () => {
                await button.scrollIntoViewIfNeeded({ timeout: 5000 });
                await button.click({ timeout: 10000 });
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

        // --- 执行声明流程 ---
        await goToAppContent();

        // 1. 广告
        console.log('Executing declaration 1/7: Ads...');
        await clickStartDeclaration('Ads');
        await selectRadio(/^No/);
        await clickMainButton('Save');
        await waitSaved(page);
        await goToAppContent();

        // 2. 应用访问权限
        console.log('Executing declaration 2/7: App access...');
        await clickStartDeclaration('App access');
        await selectRadio('All functionality in my app is available without any access restrictions');
        await clickMainButton('Save');
        await waitSaved(page);
        await goToAppContent();

        // 3. 目标受众和内容
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
        await goToAppContent();

        // 4. 广告 ID
        console.log('Executing declaration 4/7: Advertising ID...');
        await clickStartDeclaration('Advertising ID');
        await selectRadio(/^Yes/);
        await selectCheckbox('Developer communications');
        await clickMainButton('Save');
        await waitSaved(page);
        await page.evaluate(() => window.scrollTo(0, 0));
        await delay(page, 2000);
        await goToAppContent();

        // 5. 政府应用
        console.log('Executing declaration 5/7: Government apps...');
        await clickStartDeclaration('Government apps');
        await selectRadio(/^No/);
        await clickMainButton('Save');
        await waitSaved(page);
        await goToAppContent();

        // 6. 金融功能
        console.log('Executing declaration 6/7: Financial features...');
        await clickStartDeclaration('Financial features');
        await selectCheckbox("My app doesn't provide any financial features");
        await clickMainButton('Next');
        await clickMainButton('Save');
        await waitSaved(page);
        await goToAppContent();

        // 7. 健康应用
        console.log('Executing declaration 7/7: Health apps...');
        await clickStartDeclaration('Health apps');
        await selectCheckbox('My app does not have any health features');
        await clickMainButton('Next');
        await clickMainButton('Save');
        await waitSaved(page);
        await goToAppContent();

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

        // 12. Select Country / region header checkbox and ensure checked
        console.log('Selecting "Country / region" checkbox...');
        const countryRegionHeaderText = page.locator('text=Country / region').first();
        const countryRegionHeaderRow = page.locator(
            'tr:has-text("Country / region"), [role="row"]:has-text("Country / region"), div:has-text("Country / region")'
        ).first();
        const countryRegionCheckboxInput = countryRegionHeaderRow.locator('input[type="checkbox"]').first();
        const countryRegionCheckboxTarget = countryRegionHeaderRow.locator('[role="checkbox"], .mdc-checkbox').first();

        await retryAction(async () => {
            await countryRegionHeaderText.waitFor({ state: 'visible', timeout: 15000 });
            await countryRegionHeaderRow.scrollIntoViewIfNeeded({ timeout: 5000 }).catch(() => { });

            const checkedBefore = await countryRegionCheckboxInput.isChecked().catch(() => false);
            if (!checkedBefore) {
                await countryRegionCheckboxInput.check({ force: true, timeout: 10000 }).catch(async () => {
                    if (await countryRegionCheckboxTarget.isVisible().catch(() => false)) {
                        await countryRegionCheckboxTarget.click({ timeout: 10000 });
                    } else {
                        await countryRegionHeaderRow.click({ timeout: 10000 });
                    }
                });
            }

            const checkedAfter = await countryRegionCheckboxInput.isChecked().catch(() => false);
            const ariaCheckedAfter = await countryRegionCheckboxTarget.getAttribute('aria-checked').catch(() => null);
            if (!checkedAfter && ariaCheckedAfter !== 'true') {
                throw new Error('Country / region checkbox is still not checked.');
            }
        }, 'Ensure Country / region checkbox checked');
        await delay(page, 2000);

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

        await page.click('text=All apps').catch(() => { });

        const durationSeconds = Math.round((Date.now() - startTime) / 1000);
        console.log(`Finished: ${appName} (${packageName}), Duration: ${formatDuration(durationSeconds)}`);
        await browser.close();
    } catch (err) {
        console.error(`[FATAL ERROR] In iteration: ${err.message}`);
        if (browser) await browser.close().catch(() => { });
        throw err;
    }
}

(async () => {
    const { inputFileArg } = parseCliArgs();
    const { appListUrl, configPath } = loadDeveloperConsoleAppListUrl();
    const inputFilePath = resolveInputExcelFile(inputFileArg);
    const { tasks, sheetName } = loadTasksFromExcel(inputFilePath);
    const selectedTasks = tasks;

    console.log(`Developer console URL loaded from: ${configPath}`);
    console.log(`Loaded ${tasks.length} rows from Excel: ${path.basename(inputFilePath)} (sheet: ${sheetName})`);
    console.log(`Will execute ${selectedTasks.length} iteration(s), matching Excel valid rows.`);

    for (let i = 0; i < selectedTasks.length; i++) {
        const task = selectedTasks[i];
        console.log(
            `Iteration ${i + 1}/${selectedTasks.length} | Excel row ${task.rowNumber} | ` +
            `App="${task.appName}" | Package="${task.packageName}"`
        );
        try {
            await runOnce(task, appListUrl);
        } catch (e) {
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
})().catch(err => {
    console.error(`[INIT ERROR] ${err.message}`);
    process.exit(1);
});
