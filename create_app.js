const { chromium } = require('playwright');

// Default delay (ms), increase this if VPS is slow
const DELAY = 3000;

function getAppName() {
    const d = new Date();
    return d.getFullYear().toString() +
        String(d.getMonth() + 1).padStart(2, '0') +
        String(d.getDate()).padStart(2, '0') +
        String(d.getHours()).padStart(2, '0') +
        String(d.getMinutes()).padStart(2, '0') +
        String(d.getSeconds()).padStart(2, '0');
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

// Global retry wrapper for UI interactions
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
    // Fuzzy match various "saved" hints (case-insensitive)
    try {
        await page.locator('text=/saved/i').first().waitFor({ state: 'visible', timeout: 30000 });
        console.log('Save detected.');
    } catch (e) {
        console.log('Save confirmation timeout, continuing...');
    }
    await delay(page, 5000); // Give extra time for backend processing after saving
}

async function runOnce() {
    const startTime = Date.now();
    const name = getAppName();
    console.log('Creating:', name);

    let browser;
    try {
        browser = await chromium.connectOverCDP('http://localhost:9222', { timeout: 120000 });
        const context = browser.contexts()[0];
        let page = context.pages().find(p => p.url().includes('play.google.com'));

        if (!page) {
            page = await context.newPage();
        }

        await page.bringToFront();
        page.setDefaultTimeout(60000); // Set global default timeout to 60 seconds

        // Go directly to the app list page, skipping account selection
        const DEV_URL = 'https://play.google.com/console/u/0/developers/5719511147760424406';
        await page.goto(DEV_URL + '/app-list', { timeout: 90000, waitUntil: 'domcontentloaded' });
        await delay(page, 8000);

        console.log('Clicking "Create app" button on list page...');
        const createBtn = page.locator('[debug-id="create-app-button"], a:has-text("Create app"), button:has-text("Create app")').first();
        await retryAction(async () => {
            await createBtn.click({ timeout: 10000 });
        }, 'Click Create App button');
        await delay(page, 10000);

        // Fill App name
        console.log('Filling App name:', name);
        const nameInput = page.getByLabel('App name', { exact: true }).first();
        try {
            await nameInput.waitFor({ state: 'visible', timeout: 10000 });
            await nameInput.fill(name);
        } catch (e) {
            console.log('getByLabel failed, trying backup selector...');
            const backupInput = page.locator('input[debugid^="acx"], input[aria-label*="App name"], input').first();
            await backupInput.fill(name);
        }
        await delay(page, 2000);

        // Select App or Game (random choice)
        const isApp = Math.random() < 0.5;
        const typeLabel = isApp ? 'App' : 'Game';
        const typeId = isApp ? 'app-radio' : 'game-radio';
        console.log(`Selecting type: ${typeLabel}`);
        await retryAction(async () => {
            await page.locator(`[debug-id="${typeId}"]`).first().click({ timeout: 5000 });
        }, `Select type: ${typeLabel}`);
        await delay(page, 2000);

        // Select Free
        console.log('Selecting mode: Free');
        await retryAction(async () => {
            await page.getByText('Free', { exact: true }).first().click({ timeout: 5000 });
        }, 'Select Free mode');
        await delay(page, 2000);

        // Declarations
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

        // Click Create
        console.log('Clicking "Create" submit button...');
        const submitBtn = page.locator('material-button[debug-id="create-app-button"] button').first();
        await retryAction(async () => {
            await submitBtn.scrollIntoViewIfNeeded({ timeout: 5000 });
            await submitBtn.click({ timeout: 10000 });
        }, 'Final Create button');
        await page.waitForURL(/app-dashboard|app\/setup|apps\/publish/, { timeout: 90000 });
        console.log('App created successfully, navigating to dashboard...');
        await delay(page, 8000);

        // Extract app base path
        const currentUrl = page.url();
        const baseMatch = currentUrl.match(/(.*\/app\/\d+)/);
        if (!baseMatch) {
            throw new Error('Could not extract app base path from URL: ' + currentUrl);
        }
        const appBasePath = baseMatch[1];
        console.log('App path:', appBasePath);

        // --- Helper functions in closure ---
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

        // --- Execute Declarations ---
        await goToAppContent();

        // 1. Ads
        console.log('Executing declaration 1/7: Ads...');
        await clickStartDeclaration('Ads');
        await selectRadio(/^No/);
        await clickMainButton('Save');
        await waitSaved(page);
        await goToAppContent();

        // 2. App access
        console.log('Executing declaration 2/7: App access...');
        await clickStartDeclaration('App access');
        await selectRadio('All functionality in my app is available without any access restrictions');
        await clickMainButton('Save');
        await waitSaved(page);
        await goToAppContent();

        // 3. Target audience and content
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

        // 4. Advertising ID
        console.log('Executing declaration 4/7: Advertising ID...');
        await clickStartDeclaration('Advertising ID');
        await selectRadio(/^Yes/);
        await selectCheckbox('Developer communications');
        await clickMainButton('Save');
        await waitSaved(page);
        await page.evaluate(() => window.scrollTo(0, 0));
        await delay(page, 2000);
        await goToAppContent();

        // 5. Government apps
        console.log('Executing declaration 5/7: Government apps...');
        await clickStartDeclaration('Government apps');
        await selectRadio(/^No/);
        await clickMainButton('Save');
        await waitSaved(page);
        await goToAppContent();

        // 6. Financial features
        console.log('Executing declaration 6/7: Financial features...');
        await clickStartDeclaration('Financial features');
        await selectCheckbox("My app doesn't provide any financial features");
        await clickMainButton('Next');
        await clickMainButton('Save');
        await waitSaved(page);
        await goToAppContent();

        // 7. Health apps
        console.log('Executing declaration 7/7: Health apps...');
        await clickStartDeclaration('Health apps');
        await selectCheckbox('My app does not have any health features');
        await clickMainButton('Next');
        await clickMainButton('Save');
        await waitSaved(page);
        await goToAppContent();

        await page.click('text=All apps').catch(() => { });

        const durationSeconds = Math.round((Date.now() - startTime) / 1000);
        console.log(`Finished: ${name}, Duration: ${formatDuration(durationSeconds)}`);
        await browser.close();
    } catch (err) {
        console.error(`[FATAL ERROR] In iteration: ${err.message}`);
        if (browser) await browser.close().catch(() => { });
        throw err;
    }
}

const count = parseInt(process.argv[2] || '1');

(async () => {
    for (let i = 0; i < count; i++) {
        console.log(`Iteration ${i + 1}/${count}`);
        try {
            await runOnce();
        } catch (e) {
            console.error(`Iteration ${i + 1} failed but continuing...`);
        }

        if (i < count - 1) {
            const wait = 60000 + Math.random() * 120000;
            console.log('Wait until next iteration:', formatDuration(Math.round(wait / 1000)));
            await new Promise(r => setTimeout(r, wait));
        }
    }
})();
