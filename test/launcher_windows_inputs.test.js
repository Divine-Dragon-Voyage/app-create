const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const assert = require('node:assert/strict');

const launcherSource = fs.readFileSync(
    path.join(__dirname, '..', 'user_ops', 'launcher_windows.ps1'),
    'utf8'
);

test('Windows launcher collects AppGenie credentials needed by privacy policy flow', () => {
    assert.match(launcherSource, /\$labelWebUser\.Text\s*=\s*"灵构账号"/);
    assert.match(launcherSource, /\$txtWebUser\s*=\s*New-Object System\.Windows\.Forms\.TextBox/);
    assert.match(launcherSource, /\$labelWebPass\.Text\s*=\s*"灵构密码"/);
    assert.match(launcherSource, /\$txtWebPass\s*=\s*New-Object System\.Windows\.Forms\.TextBox/);
    assert.match(launcherSource, /\$labelContactEmail\.Text\s*=\s*"账号邮箱"/);
    assert.match(launcherSource, /\$txtContactEmail\s*=\s*New-Object System\.Windows\.Forms\.TextBox/);
});

test('Windows launcher passes AppGenie credentials through environment variables', () => {
    assert.match(launcherSource, /\$selectedWebUsername\s*=\s*\$null/);
    assert.match(launcherSource, /\$selectedWebPassword\s*=\s*\$null/);
    assert.match(launcherSource, /\$selectedContactEmail\s*=\s*\$null/);
    assert.match(launcherSource, /\$env:APP_CREATE_WEB_USERNAME\s*=\s*\$selectedWebUsername/);
    assert.match(launcherSource, /\$env:APP_CREATE_WEB_PASSWORD\s*=\s*\$selectedWebPassword/);
    assert.match(launcherSource, /\$env:APP_CREATE_CONTACT_EMAIL\s*=\s*\$selectedContactEmail/);
});
