const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');

const createAppSource = fs.readFileSync(path.join(__dirname, '..', 'create_app.js'), 'utf8');

test('app content declaration 2 uses Sign in details instead of retired App access', () => {
    assert.match(
        createAppSource,
        /Executing declaration 2\/7: Sign in details[\s\S]*clickStartDeclaration\('Sign in details'\)[\s\S]*selectRadio\(\s*\/\^No\/\s*\)/
    );
    assert.doesNotMatch(createAppSource, /clickStartDeclaration\('App access'\)/);
    assert.doesNotMatch(createAppSource, /All functionality in my app is available without any access restrictions/);
});
