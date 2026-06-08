const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');

const createAppSource = fs.readFileSync(path.join(__dirname, '..', 'create_app.js'), 'utf8');

test('app content declaration 2 uses Play Console Sign-in details label instead of retired App access', () => {
    assert.match(
        createAppSource,
        /Executing declaration 2\/7: Sign-in details[\s\S]*clickStartDeclaration\('Sign-in details'\)[\s\S]*selectRadio\(\s*\/\^No\/\s*\)/
    );
    assert.doesNotMatch(createAppSource, /clickStartDeclaration\('App access'\)/);
    assert.doesNotMatch(createAppSource, /All functionality in my app is available without any access restrictions/);
});

test('start declaration finder falls back between Sign-in details spelling variants', () => {
    assert.match(
        createAppSource,
        /function getDeclarationTitleVariants\(sectionTitle\)[\s\S]*Sign-in details[\s\S]*Sign in details/
    );
    assert.match(createAppSource, /const sectionTitleVariants = getDeclarationTitleVariants\(sectionTitle\)/);
    assert.match(createAppSource, /Start \$\{title\} declaration/);
    assert.match(createAppSource, /normalize-space\(text\(\)\)="\$\{title\}"/);
});
