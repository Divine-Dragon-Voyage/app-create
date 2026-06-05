const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const assert = require('node:assert/strict');

const script = fs.readFileSync(path.join(__dirname, '..', 'create_app.js'), 'utf8');

test('next iteration wait is randomized between 15 and 25 seconds', () => {
    assert.match(script, /const wait = 15000 \+ Math\.random\(\) \* 10000;/);
    assert.match(script, /Wait until next iteration:/);
    assert.doesNotMatch(script, /const wait = 30000 \+ Math\.random\(\) \* 40000;/);
});
