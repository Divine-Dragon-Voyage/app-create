const test = require('node:test');
const assert = require('node:assert/strict');
const {
    DATA_SAFETY_COLLECTION_SECURITY_ACTIONS,
    DATA_SAFETY_COLLECTION_SECURITY_STEP_SELECTOR,
    DATA_SAFETY_DATA_TYPES_ACTIONS,
    DATA_SAFETY_DEVICE_IDS_CHECKBOX_TEXT,
    DATA_SAFETY_DEVICE_IDS_SYNC_WAIT_MS,
    DATA_SAFETY_USAGE_ACTIONS,
    DATA_SAFETY_ACCOUNT_CREATION_CHECKBOX_SELECTORS,
    DATA_SAFETY_PRIMARY_NEXT_BUTTON_SELECTOR,
    DATA_SAFETY_NEXT_BUTTON_SELECTORS,
    DATA_SAFETY_OUTSIDE_APP_LOGIN_GROUP_SELECTORS,
    DATA_SAFETY_OUTSIDE_APP_LOGIN_NO_RADIO_INDEX,
    DATA_SAFETY_SECTION_SELECTORS,
    DATA_SAFETY_DATA_DELETION_GROUP_SELECTORS,
    DATA_SAFETY_DATA_DELETION_NO_RADIO_INDEX,
    DATA_SAFETY_DIRECT_MATERIAL_RADIO_SELECTOR,
    DATA_SAFETY_OUTSIDE_APP_LOGIN_NO_ANSWER_TEXT,
    DATA_SAFETY_DATA_DELETION_NO_ANSWER_TEXT,
    pickLastVisibleDataSafetyNextButton
} = require('../data_safety_flow');

test('data collection and security answers yes yes no-account no no', () => {
    assert.equal(DATA_SAFETY_COLLECTION_SECURITY_STEP_SELECTOR, 'material-radio-group[debug-id="personal-data-yes"]');
    assert.deepEqual(
        DATA_SAFETY_COLLECTION_SECURITY_ACTIONS.map(action => ({
            type: action.type,
            answer: action.answer.source
        })),
        [
            { type: 'radio', answer: '^Yes$' },
            { type: 'radio', answer: '^Yes$' },
            { type: 'checkbox', answer: 'My app does not allow users to create an account' },
            { type: 'radio', answer: '^No$' },
            { type: 'radio', answer: '^No$' }
        ]
    );

    assert.match(
        'Does your app collect or share any of the required user data types?',
        DATA_SAFETY_COLLECTION_SECURITY_ACTIONS[0].question
    );
    assert.match(
        'Is all of the user data collected by your app encrypted in transit?',
        DATA_SAFETY_COLLECTION_SECURITY_ACTIONS[1].question
    );
    assert.match(
        'Can users login to your app with accounts created outside of the app?',
        DATA_SAFETY_COLLECTION_SECURITY_ACTIONS[3].question
    );
    assert.match(
        'Can users log in to your app with accounts created outside of the app?',
        DATA_SAFETY_COLLECTION_SECURITY_ACTIONS[3].question
    );
    assert.match(
        'Do you provide a way for users to request that their data is deleted?',
        DATA_SAFETY_COLLECTION_SECURITY_ACTIONS[4].question
    );
});

test('data types selects only device or other ids', () => {
    assert.equal(DATA_SAFETY_DEVICE_IDS_CHECKBOX_TEXT, 'Device or other IDs');
    assert.equal(DATA_SAFETY_DEVICE_IDS_SYNC_WAIT_MS, 5000);
    assert.deepEqual(
        DATA_SAFETY_DATA_TYPES_ACTIONS.map(action => ({
            type: action.type,
            section: action.section.source,
            answer: action.answer.source
        })),
        [
            {
                type: 'data-type',
                section: 'Device or other IDs',
                answer: 'Device or other IDs'
            }
        ]
    );
});

test('data safety next button prefers stable bottom action selectors', () => {
    assert.equal(DATA_SAFETY_PRIMARY_NEXT_BUTTON_SELECTOR, 'button[debug-id="button-next"]');
    assert.ok(DATA_SAFETY_NEXT_BUTTON_SELECTORS.includes('button[debug-id="button-next"]'));
    assert.ok(DATA_SAFETY_NEXT_BUTTON_SELECTORS.includes('button[debug-id="next-button"]'));
    assert.ok(DATA_SAFETY_NEXT_BUTTON_SELECTORS.includes('button[debug-id="main-button"]:has-text("Next")'));
});

test('data safety next picker skips stale hidden candidates but keeps disabled visible next', () => {
    assert.equal(
        pickLastVisibleDataSafetyNextButton([
            { text: 'Next', visible: true, disabled: false },
            { text: 'Next', visible: false, disabled: false },
            { text: 'Next', visible: true, disabled: true }
        ]),
        2
    );
    assert.equal(
        pickLastVisibleDataSafetyNextButton([
            { text: 'Next', visible: true, disabled: true },
            { text: 'Save', visible: true, disabled: false },
            { text: 'Next', visible: true, disabled: false }
        ]),
        2
    );
});

test('device or other ids usage answers collected ephemeral required app functionality then save', () => {
    const collectedAction = DATA_SAFETY_USAGE_ACTIONS.find(action => action.label === 'Collected');
    const ephemeralAction = DATA_SAFETY_USAGE_ACTIONS.find(action => action.label === 'processed ephemerally');
    const requiredAction = DATA_SAFETY_USAGE_ACTIONS.find(action => action.label === 'data collection required');

    assert.equal(collectedAction.selector, 'material-checkbox[debug-id="collected-checkbox"]');
    assert.equal(collectedAction.revealsSelector, 'material-radio-group[debug-id="ephemerality-question"]');
    assert.equal(ephemeralAction.groupSelector, 'material-radio-group[debug-id="ephemerality-question"]');
    assert.equal(requiredAction.groupSelector, 'material-radio-group[debug-id="user-control-question"]');
    assert.deepEqual(
        DATA_SAFETY_USAGE_ACTIONS.map(action => ({
            type: action.type,
            label: action.label
        })),
        [
            { type: 'start', label: 'Device or other IDs questions' },
            { type: 'checkbox', label: 'Collected' },
            { type: 'radio', label: 'processed ephemerally' },
            { type: 'radio', label: 'data collection required' },
            { type: 'checkbox', label: 'App functionality' },
            { type: 'save', label: 'Save' }
        ]
    );
});

test('data collection account creation controls use stable debug selectors', () => {
    assert.ok(DATA_SAFETY_ACCOUNT_CREATION_CHECKBOX_SELECTORS.includes('material-checkbox[debug-id="acm-checkboxes"]'));
    assert.ok(DATA_SAFETY_OUTSIDE_APP_LOGIN_GROUP_SELECTORS.includes('material-radio-group[debug-id="has-outside-app-accounts"]'));
    assert.ok(DATA_SAFETY_OUTSIDE_APP_LOGIN_GROUP_SELECTORS.includes('[role="radiogroup"][aria-label*="created outside of the app"]'));
});

test('outside app login answer uses direct no radio under stable group', () => {
    const outsideLoginAction = DATA_SAFETY_COLLECTION_SECURITY_ACTIONS.find(action => action.label === 'outside app login = No');
    assert.ok(outsideLoginAction);
    assert.equal(outsideLoginAction.type, 'radio');
    assert.match('Can users log in to your app with accounts created outside of the app?', outsideLoginAction.question);
    assert.match('No', outsideLoginAction.answer);
    assert.equal(DATA_SAFETY_OUTSIDE_APP_LOGIN_NO_RADIO_INDEX, 1);
    assert.equal(DATA_SAFETY_DIRECT_MATERIAL_RADIO_SELECTOR, 'material-radio');
    assert.equal(DATA_SAFETY_OUTSIDE_APP_LOGIN_NO_ANSWER_TEXT, 'No');
});

test('data deletion request controls use stable debug selectors', () => {
    assert.ok(DATA_SAFETY_SECTION_SELECTORS.includes('expandable-container[debug-id="data-safety-section"]'));
    assert.ok(DATA_SAFETY_DATA_DELETION_GROUP_SELECTORS.includes('material-radio-group[debug-id="data-deletion"]'));
    assert.ok(DATA_SAFETY_DATA_DELETION_GROUP_SELECTORS.includes('[role="radiogroup"][aria-label*="request that their data is deleted"]'));
});

test('data deletion request answer uses no under stable data deletion group', () => {
    const deletionAction = DATA_SAFETY_COLLECTION_SECURITY_ACTIONS.find(action => action.label === 'data deletion request = No');
    assert.ok(deletionAction);
    assert.equal(deletionAction.type, 'radio');
    assert.match('Do you provide a way for users to request that their data is deleted?', deletionAction.question);
    assert.match('No', deletionAction.answer);
    assert.equal(DATA_SAFETY_DATA_DELETION_NO_RADIO_INDEX, 1);
    assert.equal(DATA_SAFETY_DIRECT_MATERIAL_RADIO_SELECTOR, 'material-radio');
    assert.equal(DATA_SAFETY_DATA_DELETION_NO_ANSWER_TEXT, 'No');
});
