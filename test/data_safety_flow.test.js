const test = require('node:test');
const assert = require('node:assert/strict');
const {
    DATA_SAFETY_COLLECTION_SECURITY_ACTIONS,
    DATA_SAFETY_DATA_TYPES_ACTIONS,
    DATA_SAFETY_USAGE_ACTIONS,
    DATA_SAFETY_ACCOUNT_CREATION_CHECKBOX_SELECTORS,
    DATA_SAFETY_OUTSIDE_APP_LOGIN_GROUP_SELECTORS
} = require('../data_safety_flow');

test('data collection and security answers yes yes no-account no no', () => {
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

test('device or other ids usage answers collected ephemeral required app functionality then save', () => {
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
