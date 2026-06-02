const DATA_SAFETY_COLLECTION_SECURITY_ACTIONS = [
    {
        label: 'required user data types = Yes',
        type: 'radio',
        question: /Does your app collect or share any of the required user data types\?/i,
        answer: /^Yes$/i
    },
    {
        label: 'encrypted in transit = Yes',
        type: 'radio',
        question: /Is all of the user data collected by your app encrypted in transit\?/i,
        answer: /^Yes$/i
    },
    {
        label: 'no in-app account creation',
        type: 'checkbox',
        answer: /My app does not allow users to create an account/i
    },
    {
        label: 'outside app login = No',
        type: 'radio',
        question: /Can users log(?:in| in) to your app with accounts created outside of the app\?/i,
        answer: /^No$/i
    },
    {
        label: 'data deletion request = No',
        type: 'radio',
        question: /Do you provide a way for users to request that their data is deleted\?/i,
        answer: /^No$/i
    }
];

const DATA_SAFETY_COLLECTION_SECURITY_STEP_SELECTOR = 'material-radio-group[debug-id="personal-data-yes"]';

const DATA_SAFETY_DATA_TYPES_ACTIONS = [
    {
        label: 'Device or other IDs data type',
        type: 'data-type',
        section: /Device or other IDs/i,
        answer: /Device or other IDs/i
    }
];

const DATA_SAFETY_PRIMARY_NEXT_BUTTON_SELECTOR = 'button[debug-id="button-next"]';

const DATA_SAFETY_NEXT_BUTTON_SELECTORS = [
    DATA_SAFETY_PRIMARY_NEXT_BUTTON_SELECTOR,
    'button[debug-id="next-button"]',
    'button[debug-id="main-button"]:has-text("Next")',
    'button:has-text("Next")'
];

function pickLastVisibleDataSafetyNextButton(candidates) {
    for (let i = candidates.length - 1; i >= 0; i--) {
        const candidate = candidates[i] || {};
        const text = String(candidate.text || '').replace(/\s+/g, ' ').trim();
        if (text !== 'Next') continue;
        if (!candidate.visible) continue;
        return i;
    }
    return -1;
}

const DATA_SAFETY_USAGE_ACTIONS = [
    {
        label: 'Device or other IDs questions',
        type: 'start',
        target: /Device or other IDs/i
    },
    {
        label: 'Collected',
        type: 'checkbox',
        answer: /^Collected$/i
    },
    {
        label: 'processed ephemerally',
        type: 'radio',
        answer: /Yes,\s*this collected data is processed ephemerally/i
    },
    {
        label: 'data collection required',
        type: 'radio',
        answer: /Data collection is required/i
    },
    {
        label: 'App functionality',
        type: 'checkbox',
        answer: /^App functionality$/i
    },
    {
        label: 'Save',
        type: 'save'
    }
];

const DATA_SAFETY_ACCOUNT_CREATION_CHECKBOX_SELECTORS = [
    'material-checkbox[debug-id="acm-checkboxes"]',
    'material-checkbox:has(label:has-text("My app does not allow users to create an account"))',
    'material-checkbox:has-text("My app does not allow users to create an account")'
];

const DATA_SAFETY_SECTION_SELECTORS = [
    'expandable-container[debug-id="data-safety-section"]',
    '[debug-id="data-safety-section"]'
];

const DATA_SAFETY_OUTSIDE_APP_LOGIN_GROUP_SELECTORS = [
    'material-radio-group[debug-id="has-outside-app-accounts"]',
    'material-radio-group[aria-label*="created outside of the app"]',
    '[role="radiogroup"][aria-label*="created outside of the app"]'
];

const DATA_SAFETY_DATA_DELETION_GROUP_SELECTORS = [
    'material-radio-group[debug-id="data-deletion"]',
    'material-radio-group[aria-label*="request that their data is deleted"]',
    '[role="radiogroup"][aria-label*="request that their data is deleted"]'
];

const DATA_SAFETY_OUTSIDE_APP_LOGIN_NO_RADIO_INDEX = 1;
const DATA_SAFETY_DATA_DELETION_NO_RADIO_INDEX = 1;
const DATA_SAFETY_DIRECT_MATERIAL_RADIO_SELECTOR = 'material-radio';
const DATA_SAFETY_OUTSIDE_APP_LOGIN_NO_ANSWER_TEXT = 'No';
const DATA_SAFETY_DATA_DELETION_NO_ANSWER_TEXT = 'No';

module.exports = {
    DATA_SAFETY_COLLECTION_SECURITY_ACTIONS,
    DATA_SAFETY_COLLECTION_SECURITY_STEP_SELECTOR,
    DATA_SAFETY_DATA_TYPES_ACTIONS,
    DATA_SAFETY_PRIMARY_NEXT_BUTTON_SELECTOR,
    DATA_SAFETY_NEXT_BUTTON_SELECTORS,
    DATA_SAFETY_USAGE_ACTIONS,
    DATA_SAFETY_ACCOUNT_CREATION_CHECKBOX_SELECTORS,
    DATA_SAFETY_SECTION_SELECTORS,
    DATA_SAFETY_OUTSIDE_APP_LOGIN_GROUP_SELECTORS,
    DATA_SAFETY_DATA_DELETION_GROUP_SELECTORS,
    DATA_SAFETY_OUTSIDE_APP_LOGIN_NO_RADIO_INDEX,
    DATA_SAFETY_DATA_DELETION_NO_RADIO_INDEX,
    DATA_SAFETY_DIRECT_MATERIAL_RADIO_SELECTOR,
    DATA_SAFETY_OUTSIDE_APP_LOGIN_NO_ANSWER_TEXT,
    DATA_SAFETY_DATA_DELETION_NO_ANSWER_TEXT,
    pickLastVisibleDataSafetyNextButton
};
