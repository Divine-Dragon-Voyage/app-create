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
        question: /Can users login to your app with accounts created outside of the app\?/i,
        answer: /^No$/i
    },
    {
        label: 'data deletion request = No',
        type: 'radio',
        question: /Do you provide a way for users to request that their data is deleted\?/i,
        answer: /^No$/i
    }
];

const DATA_SAFETY_DATA_TYPES_ACTIONS = [
    {
        label: 'Device or other IDs data type',
        type: 'data-type',
        section: /Device or other IDs/i,
        answer: /Device or other IDs/i
    }
];

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

module.exports = {
    DATA_SAFETY_COLLECTION_SECURITY_ACTIONS,
    DATA_SAFETY_DATA_TYPES_ACTIONS,
    DATA_SAFETY_USAGE_ACTIONS
};
