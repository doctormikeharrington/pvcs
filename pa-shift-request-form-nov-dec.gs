/**
 * PVCS / CRC — PA Additional Shift Request Form: November–December 2026
 *
 * Structure matches the Sep–Oct 2026 form, PLUS the new per-pay-period maximums
 * Mike asked for on 2026-08-14.
 *
 * Questions, in order:
 *   1. Your name (dropdown, 10 PAs)
 *   2. Maximum number of extra shifts you want this period  (overall cap, as before)
 *   3–7. Maximum number of extra shifts in EACH pay period  (new)
 *   8. 1st–9th choice shift (dropdowns)
 *   9. Any other shifts you would be willing to work (checkboxes)
 *
 * SHIFT OPTIONS populated 2026-10-03 — run addShiftOptions() against the live
 * form (FORM_ID below), then setDeadline() to write the deadline into the
 * description.
 *
 * Run createPAForm() ONCE. The log prints the edit link, live link and
 * responses sheet. Then run addShiftOptions() later to populate the choices.
 */

var PA_NAMES = [
  'Kylee Barnabe',
  'Olivia Coneys',
  'Jillian Desautels',
  'Brittany Devaney',
  'Daniel Fillion',
  'Karin Love',
  'Alana Ramnauth',
  'Lindsey Shumila',
  'Dana Skaritko',
  'Jennifer Wilson'
];

var PAY_PERIODS = [
  'October 23 – November 5, 2026',
  'November 6 – November 19, 2026',
  'November 20 – December 3, 2026',
  'December 4 – December 17, 2026',
  'December 18 – December 31, 2026'
];

var RANKS = ['1st', '2nd', '3rd', '4th', '5th', '6th', '7th', '8th', '9th'];

/**
 * The shifts offered this round, built 2026-10-03 from the published Nov–Dec
 * schedule: every evening Nov 1 – Dec 31 (nothing is assigned in the evening
 * row for this period), plus the day shifts left with only ONE PA after
 * Brittany's and Alana's pickups — the 9 single-PA weekend days and Christmas
 * Day. 71 shifts in total.
 */
var DAY_NAMES = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
var MONTH_NAMES = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
                   'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

var OPEN_DAY_SHIFTS = [
  'Sun Nov 1 — Day (0800–1800)',
  'Sat Nov 7 — Day (0800–1800)',
  'Sun Nov 8 — Day (0800–1800)',
  'Sun Nov 15 — Day (0800–1800)',
  'Sat Nov 28 — Day (0800–1800)',
  'Sun Dec 6 — Day (0800–1800)',
  'Sat Dec 19 — Day (0800–1800)',
  'Sun Dec 20 — Day (0800–1800)',
  'Fri Dec 25 — Day (0800–1800)',
  'Sat Dec 26 — Day (0800–1800)'
];

/** Every evening Nov 1 – Dec 31 2026, in schedule order. */
function buildEvenings_() {
  var out = [];
  var d = new Date(2026, 10, 1);          // Nov 1 (month is 0-based)
  var end = new Date(2026, 11, 31);       // Dec 31
  while (d <= end) {
    out.push(DAY_NAMES[d.getDay()] + ' ' + MONTH_NAMES[d.getMonth()] + ' ' +
             d.getDate() + ' — Evening (1730–2230)');
    d.setDate(d.getDate() + 1);
  }
  return out;
}

var SHIFTS = buildEvenings_().concat(OPEN_DAY_SHIFTS);

/** Sanity check before touching the live form. */
function logShifts() {
  Logger.log(SHIFTS.join('\n'));
  Logger.log('TOTAL: ' + SHIFTS.length + ' (expect 71 = 61 evenings + 10 days)');
}

var FORM_TITLE = 'PA PVCS — Additional Shift Request Form — November and December 2026';

function createPAForm() {
  var form = FormApp.create(FORM_TITLE);
  form.setDescription(
    'Please submit your requests for additional shifts on the PVCS service for ' +
    'November 1 – December 31, 2026.\n\n' +
    'DEADLINE: [SET BEFORE SENDING]\n\n' +
    'Rank the shifts you would like in order of preference, and tick any other shifts ' +
    'you would be willing to work. Assignment follows the published process: requests ' +
    'are processed in seniority order, rotating one shift at a time, and no one is ' +
    'given more shifts than the maximums they state below.\n\n' +
    'You will be emailed a copy of your answers with a link that lets you review ' +
    'and change them any time before the deadline.'
  );
  setResponseEditSettings_(form);      // collect email + allow response editing
  form.setLimitOneResponsePerUser(false);
  form.setProgressBar(true);
  // NOTE: "Send responders a copy of their response" has NO Apps Script method.
  // Set it by hand once: form Settings tab > Responses > set it to "Always".
  // While there, confirm "Collect email addresses" reads "Responder input", not
  // "Verified" — Verified forces a Google sign-in that PA work accounts may not have.

  form.addListItem()
    .setTitle('Your name')
    .setChoiceValues(PA_NAMES)
    .setRequired(true);

  addNumberItem_(form,
    'Maximum number of extra shifts you want this period',
    'Across the whole period (November 1 – December 31, 2026).',
    true);

  for (var p = 0; p < PAY_PERIODS.length; p++) {
    addNumberItem_(form,
      'Maximum extra shifts — pay period ' + PAY_PERIODS[p],
      'Your cap for this pay period only. You will never be given more than this, ' +
      'even if your overall maximum is higher.',
      true);
  }

  for (var r = 0; r < RANKS.length; r++) {
    form.addListItem()
      .setTitle(RANKS[r] + ' choice shift')
      .setRequired(false);
  }

  form.addCheckboxItem()
    .setTitle('Any other shifts you would be willing to work')
    .setRequired(false);

  var ss = SpreadsheetApp.create('PA Additional Shift Requests (Responses) — Nov-Dec 2026');
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());

  Logger.log('EDIT link (you): ' + form.getEditUrl());
  Logger.log('LIVE link (send to PAs): ' + form.getPublishedUrl());
  Logger.log('Responses sheet: ' + ss.getUrl());
  Logger.log('Form id: ' + form.getId());
  Logger.log('NEXT: fill in SHIFTS, then run addShiftOptions().');
}

/** Whole-number question with a 0–40 validation, used for every maximum. */
function addNumberItem_(form, title, help, required) {
  var item = form.addTextItem().setTitle(title).setHelpText(help).setRequired(!!required);
  item.setValidation(
    FormApp.createTextValidation()
      .requireNumberBetween(0, 40)
      .setHelpText('Enter a whole number (0 if you do not want any).')
      .build()
  );
  return item;
}

/**
 * Populates the 1st–9th choice dropdowns and the "any other shifts" checkboxes
 * from SHIFTS. Safe to re-run — it overwrites the options each time.
 * Set FORM_ID below (printed by createPAForm) before running.
 */
// Created 2026-08-14. Live link:
//   https://docs.google.com/forms/d/e/1FAIpQLSfurkWtMqRYbTWoc1RICnj7RYw15s2RnOr2m1U0wPwg0d7baw/viewform
// Edit link:
//   https://docs.google.com/forms/d/19vz74w8rK4Qfc-5yTocrd124MHb77ces9dDlqtg28R8/edit
// Responses sheet:
//   https://docs.google.com/spreadsheets/d/1WJS5w0TTWFnNKzd7SA65DnY4uVEZobcKAASw1mJTuOY/edit
var FORM_ID = '19vz74w8rK4Qfc-5yTocrd124MHb77ces9dDlqtg28R8';

/**
 * Applies the review-and-resubmit settings to the form that ALREADY exists
 * (the one created 2026-08-14). Run once, before sending the form out.
 * Afterwards, set "Send responders a copy of their response" = Always by hand:
 *   form Settings tab > Responses. There is no API for that one toggle.
 */
function enableResponseEdits() {
  var form = FormApp.openById(FORM_ID);
  setResponseEditSettings_(form);
  Logger.log('Collect email: ' + form.collectsEmail() +
             ' | Allow edits: ' + form.canEditResponse() +
             ' | NEXT: turn on "Send responders a copy of their response" in Settings > Responses.');
}

/**
 * Collect email + allow response editing. setEmailCollectionType() is the newer
 * method and gives "Responder input"; older projects only have setCollectEmail(),
 * which may land on "Verified" — confirm in the UI.
 * Duplicated from enable-response-edits.gs so this file runs on its own.
 */
function setResponseEditSettings_(form) {
  try {
    form.setEmailCollectionType(FormApp.EmailCollectionType.RESPONDER_INPUT);
  } catch (e) {
    form.setCollectEmail(true);
  }
  form.setAllowResponseEdits(true);
}

/**
 * Writes the deadline into the live form's description. Run once.
 * Everything else in the description is unchanged from createPAForm().
 */
function setDeadline() {
  var form = FormApp.openById(FORM_ID);
  form.setDescription(
    'Please submit your requests for additional shifts on the PVCS service for ' +
    'November 1 – December 31, 2026.\n\n' +
    'DEADLINE: 5:00 PM, Thursday, October 8, 2026\n\n' +
    'Rank the shifts you would like in order of preference, and tick any other shifts ' +
    'you would be willing to work. Assignment follows the published process: requests ' +
    'are processed in seniority order, rotating one shift at a time, and no one is ' +
    'given more shifts than the maximums they state below.\n\n' +
    'You will be emailed a copy of your answers with a link that lets you review ' +
    'and change them any time before the deadline.'
  );
  Logger.log('Deadline set. Live link: ' + form.getPublishedUrl());
}

function addShiftOptions() {
  if (!SHIFTS.length) throw new Error('SHIFTS is empty — add the shift labels first.');
  if (!FORM_ID) throw new Error('Set FORM_ID to the id printed by createPAForm().');
  var form = FormApp.openById(FORM_ID);
  var items = form.getItems();
  var lists = 0, checks = 0;
  for (var i = 0; i < items.length; i++) {
    var t = items[i].getTitle();
    if (items[i].getType() === FormApp.ItemType.LIST && t.indexOf('choice shift') > -1) {
      items[i].asListItem().setChoiceValues(SHIFTS);
      lists++;
    }
    if (items[i].getType() === FormApp.ItemType.CHECKBOX &&
        t.indexOf('Any other shifts') === 0) {
      items[i].asCheckboxItem().setChoiceValues(SHIFTS);
      checks++;
    }
  }
  Logger.log('Populated ' + lists + ' ranked dropdowns and ' + checks +
             ' checkbox list(s) with ' + SHIFTS.length + ' shifts.');
}
