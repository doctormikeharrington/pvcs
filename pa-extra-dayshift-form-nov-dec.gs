/**
 * PVCS / CRC — Open Day-Shift Pickup Form: November–December 2026
 * For THREE PAs only — Lindsey Shumila, Alana Ramnauth, Brittany Devaney.
 *
 * They tick a box beside every open day shift (0800–1800) they are willing to
 * pick up. No ranking. Open days = the days currently short a second PA (0 or 1
 * working) in the Nov–Dec base schedule; Lindsey (8-4) counts as one PA.
 * Each day is then assigned by RANDOM draw among everyone who ticked it.
 *
 * Questions:
 *   1. Your name (dropdown, 3 PAs)
 *   2. Maximum number of extra shifts you want this period (0–40)
 *   3. Open days you are willing to pick up (checkboxes — tick any/all)
 *
 * Run createForm() ONCE — the log prints the edit link, live link and responses
 * sheet. If the base schedule changes, edit OPEN_DAYS + run refreshDays().
 * Email collection + response editing match the other PA forms. Built 2026-09-16.
 */

var FORM_TITLE = 'PVCS PA — Open Day-Shift Pickup — November and December 2026';

var PA_NAMES = [
  'Lindsey Shumila',
  'Alana Ramnauth',
  'Brittany Devaney'
];

// Open day shifts (0 or 1 PA working), Nov 1 – Dec 31 2026, every day Mon–Sun.
// Computed from the base schedule. Edit + run refreshDays() to change.
var OPEN_DAYS = [
  'Sun Nov 1 — Day (0800–1800)',
  'Mon Nov 2 — Day (0800–1800)',
  'Tue Nov 3 — Day (0800–1800)',
  'Wed Nov 4 — Day (0800–1800)',
  'Sat Nov 7 — Day (0800–1800)',
  'Sun Nov 8 — Day (0800–1800)',
  'Tue Nov 10 — Day (0800–1800)',
  'Wed Nov 11 — Day (0800–1800)',
  'Thu Nov 12 — Day (0800–1800)',
  'Fri Nov 13 — Day (0800–1800)',
  'Sat Nov 14 — Day (0800–1800)',
  'Sun Nov 15 — Day (0800–1800)',
  'Sat Nov 21 — Day (0800–1800)',
  'Sat Nov 28 — Day (0800–1800)',
  'Sun Nov 29 — Day (0800–1800)',
  'Sat Dec 5 — Day (0800–1800)',
  'Sun Dec 6 — Day (0800–1800)',
  'Sat Dec 12 — Day (0800–1800)',
  'Sun Dec 13 — Day (0800–1800)',
  'Fri Dec 18 — Day (0800–1800)',
  'Sat Dec 19 — Day (0800–1800)',
  'Sun Dec 20 — Day (0800–1800)',
  'Thu Dec 24 — Day (0800–1800)',
  'Fri Dec 25 — Day (0800–1800)',
  'Sat Dec 26 — Day (0800–1800)',
  'Sun Dec 27 — Day (0800–1800)'
]; // 26 open day-shifts

// Filled in by createForm() (printed to the log). Set before running refreshDays().
var FORM_ID = '';

function createForm() {
  var form = FormApp.create(FORM_TITLE);
  form.setDescription(
    'For Lindsey, Alana and Brittany only. Tick every open day shift (0800–1800) ' +
    'you are willing to pick up — these are the days currently short a second PA. ' +
    'You do not need to rank them; just check every day you would take.\n\n' +
    'Each day is then assigned by random draw among everyone who ticked it.\n\n' +
    'DEADLINE: [SET BEFORE SENDING]\n\n' +
    'You will be emailed a copy of your answers with a link that lets you review ' +
    'and change them any time before the deadline.'
  );
  setResponseEditSettings_(form);        // collect email (responder input) + allow edits
  form.setLimitOneResponsePerUser(false);
  // NOTE: "Send responders a copy of their response" has NO Apps Script method.
  // Set it by hand once: Settings tab > Responses > "Always". While there,
  // confirm "Collect email addresses" reads "Responder input" (not "Verified").

  form.addListItem()
    .setTitle('Your name')
    .setChoiceValues(PA_NAMES)
    .setRequired(true);

  form.addTextItem()
    .setTitle('Maximum number of extra shifts you want this period')
    .setHelpText('Whole number across all of November–December 2026 (enter 0 if none).')
    .setValidation(FormApp.createTextValidation()
      .setHelpText('Enter a whole number from 0 to 40.')
      .requireNumberBetween(0, 40)
      .build())
    .setRequired(true);

  form.addCheckboxItem()
    .setTitle('Open days you are willing to pick up')
    .setHelpText('Tick the box beside every day you would take. Leave a day unticked if you do not want it.')
    .setChoiceValues(OPEN_DAYS)
    .setRequired(true);

  var ss = SpreadsheetApp.create('PA Open Day-Shift Pickup (Responses) — Nov-Dec 2026');
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());

  Logger.log('EDIT link (you):        ' + form.getEditUrl());
  Logger.log('LIVE link (send to 3):  ' + form.getPublishedUrl());
  Logger.log('Responses sheet:        ' + ss.getUrl());
  Logger.log('Form id:                ' + form.getId());
  Logger.log('NEXT: set the deadline, and turn on "Send a copy" in Settings > Responses.');
}

/** Re-apply OPEN_DAYS to the checkbox list. */
function refreshDays() {
  if (!OPEN_DAYS.length) throw new Error('OPEN_DAYS is empty.');
  if (!FORM_ID) throw new Error('Set FORM_ID to the id printed by createForm().');
  var form = FormApp.openById(FORM_ID);
  var items = form.getItems(), checks = 0;
  for (var i = 0; i < items.length; i++) {
    if (items[i].getType() === FormApp.ItemType.CHECKBOX &&
        items[i].getTitle().indexOf('Open days') === 0) {
      items[i].asCheckboxItem().setChoiceValues(OPEN_DAYS); checks++;
    }
  }
  Logger.log('Refreshed ' + checks + ' checkbox list(s) with ' + OPEN_DAYS.length + ' days.');
}

/**
 * Collect email + allow response editing. setEmailCollectionType() gives
 * "Responder input"; older projects fall back to setCollectEmail() ("Verified" —
 * confirm in the UI). Same helper as the other PA forms.
 */
function setResponseEditSettings_(form) {
  try {
    form.setEmailCollectionType(FormApp.EmailCollectionType.RESPONDER_INPUT);
  } catch (e) {
    form.setCollectEmail(true);
  }
  form.setAllowResponseEdits(true);
}
