/**
 * PVCS CRC — Extra-Shift Request Form (November – December 2026)
 * ------------------------------------------------------------
 * For the casual PA / moonlighting resident group (Olivia Coneys,
 * Kylee Barnabe, et al.) — same process as the Sep–Oct round in
 * resident-shift-form.gs.
 *
 * ONE-TIME SETUP: paste this file into a new Apps Script project at
 * script.google.com, select createExtraShiftForm, click Run.
 * The execution log prints three links: the form EDIT link, the LIVE link
 * to send out, and the responses spreadsheet.
 *
 * Shifts offered:
 * - EVENINGS: every evening Nov 1 – Dec 31 2026 (61). Nothing is assigned in
 *   the evening row for this period, so the whole range is open. Built by
 *   buildEvenings_() rather than typed out — run logEvenings() to eyeball it.
 * - WEEKEND DAYS: the 9 Sat/Sun day shifts left with only ONE PA after
 *   Brittany's and Alana's picked-up extras were applied, plus Christmas Day
 *   (Fri Dec 25), which is also single-PA.
 *
 * Assignment: shifts are allocated in seniority order, never exceeding
 * anyone's stated max, and only ever from the shifts they ticked.
 *
 * BEFORE SENDING: turn on Settings > Responses > "Send responders a copy of
 * their response". Deadline is already set to Thu Oct 8, 2026, 5:00 PM.
 */

var EVENING_START = new Date(2026, 10, 1);   // Nov 1 2026 (month is 0-based)
var EVENING_END   = new Date(2026, 11, 31);  // Dec 31 2026

var WEEKEND_DAY_SHIFTS = [
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
];  // 9 single-PA weekend days + Christmas Day

var DAYS = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
var MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
              'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

/** Every evening from EVENING_START to EVENING_END inclusive. */
function buildEvenings_() {
  var out = [];
  var d = new Date(EVENING_START.getTime());
  while (d <= EVENING_END) {
    out.push(DAYS[d.getDay()] + ' ' + MONTHS[d.getMonth()] + ' ' + d.getDate() +
             ' — Evening (1700–2200)');
    d.setDate(d.getDate() + 1);
  }
  return out;
}

/** Sanity check before creating the form: prints the list and the count. */
function logEvenings() {
  var e = buildEvenings_();
  Logger.log(e.join('\n'));
  Logger.log('TOTAL EVENINGS: ' + e.length + ' (expect 61)');
  Logger.log('WEEKEND/HOLIDAY DAYS: ' + WEEKEND_DAY_SHIFTS.length + ' (expect 10)');
}

function createExtraShiftForm() {
  var form = FormApp.create('CRC Extra Shifts — November & December 2026');
  form.setDescription(
    'For PAs and residents wishing to pick up extra CRC shifts (Nov 1 – Dec 31, 2026).\n\n' +
    'DEADLINE: 5:00 PM, Thursday, October 8, 2026\n\n' +
    'How assignment works: after the deadline, shifts are allocated in seniority ' +
    'order, never exceeding your stated maximum. You will only ever be assigned ' +
    'shifts you tick below.\n\n' +
    'You will be emailed a copy of your answers with a link that lets you review ' +
    'and change them any time before the deadline.'
  );
  setResponseEditSettings_(form);   // collect email + allow response editing
  form.setLimitOneResponsePerUser(false);
  // NOTE: "Send responders a copy of their response" has NO Apps Script method.
  // After creating the form, set it by hand once:
  //   Settings tab > Responses > Send responders a copy of their response > Always
  // Also confirm "Collect email addresses" reads "Responder input", not "Verified".

  form.addTextItem().setTitle('Your name').setRequired(true);

  var maxItem = form.addTextItem()
    .setTitle('Maximum number of extra shifts you want this period')
    .setRequired(true);
  maxItem.setValidation(
    FormApp.createTextValidation()
      .requireNumberBetween(0, 40)
      .setHelpText('Enter a whole number from 0 to 40.')
      .build()
  );

  form.addCheckboxItem()
    .setTitle('Evening shifts you are willing to work')
    .setChoiceValues(buildEvenings_())
    .setRequired(false);

  form.addCheckboxItem()
    .setTitle('Weekend day shifts you are willing to work')
    .setChoiceValues(WEEKEND_DAY_SHIFTS)
    .setRequired(false);

  var ss = SpreadsheetApp.create('CRC Extra Shifts — Nov-Dec 2026 (Responses)');
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());

  Logger.log('EDIT link (you): ' + form.getEditUrl());
  Logger.log('LIVE link (send out): ' + form.getPublishedUrl());
  Logger.log('Responses sheet: ' + ss.getUrl());
  Logger.log('NEXT: Settings > Responses > "Send responders a copy of their response" = Always.');
}

/**
 * Collect email + allow response editing, so responders get a copy of their
 * answers with an "Edit response" link and can change them before the deadline.
 * An edit overwrites the same row in the responses sheet rather than adding one.
 */
function setResponseEditSettings_(form) {
  try {
    form.setEmailCollectionType(FormApp.EmailCollectionType.RESPONDER_INPUT);
  } catch (e) {
    form.setCollectEmail(true);   // older projects: may land on "Verified", check in the UI
  }
  form.setAllowResponseEdits(true);
}
