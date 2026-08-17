/**
 * PVCS CRC — Resident Extra-Shift Request Form (Sep–Oct 2026)
 * ------------------------------------------------------------
 * ONE-TIME SETUP: paste this file into a new Apps Script project at
 * script.google.com, select createResidentForm, click Run.
 * The execution log prints three links: the form EDIT link, the LIVE link
 * to send to residents, and the responses spreadsheet.
 *
 * Process (agreed 2026-08-12):
 * - Residents enter name, max number of shifts, and tick the shifts they
 *   are willing to work (12 open evenings + 8 open weekend day shifts).
 * - Deadline: midnight tonight, Wed Aug 12. No auto-close: open the form
 *   editor Thursday morning → Responses → toggle "Accepting responses" off.
 * - Schedule generation is done by Claude from the responses sheet
 *   (no seniority): distribute shifts as evenly as possible across
 *   respondents, never exceeding anyone's stated max, and space each
 *   person's shifts out as much as possible.
 */

var EVENING_SHIFTS = [
  'Mon Aug 31 — Evening (1700–2200)',
  'Mon Sep 14 — Evening (1700–2200)',
  'Thu Sep 17 — Evening (1700–2200)',
  'Thu Sep 24 — Evening (1700–2200)',
  'Fri Sep 25 — Evening (1700–2200)',
  'Sat Sep 26 — Evening (1700–2200)',
  'Sat Oct 10 — Evening (1700–2200)',
  'Sun Oct 11 — Evening (1700–2200)',
  'Sat Oct 17 — Evening (1700–2200)',
  'Sun Oct 18 — Evening (1700–2200)',
  'Sat Oct 24 — Evening (1700–2200)',
  'Sat Oct 31 — Evening (1700–2200)'
];

var WEEKEND_DAY_SHIFTS = [
  'Sat Sep 26 — Day (0800–1800)',
  'Sat Oct 3 — Day (0800–1800)',
  'Sun Oct 4 — Day (0800–1800)',
  'Sat Oct 10 — Day (0800–1800)',
  'Sun Oct 11 — Day (0800–1800)',
  'Sat Oct 17 — Day (0800–1800)',
  'Sun Oct 18 — Day (0800–1800)',
  'Sun Oct 25 — Day (0800–1800)'
];

function createResidentForm() {
  var form = FormApp.create('Resident Extra Shifts — September & October 2026');
  form.setDescription(
    'For residents wishing to pick up extra CRC shifts (Aug 31 – Oct 31 period).\n\n' +
    'DEADLINE: midnight tonight, Wednesday, August 12.\n\n' +
    'How assignment works: after the deadline, shifts are distributed as evenly as ' +
    'possible among everyone who responds (no seniority), never exceeding your stated ' +
    'maximum, with your shifts spaced out across the period as much as possible. ' +
    'You will only ever be assigned shifts you tick below.\n\n' +
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
      .requireNumberBetween(0, 20)
      .setHelpText('Enter a whole number from 0 to 20.')
      .build()
  );

  form.addCheckboxItem()
    .setTitle('Evening shifts you are willing to work')
    .setChoiceValues(EVENING_SHIFTS)
    .setRequired(false);

  form.addCheckboxItem()
    .setTitle('Weekend day shifts you are willing to work')
    .setChoiceValues(WEEKEND_DAY_SHIFTS)
    .setRequired(false);

  var ss = SpreadsheetApp.create('Resident Extra Shifts — Sep-Oct 2026 (Responses)');
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());

  Logger.log('EDIT link (you): ' + form.getEditUrl());
  Logger.log('LIVE link (send to residents): ' + form.getPublishedUrl());
  Logger.log('Responses sheet: ' + ss.getUrl());
  Logger.log('NEXT: Settings > Responses > "Send responders a copy of their response" = Always.');
}

/**
 * Collect email + allow response editing, so residents get a copy of their
 * answers with an "Edit response" link and can change them before the deadline.
 * An edit overwrites the same row in the responses sheet rather than adding one.
 * Duplicated from enable-response-edits.gs so this file runs on its own.
 */
function setResponseEditSettings_(form) {
  try {
    form.setEmailCollectionType(FormApp.EmailCollectionType.RESPONDER_INPUT);
  } catch (e) {
    form.setCollectEmail(true);   // older projects: may land on "Verified", check in the UI
  }
  form.setAllowResponseEdits(true);
}
