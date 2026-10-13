/**
 * PVCS — Resident Moonlighting Shift Request Form (Nov–Dec 2026)
 * ------------------------------------------------------------
 * ONE-TIME SETUP: paste this file into a new Apps Script project at
 * script.google.com, select createResidentForm, click Run.
 * The execution log prints the form EDIT link, the LIVE link to send to
 * residents, and the responses spreadsheet.
 *
 * These are the 13 shifts left after the PA seniority round (Oct 8) and the
 * PA leftover first-come-first-served round (closed 22:00 Mon Oct 12).
 *
 * Process (same as the Sep–Oct 2026 round):
 * - Residents enter name, max number of shifts, and tick every shift they
 *   are willing to work.
 * - Deadline: 12:00 noon, Friday, October 16, 2026. NO auto-close — the
 *   deadline is stated in the email and in the form description only.
 *   To close by hand: form editor > Responses > toggle "Accepting responses" off.
 * - Assignment is generated from the responses sheet with NO seniority:
 *   distribute as evenly as possible across respondents, never exceeding
 *   anyone's stated maximum, only ever assigning ticked shifts, and spacing
 *   each person's shifts out across the period as much as possible.
 */

var EVENING_SHIFTS = [
  'Sun Nov 15 — Evening (1700–2200)',
  'Fri Nov 27 — Evening (1700–2200)',
  'Sun Dec 6 — Evening (1700–2200)',
  'Fri Dec 18 — Evening (1700–2200)',
  'Sat Dec 19 — Evening (1700–2200)',
  'Thu Dec 24 — Evening (1700–2200) (Christmas Eve)'
];

var DAY_SHIFTS = [
  'Sun Nov 1 — Day (0800–1700)',
  'Sat Nov 7 — Day (0800–1700)',
  'Sat Nov 28 — Day (0800–1700)',
  'Sun Dec 6 — Day (0800–1700)',
  'Sat Dec 19 — Day (0800–1700)',
  'Sun Dec 20 — Day (0800–1700)',
  'Fri Dec 25 — Day (0800–1700) (Christmas Day)'
];

function createResidentForm() {
  var form = FormApp.create('Resident Moonlighting Shifts — November & December 2026');
  form.setDescription(
    'For residents wishing to pick up moonlighting shifts at PVCS in the November–December 2026 period.\n\n' +
    'DEADLINE: 12:00 noon, Friday, October 16, 2026.\n\n' +
    'How assignment works: after the deadline, shifts are distributed as evenly as ' +
    'possible among everyone who responds (no seniority), never exceeding your stated ' +
    'maximum, with your shifts spaced out across the period as much as possible. ' +
    'You will only ever be assigned shifts you tick below.\n\n' +
    'You will be emailed a copy of your answers with a link that lets you review ' +
    'and change them any time before the deadline.'
  );
  setResponseEditSettings_(form);
  form.setLimitOneResponsePerUser(false);
  // NOTE: "Send responders a copy of their response" has NO Apps Script method.
  // After creating the form, set it by hand once:
  //   Settings tab > Responses > Send responders a copy of their response > Always
  // Also confirm "Collect email addresses" reads "Responder input", not "Verified".

  form.addTextItem().setTitle('Your name').setRequired(true);

  var maxItem = form.addTextItem()
    .setTitle('Maximum number of moonlighting shifts you want this period')
    .setRequired(true);
  maxItem.setValidation(
    FormApp.createTextValidation()
      .requireNumberBetween(0, 13)
      .setHelpText('Enter a whole number from 0 to 13.')
      .build()
  );

  form.addCheckboxItem()
    .setTitle('Evening shifts you are willing to work')
    .setChoiceValues(EVENING_SHIFTS)
    .setRequired(false);

  form.addCheckboxItem()
    .setTitle('Day shifts you are willing to work')
    .setChoiceValues(DAY_SHIFTS)
    .setRequired(false);

  var ss = SpreadsheetApp.create('Resident Moonlighting Shifts — Nov-Dec 2026 (Responses)');
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());

  Logger.log('EDIT link (you): ' + form.getEditUrl());
  Logger.log('LIVE link (send to residents): ' + form.getPublishedUrl());
  Logger.log('Responses sheet: ' + ss.getUrl());
  Logger.log('NEXT: Settings > Responses > "Send responders a copy of their response" = Always.');
}

function setResponseEditSettings_(form) {
  try {
    form.setEmailCollectionType(FormApp.EmailCollectionType.RESPONDER_INPUT);
  } catch (e) {
    form.setCollectEmail(true);
  }
  form.setAllowResponseEdits(true);
}
