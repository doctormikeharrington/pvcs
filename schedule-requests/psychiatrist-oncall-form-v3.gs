/**
 * ============================================================================
 *  ON-CALL REQUEST FORM v3 — PSYCHIATRISTS — Aug 31 – Oct 11, 2026 (BRANCHED)
 * ============================================================================
 *
 *  Adds shift-based branching: people who pick "Day" only see DAY sections,
 *  "Evening" only see EVENING sections, "Both" see a copy of the DAY block
 *  then the EVENING block. (Google Forms section jumps are static, so the
 *  "Both" path needs its own copy of the DAY sections; the scheduler merges
 *  the duplicate response columns automatically.)
 *
 *  Path map:
 *    S1 Name + shift + max weekends
 *      Day     -> DAY-only block  (counts, full weeks, weekends, Mon–Fri) -> end
 *      Both    -> DAY-both block  (same 4 sections)  -> EVENING block -> end
 *      Evening -> EVENING block   (counts, full weeks, weekends, Mon–Fri) -> end
 *
 *  rebuildPsychiatristFormV3() trashes any older psychiatrist form/sheet for
 *  this date range first, then builds fresh (avoids duplicate piles).
 * ============================================================================
 */

function rebuildPsychiatristFormV3() {
  trashOldV3_();
  createPsychiatristFormV3();
}

function createPsychiatristFormV3() {

  // ======================= CONFIG — EDIT EACH CYCLE =======================

  var DATE_RANGE = 'Aug 31 – Oct 11, 2026';

  var ROSTER = {
    'James Bolton': 'High',
    'Michael Harrington': 'High',
    'Jitender Sareen': 'Low',
    'Christopher Classen': 'Low',
    'Antonio Paletta': 'Low',
    'Chijoke Okoye': 'Low',
    'Eunice Gill': 'Low',
    'Navjot Brainch': 'Low',
    'Jennifer Ruzhynsky': 'Low',
    'Geoffrey Konrad': 'Low'
  };

  var FULL_WEEKS = [
    'Week 1: Mon Aug 31 – Sun Sep 6',
    'Week 2: Mon Sep 7 – Sun Sep 13',
    'Week 3: Mon Sep 14 – Sun Sep 20',
    'Week 4: Mon Sep 21 – Sun Sep 27',
    'Week 5: Mon Sep 28 – Sun Oct 4',
    'Week 6: Mon Oct 5 – Sun Oct 11'
  ];
  var WEEKENDS = [
    'Weekend 1: Sat Sep 5 – Sun Sep 6',
    'Weekend 2: Sat Sep 12 – Sun Sep 13',
    'Weekend 3: Sat Sep 19 – Sun Sep 20',
    'Weekend 4: Sat Sep 26 – Sun Sep 27',
    'Weekend 5: Sat Oct 3 – Sun Oct 4',
    'Weekend 6: Sat Oct 10 – Sun Oct 11'
  ];
  var WEEKDAYS = [
    'Weekdays 1: Mon Aug 31 – Fri Sep 4',
    'Weekdays 2: Mon Sep 7 – Fri Sep 11',
    'Weekdays 3: Mon Sep 14 – Fri Sep 18',
    'Weekdays 4: Mon Sep 21 – Fri Sep 25',
    'Weekdays 5: Mon Sep 28 – Fri Oct 2',
    'Weekdays 6: Mon Oct 5 – Fri Oct 9'
  ];

  // =========================== END CONFIG ================================

  var APPROVAL_NOTE =
    'Day-time scheduling requires prior approval. ' +
    'Email jbolton@hsc.mb.ca or mharrington@hsc.mb.ca to discuss.';

  var RANK_WARNING = 'DO NOT RANK PERIODS YOU ARE UNABLE TO WORK. LEAVE THESE UNSELECTED';
  var RANK_HELP =
    'Fill in starting with your most preferred choice. Use as many choices as ' +
    'you like and leave the rest blank. Do not pick the same period twice.\n\n' +
    RANK_WARNING;

  var form = FormApp.create('On-Call Schedule Request – Psychiatrists – ' + DATE_RANGE);
  form.setDescription(
    'Please submit your on-call availability and preferences for ' + DATE_RANGE + '. ' +
    'Your responses are used to build the on-call schedule.');
  form.setProgressBar(true);

  // --------------------------- SECTION 1 ---------------------------------
  form.addListItem()
    .setTitle('Name')
    .setChoiceValues(Object.keys(ROSTER))
    .setRequired(true);

  var shiftQ = form.addMultipleChoiceItem()
    .setTitle('Which shifts would you like to be considered for?')
    .setHelpText(APPROVAL_NOTE)
    .setRequired(true);

  addWholeNumber_(form,
    'What is the MAXIMUM number of WEEKENDS you will work over this period (full weeks Mon-Sun will count as one weekend)');

  // ------------------- DAY BLOCK (helper, used twice) --------------------
  function addDayBlock_() {
    var pbCounts = form.addPageBreakItem()
      .setTitle('DAY shifts — number of weeks requested')
      .setHelpText('Enter whole numbers. Mon-Fri blocks count toward your full-week numbers.');
    addWholeNumber_(form, 'How many is your OPTIMAL number of Full Weeks of DAY shifts (Monday to Sunday)');
    addWholeNumber_(form, 'How many is your MAXIMUM number of Full Weeks of DAY shifts');
    addRankSection_(form, 'DAY full weeks — rank your preferred weeks', 'DAY FULL WEEK', FULL_WEEKS, RANK_HELP);
    addRankSection_(form, 'DAY weekends — rank your preferred weekends', 'DAY WEEKEND', WEEKENDS, RANK_HELP);
    addRankSection_(form, 'DAY Monday–Friday blocks — rank your preferred blocks', 'DAY MON-FRI', WEEKDAYS, RANK_HELP);
    return pbCounts;
  }

  var pbDayOnly = addDayBlock_();   // path for "Day"
  var pbDayBoth = addDayBlock_();   // path for "Both" (then continues to EVENING)

  // --------------------------- EVENING BLOCK -----------------------------
  var pbEve = form.addPageBreakItem()
    .setTitle('EVENING shifts — number of weeks requested')
    .setHelpText('Enter whole numbers. Mon-Fri blocks count toward your full-week numbers.');
  addWholeNumber_(form, 'How many is your OPTIMAL number of Full Weeks of EVENING shifts (Monday to Sunday)');
  addWholeNumber_(form, 'How many is your MAXIMUM number of Full Weeks of EVENING shifts');
  addRankSection_(form, 'EVENING full weeks — rank your preferred weeks', 'EVENING FULL WEEK', FULL_WEEKS, RANK_HELP);
  addRankSection_(form, 'EVENING weekends — rank your preferred weekends', 'EVENING WEEKEND', WEEKENDS, RANK_HELP);
  addRankSection_(form, 'EVENING Monday–Friday blocks — rank your preferred blocks', 'EVENING MON-FRI', WEEKDAYS, RANK_HELP);

  // --------------------------- FINAL SECTION -----------------------------
  var pbFinal = form.addPageBreakItem()
    .setTitle('What happens next')
    .setHelpText('Any days that remain unfilled after this process will be sent out ' +
                 'with the option to take individual days.');

  form.setConfirmationMessage(
    'Thank you. Any days that remain unfilled after this process will be sent out ' +
    'with the option to take individual days.');

  // ----------------------------- ROUTING ---------------------------------
  // Shift answer decides the first block.
  shiftQ.setChoices([
    shiftQ.createChoice('Day', pbDayOnly),
    shiftQ.createChoice('Evening', pbEve),
    shiftQ.createChoice('Both', pbDayBoth)
  ]);
  // Day-only path: after its last section (the section right before pbDayBoth),
  // jump straight to the final section instead of falling into the Both block.
  pbDayBoth.setGoToPage(pbFinal);
  // Both path falls linearly from its DAY block into the EVENING block (default),
  // and the EVENING block falls into the final section (default).

  // ----------------- RESPONSES SHEET + HIDDEN ROSTER TAB ------------------
  var ss = SpreadsheetApp.create('On-Call Requests (Responses) – Psychiatrists – ' + DATE_RANGE);
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());

  var roster = ss.insertSheet('Roster');
  var rows = [['Name', 'Priority (High/Low)', '', 'High = has all computer programs. Edit freely; the scheduler reads this tab.']];
  Object.keys(ROSTER).forEach(function (n) { rows.push([n, ROSTER[n], '', '']); });
  roster.getRange(1, 1, rows.length, 4).setValues(rows);
  roster.getRange(1, 1, 1, 4).setFontWeight('bold');
  roster.setColumnWidth(1, 180); roster.setColumnWidth(2, 140); roster.setColumnWidth(4, 420);

  Logger.log('FORM (edit):   ' + form.getEditUrl());
  Logger.log('FORM (share):  ' + form.getPublishedUrl());
  Logger.log('RESPONSES:     ' + ss.getUrl());
}


function addWholeNumber_(form, title) {
  var validation = FormApp.createTextValidation()
    .requireNumberGreaterThanOrEqualTo(0)
    .build();
  form.addTextItem().setTitle(title).setRequired(true).setValidation(validation);
}

function addRankSection_(form, sectionTitle, prefix, choices, help) {
  form.addPageBreakItem().setTitle(sectionTitle).setHelpText(help);
  for (var i = 1; i <= choices.length; i++) {
    form.addListItem()
      .setTitle(prefix + ' — ' + ordinal_(i) + ' choice')
      .setChoiceValues(choices);
  }
}

function ordinal_(n) {
  var s = ['th', 'st', 'nd', 'rd'];
  var v = n % 100;
  return n + (s[(v - 20) % 10] || s[v] || s[0]);
}

/** Trash older Aug31–Oct11 psychiatrist form + responses sheet (reversible). */
function trashOldV3_() {
  var queries = [
    "title contains 'On-Call Schedule Request – Psychiatrists – Aug 31' and mimeType = 'application/vnd.google-apps.form'",
    "title contains 'On-Call Requests (Responses) – Psychiatrists – Aug 31' and mimeType = 'application/vnd.google-apps.spreadsheet'"
  ];
  queries.forEach(function (q) {
    var files = DriveApp.searchFiles(q);
    while (files.hasNext()) files.next().setTrashed(true);
  });
}
