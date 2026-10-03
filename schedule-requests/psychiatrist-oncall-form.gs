/**
 * ============================================================================
 *  ON-CALL SCHEDULE REQUEST FORM  —  PSYCHIATRISTS  (TEMPLATE)
 * ============================================================================
 *
 *  WHAT THIS DOES
 *  --------------
 *  rebuildPsychiatristForm()  ->  trashes any existing psychiatrist form +
 *                                 response sheet for this date range, then
 *                                 builds a fresh one. Safe to re-run; it will
 *                                 not pile up duplicates. (Creates NEW links.)
 *  createPsychiatristForm()   ->  just builds a new form + linked sheet.
 *  applyEdits()               ->  one-time patch of the CURRENT live form
 *                                 (keeps existing links). See FORM_ID inside.
 *
 *  HOW TO USE EACH CYCLE
 *  ---------------------
 *  1. Edit the CONFIG block: set DATE_RANGE and list the week blocks.
 *  2. Run rebuildPsychiatristForm().
 *  3. The Logs (View > Execution log) print the form edit link, the live
 *     (share) link, and the responses spreadsheet link.
 *
 *  RANKING METHOD
 *  --------------
 *  Google Forms has no drag-to-reorder question. Instead, each shift type has
 *  one dropdown per rank slot — "1st choice week", "2nd choice week", ... —
 *  each listing every week. Because each rank is its own question, a rank can
 *  only be used once, and respondents fill only as many slots as they wish.
 *
 *  NOTE: Bold text in section descriptions can only be applied in the Forms
 *  UI, not via this API. The warning below is added in CAPS by the script;
 *  bolding is done manually in the form editor.
 * ============================================================================
 */

var RANK_WARNING = '⚠️ DO NOT RANK WEEKS YOU ARE UNABLE TO WORK. LEAVE THESE UNSELECTED ⚠️';

// Name dropdown, alphabetical by LAST name. Update each cycle as needed.
var PSYCHIATRIST_NAMES = [
  'Adegoke Adelufosi',
  'Abdellah Bezzahou',
  'James Bolton',
  'Navjot Brainch',
  'Kristen Braun',
  'Christopher Classen',
  'Eunice Gill',
  'Michael Harrington',
  'Geoffrey Konrad',
  'Nina Kuzenko',
  'Chijioke Okoye',
  'Joshua Palay',
  'Antonio Paletta',
  'Samantha Reimer',
  'Jennifer Ruzhynsky',
  'Jitender Sareen',
  'Shauna Sawich',
  'Melina Zylberman'
];

var RANK_HELP =
  'Fill in starting with your most preferred week. Use as many choices as you ' +
  'like and leave the rest blank. Do not pick the same week twice.\n\n' +
  RANK_WARNING;


/**
 * ONE-TIME patch for the form that is already live (does not change links).
 * Set FORM_ID to the current form's id.
 */
function applyEdits() {
  var FORM_ID = '1l2AmzUnvDta_Msu2fnqVDg74NruCATOmtIkg3Dvc0jQ';
  var form = FormApp.openById(FORM_ID);

  // 1) Reword the four week-count questions.
  var renames = {
    'Optimal number of Full Weeks of DAY shifts':
      'How many is the OPTIMAL number of Full Weeks of DAY shifts',
    'Maximum number of Full Weeks of DAY shifts':
      'How many is the MAXIMUM number of Full Weeks of DAY shifts',
    'Optimal number of Full Weeks of EVENING shifts':
      'How many is the OPTIMAL number of Full Weeks of EVENING shifts',
    'Maximum number of Full Weeks of EVENING shifts':
      'How many is the MAXIMUM number of Full Weeks of EVENING shifts'
  };
  form.getItems(FormApp.ItemType.TEXT).forEach(function (item) {
    if (renames[item.getTitle()]) {
      item.setTitle(renames[item.getTitle()]);
    }
  });

  // 2) Append the CAPS warning to the two ranking section descriptions.
  form.getItems(FormApp.ItemType.PAGE_BREAK).forEach(function (item) {
    var pb = item.asPageBreakItem();
    if (pb.getTitle().indexOf('rank your preferred weeks') !== -1) {
      var help = pb.getHelpText() || '';
      if (help.indexOf(RANK_WARNING) === -1) {
        pb.setHelpText(help + '\n\n' + RANK_WARNING);
      }
    }
  });
}


function rebuildPsychiatristForm() {
  trashOldPsychFiles_();
  createPsychiatristForm();
}


function createPsychiatristForm() {

  // ======================= CONFIG — EDIT EACH CYCLE =======================

  // Appears in the form title and the responses sheet name.
  var DATE_RANGE = 'August 31 – November 8, 2026';

  // List the week blocks for this cycle (Monday–Sunday each).
  // Day and Evening use the same list by default; if they differ, edit
  // EVENING_WEEKS separately below.
  var DAY_WEEKS = [
    'Week 1: Mon Aug 31 – Sun Sep 6',
    'Week 2: Mon Sep 7 – Sun Sep 13',
    'Week 3: Mon Sep 14 – Sun Sep 20',
    'Week 4: Mon Sep 21 – Sun Sep 27',
    'Week 5: Mon Sep 28 – Sun Oct 4',
    'Week 6: Mon Oct 5 – Sun Oct 11',
    'Week 7: Mon Oct 12 – Sun Oct 18',
    'Week 8: Mon Oct 19 – Sun Oct 25',
    'Week 9: Mon Oct 26 – Sun Nov 1',
    'Week 10: Mon Nov 2 – Sun Nov 8'
  ];

  var EVENING_WEEKS = DAY_WEEKS;   // change if evening weeks differ

  // =========================== END CONFIG ================================


  var form = FormApp.create(formTitle_(DATE_RANGE));
  form.setDescription(
    'Please submit your on-call availability and preferences for ' + DATE_RANGE + '. ' +
    'Your responses are used to build the on-call schedule.');
  form.setProgressBar(true);
  // form.setCollectEmail(true);  // uncomment to auto-capture sign-in email

  // ----------------------------- PAGE 1 ---------------------------------
  form.addListItem()
    .setTitle('Name')
    .setChoiceValues(PSYCHIATRIST_NAMES)
    .setRequired(true);

  form.addMultipleChoiceItem()
    .setTitle('Which shifts would you like to be considered for?')
    .setChoiceValues(['Day', 'Evening', 'Both'])
    .setRequired(true);

  // ----------------------------- PAGE 2 ---------------------------------
  form.addPageBreakItem()
    .setTitle('Number of weeks requested')
    .setHelpText('Enter whole numbers. Leave Day fields at 0 if you only want Evening, and vice versa.');

  addWholeNumber_(form, 'How many is your OPTIMAL number of Full Weeks of DAY shifts');
  addWholeNumber_(form, 'How many is your MAXIMUM number of Full Weeks of DAY shifts');
  addWholeNumber_(form, 'How many is your OPTIMAL number of Full Weeks of EVENING shifts');
  addWholeNumber_(form, 'How many is your MAXIMUM number of Full Weeks of EVENING shifts');

  // ----------------------------- PAGE 3 ---------------------------------
  form.addPageBreakItem()
    .setTitle('DAY SHIFTS — rank your preferred weeks')
    .setHelpText(RANK_HELP);

  addRankedChoiceDropdowns_(form, 'DAY SHIFTS', DAY_WEEKS);

  // ----------------------------- PAGE 4 ---------------------------------
  form.addPageBreakItem()
    .setTitle('EVENING SHIFTS — rank your preferred weeks')
    .setHelpText(RANK_HELP);

  addRankedChoiceDropdowns_(form, 'EVENING SHIFTS', EVENING_WEEKS);

  // ----------------------- LINK RESPONSES TO SHEET -----------------------
  var ss = SpreadsheetApp.create(sheetTitle_(DATE_RANGE));
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());

  // ------------------------------- LINKS ---------------------------------
  Logger.log('FORM (edit):   ' + form.getEditUrl());
  Logger.log('FORM (share):  ' + form.getPublishedUrl());
  Logger.log('RESPONSES:     ' + ss.getUrl());
}


/** Adds a short-answer question validated as a whole number >= 0. */
function addWholeNumber_(form, title) {
  var validation = FormApp.createTextValidation()
    .requireNumberGreaterThanOrEqualTo(0)
    .build();
  form.addTextItem()
    .setTitle(title)
    .setRequired(true)
    .setValidation(validation);
}


/**
 * Adds one dropdown per rank slot for a shift type:
 *   "DAY — 1st choice week", "DAY — 2nd choice week", ...
 * Each dropdown lists every week. Not required, so people fill only the
 * slots they want. One question per rank => a rank cannot be reused.
 */
function addRankedChoiceDropdowns_(form, shiftLabel, weeks) {
  for (var i = 1; i <= weeks.length; i++) {
    form.addListItem()
      .setTitle(shiftLabel + ' — ' + ordinal_(i) + ' choice week')
      .setChoiceValues(weeks);
  }
}


/** 1 -> "1st", 2 -> "2nd", etc. */
function ordinal_(n) {
  var s = ['th', 'st', 'nd', 'rd'];
  var v = n % 100;
  return n + (s[(v - 20) % 10] || s[v] || s[0]);
}


/** Consistent titles so rebuild can find and clean up old copies. */
function formTitle_(dateRange) {
  return 'On-Call Schedule Request – Psychiatrists – ' + dateRange;
}
function sheetTitle_(dateRange) {
  return 'On-Call Requests (Responses) – Psychiatrists – ' + dateRange;
}


/**
 * Moves any existing psychiatrist forms / response sheets to Trash so a
 * rebuild leaves exactly one clean copy. (Trash is reversible.)
 */
function trashOldPsychFiles_() {
  var queries = [
    "title contains 'On-Call Schedule Request – Psychiatrists' and mimeType = 'application/vnd.google-apps.form'",
    "title contains 'On-Call Requests (Responses) – Psychiatrists' and mimeType = 'application/vnd.google-apps.spreadsheet'"
  ];
  queries.forEach(function (q) {
    var files = DriveApp.searchFiles(q);
    while (files.hasNext()) {
      files.next().setTrashed(true);
    }
  });
}
