/**
 * ============================================================================
 *  ON-CALL SCHEDULE REQUEST FORM v2 — PSYCHIATRISTS — Aug 31 – Oct 11, 2026
 * ============================================================================
 *
 *  Creates the new-style request form plus a linked responses spreadsheet
 *  that includes a "Roster" tab holding each person's HIDDEN priority
 *  (High = has all computer programs, Low = does not). Respondents never see
 *  priority; you change it by editing the Roster tab. To add people later:
 *  add a row on the Roster tab AND add their name to the form's Name
 *  dropdown (or just re-run this builder before sending out).
 *
 *  Structure:
 *   Page 1  Name (dropdown) + shift preference (Day/Evening/Both)
 *   Page 2  Optimal & maximum FULL WEEKS (day, evening) + max WEEKENDS
 *   Pages   Ranked choices: FULL WEEK, then WEEKEND, then MON–FRI
 *           (each for DAY and EVENING)
 *   End     Note that unfilled days will be offered as individual days.
 * ============================================================================
 */

function createPsychiatristFormV2() {

  // ======================= CONFIG — EDIT EACH CYCLE =======================

  var DATE_RANGE = 'Aug 31 – Oct 11, 2026';

  // name -> 'High' or 'Low'   (High = has all computer programs)
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

  // ----------------------------- PAGE 1 ---------------------------------
  form.addListItem()
    .setTitle('Name')
    .setChoiceValues(Object.keys(ROSTER))
    .setRequired(true);

  form.addMultipleChoiceItem()
    .setTitle('Which shifts would you like to be considered for?')
    .setHelpText(APPROVAL_NOTE)
    .setChoiceValues(['Day', 'Evening', 'Both'])
    .setRequired(true);

  // ----------------------------- PAGE 2 ---------------------------------
  form.addPageBreakItem()
    .setTitle('Number of weeks and weekends requested')
    .setHelpText('Enter whole numbers. Leave Day fields at 0 if you only want Evening, and vice versa. ' +
                 'Mon–Fri blocks count toward your full-week numbers. Full weeks (Mon-Sun) count as ' +
                 'one weekend toward your weekend total.');

  addWholeNumber_(form, 'How many is your OPTIMAL number of Full Weeks of DAY shifts (Monday to Sunday)');
  addWholeNumber_(form, 'How many is your MAXIMUM number of Full Weeks of DAY shifts');
  addWholeNumber_(form, 'How many is your OPTIMAL number of Full Weeks of EVENING shifts (Monday to Sunday)');
  addWholeNumber_(form, 'How many is your MAXIMUM number of Full Weeks of EVENING shifts');
  addWholeNumber_(form, 'What is the MAXIMUM number of WEEKENDS you will work over this period (full weeks Mon-Sun will count as one weekend)');

  // ------------------- RANKING SECTIONS (6 of them) ----------------------
  addRankSection_(form, 'DAY full weeks — rank your preferred weeks',
                  'DAY FULL WEEK', FULL_WEEKS, RANK_HELP);
  addRankSection_(form, 'EVENING full weeks — rank your preferred weeks',
                  'EVENING FULL WEEK', FULL_WEEKS, RANK_HELP);
  addRankSection_(form, 'DAY weekends — rank your preferred weekends',
                  'DAY WEEKEND', WEEKENDS, RANK_HELP);
  addRankSection_(form, 'EVENING weekends — rank your preferred weekends',
                  'EVENING WEEKEND', WEEKENDS, RANK_HELP);
  addRankSection_(form, 'DAY Monday–Friday blocks — rank your preferred blocks',
                  'DAY MON-FRI', WEEKDAYS, RANK_HELP);
  addRankSection_(form, 'EVENING Monday–Friday blocks — rank your preferred blocks',
                  'EVENING MON-FRI', WEEKDAYS, RANK_HELP);

  // ------------------------- CLOSING SECTION -----------------------------
  form.addPageBreakItem()
    .setTitle('What happens next')
    .setHelpText('Any days that remain unfilled after this process will be sent out ' +
                 'with the option to take individual days.');

  form.setConfirmationMessage(
    'Thank you. Any days that remain unfilled after this process will be sent out ' +
    'with the option to take individual days.');

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


/** Short-answer question validated as a whole number >= 0. */
function addWholeNumber_(form, title) {
  var validation = FormApp.createTextValidation()
    .requireNumberGreaterThanOrEqualTo(0)
    .build();
  form.addTextItem().setTitle(title).setRequired(true).setValidation(validation);
}


/** A page break + one dropdown per rank slot ("PREFIX — 1st choice", ...). */
function addRankSection_(form, sectionTitle, prefix, choices, help) {
  form.addPageBreakItem().setTitle(sectionTitle).setHelpText(help);
  for (var i = 1; i <= choices.length; i++) {
    form.addListItem()
      .setTitle(prefix + ' — ' + ordinal_(i) + ' choice')
      .setChoiceValues(choices);
  }
}


/** 1 -> "1st", 2 -> "2nd", etc. */
function ordinal_(n) {
  var s = ['th', 'st', 'nd', 'rd'];
  var v = n % 100;
  return n + (s[(v - 20) % 10] || s[v] || s[0]);
}
