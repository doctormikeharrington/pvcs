/**
 * ============================================================================
 *  PHYSICIAN ASSISTANT (CRC) — OPEN-SHIFT REQUESTS, SCHEDULING & SHIFT SWAPS
 * ============================================================================
 *
 *  ONE-TIME INSTALL
 *  ----------------
 *  1. Create a new Google Sheet (suggested name: "PA Scheduling — CRC").
 *  2. Extensions > Apps Script. Delete any code there, paste this whole file,
 *     Save.
 *  3. In the editor, run setupSystem() once and authorize when prompted.
 *     This creates:
 *       - "Roster" tab       (names, emails, seniority — EDIT THIS)
 *       - "PVCS Roster" tab  (the smaller Step 1 group — EDIT THIS;
 *                             seniority is a random placeholder)
 *       - "Open Shifts" tab  (you fill this each cycle)
 *       - "Swap Log" tab     (script-managed)
 *       - the GIVE-AWAY form and the CLAIM form (permanent; links are logged
 *         and stored — see "PA Scheduler > Show form links" menu)
 *       - the form-submit trigger that runs the swap email flow
 *  4. Open the Roster tab: fix the placeholder EMAILS and edit the SENIORITY
 *     numbers (1 = most senior). The seniority here was randomly assigned.
 *
 *  EACH SCHEDULING CYCLE
 *  ---------------------
 *  1. On the "Open Shifts" tab set the period label (cell B1, e.g.
 *     "August 2026") and list the open shifts below: one row per shift,
 *     Date in column A (e.g. "Sat Aug 8"), Morning/Evening in column B.
 *
 *  STEP 1 — PVCS day-shift pre-selection (BEFORE the general round)
 *  2. Menu: PA Scheduler > STEP 1: Build PVCS day-shift form. This builds a
 *     form containing ONLY the Morning (day) shifts, with only the PAs on the
 *     "PVCS Roster" tab in the name pull-down. Responses land in
 *     "PVCS Requests — <period label>".
 *  3. When their responses are in: PA Scheduler > STEP 1: Generate PVCS
 *     day-shift schedule. Written to "PVCS Schedule — <period label>".
 *     Allocation model is identical to the general round (seniority
 *     round-robin, using the seniority on the "PVCS Roster" tab).
 *
 *  STEP 2 — general round (everyone; unchanged from before)
 *  4. Menu: PA Scheduler > STEP 2: Build shift request form. Day shifts
 *     already assigned in Step 1 are automatically left out; everything
 *     else (evening shifts + any day shifts not taken in Step 1) is offered
 *     to the full roster. Responses land in "Requests — <period label>".
 *  5. When responses are in: PA Scheduler > STEP 2: Generate schedule. The
 *     schedule is written to a "Schedule — <period label>" tab.
 *
 *  ALLOCATION RULES (as agreed)
 *  ----------------------------
 *  - People are processed in seniority order (1 first).
 *  - Round-robin: each round, each person (in seniority order) receives their
 *    highest-ranked shift that is still open — at most one per round.
 *  - A person is skipped once they reach the maximum number of extra shifts
 *    they asked for, and is only ever placed in a shift they ranked.
 *  - Rounds repeat until no one can be given anything. Unrequested shifts
 *    stay open ("OPEN — fill manually").
 *  - If someone submits the form twice, only their latest response counts.
 *  - Flags are printed for duplicate picks and for anyone assigned both the
 *    Morning and Evening shift on the same day (allowed, but flagged).
 *
 *  SHIFT GIVE-AWAY / SWAP FLOW
 *  ---------------------------
 *  - A PA who cannot work a scheduled shift fills the GIVE-AWAY form.
 *  - The script emails everyone else on the Roster with the shift details and
 *    a pre-filled CLAIM form link. First submitted claim wins.
 *  - The winner and the requester are emailed; everyone else is told the
 *    shift is taken. Late claims get a "sorry, already taken" email.
 *  - Everything is recorded on the "Swap Log" tab, and the matching row on
 *    the Schedule tab is updated automatically when one is found.
 * ============================================================================
 */


// ============================ CONFIG =======================================

// Roster written by setupSystem(). AFTER SETUP, EDIT THE "Roster" TAB, not
// this list. Seniority below is RANDOM placeholder order (1 = most senior).
var INITIAL_ROSTER = [
  // [Name, Email (placeholder — edit in the Roster tab), Seniority]
  ['Daniel Fillion',   'daniel.fillion@example.com',   1],
  ['Karin Love',       'karin.love@example.com',       2],
  ['Jill Desautels',   'jill.desautels@example.com',   3],
  ['Dana Skaritko',    'dana.skaritko@example.com',    4],
  ['Olivia Coneys',    'olivia.coneys@example.com',    5],
  ['Brittany Devaney', 'brittany.devaney@example.com', 6],
  ['Alana Ramnauth',   'alana.ramnauth@example.com',   7],
  ['Lindsey Shumila',  'lindsey.shumila@example.com',  8]
];

// PVCS roster: the smaller group that pre-selects open DAY shifts (Step 1).
// Written to the "PVCS Roster" tab by setup. AFTER SETUP, EDIT THAT TAB, not
// this list. Seniority below is a RANDOM placeholder (1 = most senior) —
// to be replaced with the real rankings.
var INITIAL_PVCS_ROSTER = [
  ['Alana Ramnauth',   'alana.ramnauth@example.com',   1],
  ['Jill Desautels',   'jill.desautels@example.com',   2],
  ['Lindsey Shumila',  'lindsey.shumila@example.com',  3],
  ['Brittany Devaney', 'brittany.devaney@example.com', 4]
];

var MORNING_LABEL = 'Day (0800–1800)';
var EVENING_LABEL = 'Evening (1730–2230 or 1800–2300)';

// Number of explicit ranked-choice questions on request forms.
var NUM_RANKS = 8;

// Question titles (used to read responses — do not reword in the live forms).
var Q_NAME  = 'Your name';
var Q_MAX   = 'Maximum number of extra shifts you want this period';
var Q_WILLING = 'Any other shifts you would be willing to work';
var Q_GA_DATE  = 'Date of the shift you cannot work (e.g. Sat Aug 8)';
var Q_GA_SHIFT = 'Which shift';
var Q_GA_NOTE  = 'Note (optional)';
var Q_CLAIM_SHIFT = 'Shift you are claiming (do not edit)';

// Tab names
var TAB_ROSTER = 'Roster';
var TAB_PVCS   = 'PVCS Roster';
var TAB_OPEN   = 'Open Shifts';
var TAB_SWAP   = 'Swap Log';
var TAB_GIVEAWAY_RESP = 'Give-Away Responses';
var TAB_CLAIM_RESP    = 'Claim Responses';

var RANK_HELP =
  'Rank up to ' + NUM_RANKS + ' shifts, starting with the one you want most. ' +
  'Use as many of the choices as you like and leave the rest blank. Do not ' +
  'pick the same shift twice. Below the ranked choices you can check any ' +
  'other shifts you would also accept.\n\n' +
  'DO NOT RANK OR CHECK SHIFTS YOU ARE UNABLE TO WORK.';

/** Adds the ranking section: NUM_RANKS dropdowns + a willing-to-work
 *  checkbox list. Used by both request-form builders. */
function addRankingQuestions_(form, shifts) {
  form.addPageBreakItem()
    .setTitle('Rank your top ' + NUM_RANKS + ' shifts')
    .setHelpText(RANK_HELP);
  var n = Math.min(NUM_RANKS, shifts.length);
  for (var i = 1; i <= n; i++) {
    form.addListItem()
      .setTitle(ordinal_(i) + ' choice shift')
      .setChoiceValues(shifts);
  }
  form.addCheckboxItem()
    .setTitle(Q_WILLING)
    .setHelpText('Check every other shift you would accept. These count after ' +
                 'your ranked choices, in date order.')
    .setChoiceValues(shifts);
}

// ========================== END CONFIG =====================================


function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('PA Scheduler')
    .addItem('STEP 1: Build PVCS day-shift form', 'buildPvcsForm')
    .addItem('STEP 1: Generate PVCS day-shift schedule', 'generatePvcsSchedule')
    .addSeparator()
    .addItem('STEP 2: Build shift request form (everyone)', 'buildRequestForm')
    .addItem('STEP 2: Generate schedule', 'generateSchedule')
    .addSeparator()
    .addItem('Show form links', 'showFormLinks')
    .addItem('Run test with sample data', 'runSampleTest')
    .addSeparator()
    .addItem('Setup (run once)', 'setupSystem')
    .addToUi();
}


// ============================ SETUP ========================================

function setupSystem() {
  var ss = SpreadsheetApp.getActive();
  var props = PropertiesService.getDocumentProperties();

  // ---- Roster tab ----
  if (!ss.getSheetByName(TAB_ROSTER)) {
    var r = ss.insertSheet(TAB_ROSTER);
    var rows = [['Name', 'Email', 'Seniority (1 = most senior)']].concat(INITIAL_ROSTER);
    r.getRange(1, 1, rows.length, 3).setValues(rows);
    r.getRange(1, 1, 1, 3).setFontWeight('bold');
    r.setColumnWidth(1, 160); r.setColumnWidth(2, 240); r.setColumnWidth(3, 180);
  }

  // ---- PVCS Roster tab (Step 1 group) ----
  if (!ss.getSheetByName(TAB_PVCS)) {
    var pv = ss.insertSheet(TAB_PVCS);
    var pvRows = [['Name', 'Email', 'Seniority (1 = most senior) — RANDOM placeholder, edit me']]
      .concat(INITIAL_PVCS_ROSTER);
    pv.getRange(1, 1, pvRows.length, 3).setValues(pvRows);
    pv.getRange(1, 1, 1, 3).setFontWeight('bold');
    pv.setColumnWidth(1, 160); pv.setColumnWidth(2, 240); pv.setColumnWidth(3, 320);
  }

  // ---- Open Shifts tab ----
  if (!ss.getSheetByName(TAB_OPEN)) {
    var o = ss.insertSheet(TAB_OPEN);
    o.getRange('A1').setValue('Period label:').setFontWeight('bold');
    o.getRange('B1').setValue('August 2026');
    o.getRange('A2:B2').setValues([['Date (e.g. Sat Aug 8)', 'Shift']]).setFontWeight('bold');
    // Sample: every day in August 2026, Morning and Evening (62 rows).
    // Delete the rows you don't need, or replace the whole list each cycle.
    var sample = [];
    for (var d = 1; d <= 31; d++) {
      var dayLabel = Utilities.formatDate(new Date(2026, 7, d), Session.getScriptTimeZone(), 'EEE MMM d');
      sample.push([dayLabel, 'Morning']);
      sample.push([dayLabel, 'Evening']);
    }
    o.getRange(3, 1, sample.length, 2).setValues(sample);
    var rule = SpreadsheetApp.newDataValidation().requireValueInList(['Morning', 'Evening'], true).build();
    o.getRange('B3:B200').setDataValidation(rule);
    o.setColumnWidth(1, 180); o.setColumnWidth(2, 120);
  }

  // ---- Swap Log tab ----
  if (!ss.getSheetByName(TAB_SWAP)) {
    var s = ss.insertSheet(TAB_SWAP);
    s.getRange(1, 1, 1, 6).setValues([['Requested', 'Requester', 'Shift', 'Status', 'Claimed by', 'Claimed at']])
      .setFontWeight('bold');
    s.setColumnWidths(1, 6, 160);
  }

  // ---- Give-away form (permanent) ----
  if (!props.getProperty('GIVEAWAY_FORM_ID')) {
    var ga = FormApp.create('PA PVCS — I cannot work a scheduled shift');
    ga.setDescription(
      'Use this form when you cannot work a shift you are scheduled for at CRC. ' +
      'Everyone on the team will be emailed and the first person to claim the shift takes it.');
    ga.addListItem().setTitle(Q_NAME).setChoiceValues(rosterNames_()).setRequired(true);
    ga.addTextItem().setTitle(Q_GA_DATE).setRequired(true);
    ga.addListItem().setTitle(Q_GA_SHIFT).setChoiceValues([MORNING_LABEL, EVENING_LABEL]).setRequired(true);
    ga.addTextItem().setTitle(Q_GA_NOTE);
    linkFormToTab_(ss, ga, TAB_GIVEAWAY_RESP);
    props.setProperty('GIVEAWAY_FORM_ID', ga.getId());
  }

  // ---- Claim form (permanent) ----
  if (!props.getProperty('CLAIM_FORM_ID')) {
    var cl = FormApp.create('PA PVCS — Claim an available shift');
    cl.setDescription(
      'First submitted claim gets the shift. You will receive a confirmation email either way.');
    cl.addListItem().setTitle(Q_NAME).setChoiceValues(rosterNames_()).setRequired(true);
    cl.addTextItem().setTitle(Q_CLAIM_SHIFT).setRequired(true);
    linkFormToTab_(ss, cl, TAB_CLAIM_RESP);
    props.setProperty('CLAIM_FORM_ID', cl.getId());
  }

  // ---- Form-submit trigger ----
  var hasTrigger = ScriptApp.getProjectTriggers().some(function (t) {
    return t.getHandlerFunction() === 'onAnyFormSubmit';
  });
  if (!hasTrigger) {
    ScriptApp.newTrigger('onAnyFormSubmit').forSpreadsheet(ss).onFormSubmit().create();
  }

  showFormLinks();
}


/** Links a form's responses to this spreadsheet and names the response tab. */
function linkFormToTab_(ss, form, tabName) {
  var before = ss.getSheets().map(function (s) { return s.getName(); });
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
  SpreadsheetApp.flush();
  var sheets = SpreadsheetApp.openById(ss.getId()).getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (before.indexOf(sheets[i].getName()) === -1) {
      if (ss.getSheetByName(tabName)) {                      // old tab from a rebuild
        ss.getSheetByName(tabName).setName(tabName + ' (old ' + new Date().getTime() + ')');
      }
      sheets[i].setName(tabName);
      return;
    }
  }
}


function showFormLinks() {
  var props = PropertiesService.getDocumentProperties();
  var lines = [];
  ['GIVEAWAY_FORM_ID', 'CLAIM_FORM_ID', 'PVCS_FORM_ID', 'REQUEST_FORM_ID'].forEach(function (key) {
    var id = props.getProperty(key);
    if (!id) return;
    try {                                  // auto-restore if accidentally trashed
      var file = DriveApp.getFileById(id);
      if (file.isTrashed()) file.setTrashed(false);
    } catch (e) {}
    try {
      var f = FormApp.openById(id);
      lines.push(f.getTitle() + '\n  send this link: ' + f.getPublishedUrl() + '\n  edit: ' + f.getEditUrl());
    } catch (e) {}
  });
  var msg = lines.length ? lines.join('\n\n') : 'No forms yet — run Setup first.';
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) {}
}


// ============ STEP 1 — PVCS DAY-SHIFT PRE-SELECTION (each cycle) ===========

/** Builds the Step 1 form: Morning (day) shifts only, PVCS Roster names only.
 *  The process is otherwise identical to the general request form. */
function buildPvcsForm() {
  var ss = SpreadsheetApp.getActive();
  var props = PropertiesService.getDocumentProperties();
  var open = readOpenShifts_(ss);
  var dayShifts = open.shifts.filter(function (s) {
    return s.indexOf(MORNING_LABEL) !== -1;
  });
  if (!dayShifts.length) {
    SpreadsheetApp.getUi().alert('No Morning (day) shifts found on the "' + TAB_OPEN + '" tab.');
    return;
  }

  // Trash any previous PVCS form for this same period label.
  var title = 'PA PVCS — Day-Shift Pre-Selection — ' + open.label;
  var files = DriveApp.searchFiles(
    "title = '" + title.replace(/'/g, "\\'") + "' and mimeType = 'application/vnd.google-apps.form'");
  while (files.hasNext()) files.next().setTrashed(true);

  var form = FormApp.create(title);
  form.setDescription(
    'Open CRC DAY shifts for ' + open.label + ', offered first to PAs whose regular work is PVCS. ' +
    'Rank the shifts you are willing to take. Shifts are assigned by seniority from these rankings. ' +
    'Day shifts not taken here go into the general round with the evening shifts.');
  form.setProgressBar(true);

  form.addListItem().setTitle(Q_NAME).setChoiceValues(rosterNames_()).setRequired(true);

  form.addTextItem()
    .setTitle(Q_MAX)
    .setHelpText('Whole number. You will not be given more shifts than this.')
    .setRequired(true)
    .setValidation(FormApp.createTextValidation().requireNumberGreaterThanOrEqualTo(0).build());

  addRankingQuestions_(form, dayShifts);

  linkFormToTab_(ss, form, 'PVCS Requests — ' + open.label);
  props.setProperty('PVCS_FORM_ID', form.getId());

  var msg = 'PVCS day-shift form created for ' + open.label + ' (' + dayShifts.length + ' day shifts).\n\n' +
            'Send this link (PVCS group only):\n' + form.getPublishedUrl() + '\n\nEdit:\n' + form.getEditUrl();
  Logger.log(msg);
  SpreadsheetApp.getUi().alert(msg);
}


/** Generates the Step 1 schedule with the same allocation model, using the
 *  seniority on the "PVCS Roster" tab. */
function generatePvcsSchedule() {
  var ss = SpreadsheetApp.getActive();
  var open = readOpenShifts_(ss);
  var dayShifts = open.shifts.filter(function (s) {
    return s.indexOf(MORNING_LABEL) !== -1;
  });
  var respSheet = ss.getSheetByName('PVCS Requests — ' + open.label);
  if (!respSheet) {
    SpreadsheetApp.getUi().alert(
      'No "PVCS Requests — ' + open.label + '" tab found.\n' +
      'Check that the period label on the Open Shifts tab matches the form you sent out.');
    return;
  }
  var people = readRequests_(respSheet, pvcsSeniority_());
  if (!people.length) { SpreadsheetApp.getUi().alert('No responses yet.'); return; }
  writeSchedule_(ss, 'PVCS Schedule — ' + open.label, people, dayShifts,
                 open.label + ' (PVCS day shifts)', '(not taken — goes to Step 2)');
  SpreadsheetApp.getUi().alert(
    'Done — see the "PVCS Schedule — ' + open.label + '" tab.\n\n' +
    'When you build the Step 2 request form, day shifts assigned here are left out automatically.');
}


/** Shifts already assigned on the "PVCS Schedule — <label>" tab. */
function pvcsTakenShifts_(ss, label) {
  var taken = {};
  var sh = ss.getSheetByName('PVCS Schedule — ' + label);
  if (!sh) return taken;
  var data = sh.getDataRange().getValues();
  for (var r = 0; r < data.length; r++) {
    var shift = String(data[r][0]).trim();
    var who = String(data[r][1]).trim();
    if (shift.indexOf(' — ') !== -1 && who && who.charAt(0) !== '(') taken[shift] = true;
  }
  return taken;
}


// ===================== SHIFT REQUEST FORM (each cycle) =====================

function buildRequestForm() {
  var ss = SpreadsheetApp.getActive();
  var props = PropertiesService.getDocumentProperties();
  var open = readOpenShifts_(ss);

  // Leave out day shifts already assigned in Step 1 (PVCS pre-selection).
  var taken = pvcsTakenShifts_(ss, open.label);
  open.shifts = open.shifts.filter(function (s) { return !taken[s]; });

  if (!open.shifts.length) {
    SpreadsheetApp.getUi().alert('No shifts found on the "' + TAB_OPEN + '" tab (or all were taken in Step 1).');
    return;
  }

  // Trash any previous request form for this same period label.
  var title = 'PA PVCS — Additional Shift Request Form — ' + open.label;
  var files = DriveApp.searchFiles(
    "title = '" + title.replace(/'/g, "\\'") + "' and mimeType = 'application/vnd.google-apps.form'");
  while (files.hasNext()) files.next().setTrashed(true);

  var form = FormApp.create(title);
  form.setDescription(
    'Open PVCS shifts for ' + open.label + '. Rank the shifts you are willing to take. ' +
    'Shifts are assigned by seniority from these rankings.');
  form.setProgressBar(true);

  form.addListItem().setTitle(Q_NAME).setChoiceValues(rosterNames_()).setRequired(true);

  form.addTextItem()
    .setTitle(Q_MAX)
    .setHelpText('Whole number. You will not be given more shifts than this.')
    .setRequired(true)
    .setValidation(FormApp.createTextValidation().requireNumberGreaterThanOrEqualTo(0).build());

  addRankingQuestions_(form, open.shifts);

  linkFormToTab_(ss, form, 'Requests — ' + open.label);
  props.setProperty('REQUEST_FORM_ID', form.getId());

  var msg = 'Request form created for ' + open.label + ' (' + open.shifts.length + ' shifts).\n\n' +
            'Send this link:\n' + form.getPublishedUrl() + '\n\nEdit:\n' + form.getEditUrl();
  Logger.log(msg);
  SpreadsheetApp.getUi().alert(msg);
}


/** Reads the period label and shift list from the Open Shifts tab. */
function readOpenShifts_(ss) {
  var sh = ss.getSheetByName(TAB_OPEN);
  if (!sh) return { label: '', shifts: [] };
  var label = asText_(sh.getRange('B1').getValue(), 'MMMM yyyy') || 'Untitled period';
  var last = sh.getLastRow();
  var shifts = [];
  if (last >= 3) {
    sh.getRange(3, 1, last - 2, 2).getValues().forEach(function (row) {
      var date = asText_(row[0], 'EEE MMM d'), shift = String(row[1]).trim();
      if (date && shift) shifts.push(shiftLabel_(date, shift));
    });
  }
  return { label: label, shifts: shifts };
}

/** Sheets auto-converts entries like "August 2026" or "Aug 8" into Date
 *  values; render those back as friendly text instead of a full timestamp. */
function asText_(v, dateFormat) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), dateFormat);
  }
  return String(v).trim();
}

function shiftLabel_(dateText, shiftType) {
  var t = /^m/i.test(shiftType) ? MORNING_LABEL : EVENING_LABEL;
  return dateText + ' — ' + t;
}


// ========================= SCHEDULE GENERATION =============================

function generateSchedule() {
  var ss = SpreadsheetApp.getActive();
  var open = readOpenShifts_(ss);

  // Same filter as the Step 2 form: day shifts assigned in Step 1 are out.
  var taken = pvcsTakenShifts_(ss, open.label);
  open.shifts = open.shifts.filter(function (s) { return !taken[s]; });

  var respSheet = ss.getSheetByName('Requests — ' + open.label);
  if (!respSheet) {
    SpreadsheetApp.getUi().alert(
      'No "Requests — ' + open.label + '" tab found.\n' +
      'Check that the period label on the Open Shifts tab matches the form you sent out.');
    return;
  }
  var people = readRequests_(respSheet, rosterSeniority_());
  if (!people.length) { SpreadsheetApp.getUi().alert('No responses yet.'); return; }
  writeSchedule_(ss, 'Schedule — ' + open.label, people, open.shifts, open.label);
  SpreadsheetApp.getUi().alert('Done — see the "Schedule — ' + open.label + '" tab.');
}


/** Reads request-form responses; keeps only each person's LATEST response.
 *  `seniority` is a name -> rank map (main roster or PVCS roster). */
function readRequests_(sheet, seniority) {
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  var headers = data[0].map(function (h) { return String(h).trim(); });

  var cName = headers.indexOf(Q_NAME);
  var cMax = headers.indexOf(Q_MAX);
  var rankCols = [];
  headers.forEach(function (h, i) {
    var m = h.match(/^(\d+)(st|nd|rd|th) choice shift$/);
    if (m) rankCols.push({ idx: i, ord: parseInt(m[1], 10) });
  });
  rankCols.sort(function (a, b) { return a.ord - b.ord; });

  seniority = seniority || rosterSeniority_();
  var latest = {};                      // name -> {row, timestamp}
  for (var r = 1; r < data.length; r++) {
    var name = String(data[r][cName] || '').trim();
    if (!name) continue;
    var ts = data[r][0] instanceof Date ? data[r][0].getTime() : r;
    if (!latest[name] || ts >= latest[name].ts) latest[name] = { row: data[r], ts: ts };
  }

  var cWill = headers.indexOf(Q_WILLING);
  var out = [];
  Object.keys(latest).forEach(function (name) {
    var row = latest[name].row;
    var ranks = [];
    rankCols.forEach(function (c) {
      var v = String(row[c.idx] || '').trim();
      if (v) ranks.push(v);
    });
    // Willing-to-work checkboxes count after the ranked choices, in the order
    // they appear on the form (date order). Shift labels contain no commas.
    if (cWill !== -1) {
      String(row[cWill] || '').split(', ').forEach(function (s) {
        s = s.trim();
        if (s && ranks.indexOf(s) === -1) ranks.push(s);
      });
    }
    var maxN = parseInt(row[cMax], 10);
    out.push({
      name: name,
      seniority: seniority.hasOwnProperty(name) ? seniority[name] : 999,
      max: isNaN(maxN) ? 0 : maxN,
      ranks: ranks
    });
  });
  out.sort(function (a, b) { return a.seniority - b.seniority; });
  return out;
}


/**
 * Seniority round-robin allocation:
 * repeat { for each person in seniority order: give their highest-ranked
 * still-open shift, unless they are at their max } until nothing changes.
 */
function allocate_(people, shifts) {
  var assigned = {};
  shifts.forEach(function (s) { assigned[s] = null; });
  var counts = {};
  people.forEach(function (p) { counts[p.name] = 0; });

  var progress = true;
  while (progress) {
    progress = false;
    people.forEach(function (p) {
      if (counts[p.name] >= p.max) return;
      for (var i = 0; i < p.ranks.length; i++) {
        var s = p.ranks[i];
        if (assigned.hasOwnProperty(s) && assigned[s] === null) {
          assigned[s] = p.name;
          counts[p.name]++;
          progress = true;
          return;                         // one shift per person per round
        }
      }
    });
  }
  return { assigned: assigned, counts: counts };
}


function writeSchedule_(ss, tabName, people, shifts, label, openText) {
  openText = openText || '(OPEN — fill manually)';
  var flags = [];

  // De-duplicate each person's ranking; flag duplicates.
  people.forEach(function (p) {
    var seen = {}, clean = [], dups = [];
    p.ranks.forEach(function (s) { if (seen[s]) dups.push(s); else { seen[s] = true; clean.push(s); } });
    p.ranks = clean;
    if (dups.length) flags.push(p.name + ': duplicate pick(s) — ' + dups.join(', '));
  });

  var res = allocate_(people, shifts);

  // Flag same-day Morning+Evening doubles.
  var byPersonDate = {};
  shifts.forEach(function (s) {
    var who = res.assigned[s];
    if (!who) return;
    var key = who + '|' + s.split(' — ')[0];
    if (byPersonDate[key]) flags.push(who + ': assigned BOTH shifts on ' + s.split(' — ')[0]);
    byPersonDate[key] = true;
  });

  var sh = ss.getSheetByName(tabName) || ss.insertSheet(tabName);
  sh.clear();
  var rows = [];
  rows.push(['CRC PA Schedule — ' + label + ' — generated ' + new Date().toLocaleString()]);
  rows.push([]);
  rows.push(['Shift', 'Assigned to']);
  shifts.forEach(function (s) {
    rows.push([s, res.assigned[s] || openText]);
  });
  rows.push([]);
  rows.push(['Assigned / Requested max (in seniority order)']);
  people.forEach(function (p) {
    rows.push([p.name + '  (seniority ' + p.seniority + ')', res.counts[p.name] + ' / ' + p.max]);
  });
  rows.push([]);
  rows.push(['Flags (please review)']);
  if (!flags.length) flags.push('none');
  flags.forEach(function (f) { rows.push([f]); });

  rows.forEach(function (r) { while (r.length < 2) r.push(''); });
  sh.getRange(1, 1, rows.length, 2).setValues(rows);
  sh.getRange(3, 1, 1, 2).setFontWeight('bold');
  sh.setColumnWidth(1, 300); sh.setColumnWidth(2, 220);
  sh.setFrozenRows(3);
}


// ====================== GIVE-AWAY / CLAIM EMAIL FLOW =======================

/** Installed trigger: routes any form submission by its response tab. */
function onAnyFormSubmit(e) {
  var tab = e.range.getSheet().getName();
  if (tab === TAB_GIVEAWAY_RESP) handleGiveAway_(e);
  else if (tab === TAB_CLAIM_RESP) handleClaim_(e);
  // Request-form submissions need no live handling.
}


function handleGiveAway_(e) {
  var ss = SpreadsheetApp.getActive();
  var v = e.namedValues;
  var requester = first_(v[Q_NAME]);
  var shift = shiftLabel_(first_(v[Q_GA_DATE]), first_(v[Q_GA_SHIFT]));
  var note = first_(v[Q_GA_NOTE]);

  // Log the request as OPEN.
  ss.getSheetByName(TAB_SWAP).appendRow([new Date(), requester, shift, 'OPEN', '', '']);

  // Pre-filled claim link.
  var claimForm = FormApp.openById(PropertiesService.getDocumentProperties().getProperty('CLAIM_FORM_ID'));
  var url = claimForm.getPublishedUrl();
  try {
    var items = claimForm.getItems(FormApp.ItemType.TEXT);
    for (var i = 0; i < items.length; i++) {
      if (items[i].getTitle() === Q_CLAIM_SHIFT) {
        var resp = claimForm.createResponse()
          .withItemResponse(items[i].asTextItem().createResponse(shift));
        url = resp.toPrefilledUrl();
        break;
      }
    }
  } catch (err) {}

  var body =
    requester + ' cannot work the following shift at CRC and is asking for someone to take it:\n\n' +
    '    ' + shift + '\n' +
    (note ? '\nNote from ' + requester + ': ' + note + '\n' : '') +
    '\nFirst person to claim it gets it. Claim here:\n' + url + '\n\n' +
    '(This is an automated message from the PA scheduling system.)';

  emailRoster_('Shift available: ' + shift, body, [requester]);
}


function handleClaim_(e) {
  var ss = SpreadsheetApp.getActive();
  var v = e.namedValues;
  var claimant = first_(v[Q_NAME]);
  var shift = first_(v[Q_CLAIM_SHIFT]);
  var emails = rosterEmails_();

  // Find the oldest still-OPEN swap request for this shift.
  var sw = ss.getSheetByName(TAB_SWAP);
  var data = sw.getDataRange().getValues();
  var rowIdx = -1;
  for (var r = 1; r < data.length; r++) {
    if (String(data[r][2]).trim() === String(shift).trim() && data[r][3] === 'OPEN') { rowIdx = r; break; }
  }

  if (rowIdx === -1) {
    if (emails[claimant]) {
      MailApp.sendEmail(emails[claimant], 'Shift already taken: ' + shift,
        'Sorry — someone claimed this shift before you, or the request is no longer open.\n\n' +
        '(Automated message from the PA scheduling system.)');
    }
    return;
  }

  var requester = String(data[rowIdx][1]).trim();
  sw.getRange(rowIdx + 1, 4, 1, 3).setValues([['CLAIMED', claimant, new Date()]]);

  // Update the schedule tab if we can find the matching row.
  var updated = updateScheduleRow_(ss, shift, requester, claimant);

  if (emails[claimant]) {
    MailApp.sendEmail(emails[claimant], 'Confirmed — the shift is yours: ' + shift,
      'You were first. You are now scheduled for:\n\n    ' + shift + '\n\n(taking it over from ' + requester + ').\n\n' +
      '(Automated message from the PA scheduling system.)');
  }
  if (emails[requester]) {
    MailApp.sendEmail(emails[requester], 'Your shift has been covered: ' + shift,
      claimant + ' has claimed your shift:\n\n    ' + shift + '\n\n' +
      (updated ? 'The schedule tab has been updated.' :
                 'NOTE: no matching row was found on a Schedule tab — please update the master schedule by hand.') +
      '\n\n(Automated message from the PA scheduling system.)');
  }
  emailRoster_('Shift taken: ' + shift,
    'The shift below has been claimed by ' + claimant + '. No further claims are needed.\n\n    ' + shift +
    '\n\n(Automated message from the PA scheduling system.)',
    [claimant, requester]);
}


/** Replaces requester with claimant on any "Schedule — ..." tab. */
function updateScheduleRow_(ss, shift, requester, claimant) {
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getName().indexOf('Schedule — ') !== 0) continue;
    var data = sheets[i].getDataRange().getValues();
    for (var r = 0; r < data.length; r++) {
      if (String(data[r][0]).trim() === shift && String(data[r][1]).trim().indexOf(requester) === 0) {
        sheets[i].getRange(r + 1, 2).setValue(claimant + ' (swap from ' + requester + ')');
        return true;
      }
    }
  }
  return false;
}


// ============================ ROSTER HELPERS ===============================

function rosterRows_() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss ? ss.getSheetByName(TAB_ROSTER) : null;
  if (!sh || sh.getLastRow() < 2) {
    return INITIAL_ROSTER.map(function (r) { return { name: r[0], email: r[1], seniority: r[2] }; });
  }
  return sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues()
    .filter(function (r) { return String(r[0]).trim(); })
    .map(function (r) {
      return { name: String(r[0]).trim(), email: String(r[1]).trim(), seniority: parseInt(r[2], 10) || 999 };
    });
}

function rosterNames_() { return rosterRows_().map(function (r) { return r.name; }); }

/** PVCS roster (Step 1 group) — read from the "PVCS Roster" tab. */
function pvcsRosterRows_() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss ? ss.getSheetByName(TAB_PVCS) : null;
  if (!sh || sh.getLastRow() < 2) {
    return INITIAL_PVCS_ROSTER.map(function (r) { return { name: r[0], email: r[1], seniority: r[2] }; });
  }
  return sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues()
    .filter(function (r) { return String(r[0]).trim(); })
    .map(function (r) {
      return { name: String(r[0]).trim(), email: String(r[1]).trim(), seniority: parseInt(r[2], 10) || 999 };
    });
}

function pvcsRosterNames_() { return pvcsRosterRows_().map(function (r) { return r.name; }); }

function pvcsSeniority_() {
  var out = {};
  pvcsRosterRows_().forEach(function (r) { out[r.name] = r.seniority; });
  return out;
}

function rosterSeniority_() {
  var out = {};
  rosterRows_().forEach(function (r) { out[r.name] = r.seniority; });
  return out;
}

function rosterEmails_() {
  var out = {};
  rosterRows_().forEach(function (r) {
    if (r.email && r.email.indexOf('example.com') === -1) out[r.name] = r.email;
  });
  return out;
}

/** Emails everyone on the roster except the names in `except`. */
function emailRoster_(subject, body, except) {
  var emails = rosterEmails_();
  var skipped = [];
  rosterRows_().forEach(function (r) {
    if (except && except.indexOf(r.name) !== -1) return;
    if (emails[r.name]) MailApp.sendEmail(emails[r.name], subject, body);
    else skipped.push(r.name);
  });
  if (skipped.length) Logger.log('No valid email (placeholder?) — not emailed: ' + skipped.join(', '));
}

function first_(arr) { return arr && arr.length ? String(arr[0]).trim() : ''; }

function ordinal_(n) {
  var s = ['th', 'st', 'nd', 'rd'], v = n % 100;
  return n + (s[(v - 20) % 10] || s[v] || s[0]);
}


// ============================== SAMPLE TEST ================================

/** Writes a sample schedule to "Schedule (TEST)" without touching real data. */
function runSampleTest() {
  var shifts = [];
  ['Sat Aug 1', 'Sun Aug 2', 'Sat Aug 8', 'Sun Aug 9', 'Sat Aug 15'].forEach(function (d) {
    shifts.push(shiftLabel_(d, 'Morning'));
    shifts.push(shiftLabel_(d, 'Evening'));
  });
  var people = [
    { name: 'Daniel Fillion',   seniority: 1, max: 2, ranks: [shifts[0], shifts[2], shifts[4], shifts[6]] },
    { name: 'Karin Love',       seniority: 2, max: 3, ranks: [shifts[0], shifts[1], shifts[2], shifts[3], shifts[5]] },
    { name: 'Jill Desautels',   seniority: 3, max: 1, ranks: [shifts[0], shifts[4]] },
    { name: 'Dana Skaritko',    seniority: 4, max: 4, ranks: shifts.slice() },
    { name: 'Olivia Coneys',    seniority: 5, max: 0, ranks: [shifts[7]] },
    { name: 'Brittany Devaney', seniority: 6, max: 2, ranks: [shifts[8], shifts[9]] }
  ];
  writeSchedule_(SpreadsheetApp.getActive(), 'Schedule (TEST)', people, shifts, 'TEST');
  SpreadsheetApp.getUi().alert('Sample schedule written to the "Schedule (TEST)" tab.');
}


/**
 * ONE-OFF PATCH (July 2026): rename existing "PA CRC" forms to "PA PVCS"
 * and add Kylee Barnabe to the Roster + the live forms' name dropdowns.
 * Safe to run more than once.
 */
function applyPvcsRenameAndAddKylee() {
  var SS_ID = '1pj9-c6eRi4gddNlJv3sWTEUR3Y-rHVNxSrHnokXkaQc';
  var NEW_PA = 'Kylee Barnabe';
  var ss = SpreadsheetApp.openById(SS_ID);

  // --- add Kylee to the Roster tab (placeholder email; edit it there) ---
  var roster = ss.getSheetByName(TAB_ROSTER);
  var data = roster.getDataRange().getValues();
  var names = [], maxSen = 0, present = false;
  for (var r = 1; r < data.length; r++) {
    var n = String(data[r][0]).trim();
    if (!n) continue;
    names.push(n);
    var sen = parseInt(data[r][2], 10);
    if (!isNaN(sen) && sen > maxSen) maxSen = sen;
    if (n === NEW_PA) present = true;
  }
  if (!present) {
    roster.appendRow([NEW_PA, 'kylee.barnabe@example.com', maxSen + 1]);
    names.push(NEW_PA);
  }

  // --- rename all PA CRC forms; refresh name dropdowns with full roster ---
  var renamed = [];
  var files = DriveApp.searchFiles(
    "title contains 'PA CRC' and mimeType = 'application/vnd.google-apps.form'");
  while (files.hasNext()) {
    var f = files.next();
    var oldTitle = f.getName();
    var newTitle = oldTitle
      .replace('PA CRC — PVCS Day-Shift', 'PA PVCS — Day-Shift')
      .replace('PA CRC', 'PA PVCS');
    var form = FormApp.openById(f.getId());
    f.setName(newTitle);
    form.setTitle(newTitle);
    // Update the "Your name" dropdown, except on the PVCS pre-selection form
    // (that one is restricted to the PVCS Roster group).
    if (newTitle.indexOf('Day-Shift Pre-Selection') === -1) {
      var items = form.getItems(FormApp.ItemType.LIST);
      for (var i = 0; i < items.length; i++) {
        if (items[i].getTitle() === Q_NAME) {
          items[i].asListItem().setChoiceValues(names);
        }
      }
    }
    renamed.push(oldTitle + '  ->  ' + newTitle);
  }
  Logger.log(renamed.length ? renamed.join('\n') : 'No PA CRC forms found.');
  Logger.log('Roster now includes: ' + names.join(', '));
}
