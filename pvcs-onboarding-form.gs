/**
 * PVCS On-Boarding Form + Admin roster (Apps Script)
 * ==================================================
 * Phase 1 (already run): buildPvcsOnboardingForm() created the Google Form and
 *   the email-notification trigger.
 * Phase 2 (run setupAdminAssets once): creates the admin roster spreadsheet,
 *   a shift-record sheet + notes doc per person, adds a name-key dropdown to the
 *   form, and wires submissions to auto-fill the roster (which the website reads).
 *
 * After running setupAdminAssets, open View -> Logs and send Claude the
 * "ROSTER_SHEET_ID" it prints.
 */

var NOTIFY_EMAIL  = 'doctormikeharrington@gmail.com';
var FORM_TITLE    = 'PVCS On-Boarding Checklist';
var FORM_PROP_KEY = 'PVCS_ONBOARDING_FORM_ID';
var ROSTER_PROP   = 'PVCS_ROSTER_SS_ID';
var FOLDER_NAME   = 'PVCS Admin';
var KEY_QUESTION  = 'Your name as it appears on the PVCS schedule';

var PSYCHIATRISTS = ['Chris','James','Nina','CJ','Jitender','Goke','Antonio','Geoff','Abdellah','Eunice','Nav','Kristen'];
var PAS           = ['Jill','Lindsay','Britney','Dan','Kylee','Dana','Olivia'];

// Admin-table columns -> exact Form question titles (order matters; matches website)
var FIELD_MAP = [
  ['Work email',        'Work email'],
  ['Cell',              '1. Cell number'],
  ['Momentum',          '2. Do you have access to CHR Momentum?'],
  ['eChart',            '3. Do you have access to echart?'],
  ['EPR',               '4. Do you have access to Electronic Patient Record (EPR)?'],
  ['TigerConnect',      '5. Do you have TigerConnect messaging active on your phone?'],
  ['Scribeberry',       '6. Have you signed up for a Scribeberry account? (instructions: pvcsmanitoba.ca/scribeberry.html)'],
  ['Scribeberry email', '7. What email address do you use for your Scribeberry account?']
];

/* ============================ PHASE 2 SETUP ============================ */

function setupAdminAssets() {
  var form = FormApp.openById(PropertiesService.getScriptProperties().getProperty(FORM_PROP_KEY));
  var folder = getOrCreateFolder_(FOLDER_NAME);

  // 1) Master roster spreadsheet (also the form's response destination).
  var ss = getOrCreateRoster_(folder, form);

  // 2) Add the name-key dropdown to the form (once), at the top.
  ensureKeyQuestion_(form);

  // 3) Build the Psychiatrists and PAs tabs with per-person sheets + docs.
  var psyHeaders = ['Name','Shift URL','Notes URL'].concat(FIELD_MAP.map(function (f) { return f[0]; })).concat(['Day shifts','Night shifts']);
  var paHeaders  = ['Name','Shift URL','Notes URL','Day shifts','Evening shifts'];

  buildRosterTab_(ss, folder, 'Psychiatrists', PSYCHIATRISTS, psyHeaders, ['Date','Day / Night','Notes']);
  buildRosterTab_(ss, folder, 'PAs',           PAS,           paHeaders,  ['Date','Day / Evening','Notes']);

  // 4) Make the roster link-viewable so the website can read it (gviz). The
  //    per-person sheets/docs stay private (open them while signed in).
  DriveApp.getFileById(ss.getId()).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  // 5) Refresh the on-submit trigger so it also updates the roster.
  installNotifyTrigger();

  Logger.log('ROSTER_SHEET_ID: ' + ss.getId());
  Logger.log('Roster spreadsheet URL: ' + ss.getUrl());
  MailApp.sendEmail(NOTIFY_EMAIL, 'PVCS admin roster created',
    'Roster spreadsheet ID: ' + ss.getId() + '\nURL: ' + ss.getUrl() +
    '\n\nSend the ROSTER_SHEET_ID to Claude to finish wiring the admin page.');
}

function getOrCreateFolder_(name) {
  var it = DriveApp.getFoldersByName(name);
  return it.hasNext() ? it.next() : DriveApp.createFolder(name);
}

function getOrCreateRoster_(folder, form) {
  var id = PropertiesService.getScriptProperties().getProperty(ROSTER_PROP);
  if (id) { try { return SpreadsheetApp.openById(id); } catch (e) {} }
  var ss = SpreadsheetApp.create('PVCS Admin Roster');
  // move into folder
  var file = DriveApp.getFileById(ss.getId());
  folder.addFile(file); DriveApp.getRootFolder().removeFile(file);
  // bind form responses to this spreadsheet
  try { form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId()); } catch (e) {}
  PropertiesService.getScriptProperties().setProperty(ROSTER_PROP, ss.getId());
  return ss;
}

function ensureKeyQuestion_(form) {
  var exists = form.getItems(FormApp.ItemType.LIST).some(function (it) {
    return it.getTitle() === KEY_QUESTION;
  });
  if (exists) return;
  var item = form.addListItem();
  item.setTitle(KEY_QUESTION)
      .setHelpText('Pick your name so your answers attach to the right row. New team members: choose "Other".')
      .setChoiceValues(PSYCHIATRISTS.concat(PAS).concat(['Other (new team member)']))
      .setRequired(true);
  form.moveItem(item.getIndex(), 0); // put it first
}

function buildRosterTab_(ss, folder, tabName, names, headers, shiftCols) {
  var sh = ss.getSheetByName(tabName) || ss.insertSheet(tabName);
  sh.clear();
  sh.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold').setBackground('#1f3a5f').setFontColor('#ffffff');

  for (var i = 0; i < names.length; i++) {
    var name = names[i];
    var shiftSheet = SpreadsheetApp.create(name + ' - Shift Record');
    var sf = DriveApp.getFileById(shiftSheet.getId()); folder.addFile(sf); DriveApp.getRootFolder().removeFile(sf);
    var s0 = shiftSheet.getSheets()[0];
    s0.getRange(1, 1, 1, shiftCols.length).setValues([shiftCols]).setFontWeight('bold');

    var doc = DocumentApp.create(name + ' - Notes & Archive');
    var df = DriveApp.getFileById(doc.getId()); folder.addFile(df); DriveApp.getRootFolder().removeFile(df);
    doc.getBody().appendParagraph(name + ' - Notes & Archived Emails').setHeading(DocumentApp.ParagraphHeading.HEADING1);
    doc.saveAndClose();

    sh.getRange(i + 2, 1).setValue(name);
    sh.getRange(i + 2, 2).setValue(shiftSheet.getUrl());
    sh.getRange(i + 2, 3).setValue(doc.getUrl());
  }

  // Conditional formatting: green for "Yes", red for "No" across answer columns.
  if (tabName === 'Psychiatrists') {
    var firstAnsCol = 4;                   // column D (after Name, Shift, Notes)
    var lastAnsCol  = 3 + FIELD_MAP.length; // 22 = column V
    var rng = sh.getRange(2, firstAnsCol, names.length, lastAnsCol - firstAnsCol + 1);
    applyYesNoFormatting_(sh, rng);
  }
  sh.setFrozenRows(1);
}

function applyYesNoFormatting_(sh, rng) {
  var rules = sh.getConditionalFormatRules();
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Yes')
    .setBackground('#16a34a').setFontColor('#ffffff').setRanges([rng]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('No')
    .setBackground('#dc2626').setFontColor('#ffffff').setRanges([rng]).build());
  sh.setConditionalFormatRules(rules);
}

/* ====================== FORM SUBMIT -> ROSTER + EMAIL ====================== */

function onPvcsFormSubmit(e) {
  try { updateRosterFromResponse_(e); } catch (err) { Logger.log('roster update failed: ' + err); }

  var responses = e.response.getItemResponses();
  var lines = [];
  for (var i = 0; i < responses.length; i++) {
    var r = responses[i];
    var ans = r.getResponse();
    if (Array.isArray(ans)) { ans = ans.join(', '); }
    lines.push(r.getItem().getTitle() + ':\n  ' + (ans === '' ? '(blank)' : ans));
  }
  MailApp.sendEmail(NOTIFY_EMAIL, 'New PVCS On-Boarding Checklist submission',
    'A new PVCS On-Boarding Checklist was submitted.\n\n' + lines.join('\n\n') + '\n\nSubmitted: ' + new Date());
}

function updateRosterFromResponse_(e) {
  var id = PropertiesService.getScriptProperties().getProperty(ROSTER_PROP);
  if (!id) return;
  var ss = SpreadsheetApp.openById(id);

  // Build a title->answer map from this submission.
  var ans = {};
  e.response.getItemResponses().forEach(function (r) {
    var v = r.getResponse();
    if (Array.isArray(v)) v = v.join(', ');
    ans[r.getItem().getTitle()] = v;
  });

  var name = String(ans['First Name'] || '').trim();
  if (!name) return;

  // Route by the form's Position answer. Only Psychiatrists get roster rows;
  // PAs/Residents/Clinical Assistants/Other are left for manual handling.
  if (!/psychiatrist/i.test(String(ans['Position'] || ''))) return;

  var sh = ss.getSheetByName('Psychiatrists');
  var last = sh.getLastRow();
  var row = -1;
  if (last >= 2) {
    var names = sh.getRange(2, 1, last - 1, 1).getValues()
      .map(function (r) { return String(r[0]).trim().toLowerCase(); });
    row = names.indexOf(name.toLowerCase());
  }

  if (row < 0) {
    // New person: append a row, reusing their Shift Record / Notes files if
    // they already exist in the PVCS Admin folder (avoids duplicates), else
    // creating fresh ones.
    var folder = getOrCreateFolder_(FOLDER_NAME);
    var links = getOrCreatePersonFiles_(folder, name, ['Date', 'Day / Night', 'Notes']);
    row = (sh.getLastRow() >= 2 ? sh.getLastRow() : 1) - 1; // 0-based data index of new row
    sh.getRange(row + 2, 1, 1, 3).setValues([[name, links[0], links[1]]]);
  }

  var values = FIELD_MAP.map(function (f) { return ans[f[1]] || ''; });
  sh.getRange(row + 2, 4, 1, values.length).setValues([values]); // columns D.. onward

  // Keep the roster alphabetical by first name.
  var lastRow = sh.getLastRow();
  if (lastRow >= 3) {
    var cols = sh.getLastColumn();
    var data = sh.getRange(2, 1, lastRow - 1, cols).getValues();
    data.sort(function (a, b) { return String(a[0]).toLowerCase() < String(b[0]).toLowerCase() ? -1 : 1; });
    sh.getRange(2, 1, data.length, cols).setValues(data);
  }
}

// Find (by exact file name in the PVCS Admin folder) or create the per-person
// Shift Record spreadsheet and Notes doc. Returns [shiftUrl, notesUrl].
function getOrCreatePersonFiles_(folder, name, shiftCols) {
  var shiftTitle = name + ' - Shift Record';
  var docTitle   = name + ' - Notes & Archive';
  var shiftUrl = findFileUrl_(folder, shiftTitle);
  var docUrl   = findFileUrl_(folder, docTitle);

  if (!shiftUrl) {
    var s = SpreadsheetApp.create(shiftTitle);
    var sf = DriveApp.getFileById(s.getId()); folder.addFile(sf); DriveApp.getRootFolder().removeFile(sf);
    s.getSheets()[0].getRange(1, 1, 1, shiftCols.length).setValues([shiftCols]).setFontWeight('bold');
    shiftUrl = s.getUrl();
  }
  if (!docUrl) {
    var d = DocumentApp.create(docTitle);
    var df = DriveApp.getFileById(d.getId()); folder.addFile(df); DriveApp.getRootFolder().removeFile(df);
    d.getBody().appendParagraph(name + ' - Notes & Archived Emails').setHeading(DocumentApp.ParagraphHeading.HEADING1);
    d.saveAndClose();
    docUrl = d.getUrl();
  }
  return [shiftUrl, docUrl];
}

function findFileUrl_(folder, fileName) {
  var it = folder.getFilesByName(fileName);
  return it.hasNext() ? it.next().getUrl() : '';
}

/* ===== Maintenance: add the new PAs + sort both rosters alphabetically ===== */

function organizeRoster() {
  var ss = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty(ROSTER_PROP));
  var folder = getOrCreateFolder_(FOLDER_NAME);
  var byName = function (a, b) { return String(a).toLowerCase() < String(b).toLowerCase() ? -1 : 1; };

  // Psychiatrists: sort existing rows alphabetically (preserves all answer data).
  var psy = ss.getSheetByName('Psychiatrists');
  var pCols = psy.getLastColumn();
  if (psy.getLastRow() >= 2) {
    var pData = psy.getRange(2, 1, psy.getLastRow() - 1, pCols).getValues();
    pData.sort(function (a, b) { return byName(a[0], b[0]); });
    psy.getRange(2, 1, pData.length, pCols).setValues(pData);
  }

  // PAs: reuse existing per-person files, create files for new names, then sort.
  var paNames = ['Jill','Lindsay','Britney','Dan','Kylee','Dana','Olivia'].sort(byName);
  var pa = ss.getSheetByName('PAs');
  var existing = {};
  if (pa.getLastRow() >= 2) {
    pa.getRange(2, 1, pa.getLastRow() - 1, 3).getValues().forEach(function (r) {
      if (r[0]) existing[r[0]] = [r[1], r[2]];
    });
  }
  var rows = paNames.map(function (name) {
    var urls = existing[name];
    if (!urls) {
      var s = SpreadsheetApp.create(name + ' - Shift Record');
      var sf = DriveApp.getFileById(s.getId()); folder.addFile(sf); DriveApp.getRootFolder().removeFile(sf);
      s.getSheets()[0].getRange(1, 1, 1, 3).setValues([['Date','Day / Evening','Notes']]).setFontWeight('bold');
      var doc = DocumentApp.create(name + ' - Notes & Archive');
      var df = DriveApp.getFileById(doc.getId()); folder.addFile(df); DriveApp.getRootFolder().removeFile(df);
      doc.getBody().appendParagraph(name + ' - Notes & Archived Emails').setHeading(DocumentApp.ParagraphHeading.HEADING1);
      doc.saveAndClose();
      urls = [s.getUrl(), doc.getUrl()];
    }
    return [name, urls[0], urls[1]];
  });
  if (pa.getLastRow() >= 2) pa.getRange(2, 1, pa.getLastRow() - 1, 3).clearContent();
  pa.getRange(2, 1, rows.length, 3).setValues(rows);

  // Form name-key dropdown: sorted psychiatrists + sorted PAs + Other.
  var psyNames = PSYCHIATRISTS.slice().sort(byName);
  var form = FormApp.openById(PropertiesService.getScriptProperties().getProperty(FORM_PROP_KEY));
  form.getItems(FormApp.ItemType.LIST).forEach(function (it) {
    if (it.getTitle() === KEY_QUESTION) {
      it.asListItem().setChoiceValues(psyNames.concat(paNames).concat(['Other (new team member)']));
    }
  });
  Logger.log('Organized. Psychiatrists + PAs sorted; PAs = ' + paNames.join(', '));
}

/* ============================ PHASE 1 (kept) ============================ */

function installNotifyTrigger() {
  var formId = PropertiesService.getScriptProperties().getProperty(FORM_PROP_KEY);
  if (!formId) throw new Error('No form id stored yet - run buildPvcsOnboardingForm first.');
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'onPvcsFormSubmit') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('onPvcsFormSubmit').forForm(formId).onFormSubmit().create();
}

function testNotifyEmail() {
  MailApp.sendEmail(NOTIFY_EMAIL, 'PVCS form notifications are working',
    'This confirms the PVCS On-Boarding form will email you here whenever someone submits it.');
}
