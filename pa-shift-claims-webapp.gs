/**
 * PVCS PA shift give-away → claim system.
 *
 * Flow: a PA submits the existing "PA CRC — I cannot work a scheduled shift" Google Form
 * → onGiveAway trigger posts that shift as OPEN on the claim page and emails every PA
 * → first PA to claim it gets it (atomic via LockService); shift disappears from the page.
 *
 * setup() is safe to re-run: it only creates what is missing and never clears data.
 */
const GIVEAWAY_FORM_ID = '1Bgl7CygReQqzt3B6ez8Y_m1tM0WO_dykvEvzxdXq2-E';
const ADMIN_EMAIL = 'harringtonmike@hotmail.com';
const TZ = 'America/Winnipeg';
const SHIFTS_TAB = 'Shifts';
const TEAM_TAB = 'PA Team';
const PROP_SHEET_ID = 'PA_SHEET_ID';

// Name, email. Blank email = not emailed (fill it in on the "PA Team" tab).
const PA_TEAM = [
  ['Kylee Barnabe', 'kbarnabe2@sharedhealthmb.ca'],
  ['Olivia Coneys', 'oconeys@sharedhealthmb.ca'],
  ['Jillian Desautels', 'jdesautels4@sharedhealthmb.ca'],
  ['Brittany Devaney', 'bdevaney@sharedhealthmb.ca'],
  ['Daniel Fillion', 'dfillion3@sharedhealthmb.ca'],
  ['Karin Love', 'klove@hsc.mb.ca'],
  ['Alana Ramnauth', 'aramnauth@hsc.mb.ca'],
  ['Lindsey Shumila', 'lshumila@wrha.mb.ca'],
  ['Dana Skaritko', 'dskaritko2@sharedhealthmb.ca'],
  ['Jennifer Wilson', 'jwilson2@sharedhealthmb.ca']
];

// ---------- setup (idempotent) ----------
function setup() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty(PROP_SHEET_ID);
  let ss;
  if (id) {
    ss = SpreadsheetApp.openById(id);
  } else {
    ss = SpreadsheetApp.create('PVCS PA Shift Claims');
    id = ss.getId();
    props.setProperty(PROP_SHEET_ID, id);
  }

  let sh = ss.getSheetByName(SHIFTS_TAB);
  if (!sh) {
    sh = ss.insertSheet(SHIFTS_TAB);
    sh.getRange(1, 1, 1, 8).setValues([['ID', 'Shift', 'Status', 'Claimed by', 'Email',
      'Claimed at', 'Given up by', 'Posted at']]);
    sh.getRange(1, 1, 1, 8).setFontWeight('bold');
    sh.setColumnWidth(2, 320); sh.setColumnWidth(4, 160); sh.setColumnWidth(5, 220);
    sh.setColumnWidth(6, 150); sh.setColumnWidth(7, 150); sh.setColumnWidth(8, 150);
    sh.setFrozenRows(1);
  }

  let tm = ss.getSheetByName(TEAM_TAB);
  if (!tm) {
    tm = ss.insertSheet(TEAM_TAB);
    const rows = [['Name', 'Email']].concat(PA_TEAM);
    tm.getRange(1, 1, rows.length, 2).setValues(rows);
    tm.getRange(1, 1, 1, 2).setFontWeight('bold');
    tm.setColumnWidth(1, 160); tm.setColumnWidth(2, 260);
    tm.setFrozenRows(1);
  }

  const s1 = ss.getSheetByName('Sheet1');
  if (s1 && ss.getSheets().length > 1 && s1.getLastRow() === 0) ss.deleteSheet(s1);

  // install the form-submit trigger exactly once
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'onGiveAway') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('onGiveAway').forForm(GIVEAWAY_FORM_ID).onFormSubmit().create();

  Logger.log('sheet=' + ss.getUrl() + ' | webapp=' + (ScriptApp.getService().getUrl() || 'not deployed yet'));
}

// ---------- helpers ----------
function sheet_(name) {
  const id = PropertiesService.getScriptProperties().getProperty(PROP_SHEET_ID);
  return SpreadsheetApp.openById(id).getSheetByName(name);
}

function teamEmails_() {
  const data = sheet_(TEAM_TAB).getDataRange().getValues();
  const out = [];
  for (let i = 1; i < data.length; i++) {
    const e = String(data[i][1] || '').trim();
    if (e) out.push(e);
  }
  return out;
}

function getOpenShifts_() {
  const data = sheet_(SHIFTS_TAB).getDataRange().getValues();
  const out = [];
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][2]).toUpperCase() === 'OPEN') out.push({ id: data[i][0], label: data[i][1] });
  }
  return out;
}

// ---------- claim page ----------
function getOpenShifts() { return getOpenShifts_(); }

function claimShift(id, name, email) {
  name = String(name || '').trim();
  email = String(email || '').trim();
  if (!name || !email) {
    return { ok: false, msg: 'Please enter both your name and your email address.', shifts: getOpenShifts_() };
  }
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
    return { ok: false, msg: "That email address doesn't look right — please check it and try again.", shifts: getOpenShifts_() };
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = sheet_(SHIFTS_TAB);
    const data = sh.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        if (String(data[i][2]).toUpperCase() !== 'OPEN') {
          return { ok: false, msg: 'Sorry — that shift was just claimed by someone else. Please choose another.', shifts: getOpenShifts_() };
        }
        const stamp = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss');
        sh.getRange(i + 1, 3, 1, 4).setValues([['CLAIMED', name, email, stamp]]);
        const label = data[i][1];
        const gaveUp = data[i][6];
        try {
          MailApp.sendEmail({
            to: email,
            cc: ADMIN_EMAIL,
            subject: 'PVCS shift confirmed: ' + label,
            body: name + ',\n\nYou have claimed this shift:\n\n' + label +
                  (gaveUp ? '\n(originally scheduled to ' + gaveUp + ')' : '') +
                  '\n\nRecorded ' + stamp + '. Dr. Harrington has been copied on this confirmation.'
          });
        } catch (err) {}
        return { ok: true, msg: 'Confirmed — you have the shift: ' + label + '. A confirmation email is on its way.', shifts: getOpenShifts_() };
      }
    }
    return { ok: false, msg: 'Shift not found — please reload the page.', shifts: getOpenShifts_() };
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('PVCS PA — Open Shifts')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ---------- give-away form trigger ----------
function onGiveAway(e) {
  const answers = {};
  e.response.getItemResponses().forEach(function (ir) {
    answers[ir.getItem().getTitle()] = ir.getResponse();
  });
  const who = answers['Your name'] || 'A PA';
  let date = '';
  let which = '';
  let note = '';
  Object.keys(answers).forEach(function (k) {
    if (k.indexOf('Date of the shift') === 0) date = answers[k];
    else if (k.indexOf('Which shift') === 0) which = answers[k];
    else if (k.indexOf('Note') === 0) note = answers[k];
  });

  const label = date + ' — ' + which;
  const stamp = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss');

  const sh = sheet_(SHIFTS_TAB);
  const ids = sh.getDataRange().getValues().slice(1).map(function (r) { return Number(r[0]) || 0; });
  const newId = (ids.length ? Math.max.apply(null, ids) : 0) + 1;
  sh.appendRow([newId, label, 'OPEN', '', '', '', who, stamp]);

  const url = ScriptApp.getService().getUrl();
  const emails = teamEmails_();
  const body = who + ' is unable to work the following shift:\n\n' + label +
    (note ? '\n\nNote: ' + note : '') +
    '\n\nFirst come, first served — claim it here:\n' + url +
    '\n\nThe shift disappears from that page as soon as someone claims it, and the claimer gets an automatic confirmation.';
  try {
    MailApp.sendEmail({
      to: ADMIN_EMAIL,
      bcc: emails.join(','),
      subject: 'PVCS PA shift available: ' + label,
      body: body
    });
  } catch (err) {
    Logger.log('mail failed: ' + err);
  }
  Logger.log('posted ' + label + ' from ' + who + ' to ' + emails.length + ' PAs');
}
