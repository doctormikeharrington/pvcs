/**
 * PVCS Open Shift Claims + Give-Up web app
 * Sheet: "PVCS Open Shift Claims — Aug/Sep 2026" (18cGVmmviGtWPzZYwh-hhLZet5NTDNatCaSY62NbhrM8)
 * Deployed as web app: execute as Mike, accessible to anyone with the link.
 * Claim page (default URL): shows only OPEN rows of "Shifts"; claiming is atomic via LockService.
 * Give-up page (URL + ?page=giveup): psychiatrist picks their scheduled shift; it goes OPEN
 * on the claim page and the whole team (Team tab emails) is notified automatically.
 */
const SHEET_ID = '18cGVmmviGtWPzZYwh-hhLZet5NTDNatCaSY62NbhrM8';
const TAB_NAME = 'Shifts';
const SCHED_TAB = 'Schedule';
const TEAM_TAB = 'Team';
const TZ = 'America/Winnipeg';
const ADMIN_EMAIL = 'harringtonmike@hotmail.com';
const CLAIM_URL = 'https://script.google.com/macros/s/AKfycbxcCUN_MtRp-o90xU0kfqKUF5AKlwBodaa9R3avtXL1U-lU6Yfuz5jWuOI4KUn61C7PBw/exec';

const DAY_LABEL = 'Day (08:00–17:00)';
const EVE_LABEL = 'Evening (17:00–08:00)';

// Run once to create/populate all tabs. CAUTION: wipes existing claims/give-ups.
function setup() {
  const ss = SpreadsheetApp.openById(SHEET_ID);

  // ---- Shifts tab (claimable shifts) ----
  let sh = ss.getSheetByName(TAB_NAME);
  if (!sh) sh = ss.insertSheet(TAB_NAME);
  sh.clear();
  const rows = [
    ['ID', 'Shift', 'Status', 'Claimed by', 'Email', 'Claimed at', 'Key'],
    [1, 'Monday, August 31 — Evening (17:00–08:00)', 'ASSIGNED', 'Dr. Zylberman', '', '2026-08-07 (assigned)', '2026-08-31|Evening'],
    [2, 'Tuesday, September 1 — Evening (17:00–08:00)', 'ASSIGNED', 'Dr. Zylberman', '', '2026-08-07 (assigned)', '2026-09-01|Evening'],
    [3, 'Wednesday, September 2 — Evening (17:00–08:00)', 'ASSIGNED', 'Dr. Zylberman', '', '2026-08-07 (assigned)', '2026-09-02|Evening'],
    [4, 'Thursday, September 3 — Evening (17:00–08:00)', 'ASSIGNED', 'Dr. Zylberman', '', '2026-08-07 (assigned)', '2026-09-03|Evening'],
    [5, 'Friday, September 4 — Evening (17:00–08:00)', 'ASSIGNED', 'Dr. Zylberman', '', '2026-08-07 (assigned)', '2026-09-04|Evening'],
    [6, 'Saturday, September 5 — Evening (17:00–08:00)', 'OPEN', '', '', '', '2026-09-05|Evening'],
    [7, 'Sunday, September 6 — Evening (17:00–08:00)', 'OPEN', '', '', '', '2026-09-06|Evening'],
    [8, 'Saturday, September 19 — Evening (17:00–08:00)', 'HOLD', '', '', '', '2026-09-19|Evening'],
    [9, 'Sunday, September 20 — Evening (17:00–08:00)', 'HOLD', '', '', '', '2026-09-20|Evening']
  ];
  sh.getRange(1, 1, rows.length, 7).setValues(rows);
  sh.getRange(1, 1, 1, 7).setFontWeight('bold');
  sh.setColumnWidth(1, 50); sh.setColumnWidth(2, 320); sh.setColumnWidth(3, 90);
  sh.setColumnWidth(4, 160); sh.setColumnWidth(5, 220); sh.setColumnWidth(6, 160); sh.setColumnWidth(7, 150);
  sh.setFrozenRows(1);

  // ---- Schedule tab (full assignment schedule, Option B final) ----
  // per week: [dayMonFri, dayWeekend, eveMonFri, eveWeekend]; null = special-cased below
  const wk = [
    ['Dr. Kuzenko', 'Dr. Kuzenko', 'Dr. Zylberman', null],   // W1: eve weekend OPEN (claimable)
    ['Dr. Classen', 'Dr. Classen', 'Dr. Okoye', 'Dr. Okoye'],
    ['Dr. Okoye', 'Dr. Okoye', 'Dr. Sawich', null],          // W3: eve weekend HOLD (TBD)
    ['Dr. Brainch', 'Dr. Okoye', 'Dr. Ruzhynsky', 'Dr. Ruzhynsky'],
    ['Dr. Gill', 'Dr. Gill', 'Dr. Classen', 'Dr. Classen'],
    ['Dr. Sawich', 'Dr. Sawich', 'Dr. Palay', 'Dr. Palay'],
    ['Dr. Kuzenko', 'Dr. Kuzenko', 'Dr. Ruzhynsky', 'Dr. Ruzhynsky'],
    ['Dr. Konrad', 'Dr. Konrad', 'Dr. Okoye', 'Dr. Okoye'],
    ['Dr. Bolton', 'Dr. Bolton', 'Dr. Classen', 'Dr. Classen'],
    ['Dr. Konrad', 'Dr. Konrad', 'Dr. Okoye', 'Dr. Okoye']
  ];
  const dayNames = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'];
  const schedRows = [['Date', 'Weekday', 'Type', 'Assigned to', 'Status', 'Notes']];
  for (let w = 0; w < 10; w++) {
    for (let d = 0; d < 7; d++) {
      const dt = new Date(2026, 7, 31, 12);
      dt.setDate(dt.getDate() + w * 7 + d);
      const iso = Utilities.formatDate(dt, TZ, 'yyyy-MM-dd');
      const wd = dayNames[d];
      const weekend = d >= 5;
      const dayAssigned = weekend ? wk[w][1] : wk[w][0];
      let eveAssigned = weekend ? wk[w][3] : wk[w][2];
      schedRows.push([iso, wd, 'Day', dayAssigned, 'SCHEDULED', '']);
      if (eveAssigned === null) {
        const st = (w === 0) ? 'OPEN' : 'HOLD';
        schedRows.push([iso, wd, 'Evening', '', st, (w === 0) ? 'claimable' : 'TBD — to be scheduled separately']);
      } else {
        schedRows.push([iso, wd, 'Evening', eveAssigned, 'SCHEDULED', '']);
      }
    }
  }
  let sc = ss.getSheetByName(SCHED_TAB);
  if (!sc) sc = ss.insertSheet(SCHED_TAB);
  sc.clear();
  sc.getRange(1, 1, schedRows.length, 6).setValues(schedRows);
  sc.getRange(1, 1, 1, 6).setFontWeight('bold');
  sc.setFrozenRows(1);
  sc.setColumnWidth(1, 100); sc.setColumnWidth(2, 100); sc.setColumnWidth(3, 80);
  sc.setColumnWidth(4, 150); sc.setColumnWidth(5, 110); sc.setColumnWidth(6, 260);

  // ---- Team tab (broadcast list — Mike fills in emails) ----
  let tm = ss.getSheetByName(TEAM_TAB);
  if (!tm) tm = ss.insertSheet(TEAM_TAB);
  tm.clear();
  const team = [
    ['Name', 'Email'],
    ['Dr. Adelufosi', ''], ['Dr. Bezzahou', ''], ['Dr. Bolton', ''], ['Dr. Brainch', ''],
    ['Dr. Braun', ''], ['Dr. Classen', ''], ['Dr. Gill', ''], ['Dr. Harrington', ADMIN_EMAIL],
    ['Dr. Konrad', ''], ['Dr. Kuzenko', ''], ['Dr. Okoye', ''], ['Dr. Palay', ''],
    ['Dr. Paletta', ''], ['Dr. Reimer', ''], ['Dr. Ruzhynsky', ''], ['Dr. Sareen', ''],
    ['Dr. Sawich', ''], ['Dr. Zylberman', '']
  ];
  tm.getRange(1, 1, team.length, 2).setValues(team);
  tm.getRange(1, 1, 1, 2).setFontWeight('bold');
  tm.setColumnWidth(1, 140); tm.setColumnWidth(2, 260);
  tm.setFrozenRows(1);

  const s1 = ss.getSheetByName('Sheet1');
  if (s1 && ss.getSheets().length > 1) ss.deleteSheet(s1);
}

// ---------- helpers ----------
function labelFor_(iso, weekday, type) {
  const m = ['January','February','March','April','May','June','July','August','September','October','November','December'];
  const parts = iso.split('-');
  const month = m[Number(parts[1]) - 1];
  const dayNum = Number(parts[2]);
  return weekday + ', ' + month + ' ' + dayNum + ' — ' + (type === 'Day' ? DAY_LABEL : EVE_LABEL);
}

function teamEmails_() {
  const tm = SpreadsheetApp.openById(SHEET_ID).getSheetByName(TEAM_TAB);
  const data = tm.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < data.length; i++) {
    const e = String(data[i][1] || '').trim();
    if (e) out.push(e);
  }
  return out;
}

function getOpenShifts_() {
  const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(TAB_NAME);
  const data = sh.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][2]).toUpperCase() === 'OPEN') {
      out.push({ id: data[i][0], label: data[i][1] });
    }
  }
  return out;
}

// ---------- claim page API ----------
function getOpenShifts() {
  return getOpenShifts_();
}

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
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sh = ss.getSheetByName(TAB_NAME);
    const data = sh.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        if (String(data[i][2]).toUpperCase() !== 'OPEN') {
          return { ok: false, msg: 'Sorry — that shift was just claimed by someone else. Please choose another.', shifts: getOpenShifts_() };
        }
        const stamp = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss');
        sh.getRange(i + 1, 3, 1, 4).setValues([['CLAIMED', name, email, stamp]]);
        // reflect in Schedule tab
        const key = String(data[i][6] || '');
        if (key) {
          const kp = key.split('|');
          const sc = ss.getSheetByName(SCHED_TAB);
          const sd = sc.getDataRange().getValues();
          for (let j = 1; j < sd.length; j++) {
            if (Utilities.formatDate(new Date(sd[j][0]), TZ, 'yyyy-MM-dd') === kp[0] && sd[j][2] === kp[1]) {
              sc.getRange(j + 1, 4, 1, 3).setValues([[name, 'SCHEDULED', 'claimed ' + stamp]]);
              break;
            }
          }
        }
        // notify Mike
        try {
          MailApp.sendEmail(ADMIN_EMAIL, 'PVCS shift claimed: ' + data[i][1],
            name + ' (' + email + ') claimed:\n\n' + data[i][1] + '\n\nRecorded ' + stamp + '.');
        } catch (e) {}
        return { ok: true, msg: 'You have the shift: ' + data[i][1] + '. Dr. Harrington will follow up by email to confirm.', shifts: getOpenShifts_() };
      }
    }
    return { ok: false, msg: 'Shift not found — please reload the page.', shifts: getOpenShifts_() };
  } finally {
    lock.releaseLock();
  }
}

// ---------- give-up page API ----------
function getTeamNames() {
  const tm = SpreadsheetApp.openById(SHEET_ID).getSheetByName(TEAM_TAB);
  const data = tm.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) out.push(String(data[i][0]));
  }
  return out;
}

function getMyShifts(name) {
  name = String(name || '').trim();
  const sc = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SCHED_TAB);
  const data = sc.getDataRange().getValues();
  const today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  const out = [];
  for (let i = 1; i < data.length; i++) {
    const iso = Utilities.formatDate(new Date(data[i][0]), TZ, 'yyyy-MM-dd');
    if (String(data[i][3]) === name && String(data[i][4]) === 'SCHEDULED' && iso >= today) {
      out.push({ key: iso + '|' + data[i][2], label: labelFor_(iso, data[i][1], data[i][2]) });
    }
  }
  return out;
}

function giveUpShift(name, key) {
  name = String(name || '').trim();
  key = String(key || '');
  const kp = key.split('|');
  if (!name || kp.length !== 2) return { ok: false, msg: 'Something went wrong — please reload the page.', shifts: [] };
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sc = ss.getSheetByName(SCHED_TAB);
    const sd = sc.getDataRange().getValues();
    const stamp = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss');
    for (let j = 1; j < sd.length; j++) {
      const iso = Utilities.formatDate(new Date(sd[j][0]), TZ, 'yyyy-MM-dd');
      if (iso === kp[0] && sd[j][2] === kp[1]) {
        if (String(sd[j][3]) !== name || String(sd[j][4]) !== 'SCHEDULED') {
          return { ok: false, msg: 'That shift is no longer listed under your name — please reload the page.', shifts: getMyShifts(name) };
        }
        const label = labelFor_(iso, sd[j][1], sd[j][2]);
        // mark schedule
        sc.getRange(j + 1, 4, 1, 3).setValues([['', 'GIVEN UP', 'given up by ' + name + ' ' + stamp]]);
        // add to claimable shifts
        const sh = ss.getSheetByName(TAB_NAME);
        const ids = sh.getDataRange().getValues().slice(1).map(function (r) { return Number(r[0]) || 0; });
        const newId = (ids.length ? Math.max.apply(null, ids) : 0) + 1;
        sh.appendRow([newId, label, 'OPEN', '', '', '', key]);
        // broadcast
        const emails = teamEmails_();
        try {
          MailApp.sendEmail({
            to: ADMIN_EMAIL,
            bcc: emails.join(','),
            subject: 'PVCS open shift available: ' + label,
            body: name + ' has given up the following shift:\n\n' + label +
                  '\n\nFirst come, first served — claim it here:\n' + CLAIM_URL +
                  '\n\nOnce someone claims it, it disappears from the claim page.'
          });
        } catch (e) {}
        return { ok: true, msg: 'Done — your shift (' + label + ') is now posted and the team has been emailed. Dr. Harrington will follow up to confirm.', shifts: getMyShifts(name) };
      }
    }
    return { ok: false, msg: 'Shift not found — please reload the page.', shifts: getMyShifts(name) };
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  const page = (e && e.parameter && e.parameter.page) || '';
  if (page === 'giveup') {
    return HtmlService.createHtmlOutputFromFile('GiveUp')
      .setTitle('PVCS — Give Up a Shift')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('PVCS Open Shifts')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}
