/**
 * PVCS — Psychiatrist open-shift claim page (Nov 9 2026 – Jan 3 2027).
 *
 * Same pattern as final-shift-claims.gs: page lists the open shifts, first
 * person to claim one gets it (atomic via LockService), the shift disappears
 * from the list, and the claimer + Dr. Harrington get a confirmation email.
 *
 * Shifts are listed as INDIVIDUAL DAYS so a psychiatrist can take part of a
 * block. Anyone can claim several in a row — each claim is a separate click.
 *
 * setup() is FIRST in this file on purpose (the editor's Run button runs the
 * first function). It is idempotent — safe to re-run, never clears data.
 * To offer more shifts later: add a row to the Shifts tab
 * (ID, Shift label, OPEN) — the page picks it up automatically, no redeploy.
 */
const ADMIN_EMAIL = 'harringtonmike@hotmail.com';
const TZ = 'America/Winnipeg';
const SHIFTS_TAB = 'Shifts';
const PROP_SHEET_ID = 'PSYCH_NOV_JAN_SHEET_ID';

/* The four open blocks from the Nov 9 – Jan 3 schedule, broken into days:
 *   EVENING  Mon Nov 9  – Fri Nov 13  (Mon–Fri only; Nov 14–15 = Dr. Ruzhynsky)
 *   DAY      Sat Nov 21 – Sun Nov 22  (weekend of Week 2)
 *   DAY      Mon Dec 14 – Sun Dec 20  (full Week 6)
 *   EVENING  Mon Dec 28 – Sun Jan 3   (full Week 8)
 */
const INITIAL_SHIFTS = [
  ['Mon Nov 9 — Evening'],
  ['Tue Nov 10 — Evening'],
  ['Wed Nov 11 — Evening'],
  ['Thu Nov 12 — Evening'],
  ['Fri Nov 13 — Evening'],
  ['Sat Nov 21 — Day'],
  ['Sun Nov 22 — Day'],
  ['Mon Dec 14 — Day'],
  ['Tue Dec 15 — Day'],
  ['Wed Dec 16 — Day'],
  ['Thu Dec 17 — Day'],
  ['Fri Dec 18 — Day'],
  ['Sat Dec 19 — Day'],
  ['Sun Dec 20 — Day'],
  ['Mon Dec 28 — Evening'],
  ['Tue Dec 29 — Evening'],
  ['Wed Dec 30 — Evening'],
  ['Thu Dec 31 — Evening (New Year\u2019s Eve)'],
  ['Fri Jan 1 — Evening (New Year\u2019s Day)'],
  ['Sat Jan 2 — Evening'],
  ['Sun Jan 3 — Evening']
];

function setup() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty(PROP_SHEET_ID);
  let ss;
  if (id) {
    ss = SpreadsheetApp.openById(id);
  } else {
    ss = SpreadsheetApp.create('PVCS Psychiatrist Shift Claims — Nov 2026-Jan 2027');
    id = ss.getId();
    props.setProperty(PROP_SHEET_ID, id);
  }
  let sh = ss.getSheetByName(SHIFTS_TAB);
  if (!sh) {
    sh = ss.getSheets()[0].getName() === 'Sheet1' && ss.getSheets().length === 1
      ? ss.getSheets()[0].setName(SHIFTS_TAB)
      : ss.insertSheet(SHIFTS_TAB);
    sh.getRange(1, 1, 1, 6).setValues([['ID', 'Shift', 'Status', 'Claimed by', 'Email', 'Claimed at']]);
    sh.getRange(1, 1, 1, 6).setFontWeight('bold');
    sh.setColumnWidth(2, 320); sh.setColumnWidth(4, 160); sh.setColumnWidth(5, 220); sh.setColumnWidth(6, 150);
    sh.setFrozenRows(1);
    const rows = INITIAL_SHIFTS.map(function (s, i) { return [i + 1, s[0], 'OPEN', '', '', '']; });
    sh.getRange(2, 1, rows.length, 6).setValues(rows);
  }
  Logger.log('sheet=' + ss.getUrl());
  Logger.log('webapp=' + (ScriptApp.getService().getUrl() || 'not deployed yet'));
}

function sheet_() {
  const id = PropertiesService.getScriptProperties().getProperty(PROP_SHEET_ID);
  return SpreadsheetApp.openById(id).getSheetByName(SHIFTS_TAB);
}

function getOpenShifts_() {
  const data = sheet_().getDataRange().getValues();
  const out = [];
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][2]).toUpperCase() === 'OPEN') out.push({ id: data[i][0], label: data[i][1] });
  }
  return out;
}

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
    const sh = sheet_();
    const data = sh.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        if (String(data[i][2]).toUpperCase() !== 'OPEN') {
          return { ok: false, msg: 'Sorry — that shift was just claimed by someone else.', shifts: getOpenShifts_() };
        }
        const stamp = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss');
        sh.getRange(i + 1, 3, 1, 4).setValues([['CLAIMED', name, email, stamp]]);
        const label = data[i][1];
        try {
          MailApp.sendEmail({
            to: email,
            cc: ADMIN_EMAIL,
            subject: 'PVCS shift confirmed: ' + label,
            body: name + ',\n\nYou have claimed this PVCS shift:\n\n' + label +
                  '\n\nRecorded ' + stamp + '. Dr. Harrington has been copied on this confirmation ' +
                  'and will add it to the published schedule.'
          });
        } catch (err) {}
        return { ok: true, msg: 'Confirmed — you have ' + label + '. A confirmation email is on its way. You may claim another shift below.', shifts: getOpenShifts_() };
      }
    }
    return { ok: false, msg: 'Shift not found — please reload the page.', shifts: getOpenShifts_() };
  } finally {
    lock.releaseLock();
  }
}

function claimPage_() {
  var html = `<!DOCTYPE html><html><head><meta name="viewport" content="width=device-width, initial-scale=1">
<style>
body{font-family:-apple-system,Segoe UI,Roboto,Arial,sans-serif;background:#f4f6f9;margin:0;padding:24px;color:#1f2937}
.card{max-width:560px;margin:0 auto;background:#fff;border:1px solid #dde3ea;border-radius:10px;padding:24px}
h1{font-size:1.25rem;color:#1f3a5f;margin:0 0 4px}
p.sub{color:#5a6675;font-size:.9rem;margin:0 0 18px}
.shift{border:1px solid #dde3ea;border-radius:8px;padding:12px 14px;margin-bottom:10px}
.shift b{color:#1f3a5f}
button{background:#1f3a5f;color:#fff;border:0;border-radius:6px;padding:8px 14px;font-size:.9rem;cursor:pointer}
button.confirm{background:#2c7a7b}
button.cancel{background:#94a3b8}
input{width:100%;box-sizing:border-box;padding:8px;margin:6px 0;border:1px solid #cbd5e1;border-radius:6px;font-size:.95rem}
.row{display:flex;justify-content:space-between;align-items:center;gap:10px}
.msg{padding:10px 14px;border-radius:8px;margin-bottom:14px;font-size:.92rem}
.ok{background:#ecfdf5;border:1px solid #34d399}
.bad{background:#fef2f2;border:1px solid #fca5a5}
.empty{color:#5a6675;font-style:italic}
</style></head><body><div class="card">
<h1>PVCS &mdash; Open Psychiatrist Shifts (Nov 9 2026 &ndash; Jan 3 2027)</h1>
<p class="sub">First come, first served &mdash; a shift disappears as soon as someone claims it. Claim as many as you like, one at a time. You will receive a confirmation email for each.</p>
<div id="msg"></div><div id="list"><p class="empty">Loading&hellip;</p></div>
</div>
<script>
var open=null,last=[];
function esc(s){return String(s).replace(/[&<>"]/g,function(c){return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c];});}
function msg(t,ok){document.getElementById('msg').innerHTML = t ? ('<div class="msg '+(ok?'ok':'bad')+'">'+esc(t)+'</div>') : '';}
function render(shifts){
  last = shifts;
  var el = document.getElementById('list');
  if(!shifts.length){ el.innerHTML = '<p class="empty">No open shifts remain.</p>'; return; }
  el.innerHTML = shifts.map(function(s){
    var head;
    if(open === s.id){
      head = '<b>'+esc(s.label)+'</b>' +
             '<input id="nm" placeholder="Your name">' +
             '<input id="em" type="email" placeholder="Your email">' +
             '<div class="row"><button class="confirm" data-act="go" data-id="'+s.id+'">Confirm claim</button>' +
             '<button class="cancel" data-act="cancel">Cancel</button></div>';
    } else {
      head = '<div class="row"><b>'+esc(s.label)+'</b>' +
             '<button data-act="pick" data-id="'+s.id+'">Claim</button></div>';
    }
    return '<div class="shift">'+head+'</div>';
  }).join('');
}
function pick(id){ open = id; render(last); }
function go(id){
  var n = document.getElementById('nm').value, e = document.getElementById('em').value;
  msg('Claiming…', true);
  google.script.run.withSuccessHandler(function(r){ open=null; msg(r.msg, r.ok); render(r.shifts); }).claimShift(id, n, e);
}
// NOTE: inline onclick attributes get stripped by the Apps Script sanitizer,
// so button actions MUST go through this delegated listener.
document.getElementById('list').addEventListener('click', function(ev){
  var b = ev.target;
  while (b && b.tagName !== 'BUTTON') b = b.parentElement;
  if (!b) return;
  var act = b.getAttribute('data-act');
  var id = Number(b.getAttribute('data-id'));
  if (act === 'pick') pick(id);
  else if (act === 'cancel') pick(null);
  else if (act === 'go') go(id);
});
google.script.run.withSuccessHandler(render).getOpenShifts();
</scr` + `ipt></body></html>`;
  return HtmlService.createHtmlOutput(html)
    .setTitle('PVCS — Open Psychiatrist Shifts')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/* ==========================================================================
 * GIVE-UP PAGE  (added 2026-10-19)
 * URL: <web app url>?page=giveup
 * A psychiatrist picks their name, sees the shifts they are scheduled for,
 * and releases one. It is marked GIVEN UP on the Schedule tab and appended to
 * the Shifts tab as OPEN, so it appears on the claim page immediately.
 * Run setupSchedule() ONCE to build the Schedule tab.
 * ======================================================================== */

const SCHEDULE_TAB = 'Schedule';
const SCHEDULE_START = '2026-11-09';
// One entry per day from SCHEDULE_START: "DayName|EveningName", days separated by ";"
const SCHEDULE_PAIRS = 'Brainch|Lewis;Brainch|Lewis;Brainch|Lewis;Brainch|Lewis;Brainch|Okoye;Brainch|Ruzhynsky;Brainch|Ruzhynsky;Sawich|Palay;Sawich|Palay;Sawich|Palay;Sawich|Palay;Sawich|Palay;Okoye|Palay;Okoye|Palay;Kuzenko|Sala;Kuzenko|Sala;Kuzenko|Sala;Kuzenko|Sala;Kuzenko|Sala;Kuzenko|Sala;Kuzenko|Sala;Konrad|Classen;Konrad|Classen;Konrad|Classen;Konrad|Classen;Konrad|Classen;Konrad|Classen;Konrad|Classen;Classen|Sala;Classen|Sala;Classen|Sala;Classen|Sala;Classen|Sala;Classen|Sala;Classen|Sala;Paletta|Sylvestre;Paletta|Sylvestre;Okoye|Sylvestre;OPEN|Sylvestre;Okoye|Sylvestre;OPEN|Sylvestre;Okoye|Sylvestre;Brainch|Okoye;Brainch|Okoye;Brainch|Okoye;Brainch|Okoye;Brainch|Okoye;Brainch|Okoye;Brainch|Okoye;Okoye|OPEN;Okoye|Palay;Okoye|OPEN;Okoye|OPEN;Okoye|OPEN;Okoye|OPEN;Okoye|OPEN';

function scheduleRows_() {
  const p = SCHEDULE_START.split('-');
  const start = new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]));
  const out = [];
  SCHEDULE_PAIRS.split(';').forEach(function (pair, i) {
    const d = new Date(start.getTime());
    d.setDate(d.getDate() + i);
    const iso = Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
    const nm = pair.split('|');
    out.push([iso, 'Day', nm[0]]);
    out.push([iso, 'Evening', nm[1]]);
  });
  return out;
}

function setupSchedule() {
  const id = PropertiesService.getScriptProperties().getProperty(PROP_SHEET_ID);
  const ss = SpreadsheetApp.openById(id);
  let sh = ss.getSheetByName(SCHEDULE_TAB);
  if (!sh) sh = ss.insertSheet(SCHEDULE_TAB);
  sh.clear();
  sh.getRange(1, 1, 1, 5)
    .setValues([['Date', 'Type', 'Assigned to', 'Status', 'Given up at']])
    .setFontWeight('bold');
  const out = scheduleRows_().map(function (r) {
    return [r[0], r[1], r[2], r[2] === 'OPEN' ? 'OPEN' : 'SCHEDULED', ''];
  });
  sh.getRange(2, 1, out.length, 5).setValues(out);
  sh.setFrozenRows(1);
  Logger.log('Schedule tab written: ' + out.length + ' rows');
}

function scheduleSheet_() {
  const id = PropertiesService.getScriptProperties().getProperty(PROP_SHEET_ID);
  return SpreadsheetApp.openById(id).getSheetByName(SCHEDULE_TAB);
}

function isoOf_(v) {
  return (v instanceof Date) ? Utilities.formatDate(v, TZ, 'yyyy-MM-dd') : String(v);
}

function labelFor_(iso, type) {
  const p = iso.split('-');
  const d = new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]));
  return Utilities.formatDate(d, TZ, 'EEE MMM d') + ' — ' + type;
}

function getNames() {
  const vals = scheduleSheet_().getDataRange().getValues().slice(1);
  const seen = {};
  vals.forEach(function (r) {
    if (r[2] && r[2] !== 'OPEN' && r[3] === 'SCHEDULED') seen[r[2]] = true;
  });
  return Object.keys(seen).sort();
}

function getMyShifts(name) {
  if (!name) return [];
  const vals = scheduleSheet_().getDataRange().getValues().slice(1);
  const out = [];
  vals.forEach(function (r) {
    if (String(r[2]) === String(name) && r[3] === 'SCHEDULED') {
      const iso = isoOf_(r[0]);
      out.push({ key: iso + '|' + r[1], label: labelFor_(iso, r[1]) });
    }
  });
  return out;
}

function giveUpShift(key, name) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const sh = scheduleSheet_();
    const vals = sh.getDataRange().getValues();
    for (let i = 1; i < vals.length; i++) {
      if (isoOf_(vals[i][0]) + '|' + vals[i][1] !== key) continue;
      if (String(vals[i][2]) !== String(name) || vals[i][3] !== 'SCHEDULED') {
        return { ok: false, msg: 'That shift is no longer yours to give up.', shifts: getMyShifts(name) };
      }
      const stamp = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss');
      sh.getRange(i + 1, 4, 1, 2).setValues([['GIVEN UP', stamp]]);

      const label = labelFor_(isoOf_(vals[i][0]), vals[i][1]);
      const shifts = sheet_();
      const sv = shifts.getDataRange().getValues();
      let maxId = 0;
      for (let k = 1; k < sv.length; k++) maxId = Math.max(maxId, Number(sv[k][0]) || 0);
      shifts.appendRow([maxId + 1, label, 'OPEN', '', '', '']);

      try {
        MailApp.sendEmail({
          to: ADMIN_EMAIL,
          bcc: teamEmails_().join(','),
          subject: 'PVCS open shift available: ' + label,
          body: 'An on-call shift has become available.\n\n'
              + '    ' + label + '\n\n'
              + 'Given up by Dr. ' + name + '.\n\n'
              + 'First come, first served. Claim it here:\n'
              + CLAIM_URL + '\n\n'
              + 'You are receiving this as a member of the PVCS psychiatrist group.\n'
        });
      } catch (err) {
        Logger.log('mail failed: ' + err);
      }
      return { ok: true, msg: 'Given up: ' + label + '. It is now open for others to claim.', shifts: getMyShifts(name) };
    }
    return { ok: false, msg: 'Shift not found.', shifts: getMyShifts(name) };
  } finally {
    lock.releaseLock();
  }
}

function giveUpPage_() {
  var html = `<!DOCTYPE html><html><head><meta name="viewport" content="width=device-width, initial-scale=1">
<style>
body{font-family:-apple-system,Segoe UI,Roboto,Arial,sans-serif;background:#f4f6f9;margin:0;padding:24px;color:#1f2937}
.card{max-width:560px;margin:0 auto;background:#fff;border:1px solid #dde3ea;border-radius:10px;padding:24px}
h1{font-size:1.25rem;color:#1f3a5f;margin:0 0 4px}
p.sub{color:#5a6675;font-size:.9rem;margin:0 0 18px}
.shift{border:1px solid #dde3ea;border-radius:8px;padding:12px 14px;margin-bottom:10px}
.shift b{color:#1f3a5f}
button{background:#1f3a5f;color:#fff;border:0;border-radius:6px;padding:8px 14px;font-size:.9rem;cursor:pointer}
button.confirm{background:#b91c1c}
button.cancel{background:#94a3b8}
select{width:100%;box-sizing:border-box;padding:8px;margin:6px 0 16px;border:1px solid #cbd5e1;border-radius:6px;font-size:.95rem}
.row{display:flex;justify-content:space-between;align-items:center;gap:10px}
.msg{padding:10px 14px;border-radius:8px;margin-bottom:14px;font-size:.92rem}
.ok{background:#ecfdf5;border:1px solid #34d399}
.bad{background:#fef2f2;border:1px solid #fca5a5}
.empty{color:#5a6675;font-style:italic}
.warn{font-size:.85rem;color:#5a6675;margin:6px 0 0}
</style></head><body><div class="card">
<h1>PVCS &mdash; Give Up a Shift (Nov 9 2026 &ndash; Jan 3 2027)</h1>
<p class="sub">Choose your name, then release any shift you cannot work. It becomes available on the claim page straight away. Your name stays on the published schedule until someone else claims it.</p>
<div id="msg"></div>
<select id="who"><option value="">Loading names&hellip;</option></select>
<div id="list"></div>
</div>
<script>
var open=null,last=[],who='';
function esc(s){return String(s).replace(/[&<>"]/g,function(c){return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c];});}
function msg(t,ok){document.getElementById('msg').innerHTML = t ? ('<div class="msg '+(ok?'ok':'bad')+'">'+esc(t)+'</div>') : '';}
function render(shifts){
  last = shifts || [];
  var el = document.getElementById('list');
  if(!who){ el.innerHTML = ''; return; }
  if(!last.length){ el.innerHTML = '<p class="empty">No shifts scheduled for you in this period.</p>'; return; }
  el.innerHTML = last.map(function(s){
    var head;
    if(open === s.key){
      head = '<b>'+esc(s.label)+'</b>' +
             '<p class="warn">This cannot be undone from here.</p>' +
             '<div class="row"><button class="confirm" data-act="go" data-key="'+esc(s.key)+'">Yes, give it up</button>' +
             '<button class="cancel" data-act="cancel">Cancel</button></div>';
    } else {
      head = '<div class="row"><b>'+esc(s.label)+'</b>' +
             '<button data-act="pick" data-key="'+esc(s.key)+'">Give up</button></div>';
    }
    return '<div class="shift">'+head+'</div>';
  }).join('');
}
function load(){
  msg('', true); open=null;
  if(!who){ render([]); return; }
  google.script.run.withSuccessHandler(render).getMyShifts(who);
}
document.getElementById('who').addEventListener('change', function(ev){ who = ev.target.value; load(); });
// NOTE: inline onclick attributes get stripped by the Apps Script sanitizer,
// so button actions MUST go through this delegated listener.
document.getElementById('list').addEventListener('click', function(ev){
  var b = ev.target;
  while (b && b.tagName !== 'BUTTON') b = b.parentElement;
  if (!b) return;
  var act = b.getAttribute('data-act');
  if (act === 'pick') { open = b.getAttribute('data-key'); render(last); }
  else if (act === 'cancel') { open = null; render(last); }
  else if (act === 'go') {
    msg('Releasing...', true);
    google.script.run.withSuccessHandler(function(r){
      open=null; msg(r.msg, r.ok); render(r.shifts);
    }).giveUpShift(b.getAttribute('data-key'), who);
  }
});
google.script.run.withSuccessHandler(function(names){
  var sel = document.getElementById('who');
  sel.innerHTML = '<option value="">Select your name...</option>' +
    names.map(function(n){ return '<option>'+esc(n)+'</option>'; }).join('');
}).getNames();
</scr` + `ipt></body></html>`;
  return HtmlService.createHtmlOutput(html)
    .setTitle('PVCS - Give Up a Shift')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function doGet(e) {
  if (e && e.parameter && e.parameter.page === 'giveup') return giveUpPage_();
  return claimPage_();
}


/* ==========================================================================
 * PSYCHIATRIST GROUP BROADCAST  (added 2026-10-19)
 * The standard PVCS psychiatrist email group, supplied by Dr. Harrington.
 * Run setupTeam() ONCE to write the Team tab; teamEmails_() reads the tab,
 * so the list can be edited in the sheet afterwards without touching code.
 * ======================================================================== */

const TEAM_TAB = 'Team';
const CLAIM_URL = 'https://script.google.com/macros/s/AKfycbyd3nsDgyKAWQ-uIW_tUslfZAU-lwDbDGxWb7pF8rTmUw9CI0F5wnq2ZdGc-ZZp_l95/exec';

const TEAM = [
  ['Adegoke Adelufosi', 'aadelufosi@nrha.ca'],
  ['Abdellah Bezzahou', 'ABezzahou@sbgh.mb.ca'],
  ['James Bolton', 'JBOLTON@hsc.mb.ca'],
  ['Navjot Brainch', 'nbrainch@hsc.mb.ca'],
  ['Kristen Braun', 'kbraun9@wrha.mb.ca'],
  ['Christopher Classen', 'cclassen@manitoba-physicians.ca'],
  ['Eunice Gill', 'EGill@hsc.mb.ca'],
  ['Geoffrey Konrad', 'gkonrad2@manitoba-physicians.ca'],
  ['Nina Kuzenko', 'nkuzenko@hsc.mb.ca'],
  ['Kylee Lewis', 'Klewis6@manitoba-physicians.ca'],
  ['Chijioke Okoye', 'cokoye@sharedhealthmb.ca'],
  ['Josh Palay', 'jpalay@sharedhealthmb.ca'],
  ['Antonio Paletta', 'apaletta@wrha.mb.ca'],
  ['Samantha Reimer', 'reimer35@myumanitoba.ca'],
  ['Jennifer Ruzhynsky', 'jruzhynsky@hsc.mb.ca'],
  ['Tanya Sala', 'tsala@hsc.mb.ca'],
  ['Jitender Sareen', 'JSareen@manitoba-physicians.ca'],
  ['Shauna Sawich', 'ssawich@hsc.mb.ca'],
  ['Devon Sylvestre', 'dsylvestre4@sbgh.mb.ca'],
  ['Melina Zylberman', 'mzylberman@manitoba-physicians.ca']
];

function setupTeam() {
  const id = PropertiesService.getScriptProperties().getProperty(PROP_SHEET_ID);
  const ss = SpreadsheetApp.openById(id);
  let sh = ss.getSheetByName(TEAM_TAB);
  if (!sh) sh = ss.insertSheet(TEAM_TAB);
  sh.clear();
  sh.getRange(1, 1, 1, 2).setValues([['Name', 'Email']]).setFontWeight('bold');
  sh.getRange(2, 1, TEAM.length, 2).setValues(TEAM);
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, 2);
  Logger.log('Team tab written: ' + TEAM.length + ' addresses');
}

function teamEmails_() {
  try {
    const id = PropertiesService.getScriptProperties().getProperty(PROP_SHEET_ID);
    const sh = SpreadsheetApp.openById(id).getSheetByName(TEAM_TAB);
    if (!sh) return [];
    return sh.getDataRange().getValues().slice(1)
      .map(function (r) { return String(r[1] || '').trim(); })
      .filter(function (e) { return e && e.indexOf('@') > 0 && e.toLowerCase() !== ADMIN_EMAIL.toLowerCase(); });
  } catch (err) {
    Logger.log('teamEmails_ failed: ' + err);
    return [];
  }
}
