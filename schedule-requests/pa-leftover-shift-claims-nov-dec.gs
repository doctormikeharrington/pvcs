/**
 * PVCS — PA leftover-shift claim page (Nov 1 – Dec 31 2026).
 *
 * Same model as psych-shift-claims-nov9-jan3.gs: the page lists the shifts
 * left over after the seniority round, the first person to claim one gets it
 * (atomic via LockService), it disappears from the list, and the claimer +
 * Dr. Harrington get a confirmation email.
 *
 * CLAIM WINDOW: the page closes itself at CLOSES_AT (22:00 Mon Oct 12 2026).
 * After that it shows a "closed" notice and refuses claims — no action needed.
 * To reopen later for moonlighting residents: set AUDIENCE = 'RESIDENT',
 * set CLOSES_AT to the new cutoff, save, then Deploy > Manage deployments >
 * pencil > Version: New version > Deploy (same URL).
 *
 * setup() is FIRST in this file on purpose (the editor's Run button runs the
 * first function). It is idempotent — safe to re-run, never clears data.
 * To offer another shift later: add a row to the Shifts tab
 * (next ID, label, OPEN) — no redeploy needed.
 */
const ADMIN_EMAIL = 'harringtonmike@hotmail.com';
const TZ = 'America/Winnipeg';
const SHIFTS_TAB = 'Shifts';
const PROP_SHEET_ID = 'PA_LEFTOVER_NOV_DEC_SHEET_ID';

const AUDIENCE = 'PA';                                  // 'PA' or 'RESIDENT'
const CLOSES_AT = new Date('2026-10-12T22:00:00-05:00'); // 22:00 Mon Oct 12, Winnipeg

/* The 18 shifts nobody was assigned in the Nov–Dec seniority round. */
const INITIAL_SHIFTS = [
  ['Sun Nov 1 — Day (0800–1800)'],
  ['Sat Nov 7 — Day (0800–1800)'],
  ['Sun Nov 15 — Evening (1730–2230)'],
  ['Wed Nov 18 — Evening (1730–2230)'],
  ['Fri Nov 27 — Evening (1730–2230)'],
  ['Sat Nov 28 — Day (0800–1800)'],
  ['Sun Dec 6 — Day (0800–1800)'],
  ['Sun Dec 6 — Evening (1730–2230)'],
  ['Sun Dec 13 — Evening (1730–2230)'],
  ['Fri Dec 18 — Evening (1730–2230)'],
  ['Sat Dec 19 — Day (0800–1800)'],
  ['Sat Dec 19 — Evening (1730–2230)'],
  ['Sun Dec 20 — Day (0800–1800)'],
  ['Wed Dec 23 — Evening (1730–2230)'],
  ['Thu Dec 24 — Evening (1730–2230) (Christmas Eve)'],
  ['Fri Dec 25 — Day (0800–1800) (Christmas Day)'],
  ['Fri Dec 25 — Evening (1730–2230) (Christmas Day)'],
  ['Sat Dec 26 — Evening (1730–2230)']
];

function setup() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty(PROP_SHEET_ID);
  let ss;
  if (id) {
    ss = SpreadsheetApp.openById(id);
  } else {
    ss = SpreadsheetApp.create('PVCS PA Leftover Shift Claims — Nov-Dec 2026');
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
    sh.setColumnWidth(2, 340); sh.setColumnWidth(4, 160); sh.setColumnWidth(5, 220); sh.setColumnWidth(6, 150);
    sh.setFrozenRows(1);
    const rows = INITIAL_SHIFTS.map(function (s, i) { return [i + 1, s[0], 'OPEN', '', '', '']; });
    sh.getRange(2, 1, rows.length, 6).setValues(rows);
  }
  Logger.log('sheet=' + ss.getUrl());
  Logger.log('webapp=' + (ScriptApp.getService().getUrl() || 'not deployed yet'));
}

function isClosed_() { return new Date() >= CLOSES_AT; }
function closesText_() { return Utilities.formatDate(CLOSES_AT, TZ, "HH:mm 'on' EEEE, MMMM d"); }
function who_() { return AUDIENCE === 'RESIDENT' ? 'moonlighting residents' : 'PAs'; }

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

function getState() {
  return { closed: isClosed_(), closes: closesText_(), who: who_(), shifts: isClosed_() ? [] : getOpenShifts_() };
}

function claimShift(id, name, email) {
  if (isClosed_()) {
    return { ok: false, closed: true, closes: closesText_(), msg: 'Claiming closed at ' + closesText_() + '.', shifts: [] };
  }
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

function doGet() {
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
.closed{background:#fff8e1;border:1px solid #f1d58a;padding:14px;border-radius:8px}
</style></head><body><div class="card">
<h1>PVCS &mdash; Open Shifts, November&ndash;December 2026</h1>
<p class="sub" id="sub">Loading&hellip;</p>
<div id="msg"></div><div id="list"><p class="empty">Loading&hellip;</p></div>
</div>
<script>
var open=null,last=[];
function esc(s){return String(s).replace(/[&<>"]/g,function(c){return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c];});}
function msg(t,ok){document.getElementById('msg').innerHTML = t ? ('<div class="msg '+(ok?'ok':'bad')+'">'+esc(t)+'</div>') : '';}
function showClosed(closes){
  document.getElementById('list').innerHTML = '<div class="closed"><b>Claiming is closed.</b> This round closed at '+esc(closes)+'. Remaining shifts will be offered again shortly.</div>';
}
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
  google.script.run.withSuccessHandler(function(r){
    open=null; msg(r.msg, r.ok);
    if (r.closed) showClosed(r.closes); else render(r.shifts);
  }).claimShift(id, n, e);
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
google.script.run.withSuccessHandler(function(st){
  document.getElementById('sub').innerHTML = 'Open to '+esc(st.who)+'. First come, first served &mdash; a shift disappears as soon as someone claims it. Claim as many as you like, one at a time; you will get a confirmation email for each. <b>Claiming closes at '+esc(st.closes)+'.</b>';
  if (st.closed) showClosed(st.closes); else render(st.shifts);
}).getState();
</scr` + `ipt></body></html>`;
  return HtmlService.createHtmlOutput(html)
    .setTitle('PVCS — Open PA Shifts')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}
