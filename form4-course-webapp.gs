/**
 * PVCS — Physician Assistant/Clinical Assistant Form 4 Certification Course
 * Results web app (Google Apps Script).
 *
 * Receives a learner's test answers from pvcsmanitoba.ca/form4course/,
 * re-scores them here (authoritative), logs every attempt to a Google Sheet,
 * emails the learner their result (with a PDF certificate if they pass),
 * and notifies the course administrators of every attempt (pass or fail).
 *
 * SETUP (one time):
 *   1. script.google.com → New project → paste this file → save.
 *   2. Select function `setup` → Run → authorize. It creates the results sheet
 *      and sends a sample certificate to ADMIN_EMAILS[0] so you can see it.
 *   3. Deploy → New deployment → type: Web app → Execute as: Me →
 *      Who has access: Anyone → Deploy. Copy the /exec URL.
 *   4. Paste that URL into RESULTS_ENDPOINT in source/form4course/index.html,
 *      then run update-site.command.
 *   After any later code change: Deploy → Manage deployments → pencil →
 *   Version: New version → Deploy (the /exec URL stays the same).
 */

// ---- CONFIG -------------------------------------------------------------
const ADMIN_EMAILS = ['doctormikeharrington@gmail.com', 'jbolton@hsc.mb.ca'];
const NOTIFY_ADMINS_ON_FAIL = true;            // notify admins of failed attempts as well as passes
const PASS_MARK = 3;                           // out of 5
const COURSE_TITLE = 'Physician Assistant/Clinical Assistant Form 4 Certification Course';
const SIGNATORY = 'Dr. Ogo Chukwujama';
const SIGNATORY_TITLE = 'Chief Provincial Psychiatrist, Manitoba';
const TZ = 'America/Winnipeg';

// Answer key: index of the correct option (0 = A) — must match the course page.
const ANSWER_KEY = { q1: 1, q2: 2, q3: 1, q4: 2, q5: 1 };
// -------------------------------------------------------------------------

function doGet() {
  return json_({ ok: true, service: 'PVCS Form 4 course results' });
}

function doPost(e) {
  try {
    const d = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const L = {
      first: clean_(d.first, 60), last: clean_(d.last, 60), role: clean_(d.role, 40),
      cpsm: clean_(d.cpsm, 20), email: clean_(d.email, 120).toLowerCase()
    };
    if (!L.first || !L.last) return json_({ ok: false, error: 'missing name' });
    if (['Physician Assistant', 'Clinical Assistant'].indexOf(L.role) < 0) return json_({ ok: false, error: 'invalid role' });
    if (!/^[A-Za-z0-9\-]{2,20}$/.test(L.cpsm)) return json_({ ok: false, error: 'invalid CPSM number' });
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(L.email)) return json_({ ok: false, error: 'invalid email' });

    const answers = d.answers || {};
    const keys = Object.keys(ANSWER_KEY);
    let score = 0;
    keys.forEach(function (k) { if (Number(answers[k]) === ANSWER_KEY[k]) score++; });
    const total = keys.length;
    const passed = score >= PASS_MARK;
    const now = new Date();
    const certId = passed ? makeCertId_(now) : '';

    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      getSheet_().appendRow([
        now, L.first, L.last, L.role, L.cpsm, L.email, score, total,
        passed ? 'PASS' : 'FAIL', certId, JSON.stringify(answers), clean_(d.version, 20)
      ]);
    } finally { lock.releaseLock(); }

    sendLearnerEmail_(L, score, total, passed, certId, now);
    if (passed || NOTIFY_ADMINS_ON_FAIL) sendAdminEmail_(L, score, total, passed, certId, now);

    return json_({ ok: true, score: score, total: total, passed: passed, certId: certId });
  } catch (err) {
    console.error(err);
    return json_({ ok: false, error: 'server error' });
  }
}

// ---- email ---------------------------------------------------------------
function sendLearnerEmail_(L, score, total, passed, certId, when) {
  const dateStr = Utilities.formatDate(when, TZ, 'MMMM d, yyyy');
  if (passed) {
    const pdf = certificatePdf_(L, score, total, certId, when);
    MailApp.sendEmail({
      to: L.email,
      subject: 'Certificate — ' + COURSE_TITLE,
      name: 'PVCS — Provincial Virtual Crisis Service',
      replyTo: ADMIN_EMAILS[0],
      htmlBody:
        '<p>Dear ' + esc_(L.first) + ',</p>' +
        '<p>Congratulations — you passed the <strong>' + COURSE_TITLE + '</strong> on ' + dateStr +
        ' with a score of <strong>' + score + '/' + total + '</strong>.</p>' +
        '<p>Your certificate is attached (Certificate ID ' + certId + '). Please keep it for your records.</p>' +
        '<p>Provincial Virtual Crisis Service</p>' + certificateHtml_(L, score, total, certId, when),
      attachments: [pdf]
    });
  } else {
    MailApp.sendEmail({
      to: L.email,
      subject: 'Result — ' + COURSE_TITLE,
      name: 'PVCS — Provincial Virtual Crisis Service',
      replyTo: ADMIN_EMAILS[0],
      htmlBody:
        '<p>Dear ' + esc_(L.first) + ',</p>' +
        '<p>Thank you for completing the <strong>' + COURSE_TITLE + '</strong> test on ' + dateStr +
        '. Your score was <strong>' + score + '/' + total + '</strong>; a score of ' + PASS_MARK + '/' + total +
        ' is needed to pass.</p><p>Please review the course material and take the test again.</p>' +
        '<p>Provincial Virtual Crisis Service</p>'
    });
  }
}

function sendAdminEmail_(L, score, total, passed, certId, when) {
  const name = L.first + ' ' + L.last;
  MailApp.sendEmail({
    to: ADMIN_EMAILS.join(','),
    subject: (passed ? 'Passed: ' : 'Not passed: ') + name + ' (' + L.role + ') — Form 4 Certification Course',
    htmlBody:
      '<p><strong>' + esc_(name) + '</strong> has ' + (passed ? 'successfully completed' : 'attempted') +
      ' the ' + COURSE_TITLE + '.</p>' +
      '<table cellpadding="4" style="border-collapse:collapse;font-family:Arial,sans-serif;font-size:14px">' +
      row_('Role', L.role) + row_('CPSM registration no.', L.cpsm) + row_('Email', L.email) +
      row_('Score', score + '/' + total + (passed ? ' — PASS' : ' — not passed')) +
      (certId ? row_('Certificate ID', certId) : '') +
      row_('Completed', Utilities.formatDate(when, TZ, 'yyyy-MM-dd HH:mm') + ' (Winnipeg)') +
      '</table><p style="font-size:12px;color:#666">All attempts are logged in the "PVCS Form 4 Course Results" sheet: ' +
      getSheet_().getParent().getUrl() + '</p>'
  });
}

function row_(k, v) { return '<tr><td style="color:#555">' + esc_(k) + '</td><td><strong>' + esc_(v) + '</strong></td></tr>'; }

// ---- certificate ----------------------------------------------------------
function certificateHtml_(L, score, total, certId, when) {
  const dateStr = Utilities.formatDate(when, TZ, 'MMMM d, yyyy');
  return '' +
    '<div style="max-width:640px;margin:24px auto;border:10px solid #1f3a5f;background:#fffdf7;font-family:Georgia,\'Times New Roman\',serif;color:#1a202c">' +
    '<div style="margin:8px;border:2px solid #c9a227;padding:36px 34px 28px;text-align:center">' +
    '<div style="font-size:30px;color:#1f3a5f;margin:0 0 4px">Certificate of Completion</div>' +
    '<div style="font-style:italic;color:#4a5568;margin-bottom:22px">' + COURSE_TITLE + '</div>' +
    '<div style="font-size:15px">This certifies that</div>' +
    '<div style="font-size:28px;font-weight:bold;margin:10px auto 12px;padding-bottom:6px;border-bottom:1px solid #c9a227;display:inline-block">' + esc_(L.first + ' ' + L.last) + '</div>' +
    '<div style="font-size:15px;line-height:1.55;margin:0 auto 26px;max-width:500px">' + esc_(L.role) + ', CPSM Registration No. ' + esc_(L.cpsm) +
    ', has successfully completed the ' + COURSE_TITLE + ' on <em>The Mental Health Act</em> of Manitoba and Form 4 &mdash; Application for Involuntary Psychiatric Assessment, with a score of ' + score + '/' + total + '.</div>' +
    '<table width="100%" style="font-family:Arial,sans-serif;font-size:11px;color:#4a5568;margin-top:18px"><tr>' +
    '<td style="text-align:left;vertical-align:bottom"><div style="border-top:1px solid #4a5568;padding-top:6px;width:300px">' + SIGNATORY + '<br>' + SIGNATORY_TITLE + '</div></td>' +
    '<td style="text-align:right;vertical-align:bottom">Date: ' + dateStr + '<br>Certificate ID: ' + certId + '</td>' +
    '</tr></table></div></div>';
}

function certificatePdf_(L, score, total, certId, when) {
  const html = '<html><body style="margin:0;padding:30px 0">' + certificateHtml_(L, score, total, certId, when) + '</body></html>';
  const fname = 'PVCS Form 4 Certificate - ' + L.last + ', ' + L.first + '.pdf';
  return Utilities.newBlob(html, MimeType.HTML, 'cert.html').getAs(MimeType.PDF).setName(fname);
}

function makeCertId_(when) {
  const rand = Utilities.getUuid().replace(/-/g, '').slice(0, 6).toUpperCase();
  return 'PVCS-F4-' + Utilities.formatDate(when, TZ, 'yyyyMMdd') + '-' + rand;
}

// ---- sheet ------------------------------------------------------------------
function getSheet_() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty('SHEET_ID'), ss;
  if (id) { try { ss = SpreadsheetApp.openById(id); } catch (e) { ss = null; } }
  if (!ss) {
    ss = SpreadsheetApp.create('PVCS Form 4 Course Results');
    props.setProperty('SHEET_ID', ss.getId());
    const sh = ss.getSheets()[0];
    sh.setName('Results');
    sh.appendRow(['Timestamp', 'First name', 'Last name', 'Role', 'CPSM reg no.', 'Email', 'Score', 'Out of', 'Result', 'Certificate ID', 'Answers', 'Course version']);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, 12).setFontWeight('bold');
  }
  return ss.getSheetByName('Results') || ss.getSheets()[0];
}

// ---- helpers ------------------------------------------------------------------
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function clean_(v, max) { return String(v == null ? '' : v).replace(/[\u0000-\u001f<>]/g, '').trim().slice(0, max); }
function esc_(s) { return String(s).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; }); }

/** Run once from the editor: authorizes, creates the results sheet, emails a sample certificate to you. */
function setup() {
  const sh = getSheet_();
  const L = { first: 'Sample', last: 'Learner', role: 'Physician Assistant', cpsm: '00000', email: ADMIN_EMAILS[0] };
  const now = new Date();
  const certId = makeCertId_(now);
  MailApp.sendEmail({
    to: ADMIN_EMAILS[0],
    subject: '[SAMPLE] Certificate — ' + COURSE_TITLE,
    htmlBody: '<p>Sample of the certificate learners receive. Results sheet: ' + sh.getParent().getUrl() + '</p>' + certificateHtml_(L, 5, 5, certId, now),
    attachments: [certificatePdf_(L, 5, 5, certId, now)]
  });
  Logger.log('Results sheet: ' + sh.getParent().getUrl());
}
