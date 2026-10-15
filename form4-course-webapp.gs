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
const PASS_MARK = 4;                           // out of 5
const COURSE_TITLE = 'Physician Assistant/Clinical Assistant Form 4 Certification Course';
const SIGNATORY = 'Dr. Ogo Chukwujama';
const SIGNATORY_TITLE = 'Chief Provincial Psychiatrist, Manitoba';
// Signature image (black ink, transparent PNG), shown above the signatory line.
const SIGNATURE_PNG_B64 = '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';
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
        '<p>Provincial Virtual Crisis Service</p>' + certificateHtml_(L, score, total, certId, when, 'cid:sig'),
      inlineImages: { sig: sigBlob_() },
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
function sigBlob_() { return Utilities.newBlob(Utilities.base64Decode(SIGNATURE_PNG_B64), 'image/png', 'signature.png'); }

// sigSrc: 'cid:sig' for email bodies (Gmail blocks data: images); default data URI for the PDF.
function certificateHtml_(L, score, total, certId, when, sigSrc) {
  const sig = sigSrc || ('data:image/png;base64,' + SIGNATURE_PNG_B64);
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
    '<td style="text-align:left;vertical-align:bottom"><div style="width:300px;height:56px;text-align:left"><img src="' + sig + '" alt="Signature" height="52" style="height:52px;margin:0 0 2px 10px"></div><div style="border-top:1px solid #4a5568;padding-top:6px;width:300px">' + SIGNATORY + '<br>' + SIGNATORY_TITLE + '</div></td>' +
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
    htmlBody: '<p>Sample of the certificate learners receive. Results sheet: ' + sh.getParent().getUrl() + '</p>' + certificateHtml_(L, 5, 5, certId, now, 'cid:sig'),
    inlineImages: { sig: sigBlob_() },
    attachments: [certificatePdf_(L, 5, 5, certId, now)]
  });
  Logger.log('Results sheet: ' + sh.getParent().getUrl());
}
