/**
 * Re-issue Form 4 course certificates in the current format (no CPSM number).
 *
 * Add this as a SEPARATE file in the same Apps Script project as Code.gs
 * (Files + → Script → name it "Reissue"), paste this in, save.
 * It reuses certificatePdf_ / certificateHtml_ / sigBlob_ from Code.gs, and keeps
 * each person's ORIGINAL certificate ID, completion date and score.
 *
 * 1. Leave DRY_RUN = true, select reissueCertificates, Run. You get ONE sample
 *    email (the first person's certificate) and the Execution log lists who
 *    would be sent one. Nobody else is emailed.
 * 2. Set DRY_RUN = false, save, Run again. Each person below gets their
 *    updated certificate. Running it twice sends twice.
 */
const DRY_RUN = true;

// Certificate IDs to re-issue (from the Results sheet). Edit as needed.
const REISSUE_IDS = [
  'PVCS-F4-20261008-3EDF7F', // Olivia Coneys
  'PVCS-F4-20261016-AFA9EA', // Daniel Fillion (his later attempt; Oct 8 duplicate skipped)
  'PVCS-F4-20261022-D47706', // Kylee Barnabe
  'PVCS-F4-20261023-DB02EF'  // Dana Skaritko (sheet email has a typo — corrected below)
];

// Optional: send to a different address than the one in the sheet (e.g. a typo).
const EMAIL_OVERRIDES = {
  'PVCS-F4-20261023-DB02EF': 'dskaritko2@sharedhealthmb.ca' // Dana Skaritko (sheet has sharedheathmb.ca)
};

function reissueCertificates() {
  const data = getSheet_().getDataRange().getValues();
  const h = data[0].map(String);
  const col = function (name) { const i = h.indexOf(name); if (i < 0) throw new Error('Missing column: ' + name); return i; };
  const c = { ts: col('Timestamp'), first: col('First name'), last: col('Last name'), role: col('Role'), email: col('Email'),
              score: col('Score'), total: col('Out of'), result: col('Result'), id: col('Certificate ID') };
  const cap = function (s) { s = String(s).trim(); return s === s.toLowerCase() ? s.replace(/\b\w/g, function (m) { return m.toUpperCase(); }) : s; };

  const found = {};
  data.slice(1).forEach(function (r) { if (REISSUE_IDS.indexOf(String(r[c.id])) >= 0 && r[c.result] === 'PASS') found[String(r[c.id])] = r; });
  const missing = REISSUE_IDS.filter(function (id) { return !found[id]; });
  if (missing.length) Logger.log('NOT FOUND / not a pass: ' + missing.join(', '));

  let sent = 0;
  REISSUE_IDS.forEach(function (id) {
    const r = found[id]; if (!r) return;
    const L = { first: cap(r[c.first]), last: cap(r[c.last]), role: String(r[c.role]), email: EMAIL_OVERRIDES[id] || String(r[c.email]).trim() };
    const when = r[c.ts] instanceof Date ? r[c.ts] : new Date(r[c.ts]);
    const score = Number(r[c.score]), total = Number(r[c.total]);
    const to = DRY_RUN ? ADMIN_EMAILS[0] : L.email;
    Logger.log((DRY_RUN ? '[dry run] would send to ' : 'Sending to ') + L.email + ' — ' + L.first + ' ' + L.last + ' (' + id + ')');
    if (DRY_RUN && sent > 0) return;
    MailApp.sendEmail({
      to: to,
      subject: (DRY_RUN ? '[SAMPLE] ' : '') + (EMAIL_OVERRIDES[id] ? 'Certificate — ' : 'Updated certificate — ') + COURSE_TITLE,
      name: 'PVCS — Provincial Virtual Crisis Service',
      replyTo: ADMIN_EMAILS[0],
      htmlBody:
        '<p>Dear ' + esc_(L.first) + ',</p>' +
        '<p>Thank you for completing the <strong>' + COURSE_TITLE + '</strong>. ' +
        (EMAIL_OVERRIDES[id]
          ? 'The email address entered when you took the course contained a typo, so your certificate may not have reached you. ' +
            'Your certificate is attached (Certificate ID ' + id + '). Please keep it for your records.</p>'
          : 'We have updated the certificate format, and your updated certificate is attached. It replaces the one you received ' +
            'earlier; your completion date, score and certificate ID (' + id + ') are unchanged.</p>') +
        '<p>Provincial Virtual Crisis Service</p>' + certificateHtml_(L, score, total, id, when, 'cid:sig'),
      inlineImages: { sig: sigBlob_() },
      attachments: [certificatePdf_(L, score, total, id, when)]
    });
    sent++;
  });
  Logger.log((DRY_RUN ? 'Dry run: 1 sample sent to ' + ADMIN_EMAILS[0] + '. ' : 'Done: ' + sent + ' sent. ') + Object.keys(found).length + ' certificate(s) matched.');
}
