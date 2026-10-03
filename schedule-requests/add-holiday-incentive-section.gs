/**
 * addHolidayIncentiveSection
 * ---------------------------
 * Adds a "Christmas & New Year's — special preferences" section to the
 * Nov 9 2026 – Jan 3 2027 psychiatrist shift request form.
 *
 * Section (placed BEFORE the DAY-shift ranking section):
 *   • Intro text: incentive framing (extra priority on another week for
 *     the person who takes the holiday shift).
 *   • Q1  Week 7 — Dec 21–27, 2026 (Christmas week)
 *   • Q2  Week 8 — Dec 28, 2026 – Jan 3, 2027 (New Year's week)
 *   • Header: Alternative — individual holiday days
 *   • Q3–Q6  Dec 25, Dec 26, Dec 31, Jan 1
 *
 * All 6 MCQ items share the same three options and are OPTIONAL.
 * Idempotent — safe to re-run; will not duplicate.
 *
 * Paste this into the Apps Script project 1fgdHD1QxvIRzij4bW_jDXPWJBDt6ztDgpk0z2HXn1A-31gSozOPayDz8
 * (the Aug 31 – Nov 8 psychiatrist form's bound project, already authorized
 * for Forms/Sheets/Drive scopes). Select 'addHolidayIncentiveSection' from
 * the function dropdown and Run.
 */

function addHolidayIncentiveSection() {
  var FORM_ID = '1bvojJ3zQLJTN_1nUxp2vVK3KL6gQuU9aysSG52jtoU8';   // Nov 9 – Jan 3 form
  var SECTION_TITLE = "Christmas & New Year's — special preferences";
  var SECTION_DESC =
    "In recognition that these are days many team members prefer not to work: " +
    "anyone who works one of these shifts will be given additional priority " +
    "for their highest-requested other week on service.";
  var ALT_HEADER = "Alternative — individual holiday days";
  var ALT_HELP =
    "If you would rather answer day by day than by the full week, use these. " +
    "The same priority incentive applies.";
  var OPTIONS = [
    "I would prefer to work",
    "I am willing to work if needed",
    "I would prefer not to work"
  ];
  var QS = [
    { title: "Week 7 — Dec 21–27, 2026 (Christmas week)" },
    { title: "Week 8 — Dec 28, 2026 – Jan 3, 2027 (New Year's week)" },
    { title: "Christmas Day — Fri Dec 25, 2026" },
    { title: "Boxing Day — Sat Dec 26, 2026" },
    { title: "New Year's Eve — Thu Dec 31, 2026" },
    { title: "New Year's Day — Fri Jan 1, 2027" }
  ];

  var form = FormApp.openById(FORM_ID);
  var items = form.getItems();

  // Idempotency: bail if the section header is already present.
  for (var i = 0; i < items.length; i++) {
    if (items[i].getTitle() === SECTION_TITLE) {
      Logger.log('Section already present at index ' + i + ' — nothing to do.');
      return;
    }
  }

  // Anchor: the first item titled "How many Full Weeks..." mentioning DAY.
  // If it sits immediately after a PageBreak, insert BEFORE that PageBreak.
  var anchor = -1;
  for (var j = 0; j < items.length; j++) {
    var t = items[j].getTitle() || '';
    if (t.indexOf('How many Full Weeks') === 0 && t.indexOf('DAY') !== -1) {
      anchor = (j > 0 && items[j - 1].getType() === FormApp.ItemType.PAGE_BREAK) ? (j - 1) : j;
      break;
    }
  }
  if (anchor < 0) {
    throw new Error('Could not find "How many Full Weeks ... DAY ..." question to anchor insertion.');
  }
  Logger.log('Insertion anchor index: ' + anchor);

  // Create items (they are appended to end, then moved into place in order).
  var created = [];
  created.push(form.addPageBreakItem().setTitle(SECTION_TITLE).setHelpText(SECTION_DESC));
  created.push(form.addMultipleChoiceItem().setTitle(QS[0].title).setChoiceValues(OPTIONS).setRequired(false));
  created.push(form.addMultipleChoiceItem().setTitle(QS[1].title).setChoiceValues(OPTIONS).setRequired(false));
  created.push(form.addSectionHeaderItem().setTitle(ALT_HEADER).setHelpText(ALT_HELP));
  created.push(form.addMultipleChoiceItem().setTitle(QS[2].title).setChoiceValues(OPTIONS).setRequired(false));
  created.push(form.addMultipleChoiceItem().setTitle(QS[3].title).setChoiceValues(OPTIONS).setRequired(false));
  created.push(form.addMultipleChoiceItem().setTitle(QS[4].title).setChoiceValues(OPTIONS).setRequired(false));
  created.push(form.addMultipleChoiceItem().setTitle(QS[5].title).setChoiceValues(OPTIONS).setRequired(false));

  // Slide each into position at anchor, anchor+1, anchor+2, ...
  for (var k = 0; k < created.length; k++) {
    form.moveItem(created[k].getIndex(), anchor + k);
  }

  Logger.log('Inserted ' + created.length + ' items starting at index ' + anchor + '.');
}
