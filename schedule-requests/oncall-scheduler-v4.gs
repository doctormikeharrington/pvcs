/**
 * ============================================================================
 *  ON-CALL SCHEDULE GENERATOR v4  (bound to the form's responses spreadsheet)
 *  For the REBUILT Aug 31 – Nov 8, 2026 psychiatrist form (2026-08-04).
 * ============================================================================
 *
 *  INSTALL: open the responses Google Sheet > Extensions > Apps Script,
 *  paste this file, Save, reload the sheet. Menu: "On-Call Scheduler".
 *
 *  FORM QUESTIONS READ (new format):
 *   - "How many Full Weeks (Monday to Sunday) of DAY/EVENING shifts do you want?"
 *   - "What is the MAXIMUM number of WEEKENDS you will work over this period"
 *   - Rankings: DAY/EVENING FULL WEEK, MON-FRI, WEEKEND — 1st..10th choice.
 *
 *  OLD-FORMAT FALLBACK: rows submitted before 2026-08-04 (old columns) are
 *  still scheduled — old OPTIMAL becomes full-weeks-wanted, old week rankings
 *  become full-week rankings, weekend cap treated as UNLIMITED (their old
 *  MAXIMUM answers consented to those weeks). Each such row is flagged.
 *  A resubmission on the new form overrides (last submission wins).
 *
 *  RULES IMPLEMENTED:
 *   - Priority from the "Roster" tab (Name | High/Low). Highs shuffled first,
 *     Lows shuffled below; combined order used for every stage. If there is
 *     no Roster tab everyone is treated equally (and a flag reminds you).
 *   - Stage 1: FULL WEEKS (until full-weeks-wanted)
 *     Stage 2: WEEKENDS   (until weekend maximum)
 *     Stage 3: MON–FRI    (counts toward full-weeks-wanted)
 *     Day and Evening are allocated together in each stage (interleaved).
 *   - HARD WEEKEND CAP: a full week counts as ONE weekend and the weekend
 *     maximum limits full weeks too — no one is assigned more full weeks +
 *     weekend blocks than their stated weekend maximum. People whose
 *     full-weeks-wanted exceeds their weekend max are flagged.
 *   - No one works two consecutive full weeks in ANY shift combination.
 *   - No day-overlap between a person's Day and Evening assignments.
 *   - Only ranked periods are ever assigned; people are skipped when their
 *     ranked choices are gone.
 *   - Output: "Schedule" tab + unfilled shifts + review flags.
 *     Sample runs go to "Schedule (TEST)".
 * ============================================================================
 */

var NWEEKS_ = 10;
var UNCAPPED_ = 999;   // weekend cap used for old-format fallback rows

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('On-Call Scheduler')
    .addItem('Generate schedule (Day & Evening)', 'generateSchedule')
    .addSeparator()
    .addItem('Run test with sample data', 'runSampleTest')
    .addToUi();
}

function generateSchedule() {
  var ss = SpreadsheetApp.getActive();
  var sheet = findResponseSheet_(ss);
  if (!sheet) { alert_('Could not find a form-responses tab.'); return; }
  var responses = readResponses_(sheet);
  if (!responses.length) { alert_('No responses found yet.'); return; }
  writeSchedule_(ss, 'Schedule', responses, readRoster_(ss), getLabels_(ss));
  alert_('Schedule generated on the "Schedule" tab.');
}

function runSampleTest() {
  var ss = SpreadsheetApp.getActive();
  writeSchedule_(ss, 'Schedule (TEST)', sampleResponses_(), readRoster_(ss), getLabels_(ss));
  alert_('Sample schedule written to the "Schedule (TEST)" tab.');
}

/** Alert in the sheet UI; falls back to a log when run from the editor. */
function alert_(msg) {
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { Logger.log(msg); }
}


// ------------------------------ inputs -------------------------------------

function readRoster_(ss) {
  var sh = ss.getSheetByName('Roster');
  var map = null;                       // null = no Roster tab at all
  if (!sh) return map;
  map = {};
  var data = sh.getDataRange().getValues();
  for (var r = 1; r < data.length; r++) {
    var n = String(data[r][0]).trim();
    var p = String(data[r][1]).trim().toLowerCase();
    if (n) map[n] = (p.indexOf('h') === 0) ? 'High' : 'Low';
  }
  return map;
}

function findResponseSheet_(ss) {
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    var h = sheets[i].getRange(1, 1, 1, sheets[i].getLastColumn() || 1).getValues()[0].join('|');
    if (h.indexOf('Which shifts') !== -1) return sheets[i];
  }
  return ss.getSheetByName('Form Responses 1');
}

/** Block labels for output, from the linked form's dropdowns (with fallback). */
function getLabels_(ss) {
  var labels = { full: [], wkend: [], wkday: [] };
  try {
    var form = FormApp.openByUrl(ss.getFormUrl());
    var items = form.getItems(FormApp.ItemType.LIST);
    for (var i = 0; i < items.length; i++) {
      var t = items[i].getTitle();
      var ch = items[i].asListItem().getChoices().map(function (c) { return c.getValue(); });
      if (t.indexOf('DAY FULL WEEK — 1st') === 0) labels.full = ch;
      if (t.indexOf('DAY WEEKEND — 1st') === 0) labels.wkend = ch;
      if (t.indexOf('DAY MON-FRI — 1st') === 0) labels.wkday = ch;
    }
  } catch (e) {}
  for (var w = 0; w < NWEEKS_; w++) {
    if (!labels.full[w]) labels.full[w] = 'Week ' + (w + 1);
    if (!labels.wkend[w]) labels.wkend[w] = 'Weekend ' + (w + 1);
    if (!labels.wkday[w]) labels.wkday[w] = 'Weekdays ' + (w + 1);
  }
  return labels;
}

/** Parse "Week 3: ..." / "Weekend 3: ..." -> index 2. */
function blockIndex_(v) {
  var m = String(v).match(/(\d+)/);
  return m ? parseInt(m[1], 10) - 1 : -1;
}

function readResponses_(sheet) {
  var data = sheet.getDataRange().getValues();
  var headers = data[0].map(function (h) { return String(h).trim(); });

  function colsAll(test) { var out = []; for (var i = 0; i < headers.length; i++) if (test(headers[i])) out.push(i); return out; }
  function firstVal(row, colsArr) {
    for (var i = 0; i < colsArr.length; i++) {
      var v = String(row[colsArr[i]] === undefined ? '' : row[colsArr[i]]).trim();
      if (v !== '') return v;
    }
    return '';
  }
  function rankCols(prefix) {
    var cols = [];
    for (var i = 0; i < headers.length; i++) {
      if (headers[i].indexOf(prefix + ' — ') === 0) {
        var m = headers[i].match(/(\d+)(st|nd|rd|th)/);
        cols.push({ idx: i, ord: m ? parseInt(m[1], 10) : 999 });
      }
    }
    cols.sort(function (a, b) { return a.ord - b.ord; });
    return cols.map(function (c) { return c.idx; });
  }
  // Old-format ranking columns: "DAY SHIFTS — 3rd choice week" / "DAY — 11th choice week"
  function oldRankCols(word) {
    var re = new RegExp('^' + word + '( SHIFTS)? — \\d+(st|nd|rd|th) choice week$');
    var cols = [];
    for (var i = 0; i < headers.length; i++) {
      if (re.test(headers[i])) {
        var m = headers[i].match(/(\d+)(st|nd|rd|th)/);
        cols.push({ idx: i, ord: m ? parseInt(m[1], 10) : 999 });
      }
    }
    cols.sort(function (a, b) { return a.ord - b.ord; });
    return cols.map(function (c) { return c.idx; });
  }

  // NEW-format columns
  var cName   = colsAll(function (h) { return h === 'Name'; });
  var cShift  = colsAll(function (h) { return h.indexOf('Which shifts') !== -1; });
  var cDayWant = colsAll(function (h) { return /Full Weeks/.test(h) && /DAY/.test(h) && /do you want/.test(h); });
  var cEveWant = colsAll(function (h) { return /Full Weeks/.test(h) && /EVENING/.test(h) && /do you want/.test(h); });
  var cWkends  = colsAll(function (h) { return /MAXIMUM number of WEEKENDS/.test(h); });
  var rc = {
    Dfull: rankCols('DAY FULL WEEK'), Efull: rankCols('EVENING FULL WEEK'),
    Dwkend: rankCols('DAY WEEKEND'), Ewkend: rankCols('EVENING WEEKEND'),
    Dwkday: rankCols('DAY MON-FRI'), Ewkday: rankCols('EVENING MON-FRI')
  };
  // OLD-format columns (pre-2026-08-04 responses)
  var cDayOptOld = colsAll(function (h) { return /OPTIMAL/.test(h) && /DAY/.test(h) && /Full Weeks/.test(h); });
  var cEveOptOld = colsAll(function (h) { return /OPTIMAL/.test(h) && /EVENING/.test(h) && /Full Weeks/.test(h); });
  var oldDayRanks = oldRankCols('DAY');
  var oldEveRanks = oldRankCols('EVENING');

  function picks(row, cols) {
    var seen = {}, out = [], dups = [];
    cols.forEach(function (c) {
      var v = String(row[c] || '').trim();
      if (!v) return;
      var w = blockIndex_(v);
      if (w < 0 || w >= NWEEKS_) return;
      if (seen[w]) dups.push(v); else { seen[w] = true; out.push(w); }
    });
    return { list: out, dups: dups };
  }
  function emptyPicks() { return { list: [], dups: [] }; }

  var out = [];
  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    if (!cName.length || !firstVal(row, cName)) continue;

    var isNew = firstVal(row, cDayWant) !== '' || firstVal(row, cEveWant) !== '' ||
                firstVal(row, cWkends) !== '';
    var p;
    if (isNew) {
      p = {
        name: firstVal(row, cName),
        shift: firstVal(row, cShift),
        oldFormat: false,
        dayWant: toInt_(firstVal(row, cDayWant)),
        eveWant: toInt_(firstVal(row, cEveWant)),
        maxWkends: toInt_(firstVal(row, cWkends)),
        Dfull: picks(row, rc.Dfull), Efull: picks(row, rc.Efull),
        Dwkend: picks(row, rc.Dwkend), Ewkend: picks(row, rc.Ewkend),
        Dwkday: picks(row, rc.Dwkday), Ewkday: picks(row, rc.Ewkday)
      };
    } else {
      p = {
        name: firstVal(row, cName),
        shift: firstVal(row, cShift),
        oldFormat: true,
        dayWant: toInt_(firstVal(row, cDayOptOld)),
        eveWant: toInt_(firstVal(row, cEveOptOld)),
        maxWkends: UNCAPPED_,
        Dfull: picks(row, oldDayRanks), Efull: picks(row, oldEveRanks),
        Dwkend: emptyPicks(), Ewkend: emptyPicks(),
        Dwkday: emptyPicks(), Ewkday: emptyPicks()
      };
    }
    out.push(p);
  }
  return out;
}

function toInt_(v) { var n = parseInt(v, 10); return isNaN(n) ? 0 : n; }


// ---------------------------- allocation -----------------------------------

function allocate_(responses, roster, seed) {
  var people = {};
  responses.forEach(function (p) { people[p.name] = p; });  // last submission wins

  var rmap = roster || {};
  var names = Object.keys(people);
  var hi = names.filter(function (n) { return rmap[n] === 'High'; });
  var lo = names.filter(function (n) { return rmap[n] !== 'High'; });
  shuffle_(hi, mulberry32_(seed));
  shuffle_(lo, mulberry32_(seed + 1));
  var order = hi.concat(lo);

  var slots = { D: emptySlots_(), E: emptySlots_() };
  var wkCount = {}, wkendCount = {};
  names.forEach(function (n) { wkCount[n] = { D: 0, E: 0 }; wkendCount[n] = 0; });

  function want(n, s) { return s === 'D' ? people[n].dayWant : people[n].eveWant; }
  function inShift(n, s) {
    var sh = people[n].shift;
    return s === 'D' ? (sh === 'Day' || sh === 'Both') : (sh === 'Evening' || sh === 'Both');
  }
  function holds(n, s, b, w) { return slots[s][b][w] === n; }
  function other(s) { return s === 'D' ? 'E' : 'D'; }

  function canFull(n, s, w) {
    if (slots[s].full[w] !== null) return false;
    if (wkCount[n][s] >= want(n, s)) return false;
    // HARD CAP: a full week consumes one weekend.
    if (wkendCount[n] + 1 > people[n].maxWkends) return false;
    var o = other(s);
    if (holds(n, o, 'full', w) || holds(n, o, 'wkend', w) || holds(n, o, 'wkday', w)) return false;
    // No consecutive full weeks in ANY shift combination.
    var shifts = ['D', 'E'];
    for (var si = 0; si < shifts.length; si++) for (var d = -1; d <= 1; d += 2) {
      var w2 = w + d;
      if (w2 >= 0 && w2 < NWEEKS_ && holds(n, shifts[si], 'full', w2)) return false;
    }
    return true;
  }
  function canWkend(n, s, w) {
    if (slots[s].full[w] !== null || slots[s].wkend[w] !== null) return false;
    if (wkendCount[n] + 1 > people[n].maxWkends) return false;
    var o = other(s);
    if (holds(n, o, 'full', w) || holds(n, o, 'wkend', w)) return false;
    return true;
  }
  function canWkday(n, s, w) {
    if (slots[s].full[w] !== null || slots[s].wkday[w] !== null) return false;
    if (wkCount[n][s] >= want(n, s)) return false;
    var o = other(s);
    if (holds(n, o, 'full', w) || holds(n, o, 'wkday', w)) return false;
    return true;
  }
  function ranksFor(n, s, b) {
    var k = (s === 'D' ? 'D' : 'E') + b;
    return people[n][k] ? people[n][k].list : [];
  }
  function pass(b, s) {
    var progress = false;
    order.forEach(function (n) {
      if (!inShift(n, s)) return;
      var ranks = ranksFor(n, s, b);
      for (var i = 0; i < ranks.length; i++) {
        var w = ranks[i], ok;
        if (b === 'full') ok = canFull(n, s, w);
        else if (b === 'wkend') ok = canWkend(n, s, w);
        else ok = canWkday(n, s, w);
        if (ok) {
          slots[s][b][w] = n;
          if (b === 'full' || b === 'wkday') wkCount[n][s]++;
          if (b === 'full' || b === 'wkend') wkendCount[n]++;
          progress = true; return;
        }
      }
    });
    return progress;
  }

  // Stage order: full weeks -> weekends -> Mon–Fri.
  ['full', 'wkend', 'wkday'].forEach(function (b) {
    var progress = true;
    while (progress) {
      var d = pass(b, 'D');
      var e = pass(b, 'E');
      progress = d || e;
    }
  });

  var unfilled = [];
  ['D', 'E'].forEach(function (s) {
    for (var w = 0; w < NWEEKS_; w++) {
      if (slots[s].full[w] === null) {
        if (slots[s].wkday[w] === null) unfilled.push({ shift: s, b: 'wkday', w: w });
        if (slots[s].wkend[w] === null) unfilled.push({ shift: s, b: 'wkend', w: w });
      }
    }
  });
  return { order: order, hi: hi, slots: slots, wkCount: wkCount, wkendCount: wkendCount, unfilled: unfilled };
}

function emptySlots_() {
  var mk = function () { var a = []; for (var i = 0; i < NWEEKS_; i++) a.push(null); return a; };
  return { full: mk(), wkend: mk(), wkday: mk() };
}

function shuffle_(arr, rand) {
  for (var i = arr.length - 1; i > 0; i--) {
    var j = Math.floor(rand() * (i + 1));
    var t = arr[i]; arr[i] = arr[j]; arr[j] = t;
  }
  return arr;
}

function mulberry32_(a) {
  return function () {
    a |= 0; a = (a + 0x6D2B79F5) | 0;
    var t = Math.imul(a ^ (a >>> 15), 1 | a);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}


// ------------------------------ output --------------------------------------

function writeSchedule_(ss, tabName, responses, roster, labels) {
  var flags = [];
  var latest = {};
  responses.forEach(function (p) { latest[p.name] = p; });   // last submission wins
  Object.keys(latest).forEach(function (n) {
    var p = latest[n];
    ['Dfull', 'Efull', 'Dwkend', 'Ewkend', 'Dwkday', 'Ewkday'].forEach(function (k) {
      if (p[k] && p[k].dups.length) flags.push(p.name + ': duplicate pick(s) in ' + k + ' — ' + p[k].dups.join(', '));
    });
    if (p.oldFormat)
      flags.push(p.name + ': OLD-FORM response — full-weeks-wanted taken from old OPTIMAL, weekend cap treated as unlimited. Collect their weekend max / Mon-Fri / weekend preferences manually.');
    if (!p.oldFormat && p.maxWkends < Math.max(p.dayWant, p.eveWant))
      flags.push(p.name + ': weekend max (' + p.maxWkends + ') is below full-weeks-wanted (Day ' + p.dayWant + ' / Eve ' + p.eveWant + ') — the hard cap limited their full weeks.');
    if (roster && !(p.name in roster)) flags.push(p.name + ': not on the Roster tab — treated as Low priority.');
  });
  if (!roster) flags.push('No "Roster" tab found — everyone treated as equal priority. Add a Roster tab (Name | High/Low) to use priority.');

  var seed = Math.floor(Math.random() * 2147483647);
  var r = allocate_(responses, roster, seed);

  function cell(s, b, w) {
    if (r.slots[s].full[w]) return b === 'full' ? r.slots[s].full[w] : '(covered by full week)';
    if (b === 'full') return '—';
    return r.slots[s][b][w] || '(OPEN)';
  }

  var sh = ss.getSheetByName(tabName) || ss.insertSheet(tabName);
  sh.clear();
  var rows = [];
  rows.push(['On-Call Schedule — generated ' + new Date().toLocaleString(), '', '', '', '', '', '']);
  rows.push(['Random seed: ' + seed + '   Order used (High priority first): ' + r.order.join(', '), '', '', '', '', '', '']);
  rows.push([]);
  rows.push(['Week', 'DAY full week', 'DAY Mon–Fri', 'DAY weekend', 'EVENING full week', 'EVENING Mon–Fri', 'EVENING weekend']);
  for (var w = 0; w < NWEEKS_; w++) {
    rows.push([labels.full[w],
      cell('D', 'full', w), cell('D', 'wkday', w), cell('D', 'wkend', w),
      cell('E', 'full', w), cell('E', 'wkday', w), cell('E', 'wkend', w)]);
  }
  rows.push([]);
  rows.push(['UNFILLED SHIFTS (offer as individual days)']);
  if (!r.unfilled.length) rows.push(['none — everything assigned']);
  r.unfilled.forEach(function (u) {
    var lab = u.b === 'wkday' ? labels.wkday[u.w] : labels.wkend[u.w];
    rows.push([(u.shift === 'D' ? 'DAY — ' : 'EVENING — ') + lab]);
  });
  rows.push([]);
  rows.push(['Person', 'Priority', 'DAY weeks (assigned/wanted)', 'EVE weeks (assigned/wanted)', 'Weekends (assigned/max)']);
  Object.keys(latest).forEach(function (n) {
    var p = latest[n];
    rows.push([p.name, (roster && roster[p.name]) || 'Low',
      r.wkCount[p.name].D + ' / ' + p.dayWant,
      r.wkCount[p.name].E + ' / ' + p.eveWant,
      r.wkendCount[p.name] + ' / ' + (p.maxWkends >= UNCAPPED_ ? 'n/a (old form)' : p.maxWkends)]);
  });
  rows.push([]);
  rows.push(['Flags (please review)']);
  if (!flags.length) flags.push('none');
  flags.forEach(function (f) { rows.push([f]); });

  var width = 7;
  rows.forEach(function (rw) { while (rw.length < width) rw.push(''); });
  sh.getRange(1, 1, rows.length, width).setValues(rows);
  sh.getRange(4, 1, 1, width).setFontWeight('bold');
  sh.setColumnWidth(1, 210);
  for (var c = 2; c <= 7; c++) sh.setColumnWidth(c, 170);
  sh.setFrozenRows(4);
}


// ---------------------------- sample data -----------------------------------

function sampleResponses_() {
  function pick(list) { return { list: list, dups: [] }; }
  function e() { return { list: [], dups: [] }; }
  return [
    { name: 'Dr. Adams', shift: 'Both', oldFormat: false, dayWant: 2, eveWant: 1, maxWkends: 3,
      Dfull: pick([0, 1, 2, 3]), Efull: pick([4, 5]), Dwkend: pick([0, 1]), Ewkend: e(),
      Dwkday: pick([4, 5]), Ewkday: pick([2, 3]) },
    { name: 'Dr. Brar', shift: 'Day', oldFormat: false, dayWant: 3, eveWant: 0, maxWkends: 2,
      Dfull: pick([0, 2, 4, 6, 8]), Efull: e(), Dwkend: pick([2, 3]), Ewkend: e(),
      Dwkday: pick([1, 3, 5, 7, 9]), Ewkday: e() },
    { name: 'Dr. Chen', shift: 'Both', oldFormat: false, dayWant: 1, eveWant: 1, maxWkends: 2,
      Dfull: pick([0, 1]), Efull: pick([0, 1]), Dwkend: pick([4, 5]), Ewkend: e(),
      Dwkday: e(), Ewkday: pick([0, 1]) },
    { name: 'Dr. Diaz', shift: 'Evening', oldFormat: false, dayWant: 0, eveWant: 2, maxWkends: 1,
      Dfull: e(), Efull: pick([0, 1, 2, 6, 7]), Dwkend: e(), Ewkend: pick([3]),
      Dwkday: e(), Ewkday: pick([3, 4, 5]) },
    { name: 'Dr. Evans', shift: 'Day', oldFormat: false, dayWant: 2, eveWant: 0, maxWkends: 0,
      Dfull: pick([1, 2, 3]), Efull: e(), Dwkend: e(), Ewkend: e(),
      Dwkday: pick([0, 1, 2, 3, 4, 5, 6, 7, 8, 9]), Ewkday: e() },
    { name: 'Dr. Foster', shift: 'Evening', oldFormat: false, dayWant: 0, eveWant: 2, maxWkends: 6,
      Dfull: e(), Efull: pick([2, 3, 4, 5, 8, 9]), Dwkend: e(), Ewkend: pick([0, 1, 6, 7]),
      Dwkday: e(), Ewkday: e() },
    // Old-format fallback row (like Okoye/Gill/Kuzenko):
    { name: 'Dr. Gupta', shift: 'Both', oldFormat: true, dayWant: 1, eveWant: 2, maxWkends: UNCAPPED_,
      Dfull: pick([3, 4, 6, 9]), Efull: pick([2, 6, 7, 9]), Dwkend: e(), Ewkend: e(),
      Dwkday: e(), Ewkday: e() }
  ];
}
