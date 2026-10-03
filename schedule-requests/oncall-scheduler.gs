/**
 * ============================================================================
 *  ON-CALL SCHEDULE GENERATOR  (bound to the form's responses spreadsheet)
 * ============================================================================
 *
 *  INSTALL: open the responses Google Sheet > Extensions > Apps Script,
 *  paste this file, Save. Reload the sheet — an "On-Call Scheduler" menu
 *  appears.
 *
 *  HOW IT WORKS (matches the agreed rules):
 *   - Day and Evening are scheduled independently, one person per week each.
 *   - Everyone who asked for a shift is shuffled into a random order
 *     (the random seed is recorded so a run can be reproduced).
 *   - Round-robin, one week per pass: each person in turn receives their
 *     highest still-open ranked week, until they reach their OPTIMAL count.
 *   - If weeks remain open after everyone hits optimal, further passes top
 *     people up toward their MAXIMUM (never exceeding it).
 *   - A person is only ever placed in a week they ranked. If all their
 *     ranked weeks are taken, they are skipped — some weeks may stay open
 *     for you to fill by hand.
 *   - No one is given both the Day and the Evening slot in the same week.
 *   - No one is given the same shift type two weeks in a row (no back-to-back
 *     full weeks of Day; likewise for Evening).
 *   - Duplicate week picks within one person's ranking are flagged.
 *
 *  Output is written to a "Schedule" tab (sample runs go to "Schedule (TEST)").
 * ============================================================================
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('On-Call Scheduler')
    .addItem('Generate schedule (Day & Evening)', 'generateSchedule')
    .addSeparator()
    .addItem('Run test with sample data', 'runSampleTest')
    .addToUi();
}


/** Main entry: build the schedule from the real form responses. */
function generateSchedule() {
  var ss = SpreadsheetApp.getActive();
  var sheet = findResponseSheet_(ss);
  if (!sheet) { SpreadsheetApp.getUi().alert('Could not find a form-responses tab.'); return; }
  var weeks = getWeeksFromForm_(ss);
  var responses = readResponses_(sheet, weeks);
  if (!responses.length) { SpreadsheetApp.getUi().alert('No responses found yet.'); return; }
  writeSchedule_(ss, 'Schedule', responses, weeks);
  SpreadsheetApp.getUi().alert('Schedule generated on the "Schedule" tab.');
}


/** Test run with built-in dummy data; does not touch real responses. */
function runSampleTest() {
  var weeks = [];
  for (var i = 1; i <= 12; i++) weeks.push('Week ' + i + ': (sample dates)');
  var responses = sampleResponses_(weeks);
  writeSchedule_(SpreadsheetApp.getActive(), 'Schedule (TEST)', responses, weeks);
  SpreadsheetApp.getUi().alert('Sample schedule written to the "Schedule (TEST)" tab.');
}


// ----------------------------- reading -------------------------------------

function findResponseSheet_(ss) {
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    var h = sheets[i].getRange(1, 1, 1, sheets[i].getLastColumn() || 1).getValues()[0].join('|');
    if (h.indexOf('Which shifts') !== -1) return sheets[i];   // the form responses tab
  }
  // fall back to a tab literally named like the form responses
  return ss.getSheetByName('Form Responses 1') || ss.getSheets()[0];
}

/** Pull the canonical, ordered week list from the linked form's DAY ranking. */
function getWeeksFromForm_(ss) {
  try {
    var form = FormApp.openByUrl(ss.getFormUrl());
    var items = form.getItems(FormApp.ItemType.LIST);
    for (var i = 0; i < items.length; i++) {
      if (items[i].getTitle().indexOf('DAY') === 0) {
        return items[i].asListItem().getChoices().map(function (c) { return c.getValue(); });
      }
    }
  } catch (e) {}
  return [];
}

function readResponses_(sheet, weeks) {
  var data = sheet.getDataRange().getValues();
  var headers = data[0].map(function (h) { return String(h).trim(); });

  function colIndex(test) { for (var i = 0; i < headers.length; i++) if (test(headers[i])) return i; return -1; }
  function rankCols(prefix) {
    var cols = [];
    for (var i = 0; i < headers.length; i++) {
      var h = headers[i];
      if (h.indexOf(prefix) === 0 && /choice week/i.test(h)) {
        var m = h.match(/(\d+)(st|nd|rd|th)/);
        cols.push({ idx: i, ord: m ? parseInt(m[1], 10) : 999 });
      }
    }
    cols.sort(function (a, b) { return a.ord - b.ord; });
    return cols.map(function (c) { return c.idx; });
  }

  var cName = colIndex(function (h) { return h === 'Name'; });
  var cShift = colIndex(function (h) { return h.indexOf('Which shifts') !== -1; });
  var cDayOpt = colIndex(function (h) { return /OPTIMAL/.test(h) && /DAY/.test(h); });
  var cDayMax = colIndex(function (h) { return /MAXIMUM/.test(h) && /DAY/.test(h); });
  var cEveOpt = colIndex(function (h) { return /OPTIMAL/.test(h) && /EVENING/.test(h); });
  var cEveMax = colIndex(function (h) { return /MAXIMUM/.test(h) && /EVENING/.test(h); });
  var dayRankCols = rankCols('DAY');
  var eveRankCols = rankCols('EVENING');

  var out = [];
  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    if (cName === -1 || !String(row[cName]).trim()) continue;
    out.push({
      name: String(row[cName]).trim(),
      shift: String(row[cShift] || '').trim(),
      dayOpt: toInt_(row[cDayOpt]), dayMax: toInt_(row[cDayMax]),
      eveOpt: toInt_(row[cEveOpt]), eveMax: toInt_(row[cEveMax]),
      dayRanks: pickRanks_(row, dayRankCols),
      eveRanks: pickRanks_(row, eveRankCols)
    });
  }
  return out;
}

function pickRanks_(row, cols) {
  var out = [];
  for (var i = 0; i < cols.length; i++) {
    var v = String(row[cols[i]] || '').trim();
    if (v) out.push(v);
  }
  return out;
}

function toInt_(v) { var n = parseInt(v, 10); return isNaN(n) ? 0 : n; }


// --------------------------- allocation ------------------------------------

function dedupe_(ranks) {
  var seen = {}, out = [], dups = [];
  for (var i = 0; i < ranks.length; i++) {
    if (seen[ranks[i]]) dups.push(ranks[i]); else { seen[ranks[i]] = true; out.push(ranks[i]); }
  }
  return { ranks: out, dups: dups };
}

/**
 * Allocates Day and Evening TOGETHER so a person is never given both the Day
 * and the Evening slot in the same week. Day and Evening each keep their own
 * random order; passes are interleaved (one Day pass, then one Evening pass)
 * and share a "held weeks" map per person so cross-shift clashes are blocked.
 */
function allocateAll_(dayPool, evePool, weeks, seed) {
  var dayOrder = shuffle_(dayPool.slice(), mulberry32_(seed));
  var eveOrder = shuffle_(evePool.slice(), mulberry32_(seed + 1));

  var dayAssigned = {}, eveAssigned = {};
  weeks.forEach(function (w) { dayAssigned[w] = null; eveAssigned[w] = null; });

  var dayCounts = {}, eveCounts = {}, held = {};
  function register(p) { if (!held[p.name]) held[p.name] = {}; }
  dayPool.forEach(function (p) { dayCounts[p.name] = 0; register(p); });
  evePool.forEach(function (p) { eveCounts[p.name] = 0; register(p); });

  var weekIndex = {};
  weeks.forEach(function (w, i) { weekIndex[w] = i; });

  // True if this person already holds the adjacent week in the SAME shift type
  // (blocks back-to-back full weeks of Day, and likewise for Evening).
  function adjacentSameShift(name, w, assigned) {
    var wi = weekIndex[w];
    var prev = weeks[wi - 1], next = weeks[wi + 1];
    return (prev !== undefined && assigned[prev] === name) ||
           (next !== undefined && assigned[next] === name);
  }

  // One round-robin pass over a shift: each person gets at most one new week.
  function pass(order, assigned, counts, key) {
    var progress = false;
    order.forEach(function (p) {
      if (counts[p.name] >= p[key]) return;
      for (var i = 0; i < p.ranks.length; i++) {
        var w = p.ranks[i];
        if (assigned.hasOwnProperty(w) && assigned[w] === null && !held[p.name][w] &&
            !adjacentSameShift(p.name, w, assigned)) {
          assigned[w] = p.name; counts[p.name]++; held[p.name][w] = true;
          progress = true; return;
        }
      }
    });
    return progress;
  }

  ['opt', 'max'].forEach(function (key) {   // optimal first, then top up to max
    var progress = true;
    while (progress && (hasOpen_(dayAssigned) || hasOpen_(eveAssigned))) {
      var movedDay = pass(dayOrder, dayAssigned, dayCounts, key);
      var movedEve = pass(eveOrder, eveAssigned, eveCounts, key);
      progress = movedDay || movedEve;
    }
  });

  return { dayAssigned: dayAssigned, eveAssigned: eveAssigned, dayCounts: dayCounts, eveCounts: eveCounts };
}

function shuffle_(arr, rand) {                // seeded Fisher-Yates
  for (var i = arr.length - 1; i > 0; i--) {
    var j = Math.floor(rand() * (i + 1));
    var t = arr[i]; arr[i] = arr[j]; arr[j] = t;
  }
  return arr;
}

function hasOpen_(assigned) { for (var w in assigned) if (assigned[w] === null) return true; return false; }

/** Small seedable PRNG so runs are reproducible from the recorded seed. */
function mulberry32_(a) {
  return function () {
    a |= 0; a = (a + 0x6D2B79F5) | 0;
    var t = Math.imul(a ^ (a >>> 15), 1 | a);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}


// ---------------------------- writing --------------------------------------

function writeSchedule_(ss, tabName, responses, weeks) {
  var flags = [];
  var dayPool = [], evePool = [];
  responses.forEach(function (p) {
    var d = dedupe_(p.dayRanks), e = dedupe_(p.eveRanks);
    if (d.dups.length) flags.push(p.name + ': duplicate DAY pick(s) — ' + d.dups.join(', '));
    if (e.dups.length) flags.push(p.name + ': duplicate EVENING pick(s) — ' + e.dups.join(', '));
    if ((p.shift === 'Day' || p.shift === 'Both') && (p.dayOpt > 0 || d.ranks.length))
      dayPool.push({ name: p.name, opt: p.dayOpt, max: Math.max(p.dayMax, p.dayOpt), ranks: d.ranks });
    if ((p.shift === 'Evening' || p.shift === 'Both') && (p.eveOpt > 0 || e.ranks.length))
      evePool.push({ name: p.name, opt: p.eveOpt, max: Math.max(p.eveMax, p.eveOpt), ranks: e.ranks });
  });

  var seed = Math.floor(Math.random() * 2147483647);
  var res = allocateAll_(dayPool, evePool, weeks, seed);
  var day = { assigned: res.dayAssigned, counts: res.dayCounts };
  var eve = { assigned: res.eveAssigned, counts: res.eveCounts };

  var sh = ss.getSheetByName(tabName) || ss.insertSheet(tabName);
  sh.clear();
  var rows = [];
  rows.push(['On-Call Schedule — generated ' + new Date().toLocaleString()]);
  rows.push(['Random seed: Day=' + seed + '  Evening=' + (seed + 1) + '   (re-running picks a new seed)']);
  rows.push([]);
  rows.push(['Week', 'Day on-call', 'Evening on-call']);
  weeks.forEach(function (w) {
    rows.push([w, day.assigned[w] || '(OPEN — fill manually)', eve.assigned[w] || '(OPEN — fill manually)']);
  });
  rows.push([]);
  rows.push(['Assigned / Optimal / Max — DAY']);
  dayPool.forEach(function (p) { rows.push([p.name, day.counts[p.name] + ' / ' + p.opt + ' / ' + p.max]); });
  rows.push([]);
  rows.push(['Assigned / Optimal / Max — EVENING']);
  evePool.forEach(function (p) { rows.push([p.name, eve.counts[p.name] + ' / ' + p.opt + ' / ' + p.max]); });
  rows.push([]);
  rows.push(['Flags (please review)']);
  if (!flags.length) flags.push('none');
  flags.forEach(function (f) { rows.push([f]); });

  var width = 3;
  rows.forEach(function (r) { while (r.length < width) r.push(''); });
  sh.getRange(1, 1, rows.length, width).setValues(rows);
  sh.getRange(4, 1, 1, width).setFontWeight('bold');
  sh.setColumnWidth(1, 220); sh.setColumnWidth(2, 200); sh.setColumnWidth(3, 200);
  sh.setFrozenRows(4);
}


// --------------------------- sample data -----------------------------------

function sampleResponses_(w) {
  return [
    { name: 'Dr. Adams',  shift: 'Both',    dayOpt: 2, dayMax: 3, eveOpt: 2, eveMax: 2, dayRanks: [w[0], w[1], w[2], w[3]], eveRanks: [w[0], w[1], w[4], w[5]] },
    { name: 'Dr. Brar',   shift: 'Day',     dayOpt: 3, dayMax: 4, eveOpt: 0, eveMax: 0, dayRanks: [w[0], w[2], w[4], w[6], w[8]], eveRanks: [] },
    { name: 'Dr. Chen',   shift: 'Evening', dayOpt: 0, dayMax: 0, eveOpt: 3, eveMax: 4, dayRanks: [], eveRanks: [w[0], w[1], w[2], w[9], w[10], w[11]] },
    { name: 'Dr. Diaz',   shift: 'Day',     dayOpt: 2, dayMax: 2, eveOpt: 0, eveMax: 0, dayRanks: [w[1], w[3], w[5]], eveRanks: [] },
    { name: 'Dr. Evans',  shift: 'Both',    dayOpt: 1, dayMax: 2, eveOpt: 2, eveMax: 3, dayRanks: [w[0], w[11]], eveRanks: [w[3], w[4], w[5], w[6]] },
    { name: 'Dr. Foster', shift: 'Day',     dayOpt: 2, dayMax: 3, eveOpt: 0, eveMax: 0, dayRanks: [w[7], w[8], w[9], w[10], w[11]], eveRanks: [] },
    { name: 'Dr. Gupta',  shift: 'Evening', dayOpt: 0, dayMax: 0, eveOpt: 2, eveMax: 2, dayRanks: [], eveRanks: [w[7], w[8], w[9]] },
    { name: 'Dr. Hill',   shift: 'Day',     dayOpt: 1, dayMax: 1, eveOpt: 0, eveMax: 0, dayRanks: [w[0]], eveRanks: [] },
    { name: 'Dr. Ito',    shift: 'Both',    dayOpt: 2, dayMax: 2, eveOpt: 1, eveMax: 1, dayRanks: [w[4], w[5], w[6], w[7]], eveRanks: [w[11]] },
    { name: 'Dr. Jones',  shift: 'Evening', dayOpt: 0, dayMax: 0, eveOpt: 3, eveMax: 3, dayRanks: [], eveRanks: [w[0], w[1], w[2]] }
  ];
}
