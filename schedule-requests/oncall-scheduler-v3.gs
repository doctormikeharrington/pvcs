/**
 * ============================================================================
 *  ON-CALL SCHEDULE GENERATOR v3  (bound to the v3 BRANCHED responses sheet)
 * ============================================================================
 *
 *  INSTALL: open the v2 responses Google Sheet > Extensions > Apps Script,
 *  paste this file, Save, reload the sheet. Menu: "On-Call Scheduler".
 *
 *  RULES IMPLEMENTED:
 *   - Priority comes from the "Roster" tab (Name | High/Low). Highs are
 *     shuffled first, Lows shuffled below them; the combined order is used
 *     for every stage. Respondents never see priority.
 *   - Stage 1: FULL WEEKS  (optimal first, then top up to maximum)
 *     Stage 2: WEEKENDS    (until weekend maximum)
 *     Stage 3: MON–FRI     (optimal first, then maximum; counts as weeks)
 *     Day and Evening are allocated together in each stage (interleaved).
 *   - Mon–Fri blocks count toward full-week optimal/maximum.
 *   - A full week counts as ONE weekend toward the weekend total, but the
 *     weekend maximum is IGNORED while full weeks are assigned — it only
 *     limits Stage 2 (weekend blocks).
 *   - No day-overlap between a person's Day and Evening assignments.
 *   - No one gets two consecutive full weeks (any shift combination).
 *   - Only ranked periods are ever assigned; people are skipped when their
 *     ranked choices are gone.
 *   - Output: "Schedule" tab + list of unfilled shifts + review flags.
 *     Sample runs go to "Schedule (TEST)".
 * ============================================================================
 */

var NWEEKS_ = 6;

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
  if (!sheet) { SpreadsheetApp.getUi().alert('Could not find a form-responses tab.'); return; }
  var responses = readResponses_(sheet);
  if (!responses.length) { SpreadsheetApp.getUi().alert('No responses found yet.'); return; }
  writeSchedule_(ss, 'Schedule', responses, readRoster_(ss), getLabels_(ss));
  SpreadsheetApp.getUi().alert('Schedule generated on the "Schedule" tab.');
}

function runSampleTest() {
  var ss = SpreadsheetApp.getActive();
  writeSchedule_(ss, 'Schedule (TEST)', sampleResponses_(), readRoster_(ss), getLabels_(ss));
  SpreadsheetApp.getUi().alert('Sample schedule written to the "Schedule (TEST)" tab.');
}


// ------------------------------ inputs -------------------------------------

function readRoster_(ss) {
  var sh = ss.getSheetByName('Roster');
  var map = {};
  if (!sh) return map;
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
      var ch = function () { return items[i].asListItem().getChoices().map(function (c) { return c.getValue(); }); };
      if (t.indexOf('DAY FULL WEEK — 1st') === 0) labels.full = ch();
      if (t.indexOf('DAY WEEKEND — 1st') === 0) labels.wkend = ch();
      if (t.indexOf('DAY MON-FRI — 1st') === 0) labels.wkday = ch();
    }
  } catch (e) {}
  for (var w = 0; w < NWEEKS_; w++) {
    if (!labels.full[w]) labels.full[w] = 'Week ' + (w + 1);
    if (!labels.wkend[w]) labels.wkend[w] = 'Weekend ' + (w + 1);
    if (!labels.wkday[w]) labels.wkday[w] = 'Weekdays ' + (w + 1);
  }
  return labels;
}

/** Parse "Week 3: ..." / "Weekend 3: ..." / "Weekdays 3: ..." -> index 2. */
function blockIndex_(v) {
  var m = String(v).match(/(\d+)/);
  return m ? parseInt(m[1], 10) - 1 : -1;
}

function readResponses_(sheet) {
  var data = sheet.getDataRange().getValues();
  var headers = data[0].map(function (h) { return String(h).trim(); });
  // The "Both" path duplicates the DAY sections, so DAY headers can appear
  // twice. Collect ALL matching columns and use the first non-empty value.
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
  var cName = colsAll(function (h) { return h === 'Name'; });
  var cShift = colsAll(function (h) { return h.indexOf('Which shifts') !== -1; });
  var cDayOpt = colsAll(function (h) { return /OPTIMAL/.test(h) && /DAY/.test(h) && /Full Weeks/.test(h); });
  var cDayMax = colsAll(function (h) { return /MAXIMUM/.test(h) && /DAY/.test(h) && /Full Weeks/.test(h); });
  var cEveOpt = colsAll(function (h) { return /OPTIMAL/.test(h) && /EVENING/.test(h) && /Full Weeks/.test(h); });
  var cEveMax = colsAll(function (h) { return /MAXIMUM/.test(h) && /EVENING/.test(h) && /Full Weeks/.test(h); });
  var cWkends = colsAll(function (h) { return /WEEKENDS/.test(h) && /MAXIMUM/.test(h); });
  var rc = {
    Dfull: rankCols('DAY FULL WEEK'), Efull: rankCols('EVENING FULL WEEK'),
    Dwkend: rankCols('DAY WEEKEND'), Ewkend: rankCols('EVENING WEEKEND'),
    Dwkday: rankCols('DAY MON-FRI'), Ewkday: rankCols('EVENING MON-FRI')
  };
  function picks(row, cols) {
    var seen = {}, out = [], dups = [];
    cols.forEach(function (c) {
      var v = String(row[c] || '').trim();
      if (!v) return;
      var w = blockIndex_(v);
      if (w < 0) return;
      if (seen[w]) dups.push(v); else { seen[w] = true; out.push(w); }
    });
    return { list: out, dups: dups };
  }
  var out = [];
  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    if (!cName.length || !firstVal(row, cName)) continue;
    out.push({
      name: firstVal(row, cName),
      shift: firstVal(row, cShift),
      dayOpt: toInt_(firstVal(row, cDayOpt)), dayMax: toInt_(firstVal(row, cDayMax)),
      eveOpt: toInt_(firstVal(row, cEveOpt)), eveMax: toInt_(firstVal(row, cEveMax)),
      maxWkends: toInt_(firstVal(row, cWkends)),
      Dfull: picks(row, rc.Dfull), Efull: picks(row, rc.Efull),
      Dwkend: picks(row, rc.Dwkend), Ewkend: picks(row, rc.Ewkend),
      Dwkday: picks(row, rc.Dwkday), Ewkday: picks(row, rc.Ewkday)
    });
  }
  return out;
}

function toInt_(v) { var n = parseInt(v, 10); return isNaN(n) ? 0 : n; }


// ---------------------------- allocation -----------------------------------

function allocateV2_(responses, roster, seed) {
  var people = {};
  responses.forEach(function (p) { people[p.name] = p; });  // last submission wins

  var names = Object.keys(people);
  var hi = names.filter(function (n) { return roster[n] === 'High'; });
  var lo = names.filter(function (n) { return roster[n] !== 'High'; });
  shuffle_(hi, mulberry32_(seed));
  shuffle_(lo, mulberry32_(seed + 1));
  var order = hi.concat(lo);

  var slots = { D: emptySlots_(), E: emptySlots_() };
  var wkCount = {}, wkendCount = {};
  names.forEach(function (n) { wkCount[n] = { D: 0, E: 0 }; wkendCount[n] = 0; });

  function optmax(n, s, key) {
    var p = people[n];
    if (s === 'D') return key === 'opt' ? p.dayOpt : Math.max(p.dayMax, p.dayOpt);
    return key === 'opt' ? p.eveOpt : Math.max(p.eveMax, p.eveOpt);
  }
  function inShift(n, s) {
    var sh = people[n].shift;
    return s === 'D' ? (sh === 'Day' || sh === 'Both') : (sh === 'Evening' || sh === 'Both');
  }
  function holds(n, s, b, w) { return slots[s][b][w] === n; }
  function other(s) { return s === 'D' ? 'E' : 'D'; }

  function canFull(n, s, w, key) {
    if (slots[s].full[w] !== null) return false;
    if (wkCount[n][s] >= optmax(n, s, key)) return false;
    // NOTE: the weekend cap is deliberately NOT checked for full weeks.
    // Full weeks still count 1 weekend each; the cap only limits Stage 2.
    var o = other(s);
    if (holds(n, o, 'full', w) || holds(n, o, 'wkend', w) || holds(n, o, 'wkday', w)) return false;
    for (var s2 in slots) for (var d = -1; d <= 1; d += 2) {
      var w2 = w + d;
      if (w2 >= 0 && w2 < NWEEKS_ && holds(n, s2, 'full', w2)) return false;
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
  function canWkday(n, s, w, key) {
    if (slots[s].full[w] !== null || slots[s].wkday[w] !== null) return false;
    if (wkCount[n][s] >= optmax(n, s, key)) return false;
    var o = other(s);
    if (holds(n, o, 'full', w) || holds(n, o, 'wkday', w)) return false;
    return true;
  }
  function ranksFor(n, s, b) {
    var k = (s === 'D' ? 'D' : 'E') + b;
    return people[n][k] ? people[n][k].list : [];
  }
  function pass(b, s, key) {
    var progress = false;
    order.forEach(function (n) {
      if (!inShift(n, s)) return;
      var ranks = ranksFor(n, s, b);
      for (var i = 0; i < ranks.length; i++) {
        var w = ranks[i], ok;
        if (b === 'full') ok = canFull(n, s, w, key);
        else if (b === 'wkend') ok = canWkend(n, s, w);
        else ok = canWkday(n, s, w, key);
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
  [['full', ['opt', 'max']], ['wkend', ['max']], ['wkday', ['opt', 'max']]].forEach(function (stage) {
    stage[1].forEach(function (key) {
      var progress = true;
      while (progress) {
        var d = pass(stage[0], 'D', key);
        var e = pass(stage[0], 'E', key);
        progress = d || e;
      }
    });
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
  responses.forEach(function (p) {
    ['Dfull', 'Efull', 'Dwkend', 'Ewkend', 'Dwkday', 'Ewkday'].forEach(function (k) {
      if (p[k] && p[k].dups.length) flags.push(p.name + ': duplicate pick(s) in ' + k + ' — ' + p[k].dups.join(', '));
    });
    if (!(p.name in roster)) flags.push(p.name + ': not on the Roster tab — treated as Low priority.');
  });

  var seed = Math.floor(Math.random() * 2147483647);
  var r = allocateV2_(responses, roster, seed);

  function cell(s, b, w) {
    if (r.slots[s].full[w]) return b === 'full' ? r.slots[s].full[w] : '(covered by full week)';
    if (b === 'full') return '—';
    return r.slots[s][b][w] || '(OPEN)';
  }

  var sh = ss.getSheetByName(tabName) || ss.insertSheet(tabName);
  sh.clear();
  var rows = [];
  rows.push(['On-Call Schedule — generated ' + new Date().toLocaleString(), '', '', '', '', '', '']);
  rows.push(['Random seed: ' + seed + '   Priority order used (High first): ' + r.order.join(', '), '', '', '', '', '', '']);
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
  rows.push(['Person', 'Priority', 'DAY weeks (assigned/opt/max)', 'EVE weeks (assigned/opt/max)', 'Weekends (assigned/max)']);
  var seen = {};
  responses.forEach(function (p) {
    if (seen[p.name]) return; seen[p.name] = true;
    rows.push([p.name, roster[p.name] || 'Low',
      r.wkCount[p.name].D + ' / ' + p.dayOpt + ' / ' + p.dayMax,
      r.wkCount[p.name].E + ' / ' + p.eveOpt + ' / ' + p.eveMax,
      r.wkendCount[p.name] + ' / ' + p.maxWkends]);
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
  return [
    { name: 'James Bolton', shift: 'Both', dayOpt: 2, dayMax: 3, eveOpt: 1, eveMax: 2, maxWkends: 3,
      Dfull: pick([0, 1, 2, 3]), Efull: pick([4, 5]), Dwkend: pick([0, 1]), Ewkend: pick([]),
      Dwkday: pick([4, 5]), Ewkday: pick([2, 3]) },
    { name: 'Michael Harrington', shift: 'Day', dayOpt: 3, dayMax: 4, eveOpt: 0, eveMax: 0, maxWkends: 2,
      Dfull: pick([0, 2, 4]), Efull: pick([]), Dwkend: pick([2, 3]), Ewkend: pick([]),
      Dwkday: pick([1, 3, 5]), Ewkday: pick([]) },
    { name: 'Jitender Sareen', shift: 'Both', dayOpt: 1, dayMax: 2, eveOpt: 1, eveMax: 1, maxWkends: 2,
      Dfull: pick([0, 1]), Efull: pick([0, 1]), Dwkend: pick([4, 5]), Ewkend: pick([]),
      Dwkday: pick([]), Ewkday: pick([0, 1]) },
    { name: 'Christopher Classen', shift: 'Evening', dayOpt: 0, dayMax: 0, eveOpt: 2, eveMax: 3, maxWkends: 1,
      Dfull: pick([]), Efull: pick([0, 1, 2]), Dwkend: pick([]), Ewkend: pick([3]),
      Dwkday: pick([]), Ewkday: pick([3, 4, 5]) },
    { name: 'Antonio Paletta', shift: 'Day', dayOpt: 2, dayMax: 2, eveOpt: 0, eveMax: 0, maxWkends: 0,
      Dfull: pick([1, 2, 3]), Efull: pick([]), Dwkend: pick([]), Ewkend: pick([]),
      Dwkday: pick([0, 1, 2, 3, 4, 5]), Ewkday: pick([]) },
    { name: 'Eunice Gill', shift: 'Evening', dayOpt: 0, dayMax: 0, eveOpt: 2, eveMax: 2, maxWkends: 6,
      Dfull: pick([]), Efull: pick([2, 3, 4, 5]), Dwkend: pick([]), Ewkend: pick([0, 1, 2, 3, 4, 5]),
      Dwkday: pick([]), Ewkday: pick([]) }
  ];
}
