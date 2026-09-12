/* TAT Raw Data Table -- a Tableau dashboard extension that renders one worksheet as an interactive table:
 * per-column filters (text, and a date range slider for date columns), a table-wide search box, sortable columns,
 * paging, and export to Excel with REAL data types (dates as Excel dates, numbers as numbers, clean headers) or CSV.
 * The export honours whatever is filtered/searched on screen; "All rows" re-reads the full filtered data set from
 * Tableau first (it flips the dashboard's row-cap parameter, reads, and puts it back).
 * Libraries: Tableau Extensions API 1.17 (Tableau's own @tableau/tabextsandbox package), SheetJS 0.18.5 (Apache-2.0).
 * The conversion / shaping / export functions are pure so they can be unit-tested in Node (see selftest.js).
 */
(function (root) {
  'use strict';

  var DEFAULTS = {
    sheet: 'Raw Data Table',        // worksheet on the dashboard to read
    filename: 'TAT Raw Data',       // file name prefix; a timestamp is appended
    dateFormat: 'dd/mm/yyyy',       // Excel number format for date columns
    dateTimeFormat: 'dd/mm/yyyy hh:mm',
    label: 'Excel',             // text on the green export button
    param: 'Rows shown',            // parameter that caps how many rows the worksheet returns
    paramAll: '10000000',           // its "all rows" value
    allRows: 'true',                // default state of the "All rows" checkbox
    pageSize: '100',
    // Tableau's summary data hands columns back as dimensions A-Z then measures, NOT in the worksheet's shelf order.
    // This is the order we want left to right; anything not listed keeps its place after the listed ones.
    columns: 'WO, Product Name, Product Description, Product Category, Deliverable Services, Service Level Clean, Ship to Service Account, WDF Region, WDF Hub, WDB Region, WDB Hub, Transhipment, RTK/RTF, SOT, Status, Net TAT (Days), Target TAT (Days), Parts Delay, WO Delay Reason, Fiscal Year, Display Quarter, WO Created Date, WO Received Date, WO Closed Date'
  };

  // ---------------------------------------------------------------- pure helpers
  var AGG_RE = /^(SUM|AVG|MIN|MAX|CNT|CNTD|COUNT|COUNTD|ATTR|AGG|MEDIAN|STDEV|VAR|TOTAL)\((.*)\)$/i;
  function cleanName(n) { var m = AGG_RE.exec(n); return m ? m[2] : n; }

  // Excel serial for a calendar date (1900 date system; valid for dates after 28 Feb 1900)
  function serial(y, m, d) { return Math.round((Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 86400000); }
  var DT_RE = /^(\d{4})-(\d{2})-(\d{2})(?:[ T](\d{2}):(\d{2})(?::(\d{2}(?:\.\d+)?))?)?/;
  // value = the API's string form ("2023-04-04" / "2023-04-04 13:45:00"); nativeValue = Date (UTC) -- both handled
  function toSerial(value, nativeValue) {
    if (typeof value === 'string') {
      var m = DT_RE.exec(value);
      if (m) {
        var s = serial(+m[1], +m[2], +m[3]);
        if (m[4] !== undefined) s += (+m[4] * 3600 + +m[5] * 60 + (m[6] ? parseFloat(m[6]) : 0)) / 86400;
        return s;
      }
    }
    var dt = nativeValue instanceof Date ? nativeValue : (value instanceof Date ? value : null);
    if (dt && !isNaN(dt.getTime())) {
      return serial(dt.getUTCFullYear(), dt.getUTCMonth() + 1, dt.getUTCDate()) + (dt.getUTCHours() * 3600 + dt.getUTCMinutes() * 60 + dt.getUTCSeconds()) / 86400;
    }
    return null;
  }
  function fromSerial(s) { return new Date(Date.UTC(1899, 11, 30) + Math.round(s * 86400000)); }
  function isNull(dv) { return dv === null || dv === undefined || dv.value === '%null%' || dv.value === null || dv.value === undefined || (dv.nativeValue === null && dv.formattedValue === 'Null'); }
  function isNum(t) { return t === 'int' || t === 'float'; }
  function isDate(t) { return t === 'date' || t === 'date-time'; }

  /** columns: [{fieldName, dataType}], rows: [[DataValue,...]]
   *  -> { header, keep, types, cells (for Excel), display (strings), sortKeys (number|string) } */
  function shape(columns, rows, opts) {
    opts = opts || {};
    var seen = {}, keep = [], header = [];
    columns.forEach(function (c, i) {
      var name = cleanName(c.fieldName);
      if (seen[name]) return;                       // the text-mark duplicate of a measure, etc.
      seen[name] = true; keep.push(i); header.push(name);
    });
    // put the columns in the requested left-to-right order; unlisted ones keep their relative place, after the listed ones
    var want = (opts.columns === undefined ? DEFAULTS.columns : opts.columns).split(',')
                 .map(function (s) { return s.trim().toLowerCase(); }).filter(Boolean);
    if (want.length) {
      var ranked = header.map(function (h, i) { var p = want.indexOf(h.toLowerCase()); return [p < 0 ? want.length + i : p, i]; });
      ranked.sort(function (a, b) { return a[0] - b[0] || a[1] - b[1]; });
      keep = ranked.map(function (r) { return keep[r[1]]; });
      header = ranked.map(function (r) { return header[r[1]]; });
    }
    var dateFmt = opts.dateFormat || DEFAULTS.dateFormat, dtFmt = opts.dateTimeFormat || DEFAULTS.dateTimeFormat;
    var types = keep.map(function (i) { return columns[i].dataType; });
    var cells = new Array(rows.length), display = new Array(rows.length), sortKeys = new Array(rows.length);
    for (var r = 0; r < rows.length; r++) {
      var row = rows[r], line = new Array(keep.length), disp = new Array(keep.length), sk = new Array(keep.length);
      for (var k = 0; k < keep.length; k++) {
        var t = types[k], dv = row[keep[k]];
        if (isNull(dv)) { line[k] = null; disp[k] = ''; sk[k] = null; continue; }
        var txt = dv.formattedValue !== undefined && dv.formattedValue !== null ? String(dv.formattedValue) : String(dv.value);
        if (isDate(t)) {
          var s = toSerial(dv.value, dv.nativeValue);
          line[k] = s === null ? { t: 's', v: txt } : { t: 'n', v: s, z: t === 'date-time' ? dtFmt : dateFmt };
          sk[k] = s;
        } else if (isNum(t)) {
          var n = typeof dv.value === 'number' ? dv.value : Number(dv.value);
          line[k] = isNaN(n) ? { t: 's', v: txt } : { t: 'n', v: n };
          sk[k] = isNaN(n) ? txt.toLowerCase() : n;
        } else if (t === 'bool') {
          line[k] = { t: 'b', v: dv.value === true || dv.value === 'true' }; sk[k] = txt.toLowerCase();
        } else {
          line[k] = { t: 's', v: txt }; sk[k] = txt.toLowerCase();
        }
        disp[k] = txt;
      }
      cells[r] = line; display[r] = disp; sortKeys[r] = sk;
    }
    return { header: header, keep: keep, types: types, cells: cells, display: display, sortKeys: sortKeys };
  }

  /** a shaped object holding only the given row indices, in that order */
  function subset(shaped, idx) {
    return { header: shaped.header, keep: shaped.keep, types: shaped.types,
             cells: idx.map(function (i) { return shaped.cells[i]; }),
             display: idx.map(function (i) { return shaped.display[i]; }),
             sortKeys: idx.map(function (i) { return shaped.sortKeys[i]; }) };
  }

  /** filters: [{text, from, to}] per column; search: string -> array of row indices that pass */
  function filterIndices(shaped, filters, search) {
    var out = [], q = (search || '').trim().toLowerCase(), n = shaped.header.length;
    for (var r = 0; r < shaped.display.length; r++) {
      var disp = shaped.display[r], sk = shaped.sortKeys[r], ok = true;
      for (var c = 0; c < n && ok; c++) {
        var f = filters[c];
        if (!f) continue;
        if (f.text && disp[c].toLowerCase().indexOf(f.text.toLowerCase()) === -1) ok = false;
        if (ok && (f.from !== null && f.from !== undefined)) { if (sk[c] === null || sk[c] < f.from) ok = false; }
        if (ok && (f.to !== null && f.to !== undefined)) { if (sk[c] === null || sk[c] > f.to + 0.99999) ok = false; }
      }
      if (ok && q) {
        var hit = false;
        for (var s = 0; s < n; s++) if (disp[s].toLowerCase().indexOf(q) !== -1) { hit = true; break; }
        ok = hit;
      }
      if (ok) out.push(r);
    }
    return out;
  }

  /** stable sort of row indices by column c (dir 1 asc, -1 desc); nulls last */
  function sortIndices(shaped, idx, c, dir) {
    if (!dir || c < 0) return idx;
    var keyed = idx.map(function (i, p) { return [shaped.sortKeys[i][c], i, p]; });
    keyed.sort(function (a, b) {
      var x = a[0], y = b[0];
      if (x === null && y === null) return a[2] - b[2];
      if (x === null) return 1;
      if (y === null) return -1;
      if (x < y) return -dir; if (x > y) return dir;
      return a[2] - b[2];
    });
    return keyed.map(function (k) { return k[1]; });
  }

  /** SheetJS workbook from a shaped object; XLSX = the SheetJS global (or require('xlsx') in Node) */
  function buildWorkbook(XLSX, shaped, sheetName) {
    var aoa = [shaped.header.map(function (h) { return { t: 's', v: h }; })].concat(shaped.cells);
    var ws = XLSX.utils.aoa_to_sheet(aoa, { cellDates: false });
    ws['!cols'] = shaped.header.map(function (h, k) {
      var w = h.length;
      for (var r = 0; r < Math.min(shaped.cells.length, 500); r++) {
        var c = shaped.cells[r][k]; if (!c) continue;
        var len = c.z ? 10 : (typeof c.v === 'number' ? String(Math.round(c.v * 100) / 100).length : String(c.v).length);
        if (len > w) w = len;
      }
      return { wch: Math.min(Math.max(w + 2, 8), 60) };
    });
    if (shaped.cells.length) ws['!autofilter'] = { ref: XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: shaped.cells.length, c: shaped.header.length - 1 } }) };
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, (sheetName || 'Data').replace(/[\\\/\?\*\[\]:]/g, '').slice(0, 31) || 'Data');
    return wb;
  }

  /** CSV (UTF-8 BOM, ISO dates so Excel parses them on open) */
  function buildCsv(shaped) {
    function esc(s) { s = String(s); return /[",\r\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s; }
    function cell(c, k) {
      if (c === null) return '';
      if (c.z && typeof c.v === 'number') {
        var d = fromSerial(c.v), p = function (n) { return String(n).padStart(2, '0'); };
        var iso = d.getUTCFullYear() + '-' + p(d.getUTCMonth() + 1) + '-' + p(d.getUTCDate());
        return shaped.types[k] === 'date-time' ? iso + ' ' + p(d.getUTCHours()) + ':' + p(d.getUTCMinutes()) : iso;
      }
      return esc(c.v);
    }
    var lines = [shaped.header.map(esc).join(',')];
    for (var r = 0; r < shaped.cells.length; r++) lines.push(shaped.cells[r].map(cell).join(','));
    return '﻿' + lines.join('\r\n');
  }

  function stamp() { var d = new Date(), p = function (n) { return String(n).padStart(2, '0'); }; return d.getFullYear() + '-' + p(d.getMonth() + 1) + '-' + p(d.getDate()) + ' ' + p(d.getHours()) + p(d.getMinutes()); }
  function fmtDate(s) { var d = fromSerial(s), p = function (n) { return String(n).padStart(2, '0'); }; return p(d.getUTCDate()) + '/' + p(d.getUTCMonth() + 1) + '/' + d.getUTCFullYear(); }

  var api = { cleanName: cleanName, serial: serial, toSerial: toSerial, fromSerial: fromSerial, shape: shape, subset: subset,
              filterIndices: filterIndices, sortIndices: sortIndices, buildWorkbook: buildWorkbook, buildCsv: buildCsv,
              DEFAULTS: DEFAULTS, stamp: stamp, fmtDate: fmtDate };
  if (typeof module !== 'undefined' && module.exports) { module.exports = api; return; }   // Node self-test
  root.TatExport = api;

  // ================================================================ browser / Tableau side
  var $ = function (id) { return document.getElementById(id); };
  var cfg = {}, S = { shaped: null, filters: [], search: '', sortCol: -1, sortDir: 0, page: 1, pageSize: 100,
                      view: [], bounds: [], busy: false, suppress: false, sheet: null, cap: 0, capLabel: '' };

  function setting(k) { var v = tableau.extensions.settings.get(k); return (v === undefined || v === null || v === '') ? DEFAULTS[k] : v; }
  function status(msg, kind) { var el = $('status'); el.textContent = msg || ''; el.className = kind || ''; }
  function esc(s) { return String(s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }
  function busy(on) {
    S.busy = on;
    ['xl', 'csv', 'first', 'prev', 'next', 'last'].forEach(function (id) { $(id).disabled = on; });
    if (!on) renderPager();
  }

  function findSheet() {
    var ws = tableau.extensions.dashboardContent.dashboard.worksheets;
    for (var i = 0; i < ws.length; i++) if (ws[i].name === cfg.sheet) return ws[i];
    throw new Error('Worksheet "' + cfg.sheet + '" is not on this dashboard');
  }
  function readAll(sheet, onProgress) {
    var opts = { ignoreSelection: true };
    if (typeof sheet.getSummaryDataReaderAsync !== 'function') {
      return sheet.getSummaryDataAsync(opts).then(function (t) { return { columns: t.columns, rows: t.data }; });
    }
    return sheet.getSummaryDataReaderAsync(10000, opts).then(function (reader) {
      var columns = null, rows = [], pages = reader.pageCount, chain = Promise.resolve();
      for (var p = 0; p < pages; p++) (function (p) {
        chain = chain.then(function () {
          return reader.getPageAsync(p).then(function (t) {
            if (!columns) columns = t.columns;
            for (var i = 0; i < t.data.length; i++) rows.push(t.data[i]);
            if (onProgress) onProgress(rows.length, reader.totalRowCount);
          });
        });
      })(p);
      return chain.then(function () { return reader.releaseAsync().then(function () { return { columns: columns || [], rows: rows }; }); });
    });
  }
  function countRows(sheet) {
    if (typeof sheet.getSummaryDataReaderAsync !== 'function') return Promise.resolve(null);
    return sheet.getSummaryDataReaderAsync(1, { ignoreSelection: true })
      .then(function (r) { var n = r.totalRowCount; return r.releaseAsync().then(function () { return n; }); })
      .catch(function () { return null; });
  }

  // Tableau's rule for list parameters WITH aliases: changeValueAsync must receive the ALIAS (docs: "the aliased value
  // should be passed"). Passing the raw value is refused by Desktop with 'set-parameter-value ... bad value: value-string'.
  function candidates(p, raw) {
    var list = (p.allowableValues && p.allowableValues.allowableValues) || [], hit = null, out = [], seen = {};
    for (var i = 0; i < list.length; i++) if (String(list[i].value) === String(raw)) { hit = list[i]; break; }
    function add(v) { var k = typeof v + ':' + String(v); if (v !== undefined && v !== null && v !== '' && !seen[k]) { seen[k] = true; out.push(v); } }
    if (hit) { if (hit.hasAlias) add(hit.aliasValue); add(hit.formattedValue); add(hit.value); }
    add(String(raw)); if (!isNaN(Number(raw))) add(Number(raw));
    return out;
  }
  function setParam(p, raw) {
    var tries = candidates(p, raw), errors = [];
    return (function next(i) {
      if (i >= tries.length) return Promise.reject(new Error('Could not set "' + p.name + '" to ' + raw + ' (' + errors.join(' | ') + ')'));
      return p.changeValueAsync(tries[i]).then(function (dv) {
        if (dv && String(dv.value) === String(raw)) return dv;
        errors.push(JSON.stringify(tries[i]) + ' -> ' + (dv ? dv.value : '?')); return next(i + 1);
      }, function (e) { errors.push(JSON.stringify(tries[i]) + ': ' + (e && e.message ? e.message : e)); return next(i + 1); });
    })(0);
  }
  function withAllRows(sheet, fn) {
    var dash = tableau.extensions.dashboardContent.dashboard, param = null, previous = null, flipped = false, listener = null;
    S.suppress = true;
    return dash.findParameterAsync(cfg.param).then(function (p) {
      param = p;
      if (!p) { status('Parameter "' + cfg.param + '" not found; exporting what is loaded', 'warn'); return fn(); }
      previous = String(p.currentValue.value);
      if (previous === String(cfg.paramAll)) return fn();
      return countRows(sheet).then(function (n) {
        if (n !== null && !isNaN(Number(previous)) && n < Number(previous)) return fn();
        return flip();
      });
      function flip() {
        var changed = new Promise(function (resolve) {
          var done = false, finish = function () { if (!done) { done = true; resolve(); } };
          try { listener = sheet.addEventListener(tableau.TableauEventType.SummaryDataChanged, finish); } catch (e) { /* older API */ }
          setTimeout(finish, 90000);
          setTimeout(function () { if (!listener) finish(); }, 6000);
        });
        status('Loading all rows from the data source…', 'warn');
        return setParam(p, cfg.paramAll).then(function () { flipped = true; return changed; }).then(function () {
          if (listener) { try { listener(); } catch (e) { /* ignore */ } listener = null; }
          return fn();
        });
      }
    }).then(function (r) { return restore().then(function () { unsuppress(); return r; }); },
            function (e) { return restore().then(function () { unsuppress(); throw e; }); });
    // the restore fires one more SummaryDataChanged a moment later; ignore that one too (the on-screen data is unchanged)
    function unsuppress() { setTimeout(function () { S.suppress = false; }, 1500); }
    function restore() {
      if (!flipped || !param || previous === null) return Promise.resolve();
      return setParam(param, previous).then(function () { flipped = false; },
        function (e) { status('Exported, but "' + cfg.param + '" could not be set back: ' + e.message, 'warn'); });
    }
  }

  // ---------------------------------------------------------------- rendering
  function bounds(shaped) {
    return shaped.types.map(function (t, c) {
      if (!isDate(t)) return null;
      var lo = null, hi = null;
      for (var r = 0; r < shaped.sortKeys.length; r++) {
        var v = shaped.sortKeys[r][c]; if (v === null) continue;
        v = Math.floor(v);
        if (lo === null || v < lo) lo = v;
        if (hi === null || v > hi) hi = v;
      }
      return lo === null ? null : { lo: lo, hi: hi };
    });
  }
  function renderHead() {
    var h = S.shaped.header, out = '<tr>';
    for (var c = 0; c < h.length; c++) {
      var arr = S.sortCol === c ? '<span class="arrow">' + (S.sortDir > 0 ? '▲' : '▼') + '</span>' : '';
      out += '<th data-c="' + c + '" title="' + esc(h[c]) + ' – click to sort">' + esc(h[c]) + arr + '</th>';
    }
    out += '</tr><tr class="filters">';
    for (var f = 0; f < h.length; f++) {
      var b = S.bounds[f];
      if (b && b.hi > b.lo) {
        var fl = S.filters[f], from = fl.from === null || fl.from === undefined ? b.lo : fl.from, to = fl.to === null || fl.to === undefined ? b.hi : fl.to;
        out += '<th><div class="rng"><div class="lbl"><span id="lo' + f + '">' + fmtDate(from) + '</span><span id="hi' + f + '">' + fmtDate(to) + '</span></div>' +
               '<div class="track">' +
               '<input type="range" data-c="' + f + '" data-e="lo" min="' + b.lo + '" max="' + b.hi + '" value="' + from + '">' +
               '<input type="range" data-c="' + f + '" data-e="hi" min="' + b.lo + '" max="' + b.hi + '" value="' + to + '">' +
               '</div></div></th>';
      } else {
        out += '<th><input type="text" size="1" data-c="' + f + '" placeholder="Filter value" value="' + esc(S.filters[f].text || '') + '" autocomplete="off"></th>';
      }
    }
    $('head').innerHTML = out + '</tr>';
  }
  function renderBody() {
    var start = (S.page - 1) * S.pageSize, end = Math.min(start + S.pageSize, S.view.length), rows = [];
    for (var i = start; i < end; i++) {
      var d = S.shaped.display[S.view[i]], tds = '';
      for (var c = 0; c < d.length; c++) tds += '<td class="' + (isNum(S.shaped.types[c]) ? 'num' : '') + '" title="' + esc(d[c]) + '">' + esc(d[c]) + '</td>';
      rows.push('<tr>' + tds + '</tr>');
    }
    $('body').innerHTML = rows.join('');
    var none = S.view.length === 0;
    $('empty').style.display = none ? 'block' : 'none';
    $('empty').textContent = none ? (S.shaped.display.length ? 'No rows match the filters on this table.' : 'No rows in the current dashboard filters.') : '';
  }
  function renderPager() {
    var pages = Math.max(1, Math.ceil(S.view.length / S.pageSize));
    if (S.page > pages) S.page = pages;
    $('pageinfo').textContent = S.view.length ? 'Page ' + S.page + ' of ' + pages : '–';
    $('first').disabled = $('prev').disabled = S.busy || S.page <= 1;
    $('next').disabled = $('last').disabled = S.busy || S.page >= pages;
    var total = S.shaped ? S.shaped.display.length : 0;
    var atCap = S.cap > 0 && total >= S.cap;
    $('count').innerHTML = (S.view.length === total ? '<b>' + total.toLocaleString() + '</b> rows'
                                                    : '<b>' + S.view.length.toLocaleString() + '</b> of ' + total.toLocaleString() + ' rows')
                         + (atCap ? ' <span class="cap">limit reached</span>' : '');
    $('count').title = total.toLocaleString() + ' row' + (total === 1 ? '' : 's') + ' loaded from "' + cfg.sheet
                     + '" with the dashboard filters as they are now'
                     + (S.capLabel ? '\n"' + cfg.param + '" is set to ' + S.capLabel
                                   + (atCap ? ' - there may be more rows than this' : ' - this is everything those filters return') : '');
    var active = S.search || S.filters.some(function (f) { return f.text || f.from !== null && f.from !== undefined || f.to !== null && f.to !== undefined; });
    $('clear').hidden = !active;
  }
  function applyView(resetPage) {
    S.view = sortIndices(S.shaped, filterIndices(S.shaped, S.filters, S.search), S.sortCol, S.sortDir);
    if (resetPage !== false) S.page = 1;
    renderBody(); renderPager();
  }

  // ---------------------------------------------------------------- load + events
  // what "Rows shown" is set to right now, so the row count can say whether it is the limiter
  function readCap() {
    try {
      return tableau.extensions.dashboardContent.dashboard.findParameterAsync(cfg.param).then(function (p) {
        var v = p && p.currentValue;
        S.cap = v && !isNaN(Number(v.value)) ? Number(v.value) : 0;
        S.capLabel = v ? (v.formattedValue || String(v.value)) : '';
      }, function () { S.cap = 0; S.capLabel = ''; });
    } catch (e) { S.cap = 0; S.capLabel = ''; return Promise.resolve(); }
  }

  function load(first) {
    if (S.busy) return Promise.resolve();
    var keep = $('status').className === 'ok' ? $('status').textContent : null;   // don't wipe an export confirmation
    busy(true); status('Reading rows…', 'warn');
    return readCap()
      .then(function () { return readAll(S.sheet, function (n, t) { status('Reading rows… ' + n.toLocaleString() + (t ? ' of ' + t.toLocaleString() : ''), 'warn'); }); })
      .then(function (data) {
        var was = S.shaped ? S.shaped.header.join('') : null;
        S.shaped = shape(data.columns, data.rows, cfg);
        S.bounds = bounds(S.shaped);
        if (was !== S.shaped.header.join('')) {                 // columns changed -> fresh filters
          S.filters = S.shaped.header.map(function () { return { text: '', from: null, to: null }; });
          S.sortCol = -1; S.sortDir = 0;
        }
        S.filters.forEach(function (f, c) {                            // re-clamp date ranges into the new bounds
          var b = S.bounds[c]; if (!b) return;
          if (f.from !== null && f.from !== undefined) f.from = Math.max(b.lo, Math.min(f.from, b.hi));
          if (f.to !== null && f.to !== undefined) f.to = Math.max(b.lo, Math.min(f.to, b.hi));
        });
        renderHead(); applyView(true);
        status(keep || (first ? '' : 'Updated'), keep || !first ? 'ok' : '');
        if (!keep && !first) setTimeout(function () { if ($('status').textContent === 'Updated') status(''); }, 2500);
      })
      .catch(function (e) { status('Could not read the table: ' + (e && e.message ? e.message : e), 'error'); })
      .then(function () { busy(false); });
  }

  function exportFile(csv) {
    if (!S.shaped || S.busy) return;
    busy(true);
    var t0 = Date.now(), name = cfg.filename + ' ' + stamp();
    var work = function () {
      if (!$('all').checked) return Promise.resolve(subset(S.shaped, S.view));
      return readAll(S.sheet, function (n, t) { status('Reading all rows… ' + n.toLocaleString() + (t ? ' of ' + t.toLocaleString() : ''), 'warn'); })
        .then(function (data) {
          var full = shape(data.columns, data.rows, cfg);
          return subset(full, sortIndices(full, filterIndices(full, S.filters, S.search), S.sortCol, S.sortDir));
        });
    };
    var job = $('all').checked ? withAllRows(S.sheet, work) : work();
    job.then(function (out) {
      status('Building the file… (' + out.cells.length.toLocaleString() + ' rows)', 'warn');
      return new Promise(function (r) { setTimeout(r, 30); }).then(function () {
        var blob;
        if (csv) blob = new Blob([buildCsv(out)], { type: 'text/csv;charset=utf-8' });
        else blob = new Blob([XLSX.write(buildWorkbook(XLSX, out, cfg.sheet), { bookType: 'xlsx', type: 'array', compression: true })],
                             { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        var a = document.createElement('a');
        a.href = URL.createObjectURL(blob); a.download = name + (csv ? '.csv' : '.xlsx');
        document.body.appendChild(a); a.click();
        setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 4000);
        status(out.cells.length.toLocaleString() + ' rows exported in ' + Math.round((Date.now() - t0) / 1000) + ' s' + (csv ? ' (CSV)' : ''), 'ok');
      });
    }).catch(function (e) { status('Export failed: ' + (e && e.message ? e.message : e), 'error'); })
      .then(function () { busy(false); });
  }

  function wire() {
    var timer = null;
    $('search').addEventListener('input', function () {
      S.search = this.value;
      clearTimeout(timer); timer = setTimeout(function () { applyView(true); }, 180);
    });
    $('head').addEventListener('input', function (e) {
      var el = e.target, c = +el.getAttribute('data-c');
      if (el.type === 'text') {
        S.filters[c].text = el.value;
        clearTimeout(timer); timer = setTimeout(function () { applyView(true); }, 180);
      } else if (el.type === 'range') {
        var b = S.bounds[c], which = el.getAttribute('data-e'), v = +el.value;
        var lo = S.filters[c].from === null || S.filters[c].from === undefined ? b.lo : S.filters[c].from;
        var hi = S.filters[c].to === null || S.filters[c].to === undefined ? b.hi : S.filters[c].to;
        if (which === 'lo') { lo = Math.min(v, hi); el.value = lo; } else { hi = Math.max(v, lo); el.value = hi; }
        S.filters[c].from = lo <= b.lo ? null : lo;
        S.filters[c].to = hi >= b.hi ? null : hi;
        $('lo' + c).textContent = fmtDate(lo); $('hi' + c).textContent = fmtDate(hi);
        clearTimeout(timer); timer = setTimeout(function () { applyView(true); }, 120);
      }
    });
    $('head').addEventListener('click', function (e) {
      var th = e.target.closest ? e.target.closest('th') : null;
      if (!th || !th.hasAttribute('data-c') || th.parentNode.className === 'filters') return;
      var c = +th.getAttribute('data-c');
      if (S.sortCol !== c) { S.sortCol = c; S.sortDir = 1; }
      else if (S.sortDir === 1) S.sortDir = -1;
      else { S.sortCol = -1; S.sortDir = 0; }
      renderHead(); applyView(true);
    });
    $('clear').addEventListener('click', function () {
      S.search = ''; $('search').value = '';
      S.filters = S.shaped.header.map(function () { return { text: '', from: null, to: null }; });
      renderHead(); applyView(true);
    });
    $('first').addEventListener('click', function () { S.page = 1; renderBody(); renderPager(); });
    $('prev').addEventListener('click', function () { S.page--; renderBody(); renderPager(); });
    $('next').addEventListener('click', function () { S.page++; renderBody(); renderPager(); });
    $('last').addEventListener('click', function () { S.page = Math.max(1, Math.ceil(S.view.length / S.pageSize)); renderBody(); renderPager(); });
    $('size').addEventListener('change', function () { S.pageSize = +this.value; S.page = 1; renderBody(); renderPager(); });
    $('xl').addEventListener('click', function () { exportFile(false); });
    $('csv').addEventListener('click', function () { exportFile(true); });
  }

  document.addEventListener('DOMContentLoaded', function () {
    var started = false;
    setTimeout(function () { if (!started) status('Hosting OK - this table only works inside the Tableau dashboard', 'warn'); }, 6000);
    tableau.extensions.initializeAsync().then(function () {
      started = true;
      Object.keys(DEFAULTS).forEach(function (k) { cfg[k] = setting(k); });
      S.pageSize = Number(cfg.pageSize) || 100;
      $('size').value = String(S.pageSize);
      if (!Array.prototype.some.call($('size').options, function (o) { return o.value === String(S.pageSize); })) {
        $('size').insertAdjacentHTML('afterbegin', '<option selected>' + S.pageSize + '</option>');
      }
      $('all').checked = String(cfg.allRows) !== 'false';
      $('xl').lastChild.textContent = ' ' + cfg.label;
      S.sheet = findSheet();
      wire();
      try { S.sheet.addEventListener(tableau.TableauEventType.SummaryDataChanged, function () { if (!S.suppress) load(false); }); } catch (e) { /* older API */ }
      return load(true);
    }).catch(function (e) { status('Could not start: ' + (e && e.message ? e.message : e), 'error'); });
  });
})(typeof window !== 'undefined' ? window : this);
