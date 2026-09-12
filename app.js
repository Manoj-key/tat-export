/* TAT Excel Export -- a Tableau dashboard extension that exports one worksheet to Excel with REAL data types:
 * dates as Excel dates (dd/mm/yyyy), numbers as numbers, text as text; clean column names; optional one-click
 * "all rows" (flips the dashboard's 'Rows shown' parameter to the export value, reads, restores it).
 * Libraries: Tableau Extensions API 1.17 (Tableau's own @tableau/tabextsandbox package), SheetJS 0.18.5 (Apache-2.0).
 * The conversion and workbook-building functions are pure so they can be unit-tested in Node (see selftest.js).
 */
(function (root) {
  'use strict';

  var DEFAULTS = {
    sheet: 'Raw Data Table',        // worksheet on the dashboard to export
    filename: 'TAT Raw Data',       // file name prefix; a timestamp is appended
    dateFormat: 'dd/mm/yyyy',       // Excel number format for date columns
    dateTimeFormat: 'dd/mm/yyyy hh:mm',
    label: 'Export to Excel',
    param: 'Rows shown',            // parameter that caps the on-screen preview
    paramAll: '10000000',           // its "all rows" value
    allRows: 'true'                 // default state of the "All rows" checkbox
  };

  // ---------------------------------------------------------------- pure helpers
  var AGG_RE = /^(SUM|AVG|MIN|MAX|CNT|CNTD|COUNT|COUNTD|ATTR|AGG|MEDIAN|STDEV|VAR|TOTAL)\((.*)\)$/i;
  function cleanName(n) { var m = AGG_RE.exec(n); return m ? m[2] : n; }

  // Excel serial for a calendar date (1900 date system; valid for dates after 28 Feb 1900)
  function serial(y, m, d) {
    var ms = Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30);
    return Math.round(ms / 86400000);
  }
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
  function isNull(dv) { return dv === null || dv === undefined || dv.value === '%null%' || dv.value === null || dv.value === undefined || (dv.nativeValue === null && dv.formattedValue === 'Null'); }

  /** columns: [{fieldName, dataType, index}], rows: [[DataValue,...]] -> {header:[names], keep:[colIdx], cells:[[cellobj|null]]} */
  function shape(columns, rows, opts) {
    opts = opts || {};
    var seen = {}, keep = [], header = [];
    columns.forEach(function (c, i) {
      var name = cleanName(c.fieldName);
      if (seen[name]) return;                       // the text-mark duplicate of a measure, etc.
      seen[name] = true; keep.push(i); header.push(name);
    });
    var dateFmt = opts.dateFormat || DEFAULTS.dateFormat, dtFmt = opts.dateTimeFormat || DEFAULTS.dateTimeFormat;
    var out = new Array(rows.length);
    for (var r = 0; r < rows.length; r++) {
      var row = rows[r], line = new Array(keep.length);
      for (var k = 0; k < keep.length; k++) {
        var c = columns[keep[k]], dv = row[keep[k]];
        if (isNull(dv)) { line[k] = null; continue; }
        switch (c.dataType) {
          case 'date': { var s = toSerial(dv.value, dv.nativeValue); line[k] = s === null ? { t: 's', v: String(dv.formattedValue) } : { t: 'n', v: s, z: dateFmt }; break; }
          case 'date-time': { var s2 = toSerial(dv.value, dv.nativeValue); line[k] = s2 === null ? { t: 's', v: String(dv.formattedValue) } : { t: 'n', v: s2, z: dtFmt }; break; }
          case 'int': case 'float': { var n = typeof dv.value === 'number' ? dv.value : Number(dv.value); line[k] = isNaN(n) ? { t: 's', v: String(dv.formattedValue) } : { t: 'n', v: n }; break; }
          case 'bool': { line[k] = { t: 'b', v: dv.value === true || dv.value === 'true' }; break; }
          default: { var txt = dv.formattedValue !== undefined && dv.formattedValue !== null ? dv.formattedValue : dv.value; line[k] = { t: 's', v: String(txt) }; }
        }
      }
      out[r] = line;
    }
    return { header: header, keep: keep, cells: out, types: keep.map(function (i) { return columns[i].dataType; }) };
  }

  /** SheetJS workbook from shape() output; XLSX = the SheetJS global (or require('xlsx') in Node) */
  function buildWorkbook(XLSX, shaped, sheetName) {
    var aoa = [shaped.header.map(function (h) { return { t: 's', v: h }; })].concat(shaped.cells.map(function (line) {
      return line.map(function (cell) { return cell === null ? null : cell; });
    }));
    var ws = XLSX.utils.aoa_to_sheet(aoa, { cellDates: false });
    // column widths from the longest visible value (capped), header row frozen, autofilter on
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

  /** CSV (UTF-8 BOM, ISO dates so Excel parses them on open) -- the light-weight route for very large exports */
  function buildCsv(shaped) {
    function esc(s) { s = String(s); return /[",\r\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s; }
    function cell(c, k) {
      if (c === null) return '';
      if (c.z && typeof c.v === 'number') {                       // serial -> ISO
        var ms = Date.UTC(1899, 11, 30) + Math.round(c.v * 86400000), d = new Date(ms);
        var iso = d.getUTCFullYear() + '-' + String(d.getUTCMonth() + 1).padStart(2, '0') + '-' + String(d.getUTCDate()).padStart(2, '0');
        return shaped.types[k] === 'date-time' ? iso + ' ' + String(d.getUTCHours()).padStart(2, '0') + ':' + String(d.getUTCMinutes()).padStart(2, '0') : iso;
      }
      return esc(c.v);
    }
    var lines = [shaped.header.map(esc).join(',')];
    for (var r = 0; r < shaped.cells.length; r++) lines.push(shaped.cells[r].map(cell).join(','));
    return '\uFEFF' + lines.join('\r\n');
  }

  function stamp() { var d = new Date(), p = function (n) { return String(n).padStart(2, '0'); }; return d.getFullYear() + '-' + p(d.getMonth() + 1) + '-' + p(d.getDate()) + ' ' + p(d.getHours()) + p(d.getMinutes()); }

  var api = { cleanName: cleanName, serial: serial, toSerial: toSerial, shape: shape, buildWorkbook: buildWorkbook, buildCsv: buildCsv, DEFAULTS: DEFAULTS, stamp: stamp };
  if (typeof module !== 'undefined' && module.exports) { module.exports = api; return; }   // Node self-test
  root.TatExport = api;

  // ---------------------------------------------------------------- browser / Tableau side
  var $ = function (id) { return document.getElementById(id); };
  var cfg = {};
  function setting(k) { var v = tableau.extensions.settings.get(k); return (v === undefined || v === null || v === '') ? DEFAULTS[k] : v; }
  function status(msg, kind) { var el = $('status'); el.textContent = msg; el.className = kind || ''; }

  function findSheet() {
    var ws = tableau.extensions.dashboardContent.dashboard.worksheets;
    for (var i = 0; i < ws.length; i++) if (ws[i].name === cfg.sheet) return ws[i];
    throw new Error('Worksheet "' + cfg.sheet + '" is not on this dashboard');
  }
  function download(blob, name) {
    var a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = name;
    document.body.appendChild(a); a.click(); setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 4000);
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
            onProgress(rows.length, reader.totalRowCount);
          });
        });
      })(p);
      return chain.then(function () { return reader.releaseAsync().then(function () { return { columns: columns || [], rows: rows, total: reader.totalRowCount }; }); });
    });
  }
  function countRows(sheet) {
    if (typeof sheet.getSummaryDataReaderAsync !== 'function') return Promise.resolve(null);
    return sheet.getSummaryDataReaderAsync(1, { ignoreSelection: true }).then(function (r) { var n = r.totalRowCount; return r.releaseAsync().then(function () { return n; }); }).catch(function () { return null; });
  }
  // Tableau's rule for list parameters WITH aliases: changeValueAsync must receive the ALIAS (docs: "If the domain restriction
  // is type List and there are aliases defined for the list, the aliased value should be passed").  Passing the raw value
  // ("1000") is refused by Desktop with 'set-parameter-value ... bad value: value-string' (532B4D27).  So: find the allowable
  // entry whose raw value matches, try its alias first, then its formatted value, then the raw string / number, and verify
  // the value Tableau reports back after each attempt.
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
        if (dv && String(dv.value) === String(raw)) return dv;                      // verified
        errors.push(JSON.stringify(tries[i]) + ' -> ' + (dv ? dv.value : '?')); return next(i + 1);
      }, function (e) { errors.push(JSON.stringify(tries[i]) + ': ' + (e && e.message ? e.message : e)); return next(i + 1); });
    })(0);
  }
  // flip the preview parameter to "all", wait until the sheet has re-queried, run fn, restore the parameter (only if flipped)
  function withAllRows(sheet, fn) {
    var dash = tableau.extensions.dashboardContent.dashboard, param = null, previous = null, flipped = false, listener = null;
    return dash.findParameterAsync(cfg.param).then(function (p) {
      param = p;
      if (!p) { status('Parameter "' + cfg.param + '" not found; exporting the rows on screen', 'warn'); return fn(); }
      previous = String(p.currentValue.value);
      if (previous === String(cfg.paramAll)) return fn();
      return countRows(sheet).then(function (n) {
        if (n !== null && !isNaN(Number(previous)) && n < Number(previous)) { status('All ' + n.toLocaleString() + ' rows are already on screen', 'busy'); return fn(); }
        return flip();
      });
      function flip() {
        var changed = new Promise(function (resolve) {
          var done = false, finish = function () { if (!done) { done = true; resolve(); } };
          try { listener = sheet.addEventListener(tableau.TableauEventType.SummaryDataChanged, finish); } catch (e) { /* older API */ }
          setTimeout(finish, 90000);                                  // never hang: Snowflake live can be slow, but not forever
          setTimeout(function () { if (!listener) finish(); }, 6000);
        });
        status('Loading all rows from the data source…', 'busy');
        return setParam(p, cfg.paramAll).then(function () { flipped = true; return changed; }).then(function () {
          if (listener) { try { listener(); } catch (e) { /* ignore */ } listener = null; }
          return fn();
        });
      }
    }).then(function (result) { return restore().then(function () { return result; }); }, function (err) { return restore().then(function () { throw err; }); });
    function restore() {
      if (!flipped || !param || previous === null) return Promise.resolve();
      return setParam(param, previous).then(function () { flipped = false; }, function (e) { status('Exported, but "' + cfg.param + '" could not be set back: ' + e.message, 'warn'); });
    }
  }
  function run(csv) {
    var btn = $('go'), btn2 = $('csv'); btn.disabled = btn2.disabled = true;
    var sheet, t0 = Date.now();
    try { sheet = findSheet(); } catch (e) { status(e.message, 'error'); btn.disabled = btn2.disabled = false; return; }
    var work = function () {
      status('Reading rows…', 'busy');
      return readAll(sheet, function (n, total) { status('Reading rows… ' + n.toLocaleString() + (total ? ' of ' + total.toLocaleString() : ''), 'busy'); });
    };
    var job = $('all').checked ? withAllRows(sheet, work) : work();
    job.then(function (data) {
      status('Building the file… (' + data.rows.length.toLocaleString() + ' rows)', 'busy');
      return new Promise(function (resolve) { setTimeout(resolve, 30); }).then(function () {
        var shaped = shape(data.columns, data.rows, cfg), name = cfg.filename + ' ' + stamp();
        if (csv) { download(new Blob([buildCsv(shaped)], { type: 'text/csv;charset=utf-8' }), name + '.csv'); }
        else {
          var wb = buildWorkbook(XLSX, shaped, cfg.sheet);
          var out = XLSX.write(wb, { bookType: 'xlsx', type: 'array', compression: true });
          download(new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), name + '.xlsx');
        }
        status(data.rows.length.toLocaleString() + ' rows exported in ' + Math.round((Date.now() - t0) / 1000) + ' s' + (csv ? ' (CSV)' : ''), 'ok');
      });
    }).catch(function (err) {
      status('Export failed: ' + (err && err.message ? err.message : err), 'error');
    }).then(function () { btn.disabled = btn2.disabled = false; });
  }

  document.addEventListener('DOMContentLoaded', function () {
    // outside a Tableau dashboard (e.g. someone opens the hosted URL in a browser to check the hosting) initializeAsync never
    // resolves: say so after 6 s instead of showing "Starting…" forever -- that message doubles as the hosting check
    var started = false;
    setTimeout(function () { if (!started) status('Hosting OK - this button only works inside the Tableau dashboard', 'warn'); }, 6000);
    tableau.extensions.initializeAsync().then(function () {
      started = true;
      Object.keys(DEFAULTS).forEach(function (k) { cfg[k] = setting(k); });
      $('go').textContent = cfg.label; $('all').checked = String(cfg.allRows) !== 'false';
      $('go').addEventListener('click', function () { run(false); });
      $('csv').addEventListener('click', function () { run(true); });
      status('');
    }).catch(function (err) { status('Extension could not start: ' + (err && err.message ? err.message : err), 'error'); });
  });
})(typeof window !== 'undefined' ? window : this);
