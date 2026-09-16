/* WO Volume Raw Data Table v3 -- a Tableau dashboard extension that renders one worksheet as an interactive table
 * and reads a few tiny aggregate worksheets for exact, filter-aware insight.
 *
 *   chips row      every dashboard filter as a removable chip (getFiltersAsync / clearFilterAsync), Clear all,
 *                  the exact "rows match" count and the data-as-of stamp (from the aggregate sheets)
 *   toolbar        load size, paging, column chooser, search, one Download button
 *   insight band   summary sentence, six tiles, three minis -- all from Tableau aggregates, never from the
 *                  loaded rows (which are a preview and would be wrong)
 *   download       Excel / CSV, all filtered rows or the loaded page, all or shown columns, an "About this
 *                  export" sheet, a filter-aware file name, a Web Worker build with progress and Cancel
 *   states         skeleton on first load, empty state naming the narrowest filter, done toast
 *
 * Everything workbook-specific is a <setting>; DEFAULTS below are the fallbacks (and the values for the WO page).
 * Libraries: Tableau Extensions API 1.17, SheetJS 0.18.5 (Apache-2.0). Pure functions are exported for Node tests.
 */
(function (root) {
  'use strict';

  var DEFAULTS = {
    sheet: 'WO Table',
    filename: 'Work Order Volume',
    grain: 'work orders',                 // the noun in the count and the empty state
    dateFormat: 'dd mmm yyyy',            // unambiguous everywhere; a real date in Excel
    dateTimeFormat: 'dd mmm yyyy hh:mm',
    label: 'Download',
    param: '',                            // Tableau-side row cap parameter ('' = none; see v2 notes)
    paramAll: '10000000',
    allRows: 'true',
    pageSize: '100',
    previewRows: '1000',
    sortBy: 'WO Closed Date',
    sortDir: 'desc',
    exactColumns: 'true',
    columns: 'Work Order, WO Closed Date, Closed Fiscal Qtr, Closed Fiscal Year, Order Type, Deliverable Services, '
           + 'Customer, WDF Site, WDF Country, WDF Region, WDB Site, WDB Country, WDB Region, RTF / RTK, '
           + 'Asset Product, Product Description, Product Line, PL Revised, PL Detail, SOM Outsource Approval, '
           + 'Service Type, Keysight Care, Asset Manufacturer, Total Transit (Days)',
    defaultColumns: '',                   // comma list shown by default ('' = all of `columns`)
    // aggregate worksheets on the dashboard (rule-24 header tables); '' disables that piece
    aggSheet: 'WO Agg',                   // one row: Agg Count, Agg RTK, Agg Trn, Agg Transit, Agg Customers, Agg Products, Agg PLs, Agg Min Date, Agg Max Date
    aggAllSheet: 'WO Agg All',            // unfiltered: Agg Count, Agg Max Date
    byQuarterSheet: 'WO By Quarter',      // dim + Agg Count
    byRegionSheet: 'WO By Region',        // dim + Agg Count
    byPLSheet: 'WO By PL',                // dim + Agg Count
    byStatusSheet: '',                    // WOS page: dim + Agg Count
    periodParam: 'Period',                // the rail's period control (shown as an informational chip)
    dateField: 'WO Closed Date',          // the date column whose slider spans the WHOLE filtered range (from aggSheet)
    hideFilters: 'In Period',             // filters never shown as chips (semicolon-separated field names)
    filterLabels: 'Closed Fiscal Year=Fiscal year;Closed Fiscal Quarter=Fiscal quarter;WDF_SITE_REGION=WDF region;'
                + 'WDB_SITE_REGION=WDB region;IS_RTF_RTK=RTF / RTK;SERVICE_TYPE=Service type;KEYSIGHT_CARE=Keysight Care;'
                + 'PL Revised=Product line;CUSTOMER_NAME=Customer;WOS Status=WOS status',
    help: 'WDF Site=Work done FOR: the site that received the work order;WDF Country=Country of the work-done-for site;'
        + 'WDF Region=Region of the work-done-for site;WDB Site=Work done BY: the site that did the work;'
        + 'WDB Country=Country of the work-done-by site;WDB Region=Region of the work-done-by site;'
        + 'RTF / RTK=Return to Factory / Return to Keysight;PL Revised=Product line with WN split into WN-TA, WN-PMPS, WN-NA, WN-DTA;'
        + 'PL Detail=Product family, computed from the product-line mapping;Total Transit (Days)=Days in transit between the WDF and WDB sites',
    workbook: 'WO Volume Analytics'
  };

  // ================================================================ pure helpers
  var AGG_RE = /^(SUM|AVG|MIN|MAX|CNT|CNTD|COUNT|COUNTD|ATTR|AGG|MEDIAN|STDEV|VAR|TOTAL)\((.*)\)$/i;
  function cleanName(n) { var m = AGG_RE.exec(n); return m ? m[2] : n; }
  function serial(y, m, d) { return Math.round((Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 86400000); }
  var DT_RE = /^(\d{4})-(\d{2})-(\d{2})(?:[ T](\d{2}):(\d{2})(?::(\d{2}(?:\.\d+)?))?)?/;
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
  var MON = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  function p2(n) { return String(n).padStart(2, '0'); }
  function fmtDate(s) { var d = fromSerial(s); return p2(d.getUTCDate()) + ' ' + MON[d.getUTCMonth()] + ' ' + d.getUTCFullYear(); }
  function fmtDateTime(s) { var d = fromSerial(s); return fmtDate(s) + ' ' + p2(d.getUTCHours()) + ':' + p2(d.getUTCMinutes()); }
  function isNull(dv) { return dv === null || dv === undefined || dv.value === '%null%' || dv.value === null || dv.value === undefined || (dv.nativeValue === null && dv.formattedValue === 'Null'); }
  function isNum(t) { return t === 'int' || t === 'float'; }
  function isDate(t) { return t === 'date' || t === 'date-time'; }
  function num(dv) {
    if (isNull(dv)) return null;
    var n = typeof dv.value === 'number' ? dv.value : Number(dv.value);
    if (!isNaN(n)) return n;
    var f = String(dv.formattedValue || '').replace(/[^0-9.\-]/g, '');
    return f === '' ? null : Number(f);
  }
  function fmtInt(n) { return n === null || n === undefined || isNaN(n) ? '–' : Math.round(n).toLocaleString('en-GB'); }
  function pct(a, b, d) { if (!b) return '–'; return (100 * a / b).toFixed(d === undefined ? 1 : d) + '%'; }
  function kv(str) {   // 'a=b;c=d' -> {a:'b', c:'d'}
    var out = {};
    String(str || '').split(';').forEach(function (p) { var i = p.indexOf('='); if (i > 0) out[p.slice(0, i).trim()] = p.slice(i + 1).trim(); });
    return out;
  }
  function list(str) { return String(str || '').split(',').map(function (s) { return s.trim(); }).filter(Boolean); }

  /** columns: [{fieldName, dataType}], rows: [[DataValue,...]] -> shaped table (see v2) */
  function shape(columns, rows, opts) {
    opts = opts || {};
    var seen = {}, keep = [], header = [];
    columns.forEach(function (c, i) {
      var name = cleanName(c.fieldName);
      if (seen[name]) return;
      seen[name] = true; keep.push(i); header.push(name);
    });
    var want = list(opts.columns === undefined ? DEFAULTS.columns : opts.columns).map(function (s) { return s.toLowerCase(); });
    var exact = String(opts.exactColumns === undefined ? DEFAULTS.exactColumns : opts.exactColumns) === 'true';
    if (want.length) {
      var ranked = header.map(function (h, i) { var p = want.indexOf(h.toLowerCase()); return [p < 0 ? want.length + i : p, i, p]; });
      if (exact) ranked = ranked.filter(function (r) { return r[2] >= 0; });
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
          if (s === null) { line[k] = { t: 's', v: txt }; sk[k] = null; }
          else { line[k] = { t: 'n', v: s, z: t === 'date-time' ? dtFmt : dateFmt }; sk[k] = s; txt = t === 'date-time' ? fmtDateTime(s) : fmtDate(s); }
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
  function subset(shaped, idx) {
    return { header: shaped.header, keep: shaped.keep, types: shaped.types,
             cells: idx.map(function (i) { return shaped.cells[i]; }),
             display: idx.map(function (i) { return shaped.display[i]; }),
             sortKeys: idx.map(function (i) { return shaped.sortKeys[i]; }) };
  }
  /** keep only the given header names (in that order) */
  function project(shaped, names) {
    var ix = names.map(function (n) { return shaped.header.indexOf(n); }).filter(function (i) { return i >= 0; });
    var pick = function (row) { return ix.map(function (i) { return row[i]; }); };
    return { header: ix.map(function (i) { return shaped.header[i]; }), keep: ix.map(function (i) { return shaped.keep[i]; }),
             types: ix.map(function (i) { return shaped.types[i]; }), cells: shaped.cells.map(pick),
             display: shaped.display.map(pick), sortKeys: shaped.sortKeys.map(pick) };
  }
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
  /** the rows of the "About this export" sheet */
  function aboutRows(info) {
    var rows = [['About this export'], []];
    rows.push(['Workbook', info.workbook]); rows.push(['Table', info.table]); rows.push(['Exported', info.exported]);
    rows.push(['Rows', info.rows]); rows.push(['Columns', info.columns]);
    if (info.sort) rows.push(['Sorted by', info.sort]);
    if (info.asof) rows.push(['Data as of', info.asof]);
    rows.push([]); rows.push(['Dashboard filters']);
    if (info.chips.length) info.chips.forEach(function (c) { rows.push([c.label, c.value]); }); else rows.push(['(none)']);
    rows.push([]); rows.push(['Table filters']);
    if (info.tableFilters.length) info.tableFilters.forEach(function (f) { rows.push([f[0], f[1]]); }); else rows.push(['(none)']);
    if (info.search) rows.push(['Search', info.search]);
    return rows;
  }
  /** SheetJS workbook: the data sheet (+ optional About sheet) */
  function buildWorkbook(XLSX, shaped, sheetName, about) {
    var aoa = [shaped.header.map(function (h) { return { t: 's', v: h }; })].concat(shaped.cells);
    var ws = XLSX.utils.aoa_to_sheet(aoa, { cellDates: false });
    ws['!cols'] = shaped.header.map(function (h, k) {
      var w = h.length;
      for (var r = 0; r < Math.min(shaped.cells.length, 500); r++) {
        var c = shaped.cells[r][k]; if (!c) continue;
        var len = c.z ? 11 : (typeof c.v === 'number' ? String(Math.round(c.v * 100) / 100).length : String(c.v).length);
        if (len > w) w = len;
      }
      return { wch: Math.min(Math.max(w + 2, 8), 60) };
    });
    if (shaped.cells.length) ws['!autofilter'] = { ref: XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: shaped.cells.length, c: shaped.header.length - 1 } }) };
    ws['!freeze'] = { xSplit: 1, ySplit: 1 };
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, (sheetName || 'Data').replace(/[\\\/\?\*\[\]:]/g, '').slice(0, 31) || 'Data');
    if (about && about.length) {
      var aws = XLSX.utils.aoa_to_sheet(about.map(function (r) { return r.map(function (v) { return v === undefined || v === null ? '' : v; }); }));
      aws['!cols'] = [{ wch: 22 }, { wch: 70 }];
      XLSX.utils.book_append_sheet(wb, aws, 'About this export');
    }
    return wb;
  }
  function buildCsv(shaped) {
    function esc(s) { s = String(s); return /[",\r\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s; }
    function cell(c, k) {
      if (c === null) return '';
      if (c.z && typeof c.v === 'number') {
        var d = fromSerial(c.v);
        var iso = d.getUTCFullYear() + '-' + p2(d.getUTCMonth() + 1) + '-' + p2(d.getUTCDate());
        return shaped.types[k] === 'date-time' ? iso + ' ' + p2(d.getUTCHours()) + ':' + p2(d.getUTCMinutes()) : iso;
      }
      return esc(c.v);
    }
    var lines = [shaped.header.map(esc).join(',')];
    for (var r = 0; r < shaped.cells.length; r++) lines.push(shaped.cells[r].map(cell).join(','));
    return '﻿' + lines.join('\r\n');
  }
  function stamp(d) { d = d || new Date(); return d.getFullYear() + '-' + p2(d.getMonth() + 1) + '-' + p2(d.getDate()) + ' ' + p2(d.getHours()) + p2(d.getMinutes()); }
  /** "Work Order Volume · Americas · 2026 · 2026-09-15 1412" -- filter values, Windows-safe, bounded */
  function fileName(prefix, chips, when) {
    var parts = [prefix];
    chips.filter(function (c) { return !c.info; }).slice(0, 3).forEach(function (c) { parts.push(c.value); });
    parts.push(stamp(when));
    return parts.join(' · ').replace(/[\\\/:*?"<>|]/g, '-').replace(/\s+/g, ' ').slice(0, 150);
  }
  /** Tableau Filter objects -> chip descriptors */
  function describeFilters(filters, labels, hide) {
    var out = [];
    (filters || []).forEach(function (f) {
      if (!f || hide.indexOf(f.fieldName) >= 0) return;
      var label = labels[f.fieldName] || f.fieldName, value = null;
      if (f.filterType === 'categorical') {
        if (f.isAllSelected) return;
        var vals = (f.appliedValues || []).map(function (v) { return v.formattedValue !== undefined ? String(v.formattedValue) : String(v.value); });
        if (!vals.length) return;
        value = (f.isExcludeMode ? 'not ' : '') + (vals.length <= 2 ? vals.join(', ') : vals.slice(0, 2).join(', ') + ' +' + (vals.length - 2));
      } else if (f.filterType === 'range') {
        var lo = f.minValue && !isNull(f.minValue) ? f.minValue.formattedValue : null, hi = f.maxValue && !isNull(f.maxValue) ? f.maxValue.formattedValue : null;
        if (lo === null && hi === null) return;
        value = (lo === null ? '…' : lo) + ' – ' + (hi === null ? '…' : hi);
      } else if (f.filterType === 'relative-date') {
        value = (f.rangeType || 'relative') + ' ' + (f.rangeN !== undefined ? f.rangeN + ' ' : '') + (f.periodType || '');
      } else return;
      out.push({ field: f.fieldName, label: label, value: value, info: false });
    });
    return out;
  }
  /** one-row aggregate table -> {lowercased clean name: DataValue} */
  function aggMap(table) {
    var out = {};
    if (!table || !table.data || !table.data.length) return out;
    table.columns.forEach(function (c) { out[cleanName(c.fieldName).toLowerCase()] = table.data[0][c.index]; });
    return out;
  }
  function aggGet(map, key) {          // tolerant lookup: exact clean name, then substring
    var k = key.toLowerCase();
    if (map[k]) return map[k];
    for (var n in map) if (n.indexOf(k) >= 0) return map[n];
    return null;
  }
  /** dim + measure table -> [[label, number], ...] in the table's order */
  function series(table) {
    if (!table || !table.data || !table.columns) return [];
    var cols = table.columns, dim = -1, meas = -1;
    cols.forEach(function (c) {
      var isAgg = AGG_RE.test(c.fieldName) || /^agg /i.test(cleanName(c.fieldName));
      if (isAgg && meas < 0) meas = c.index; else if (!isAgg && dim < 0) dim = c.index;
    });
    if (dim < 0 || meas < 0) { if (cols.length >= 2) { dim = cols[0].index; meas = cols[1].index; } else return []; }
    return table.data.map(function (r) { var d = r[dim], m = r[meas]; return [isNull(d) ? '(blank)' : String(d.formattedValue !== undefined ? d.formattedValue : d.value), num(m) || 0]; });
  }

  var api = { cleanName: cleanName, serial: serial, toSerial: toSerial, fromSerial: fromSerial, fmtDate: fmtDate, shape: shape, subset: subset,
              project: project, filterIndices: filterIndices, sortIndices: sortIndices, buildWorkbook: buildWorkbook, buildCsv: buildCsv,
              aboutRows: aboutRows, fileName: fileName, describeFilters: describeFilters, aggMap: aggMap, aggGet: aggGet, series: series,
              DEFAULTS: DEFAULTS, stamp: stamp, kv: kv, list: list };
  if (typeof module !== 'undefined' && module.exports) { module.exports = api; return; }
  root.WoExport = api;

  // ================================================================ browser / Tableau side
  var $ = function (id) { return document.getElementById(id); };
  var cfg = {}, S = { shaped: null, filters: [], search: '', sortCol: -1, sortDir: 0, page: 1, pageSize: 100, view: [], bounds: [],
                      busy: false, suppress: false, sheet: null, total: 0, preview: 1000, sorted: false,
                      agg: {}, aggAll: {}, byQ: [], byR: [], byPL: [], byS: [], chips: [], period: null, periodParam: null,
                      colOrder: null, hidden: {}, labels: {}, help: {}, cancel: false, insOpen: true, worker: null };

  function setting(k) { var v = tableau.extensions.settings.get(k); return (v === undefined || v === null || v === '') ? DEFAULTS[k] : v; }
  function status(msg, kind) { var el = $('status'); el.textContent = msg || ''; el.className = kind || ''; }
  function esc(s) { return String(s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }
  function store(k, v) { try { localStorage.setItem('wov:' + cfg.sheet + ':' + k, JSON.stringify(v)); } catch (e) { /* private mode */ } }
  function recall(k, d) { try { var v = localStorage.getItem('wov:' + cfg.sheet + ':' + k); return v === null ? d : JSON.parse(v); } catch (e) { return d; } }
  function busy(on) { S.busy = on; ['dl', 'first', 'prev', 'next', 'last'].forEach(function (id) { $(id).disabled = on; }); if (!on) renderPager(); }
  var toastTimer = null;
  function toast(html, ms) { var t = $('toast'); t.innerHTML = html; t.hidden = false; clearTimeout(toastTimer); toastTimer = setTimeout(function () { t.hidden = true; }, ms || 5000); }

  function sheetByName(name) {
    if (!name) return null;
    var ws = tableau.extensions.dashboardContent.dashboard.worksheets;
    for (var i = 0; i < ws.length; i++) if (ws[i].name === name) return ws[i];
    return null;
  }
  function readAll(sheet, limit, onProgress, cancelled) {
    var opts = { ignoreSelection: true };
    if (typeof sheet.getSummaryDataReaderAsync !== 'function') {
      return sheet.getSummaryDataAsync(opts).then(function (t) { return { columns: t.columns, rows: limit ? t.data.slice(0, limit) : t.data, total: t.data.length }; });
    }
    var page = limit ? Math.min(limit, 10000) : 10000;
    return sheet.getSummaryDataReaderAsync(page, opts).then(function (reader) {
      var columns = null, rows = [], total = reader.totalRowCount, chain = Promise.resolve();
      var need = limit ? limit : Infinity, pages = Math.min(reader.pageCount, limit ? Math.ceil(limit / page) : reader.pageCount);
      for (var p = 0; p < pages; p++) (function (p) {
        chain = chain.then(function () {
          if (cancelled && cancelled()) throw new Error('cancelled');
          return reader.getPageAsync(p).then(function (t) {
            if (!columns) columns = t.columns;
            for (var i = 0; i < t.data.length && rows.length < need; i++) rows.push(t.data[i]);
            if (onProgress) onProgress(rows.length, total);
          });
        });
      })(p);
      var done = function () { return reader.releaseAsync().catch(function () {}); };
      return chain.then(function () { return done().then(function () { return { columns: columns || [], rows: rows, total: total }; }); },
                        function (e) { return done().then(function () { throw e; }); });
    });
  }
  function readTable(name) {
    var ws = sheetByName(name);
    if (!ws) return Promise.resolve(null);
    return ws.getSummaryDataAsync({ ignoreSelection: true, maxRows: 2000 }).catch(function () { return null; });
  }

  // ---------------------------------------------------------------- filters / chips
  // findParameterAsync matches the parameter's DISPLAY name (its caption, e.g. 'Period'), not the XML name
  // ('Parameter 2'). The setting normally carries the caption; if it does not resolve, scan every dashboard
  // parameter and accept a case-insensitive match on name or id (ids look like '[Parameters].[Parameter 2]').
  function paramKey(s) { return String(s || '').replace(/^\[Parameters\]\./i, '').replace(/^\[|\]$/g, '').trim().toLowerCase(); }
  function findParam(name) {
    var dash = tableau.extensions.dashboardContent.dashboard;
    if (!name || !dash) return Promise.resolve(null);
    var direct = typeof dash.findParameterAsync === 'function' ? dash.findParameterAsync(name).catch(function () { return null; }) : Promise.resolve(null);
    return direct.then(function (p) {
      if (p) return p;
      if (typeof dash.getParametersAsync !== 'function') return null;
      return dash.getParametersAsync().then(function (all) {
        var want = paramKey(name);
        for (var i = 0; i < (all || []).length; i++) if (paramKey(all[i].name) === want || paramKey(all[i].id) === want) return all[i];
        return null;
      }).catch(function () { return null; });
    });
  }
  function refreshFilters() {
    var hide = String(cfg.hideFilters || '').split(';').map(function (s) { return s.trim(); });
    var p = S.periodParam ? Promise.resolve(S.periodParam) : findParam(cfg.periodParam);
    return Promise.all([S.sheet.getFiltersAsync().catch(function () { return []; }), p]).then(function (r) {
      S.chips = describeFilters(r[0], S.labels, hide);
      if (r[1]) {
        S.periodParam = r[1];
        var cv = r[1].currentValue;
        S.chips.unshift({ field: '', label: 'Period', value: String(cv.formattedValue !== undefined ? cv.formattedValue : cv.value), info: true });
      }
      renderChips();
    });
  }
  function renderChips() {
    var el = $('chiplist');
    if (!S.chips.length) { el.innerHTML = '<span class="none">No dashboard filters &mdash; every row in the period</span>'; $('clear').hidden = true; return; }
    el.innerHTML = S.chips.map(function (c, i) {
      return '<span class="chip' + (c.info ? ' info' : '') + '" title="' + esc(c.info ? 'Change the period in the rail' : c.label + ': ' + c.value) + '">' + esc(c.label) + ': <b>' + esc(c.value) + '</b>'
           + (c.info ? '' : '<button class="x" data-i="' + i + '" title="Remove this filter">&times;</button>') + '</span>';
    }).join('');
    $('clear').hidden = !S.chips.some(function (c) { return !c.info; });
  }
  function clearFilter(field) {
    return S.sheet.clearFilterAsync(field).catch(function (e) { toast('Could not clear that filter: ' + esc(e && e.message ? e.message : e)); });
  }
  function clearAll() {
    var fs = S.chips.filter(function (c) { return !c.info; }).map(function (c) { return c.field; });
    status('Clearing ' + fs.length + ' filter' + (fs.length === 1 ? '' : 's') + '…', 'warn');
    return fs.reduce(function (p, f) { return p.then(function () { return clearFilter(f); }); }, Promise.resolve());
  }

  // ---------------------------------------------------------------- aggregates / insight band
  function refreshAgg() {
    return Promise.all([readTable(cfg.aggSheet), readTable(cfg.aggAllSheet), readTable(cfg.byQuarterSheet),
                        readTable(cfg.byRegionSheet), readTable(cfg.byPLSheet), readTable(cfg.byStatusSheet)])
      .then(function (t) {
        S.agg = aggMap(t[0]); S.aggAll = aggMap(t[1]);
        S.byQ = series(t[2]); S.byR = series(t[3]).sort(function (a, b) { return b[1] - a[1]; }); S.byPL = series(t[4]).sort(function (a, b) { return b[1] - a[1]; }); S.byS = series(t[5]).sort(function (a, b) { return b[1] - a[1]; });
        var n = num(aggGet(S.agg, 'agg count'));
        if (n !== null) S.total = n;
        renderInsights(); renderAsOf(); renderPager();
      });
  }
  function dateOf(dv) { if (!dv || isNull(dv)) return null; var s = toSerial(dv.value, dv.nativeValue); return s === null ? null : s; }
  function renderAsOf() {
    var mx = dateOf(aggGet(S.aggAll, 'agg max date')) || dateOf(aggGet(S.agg, 'agg max date'));
    var n = num(aggGet(S.agg, 'agg count'));
    $('asof').innerHTML = (n !== null ? fmtInt(n) + ' rows match' : '') + (mx !== null ? (n !== null ? ' &middot; ' : '') + 'data as of <b>' + fmtDate(mx) + '</b>' : '');
  }
  function tile(l, v, d) { return '<div class="tile"><div class="l">' + l + '</div><div class="v">' + v + '</div>' + (d || '') + '</div>'; }
  function renderInsights() {
    var A = S.agg, n = num(aggGet(A, 'agg count')), all = num(aggGet(S.aggAll, 'agg count'));
    var wosPage = !!cfg.byStatusSheet;
    var rtk = num(aggGet(A, 'agg rtk')), trn = num(aggGet(A, 'agg trn')), transit = num(aggGet(A, 'agg transit'));
    var cust = num(aggGet(A, 'agg customers')), prod = num(aggGet(A, 'agg products')), pls = num(aggGet(A, 'agg pls')), wos = num(aggGet(A, 'agg wos'));
    var dmin = dateOf(aggGet(A, 'agg min date')), dmax = dateOf(aggGet(A, 'agg max date'));
    var topPL = S.byPL.length ? S.byPL[0] : null;
    if (n === null) { $('summary').innerHTML = '<span style="color:var(--faint)">Insights unavailable &mdash; the aggregate sheets are not on this dashboard</span>'; $('insrow').innerHTML = ''; return; }
    var chipTxt = S.chips.filter(function (c) { return !c.info; }).map(function (c) { return esc(c.value); }).join(' &middot; ');
    var bits = ['<b>' + fmtInt(n) + ' ' + esc(cfg.grain) + '</b>'];
    if (chipTxt) bits.push(chipTxt);
    if (dmin !== null && dmax !== null) bits.push(fmtDate(dmin) + ' &rarr; ' + fmtDate(dmax));
    if (!wosPage && rtk !== null) bits.push('<b>' + pct(rtk, n, 1) + ' RTK</b>');
    if (trn !== null) bits.push('<b>' + pct(trn, n) + ' transhipped</b>');
    if (cust !== null) bits.push(fmtInt(cust) + ' customers');
    if (topPL) bits.push('top line ' + esc(topPL[0]));
    $('summary').innerHTML = bits.join(' &middot; ');
    // Five tiles (four on the service-line page) and three minis share the row fluidly -- see index.html.
    // The product-line count rides on the "Top product lines" mini instead of a sixth tile.
    var tiles = '', ofAll = all !== null ? pct(n, all, 0) + ' of ' + fmtInt(all) + ' all time' : '';
    if (wosPage) {
      var comp = S.byS.filter(function (s) { return /^completed$/i.test(s[0]); }).reduce(function (a, s) { return a + s[1]; }, 0);
      var canc = S.byS.filter(function (s) { return /cancel/i.test(s[0]); }).reduce(function (a, s) { return a + s[1]; }, 0);
      tiles += tile('Service lines', fmtInt(n), '<div class="d" title="' + esc(ofAll) + '">' + ofAll + '</div>');
      tiles += tile('Work orders', fmtInt(wos), '<div class="d">' + (wos ? (n / wos).toFixed(2) + ' lines per WO' : '') + '</div>');
      tiles += tile('Completed', pct(comp, n), '<div class="d">' + fmtInt(canc) + ' cancelled</div>');
    } else {
      tiles += tile('Work orders', fmtInt(n), '<div class="d" title="' + esc(ofAll) + '">' + ofAll + '</div>');
      tiles += tile('RTK / RTF', rtk === null ? '–' : pct(rtk, n, 1), rtk === null ? '' : '<div class="split"><i style="width:' + (100 * rtk / n) + '%;background:var(--teal)" title="RTK ' + fmtInt(rtk) + '"></i><i style="flex:1;background:var(--indigo)" title="RTF ' + fmtInt(n - rtk) + '"></i></div>');
      tiles += tile('Transhipped', trn === null ? '–' : pct(trn, n), '<div class="d">' + (trn === null ? '' : fmtInt(trn) + ' cross-border') + '</div>');
      tiles += tile('Avg transit', transit === null ? '–' : transit.toFixed(1) + '<small>days</small>', '<div class="d">WDF &rarr; WDB</div>');
    }
    tiles += tile('Customers', fmtInt(cust), '<div class="d">' + (prod !== null ? fmtInt(prod) + ' products' : '') + '</div>');
    var minis = '';
    if (S.byQ.length) {
      var qmax = Math.max.apply(null, S.byQ.map(function (q) { return q[1]; })) || 1;
      minis += '<div class="mini"><div class="l">By fiscal quarter</div><div class="cols">' + S.byQ.map(function (q, i) {
        var cur = i === S.byQ.length - 1;
        return '<i class="' + (cur ? 'cur lbl' : '') + '" data-v="' + fmtInt(q[1]) + '" style="height:' + Math.max(4, 40 * q[1] / qmax) + 'px" title="' + esc(q[0]) + ': ' + fmtInt(q[1]) + '"></i>';
      }).join('') + '</div><div class="xl">' + S.byQ.map(function (q) { return '<span title="' + esc(q[0]) + '">' + esc(String(q[0]).replace(/^FY\d\d\s*/, '')) + '</span>'; }).join('') + '</div></div>';
    }
    var hbars = function (title, data, alt, showPct) {
      if (!data.length) return '';
      var mx = data[0][1] || 1;
      return '<div class="mini"><div class="l" title="' + esc(title) + '">' + esc(title) + '</div><div class="hb">' + data.slice(0, 4).map(function (r) {
        return '<div title="' + esc(r[0]) + ': ' + fmtInt(r[1]) + '"><em>' + esc(r[0]) + '</em><span class="tr"><i class="' + (alt ? 'alt' : '') + '" style="width:' + Math.max(4, 100 * r[1] / mx) + '%"></i></span><s>' + (showPct ? pct(r[1], n) : fmtInt(r[1])) + '</s></div>';
      }).join('') + '</div></div>';
    };
    minis += wosPage ? hbars('By status', S.byS, false, true) : hbars('Work done by region', S.byR, false, true);
    minis += hbars(pls !== null ? 'Top of ' + fmtInt(pls) + ' product lines' : 'Top product lines', S.byPL, true, false);
    $('insrow').innerHTML = tiles + minis;
  }
  function setInsights(open) {
    S.insOpen = open; $('insrow').hidden = !open; $('instoggle').textContent = open ? 'Hide insights ▲' : 'Show insights ▼'; store('ins', open);
  }

  // ---------------------------------------------------------------- columns
  function visibleColumns() {
    var all = S.shaped.header, order = (S.colOrder || all).filter(function (h) { return all.indexOf(h) >= 0; });
    all.forEach(function (h) { if (order.indexOf(h) < 0) order.push(h); });
    return order.filter(function (h) { return !S.hidden[h]; });
  }
  function initColumns() {
    var saved = recall('cols', null);
    if (saved && saved.order) { S.colOrder = saved.order; S.hidden = saved.hidden || {}; return; }
    S.colOrder = null; S.hidden = {};
    list(cfg.defaultColumns).length && S.shaped.header.forEach(function (h) { if (list(cfg.defaultColumns).indexOf(h) < 0) S.hidden[h] = true; });
  }
  function saveColumns() { store('cols', { order: S.colOrder || S.shaped.header.slice(), hidden: S.hidden }); }
  function renderColBtn() { if (S.shaped) $('colcount').textContent = visibleColumns().length + '/' + S.shaped.header.length; }
  function openColumns() {
    closePops();
    var pop = document.createElement('div'); pop.className = 'pop'; pop.id = 'colpick';
    var draw = function (q) {
      var order = (S.colOrder || S.shaped.header.slice()).filter(function (h) { return S.shaped.header.indexOf(h) >= 0; });
      S.shaped.header.forEach(function (h) { if (order.indexOf(h) < 0) order.push(h); });
      S.colOrder = order;
      pop.innerHTML = '<h4>Columns <span>&middot; ' + visibleColumns().length + ' of ' + S.shaped.header.length + ' shown</span></h4>'
        + '<input class="s" id="colq" placeholder="Find a column" value="' + esc(q || '') + '"><div class="list">'
        + order.filter(function (h) { return !q || h.toLowerCase().indexOf(q.toLowerCase()) >= 0; }).map(function (h) {
            return '<label class="opt" draggable="true" data-h="' + esc(h) + '"><span class="h">&#8942;&#8942;</span><input type="checkbox" ' + (S.hidden[h] ? '' : 'checked') + '><span class="nm">' + esc(h) + '</span></label>';
          }).join('')
        + '</div><div class="ft"><button id="colall">Show all</button><button id="colreset">Reset to default</button><span class="hint">drag &#8942;&#8942; to reorder</span></div>';
      pop.querySelector('#colq').addEventListener('input', function () { var v = this.value, pos = this.selectionStart; draw(v); var i = pop.querySelector('#colq'); i.focus(); i.setSelectionRange(pos, pos); });
      pop.querySelector('#colall').addEventListener('click', function () { S.hidden = {}; saveColumns(); draw(''); applyView(false); renderHead(); });
      pop.querySelector('#colreset').addEventListener('click', function () { try { localStorage.removeItem('wov:' + cfg.sheet + ':cols'); } catch (e) {} initColumns(); draw(''); applyView(false); renderHead(); });
      pop.querySelectorAll('input[type=checkbox]').forEach(function (cb) {
        cb.addEventListener('change', function () {
          var h = this.closest('.opt').getAttribute('data-h');
          if (this.checked) delete S.hidden[h]; else if (visibleColumns().length > 1) S.hidden[h] = true; else { this.checked = true; return; }
          saveColumns(); pop.querySelector('h4 span').innerHTML = '&middot; ' + visibleColumns().length + ' of ' + S.shaped.header.length + ' shown';
          renderHead(); applyView(false);
        });
      });
      var dragging = null;
      pop.querySelectorAll('.opt').forEach(function (el) {
        el.addEventListener('dragstart', function (e) { dragging = el.getAttribute('data-h'); el.classList.add('drag'); e.dataTransfer.effectAllowed = 'move'; });
        el.addEventListener('dragend', function () { el.classList.remove('drag'); });
        el.addEventListener('dragover', function (e) { e.preventDefault(); });
        el.addEventListener('drop', function (e) {
          e.preventDefault(); var to = el.getAttribute('data-h'); if (!dragging || dragging === to) return;
          var o = S.colOrder.slice(); o.splice(o.indexOf(dragging), 1); o.splice(o.indexOf(to), 0, dragging); S.colOrder = o;
          saveColumns(); draw(pop.querySelector('#colq').value); renderHead(); applyView(false);
        });
      });
    };
    draw('');
    $('app').appendChild(pop);
    watchOutside(pop, $('colbtn'));
  }
  var outsideHandler = null;
  function watchOutside(pop, anchor) {
    if (outsideHandler) document.removeEventListener('mousedown', outsideHandler, true);
    outsideHandler = function (e) { if (!pop.contains(e.target) && !anchor.contains(e.target) && !S.busy) closePops(); };
    setTimeout(function () { if (outsideHandler) document.addEventListener('mousedown', outsideHandler, true); }, 0);
  }
  function closePops() {
    ['colpick', 'export'].forEach(function (id) { var p = $(id); if (p) p.remove(); });
    if (outsideHandler) { document.removeEventListener('mousedown', outsideHandler, true); outsideHandler = null; }
  }

  // ---------------------------------------------------------------- table rendering
  function bounds(shaped) {
    return shaped.types.map(function (t, c) {
      if (!isDate(t)) return null;
      var lo = null, hi = null;
      for (var r = 0; r < shaped.sortKeys.length; r++) {
        var v = shaped.sortKeys[r][c]; if (v === null) continue;
        v = Math.floor(v); if (lo === null || v < lo) lo = v; if (hi === null || v > hi) hi = v;
      }
      var full = null;
      if (cfg.dateField && shaped.header[c] === cfg.dateField) {
        var a = dateOf(aggGet(S.agg, 'agg min date')), b = dateOf(aggGet(S.agg, 'agg max date'));
        if (a !== null && b !== null) full = { lo: Math.floor(Math.min(a, b)), hi: Math.floor(Math.max(a, b)) };
      }
      return full || (lo === null ? null : { lo: lo, hi: hi });
    });
  }
  function renderHead() {
    if (!S.shaped) return;
    var vis = visibleColumns(), h = S.shaped.header, out = '<tr>';
    vis.forEach(function (name) {
      var c = h.indexOf(name), arr = S.sortCol === c ? '<span class="arrow">' + (S.sortDir > 0 ? '▲' : '▼') + '</span>' : '';
      var help = S.help[name] ? '<span class="q" title="' + esc(S.help[name]) + '">?</span>' : '';
      out += '<th data-c="' + c + '" title="' + esc((S.help[name] ? S.help[name] + ' — ' : '') + 'click to sort') + '">' + esc(name) + arr + help + '</th>';
    });
    out += '</tr><tr class="filters">';
    vis.forEach(function (name) {
      var f = h.indexOf(name), b = S.bounds[f];
      if (b && b.hi > b.lo) {
        var fl = S.filters[f], from = fl.from === null || fl.from === undefined ? b.lo : fl.from, to = fl.to === null || fl.to === undefined ? b.hi : fl.to;
        out += '<th><div class="rng"><div class="lb"><span id="lo' + f + '">' + fmtDate(from) + '</span><span id="hi' + f + '">' + fmtDate(to) + '</span></div><div class="tr">'
             + '<input type="range" data-c="' + f + '" data-e="lo" min="' + b.lo + '" max="' + b.hi + '" value="' + from + '">'
             + '<input type="range" data-c="' + f + '" data-e="hi" min="' + b.lo + '" max="' + b.hi + '" value="' + to + '"></div></div></th>';
      } else {
        out += '<th><input type="text" size="1" class="' + (S.filters[f].text ? 'on' : '') + '" data-c="' + f + '" placeholder="Filter value" value="' + esc(S.filters[f].text || '') + '" autocomplete="off"></th>';
      }
    });
    $('head').innerHTML = out + '</tr>';
    renderColBtn();
  }
  function cellHtml(name, v, t) {
    if (!v) return '';
    if (/^RTF \/ RTK$/i.test(name)) return '<span class="tag ' + (v === 'RTK' ? 'a' : 'b') + '">' + esc(v) + '</span>';
    if (/^PL Detail$/i.test(name) && /^Review needed/i.test(v)) return '<span class="tag rev" title="' + esc(v) + '">Review needed</span>';
    return esc(v);
  }
  function renderBody() {
    var vis = visibleColumns(), h = S.shaped.header, ix = vis.map(function (n) { return h.indexOf(n); });
    var start = (S.page - 1) * S.pageSize, end = Math.min(start + S.pageSize, S.view.length), rows = [];
    for (var i = start; i < end; i++) {
      var d = S.shaped.display[S.view[i]], tds = '';
      for (var k = 0; k < ix.length; k++) { var c = ix[k]; tds += '<td class="' + (isNum(S.shaped.types[c]) ? 'num' : '') + '" title="' + esc(d[c]) + '">' + cellHtml(h[c], d[c], S.shaped.types[c]) + '</td>'; }
      rows.push('<tr>' + tds + '</tr>');
    }
    $('body').innerHTML = rows.join('');
    var em = $('empty'); if (em) em.remove();
    if (S.view.length === 0) renderEmpty();
  }
  function renderEmpty() {
    var el = document.createElement('div'); el.id = 'empty';
    var tf = activeTableFilters();
    if (S.shaped.display.length === 0) {
      var ch = S.chips.filter(function (c) { return !c.info; });
      el.innerHTML = '<div class="big">No ' + esc(cfg.grain) + ' match the dashboard filters</div><div class="why">'
        + (ch.length ? esc(ch.map(function (c) { return c.label + ': ' + c.value; }).join(' · ')) + '. Remove a filter, or widen the period in the rail.' : 'Widen the period in the rail.') + '</div>'
        + (ch.length ? '<button class="btn primary" id="emptyclear">Clear the dashboard filters</button>' : '');
      $('scroll').appendChild(el);
      var b = $('emptyclear'); if (b) b.addEventListener('click', clearAll);
    } else {
      el.innerHTML = '<div class="big">No rows match the filters on this table</div><div class="why">'
        + esc(tf.map(function (f) { return f[0] + ' “' + f[1] + '”'; }).join(' · ')) + (S.search ? (tf.length ? ' · ' : '') + 'search “' + esc(S.search) + '”' : '')
        + '. The table filters narrow what the dashboard already returned — check the spelling, or clear them.</div>'
        + '<button class="btn primary" id="emptyclear">Clear the table filters</button>';
      $('scroll').appendChild(el);
      $('emptyclear').addEventListener('click', clearTableFilters);
    }
  }
  function activeTableFilters() {
    var out = [];
    S.filters.forEach(function (f, c) {
      if (!f) return;
      if (f.text) out.push([S.shaped.header[c], f.text]);
      if ((f.from !== null && f.from !== undefined) || (f.to !== null && f.to !== undefined)) {
        var b = S.bounds[c] || {};
        out.push([S.shaped.header[c], fmtDate(f.from === null || f.from === undefined ? b.lo : f.from) + ' – ' + fmtDate(f.to === null || f.to === undefined ? b.hi : f.to)]);
      }
    });
    return out;
  }
  function clearTableFilters() {
    S.search = ''; $('search').value = '';
    S.filters = S.shaped.header.map(function () { return { text: '', from: null, to: null }; });
    renderHead(); applyView(true);
  }
  function renderPager() {
    var pages = Math.max(1, Math.ceil(S.view.length / S.pageSize));
    if (S.page > pages) S.page = pages;
    $('pageinfo').textContent = S.view.length ? 'Page ' + S.page + ' of ' + pages : '–';
    $('first').disabled = $('prev').disabled = S.busy || S.page <= 1;
    $('next').disabled = $('last').disabled = S.busy || S.page >= pages;
    var loaded = S.shaped ? S.shaped.display.length : 0, all = Math.max(S.total || 0, loaded);
    var sortTxt = S.shaped && S.sortCol >= 0 ? '<span class="sort">sorted by ' + esc(S.shaped.header[S.sortCol]) + ' ' + (S.sortDir > 0 ? '↑' : '↓') + '</span>' : '';
    var tf = S.shaped ? activeTableFilters().length + (S.search ? 1 : 0) : 0;
    if (!S.shaped) $('count').innerHTML = '&nbsp;';
    else if (tf) $('count').innerHTML = '<b>' + fmtInt(S.view.length) + '</b> of ' + fmtInt(loaded) + ' loaded rows match the table filters' + sortTxt;
    else $('count').innerHTML = all > loaded ? '<b>' + fmtInt(loaded) + '</b> of ' + fmtInt(all) + ' rows loaded' + sortTxt : '<b>' + fmtInt(loaded) + '</b> rows' + sortTxt;
    $('count').title = all > loaded ? fmtInt(loaded) + ' of ' + fmtInt(all) + ' rows are loaded for browsing. Pick a bigger Load size to see more. The download is not limited by this.'
                                   : fmtInt(all) + ' row' + (all === 1 ? '' : 's') + ' — everything the dashboard filters return.';
    $('dllabel').textContent = cfg.label;
  }
  function applyView(resetPage) {
    S.view = sortIndices(S.shaped, filterIndices(S.shaped, S.filters, S.search), S.sortCol, S.sortDir);
    if (resetPage !== false) S.page = 1;
    renderBody(); renderPager();
  }
  function skeleton() {
    var n = S.shaped ? visibleColumns().length : 8, rows = [];
    if (!S.shaped) $('head').innerHTML = '<tr>' + Array.from({ length: n }, function () { return '<th style="min-width:110px">&nbsp;</th>'; }).join('') + '</tr>';
    for (var i = 0; i < 22; i++) rows.push('<tr class="skel">' + Array.from({ length: n }, function (_, j) { return '<td><i style="width:' + (40 + ((i * 7 + j * 13) % 45)) + '%"></i></td>'; }).join('') + '</tr>');
    $('body').innerHTML = rows.join('');
  }

  // ---------------------------------------------------------------- load
  function load(first) {
    if (S.busy) return Promise.resolve();
    busy(true); if (first) skeleton();
    status(first ? 'Loading…' : 'Updating…', 'warn');
    var limit = S.preview >= Number(cfg.paramAll || Infinity) ? 0 : S.preview;
    return Promise.all([refreshFilters(), refreshAgg(),
                        readAll(S.sheet, limit, function (n, t) { status('Loading rows… ' + fmtInt(n) + (t ? ' of ' + fmtInt(t) : ''), 'warn'); })])
      .then(function (r) {
        var data = r[2], was = S.shaped ? S.shaped.header.join('') : null;
        S.shaped = shape(data.columns, data.rows, cfg);
        if (!(S.total > 0) || S.total < data.rows.length) S.total = data.total === undefined ? data.rows.length : Math.max(data.total, S.total || 0);
        if (was !== S.shaped.header.join('')) {
          S.filters = S.shaped.header.map(function () { return { text: '', from: null, to: null }; });
          S.sortCol = -1; S.sortDir = 0; initColumns();
        }
        S.bounds = bounds(S.shaped);
        S.filters.forEach(function (f, c) { var b = S.bounds[c]; if (!b) return;
          if (f.from !== null && f.from !== undefined) f.from = Math.max(b.lo, Math.min(f.from, b.hi));
          if (f.to !== null && f.to !== undefined) f.to = Math.max(b.lo, Math.min(f.to, b.hi)); });
        if (!S.sorted && cfg.sortBy) { var c = S.shaped.header.indexOf(cfg.sortBy); if (c >= 0) { S.sortCol = c; S.sortDir = String(cfg.sortDir).toLowerCase() === 'asc' ? 1 : -1; } S.sorted = true; }
        renderHead(); applyView(true); status('');
      })
      .catch(function (e) { status('Could not read the table: ' + (e && e.message ? e.message : e), 'error'); })
      .then(function () { busy(false); });
  }
  function setPreview(n) {
    if (S.busy) { setTimeout(function () { setPreview(n); }, 250); return; }
    S.preview = n; store('preview', n);
    if (!cfg.param) { load(false); return; }
    busy(true); S.suppress = true;
    findParam(cfg.param).then(function (p) { return p ? p.changeValueAsync(String(n || cfg.paramAll)) : null; })
      .then(function () { return new Promise(function (r) { setTimeout(r, 400); }); })
      .then(function () { S.suppress = false; busy(false); return load(false); }, function (e) { S.suppress = false; busy(false); status('Could not change the load size: ' + (e && e.message ? e.message : e), 'error'); });
  }

  // ---------------------------------------------------------------- download
  function exportPanel() {
    closePops();
    var pop = document.createElement('div'); pop.className = 'pop'; pop.id = 'export';
    var loaded = S.shaped.display.length, all = Math.max(S.total || 0, loaded), vis = visibleColumns().length, total = S.shaped.header.length;
    var st = { fmt: recall('fmt', 'xlsx'), rows: 'all', cols: 'all', about: true };
    if (all > 250000 && st.fmt === 'xlsx') st.fmt = 'csv';
    var draw = function () {
      var nrows = st.rows === 'all' ? all : S.view.length, ncols = st.cols === 'all' ? total : vis;
      var mb = (nrows * ncols * (st.fmt === 'csv' ? 14 : 5)) / 1048576, secs = Math.max(1, Math.round(nrows / 12000));
      var tooBig = st.fmt === 'xlsx' && nrows > 1048575;
      var name = fileName(cfg.filename, S.chips, new Date()) + (st.fmt === 'csv' ? '.csv' : '.xlsx');
      pop.innerHTML = '<h4>Download</h4>'
        + '<div class="grp">Format</div>'
        + '<label class="opt"><input type="radio" name="fmt" value="xlsx" ' + (st.fmt === 'xlsx' ? 'checked' : '') + '>Excel workbook (.xlsx)<span class="sub">real dates and numbers</span></label>'
        + '<label class="opt"><input type="radio" name="fmt" value="csv" ' + (st.fmt === 'csv' ? 'checked' : '') + '>CSV<span class="sub">lighter, for very large exports</span></label>'
        + '<div class="grp">Rows</div>'
        + '<label class="opt"><input type="radio" name="rows" value="all" ' + (st.rows === 'all' ? 'checked' : '') + '>Everything the filters return<span class="sub">' + fmtInt(all) + ' rows</span></label>'
        + '<label class="opt"><input type="radio" name="rows" value="loaded" ' + (st.rows === 'loaded' ? 'checked' : '') + '>What is loaded on this page<span class="sub">' + fmtInt(S.view.length) + ' rows</span></label>'
        + '<div class="grp">Columns</div>'
        + '<label class="opt"><input type="radio" name="cols" value="all" ' + (st.cols === 'all' ? 'checked' : '') + '>All ' + total + ' columns</label>'
        + '<label class="opt"><input type="radio" name="cols" value="shown" ' + (st.cols === 'shown' ? 'checked' : '') + '>The ' + vis + ' columns shown</label>'
        + '<label class="opt" style="margin-top:4px"><input type="checkbox" id="about" ' + (st.about ? 'checked' : '') + '>Add an “About this export” sheet<span class="sub">filters · rows · data as of</span></label>'
        + '<div class="file"><b>' + esc(name) + '</b><div class="m' + (tooBig || (st.fmt === 'xlsx' && nrows > 250000) ? ' warn' : '') + '">' + fmtInt(nrows) + ' rows × ' + ncols + ' columns · about ' + (mb < 1 ? Math.max(1, Math.round(mb * 10)) / 10 : Math.round(mb)) + ' MB · ' + (secs < 5 ? 'a few seconds' : 'about ' + secs + ' s')
        + (tooBig ? ' — Excel cannot hold more than 1,048,575 rows; use CSV' : (st.fmt === 'xlsx' && nrows > 250000 ? ' — large for Excel; CSV is faster' : '')) + '</div></div>'
        + '<div id="progwrap"></div>'
        + '<div class="actions"><button class="btn primary" id="go" ' + (tooBig ? 'disabled' : '') + '>Download ' + fmtInt(nrows) + ' rows</button><button class="cancel" id="cancel">Cancel</button></div>';
      pop.querySelectorAll('input[name=fmt]').forEach(function (r) { r.addEventListener('change', function () { st.fmt = this.value; store('fmt', st.fmt); draw(); }); });
      pop.querySelectorAll('input[name=rows]').forEach(function (r) { r.addEventListener('change', function () { st.rows = this.value; draw(); }); });
      pop.querySelectorAll('input[name=cols]').forEach(function (r) { r.addEventListener('change', function () { st.cols = this.value; draw(); }); });
      pop.querySelector('#about').addEventListener('change', function () { st.about = this.checked; });
      pop.querySelector('#cancel').addEventListener('click', function () { if (S.busy) { S.cancel = true; status('Cancelling\u2026', 'warn'); } closePops(); });
      pop.querySelector('#go').addEventListener('click', function () { runExport(st, pop, name); });
    };
    draw();
    $('app').appendChild(pop);
    watchOutside(pop, $('dl'));
  }
  function progress(pop, txt, right, frac) {
    var w = pop.querySelector('#progwrap'); if (!w) return;
    w.innerHTML = '<div class="prog"><div class="t"><span>' + txt + '</span><span>' + (right || '') + '</span></div><div class="bar"><i style="width:' + Math.round(100 * (frac || 0)) + '%"></i></div></div>';
    var go = pop.querySelector('#go'); if (go) go.disabled = true;
  }
  function runExport(st, pop, name) {
    if (S.busy) return;
    busy(true); S.cancel = false;
    var t0 = Date.now(), colsWanted = st.cols === 'all' ? null : visibleColumns();
    if (st.rows !== 'loaded') progress(pop, 'Reading rows\u2026', '', 0);
    var work = st.rows === 'loaded'
      ? Promise.resolve(subset(S.shaped, S.view))
      : readAll(S.sheet, 0, function (n, t) { progress(pop, 'Reading rows… <b>' + fmtInt(n) + '</b>' + (t ? ' of ' + fmtInt(t) : ''), t ? 'about ' + Math.max(1, Math.round((t - n) / 15000)) + ' s left' : '', t ? n / t : 0); },
                function () { return S.cancel; })
          .then(function (data) { var full = shape(data.columns, data.rows, cfg); return subset(full, sortIndices(full, filterIndices(full, S.filters, S.search), S.sortCol, S.sortDir)); });
    work.then(function (out) {
      if (S.cancel) throw new Error('cancelled');
      if (colsWanted) out = project(out, colsWanted);
      progress(pop, 'Building the file… (' + fmtInt(out.cells.length) + ' rows)', '', 1);
      var about = st.about ? aboutRows({
        workbook: cfg.workbook, table: cfg.sheet, exported: stamp(), rows: out.cells.length, columns: out.header.length,
        sort: S.sortCol >= 0 ? S.shaped.header[S.sortCol] + (S.sortDir > 0 ? ' ascending' : ' descending') : '',
        asof: (function () { var mx = dateOf(aggGet(S.aggAll, 'agg max date')); return mx === null ? '' : fmtDate(mx); })(),
        chips: S.chips, tableFilters: activeTableFilters(), search: S.search }) : null;
      return (st.fmt === 'csv' ? Promise.resolve(new Blob([buildCsv(out) + (about ? '' : '')], { type: 'text/csv;charset=utf-8' }))
                               : buildXlsx(out, about));
    }).then(function (blob) {
      var a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = name;
      document.body.appendChild(a); a.click();
      setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 4000);
      closePops();
      var mb = blob.size / 1048576;
      toast('Downloaded <b>' + esc(name) + '</b> · ' + (mb < 1 ? Math.round(blob.size / 1024) + ' KB' : mb.toFixed(1) + ' MB') + ' · ' + Math.round((Date.now() - t0) / 1000) + ' s', 7000);
      status('');
    }).catch(function (e) {
      closePops();
      if (e && e.message === 'cancelled') { status('Download cancelled', 'warn'); toast('Download cancelled', 3000); }
      else status('Download failed: ' + (e && e.message ? e.message : e), 'error');
    }).then(function () { busy(false); });
  }
  /** SheetJS in a Worker so a big export does not freeze the page; falls back to the main thread */
  function buildXlsx(out, about) {
    var payload = { header: out.header, cells: out.cells, sheetName: cfg.sheet, about: about };
    return new Promise(function (resolve, reject) {
      var w = null;
      try { w = new Worker('worker.js'); } catch (e) { w = null; }
      if (!w) return resolve(mainThread());
      var settled = false;
      w.onmessage = function (ev) { settled = true; w.terminate(); if (ev.data && ev.data.error) resolve(mainThread()); else resolve(new Blob([ev.data.buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })); };
      w.onerror = function () { if (!settled) { settled = true; w.terminate(); resolve(mainThread()); } };
      try { w.postMessage(payload); } catch (e) { if (!settled) { settled = true; resolve(mainThread()); } }
    });
    function mainThread() {
      return new Blob([XLSX.write(buildWorkbook(XLSX, out, cfg.sheet, about), { bookType: 'xlsx', type: 'array', compression: true })],
                      { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    }
  }

  // ---------------------------------------------------------------- events
  function wire() {
    var timer = null;
    $('search').addEventListener('input', function () { S.search = this.value; clearTimeout(timer); timer = setTimeout(function () { applyView(true); }, 180); });
    $('search').addEventListener('keydown', function (e) { if (e.key === 'Escape') { this.value = ''; S.search = ''; applyView(true); this.blur(); } });
    $('head').addEventListener('input', function (e) {
      var el = e.target, c = +el.getAttribute('data-c');
      if (el.type === 'text') { S.filters[c].text = el.value; el.classList.toggle('on', !!el.value); clearTimeout(timer); timer = setTimeout(function () { applyView(true); }, 180); }
      else if (el.type === 'range') {
        var b = S.bounds[c], which = el.getAttribute('data-e'), v = +el.value;
        var lo = S.filters[c].from === null || S.filters[c].from === undefined ? b.lo : S.filters[c].from;
        var hi = S.filters[c].to === null || S.filters[c].to === undefined ? b.hi : S.filters[c].to;
        if (which === 'lo') { lo = Math.min(v, hi); el.value = lo; } else { hi = Math.max(v, lo); el.value = hi; }
        S.filters[c].from = lo <= b.lo ? null : lo; S.filters[c].to = hi >= b.hi ? null : hi;
        $('lo' + c).textContent = fmtDate(lo); $('hi' + c).textContent = fmtDate(hi);
        clearTimeout(timer); timer = setTimeout(function () { applyView(true); }, 120);
      }
    });
    $('head').addEventListener('click', function (e) {
      var th = e.target.closest ? e.target.closest('th') : null;
      if (!th || !th.hasAttribute('data-c') || th.parentNode.className === 'filters' || e.target.classList.contains('q')) return;
      var c = +th.getAttribute('data-c');
      if (S.sortCol !== c) { S.sortCol = c; S.sortDir = 1; } else if (S.sortDir === 1) S.sortDir = -1; else { S.sortCol = -1; S.sortDir = 0; }
      renderHead(); applyView(true);
    });
    $('body').addEventListener('dblclick', function (e) {
      var td = e.target.closest ? e.target.closest('td') : null; if (!td) return;
      var v = td.getAttribute('title') || td.textContent;
      if (navigator.clipboard && v) navigator.clipboard.writeText(v).then(function () { toast('Copied <b>' + esc(v) + '</b>', 2000); }, function () {});
    });
    $('chiplist').addEventListener('click', function (e) { var b = e.target.closest ? e.target.closest('button.x') : null; if (b) clearFilter(S.chips[+b.getAttribute('data-i')].field); });
    $('clear').addEventListener('click', clearAll);
    $('first').addEventListener('click', function () { S.page = 1; renderBody(); renderPager(); });
    $('prev').addEventListener('click', function () { S.page--; renderBody(); renderPager(); });
    $('next').addEventListener('click', function () { S.page++; renderBody(); renderPager(); });
    $('last').addEventListener('click', function () { S.page = Math.max(1, Math.ceil(S.view.length / S.pageSize)); renderBody(); renderPager(); });
    $('size').addEventListener('change', function () { S.pageSize = +this.value; store('size', S.pageSize); S.page = 1; renderBody(); renderPager(); });
    $('preview').addEventListener('change', function () { setPreview(+this.value); });
    $('colbtn').addEventListener('click', function () { if ($('colpick')) closePops(); else if (S.shaped) openColumns(); });
    $('dl').addEventListener('click', function () { if ($('export')) closePops(); else if (S.shaped) exportPanel(); });
    $('instoggle').addEventListener('click', function () { setInsights(!S.insOpen); });
    document.addEventListener('keydown', function (e) {
      var inInput = /input|select|textarea/i.test(e.target.tagName);
      if (e.key === '/' && !inInput) { e.preventDefault(); $('search').focus(); }
      else if (e.key === 'Escape') { closePops(); }
      else if (!inInput && e.key === 'ArrowRight' && !$('next').disabled) { S.page++; renderBody(); renderPager(); }
      else if (!inInput && e.key === 'ArrowLeft' && !$('prev').disabled) { S.page--; renderBody(); renderPager(); }
    });
  }

  document.addEventListener('DOMContentLoaded', function () {
    var started = false;
    skeleton(); status('Starting\u2026', 'warn');
    setTimeout(function () { if (!started) { $('hosting').style.display = 'flex'; } }, 6000);
    tableau.extensions.initializeAsync().then(function () {
      started = true;
      Object.keys(DEFAULTS).forEach(function (k) { cfg[k] = setting(k); });
      S.labels = kv(cfg.filterLabels); S.help = kv(cfg.help);
      S.pageSize = Number(recall('size', cfg.pageSize)) || 100;
      S.preview = Math.max(0, Number(recall('preview', cfg.previewRows)) || 0);
      $('preview').value = String(S.preview);
      if ($('preview').selectedIndex < 0) $('preview').insertAdjacentHTML('afterbegin', '<option selected value="' + S.preview + '">' + S.preview.toLocaleString() + '</option>');
      $('size').value = String(S.pageSize);
      if (!Array.prototype.some.call($('size').options, function (o) { return o.value === String(S.pageSize); })) $('size').insertAdjacentHTML('afterbegin', '<option selected>' + S.pageSize + '</option>');
      $('dllabel').textContent = cfg.label;
      setInsights(recall('ins', true) !== false);
      S.sheet = sheetByName(cfg.sheet);
      if (!S.sheet) { status('Worksheet "' + cfg.sheet + '" is not on this dashboard', 'error'); return; }
      wire();
      try {
        S.sheet.addEventListener(tableau.TableauEventType.SummaryDataChanged, function () { if (!S.suppress) load(false); });
        S.sheet.addEventListener(tableau.TableauEventType.FilterChanged, function () { if (!S.suppress) refreshFilters(); });
      } catch (e) { /* older API */ }
      return load(true).then(function () {
        if (S.periodParam && tableau.TableauEventType.ParameterChanged) {
          try { S.periodParam.addEventListener(tableau.TableauEventType.ParameterChanged, function () { refreshFilters(); }); } catch (e) { /* ignore */ }
        }
      });
    }).catch(function (e) { status('Could not start: ' + (e && e.message ? e.message : e), 'error'); });
  });
})(typeof window !== 'undefined' ? window : this);
