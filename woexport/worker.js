/* worker.js -- builds the .xlsx off the main thread so a large export does not freeze the table.
 * Receives {header, cells, sheetName, about}; replies {buf} (transferred) or {error}. */
importScripts('lib/xlsx.full.min.js');
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
onmessage = function (e) {
  try {
    var d = e.data;
    var buf = XLSX.write(buildWorkbook(XLSX, { header: d.header, cells: d.cells }, d.sheetName, d.about), { bookType: 'xlsx', type: 'array', compression: true });
    postMessage({ buf: buf }, [buf]);
  } catch (err) {
    postMessage({ error: String(err && err.message ? err.message : err) });
  }
};
