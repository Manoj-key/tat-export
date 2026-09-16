WO Volume Table v4 - dashboard extension for the WO Volume Analytics workbook
============================================================================
The two download pages of the workbook are this small web page, loaded by Tableau into the dashboard:
filter chips with Clear all, the insight band (exact numbers from Tableau's own aggregate sheets),
row count, Load size, paging, Columns chooser, Search, one Download button with a panel (Excel / CSV,
all rows or loaded, all columns or shown, an "About this export" sheet, a file name that carries the
filters, an estimate, progress and Cancel), a filter box under every column, a date range slider,
sortable headers, and unambiguous dates (01 Sep 2026).  The Excel file is built in a Web Worker so a
large export never freezes the page.  Nothing leaves the PC.

FILES
  index.html    the page and its styles
  app.js        the engine; this workbook's defaults are at the top
  worker.js     NEW in v3 - the background Excel build (must sit next to index.html)
  lib\tableau.extensions.1.17.0.min.js   Tableau's official runtime (@tableau/tabextsandbox)
  lib\xlsx.full.min.js                   SheetJS 0.18.5 (Apache-2.0)
  WO Volume Table.trex                   manifest, id com.keysight.kgso.wovolumetable
  serve-local.ps1 / .bat                 a loopback web server for testing on this PC

------------------------------------------------------------------------------------------------------
A.  PERMANENT  (the shipped workbook points here)
    The workbook expects   https://manoj-key.github.io/tat-export/woexport/index.html
    That folder already exists in the tat-export repository and holds v2.
    1. Replace its contents with these files.  index.html and app.js changed; worker.js is new; the two
       lib\ files are unchanged.  (The .trex, .ps1, .bat and this README are harmless there.)
    2. Commit / push.  GitHub Pages republishes in about a minute.
    3. Check: open the address above in a browser -> the page appears and after a few seconds says
       "Hosting OK - this table only works inside the Tableau dashboard".
    4. Open  WO_Volume_Analytics.twbx  -> allow the extension when Tableau asks.

------------------------------------------------------------------------------------------------------
B.  TEST NOW ON THIS PC  (two minutes, no GitHub)
    1. Unzip this folder anywhere, e.g. Documents\wo-export.
    2. Double-click  serve-local.bat  -> a window says "served at http://localhost:8766/index.html".
       Leave it open.  (Port 8766, so it does not collide with the TAT extension's 8765.)
    3. Open  "WO_Volume_Analytics (localhost test).twbx"  -> allow the extension when Tableau asks.
    Both download pages render.  This only works on a PC where the window is open.

------------------------------------------------------------------------------------------------------
NOTES
  * Tableau asks each user to allow the extension once.  On Tableau Server/Cloud the site admin adds
    the address to the extension safe list (summary data only, Full Data Access = Deny) and the prompt
    disappears for everyone.
  * Settings live in the workbook, one set per page: the table worksheet, the aggregate worksheets the
    insight band reads (aggSheet, aggAllSheet, byQuarterSheet, byRegionSheet, byPLSheet, byStatusSheet),
    the column list, file name, sort, date format, the period parameter (by its caption, "Period"),
    the filter never shown as a chip (In Period) and the filter labels / ? definitions.
  * The Load box bounds how many rows are transferred into the page for browsing; the download always
    reads everything the filters return.  The Period control in the rail is the real row cap: it is a
    context filter, so Tableau itself only builds the window.
  * Column set and order, page size, insight band open/closed and the last format are remembered per
    user in the browser (localStorage), per page.
