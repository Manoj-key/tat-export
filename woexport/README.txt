WO Volume Table - dashboard extension for the WO Volume Analytics workbook
=========================================================================
The two raw-data pages of the workbook are this small web page, loaded by Tableau into the dashboard:
row count, Preview size, All rows, paging, Search table, Excel and CSV export with real dates and
numbers, a filter box under every column, a date range slider, sortable headers.  Same engine as the
TAT Excel Export extension on the TAT & SOT workbook.  Nothing leaves the PC.

A Tableau extension is a web page, so these files need a web address.  Two ways to give them one:

------------------------------------------------------------------------------------------------------
A.  TEST NOW ON THIS PC  (two minutes, no GitHub)
    1. Unzip this folder anywhere, e.g. Documents\wo-export.
    2. Double-click  serve-local.bat  -> a window says "served at http://localhost:8766/index.html".
       Leave it open.  (Port 8766, so it does not collide with the TAT extension's 8765.)
    3. Open  "WO_Volume_Analytics (localhost test).twbx"  -> allow the extension when Tableau asks.
    Both raw-data pages render.  This only works on a PC where the window is open.

------------------------------------------------------------------------------------------------------
B.  PERMANENT  (the shipped workbook points here)
    The workbook expects   https://manoj-key.github.io/tat-export/wo-export/index.html
    i.e. a  wo-export  FOLDER inside the tat-export repository that already serves the TAT extension.
    1. In the tat-export repo, add a folder named  wo-export  and put these files in it:
         index.html   app.js   lib/tableau.extensions.1.17.0.min.js   lib/xlsx.full.min.js
       (the .trex, .ps1, .bat and this README can go in too; they are harmless.)
    2. Commit / push.  GitHub Pages republishes in about a minute.
    3. Check: open the address above in a browser -> the toolbar appears and after a few seconds says
       "Hosting OK - this table only works inside the Tableau dashboard".
    4. Open  WO_Volume_Analytics.twbx  -> allow the extension when Tableau asks.
    No new repository and no Pages settings are needed: the tat-export repo is already published.

    If you would rather host it somewhere else, say where -- the address is one build parameter and the
    workbook is re-baked in a minute.

------------------------------------------------------------------------------------------------------
NOTES
  * Tableau asks each user to allow the extension once.  If the workbook is ever published to Tableau
    Server/Cloud, the site admin adds the address to the extension safe list (summary data only,
    Full Data Access = Deny).
  * Settings (worksheet, column order, file name, sort) live in the workbook, one set per page;
    defaults are at the top of app.js.
  * The Preview box bounds how many rows are loaded into the page.  Tableau's own row cap (the
    'Rows shown' parameter on the TAT page) is not wired yet -- it needs the TAT workbook's Top-N
    block; the export is complete either way.
  * Bundled in lib\: Tableau's official Extensions API runtime (@tableau/tabextsandbox 1.17.0) and
    SheetJS 0.18.5 (Apache-2.0).  Nothing is downloaded at run time.
