WO Volume Table - the dashboard extension of the WO Volume Analytics workbook
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
  worker.js     the background Excel build (must sit next to index.html)
  lib\tableau.extensions.1.17.0.min.js   Tableau's official runtime (@tableau/tabextsandbox)
  lib\xlsx.full.min.js                   SheetJS 0.18.5 (Apache-2.0)

HOW IT IS LOADED (since v8)
  The workbook does not point at this folder. It points at the KGSO launcher one level up,
      https://manoj-key.github.io/tat-export/index.html
  with the workbook setting  app = wovolume.  The launcher initialises the Tableau API once, reads the
  setting and mounts this folder's index.html and app.js into its own page (worker.js and lib\ are
  loaded from here).  The TAT & SOT workbook uses the same launcher with no setting, which mounts
  tatexport\.  One URL on the Tableau Server safe list covers both, and any extension added later.
  This folder still works on its own (open woexport/index.html directly) for testing.

UPDATING
  Replace the files in this folder and commit.  GitHub Pages republishes in about a minute; the next
  view load runs the new version.  No workbook change, no republish, no new approval.

NOTES
  * On Tableau Server/Cloud the site admin adds the launcher URL above to the extension safe list
    (Full Data Access = No, User prompts = Hide) once; the prompt disappears for everyone.
  * Settings live in the workbook, one set per page: the table worksheet, the aggregate worksheets the
    insight band reads (aggSheet, aggAllSheet, byQuarterSheet, byRegionSheet, byPLSheet, byStatusSheet),
    the column list, file name, sort, date format, the period parameter (by its caption, "Period"),
    the filter never shown as a chip (In Period), the filter labels / ? definitions, and app.
  * The Load box bounds how many rows are transferred into the page for browsing; the download always
    reads everything the filters return.  The Period control in the rail is the real row cap: it is a
    context filter, so Tableau itself only builds the window.
  * Column set and order, page size, insight band open/closed and the last format are remembered per
    user in the browser (localStorage), per page.
