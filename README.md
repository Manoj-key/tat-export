# TAT Excel Export — Tableau dashboard extension

A small Tableau dashboard extension that exports one worksheet of a dashboard to Excel with **real data types**:
dates as Excel dates, numbers as numbers, clean column names (no `SUM(...)`), plus an optional CSV export.

It runs entirely in the viewer's browser using Tableau's Extensions API — it reads only the summary data already
shown in the view and writes the file locally. **No data is sent anywhere**, and this site has no server code.

## Files
| file | what it is |
| --- | --- |
| `index.html` | the extension's page (the button you see in the dashboard) |
| `app.js` | the export logic — data-type conversion, workbook building, "all rows" handling |
| `lib/tableau.extensions.1.17.0.min.js` | Tableau's official Extensions API runtime |
| `lib/xlsx.full.min.js` | SheetJS Community Edition 0.18.5 (Apache-2.0) |

## Use
Hosted with GitHub Pages; the dashboard points at `index.html`. Opening that address directly in a browser shows
"Hosting OK" — the button only does anything inside a Tableau dashboard.

Licences for the bundled libraries are in `lib/`.
