/* KGSO Tableau extensions -- the registry.
 *
 * One line per extension. The key is the value of the workbook setting `app`; `base` is the folder next to this
 * file that holds the extension (its own index.html, scripts and libraries). To add a new extension: upload its
 * folder, add one line here, point the new workbook at the same index.html with `app` set to the new key.
 * The safe-list entry (the URL of index.html) never changes.
 */
window.KGSO_APPS = {
  tatexport: { base: 'tatexport/', name: 'TAT & SOT Excel Export',
               about: 'The Raw Data Table of the TAT & SOT dashboard: per-column filters, search, paging, Excel / CSV export.' },
  wovolume:  { base: 'woexport/',  name: 'WO Volume Table',
               about: 'The download pages of WO Volume Analytics: filter chips, insight band, per-column filters, Excel / CSV export.' }
};
window.KGSO_DEFAULT_APP = 'tatexport';       // a workbook without the setting gets the TAT & SOT table (the older workbook)
