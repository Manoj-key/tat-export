# KGSO Tableau extensions

One hosted page for every dashboard extension of the Global Ops TAT & SOT Performance team (KGSO):

    https://manoj-key.github.io/tat-export/index.html

A workbook's extension zone points at that page and carries a setting `app` that names the table it wants.
The page (`launcher.js`) initialises the Tableau Extensions API once, reads the setting and mounts the app
from its own folder. One Tableau Server safe-list entry therefore covers every workbook, every update, and
every extension added later.

| `app`              | folder       | used by                                   |
|--------------------|--------------|-------------------------------------------|
| `tatexport` (default) | `tatexport/` | TAT and SOT Charts — the Raw Data Table  |
| `wovolume`         | `woexport/`  | WO Volume Analytics — the two download pages |

## Layout

    index.html            the launcher page (the only URL Tableau ever needs)
    launcher.js           reads `app`, mounts the folder, registers Configure…
    apps.js               the registry: one line per extension
    configure.html / .js  the Configure… dialog (pick the app for a dashboard that was not generated with the setting)
    lib/                  Tableau's Extensions API runtime, 1.17.0 (@tableau/tabextsandbox)
    tatexport/            the TAT & SOT Excel Export table (index.html, app.js, lib/)
    woexport/             the WO Volume table (index.html, app.js, worker.js, lib/)
    KGSO Extensions.trex  the manifest, for adding the extension to a new dashboard by hand
    serve-local.bat/.ps1  a loopback web server (port 8766) for testing in Tableau Desktop without GitHub

## Adding an extension later

1. Upload its folder (with its own `index.html` and scripts) next to the others.
2. Add one line to `apps.js`.
3. Point the new workbook's extension zone at the same `index.html` with the setting `app` = the new key
   (or add `KGSO Extensions.trex` to the dashboard and choose the app under Configure…).

No new safe-list entry: the URL has not changed.

## Updating an extension

Replace the files in its folder and commit. GitHub Pages republishes in about a minute; the next view load
runs the new version. No workbook change, no republish, no approval.

## Checks

Open `index.html` in a browser: after a few seconds it says *Hosting OK - this page only works inside a
Tableau dashboard*. `configure.html` lists the registered apps. After loading, the apps make no network
calls of their own; the launcher's one request is the same-origin fetch of the chosen app's `index.html`.
