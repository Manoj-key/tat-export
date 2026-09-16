/* KGSO Tableau extensions -- the launcher.
 *
 * One hosted page, one safe-list entry, every KGSO dashboard extension. A workbook's extension zone points at
 * this page and carries a setting `app` naming the extension it wants (the list is apps.js); the launcher
 * initialises the Tableau Extensions API once, reads that setting, and mounts the named app from its own
 * folder next to this file:
 *
 *   app = tatexport   (or absent)   -> tatexport/   the TAT & SOT Excel Export table   (the TAT & SOT workbook)
 *   app = wovolume                  -> woexport/    the WO Volume table                (the WO Volume Analytics workbook)
 *
 * Mounting = fetch the app's own index.html from the same folder tree (same origin, part of loading the page),
 * put its styles and body into this document, then load its scripts in order. The apps are untouched: they boot
 * on DOMContentLoaded (re-dispatched here) and call initializeAsync (answered with the launcher's own, already
 * resolved, promise). Nothing else happens on the network after that; the apps make no calls of their own.
 *
 * Configure: the launcher registers a configure callback, so the zone's menu in Tableau offers "Configure..."
 * and configure.html lets the author pick the app for a dashboard that was not generated with the setting.
 * Saving a different app restarts the page, which mounts the newly chosen one.
 */
(function () {
  'use strict';
  var APPS = window.KGSO_APPS || {}, DEFAULT_APP = window.KGSO_DEFAULT_APP || 'tatexport';
  var boot = document.getElementById('boot');
  function say(msg) {                                            // the boot line, re-created if a mount replaced the body
    if (!boot || !document.body.contains(boot)) {
      boot = document.createElement('div'); boot.id = 'boot'; document.body.appendChild(boot);
    }
    boot.textContent = msg;
  }
  function norm(v) { return String(v || DEFAULT_APP).trim().toLowerCase(); }
  function loadScript(src) {
    return new Promise(function (resolve, reject) {
      var s = document.createElement('script');
      s.src = src; s.onload = resolve; s.onerror = function () { reject(new Error('cannot load ' + src)); };
      document.body.appendChild(s);
    });
  }
  function mount(base) {
    return fetch(base + 'index.html', { cache: 'no-cache' }).then(function (r) {
      if (!r.ok) throw new Error(r.status + ' loading ' + base + 'index.html');
      return r.text();
    }).then(function (html) {
      var doc = new DOMParser().parseFromString(html, 'text/html');
      var scripts = [];
      Array.prototype.forEach.call(doc.querySelectorAll('script'), function (s) {
        var src = s.getAttribute('src');
        if (src && !/tableau\.extensions/.test(src)) scripts.push(src);    // the API runtime is already here
        s.parentNode.removeChild(s);
      });
      Array.prototype.forEach.call(doc.querySelectorAll('style'), function (s) { document.head.appendChild(s.cloneNode(true)); });
      if (doc.title) document.title = doc.title;
      window.KGSO = { base: base, ready: true };               // apps: assets (e.g. a worker) live under `base`
      document.body.innerHTML = doc.body.innerHTML;
      return scripts.reduce(function (p, src) { return p.then(function () { return loadScript(base + src); }); }, Promise.resolve());
    }).then(function () {
      // the app registered its DOMContentLoaded handler while its script ran; the real event fired long ago
      document.dispatchEvent(new Event('DOMContentLoaded'));
    });
  }
  if (!window.tableau || !window.tableau.extensions) { say('Hosting OK - this page only works inside a Tableau dashboard'); return; }
  var ext = window.tableau.extensions, current = null, started = false;
  setTimeout(function () { if (!started) say('Hosting OK - this page only works inside a Tableau dashboard'); }, 6000);
  function restart(next) { if (norm(next) !== current) window.location.reload(); }
  function configure() {                                         // Tableau: the zone menu's "Configure..."
    var url = window.location.href.replace(/[#?].*$/, '').replace(/[^\/]*$/, '') + 'configure.html';
    return ext.ui.displayDialogAsync(url, current || DEFAULT_APP, { width: 460, height: 330 })
      .then(function (chosen) { if (chosen) restart(chosen); })  // the dialog saved the setting; belt and braces
      .catch(function () { /* closed without saving */ });
  }
  var init = ext.initializeAsync({ configure: configure });
  ext.initializeAsync = function () { return init; };            // the mounted app's own call gets the same promise
  init.then(function () {
    started = true;
    current = norm(ext.settings.get('app'));
    if (ext.settings.addEventListener && window.tableau.TableauEventType) {
      ext.settings.addEventListener(window.tableau.TableauEventType.SettingsChanged, function (ev) {
        restart(ev && ev.newSettings ? ev.newSettings.app : ext.settings.get('app'));
      });
    }
    var app = APPS[current];
    if (!app) { say('Unknown app "' + current + '" - the workbook setting must be one of: ' + Object.keys(APPS).join(', ')); return; }
    return mount(app.base);
  }).catch(function (e) { say('Could not start: ' + (e && e.message ? e.message : e)); });
})();
