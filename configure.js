/* KGSO Tableau extensions -- the Configure dialog (opened from the extension zone's menu in Tableau).
 * Lists the apps in apps.js, saves the chosen key into the workbook setting `app` and closes; the launcher
 * then restarts the page with the chosen app. A workbook generated with the setting never needs this. */
(function () {
  'use strict';
  var APPS = window.KGSO_APPS || {}, DEFAULT_APP = window.KGSO_DEFAULT_APP || 'tatexport';
  var list = document.getElementById('list'), msg = document.getElementById('msg');
  var save = document.getElementById('save'), cancel = document.getElementById('cancel');
  function say(m, cls) { msg.textContent = m; msg.className = cls || ''; }
  function chosen() { var r = document.querySelector('input[name=app]:checked'); return r ? r.value : ''; }
  function mark() { Array.prototype.forEach.call(list.querySelectorAll('label.app'), function (l) {
    l.className = 'app' + (l.querySelector('input').checked ? ' on' : ''); }); }
  function render(current) {
    list.innerHTML = '';
    Object.keys(APPS).forEach(function (key) {
      var a = APPS[key], l = document.createElement('label'), r = document.createElement('input');
      l.className = 'app'; r.type = 'radio'; r.name = 'app'; r.value = key; r.checked = key === current;
      var b = document.createElement('b'); b.textContent = a.name;
      var s = document.createElement('span'); s.textContent = a.about || '';
      var c = document.createElement('code'); c.textContent = 'app = ' + key + '  (' + a.base + ')';
      l.appendChild(r); l.appendChild(b); l.appendChild(s); l.appendChild(c); list.appendChild(l);
      r.addEventListener('change', mark);
    });
    mark();
  }
  if (!window.tableau || !window.tableau.extensions) {
    render(DEFAULT_APP); save.disabled = true;
    say('Hosting OK - this dialog only opens from the extension menu inside a Tableau dashboard');
    return;
  }
  var ext = window.tableau.extensions;
  ext.initializeDialogAsync().then(function (payload) {
    var current = String(payload || ext.settings.get('app') || DEFAULT_APP).trim().toLowerCase();
    render(APPS[current] ? current : DEFAULT_APP);
    save.addEventListener('click', function () {
      var key = chosen();
      if (!key) { say('Choose a table first', 'error'); return; }
      save.disabled = true; say('Saving…');
      ext.settings.set('app', key);
      ext.settings.saveAsync().then(function () { ext.ui.closeDialog(key); })
        .catch(function (e) { save.disabled = false; say('Could not save: ' + (e && e.message ? e.message : e), 'error'); });
    });
    cancel.addEventListener('click', function () { ext.ui.closeDialog(''); });
  }).catch(function (e) { say('Could not start: ' + (e && e.message ? e.message : e), 'error'); });
})();
