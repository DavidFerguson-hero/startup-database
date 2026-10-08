// Tells the desktop app this page is still open; it quits a few minutes after the last page closes.
(function () {
  var failures = 0, banner = null;

  function showClosed() {
    if (banner) return;
    banner = document.createElement('div');
    banner.textContent = 'Startup Scout has closed. Open the app again to keep working.';
    banner.style.cssText = 'position:fixed;top:0;left:0;right:0;z-index:99999;padding:10px 16px;' +
      'background:#991B1B;color:#fff;font:600 13px system-ui,sans-serif;text-align:center';
    document.body.appendChild(banner);
  }

  function beat() {
    fetch('/api/heartbeat', { method: 'POST', cache: 'no-store' })
      .then(function (r) {
        if (!r.ok) throw new Error(r.status);
        failures = 0;
        if (banner) { banner.remove(); banner = null; }
      })
      .catch(function () { if (++failures >= 2) showClosed(); });
  }

  beat();
  setInterval(beat, 20000);
  document.addEventListener('visibilitychange', function () { if (!document.hidden) beat(); });
})();
