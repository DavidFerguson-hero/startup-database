"""
Startup Scout — desktop launcher.

Starts the Flask server in a background thread, waits until it is ready,
then opens the user's default browser at http://localhost:5001.

Works both in normal Python (dev) and as a PyInstaller-frozen bundle.
"""
import json
import os
import sys
import socket
import urllib.request
import threading
import time
import webbrowser

# ── Frozen-bundle path setup ─────────────────────────────────────────────────
# When frozen by PyInstaller, sys._MEIPASS is the temp directory where all
# bundled files are extracted.  Add it to sys.path so `import app` works.
if getattr(sys, 'frozen', False):
    _bundle_dir = sys._MEIPASS          # extracted bundle (read-only)
    _exe_dir    = os.path.dirname(sys.executable)  # directory next to the .exe/.app
    sys.path.insert(0, _bundle_dir)
else:
    _bundle_dir = os.path.dirname(os.path.abspath(__file__))
    _exe_dir    = _bundle_dir
    sys.path.insert(0, _bundle_dir)

# ── .env loader ───────────────────────────────────────────────────────────────
# Check next to the executable first (for end users who create a .env there),
# then fall back to the bundle/source directory.
# The desktop app's settings file (API key, optional DATA_DIR) lives in ~/Documents/Startup Scout
# so it survives installing a new version of the app. Earlier files don't override later ones.
_env_candidates = [os.path.join(_exe_dir, '.env'), os.path.join(_bundle_dir, '.env')]
if getattr(sys, 'frozen', False):
    _env_candidates.insert(0, os.path.join(os.path.expanduser('~/Documents'), 'Startup Scout', '.env'))
for _env_candidate in _env_candidates:
    if os.path.exists(_env_candidate):
        with open(_env_candidate) as _f:
            for _line in _f:
                _line = _line.strip()
                if _line and not _line.startswith('#') and '=' in _line:
                    _k, _v = _line.split('=', 1)
                    os.environ.setdefault(_k.strip(), _v.strip().strip('"\''))

# ── Config ───────────────────────────────────────────────────────────────────
PORT = int(os.environ.get('PORT', 5001))
HOST = '127.0.0.1'
URL  = f'http://{HOST}:{PORT}'

# Quit after this many seconds with no page open. Generous because browsers throttle
# timers in background tabs. On for the desktop app; set STARTUP_SCOUT_AUTO_QUIT to enable in dev.
AUTO_QUIT_AFTER = int(os.environ.get('STARTUP_SCOUT_AUTO_QUIT') or (180 if getattr(sys, 'frozen', False) else 0))


def _wait_for_server(timeout: int = 30) -> bool:
    """Poll localhost:PORT until it accepts connections, or timeout expires."""
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            with socket.create_connection((HOST, PORT), timeout=1):
                return True
        except OSError:
            time.sleep(0.25)
    return False


def _is_startup_scout(port: int) -> bool:
    try:
        with urllib.request.urlopen(f'http://{HOST}:{port}/api/heartbeat', timeout=2) as r:
            return json.load(r).get('app') == 'startup-scout'
    except Exception:
        return False


def _port_in_use(port: int) -> bool:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        return s.connect_ex((HOST, port)) == 0


def _run_flask():
    """Import and run the Flask app (runs forever in its thread)."""
    from app import app as flask_app
    flask_app.run(host=HOST, port=PORT, debug=False, use_reloader=False)


def _quit_when_idle():
    import app as app_module
    while True:
        time.sleep(min(15, AUTO_QUIT_AFTER / 3))
        if app_module.seconds_since_heartbeat() > AUTO_QUIT_AFTER:
            os._exit(0)


if __name__ == '__main__':
    # Already running? Just show it rather than starting a second copy.
    if _is_startup_scout(PORT):
        webbrowser.open(URL)
        sys.exit(0)
    # Something else (e.g. an older version without a heartbeat) holds the port: use the next free one.
    while _port_in_use(PORT):
        PORT += 1
    URL = f'http://{HOST}:{PORT}'

    # Start Flask in a daemon thread so it dies when the process exits
    server_thread = threading.Thread(target=_run_flask, daemon=True)
    server_thread.start()

    # Wait for Flask to be ready, then open the browser (opened anyway on timeout so the user sees the error)
    _wait_for_server(timeout=30)
    webbrowser.open(URL)

    if AUTO_QUIT_AFTER:
        threading.Thread(target=_quit_when_idle, daemon=True).start()

    # Keep the main thread alive (the daemon thread will exit when this does)
    try:
        server_thread.join()
    except KeyboardInterrupt:
        pass
