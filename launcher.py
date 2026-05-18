"""
Startup Scout — desktop launcher.

Starts the Flask server in a background thread, waits until it is ready,
then opens the user's default browser at http://localhost:5001.

Works both in normal Python (dev) and as a PyInstaller-frozen bundle.
"""
import os
import sys
import socket
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
for _env_candidate in [os.path.join(_exe_dir, '.env'), os.path.join(_bundle_dir, '.env')]:
    if os.path.exists(_env_candidate):
        with open(_env_candidate) as _f:
            for _line in _f:
                _line = _line.strip()
                if _line and not _line.startswith('#') and '=' in _line:
                    _k, _v = _line.split('=', 1)
                    os.environ.setdefault(_k.strip(), _v.strip().strip('"\''))
        break

# ── Config ───────────────────────────────────────────────────────────────────
PORT = int(os.environ.get('PORT', 5001))
HOST = '127.0.0.1'
URL  = f'http://{HOST}:{PORT}'


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


def _run_flask():
    """Import and run the Flask app (runs forever in its thread)."""
    from app import app as flask_app
    flask_app.run(host=HOST, port=PORT, debug=False, use_reloader=False)


if __name__ == '__main__':
    # Start Flask in a daemon thread so it dies when the process exits
    server_thread = threading.Thread(target=_run_flask, daemon=True)
    server_thread.start()

    # Wait for Flask to be ready, then open the browser
    if _wait_for_server(timeout=30):
        webbrowser.open(URL)
    else:
        # Server didn't start — on Mac/Windows show an error via the browser anyway
        webbrowser.open(URL)

    # Keep the main thread alive (the daemon thread will exit when this does)
    try:
        server_thread.join()
    except KeyboardInterrupt:
        pass
