import os, sys

# Load .env file if present (local dev / API keys)
_env_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), '.env')
if os.path.exists(_env_path):
    with open(_env_path) as _f:
        for _line in _f:
            _line = _line.strip()
            if _line and not _line.startswith('#') and '=' in _line:
                _k, _v = _line.split('=', 1)
                os.environ.setdefault(_k.strip(), _v.strip().strip('"\''))

sys.path.insert(0, os.path.dirname(__file__))
from app import app

port = int(os.environ.get('PORT', 5001))
app.run(host='0.0.0.0', port=port, debug=False, use_reloader=False)

