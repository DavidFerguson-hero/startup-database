"""User management for Startup Database."""
import json, os, uuid, secrets
from datetime import datetime, timedelta, timezone
from werkzeug.security import generate_password_hash, check_password_hash


def _path():
    base = os.environ.get('DATA_DIR', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, 'users.json')


def _load():
    p = _path()
    if not os.path.exists(p):
        return {}
    with open(p) as f:
        return json.load(f)


def _save(data):
    with open(_path(), 'w') as f:
        json.dump(data, f, indent=2)


def _now():
    return datetime.now(timezone.utc).isoformat()


def _utc():
    return datetime.now(timezone.utc)


def _parse_dt(s):
    dt = datetime.fromisoformat(s)
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=timezone.utc)
    return dt


# ── Queries ───────────────────────────────────────────────────────────────────

def count():
    return len(_load())


def get_all():
    return list(_load().values())


def get_by_id(uid):
    return _load().get(uid)


def get_by_email(email):
    email = email.lower().strip()
    for u in _load().values():
        if u['email'] == email:
            return u
    return None


def get_by_token(field, token):
    """Return the user whose ``field`` matches ``token``, or None."""
    for u in _load().values():
        if u.get(field) == token:
            return u
    return None


# ── Mutations ─────────────────────────────────────────────────────────────────

def create(email, name, password, role='user'):
    """Create a new (unverified) user. Returns (user_dict, verification_token)."""
    data = _load()
    uid = str(uuid.uuid4())
    token = secrets.token_urlsafe(32)
    data[uid] = {
        'id': uid,
        'email': email.lower().strip(),
        'name': name.strip(),
        'password_hash': generate_password_hash(password),
        'role': role,
        'active': True,
        'verified': False,
        'verification_token': token,
        'reset_token': None,
        'reset_token_expiry': None,
        'created_at': _now(),
        'last_login': None,
    }
    _save(data)
    return data[uid], token


def verify_email(token):
    """Mark user as verified. Returns user dict or None if token invalid."""
    data = _load()
    for u in data.values():
        if u.get('verification_token') == token:
            u['verified'] = True
            u['verification_token'] = None
            _save(data)
            return u
    return None


def authenticate(email, password):
    """
    Verify credentials. Returns (user, None) on success or (None, reason).
    reason is one of: 'invalid', 'unverified', 'disabled'
    """
    u = get_by_email(email)
    if not u:
        return None, 'invalid'
    if not u.get('verified'):
        return None, 'unverified'
    if not u.get('active'):
        return None, 'disabled'
    if check_password_hash(u['password_hash'], password):
        return u, None
    return None, 'invalid'


def make_reset_token(email):
    """Create a 1-hour password reset token. Returns (user, token) or (None, None)."""
    data = _load()
    email = email.lower().strip()
    for u in data.values():
        if u['email'] == email:
            token = secrets.token_urlsafe(32)
            u['reset_token'] = token
            u['reset_token_expiry'] = (_utc() + timedelta(hours=1)).isoformat()
            _save(data)
            return u, token
    return None, None


def reset_password(token, new_password):
    """Consume a reset token and update the password. Returns user or None."""
    data = _load()
    for u in data.values():
        if u.get('reset_token') == token:
            exp = u.get('reset_token_expiry')
            if exp and _parse_dt(exp) > _utc():
                u['password_hash'] = generate_password_hash(new_password)
                u['reset_token'] = None
                u['reset_token_expiry'] = None
                _save(data)
                return u
    return None


def update(uid, **kwargs):
    data = _load()
    if uid in data:
        data[uid].update(kwargs)
        _save(data)
        return data[uid]
    return None


def delete(uid):
    data = _load()
    if uid not in data:
        return False
    del data[uid]
    _save(data)
    return True


def record_login(uid):
    update(uid, last_login=_now())
