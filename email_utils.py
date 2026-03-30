"""Email utilities for Startup Database."""
import os, smtplib, logging
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

logger = logging.getLogger(__name__)


def _cfg():
    return {
        'host':     os.environ.get('SMTP_HOST', ''),
        'port':     int(os.environ.get('SMTP_PORT', 587)),
        'user':     os.environ.get('SMTP_USER', ''),
        'password': os.environ.get('SMTP_PASS', ''),
        'from':     os.environ.get('SMTP_FROM') or os.environ.get('SMTP_USER', 'noreply@startupdb'),
    }


def send(to, subject, html, text=None):
    cfg = _cfg()
    if not cfg['host']:
        logger.warning('SMTP_HOST not set — email NOT sent to %s | %s', to, subject)
        return False
    msg = MIMEMultipart('alternative')
    msg['Subject'] = subject
    msg['From'] = cfg['from']
    msg['To'] = to
    if text:
        msg.attach(MIMEText(text, 'plain'))
    msg.attach(MIMEText(html, 'html'))
    try:
        with smtplib.SMTP(cfg['host'], cfg['port'], timeout=10) as s:
            s.ehlo()
            if cfg['port'] != 25:
                s.starttls()
            if cfg['user']:
                s.login(cfg['user'], cfg['password'])
            s.sendmail(cfg['from'], [to], msg.as_string())
        logger.info('Email sent to %s: %s', to, subject)
        return True
    except Exception as e:
        logger.error('Failed to send email to %s: %s', to, e)
        return False


_CSS = """
  body{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;
       background:#0f172a;color:#f1f5f9;margin:0;padding:40px 20px}
  .card{background:#1e293b;border:1px solid #334155;border-radius:12px;
        max-width:480px;margin:0 auto;padding:36px}
  h2{color:#38bdf8;margin:0 0 16px;font-size:1.2rem}
  p{color:#94a3b8;line-height:1.6;margin:0 0 20px}
  .btn{display:inline-block;background:#38bdf8;color:#0f172a;padding:12px 24px;
       border-radius:8px;text-decoration:none;font-weight:700;font-size:.9rem}
  .url{word-break:break-all;color:#38bdf8;font-size:.82rem}
  .note{font-size:.78rem;color:#64748b}
"""


def _page(body):
    return f'<!DOCTYPE html><html><head><style>{_CSS}</style></head><body><div class="card">{body}</div></body></html>'


def send_verification(to, name, verify_url):
    html = _page(f"""
      <h2>Verify your email</h2>
      <p>Hi {name},<br>Click below to verify your email address and activate your Startup Database account.</p>
      <p><a href="{verify_url}" class="btn">Verify Email →</a></p>
      <p>Or copy this link:<br><span class="url">{verify_url}</span></p>
      <p class="note">This link expires in 24 hours. Didn't create an account? Ignore this email.</p>
    """)
    text = f"Hi {name},\n\nVerify your email address:\n{verify_url}\n\nExpires in 24 hours."
    return send(to, 'Verify your Startup Database account', html, text)


def send_reset(to, name, reset_url):
    html = _page(f"""
      <h2>Reset your password</h2>
      <p>Hi {name},<br>Click below to reset your Startup Database password. This link expires in 1 hour.</p>
      <p><a href="{reset_url}" class="btn">Reset Password →</a></p>
      <p>Or copy this link:<br><span class="url">{reset_url}</span></p>
      <p class="note">Didn't request a reset? You can safely ignore this email.</p>
    """)
    text = f"Hi {name},\n\nReset your password:\n{reset_url}\n\nExpires in 1 hour."
    return send(to, 'Reset your Startup Database password', html, text)
