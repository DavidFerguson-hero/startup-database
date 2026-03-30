#!/usr/bin/env python3
"""Run this script to change the login password for the Startup Database app.
Usage: python3 set_password.py"""
import os, getpass
from werkzeug.security import generate_password_hash

BASE = os.path.dirname(os.path.abspath(__file__))
PW_FILE = os.path.join(BASE, '.password_hash')

pw  = getpass.getpass('New password: ')
pw2 = getpass.getpass('Confirm password: ')

if pw != pw2:
    print('Passwords do not match.')
else:
    with open(PW_FILE, 'w') as f:
        f.write(generate_password_hash(pw, method='pbkdf2:sha256', salt_length=16))
    print('Password updated. Restart the app for it to take effect.')
