# -*- mode: python ; coding: utf-8 -*-
#
# PyInstaller spec for Startup Scout
#
# Build commands:
#   pip install pyinstaller
#   pyinstaller startup-scout.spec
#
# Output:
#   dist/Startup Scout.exe          (Windows)
#   dist/Startup Scout.app          (macOS — also zipped to Startup-Scout-Mac.zip by CI)

import sys
from PyInstaller.building.build_main import Analysis, PYZ, EXE, BUNDLE, COLLECT

block_cipher = None

a = Analysis(
    ['launcher.py'],
    pathex=[],
    binaries=[],
    datas=[
        # Bundle templates and static assets inside the executable
        ('templates', 'templates'),
        ('static',    'static'),
    ],
    hiddenimports=[
        # Flask + Werkzeug internals that PyInstaller sometimes misses
        'flask',
        'flask.templating',
        'werkzeug',
        'werkzeug.routing',
        'werkzeug.security',
        'jinja2',
        'jinja2.ext',
        'markupsafe',
        # Data
        'openpyxl',
        'openpyxl.styles',
        'openpyxl.utils',
        'pandas',
        'numpy',
        # AI
        'anthropic',
        'httpx',
        'httpcore',
        # Standard library extras
        'email.mime.text',
        'email.mime.multipart',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['tkinter', 'matplotlib', 'scipy', 'PIL'],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='Startup Scout',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,          # No terminal window — clean desktop feel
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='icon.icns' if sys.platform == 'darwin' else None,
)

# macOS: wrap into a .app bundle
if sys.platform == 'darwin':
    app = BUNDLE(
        exe,
        name='Startup Scout.app',
        icon='icon.icns',
        bundle_identifier='com.startup-scout.app',
        info_plist={
            'NSHighResolutionCapable': True,
            'CFBundleShortVersionString': '1.0.0',
            'CFBundleName': 'Startup Scout',
            'NSHumanReadableCopyright': 'EDF Energy',
        },
    )
