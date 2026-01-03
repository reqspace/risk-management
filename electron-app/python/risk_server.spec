# -*- mode: python ; coding: utf-8 -*-
# PyInstaller spec file for Risk Management Python backend

import sys
from pathlib import Path

# Get the parent directory (Risk-Management root)
ROOT_DIR = Path(SPECPATH).parent.parent

block_cipher = None

# Collect all Python files needed
a = Analysis(
    [str(ROOT_DIR / 'server.py')],
    pathex=[str(ROOT_DIR)],
    binaries=[],
    datas=[
        # Include templates and static files if any
        (str(ROOT_DIR / '.env.example'), '.'),
    ],
    hiddenimports=[
        'flask',
        'flask_cors',
        'anthropic',
        'openpyxl',
        'openpyxl.styles',
        'openpyxl.utils',
        'dotenv',
        'watchdog',
        'watchdog.observers',
        'watchdog.events',
        'apscheduler',
        'apscheduler.schedulers.background',
        'apscheduler.triggers.cron',
        'requests',
        'docx',
        'docx.shared',
        'docx.enum.text',
        'docx.enum.table',
        'email',
        'email.mime.text',
        'email.mime.multipart',
        'email.mime.base',
        'email.mime.application',
        'smtplib',
        'imaplib',
        'json',
        'datetime',
        'pathlib',
        'threading',
        'queue',
        'logging',
        'werkzeug',
        'jinja2',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'tkinter',
        'matplotlib',
        'numpy',
        'pandas',
        'scipy',
        'PIL',
        'cv2',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

# Also include the other Python modules
for pyfile in ['process.py', 'daily_digest.py', 'monthly_report.py',
               'send_monthly_report.py', 'email_reader.py']:
    filepath = ROOT_DIR / pyfile
    if filepath.exists():
        a.datas.append((pyfile, str(filepath), 'DATA'))

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='risk_server',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # No console window
    disable_windowed_traceback=False,
    argv_emulation=True,  # For macOS
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

# For macOS, create an app bundle
if sys.platform == 'darwin':
    app = BUNDLE(
        exe,
        name='risk_server.app',
        icon=None,
        bundle_identifier='com.reqspace.riskmanagement.server',
    )
