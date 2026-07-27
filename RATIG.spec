# -*- mode: python ; coding: utf-8 -*-
# RATIG EXE Build Spec — Production

import os

block_cipher = None

a = Analysis(
    ['RATIG.py'],
    pathex=[],
    binaries=[],
    datas=[('assets/ratig.ico', 'assets')],
    hiddenimports=[
        # pyserial — sering hilang di onefile build
        'serial',
        'serial.tools',
        'serial.tools.list_ports',
        'serial.tools.list_ports_windows',
        'serial.tools.list_ports_common',

        # matplotlib backend Tkinter
        'matplotlib',
        'matplotlib.backends.backend_tkagg',
        'matplotlib.backends._backend_tk',
        'matplotlib.figure',
        'matplotlib.pyplot',

        # pandas / numpy core
        'pandas',
        'pandas._libs.tslibs.np_datetime',
        'pandas._libs.tslibs.nattype',
        'pandas._libs.tslibs.timedeltas',
        'pandas._libs.tslibs.period',
        'pandas._libs.hashtable',
        'pandas._libs.lib',
        'numpy',

        # openpyxl
        'openpyxl',
        'openpyxl.cell',
        'openpyxl.styles',
        'openpyxl.utils',
        'openpyxl.workbook',

        # fpdf2
        'fpdf',

        # sqlite3 sudah built-in Python tapi declare supaya aman
        'sqlite3',

        # zoneinfo & datetime (dibutuhkan oleh pandas)
        'zoneinfo',

        # tkinter extras
        'tkinter',
        'tkinter.ttk',
        'tkinter.messagebox',
        'tkinter.filedialog',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        # Hapus dependency berat yang tidak terpakai
        'joblib',
        'sklearn',
        'scipy',
        'IPython',
        'notebook',
        'PyQt5',
        'PyQt6',
        'wx',
        'botocore',
        'boto3',
        's3transfer',
        'fsspec',
    ],
    noarchive=False,
    optimize=1,
)

# Filter out unnecessary Tcl/Tk tzdata, msgs, and encoding spam files
a.datas = [x for x in a.datas if not any(k in x[0].lower() or k in x[1].lower() for k in ['tzdata', 'msgs', 'encoding'])]

pyz = PYZ(a.pure, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='RATIG',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=['vcruntime140.dll', 'python3.dll'],
    runtime_tmpdir=None,
    console=False,                   # Tidak ada CMD window
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='assets/ratig.ico' if os.path.exists('assets/ratig.ico') else None,
    version_file=None,
)
