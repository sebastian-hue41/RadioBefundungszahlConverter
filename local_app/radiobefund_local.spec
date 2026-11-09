# -*- mode: python ; coding: utf-8 -*-
#
# PyInstaller spec of the local app: one single-file program per operating system.
#
#   pip install -r local_app/requirements-build.txt
#   pyinstaller --noconfirm --distpath <dist> --workpath <work> local_app/radiobefund_local.spec
#
# Build outside the Synology folder (--distpath/--workpath) and on the operating system the
# program is meant for: PyInstaller cannot build for another one.
#
# Inside the program: the code of core/ and local_app/, openpyxl, the Python standard library,
# the page (local_app/static/index.html) and reference_amount.xlsx. Nothing else: no map, no
# statistics, no web worker (web/), no iii framework, no tests.
import os

ROOT = os.path.dirname(SPECPATH)            # the project root; SPECPATH is the folder of this file

a = Analysis(
    [os.path.join(ROOT, "local_app", "main.py")],
    pathex=[ROOT],
    binaries=[],
    datas=[
        (os.path.join(ROOT, "local_app", "static", "index.html"), os.path.join("local_app", "static")),
        (os.path.join(ROOT, "reference_amount.xlsx"), "."),
    ],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=["web", "iii", "tests", "pytest", "tkinter", "_tkinter"],
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="befundungszahlen",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,                              # compressed programs trigger more antivirus alarms
    runtime_tmpdir=None,
    console=True,                           # the window shows that the app runs; closing it stops the server
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
