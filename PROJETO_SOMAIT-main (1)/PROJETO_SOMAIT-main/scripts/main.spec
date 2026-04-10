# -*- mode: python ; coding: utf-8 -*-

from pathlib import Path


PROJECT_ROOT = Path(SPECPATH).resolve().parent.parent
ICON_PATH = PROJECT_ROOT / 'assets' / 'app.ico'

datas = [
    (str(PROJECT_ROOT / 'app' / 'templates'), 'app/templates'),
    (str(PROJECT_ROOT / 'app' / 'static'), 'app/static'),
]

optional_datas = [
    (PROJECT_ROOT / 'config.json', '.'),
]

for source_path, target_dir in optional_datas:
    if source_path.exists():
        datas.append((str(source_path), target_dir))


a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=['pythoncom', 'pywintypes', 'win32timezone', 'win32com', 'win32com.client'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='SOMALABSDesktop',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=str(ICON_PATH) if ICON_PATH.exists() else None,
)
