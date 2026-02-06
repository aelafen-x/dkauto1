# -*- mode: python ; coding: utf-8 -*-
from pathlib import Path
from PyInstaller.compat import is_darwin

block_cipher = None

project_root = Path(globals().get("specpath", ".")).resolve()

datas = [
    (str(project_root / "points.json"), "Config files"),
    (str(project_root / "name_aliases.json"), "Config files"),
    (str(project_root / "boss_aliases.json"), "Config files"),
    (str(project_root / "prios.json"), "Config files"),
]

a = Analysis(
    ["run_app.py"],
    pathex=[str(project_root)],
    binaries=[],
    datas=datas,
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="DKP Automator",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
)

if is_darwin:
    app = BUNDLE(
        exe,
        name="DKP Automator.app",
    )
    coll = COLLECT(
        app,
        a.binaries,
        a.zipfiles,
        a.datas,
        strip=False,
        upx=True,
        name="DKP Automator",
    )
else:
    coll = COLLECT(
        exe,
        a.binaries,
        a.zipfiles,
        a.datas,
        strip=False,
        upx=True,
        name="DKP Automator",
    )
