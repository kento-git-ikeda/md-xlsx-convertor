# -*- mode: python ; coding: utf-8 -*-
from pathlib import Path
from PyInstaller.utils.hooks import collect_all

root = Path(SPECPATH)
mael_datas, mael_binaries, mael_hiddenimports = collect_all("mael")

analysis = Analysis(
    [str(root / "app.py")],
    pathex=[str(root / "src")],
    binaries=mael_binaries,
    datas=mael_datas,
    hiddenimports=mael_hiddenimports + [
        "mael.excel_builder",
        "mael.composer",
        "mael.column_config",
        "yaml",
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=1,
)

pyz = PYZ(analysis.pure)
exe = EXE(
    pyz,
    analysis.scripts,
    analysis.binaries,
    analysis.datas,
    [],
    name="md-xlsx-convertor",
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
)
