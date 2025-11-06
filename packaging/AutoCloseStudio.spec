# -*- mode: python ; coding: utf-8 -*-

from pathlib import Path

import PySide6

block_cipher = None

ROOT = Path(__file__).resolve().parents[1]
ASSET_DIR = ROOT / "assets"
PYSIDE_PLUGINS = Path(PySide6.__file__).resolve().parent / "plugins"

asset_datas = []
for name in ("app_icon.ico", "topbar_logo.png"):
    src = ASSET_DIR / name
    if src.exists():
        asset_datas.append((str(src), "assets"))

qt_plugin_datas = []
for plugin in ("platforms", "styles"):
    src = PYSIDE_PLUGINS / plugin
    if src.exists():
        qt_plugin_datas.append((str(src), f"PySide6/plugins/{plugin}"))

hiddenimports = [
    "PySide6.QtMultimedia",
    "PySide6.QtMultimediaWidgets",
    "PySide6.QtNetwork",
]

a = Analysis(
    ['desktop_scheduler_qt.py'],
    pathex=[str(ROOT)],
    binaries=[],
    datas=asset_datas + qt_plugin_datas,
    hiddenimports=hiddenimports,
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
    name='AutoCloseStudio',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=str(ASSET_DIR / 'app_icon.ico') if (ASSET_DIR / 'app_icon.ico').exists() else None,
)
coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='AutoCloseStudio',
)