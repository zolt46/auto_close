# -*- mode: python ; coding: utf-8 -*-

import importlib.util

_pathlib_spec = importlib.util.find_spec("pathlib")
if _pathlib_spec and _pathlib_spec.origin and "site-packages" in _pathlib_spec.origin:
    raise SystemExit(
        "PyInstaller와 호환되지 않는 pip backport 'pathlib'가 설치되어 있습니다. "
        "명령 프롬프트에서 'python -m pip uninstall pathlib'를 실행한 뒤 다시 시도하세요."
    )

from pathlib import Path

import PySide6

block_cipher = None
ROOT = Path(r"C:\Users\seewo\Desktop\closing_new_proj")
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