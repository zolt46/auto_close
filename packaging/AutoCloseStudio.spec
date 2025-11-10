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

ROOT = Path(r"C:\Users\seewo\Desktop\closing_new_proj\auto_close")
SRC_DIR = ROOT / "packaging"
ASSET_DIR = SRC_DIR / "assets"
PYSIDE_PLUGINS = Path(PySide6.__file__).resolve().parent / "plugins"

asset_datas = []
for name in ("app_icon.ico", "topbar_logo.png"):
    src = ASSET_DIR / name
    if src.exists():
        asset_datas.append((str(src), "assets"))

qt_plugin_datas = []
for plugin in ("platforms", "styles", "audio", "multimedia"):
    src = PYSIDE_PLUGINS / plugin
    if src.exists():
        qt_plugin_datas.append((str(src), f"PySide6/plugins/{plugin}"))

hiddenimports = ["PySide6.QtMultimedia", "PySide6.QtMultimediaWidgets", "PySide6.QtNetwork"]

a = Analysis(
    [str(SRC_DIR / 'desktop_scheduler_qt.py')],
    pathex=[str(ROOT), str(SRC_DIR)],
    binaries=[],
    datas=[(str(ASSET_DIR), 'assets')] + asset_datas + qt_plugin_datas,
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

# ★ onefile: EXE에 모든 리소스 전달, COLLECT 사용하지 않음
exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='AutoCloseStudio',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,                      # 반드시 False
    console=False,
    icon=str(ASSET_DIR / 'app_icon.ico'),
    runtime_tmpdir='.',             # EXE 옆으로 풀기 (_internal 폴더 생성)
)
