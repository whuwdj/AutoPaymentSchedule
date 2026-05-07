# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller 配置：macOS 产出 智能排款.app；Windows 产出 dist\\智能排款\\ 目录。"""
import os
import sys

block_cipher = None

_datas = []
if os.path.isfile("app.ico"):
    _datas.append(("app.ico", "."))

a = Analysis(
    ["main.py"],
    pathex=[],
    binaries=[],
    datas=_datas,
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

_app_icon = "app.ico" if os.path.isfile("app.ico") else None

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="智能排款",
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
    icon=_app_icon,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name="智能排款",
)

if sys.platform == "darwin":
    app = BUNDLE(
        coll,
        name="智能排款.app",
        icon="app.icns" if os.path.isfile("app.icns") else None,
        bundle_identifier="com.autopayment.smartschedule",
        info_plist={"NSHighResolutionCapable": "True"},
    )
