# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller spec file for 智能排款.

macOS: 打包为 .app，使用 AutoPay.jpeg 作为 .app 图标。
Windows: 打包为目录 dist\\智能排款\\，使用 app.ico 作为 EXE 图标。
"""

import os
import sys

block_cipher = None
a = Analysis(
    ["main.py"],
    pathex=[],
    binaries=[],
    datas=[
        ("AutoPay.jpeg", "."),
    ],
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

if sys.platform == "darwin":
    # macOS: 生成 .app，优先使用 app.icns，回退到 AutoPay.jpeg
    icon_file = "app.icns" if os.path.exists("app.icns") else "AutoPay.jpeg"
    exe = EXE(
        pyz,
        a.scripts,
        a.binaries,
        a.zipfiles,
        a.datas,
        [],
        name="智能排款",
        debug=False,
        bootloader_ignore_signals=False,
        strip=False,
        upx=True,
        upx_exclude=[],
        runtime_tmpdir=None,
        console=False,
        icon=icon_file,
        info_plist={
            "NSHighResolutionCapable": "True",
        },
    )
    app = BUNDLE(
        exe,
        name="智能排款.app",
        icon=icon_file,
        buffer_size=4096,
        entropy_bootloader=False,
        entitlements_file=None,
    )
else:
    # Windows/Linux: 生成目录模式，带独立图标
    icon_file = "app.ico" if os.path.exists("app.ico") else None
    exe = EXE(
        pyz,
        a.scripts,
        a.binaries,
        a.zipfiles,
        a.datas,
        [],
        name="智能排款",
        debug=False,
        bootloader_ignore_signals=False,
        strip=False,
        upx=True,
        upx_exclude=[],
        runtime_tmpdir=None,
        console=False,
        icon=icon_file,
    )