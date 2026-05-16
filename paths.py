# -*- coding: utf-8 -*-
"""工作区路径：程序目录（本包所在目录）下的 AutoPaymentScheduleFile。"""

from __future__ import annotations

import os
import shutil
import sys


def get_program_dir() -> str:
    """源码运行：本文件所在目录；PyInstaller 打包：可执行文件所在目录。

    macOS .app 内可执行文件位于 *.app/Contents/MacOS/，工作区放在 .app 同级目录，
    故解析为「包含 .app 的文件夹」。
    """
    if not getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(__file__))
    exe_dir = os.path.dirname(os.path.abspath(sys.executable))
    if sys.platform == "darwin" and exe_dir.endswith(f"{os.sep}Contents{os.sep}MacOS"):
        return os.path.abspath(os.path.join(exe_dir, "..", "..", ".."))
    return exe_dir


def _program_dir() -> str:
    return get_program_dir()


def get_base_dir() -> str:
    return os.path.join(_program_dir(), "AutoPaymentScheduleFile")


def total_sheet_dir() -> str:
    return os.path.join(get_base_dir(), "TotalSheet")


def detail_template_dir() -> str:
    return os.path.join(get_base_dir(), "DetailTemplate")


def ensure_workspace() -> None:
    os.makedirs(total_sheet_dir(), exist_ok=True)
    os.makedirs(detail_template_dir(), exist_ok=True)


def clear_workspace_upload_dirs() -> None:
    """清空 TotalSheet、DetailTemplate 目录内的文件与子目录（保留目录本身）。"""
    for d in (total_sheet_dir(), detail_template_dir()):
        if not os.path.isdir(d):
            continue
        for name in os.listdir(d):
            p = os.path.join(d, name)
            try:
                if os.path.isfile(p) or os.path.islink(p):
                    os.unlink(p)
                elif os.path.isdir(p):
                    shutil.rmtree(p)
            except OSError:
                pass
