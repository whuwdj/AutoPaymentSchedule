# -*- coding: utf-8 -*-
"""固定工作区路径：Windows 为 D:\\AutoPaymentScheduleFile；其他系统为项目内 _data 便于开发。"""

from __future__ import annotations

import os
import sys


def get_base_dir() -> str:
    if sys.platform == "win32":
        return r"D:\AutoPaymentScheduleFile"
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), "_data")


def total_sheet_dir() -> str:
    return os.path.join(get_base_dir(), "TotalSheet")


def detail_template_dir() -> str:
    return os.path.join(get_base_dir(), "DetailTemplate")


def ensure_workspace() -> None:
    os.makedirs(total_sheet_dir(), exist_ok=True)
    os.makedirs(detail_template_dir(), exist_ok=True)
