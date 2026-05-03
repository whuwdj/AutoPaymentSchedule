# -*- coding: utf-8 -*-
"""工作区路径：程序目录（本包所在目录）下的 AutoPaymentScheduleFile。"""

from __future__ import annotations

import os


def _program_dir() -> str:
    return os.path.dirname(os.path.abspath(__file__))


def get_base_dir() -> str:
    return os.path.join(_program_dir(), "AutoPaymentScheduleFile")


def total_sheet_dir() -> str:
    return os.path.join(get_base_dir(), "TotalSheet")


def detail_template_dir() -> str:
    return os.path.join(get_base_dir(), "DetailTemplate")


def ensure_workspace() -> None:
    os.makedirs(total_sheet_dir(), exist_ok=True)
    os.makedirs(detail_template_dir(), exist_ok=True)
