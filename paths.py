# -*- coding: utf-8 -*-
"""工作区路径：程序目录（本包所在目录）下的 AutoPaymentScheduleFile。"""

from __future__ import annotations

import json
import os
import shutil


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


def _is_total_sheet_excel_name(name: str) -> bool:
    lower = name.lower()
    return lower.endswith(".xlsx") or lower.endswith(".xls")


def total_sheet_startup_snapshot_dir() -> str:
    """程序启动时 TotalSheet 快照目录（用于智能排款前恢复）。"""
    return os.path.join(get_base_dir(), "_snapshots", "total_sheet_startup")


def capture_total_sheet_startup_snapshot() -> None:
    """
    在程序打开时调用：将当前 TotalSheet 下所有 Excel 复制到快照目录，
    并写入清单。若启动时 TotalSheet 无 Excel，清单为空；智能排款前恢复时不会清空用户后续上传。
    """
    snap = total_sheet_startup_snapshot_dir()
    shutil.rmtree(snap, ignore_errors=True)
    os.makedirs(snap, exist_ok=True)
    d = total_sheet_dir()
    names = sorted(n for n in os.listdir(d) if _is_total_sheet_excel_name(n))
    for n in names:
        shutil.copy2(os.path.join(d, n), os.path.join(snap, n))
    meta = os.path.join(snap, ".startup_manifest.json")
    with open(meta, "w", encoding="utf-8") as f:
        json.dump({"files": names}, f, ensure_ascii=False)


def restore_total_sheet_from_startup_snapshot() -> None:
    """
    将 TotalSheet 恢复为程序启动时的 Excel 状态（快照中有文件时才：
    先删除当前目录下所有 Excel，再从快照拷回）。启动时无总表则不做任何操作。
    """
    snap = total_sheet_startup_snapshot_dir()
    meta = os.path.join(snap, ".startup_manifest.json")
    if not os.path.isfile(meta):
        return
    with open(meta, encoding="utf-8") as f:
        names: list[str] = json.load(f).get("files", [])
    if not names:
        return
    d = total_sheet_dir()
    os.makedirs(d, exist_ok=True)
    for n in os.listdir(d):
        if _is_total_sheet_excel_name(n):
            os.remove(os.path.join(d, n))
    for n in names:
        src = os.path.join(snap, n)
        if os.path.isfile(src):
            shutil.copy2(src, os.path.join(d, n))
