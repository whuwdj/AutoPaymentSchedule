# -*- coding: utf-8 -*-
"""
智能排款：读取总表 Sheet1（程序目录 AutoPaymentScheduleFile/TotalSheet 下 Excel），
仅处理「付款优先级 = 0」的数据行；账期汇总、三条件匹配排款列、写回金额并扣减第 5 行排款余额。
"""

from __future__ import annotations

import re
from copy import copy
from dataclasses import dataclass
from datetime import date, datetime
from typing import Dict, List, Optional, Tuple

from openpyxl import load_workbook
from openpyxl.cell.cell import Cell, MergedCell
from openpyxl.utils.datetime import from_excel
from openpyxl.worksheet.worksheet import Worksheet

from excel_service import (
    _cell_text_normalized,
    _coerce_numeric_for_sum,
    _is_xls,
    _norm_header_label,
)


@dataclass
class SmartScheduleResult:
    path: str
    rows_priority0: int
    rows_scanned: int
    rows_written: int
    rows_skipped: int


def _norm_h(text: str) -> str:
    return _cell_text_normalized(text).replace("：", ":").strip()


_MONTH_HEADER_RE = re.compile(r"^(\d+)\s*个月")


def _month_bucket_n(cell_value: object) -> Optional[int]:
    """从表头如「1个月（26.3月）」提取 n；「1年以上」返回 13。"""
    t = _cell_text_normalized(cell_value)
    if not t:
        return None
    m = _MONTH_HEADER_RE.match(t)
    if m:
        return int(m.group(1))
    if "年以上" in t:
        return 13
    return None


def _find_sheet1_header_row(ws: Worksheet) -> Optional[int]:
    """定位含「主体」「协议账期」的数据表头行（通常为第 2 行）。"""
    max_r = min(ws.max_row or 0, 30)
    max_c = min(ws.max_column or 0, 120)
    best_r: Optional[int] = None
    best_score = 0
    for r in range(1, max_r + 1):
        hits = 0
        has_subject = False
        has_agreement = False
        for c in range(1, max_c + 1):
            h = _norm_h(ws.cell(row=r, column=c).value)
            if h == "主体":
                has_subject = True
                hits += 1
            elif h == "协议账期":
                has_agreement = True
                hits += 1
            elif h == "截止":
                hits += 1
        if has_subject and has_agreement and hits >= best_score:
            best_score = hits
            best_r = r
    return best_r


def _is_invoice_balance_header(text: str) -> bool:
    t = _norm_h(text)
    if t == "当月已开票余额":
        return True
    if "已开票" in t and "余额" in t:
        return True
    if re.match(r"^\d{6,8}.*余额$", t):
        return True
    return False


def _find_col_by_exact(ws: Worksheet, header_row: int, name: str) -> Optional[int]:
    for c in range(1, (ws.max_column or 0) + 1):
        if _norm_h(ws.cell(row=header_row, column=c).value) == name:
            return c
    return None


def _find_col_header_contains(ws: Worksheet, header_row: int, needle: str) -> Optional[int]:
    """表头文本包含 needle 的首列（用于「付款优先级（…）」等长表头）。"""
    for c in range(1, (ws.max_column or 0) + 1):
        h = _cell_text_normalized(ws.cell(row=header_row, column=c).value)
        if needle in h:
            return c
    return None


def _find_payment_plan_col(ws: Worksheet, header_row: int) -> Optional[int]:
    """「4月付款计划」等：表头含「付款计划」且不含「确认」（排除「付款计划确认」）。"""
    for c in range(1, (ws.max_column or 0) + 1):
        h = _cell_text_normalized(ws.cell(row=header_row, column=c).value)
        if "付款计划" in h and "确认" not in h:
            return c
    return None


def _find_invoice_balance_col(ws: Worksheet, header_row: int) -> Optional[int]:
    for c in range(1, (ws.max_column or 0) + 1):
        raw = ws.cell(row=header_row, column=c).value
        if _is_invoice_balance_header(_cell_text_normalized(raw)):
            return c
    return None


def _find_cutoff_col(ws: Worksheet, header_row: int) -> Optional[int]:
    return _find_col_by_exact(ws, header_row, "截止")


def _find_first_payment_col(ws: Worksheet, header_row: int) -> Optional[int]:
    for c in range(1, (ws.max_column or 0) + 1):
        h = _norm_h(ws.cell(row=header_row, column=c).value)
        if h == "第1笔" or h.startswith("第1笔"):
            return c
    return None


def _find_total_scheduled_col(ws: Worksheet, header_row: int) -> Optional[int]:
    for c in range(1, (ws.max_column or 0) + 1):
        h = _norm_h(ws.cell(row=header_row, column=c).value)
        if h == "合计已排款":
            return c
    return None


def _protocol_period_months(value: object) -> Optional[float]:
    """协议账期 ÷ 30 → nPeriodCount；无法解析则 None。"""
    if value is None:
        return None
    s = str(value).strip()
    if not s or s in ("-", "—", "－"):
        return None
    n = _coerce_numeric_for_sum(value)
    if n is None or n <= 0:
        return None
    return n / 30.0


def _split_tokens(s: str) -> List[str]:
    parts: List[str] = []
    for chunk in re.split(r"[/、，,|；;]+", s):
        t = chunk.strip()
        if t:
            parts.append(t)
    return parts


def _subject_contained_in_pay_region(row_subject: str, pay_cell: object) -> bool:
    """条件 2a：当前行主体包含于第七行「可支付主体」单元格全文即可。"""
    sub = row_subject.strip()
    if not sub:
        return False
    cell = _cell_text_normalized(pay_cell)
    if not cell:
        return False
    return sub in cell


def _payment_method_any_token_match(row_method: str, col_method_cell: object) -> bool:
    """条件 2b：当前行收款方式拆分后，与第八行该列收款方式拆分结果任一相同或包含关系即通过。"""
    rm = row_method.strip()
    if not rm:
        return False
    cm = _cell_text_normalized(col_method_cell)
    if not cm:
        return False
    row_toks = _split_tokens(rm)
    if not row_toks:
        row_toks = [rm]
    col_toks = _split_tokens(cm)
    if not col_toks:
        col_toks = [cm]
    for r in row_toks:
        if not r:
            continue
        if r in cm:
            return True
        for c in col_toks:
            if not c:
                continue
            if r == c or r in c or c in r:
                return True
    return False


def _read_schedule_balance(ws: Worksheet, col: int) -> float:
    """
    第 5 行「排款余额」；若为空则用第 4 行「排款总额」作为初始可排额度
    （兼容尚未维护排款余额公式的模版）。
    """
    v5 = ws.cell(row=5, column=col).value
    n5 = _coerce_numeric_for_sum(v5)
    if n5 is not None:
        return float(n5)
    v4 = ws.cell(row=4, column=col).value
    n4 = _coerce_numeric_for_sum(v4)
    if n4 is not None:
        return float(n4)
    return 0.0


def _parse_cell_datetime(value: object) -> Optional[datetime]:
    """将单元格解析为 datetime（用于放款日、付款截止日比较）。"""
    if value is None:
        return None
    if isinstance(value, datetime):
        return value
    if isinstance(value, date) and not isinstance(value, datetime):
        return datetime.combine(value, datetime.min.time())
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        try:
            return from_excel(float(value))
        except (TypeError, ValueError, OSError):
            return None
    s = str(value).strip()
    if not s or s in ("-", "—", "－"):
        return None
    if len(s) >= 19:
        try:
            return datetime.strptime(s[:19], "%Y-%m-%d %H:%M:%S")
        except ValueError:
            pass
    for fmt in ("%Y-%m-%d", "%Y/%m/%d"):
        try:
            return datetime.strptime(s[:10], fmt)
        except ValueError:
            continue
    return None


def _loan_deadline_ok(ws: Worksheet, pay_col: int, deadline_cell: object) -> bool:
    """条件 3：第 6 行放款日期 ≤ 当前行付款截止日期（同日可）。"""
    loan_raw = ws.cell(row=6, column=pay_col).value
    loan_dt = _parse_cell_datetime(loan_raw)
    due_dt = _parse_cell_datetime(deadline_cell)
    if loan_dt is None or due_dt is None:
        return False
    return loan_dt <= due_dt


def _is_priority_zero(value: object) -> bool:
    """付款优先级 = 0（付款优先级 = 1、2 不进入本逻辑）。"""
    if value is None:
        return False
    if isinstance(value, bool):
        return False
    if isinstance(value, (int, float)):
        try:
            return float(value) == 0.0
        except (TypeError, ValueError):
            return False
    s = str(value).strip()
    if not s:
        return False
    try:
        return float(s) == 0.0
    except ValueError:
        return False


def _find_row4_payout_gross_col(ws: Worksheet) -> Optional[int]:
    for c in range(1, (ws.max_column or 0) + 1):
        t = _cell_text_normalized(ws.cell(row=4, column=c).value)
        if _norm_header_label(t) == "排款总额":
            return c
    return None


def _find_row1_schedule_plan_column_span(ws: Worksheet) -> Optional[Tuple[int, int]]:
    max_c = ws.max_column or 0
    if max_c <= 0:
        return None
    for mr in ws.merged_cells.ranges:
        if mr.min_row > 1 or mr.max_row < 1:
            continue
        tl = _cell_text_normalized(ws.cell(row=1, column=mr.min_col).value)
        if "排款计划" in tl:
            return (mr.min_col, mr.max_col)
    for c in range(1, max_c + 1):
        tl = _cell_text_normalized(ws.cell(row=1, column=c).value)
        if "排款计划" not in tl:
            continue
        for mr in ws.merged_cells.ranges:
            if mr.min_row <= 1 <= mr.max_row and mr.min_col <= c <= mr.max_col:
                return (mr.min_col, mr.max_col)
        return (c, c)
    return None


def _clear_cell_if_writable(ws: Worksheet, row: int, col: int) -> None:
    cell = ws.cell(row=row, column=col)
    if isinstance(cell, MergedCell):
        return
    cell.value = None


def _copy_cell_as_is(src: Cell, dst: Cell) -> None:
    if isinstance(dst, MergedCell):
        return
    dst.value = src.value
    if src.has_style:
        dst.number_format = src.number_format
        dst.font = copy(src.font)
        dst.fill = copy(src.fill)
        dst.border = copy(src.border)
        dst.alignment = copy(src.alignment)
        dst.protection = copy(src.protection)


def _prepare_sheet_three_steps(
    ws: Worksheet,
    header_row: int,
    c_payment_plan: int,
    c_total_sched: int,
) -> None:
    """智能排款前：清空付款计划列、清空第 1 行「排款计划」列区第 9 行起、第 4 行横向区复制到第 5 行。"""
    max_r = ws.max_row or 0
    if max_r <= 0:
        return

    for r in range(header_row + 1, max_r + 1):
        _clear_cell_if_writable(ws, r, c_payment_plan)

    span = _find_row1_schedule_plan_column_span(ws)
    if span is not None:
        c_lo, c_hi = span
        for r in range(9, max_r + 1):
            for c in range(c_lo, c_hi + 1):
                _clear_cell_if_writable(ws, r, c)

    c_gross = _find_row4_payout_gross_col(ws)
    if c_gross is None:
        return
    start_c = c_gross + 1
    end_c = c_total_sched
    if start_c > end_c:
        return
    for c in range(start_c, end_c + 1):
        src = ws.cell(row=4, column=c)
        if isinstance(src, MergedCell):
            continue
        dst = ws.cell(row=5, column=c)
        _copy_cell_as_is(src, dst)


def run_smart_schedule_on_total_sheet(path: str) -> SmartScheduleResult:
    """
    在排款计划总表 Sheet1 上执行智能排款：仅处理「付款优先级 = 0」的行。
    仅支持 .xlsx。
    """
    if _is_xls(path):
        raise ValueError("智能排款仅支持 .xlsx 格式的排款计划总表。")

    wb = load_workbook(path, read_only=False, data_only=False)
    try:
        ws = wb.worksheets[0]
        header_row = _find_sheet1_header_row(ws)
        if header_row is None:
            raise ValueError("未在 Sheet1 中找到含「主体」「协议账期」的表头行。")

        c_subject = _find_col_by_exact(ws, header_row, "主体")
        c_agreement = _find_col_by_exact(ws, header_row, "协议账期")
        c_method = _find_col_by_exact(ws, header_row, "收款方式")
        c_invoice = _find_invoice_balance_col(ws, header_row)
        c_cutoff = _find_cutoff_col(ws, header_row)
        c_pay0 = _find_first_payment_col(ws, header_row)
        c_total_sched = _find_total_scheduled_col(ws, header_row)
        c_priority = _find_col_header_contains(ws, header_row, "付款优先级")
        c_due = _find_col_by_exact(ws, header_row, "付款截止日期")
        c_payment_plan = _find_payment_plan_col(ws, header_row)

        missing = []
        if c_subject is None:
            missing.append("主体")
        if c_agreement is None:
            missing.append("协议账期")
        if c_method is None:
            missing.append("收款方式")
        if c_invoice is None:
            missing.append("当月已开票余额（或同义「*余额」列）")
        if c_cutoff is None:
            missing.append("截止")
        if c_pay0 is None or c_total_sched is None:
            missing.append("第1笔 / 合计已排款")
        if c_priority is None:
            missing.append("付款优先级")
        if c_due is None:
            missing.append("付款截止日期")
        if c_payment_plan is None:
            missing.append("付款计划（表头含「付款计划」且非「付款计划确认」）")
        if missing:
            raise ValueError("表头缺少必要列：" + "、".join(missing))

        assert c_payment_plan is not None and c_total_sched is not None
        _prepare_sheet_three_steps(ws, header_row, c_payment_plan, c_total_sched)

        assert c_invoice is not None and c_cutoff is not None
        if c_invoice >= c_cutoff:
            raise ValueError("「当月已开票余额」列须位于「截止」列左侧。")

        payment_cols: List[int] = list(range(c_pay0, c_total_sched + 1))

        max_r = ws.max_row or 0
        priority0_rows: List[int] = []
        for r in range(header_row + 1, max_r + 1):
            if not _is_priority_zero(ws.cell(row=r, column=c_priority).value):
                continue
            subj = _cell_text_normalized(ws.cell(row=r, column=c_subject).value)
            if not subj:
                continue
            priority0_rows.append(r)

        rows_priority0 = len(priority0_rows)
        rows_scanned = 0
        rows_written = 0
        rows_skipped = 0

        bal_cache: Dict[int, float] = {}

        def _col_balance(c: int) -> float:
            if c not in bal_cache:
                bal_cache[c] = _read_schedule_balance(ws, c)
            return bal_cache[c]

        def _reduce_balance(c: int, amount: float) -> None:
            new_b = _col_balance(c) - amount
            bal_cache[c] = new_b
            ws.cell(row=5, column=c, value=new_b)

        for r in priority0_rows:
            subj = _cell_text_normalized(ws.cell(row=r, column=c_subject).value)
            n_period = _protocol_period_months(ws.cell(row=r, column=c_agreement).value)
            if n_period is None:
                rows_skipped += 1
                continue

            n_total = 0.0
            for c in range(c_invoice + 1, c_cutoff):
                mn = _month_bucket_n(ws.cell(row=header_row, column=c).value)
                if mn is None:
                    continue
                if mn < n_period:
                    continue
                cell_val = ws.cell(row=r, column=c).value
                num = _coerce_numeric_for_sum(cell_val)
                if num is not None:
                    n_total += num

            rows_scanned += 1

            if n_total <= 0:
                rows_skipped += 1
                continue

            row_method = _cell_text_normalized(ws.cell(row=r, column=c_method).value)
            due_cell = ws.cell(row=r, column=c_due).value

            placed = False
            for c in payment_cols:
                bal = _col_balance(c)
                if bal < n_total:
                    continue
                if not _subject_contained_in_pay_region(subj, ws.cell(row=7, column=c).value):
                    continue
                if not _payment_method_any_token_match(row_method, ws.cell(row=8, column=c).value):
                    continue
                if not _loan_deadline_ok(ws, c, due_cell):
                    continue
                ws.cell(row=r, column=c, value=n_total)
                _reduce_balance(c, n_total)
                assert c_payment_plan is not None
                old_plan = _coerce_numeric_for_sum(
                    ws.cell(row=r, column=c_payment_plan).value
                )
                plan_base = float(old_plan) if old_plan is not None else 0.0
                ws.cell(row=r, column=c_payment_plan, value=plan_base + n_total)
                rows_written += 1
                placed = True
                break
            if not placed:
                rows_skipped += 1

        wb.save(path)
    finally:
        wb.close()

    return SmartScheduleResult(
        path=path,
        rows_priority0=rows_priority0,
        rows_scanned=rows_scanned,
        rows_written=rows_written,
        rows_skipped=rows_skipped,
    )
