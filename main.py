# -*- coding: utf-8 -*-
"""
智能排款：上传总表/明细模板、总表按优先级 0→1→2 排款、从 Sheet2 读取银行、写入可支付金额与付款方式。
工作区：程序目录\\AutoPaymentScheduleFile（TotalSheet / DetailTemplate）。

启动时：清空 TotalSheet、DetailTemplate 目录内容，界面数值归零（不读总表）。
上传总表保存到 TotalSheet 后，自动读取该目录下 Excel，按 summarize_card1_payable_from_total_sheet
汇总卡片 1 应付四行；展示完成后 nCount0/1/2/Total 内存变量清零（界面不变）。卡片 2 排款额仅智能排款后更新。

界面布局与色值对齐 Figma「智能排款界面」：
https://www.figma.com/design/aabHTb8OZnSlJBQPKrNhrZ/%E6%99BA%E8%83%BD%E6%8E%92%E6%AC%BE%E7%95%8C%E9%9D%A2?node-id=0-1
"""

from __future__ import annotations

import os
import shutil
import sys
import traceback

from PyQt6 import sip
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QFont, QFontMetrics, QIcon, QValidator
from PyQt6.QtWidgets import (
    QApplication,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QStyle,
    QVBoxLayout,
    QWidget,
)

from design_tokens import Figma, build_app_stylesheet, build_message_box_stylesheet
from detail_sheet_fill import fill_detail_workbook_from_total
from excel_service import (
    EXCEL_FILTER,
    BankOption,
    SHEET1_SUMMARY_LABELS,
    find_latest_total_excel,
    read_banks,
    read_sheet1_summary_strip,
    sum_sheet1_row4_after_payout_through_total_scheduled,
    write_bank_config_xlsx,
)
from owner_section_stats import (
    OwnerCostPanelStats,
    load_owner_cost_panel_stats,
    load_sheet4_cost_display_labels,
    sheet1_payment_plan_has_numeric_entries,
)
from paths import (
    clear_workspace_upload_dirs,
    detail_template_dir,
    ensure_workspace,
    get_base_dir,
    get_program_dir,
    total_sheet_dir,
)
from smart_schedule import (
    SmartScheduleResult,
    preprocess_total_sheet_after_upload,
    run_smart_schedule_on_total_sheet,
    summarize_card1_payable_from_total_sheet,
)

_EXCEL_SUFFIXES = (".xlsx", ".xls")


class _Percent0To100Validator(QValidator):
    """仅允许 0～100 的整数，禁止前导零（仅允许单独一个「0」表示零）；空串为 Intermediate。"""

    def validate(self, input_str: str, pos: int):
        if input_str == "":
            return QValidator.State.Intermediate, input_str, pos
        if not input_str.isdigit():
            return QValidator.State.Invalid, input_str, pos
        if len(input_str) > 1 and input_str.startswith("0"):
            return QValidator.State.Invalid, input_str, pos
        if len(input_str) > 3:
            return QValidator.State.Invalid, input_str, pos
        v = int(input_str)
        if v > 100:
            return QValidator.State.Invalid, input_str, pos
        return QValidator.State.Acceptable, input_str, pos


def _is_excel_path(path: str) -> bool:
    return path.lower().endswith(_EXCEL_SUFFIXES)


def _workspace_dir_is_empty(dir_path: str) -> bool:
    """目录不存在，或除点开头的系统文件外无任何条目，视为空。"""
    if not os.path.isdir(dir_path):
        return True
    for name in os.listdir(dir_path):
        if name.startswith("."):
            continue
        return False
    return True


def _format_payable_amount(n: float) -> str:
    """展示用：千分位 + 固定两位小数（如 8,734,102.39）。"""
    return f"{n:,.2f}"


# 「查看排款」前置：本会话未完成智能排款，或付款计划列无可解析数值（视为空）
_VIEW_SCHEDULE_PREREQ_MSG = "请先点击「智能排款」按钮进行排款。"

# 为 False 时不创建「银行可支付金额及付款方式」卡片，智能排款卡片上移为第二块
SHOW_BANK_SETTINGS_CARD = False


def _app_dir() -> str:
    return get_program_dir()


def _load_window_icon() -> QIcon:
    """加载窗口图标，按优先级尝试多种格式。"""
    base = _app_dir()
    for name in ("app.ico", "app.icns", "app_icon.png", "AutoPay.jpeg"):
        path = os.path.join(base, name)
        if os.path.isfile(path):
            return QIcon(path)
    return QIcon()


class MainWindow(QWidget):
    def __init__(self) -> None:
        super().__init__()
        self.setObjectName("centralRoot")
        self._total_path: str | None = None
        self._banks: list[BankOption] = []
        self._refreshing_bank_combo = False
        self._sheet1_strip_row: int | None = None
        self._sheet1_strip_arrays: dict[str, list[str]] = {
            k: [] for k in SHEET1_SUMMARY_LABELS
        }
        # 卡片 1 应付汇总内存值（与界面 lbl_due_* 独立；展示后清零，避免下次计算累加旧值）
        self._n_count0 = 0.0
        self._n_count1 = 0.0
        self._n_count2 = 0.0
        self._n_count_total = 0.0
        self._smart_schedule_completed_ok = False
        self._owner_cost_label_col_w = 0
        self._owner_cost_amount_col_w = 0
        self._owner_cost_yuan_col_w = 0
        self._initialize_program()
        self._setup_ui()
        self._init_owner_cost_column_metrics()
        self._apply_styles()
        if SHOW_BANK_SETTINGS_CARD:
            self._refresh_pay_method_combo()
            self._refresh_bank_combo()
        self._clear_startup_numeric_fields()
        self._reset_card1_payable_count_vars()
        self._total_path = None
        if not SHOW_BANK_SETTINGS_CARD:
            self._rebuild_sheet1_strip_cache(None)
        self._refresh_owner_cost_panels(None)

    def _initialize_program(self) -> None:
        """程序启动：创建工作区并清空 TotalSheet、DetailTemplate（不读取总表）。"""
        ensure_workspace()
        clear_workspace_upload_dirs()

    def _clear_startup_numeric_fields(self) -> None:
        z = _format_payable_amount(0.0)
        self.lbl_due_strict.setText(z)
        self.lbl_due_priority1.setText(z)
        self.lbl_due_priority2.setText(z)
        self.lbl_due_month_total.setText(z)
        self.lbl_payable_total_value.setText(z)
        self.lbl_reserved_total_value.setText(z)
        self.lbl_sched_month_total.setText(z)
        self.lbl_sched_strict.setText(z)
        self.lbl_sched_priority1.setText(z)
        self.lbl_sched_priority2.setText(z)
        self.edit_priority1_pct.clear()
        self.edit_priority2_pct.clear()
        if SHOW_BANK_SETTINGS_CARD:
            self.edit_amount.setText(z)

    def _reset_card1_payable_count_vars(self) -> None:
        """nCount0/1/2/Total 内存变量清零（不改卡片 1 界面显示）。"""
        self._n_count0 = 0.0
        self._n_count1 = 0.0
        self._n_count2 = 0.0
        self._n_count_total = 0.0

    @staticmethod
    def _lbl_purple(text: str) -> QLabel:
        w = QLabel(text)
        w.setObjectName("figmaPurpleLabel")
        return w

    @staticmethod
    def _lbl_yuan_blue() -> QLabel:
        w = QLabel("元")
        w.setObjectName("suffixYuan")
        return w

    @staticmethod
    def _suffix_unit(text: str) -> QLabel:
        w = QLabel(text)
        w.setObjectName("suffixUnit")
        return w

    @staticmethod
    def _lbl_section_title(text: str) -> QLabel:
        w = QLabel(text)
        w.setObjectName("figmaSectionTitle")
        w.setWordWrap(True)
        w.setAlignment(
            Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter
        )
        return w

    # 卡片 3：标签 | 4px | 金额 | 2px | 「元」（四行共用列宽，查看排款后仍纵向对齐）
    _OWNER_COST_GAP_LABEL_TO_AMOUNT = 4
    _OWNER_COST_GAP_AMOUNT_TO_YUAN = 2

    def _init_owner_cost_column_metrics(self) -> None:
        amt_font = QFont()
        amt_font.setPointSize(Figma.BODY_PT + 1)
        amt_font.setBold(True)
        fm_amt = QFontMetrics(amt_font)
        self._owner_cost_amount_col_w = (
            fm_amt.horizontalAdvance("999,999,999,999,999,999.99") + 4
        )
        self._owner_cost_yuan_col_w = max(fm_amt.horizontalAdvance("元"), 18)

    def _ensure_owner_cost_label_col_w(self, cost_labels: tuple[str, ...]) -> None:
        """按本批成本名称（及「合计」）统一标签列宽。"""
        lbl_font = QFont()
        lbl_font.setPointSize(Figma.LABEL_PT)
        fm_lbl = QFontMetrics(lbl_font)
        texts = [f"{t}：" for t in cost_labels] + ["合计："]
        if not texts:
            texts = ["合计："]
        self._owner_cost_label_col_w = max(fm_lbl.horizontalAdvance(t) for t in texts) + 2

    def _owner_cost_block_width(self) -> int:
        return (
            self._owner_cost_label_col_w
            + self._OWNER_COST_GAP_LABEL_TO_AMOUNT
            + self._owner_cost_amount_col_w
            + self._OWNER_COST_GAP_AMOUNT_TO_YUAN
            + self._owner_cost_yuan_col_w
        )

    def _cost_stat_cell(self, label_text: str) -> tuple[QWidget, QLabel]:
        """成本名称或合计：固定列宽，标签/金额/「元」各行纵向对齐。"""
        if self._owner_cost_amount_col_w <= 0:
            self._init_owner_cost_column_metrics()
        block_w = self._owner_cost_block_width()
        w = QWidget()
        w.setFixedWidth(block_w)
        w.setSizePolicy(
            QSizePolicy.Policy.Fixed,
            QSizePolicy.Policy.Preferred,
        )
        outer = QHBoxLayout(w)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)

        lbl = self._lbl_purple(label_text)
        lbl.setFixedWidth(self._owner_cost_label_col_w)
        lbl.setAlignment(
            Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter
        )
        outer.addWidget(lbl, alignment=Qt.AlignmentFlag.AlignVCenter)
        outer.addSpacing(self._OWNER_COST_GAP_LABEL_TO_AMOUNT)

        val = QLabel("")
        val.setObjectName("summaryValue")
        val.setFixedWidth(self._owner_cost_amount_col_w)
        val.setAlignment(
            Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter
        )
        val.setSizePolicy(
            QSizePolicy.Policy.Fixed,
            QSizePolicy.Policy.Preferred,
        )
        outer.addWidget(val, alignment=Qt.AlignmentFlag.AlignVCenter)
        outer.addSpacing(self._OWNER_COST_GAP_AMOUNT_TO_YUAN)

        yuan = self._lbl_yuan_blue()
        yuan.setFixedWidth(self._owner_cost_yuan_col_w)
        yuan.setAlignment(
            Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter
        )
        outer.addWidget(yuan, alignment=Qt.AlignmentFlag.AlignVCenter)
        return w, val

    @staticmethod
    def _clear_hbox_layout(layout: QHBoxLayout) -> None:
        """清空横向布局：立即脱离父控件，避免 deleteLater 延迟导致新旧控件叠画。"""
        while layout.count():
            item = layout.takeAt(0)
            w = item.widget()
            if w is not None:
                w.setParent(None)
                w.deleteLater()
            sip.delete(item)

    def _fill_cost_summary_row_static_zeros(
        self,
        layout: QHBoxLayout,
        cost_labels: tuple[str, ...] | None = None,
    ) -> None:
        """无汇总数据时占位；若提供 cost_labels 则与 Sheet4 顺序、标签一致。"""
        layout.setSpacing(Figma.DATA_ROW_GAP)
        z = _format_payable_amount(0.0)
        titles = cost_labels if cost_labels else ("成本二部", "成本三部")
        if not titles:
            titles = ("成本二部", "成本三部")
        self._ensure_owner_cost_label_col_w(titles)
        for title in titles:
            cell, lbl = self._cost_stat_cell(f"{title}：")
            lbl.setText(z)
            layout.addWidget(cell, 0)
        tot_w, tot_lbl = self._cost_stat_cell("合计：")
        tot_lbl.setText(z)
        layout.addWidget(tot_w, 0)
        layout.addStretch(1)

    def _fill_cost_summary_row(
        self,
        layout: QHBoxLayout,
        stats,
        *,
        panel_mode: str,
        fallback_labels: tuple[str, ...] | None = None,
    ) -> None:
        """panel_mode: zero | priority1 | priority2 | all"""
        if stats is None:
            self._fill_cost_summary_row_static_zeros(layout, fallback_labels)
            return
        layout.setSpacing(Figma.DATA_ROW_GAP)
        self._ensure_owner_cost_label_col_w(stats.cost_labels)
        if panel_mode == "zero":
            amounts = stats.per_owner_zero
            total_val = stats.sum_zero
        elif panel_mode == "priority1":
            amounts = stats.per_owner_priority1
            total_val = stats.sum_priority1
        elif panel_mode == "priority2":
            amounts = stats.per_owner_priority2
            total_val = stats.sum_priority2
        else:
            amounts = stats.per_owner_all
            total_val = stats.sum_all
        for title, amt in zip(stats.cost_labels, amounts):
            cell, lbl = self._cost_stat_cell(f"{title}：")
            lbl.setText(_format_payable_amount(amt))
            layout.addWidget(cell, 0)
        tot_w, tot_lbl = self._cost_stat_cell("合计：")
        tot_lbl.setText(_format_payable_amount(total_val))
        layout.addWidget(tot_w, 0)
        layout.addStretch(1)

    def _refresh_owner_bottom_placeholder_from_path(self, path: str) -> None:
        """付款计划为空等：底部四行仅按 Sheet4 标签填 0.00，不汇总 Sheet1。"""
        fb = None
        if path and os.path.isfile(path) and str(path).lower().endswith(".xlsx"):
            try:
                fb = load_sheet4_cost_display_labels(path)
            except Exception:
                traceback.print_exc()
        self._clear_hbox_layout(self._lay_strict_cost_row)
        self._clear_hbox_layout(self._lay_priority1_cost_row)
        self._clear_hbox_layout(self._lay_priority2_cost_row)
        self._clear_hbox_layout(self._lay_overall_cost_row)
        self._fill_cost_summary_row(
            self._lay_strict_cost_row, None, panel_mode="zero", fallback_labels=fb
        )
        self._fill_cost_summary_row(
            self._lay_priority1_cost_row, None, panel_mode="priority1", fallback_labels=fb
        )
        self._fill_cost_summary_row(
            self._lay_priority2_cost_row, None, panel_mode="priority2", fallback_labels=fb
        )
        self._fill_cost_summary_row(
            self._lay_overall_cost_row, None, panel_mode="all", fallback_labels=fb
        )

    def _refresh_owner_cost_panels(self, path: str | None) -> OwnerCostPanelStats | None:
        """底部卡片 3 四行各部门明细；path 为 None 时占位 0.00。返回统计对象（供调用方扩展用）。"""
        self._clear_hbox_layout(self._lay_strict_cost_row)
        self._clear_hbox_layout(self._lay_priority1_cost_row)
        self._clear_hbox_layout(self._lay_priority2_cost_row)
        self._clear_hbox_layout(self._lay_overall_cost_row)

        fallback_labels: tuple[str, ...] | None = None
        if path and os.path.isfile(path) and str(path).lower().endswith(".xlsx"):
            try:
                fallback_labels = load_sheet4_cost_display_labels(path)
            except Exception:
                traceback.print_exc()
                fallback_labels = None

        if (
            not path
            or not os.path.isfile(path)
            or not str(path).lower().endswith(".xlsx")
        ):
            self._fill_cost_summary_row(
                self._lay_strict_cost_row,
                None,
                panel_mode="zero",
                fallback_labels=fallback_labels,
            )
            self._fill_cost_summary_row(
                self._lay_priority1_cost_row,
                None,
                panel_mode="priority1",
                fallback_labels=fallback_labels,
            )
            self._fill_cost_summary_row(
                self._lay_priority2_cost_row,
                None,
                panel_mode="priority2",
                fallback_labels=fallback_labels,
            )
            self._fill_cost_summary_row(
                self._lay_overall_cost_row,
                None,
                panel_mode="all",
                fallback_labels=fallback_labels,
            )
            return None
        try:
            stats = load_owner_cost_panel_stats(path)
        except Exception:
            traceback.print_exc()
            stats = None
        self._fill_cost_summary_row(
            self._lay_strict_cost_row,
            stats,
            panel_mode="zero",
            fallback_labels=fallback_labels,
        )
        self._fill_cost_summary_row(
            self._lay_priority1_cost_row,
            stats,
            panel_mode="priority1",
            fallback_labels=fallback_labels,
        )
        self._fill_cost_summary_row(
            self._lay_priority2_cost_row,
            stats,
            panel_mode="priority2",
            fallback_labels=fallback_labels,
        )
        self._fill_cost_summary_row(
            self._lay_overall_cost_row,
            stats,
            panel_mode="all",
            fallback_labels=fallback_labels,
        )
        return stats

    @staticmethod
    def _right_align_cell(width_px: int, widget: QWidget) -> QWidget:
        cell = QWidget()
        h = QHBoxLayout(cell)
        h.setContentsMargins(0, 0, 0, 0)
        h.addStretch(1)
        h.addWidget(widget)
        cell.setFixedWidth(width_px)
        return cell

    def _setup_ui(self) -> None:
        self.setWindowTitle("智能排款")
        self.setMinimumSize(820, 400)
        self.resize(960, 540)
        self.setWindowIcon(_load_window_icon())

        m = Figma
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        scroll.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        scroll.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding
        )

        scroll_inner = QWidget()
        scroll_layout = QVBoxLayout(scroll_inner)
        scroll_layout.setContentsMargins(
            m.PAGE_MARGIN_H, m.PAGE_MARGIN_V, m.PAGE_MARGIN_H, m.PAGE_MARGIN_V
        )
        scroll_layout.setSpacing(m.GAP_SECTION)

        scroll.setWidget(scroll_inner)
        root.addWidget(scroll, 1)

        # —— 共享：金额列宽度（顶部应付卡片 + 中部表单）——
        _amt_font = QFont()
        _amt_font.setPointSize(m.BODY_PT + 1)
        _amt_font.setBold(True)
        _fm_amt = QFontMetrics(_amt_font)
        _worst_amount = _format_payable_amount(999999999999999999.99)
        _w_num = _fm_amt.horizontalAdvance(_worst_amount)
        _w_yuan = max(_fm_amt.horizontalAdvance("元"), 18)
        _w_pct = max(_fm_amt.horizontalAdvance("%"), 14)
        _pad = 8
        w_left_tail = max(
            m.W_PCT_TAIL,
            _w_num + m.GAP_TIGHT + _w_yuan + _pad,
            m.W_PCT_INPUT + m.GAP_TIGHT + _w_pct + _pad,
        )
        w_right_gap = m.GAP_TIGHT * 4
        w_right_tail = max(
            m.W_PCT_TAIL + m.GAP_TIGHT * 2,
            _w_num + w_right_gap + _w_yuan + _pad,
        )

        _lbl_font = QFont()
        _lbl_font.setPointSize(m.LABEL_PT)
        _fm_lbl = QFontMetrics(_lbl_font)

        def _value_yuan_cell(value_lbl: QLabel) -> QWidget:
            cell = QWidget()
            cell.setFixedWidth(w_left_tail)
            cell.setMaximumWidth(w_left_tail)
            hh = QHBoxLayout(cell)
            hh.setContentsMargins(0, 0, 0, 0)
            hh.setSpacing(2)
            hh.addWidget(value_lbl, 0, Qt.AlignmentFlag.AlignLeft)
            hh.addWidget(self._lbl_yuan_blue(), 0, Qt.AlignmentFlag.AlignLeft)
            return cell

        def _value_yuan_cell_wide(value_lbl: QLabel) -> QWidget:
            cell = QWidget()
            cell.setFixedWidth(w_right_tail)
            cell.setMaximumWidth(w_right_tail)
            hh = QHBoxLayout(cell)
            hh.setContentsMargins(0, 0, 0, 0)
            hh.setSpacing(2)
            hh.addWidget(value_lbl, 0, Qt.AlignmentFlag.AlignLeft)
            hh.addWidget(self._lbl_yuan_blue(), 0, Qt.AlignmentFlag.AlignLeft)
            return cell

        def _pct_tail_widget(edit: QLineEdit) -> QWidget:
            cell = QWidget()
            cell.setFixedWidth(w_left_tail)
            cell.setMaximumWidth(w_left_tail)
            hh = QHBoxLayout(cell)
            hh.setContentsMargins(0, 0, 0, 0)
            hh.setSpacing(2)
            hh.addWidget(edit, alignment=Qt.AlignmentFlag.AlignVCenter)
            hh.addWidget(self._suffix_unit("%"), alignment=Qt.AlignmentFlag.AlignVCenter)
            return cell

        def _amount_value_lbl() -> QLabel:
            lbl = QLabel("")
            lbl.setObjectName("summaryValue")
            lbl.setAlignment(
                Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter
            )
            lbl.setSizePolicy(
                QSizePolicy.Policy.Maximum,
                QSizePolicy.Policy.Preferred,
            )
            return lbl

        def _label_value_row(lbl: QLabel, value_w: QWidget) -> QWidget:
            row = QWidget()
            row.setSizePolicy(
                QSizePolicy.Policy.Maximum,
                QSizePolicy.Policy.Preferred,
            )
            hh = QHBoxLayout(row)
            hh.setContentsMargins(0, 0, 0, 0)
            hh.setSpacing(4)
            hh.addWidget(lbl, alignment=Qt.AlignmentFlag.AlignVCenter)
            hh.addWidget(value_w, alignment=Qt.AlignmentFlag.AlignVCenter)
            return row

        _sync_label_texts = (
            "严格按账期应付额：",
            "第一优先级应付额：",
            "第二优先级应付额：",
            "本月合计应付总额：",
            "本月可支付总额：",
            "本月预留款总额：",
            "第一优先级支付计划占应付额：",
            "第二优先级支付计划占应付额：",
            "本月合计排款总额：",
            "严格按账期排款额：",
            "第一优先级排款额：",
            "第二优先级排款额：",
        )
        w_sync_label_col = max(_fm_lbl.horizontalAdvance(t) for t in _sync_label_texts)

        def _lbl_purple_fixed(text: str) -> QLabel:
            w = QLabel(text)
            w.setObjectName("figmaPurpleLabel")
            w.setFixedWidth(w_sync_label_col)
            w.setAlignment(
                Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter
            )
            return w

        # —— 卡片：顶部应付额汇总（Figma）——
        card_summary = QFrame()
        card_summary.setObjectName("elevatedCard")
        card_summary.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)
        lay_summary = QVBoxLayout(card_summary)
        lay_summary.setContentsMargins(m.CARD_PAD_H, m.CARD_PAD_V, m.CARD_PAD_H, m.CARD_PAD_V)
        lay_summary.setSpacing(m.ROW_GAP)

        self.lbl_due_strict = _amount_value_lbl()
        self.lbl_due_priority1 = _amount_value_lbl()
        self.lbl_due_priority2 = _amount_value_lbl()
        self.lbl_due_month_total = _amount_value_lbl()
        lay_summary.addWidget(
            _label_value_row(
                _lbl_purple_fixed("严格按账期应付额："),
                _value_yuan_cell(self.lbl_due_strict),
            ),
            0,
            Qt.AlignmentFlag.AlignLeft,
        )
        lay_summary.addWidget(
            _label_value_row(
                _lbl_purple_fixed("第一优先级应付额："),
                _value_yuan_cell(self.lbl_due_priority1),
            ),
            0,
            Qt.AlignmentFlag.AlignLeft,
        )
        lay_summary.addWidget(
            _label_value_row(
                _lbl_purple_fixed("第二优先级应付额："),
                _value_yuan_cell(self.lbl_due_priority2),
            ),
            0,
            Qt.AlignmentFlag.AlignLeft,
        )
        lay_summary.addWidget(
            _label_value_row(
                _lbl_purple_fixed("本月合计应付总额："),
                _value_yuan_cell(self.lbl_due_month_total),
            ),
            0,
            Qt.AlignmentFlag.AlignLeft,
        )
        scroll_layout.addWidget(card_summary)

        # —— 卡片：上传 + 可支付/预留/占比 + 排款额 + 智能排款 ——
        card_ops = QFrame()
        card_ops.setObjectName("elevatedCard")
        card_ops.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)
        lay_ops = QVBoxLayout(card_ops)
        lay_ops.setContentsMargins(m.CARD_PAD_H, m.CARD_PAD_V, m.CARD_PAD_H, m.CARD_PAD_V)
        lay_ops.setSpacing(m.ROW_GAP)

        row_upload = QHBoxLayout()
        row_upload.setSpacing(m.GAP_INLINE)
        self.btn_upload_total = QPushButton("上传排款计划总表")
        self.btn_upload_total.setObjectName("figmaUploadBtn")
        self.btn_upload_total.clicked.connect(self._on_upload_total)
        self.btn_upload_template = QPushButton("上传排款明细模版")
        self.btn_upload_template.setObjectName("figmaUploadBtn")
        self.btn_upload_template.clicked.connect(self._on_upload_template)
        row_upload.addWidget(self.btn_upload_total)
        row_upload.addWidget(self.btn_upload_template)
        row_upload.addStretch(1)
        lay_ops.addLayout(row_upload)

        self.lbl_payable_total_value = _amount_value_lbl()
        self.lbl_reserved_total_value = _amount_value_lbl()
        self.lbl_sched_month_total = _amount_value_lbl()
        self.lbl_sched_strict = _amount_value_lbl()
        self.lbl_sched_priority1 = _amount_value_lbl()
        self.lbl_sched_priority2 = _amount_value_lbl()

        self.edit_priority1_pct = QLineEdit()
        self.edit_priority1_pct.setObjectName("figmaInput")
        self.edit_priority1_pct.setMaxLength(3)
        self.edit_priority1_pct.setValidator(_Percent0To100Validator())
        self.edit_priority1_pct.setFixedWidth(m.W_PCT_INPUT)
        self.edit_priority1_pct.setSizePolicy(
            QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed
        )
        self.edit_priority2_pct = QLineEdit()
        self.edit_priority2_pct.setObjectName("figmaInput")
        self.edit_priority2_pct.setMaxLength(3)
        self.edit_priority2_pct.setValidator(_Percent0To100Validator())
        self.edit_priority2_pct.setFixedWidth(m.W_PCT_INPUT)
        self.edit_priority2_pct.setSizePolicy(
            QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed
        )

        left_col = QWidget()
        lv = QVBoxLayout(left_col)
        lv.setContentsMargins(0, 0, 0, 0)
        lv.setSpacing(m.ROW_GAP)
        lv.addWidget(
            _label_value_row(
                _lbl_purple_fixed("本月可支付总额："),
                _value_yuan_cell(self.lbl_payable_total_value),
            ),
            0,
            Qt.AlignmentFlag.AlignLeft,
        )
        lv.addWidget(
            _label_value_row(
                _lbl_purple_fixed("本月预留款总额："),
                _value_yuan_cell(self.lbl_reserved_total_value),
            ),
            0,
            Qt.AlignmentFlag.AlignLeft,
        )
        lv.addWidget(
            _label_value_row(
                _lbl_purple_fixed("第一优先级支付计划占应付额："),
                _pct_tail_widget(self.edit_priority1_pct),
            ),
            0,
            Qt.AlignmentFlag.AlignLeft,
        )
        lv.addWidget(
            _label_value_row(
                _lbl_purple_fixed("第二优先级支付计划占应付额："),
                _pct_tail_widget(self.edit_priority2_pct),
            ),
            0,
            Qt.AlignmentFlag.AlignLeft,
        )

        right_col = QWidget()
        rv = QVBoxLayout(right_col)
        rv.setContentsMargins(0, 0, 0, 0)
        rv.setSpacing(m.ROW_GAP)
        rv.addWidget(
            _label_value_row(
                _lbl_purple_fixed("严格按账期排款额："),
                _value_yuan_cell_wide(self.lbl_sched_strict),
            ),
            0,
            Qt.AlignmentFlag.AlignLeft,
        )
        rv.addWidget(
            _label_value_row(
                _lbl_purple_fixed("第一优先级排款额："),
                _value_yuan_cell_wide(self.lbl_sched_priority1),
            ),
            0,
            Qt.AlignmentFlag.AlignLeft,
        )
        rv.addWidget(
            _label_value_row(
                _lbl_purple_fixed("第二优先级排款额："),
                _value_yuan_cell_wide(self.lbl_sched_priority2),
            ),
            0,
            Qt.AlignmentFlag.AlignLeft,
        )
        rv.addWidget(
            _label_value_row(
                _lbl_purple_fixed("本月合计排款总额："),
                _value_yuan_cell_wide(self.lbl_sched_month_total),
            ),
            0,
            Qt.AlignmentFlag.AlignLeft,
        )

        card1_form = QWidget()
        card1_row = QHBoxLayout(card1_form)
        card1_row.setContentsMargins(0, 0, 0, 0)
        card1_row.setSpacing(0)
        card1_row.addWidget(left_col, 0, Qt.AlignmentFlag.AlignTop)
        card1_row.addSpacing(m.GAP_INLINE)
        card1_row.addWidget(right_col, 0, Qt.AlignmentFlag.AlignTop)
        card1_row.addStretch(1)
        self.btn_smart = QPushButton("智能排款")
        self.btn_smart.setObjectName("figmaBlue")
        self.btn_smart.setFixedWidth(80)
        self.btn_smart.clicked.connect(self._on_smart_placeholder)
        card1_row.addWidget(self.btn_smart, 0, Qt.AlignmentFlag.AlignVCenter)

        lay_ops.addWidget(card1_form)

        scroll_layout.addWidget(card_ops)

        if SHOW_BANK_SETTINGS_CARD:
            # —— 卡片 2：银行模块小标题 + 栅格 ——
            card_bank = QFrame()
            card_bank.setObjectName("elevatedCard")
            card_bank.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)
            lay_bank = QVBoxLayout(card_bank)
            lay_bank.setContentsMargins(m.CARD_PAD_H, m.CARD_PAD_V, m.CARD_PAD_H, m.CARD_PAD_V)
            lay_bank.setSpacing(14)

            micro = QLabel("银行可支付金额及付款方式设置")
            micro.setObjectName("figmaMicroTitle")
            lay_bank.addWidget(micro)

            # —— 银行/付款方式 + 可支付金额/确定（栅格：第 3 列与付款方式下拉同宽，确定右缘对齐）——
            va = Qt.AlignmentFlag.AlignVCenter
            ra = Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter

            pair = QWidget()
            grid = QGridLayout(pair)
            grid.setContentsMargins(0, 0, 0, 0)
            grid.setHorizontalSpacing(m.GAP_INLINE)
            grid.setVerticalSpacing(12)
            grid.setColumnMinimumWidth(1, m.W_COMBO_BANK)
            grid.setColumnMinimumWidth(3, m.W_COMBO_PAY_METHOD)
            grid.setColumnStretch(4, 1)

            grid.addWidget(self._lbl_purple("银 行 选 择："), 0, 0, ra)
            self.combo_bank = QComboBox()
            self.combo_bank.setObjectName("figmaCombo")
            self.combo_bank.setFixedSize(m.W_COMBO_BANK, 40)
            self.combo_bank.setSizePolicy(
                QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed
            )
            self.combo_bank.currentIndexChanged.connect(self._on_bank_index_changed)
            grid.addWidget(self.combo_bank, 0, 1, va)
            grid.addWidget(self._lbl_purple("付款方式："), 0, 2, ra)
            self.combo_pay_method = QComboBox()
            self.combo_pay_method.setObjectName("figmaCombo")
            self.combo_pay_method.setFixedSize(m.W_COMBO_PAY_METHOD, 40)
            self.combo_pay_method.setSizePolicy(
                QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed
            )
            grid.addWidget(self.combo_pay_method, 0, 3, va)

            w_amt = QWidget()
            lay_amt = QHBoxLayout(w_amt)
            lay_amt.setContentsMargins(0, 0, 0, 0)
            lay_amt.setSpacing(m.GAP_TIGHT)
            lay_amt.addWidget(self._lbl_purple("可支付金额："))
            self.edit_amount = QLineEdit()
            self.edit_amount.setObjectName("figmaAmountInput")
            self.edit_amount.setPlaceholderText("")
            self.edit_amount.setValidator(QDoubleValidator(0.0, 1e15, 2))
            self.edit_amount.setMinimumHeight(40)
            self.edit_amount.setMinimumWidth(50)
            self.edit_amount.setMaximumWidth(m.W_AMOUNT_INPUT)
            self.edit_amount.setSizePolicy(
                QSizePolicy.Policy.Maximum, QSizePolicy.Policy.Fixed
            )
            lay_amt.addWidget(self.edit_amount)
            lay_amt.addWidget(self._lbl_yuan_blue())
            lay_amt.addStretch(1)
            grid.addWidget(w_amt, 1, 0, 1, 3, va)

            self.btn_confirm = QPushButton("确定")
            self.btn_confirm.setObjectName("figmaBlueSmall")
            self.btn_confirm.setMinimumHeight(44)
            self.btn_confirm.clicked.connect(self._on_confirm)
            grid.addWidget(
                self._right_align_cell(m.W_COMBO_PAY_METHOD, self.btn_confirm), 1, 3, va
            )

            lay_bank.addWidget(pair)
            scroll_layout.addWidget(card_bank)

        # —— 底部卡片：四行各部门明细 + 右侧「查看排款」（Figma）——
        _cost_row_min_h = max(_fm_amt.height(), _fm_lbl.height()) + 14

        card_bottom = QFrame()
        card_bottom.setObjectName("elevatedCard")
        card_bottom.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)
        card_bottom.setSizePolicy(
            QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Minimum
        )
        lay_bottom_outer = QHBoxLayout(card_bottom)
        lay_bottom_outer.setContentsMargins(m.CARD_PAD_H, m.CARD_PAD_V, m.CARD_PAD_H, m.CARD_PAD_V)
        lay_bottom_outer.setSpacing(m.GAP_INLINE)

        w_bottom_left = QWidget()
        w_bottom_left.setSizePolicy(
            QSizePolicy.Policy.MinimumExpanding, QSizePolicy.Policy.Minimum
        )
        lay_bottom_left = QVBoxLayout(w_bottom_left)
        lay_bottom_left.setSpacing(0)

        lay_bottom_left.addWidget(self._lbl_section_title("严格按账期付款-各部门明细"))
        lay_bottom_left.addSpacing(12)
        self._w_strict_cost_row = QWidget()
        self._w_strict_cost_row.setMinimumHeight(_cost_row_min_h)
        self._lay_strict_cost_row = QHBoxLayout(self._w_strict_cost_row)
        self._lay_strict_cost_row.setContentsMargins(0, 0, 0, 0)
        self._lay_strict_cost_row.setSpacing(m.DATA_ROW_GAP)
        lay_bottom_left.addWidget(self._w_strict_cost_row)

        lay_bottom_left.addSpacing(m.ROW_GAP)
        lay_bottom_left.addWidget(self._lbl_section_title("第一优先级排款-各部门明细"))
        lay_bottom_left.addSpacing(12)
        self._w_priority1_cost_row = QWidget()
        self._w_priority1_cost_row.setMinimumHeight(_cost_row_min_h)
        self._lay_priority1_cost_row = QHBoxLayout(self._w_priority1_cost_row)
        self._lay_priority1_cost_row.setContentsMargins(0, 0, 0, 0)
        self._lay_priority1_cost_row.setSpacing(m.DATA_ROW_GAP)
        lay_bottom_left.addWidget(self._w_priority1_cost_row)

        lay_bottom_left.addSpacing(m.ROW_GAP)
        lay_bottom_left.addWidget(self._lbl_section_title("第二优先级排款-各部门明细"))
        lay_bottom_left.addSpacing(12)
        self._w_priority2_cost_row = QWidget()
        self._w_priority2_cost_row.setMinimumHeight(_cost_row_min_h)
        self._lay_priority2_cost_row = QHBoxLayout(self._w_priority2_cost_row)
        self._lay_priority2_cost_row.setContentsMargins(0, 0, 0, 0)
        self._lay_priority2_cost_row.setSpacing(m.DATA_ROW_GAP)
        lay_bottom_left.addWidget(self._w_priority2_cost_row)

        lay_bottom_left.addSpacing(m.ROW_GAP)
        lay_bottom_left.addWidget(self._lbl_section_title("本月合计总排款-各部门明细"))
        lay_bottom_left.addSpacing(12)
        self._w_overall_cost_row = QWidget()
        self._w_overall_cost_row.setMinimumHeight(_cost_row_min_h)
        self._lay_overall_cost_row = QHBoxLayout(self._w_overall_cost_row)
        self._lay_overall_cost_row.setContentsMargins(0, 0, 0, 0)
        self._lay_overall_cost_row.setSpacing(m.DATA_ROW_GAP)
        lay_bottom_left.addWidget(self._w_overall_cost_row)

        lay_bottom_outer.addWidget(w_bottom_left, 1)

        lay_btn_col = QVBoxLayout()
        lay_btn_col.addStretch(1)
        self.btn_view_schedule = QPushButton("查看排款")
        self.btn_view_schedule.setObjectName("figmaBlue")
        self.btn_view_schedule.setFixedWidth(80)
        self.btn_view_schedule.clicked.connect(self._on_view_schedule_placeholder)
        lay_btn_col.addWidget(
            self.btn_view_schedule,
            0,
            Qt.AlignmentFlag.AlignHCenter,
        )
        lay_btn_col.addStretch(1)
        lay_bottom_outer.addLayout(lay_btn_col, 0)

        scroll_layout.addWidget(card_bottom)

        hint = QLabel(f"工作目录：{get_base_dir()}")
        hint.setObjectName("footerHint")
        hint.setWordWrap(True)
        footer_bar = QWidget()
        footer_lay = QHBoxLayout(footer_bar)
        footer_lay.setContentsMargins(
            m.PAGE_MARGIN_H, 10, m.PAGE_MARGIN_H, m.PAGE_MARGIN_V
        )
        footer_lay.addWidget(hint)
        root.addWidget(footer_bar, 0)

    def _apply_styles(self) -> None:
        self.setStyleSheet(build_app_stylesheet())

    def _rebuild_sheet1_strip_cache(self, path: str | None) -> None:
        """从首表 Sheet1 解析五个表头词所在行，横向片段写入 _sheet1_strip_arrays 备用。"""
        self._sheet1_strip_row = None
        for k in SHEET1_SUMMARY_LABELS:
            self._sheet1_strip_arrays[k] = []
        if not path or not os.path.isfile(path):
            return
        try:
            res = read_sheet1_summary_strip(path)
            self._sheet1_strip_row = res.row_1based
            for k in SHEET1_SUMMARY_LABELS:
                self._sheet1_strip_arrays[k] = list(res.columns.get(k, ()))
        except Exception:
            traceback.print_exc()

    def _apply_card1_payable_from_total_sheet(self, path: str) -> None:
        """读取 TotalSheet 下总表，按账期窗口与付款优先级 0/1/2 汇总卡片 1 应付四行并展示；展示后内存变量清零。"""
        z = _format_payable_amount(0.0)
        caps = None
        try:
            caps = summarize_card1_payable_from_total_sheet(path)
        except Exception:
            traceback.print_exc()
        if caps is not None:
            self._n_count0 = caps.n_count0
            self._n_count1 = caps.n_count1
            self._n_count2 = caps.n_count2
            self._n_count_total = caps.n_count_total
            self.lbl_due_strict.setText(_format_payable_amount(caps.n_count0))
            self.lbl_due_priority1.setText(_format_payable_amount(caps.n_count1))
            self.lbl_due_priority2.setText(_format_payable_amount(caps.n_count2))
            self.lbl_due_month_total.setText(
                _format_payable_amount(caps.n_count_total)
            )
            self._reset_card1_payable_count_vars()
        else:
            self.lbl_due_strict.setText(z)
            self.lbl_due_priority1.setText(z)
            self.lbl_due_priority2.setText(z)
            self.lbl_due_month_total.setText(z)
            self._reset_card1_payable_count_vars()

    def _apply_payable_total_from_sheet(self, path: str | None) -> None:
        """上传总表后：读 TotalSheet 下 Excel，刷新「本月可支付总额」与卡片 1 应付四行；不更新排款额。"""
        if not path or not os.path.isfile(path):
            return
        z = _format_payable_amount(0.0)
        try:
            total = sum_sheet1_row4_after_payout_through_total_scheduled(path)
        except Exception:
            traceback.print_exc()
            total = None
        if total is None:
            self.lbl_payable_total_value.setText(z)
        else:
            self.lbl_payable_total_value.setText(_format_payable_amount(total))
        self.lbl_reserved_total_value.setText(z)
        self._apply_card1_payable_from_total_sheet(path)

    def _on_upload_total(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self, "选择排款计划列表", "", EXCEL_FILTER
        )
        if not path:
            return
        if not _is_excel_path(path):
            QMessageBox.warning(
                self, "格式不支持", "请选择 Excel 文件（.xlsx 或 .xls）。"
            )
            return
        d = total_sheet_dir()
        os.makedirs(d, exist_ok=True)
        dest = os.path.join(d, os.path.basename(path))
        same_file = os.path.normcase(os.path.abspath(path)) == os.path.normcase(
            os.path.abspath(dest)
        )
        replaced = os.path.isfile(dest) and not same_file
        try:
            # shutil.copy2：目标已有同名文件时直接覆盖
            shutil.copy2(path, dest)
        except shutil.SameFileError:
            pass
        except OSError as e:
            QMessageBox.critical(self, "上传失败", str(e))
            return
        preprocess_ok = False
        if dest.lower().endswith(".xlsx"):
            try:
                preprocess_total_sheet_after_upload(dest)
                preprocess_ok = True
            except ValueError as e:
                QMessageBox.warning(
                    self,
                    "预处理提示",
                    f"总表已上传，但预处理未完成：\\n{e}",
                )
            except Exception as e:
                QMessageBox.warning(
                    self,
                    "预处理提示",
                    f"总表已上传，但预处理失败：\\n{e}",
                )
        self._total_path = dest
        self._smart_schedule_completed_ok = False
        msg = f"已保存到：\\n{dest}"
        if same_file:
            msg = f"所选文件已在目标目录中：\\n{dest}"
        elif replaced:
            msg += "\\n（已覆盖同名文件。）"
        if dest.lower().endswith(".xlsx") and preprocess_ok:
            msg += "\\n（已预处理并保存总表。）"
        QMessageBox.information(self, "上传成功", msg)
        if SHOW_BANK_SETTINGS_CARD:
            self._refresh_bank_combo()
        else:
            self._rebuild_sheet1_strip_cache(dest)
        self._apply_payable_total_from_sheet(dest)
        z0 = _format_payable_amount(0.0)
        self.lbl_sched_strict.setText(z0)
        self.lbl_sched_priority1.setText(z0)
        self.lbl_sched_priority2.setText(z0)
        self.lbl_sched_month_total.setText(z0)

    def _on_upload_template(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self, "选择排款明细模版", "", EXCEL_FILTER
        )
        if not path:
            return
        if not _is_excel_path(path):
            QMessageBox.warning(
                self, "格式不支持", "请选择 Excel 文件（.xlsx 或 .xls）。"
            )
            return
        d = detail_template_dir()
        os.makedirs(d, exist_ok=True)
        dest = os.path.join(d, os.path.basename(path))
        same_file = os.path.normcase(os.path.abspath(path)) == os.path.normcase(
            os.path.abspath(dest)
        )
        replaced = os.path.isfile(dest) and not same_file
        try:
            shutil.copy2(path, dest)
        except shutil.SameFileError:
            pass
        except OSError as e:
            QMessageBox.critical(self, "上传失败", str(e))
            return
        msg = f"已保存到：\\n{dest}"
        if same_file:
            msg = f"所选文件已在目标目录中：\\n{dest}"
        elif replaced:
            msg += "\\n（已覆盖同名文件。）"
        QMessageBox.information(self, "上传成功", msg)

    def _refresh_pay_method_combo(self) -> None:
        if not SHOW_BANK_SETTINGS_CARD:
            return
        self.combo_pay_method.clear()
        for s in PAY_METHOD_CHOICES:
            self.combo_pay_method.addItem(s)
        self.combo_pay_method.setCurrentIndex(0)

    def _on_bank_index_changed(self, _index: int) -> None:
        if not SHOW_BANK_SETTINGS_CARD:
            return
        if self._refreshing_bank_combo:
            return
        if self.combo_pay_method.count() > 0:
            self.combo_pay_method.setCurrentIndex(0)
        self.edit_amount.clear()

    def _refresh_bank_combo(self) -> None:
        if not SHOW_BANK_SETTINGS_CARD:
            return
        self._refreshing_bank_combo = True
        try:
            self.combo_bank.clear()
            self._banks = []
            ensure_workspace()
            latest = find_latest_total_excel(total_sheet_dir())
            self._total_path = latest
            if not latest:
                self.combo_bank.addItem("（请先上传排款计划总表）")
                self.combo_bank.setEnabled(False)
                return
            self.combo_bank.setEnabled(True)
            try:
                banks, _fname = read_banks(latest)
            except Exception as e:
                self.combo_bank.addItem(f"（读取失败：{e}）")
                self.combo_bank.setEnabled(False)
                QMessageBox.warning(self, "读取总表失败", str(e))
                return
            self._banks = banks
            self.combo_bank.addItem("不限")
            for b in banks:
                self.combo_bank.addItem(b.display())
            self.combo_bank.setCurrentIndex(0)
        finally:
            self._refreshing_bank_combo = False
            self._rebuild_sheet1_strip_cache(self._total_path)

    def _current_bank(self) -> BankOption | None:
        if not SHOW_BANK_SETTINGS_CARD:
            return None
        if not self._banks or not self.combo_bank.isEnabled():
            return None
        if self.combo_bank.count() == 0:
            return None
        if self.combo_bank.itemText(0) != "不限":
            return None
        idx = self.combo_bank.currentIndex()
        if idx <= 0:
            return None
        bi = idx - 1
        if 0 <= bi < len(self._banks):
            return self._banks[bi]
        return None

    def _on_confirm(self) -> None:
        if not SHOW_BANK_SETTINGS_CARD:
            return
        bank = self._current_bank()
        if bank is None:
            QMessageBox.warning(
                self,
                "提示",
                "请先上传排款计划总表，并在「银行选择」中选择具体银行（勿选「不限」）。",
            )
            return
        amount_text = self.edit_amount.text().strip()
        method = self.combo_pay_method.currentText().strip()
        if not amount_text or not method:
            QMessageBox.warning(
                self,
                "校验失败",
                "可支付金额 / 付款方式不能为空，请填写完整。",
            )
            return
        try:
            amount = float(amount_text.replace(",", ""))
        except ValueError:
            QMessageBox.warning(self, "校验失败", "可支付金额请输入有效数字。")
            return
        latest = find_latest_total_excel(total_sheet_dir())
        if not latest:
            QMessageBox.warning(self, "提示", "未找到排款计划总表文件。")
            return
        if not latest.lower().endswith(".xlsx"):
            QMessageBox.warning(
                self,
                "格式限制",
                "保存银行配置仅支持 .xlsx 格式的排款计划总表。\\n"
                "请将总表另存为 Excel 2007+（.xlsx）后重新上传。",
            )
            return
        try:
            write_bank_config_xlsx(latest, bank, amount, method)
        except Exception as e:
            QMessageBox.critical(self, "保存失败", str(e))
            traceback.print_exc()
            return
        QMessageBox.information(self, "成功", "银行信息保存成功。")

    def _show_smart_upload_prereq(self, missing: list[str]) -> None:
        dlg = QDialog(self)
        dlg.setObjectName("uploadPrereqDlg")
        dlg.setWindowTitle("上传提示")
        dlg.setModal(True)
        dlg.setMinimumWidth(440)
        dlg.setStyleSheet(
            f"""
            QDialog#uploadPrereqDlg {{
                background-color: #E8F4FC;
                border: 1px solid #91CAFF;
                border-radius: 8px;
            }}
            QLabel#uploadPrereqTitle {{
                color: {Figma.LABEL_PURPLE};
                font-size: 17px;
                font-weight: 600;
                background: transparent;
            }}
            QLabel#uploadPrereqBody {{
                color: {Figma.LABEL_PURPLE};
                font-size: 15px;
                background: transparent;
            }}
            QLabel#uploadPrereqIcon {{
                background: transparent;
            }}
            """
        )
        root = QVBoxLayout(dlg)
        root.setSpacing(18)
        root.setContentsMargins(22, 22, 22, 20)
        row = QHBoxLayout()
        row.setSpacing(18)
        icon_lbl = QLabel()
        icon_lbl.setObjectName("uploadPrereqIcon")
        sty = self.style() or (
            QApplication.instance().style() if QApplication.instance() else None
        )
        if sty is not None:
            pm = sty.standardIcon(
                QStyle.StandardPixmap.SP_MessageBoxInformation
            ).pixmap(52, 52)
            icon_lbl.setPixmap(pm)
        icon_lbl.setFixedSize(52, 52)
        icon_lbl.setAlignment(Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignHCenter)
        row.addWidget(icon_lbl, 0, Qt.AlignmentFlag.AlignTop)
        text_col = QVBoxLayout()
        text_col.setSpacing(10)
        title = QLabel("请先上传所需文件")
        title.setObjectName("uploadPrereqTitle")
        bullets = "\\n".join(f"• {name}" for name in missing)
        body = QLabel(
            "使用「智能排款」前，请先点击上方「上传排款计划总表」或"
            "「上传排款明细模版」按钮，完成以下文件上传：\\n\\n" + bullets
        )
        body.setObjectName("uploadPrereqBody")
        body.setWordWrap(True)
        text_col.addWidget(title)
        text_col.addWidget(body)
        row.addLayout(text_col, 1)
        root.addLayout(row)
        box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok)
        box.accepted.connect(dlg.accept)
        ok_btn = box.button(QDialogButtonBox.StandardButton.Ok)
        if ok_btn is not None:
            ok_btn.setText("知道了")
            ok_btn.setObjectName("figmaBlueSmall")
        root.addWidget(box, 0, Qt.AlignmentFlag.AlignRight)
        dlg.exec()

    def _show_smart_schedule_success_dialog(self, res: SmartScheduleResult) -> None:
        dlg = QDialog(self)
        dlg.setObjectName("smartOkDlg")
        dlg.setWindowTitle("智能排款完成")
        dlg.setModal(True)
        dlg.setMinimumWidth(460)
        dlg.setStyleSheet(
            f"""
            QDialog#smartOkDlg {{
                background-color: #E8F4FC;
                border: 1px solid #91CAFF;
                border-radius: 8px;
            }}
            QLabel#smartOkHead {{
                color: {Figma.LABEL_PURPLE};
                font-size: 22px;
                font-weight: 700;
                background: transparent;
            }}
            QLabel#smartOkBody {{
                color: {Figma.AMOUNT_TEXT};
                font-size: 15px;
                background: transparent;
            }}
            QLabel#smartOkIcon {{
                background: transparent;
            }}
            """
        )
        root = QVBoxLayout(dlg)
        root.setSpacing(14)
        root.setContentsMargins(22, 20, 22, 18)
        row = QHBoxLayout()
        row.setSpacing(16)
        icon_lbl = QLabel()
        icon_lbl.setObjectName("smartOkIcon")
        sty = self.style() or (
            QApplication.instance().style() if QApplication.instance() else None
        )
        if sty is not None:
            pm = sty.standardIcon(
                QStyle.StandardPixmap.SP_MessageBoxInformation
            ).pixmap(48, 48)
            icon_lbl.setPixmap(pm)
        icon_lbl.setFixedSize(48, 48)
        icon_lbl.setAlignment(Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignHCenter)
        row.addWidget(icon_lbl, 0, Qt.AlignmentFlag.AlignTop)
        text_col = QVBoxLayout()
        text_col.setSpacing(10)
        head = QLabel("智能排款成功！")
        head.setObjectName("smartOkHead")
        body_lines = [
            f"已处理总表：{os.path.basename(res.path)}",
            f"付款优先级 = 0 且主体非空：{res.rows_priority0} 行",
            f"付款优先级 = 1 且主体非空：{res.rows_priority1} 行",
            f"付款优先级 = 2 且主体非空：{res.rows_priority2} 行",
            f"其中已解析协议账期并参与匹配：{res.rows_scanned} 行",
            f"已写入排款并扣减余额：{res.rows_written} 行",
            f"未写入（额度/三条件/账期汇总等）：{res.rows_skipped} 行",
            "（已按优先级 0→1→2 顺序处理）",
        ]
        body = QLabel("\\n".join(body_lines))
        body.setObjectName("smartOkBody")
        body.setWordWrap(True)
        text_col.addWidget(head)
        text_col.addWidget(body)
        row.addLayout(text_col, 1)
        root.addLayout(row)
        box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok)
        box.accepted.connect(dlg.accept)
        ok_btn = box.button(QDialogButtonBox.StandardButton.Ok)
        if ok_btn is not None:
            ok_btn.setText("确定")
            ok_btn.setObjectName("figmaBlueSmall")
        root.addWidget(box, 0, Qt.AlignmentFlag.AlignRight)
        dlg.exec()

    def _on_smart_placeholder(self) -> None:
        self._reset_card1_payable_count_vars()
        ensure_workspace()
        ts_dir = total_sheet_dir()
        dt_dir = detail_template_dir()
        empty_ts = _workspace_dir_is_empty(ts_dir)
        empty_dt = _workspace_dir_is_empty(dt_dir)
        if empty_ts and empty_dt:
            QMessageBox.warning(
                self,
                "提示",
                "请点击「上传排款计划总表」按钮上传排款计划总表。\\n\\n"
                "请点击「上传排款明细模版」按钮上传排款明细模版。",
            )
            return
        if empty_ts:
            QMessageBox.warning(
                self,
                "提示",
                "请点击「上传排款计划总表」按钮上传排款计划总表。",
            )
            return
        if empty_dt:
            QMessageBox.warning(
                self,
                "提示",
                "请点击「上传排款明细模版」按钮上传排款明细模版。",
            )
            return
        p1_ok = bool(self.edit_priority1_pct.text().strip())
        p2_ok = bool(self.edit_priority2_pct.text().strip())
        if not p1_ok and not p2_ok:
            QMessageBox.warning(
                self,
                "提示",
                "请先填写「第一优先级支付计划占应付额」与「第二优先级支付计划占应付额」后再进行智能排款。",
            )
            return
        if not p1_ok:
            QMessageBox.warning(
                self,
                "提示",
                "请先填写「第一优先级支付计划占应付额」后再进行智能排款。",
            )
            return
        if not p2_ok:
            QMessageBox.warning(
                self,
                "提示",
                "请先填写「第二优先级支付计划占应付额」后再进行智能排款。",
            )
            return
        has_total = bool(find_latest_total_excel(ts_dir))
        has_detail = bool(find_latest_total_excel(dt_dir))
        missing: list[str] = []
        if not has_total:
            missing.append("排款计划总表")
        if not has_detail:
            missing.append("排款明细模版")
        if missing:
            self._show_smart_upload_prereq(missing)
            return
        total_path = find_latest_total_excel(total_sheet_dir())
        if not total_path or not total_path.lower().endswith(".xlsx"):
            QMessageBox.warning(
                self,
                "格式限制",
                "智能排款仅支持 .xlsx 格式的排款计划总表。\\n"
                "请将总表另存为 Excel 2007+（.xlsx）后重新上传。",
            )
            return
        try:
            pct1 = int(self.edit_priority1_pct.text().strip(), 10)
            pct2 = int(self.edit_priority2_pct.text().strip(), 10)
        except ValueError:
            QMessageBox.warning(self, "提示", "优先级占应付额请输入有效整数（0～100）。")
            return
        try:
            res = run_smart_schedule_on_total_sheet(total_path, pct1, pct2)
        except Exception as e:
            QMessageBox.critical(self, "智能排款失败", str(e))
            traceback.print_exc()
            return
        self._smart_schedule_completed_ok = True
        if not SHOW_BANK_SETTINGS_CARD:
            self._total_path = find_latest_total_excel(total_sheet_dir())
            self._rebuild_sheet1_strip_cache(self._total_path)
        if res.reserved_total is not None:
            self.lbl_reserved_total_value.setText(
                _format_payable_amount(res.reserved_total)
            )
        else:
            self.lbl_reserved_total_value.setText("")
        self.lbl_sched_strict.setText(
            _format_payable_amount(res.priority0_total)
        )
        self.lbl_sched_priority1.setText(
            _format_payable_amount(res.priority1_total)
        )
        self.lbl_sched_priority2.setText(
            _format_payable_amount(res.priority2_total)
        )
        self.lbl_sched_month_total.setText(
            _format_payable_amount(
                res.priority0_total + res.priority1_total + res.priority2_total
            )
        )
        detail_path = find_latest_total_excel(detail_template_dir())
        if detail_path:
            if detail_path.lower().endswith(".xlsx"):
                try:
                    fill_detail_workbook_from_total(detail_path, total_path)
                except Exception as e:
                    QMessageBox.warning(
                        self,
                        "排款明细表",
                        f"排款明细表数据生成失败：\\n{e}",
                    )
                    traceback.print_exc()
            else:
                QMessageBox.warning(
                    self,
                    "排款明细表",
                    "明细模版为 .xls 时无法自动生成排款明细表，请另存为 .xlsx 后重新上传模版。",
                )
        # 第二卡片（查看排款下方两行汇总）先归零；用户点击「查看排款」时再按总表加载。
        self._refresh_owner_cost_panels(None)
        self._show_smart_schedule_success_dialog(res)

    def _on_view_schedule_placeholder(self) -> None:
        """查看排款：校验前置条件后统计并仅刷新卡片 3（底部四行各部门明细）。"""
        ensure_workspace()
        ts = total_sheet_dir()
        if _workspace_dir_is_empty(ts):
            QMessageBox.warning(
                self,
                "查看排款",
                _VIEW_SCHEDULE_PREREQ_MSG,
            )
            self._refresh_owner_cost_panels(None)
            return
        latest = find_latest_total_excel(ts)
        if not latest:
            QMessageBox.warning(
                self,
                "查看排款",
                _VIEW_SCHEDULE_PREREQ_MSG,
            )
            self._refresh_owner_cost_panels(None)
            return
        if not latest.lower().endswith(".xlsx"):
            QMessageBox.warning(
                self,
                "查看排款",
                "查看排款仅支持 .xlsx 格式的排款计划总表。",
            )
            self._refresh_owner_cost_panels(None)
            return
        plan_no_numeric = not sheet1_payment_plan_has_numeric_entries(latest)
        if (not self._smart_schedule_completed_ok) or plan_no_numeric:
            QMessageBox.warning(
                self,
                "查看排款",
                _VIEW_SCHEDULE_PREREQ_MSG,
            )
            if os.path.isfile(latest) and latest.lower().endswith(".xlsx"):
                self._refresh_owner_bottom_placeholder_from_path(latest)
            else:
                self._refresh_owner_cost_panels(None)
            return
        self._refresh_owner_cost_panels(latest)

def _apply_ui_font(app: QApplication) -> None:
    f = QFont()
    if sys.platform == "darwin":
        f.setFamily("PingFang SC")
    elif sys.platform == "win32":
        f.setFamily("Microsoft YaHei UI")
    else:
        f.setFamily("Noto Sans CJK SC")
    f.setPointSize(Figma.BODY_PT)
    app.setFont(f)


def _suppress_macos_imk_stderr_noise() -> None:
    """macOS 上 Qt 与输入法 IMK 会刷 IMKCFRunLoopWakeUpReliable，属系统噪声，过滤以免误判为报错。"""
    if sys.platform != "darwin":
        return
    needle = "IMKCFRunLoopWakeUpReliable"
    real = sys.stderr

    class _StderrFilter:
        __slots__ = ("_real",)

        def __init__(self, underlying) -> None:
            self._real = underlying

        def write(self, s: str) -> int:
            if not s:
                return 0
            if needle not in s:
                self._real.write(s)
                return len(s)
            kept = "".join(
                ln for ln in s.splitlines(keepends=True) if needle not in ln
            )
            if kept:
                self._real.write(kept)
            return len(s)

        def flush(self) -> None:
            self._real.flush()

        def __getattr__(self, name: str):
            return getattr(self._real, name)

    sys.stderr = _StderrFilter(real)  # type: ignore[misc, assignment]


def main() -> None:
    _suppress_macos_imk_stderr_noise()
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    app.setStyleSheet(build_message_box_stylesheet())
    _apply_ui_font(app)
    try:
        ensure_workspace()
    except OSError as e:
        QMessageBox.critical(
            None,
            "初始化失败",
            f"无法创建工作目录 {get_base_dir()}：\\n{e}",
        )
        sys.exit(1)
    w = MainWindow()
    w.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()