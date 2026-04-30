# -*- coding: utf-8 -*-
"""
智能排款：上传总表/明细模板、从 Sheet2 读取银行、写入可支付金额与付款方式。
工作区：程序目录\\AutoPaymentScheduleFile（TotalSheet / DetailTemplate）。

界面布局与色值对齐 Figma「智能排款界面」：
https://www.figma.com/design/aabHTb8OZnSlJBQPKrNhrZ/%E6%99%BA%E8%83%BD%E6%8E%92%E6%AC%BE%E7%95%8C%E9%9D%A2?node-id=0-1
"""

from __future__ import annotations

import os
import shutil
import sys
import traceback

from PyQt6.QtCore import Qt
from PyQt6.QtGui import QDoubleValidator, QFont, QIcon
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
    QSizePolicy,
    QStyle,
    QVBoxLayout,
    QWidget,
)

from design_tokens import Figma, build_app_stylesheet
from excel_service import (
    EXCEL_FILTER,
    BankOption,
    SHEET1_SUMMARY_LABELS,
    find_latest_total_excel,
    read_banks,
    read_sheet1_summary_strip,
    sum_sheet1_row4_after_payout_gross_column,
    write_bank_config_xlsx,
)
from paths import (
    capture_total_sheet_startup_snapshot,
    detail_template_dir,
    ensure_workspace,
    get_base_dir,
    restore_total_sheet_from_startup_snapshot,
    total_sheet_dir,
)
from smart_schedule import SmartScheduleResult, run_smart_schedule_on_total_sheet

_EXCEL_SUFFIXES = (".xlsx", ".xls")


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
    """展示用：固定两位小数（整数亦显示 .00）。"""
    return f"{n:.2f}"


# 付款方式下拉选项（启动时清空后重新加载）
PAY_METHOD_CHOICES = ("不限", "电汇", "支票", "承兑汇票", "现金")

# 为 False 时不创建「银行可支付金额及付款方式」卡片，智能排款卡片上移为第二块
SHOW_BANK_SETTINGS_CARD = False


def _app_dir() -> str:
    return os.path.dirname(os.path.abspath(__file__))


def _load_window_icon() -> QIcon:
    ico = os.path.join(_app_dir(), "app.ico")
    if os.path.isfile(ico):
        return QIcon(ico)
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
        self._setup_ui()
        self._apply_styles()
        if SHOW_BANK_SETTINGS_CARD:
            self._refresh_pay_method_combo()
            self._refresh_bank_combo()
        else:
            ensure_workspace()
            self._total_path = find_latest_total_excel(total_sheet_dir())
            self._rebuild_sheet1_strip_cache(self._total_path)
        self._clear_startup_numeric_fields()
        ensure_workspace()
        capture_total_sheet_startup_snapshot()

    def _clear_startup_numeric_fields(self) -> None:
        self.lbl_payable_total_value.setText("")
        self.edit_priority1_pct.clear()
        self.edit_priority2_pct.clear()
        if SHOW_BANK_SETTINGS_CARD:
            self.edit_amount.clear()

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
        return w

    def _data_cell(self, label_text: str) -> QWidget:
        m = Figma
        w = QWidget()
        h = QHBoxLayout(w)
        h.setContentsMargins(0, 0, 0, 0)
        h.setSpacing(0)
        h.addWidget(self._lbl_purple(label_text))
        h.addSpacing(m.DATA_LABEL_TO_YUAN)
        h.addWidget(self._suffix_unit("元"))
        h.addStretch(0)
        return w

    def _summary_data_row(self) -> QWidget:
        m = Figma
        w = QWidget()
        h = QHBoxLayout(w)
        h.setContentsMargins(0, 0, 0, 0)
        h.setSpacing(m.DATA_ROW_GAP)
        h.addWidget(self._data_cell("成本二部："), 1)
        h.addWidget(self._data_cell("成本三部："), 1)
        h.addWidget(self._data_cell("合计："), 1)
        return w

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
        self.setMinimumSize(1120, 600)
        self.setWindowIcon(_load_window_icon())

        m = Figma
        root = QVBoxLayout(self)
        root.setContentsMargins(m.PAGE_MARGIN_H, m.PAGE_MARGIN_V, m.PAGE_MARGIN_H, m.PAGE_MARGIN_V)
        root.setSpacing(m.GAP_SECTION)

        # —— 卡片 1：与 figma-smart-payment-ui.html 一致 ——
        card_top = QFrame()
        card_top.setObjectName("elevatedCard")
        card_top.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)
        lay_top = QVBoxLayout(card_top)
        lay_top.setContentsMargins(m.CARD_PAD_H, m.CARD_PAD_V, m.CARD_PAD_H, m.CARD_PAD_V)
        lay_top.setSpacing(m.ROW_GAP)

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
        lay_top.addLayout(row_upload)

        def _yuan_tail_cell() -> QWidget:
            cell = QWidget()
            cell.setFixedWidth(m.W_PCT_TAIL)
            hh = QHBoxLayout(cell)
            hh.setContentsMargins(0, 0, 0, 0)
            hh.setSpacing(0)
            hh.addSpacing(m.W_PCT_INPUT // 2)
            hh.addWidget(self._suffix_unit("元"))
            hh.addStretch(1)
            return cell

        def _pct_tail_widget(edit: QLineEdit) -> QWidget:
            cell = QWidget()
            cell.setFixedWidth(m.W_PCT_TAIL)
            hh = QHBoxLayout(cell)
            hh.setContentsMargins(0, 0, 0, 0)
            hh.setSpacing(m.GAP_TIGHT)
            hh.addWidget(edit)
            hh.addWidget(self._suffix_unit("%"))
            hh.addStretch(0)
            return cell

        form_grid = QWidget()
        g = QGridLayout(form_grid)
        g.setContentsMargins(0, 0, 0, 0)
        g.setHorizontalSpacing(m.GAP_TIGHT)
        g.setVerticalSpacing(m.ROW_GAP)
        g.setColumnMinimumWidth(1, m.W_PCT_TAIL)
        g.setColumnStretch(2, 1)
        va = Qt.AlignmentFlag.AlignVCenter

        row0_label_value = QWidget()
        rv = QHBoxLayout(row0_label_value)
        rv.setContentsMargins(0, 0, 0, 0)
        rv.setSpacing(0)
        rv.addWidget(self._lbl_purple("本月可支付总额："))
        self.lbl_payable_total_value = QLabel("")
        self.lbl_payable_total_value.setObjectName("summaryValue")
        self.lbl_payable_total_value.setAlignment(
            Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter
        )
        rv.addWidget(self.lbl_payable_total_value)

        g.addWidget(row0_label_value, 0, 0, va)
        g.addWidget(_yuan_tail_cell(), 0, 1, va)

        g.addWidget(self._lbl_purple("本月预留款总额："), 1, 0, va)
        g.addWidget(_yuan_tail_cell(), 1, 1, va)

        self.edit_priority1_pct = QLineEdit()
        self.edit_priority1_pct.setObjectName("figmaInput")
        self.edit_priority1_pct.setValidator(QDoubleValidator(0.0, 100.0, 2))
        self.edit_priority1_pct.setFixedWidth(m.W_PCT_INPUT)
        self.edit_priority1_pct.setSizePolicy(
            QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed
        )
        g.addWidget(self._lbl_purple("第一优先级支付计划占应付额："), 2, 0, va)
        g.addWidget(_pct_tail_widget(self.edit_priority1_pct), 2, 1, va)

        self.edit_priority2_pct = QLineEdit()
        self.edit_priority2_pct.setObjectName("figmaInput")
        self.edit_priority2_pct.setValidator(QDoubleValidator(0.0, 100.0, 2))
        self.edit_priority2_pct.setFixedWidth(m.W_PCT_INPUT)
        self.edit_priority2_pct.setSizePolicy(
            QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed
        )
        self.btn_smart = QPushButton("智能排款")
        self.btn_smart.setObjectName("figmaActionBtn")
        self.btn_smart.clicked.connect(self._on_smart_placeholder)
        g.addWidget(self._lbl_purple("第二优先级支付计划占应付额："), 3, 0, va)
        g.addWidget(_pct_tail_widget(self.edit_priority2_pct), 3, 1, va)
        g.addWidget(self.btn_smart, 3, 2, va | Qt.AlignmentFlag.AlignRight)

        lay_top.addWidget(form_grid)

        root.addWidget(card_top)

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
            self.edit_amount.setObjectName("figmaInput")
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
            root.addWidget(card_bank)

        # —— 卡片 2：查看排款 + 数据区（与 figma-smart-payment-ui.html 一致）——
        card_bottom = QFrame()
        card_bottom.setObjectName("elevatedCard")
        card_bottom.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)
        lay_bottom = QVBoxLayout(card_bottom)
        lay_bottom.setContentsMargins(m.CARD_PAD_H, m.CARD_PAD_V, m.CARD_PAD_H, m.CARD_PAD_V)
        lay_bottom.setSpacing(0)

        self.btn_view_schedule = QPushButton("查看排款")
        self.btn_view_schedule.setObjectName("figmaActionBtn")
        self.btn_view_schedule.clicked.connect(self._on_view_schedule_placeholder)
        lay_bottom.addWidget(self.btn_view_schedule, 0, Qt.AlignmentFlag.AlignLeft)

        lay_bottom.addSpacing(m.ROW_GAP)
        lay_bottom.addWidget(self._lbl_section_title("严格按账期付款情况"))
        lay_bottom.addSpacing(12)
        lay_bottom.addWidget(self._summary_data_row())
        lay_bottom.addSpacing(m.ROW_GAP)
        lay_bottom.addWidget(self._lbl_section_title("本月应付款总体情况"))
        lay_bottom.addSpacing(12)
        lay_bottom.addWidget(self._summary_data_row())
        lay_bottom.addStretch(1)

        root.addWidget(card_bottom, 1)

        hint = QLabel(f"工作目录：{get_base_dir()}")
        hint.setObjectName("footerHint")
        hint.setWordWrap(True)
        root.addWidget(hint)

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

    def _apply_payable_total_from_sheet(self, path: str | None) -> None:
        """Sheet1 第 4 行「排款总额」右侧数字格求和 → 本月可支付总额。"""
        if not path or not os.path.isfile(path):
            return
        try:
            total = sum_sheet1_row4_after_payout_gross_column(path)
        except Exception:
            traceback.print_exc()
            return
        if total is None:
            self.lbl_payable_total_value.setText("")
            return
        self.lbl_payable_total_value.setText(_format_payable_amount(total))

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
        self._total_path = dest
        msg = f"已保存到：\n{dest}"
        if same_file:
            msg = f"所选文件已在目标目录中：\n{dest}"
        elif replaced:
            msg += "\n（已覆盖同名文件。）"
        QMessageBox.information(self, "上传成功", msg)
        if SHOW_BANK_SETTINGS_CARD:
            self._refresh_bank_combo()
        else:
            self._rebuild_sheet1_strip_cache(dest)
        self._apply_payable_total_from_sheet(dest)

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
        msg = f"已保存到：\n{dest}"
        if same_file:
            msg = f"所选文件已在目标目录中：\n{dest}"
        elif replaced:
            msg += "\n（已覆盖同名文件。）"
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
                "保存银行配置仅支持 .xlsx 格式的排款计划总表。\n"
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
            """
            QDialog#uploadPrereqDlg {
                background-color: #E8F4FC;
                border: 1px solid #91CAFF;
                border-radius: 8px;
            }
            QLabel#uploadPrereqTitle {
                color: #0958D9;
                font-size: 17px;
                font-weight: 600;
                background: transparent;
            }
            QLabel#uploadPrereqBody {
                color: #434343;
                font-size: 15px;
                background: transparent;
            }
            QLabel#uploadPrereqIcon {
                background: transparent;
            }
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
        bullets = "\n".join(f"• {name}" for name in missing)
        body = QLabel(
            "使用「智能排款」前，请先点击上方「上传排款计划总表」或"
            "「上传排款明细模版」按钮，完成以下文件上传：\n\n" + bullets
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
            """
            QDialog#smartOkDlg {
                background-color: #E8F4FC;
                border: 1px solid #91CAFF;
                border-radius: 8px;
            }
            QLabel#smartOkHead {
                color: #0958D9;
                font-size: 22px;
                font-weight: 700;
                background: transparent;
            }
            QLabel#smartOkBody {
                color: #434343;
                font-size: 15px;
                background: transparent;
            }
            QLabel#smartOkIcon {
                background: transparent;
            }
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
            f"其中已解析协议账期并参与匹配：{res.rows_scanned} 行",
            f"已写入排款并扣减余额：{res.rows_written} 行",
            f"未写入（额度/三条件/账期汇总等）：{res.rows_skipped} 行",
            "（付款优先级 1、2 暂未处理）",
        ]
        body = QLabel("\n".join(body_lines))
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
        ensure_workspace()
        ts_dir = total_sheet_dir()
        dt_dir = detail_template_dir()
        empty_ts = _workspace_dir_is_empty(ts_dir)
        empty_dt = _workspace_dir_is_empty(dt_dir)
        if empty_ts and empty_dt:
            QMessageBox.warning(
                self,
                "提示",
                "请点击「上传排款计划总表」按钮上传排款计划总表。\n\n"
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
                "智能排款仅支持 .xlsx 格式的排款计划总表。\n"
                "请将总表另存为 Excel 2007+（.xlsx）后重新上传。",
            )
            return
        restore_total_sheet_from_startup_snapshot()
        total_path = find_latest_total_excel(total_sheet_dir())
        if not total_path or not total_path.lower().endswith(".xlsx"):
            QMessageBox.warning(
                self,
                "无法排款",
                "恢复为程序启动时的总表状态后，未找到可用的 .xlsx 总表文件。\n"
                "请重新上传排款计划总表后再试。",
            )
            return
        try:
            res = run_smart_schedule_on_total_sheet(total_path)
        except Exception as e:
            QMessageBox.critical(self, "智能排款失败", str(e))
            traceback.print_exc()
            return
        if not SHOW_BANK_SETTINGS_CARD:
            self._total_path = find_latest_total_excel(total_sheet_dir())
            self._rebuild_sheet1_strip_cache(self._total_path)
        self._show_smart_schedule_success_dialog(res)

    def _on_view_schedule_placeholder(self) -> None:
        QMessageBox.information(
            self,
            "查看排款",
            "查看排款功能将在后续版本中提供，当前为预留入口。",
        )


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


def main() -> None:
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    _apply_ui_font(app)
    try:
        ensure_workspace()
    except OSError as e:
        QMessageBox.critical(
            None,
            "初始化失败",
            f"无法创建工作目录 {get_base_dir()}：\n{e}",
        )
        sys.exit(1)
    w = MainWindow()
    w.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
