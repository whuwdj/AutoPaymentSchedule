# -*- coding: utf-8 -*-
"""
智能排款：上传总表/明细模板、从 Sheet2 读取银行、写入可支付金额与付款方式。
Windows 工作区：D:\\AutoPaymentScheduleFile

界面布局与色值对齐 Figma「智能排款界面」：
https://www.figma.com/design/aabHTb8OZnSlJBQPKrNhrZ/%E6%99%BA%E8%83%BD%E6%8E%92%E6%AC%BE%E7%95%8C%E9%9D%A2?node-id=0-1
"""

from __future__ import annotations

import os
import shutil
import sys
import traceback

from PyQt6.QtCore import Qt
from PyQt6.QtGui import QColor, QDoubleValidator, QFont, QIcon
from PyQt6.QtWidgets import (
    QApplication,
    QComboBox,
    QFileDialog,
    QFrame,
    QGraphicsDropShadowEffect,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QSizePolicy,
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
    write_bank_config_xlsx,
)
from paths import detail_template_dir, ensure_workspace, get_base_dir, total_sheet_dir

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
    def _right_align_cell(width_px: int, widget: QWidget) -> QWidget:
        cell = QWidget()
        h = QHBoxLayout(cell)
        h.setContentsMargins(0, 0, 0, 0)
        h.addStretch(1)
        h.addWidget(widget)
        cell.setFixedWidth(width_px)
        return cell

    @staticmethod
    def _apply_card_shadow(card: QFrame) -> None:
        eff = QGraphicsDropShadowEffect(card)
        eff.setBlurRadius(26)
        eff.setOffset(0, 5)
        eff.setColor(QColor(15, 23, 42, 52))
        card.setGraphicsEffect(eff)

    def _setup_ui(self) -> None:
        self.setWindowTitle("智能排款")
        self.setMinimumSize(1120, 600)
        self.setWindowIcon(_load_window_icon())

        m = Figma
        root = QVBoxLayout(self)
        root.setContentsMargins(m.PAGE_MARGIN_H, m.PAGE_MARGIN_V, m.PAGE_MARGIN_H, m.PAGE_MARGIN_V)
        root.setSpacing(m.GAP_SECTION)

        # —— 卡片 1：上传 + 全局总额 / 百分比 ——
        card_top = QFrame()
        card_top.setObjectName("elevatedCard")
        card_top.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)
        lay_top = QVBoxLayout(card_top)
        lay_top.setContentsMargins(m.CARD_PAD_H, m.CARD_PAD_V, m.CARD_PAD_H, m.CARD_PAD_V)
        lay_top.setSpacing(m.GAP_SECTION)

        row_upload = QHBoxLayout()
        row_upload.setSpacing(m.GAP_INLINE)
        self.btn_upload_total = QPushButton("上传排款计划总表")
        self.btn_upload_total.setObjectName("figmaBlue")
        self.btn_upload_total.setMinimumHeight(40)
        self.btn_upload_total.clicked.connect(self._on_upload_total)
        self.btn_upload_template = QPushButton("上传排款明细模版")
        self.btn_upload_template.setObjectName("figmaBlue")
        self.btn_upload_template.setMinimumHeight(40)
        self.btn_upload_template.clicked.connect(self._on_upload_template)
        row_upload.addWidget(self.btn_upload_total)
        row_upload.addWidget(self.btn_upload_template)
        row_upload.addStretch(1)
        lay_top.addLayout(row_upload)

        lbl_sheet_total = self._lbl_purple("本月可支付总额为：")
        lbl_sheet_pct = self._lbl_purple("支付计划占应付款：")
        w_label_col = max(
            lbl_sheet_total.sizeHint().width(),
            lbl_sheet_pct.sizeHint().width(),
        )

        summary_block = QWidget()
        grid_sum = QGridLayout(summary_block)
        grid_sum.setContentsMargins(0, 0, 0, 0)
        grid_sum.setHorizontalSpacing(m.GAP_TIGHT)
        grid_sum.setVerticalSpacing(m.GAP_SECTION)

        ra = Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter
        va = Qt.AlignmentFlag.AlignVCenter

        grid_sum.setColumnMinimumWidth(0, w_label_col)
        grid_sum.setColumnMinimumWidth(1, m.W_INPUT_SHORT_MAX)

        cell_first_yuan = QWidget()
        h_first_yuan = QHBoxLayout(cell_first_yuan)
        h_first_yuan.setContentsMargins(0, 0, 0, 0)
        h_first_yuan.addStretch(1)
        h_first_yuan.addWidget(self._lbl_yuan_blue())
        h_first_yuan.addStretch(1)
        cell_first_yuan.setFixedWidth(m.W_INPUT_SHORT_MAX)

        grid_sum.addWidget(lbl_sheet_total, 0, 0, ra)
        grid_sum.addWidget(cell_first_yuan, 0, 1, va)
        grid_sum.addWidget(self._lbl_purple("本月预留款总额："), 0, 2, va)
        spacer_reserved = QWidget()
        spacer_reserved.setFixedSize(m.W_INPUT_SHORT_MAX, 40)
        grid_sum.addWidget(spacer_reserved, 0, 3, va)
        grid_sum.addWidget(self._lbl_yuan_blue(), 0, 4, va)

        self.edit_global_percent = QLineEdit()
        self.edit_global_percent.setObjectName("figmaInput")
        self.edit_global_percent.setPlaceholderText("")
        v_pct = QDoubleValidator(0.0, 100.0, 2)
        self.edit_global_percent.setValidator(v_pct)
        self.edit_global_percent.setMinimumHeight(40)
        self.edit_global_percent.setMinimumWidth(50)
        self.edit_global_percent.setMaximumWidth(m.W_INPUT_SHORT_MAX)
        self.edit_global_percent.setSizePolicy(
            QSizePolicy.Policy.Maximum, QSizePolicy.Policy.Fixed
        )
        grid_sum.addWidget(lbl_sheet_pct, 1, 0, ra)
        grid_sum.addWidget(self.edit_global_percent, 1, 1, va)
        grid_sum.addWidget(self._lbl_purple("%"), 1, 2, va)
        grid_sum.setColumnStretch(5, 1)

        lay_top.addWidget(summary_block)

        self._apply_card_shadow(card_top)
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
            self.edit_amount.setMaximumWidth(m.W_INPUT_SHORT_MAX)
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
            self._apply_card_shadow(card_bank)
            root.addWidget(card_bank)

        # —— 卡片 2（隐藏银行模块时）/ 卡片 3：「智能排款」主操作（随窗口增高留白在卡片内）——
        card_bottom = QFrame()
        card_bottom.setObjectName("elevatedCard")
        card_bottom.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)
        lay_bottom = QVBoxLayout(card_bottom)
        lay_bottom.setContentsMargins(m.CARD_PAD_H, m.CARD_PAD_V, m.CARD_PAD_H, m.CARD_PAD_V)
        lay_bottom.setSpacing(0)

        bottom = QHBoxLayout()
        bottom.setSpacing(28)
        bottom.setAlignment(Qt.AlignmentFlag.AlignVCenter)

        self.btn_smart = QPushButton("智能排款")
        self.btn_smart.setObjectName("figmaHero")
        self.btn_smart.setMinimumWidth(220)
        self.btn_smart.setMinimumHeight(88)
        self.btn_smart.setSizePolicy(
            QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed
        )
        self.btn_smart.clicked.connect(self._on_smart_placeholder)
        bottom.addWidget(self.btn_smart, 0, Qt.AlignmentFlag.AlignVCenter)

        smart_labels = QWidget()
        v_smart_lbl = QVBoxLayout(smart_labels)
        v_smart_lbl.setContentsMargins(0, 0, 0, 0)
        v_smart_lbl.setSpacing(10)
        v_smart_lbl.addWidget(self._lbl_purple("严格按账期支付金额："))
        v_smart_lbl.addWidget(self._lbl_purple("本月应付款金额："))
        bottom.addWidget(smart_labels, 0, Qt.AlignmentFlag.AlignVCenter)
        bottom.addStretch(1)
        lay_bottom.addLayout(bottom)
        lay_bottom.addStretch(1)

        self._apply_card_shadow(card_bottom)
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

    def _on_upload_total(self) -> None:
        ensure_workspace()
        path, _ = QFileDialog.getOpenFileName(
            self, "选择排款计划总表", "", EXCEL_FILTER
        )
        if not path:
            return
        dest = os.path.join(total_sheet_dir(), os.path.basename(path))
        try:
            shutil.copy2(path, dest)
        except OSError as e:
            QMessageBox.critical(self, "上传失败", str(e))
            return
        self._total_path = dest
        QMessageBox.information(self, "上传成功", f"已保存到：\n{dest}")
        if SHOW_BANK_SETTINGS_CARD:
            self._refresh_bank_combo()
        else:
            self._rebuild_sheet1_strip_cache(dest)

    def _on_upload_template(self) -> None:
        ensure_workspace()
        path, _ = QFileDialog.getOpenFileName(
            self, "选择排款明细模板", "", EXCEL_FILTER
        )
        if not path:
            return
        dest = os.path.join(detail_template_dir(), os.path.basename(path))
        try:
            shutil.copy2(path, dest)
        except OSError as e:
            QMessageBox.critical(self, "上传失败", str(e))
            return
        QMessageBox.information(self, "上传成功", f"已保存到：\n{dest}")

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

    def _on_smart_placeholder(self) -> None:
        QMessageBox.information(
            self,
            "智能排款",
            "智能排款功能将在后续版本中提供，当前为预留入口。",
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
