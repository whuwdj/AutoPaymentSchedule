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
from PyQt6.QtGui import QDoubleValidator, QFont, QIcon
from PyQt6.QtWidgets import (
    QApplication,
    QComboBox,
    QFileDialog,
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
    find_latest_total_excel,
    read_banks,
    write_bank_config_xlsx,
)
from paths import detail_template_dir, ensure_workspace, get_base_dir, total_sheet_dir


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
        self._lbl_strict_val: QLabel | None = None
        self._lbl_month_val: QLabel | None = None
        self._setup_ui()
        self._apply_styles()
        self._refresh_bank_combo()

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
    def _lbl_yuan_purple() -> QLabel:
        w = QLabel("元")
        w.setObjectName("summaryPurple")
        return w

    def _setup_ui(self) -> None:
        self.setWindowTitle("智能排款")
        self.setMinimumSize(560, 600)
        self.setWindowIcon(_load_window_icon())

        m = Figma
        root = QVBoxLayout(self)
        root.setContentsMargins(m.PAGE_MARGIN_H, m.PAGE_MARGIN_V, m.PAGE_MARGIN_H, m.PAGE_MARGIN_V)
        root.setSpacing(m.GAP_SECTION)

        # —— 顶部：上传（稿中为蓝色实心按钮）——
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
        root.addLayout(row_upload)

        # —— 全局：可支付总额、百分比（界面与稿一致；持久化逻辑可后续接表）——
        row_total = QHBoxLayout()
        row_total.setSpacing(m.GAP_TIGHT)
        row_total.addWidget(self._lbl_purple("可支付总额："))
        self.edit_global_total = QLineEdit()
        self.edit_global_total.setObjectName("figmaInput")
        self.edit_global_total.setPlaceholderText("")
        self.edit_global_total.setValidator(QDoubleValidator(0.0, 1e15, 2))
        self.edit_global_total.setMinimumHeight(40)
        self.edit_global_total.setMinimumWidth(50)
        self.edit_global_total.setMaximumWidth(m.W_INPUT_SHORT_MAX)
        self.edit_global_total.setSizePolicy(
            QSizePolicy.Policy.Maximum, QSizePolicy.Policy.Fixed
        )
        row_total.addWidget(self.edit_global_total, 0)
        row_total.addWidget(self._lbl_yuan_blue())
        row_total.addStretch(1)
        root.addLayout(row_total)

        row_pct = QHBoxLayout()
        row_pct.setSpacing(m.GAP_TIGHT)
        row_pct.addWidget(self._lbl_purple("支付计划占应付总额百分比："))
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
        row_pct.addWidget(self.edit_global_percent, 0)
        row_pct.addWidget(self._lbl_purple("%"))
        row_pct.addStretch(1)
        root.addLayout(row_pct)

        # —— 银行模块小标题 ——
        micro = QLabel("银行可支付金额及付款方式设置")
        micro.setObjectName("figmaMicroTitle")
        root.addWidget(micro)

        # —— 银行选择 + 付款方式（同一行）——
        row_bank_pay = QHBoxLayout()
        row_bank_pay.setSpacing(m.GAP_INLINE)
        row_bank_pay.addWidget(self._lbl_purple("银 行 选 择："))
        self.combo_bank = QComboBox()
        self.combo_bank.setObjectName("figmaCombo")
        self.combo_bank.setMinimumHeight(40)
        self.combo_bank.setMinimumWidth(200)
        self.combo_bank.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed
        )
        row_bank_pay.addWidget(self.combo_bank, 1)
        row_bank_pay.addWidget(self._lbl_purple("付款方式："))
        self.combo_pay_method = QComboBox()
        self.combo_pay_method.setObjectName("figmaCombo")
        self.combo_pay_method.setMinimumHeight(40)
        self.combo_pay_method.setMinimumWidth(180)
        for s in ("不限", "电汇", "支票", "承兑汇票", "现金"):
            self.combo_pay_method.addItem(s)
        self.combo_pay_method.setCurrentIndex(0)
        self.combo_pay_method.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed
        )
        row_bank_pay.addWidget(self.combo_pay_method, 1)
        row_bank_pay.addStretch(1)
        root.addLayout(row_bank_pay)

        # —— 可支付金额 + 元 + 确定 ——
        row_amt = QHBoxLayout()
        row_amt.setSpacing(m.GAP_TIGHT)
        row_amt.addWidget(self._lbl_purple("可支付金额："))
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
        row_amt.addWidget(self.edit_amount, 0)
        row_amt.addWidget(self._lbl_yuan_blue())
        row_amt.addStretch(1)
        self.btn_confirm = QPushButton("确定")
        self.btn_confirm.setObjectName("figmaBlueSmall")
        self.btn_confirm.setMinimumHeight(44)
        self.btn_confirm.clicked.connect(self._on_confirm)
        row_amt.addWidget(self.btn_confirm, 0, Qt.AlignmentFlag.AlignRight)
        root.addLayout(row_amt)

        # —— 底部：左侧「智能排款」，右侧两行汇总 ——
        bottom = QHBoxLayout()
        bottom.setSpacing(28)
        bottom.setAlignment(Qt.AlignmentFlag.AlignTop)

        self.btn_smart = QPushButton("智能排款")
        self.btn_smart.setObjectName("figmaHero")
        self.btn_smart.setMinimumWidth(220)
        self.btn_smart.setMinimumHeight(88)
        self.btn_smart.setSizePolicy(
            QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed
        )
        self.btn_smart.clicked.connect(self._on_smart_placeholder)
        bottom.addWidget(self.btn_smart, 0, Qt.AlignmentFlag.AlignTop)

        col_sum = QVBoxLayout()
        col_sum.setSpacing(14)

        r1 = QHBoxLayout()
        r1.addWidget(self._lbl_purple("严格按账期支付金额："))
        r1.addStretch(1)
        self._lbl_strict_val = QLabel("—")
        self._lbl_strict_val.setObjectName("summaryValue")
        self._lbl_strict_val.setAlignment(
            Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter
        )
        r1.addWidget(self._lbl_strict_val)
        r1.addWidget(self._lbl_yuan_purple())
        col_sum.addLayout(r1)

        r2 = QHBoxLayout()
        r2.addWidget(self._lbl_purple("本月应付款金额："))
        r2.addStretch(1)
        self._lbl_month_val = QLabel("—")
        self._lbl_month_val.setObjectName("summaryValue")
        self._lbl_month_val.setAlignment(
            Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter
        )
        r2.addWidget(self._lbl_month_val)
        r2.addWidget(self._lbl_yuan_purple())
        col_sum.addLayout(r2)

        col_sum.addStretch(1)
        bottom.addLayout(col_sum, 1)
        root.addLayout(bottom)

        root.addStretch(1)

        hint = QLabel(f"工作目录：{get_base_dir()}")
        hint.setObjectName("footerHint")
        hint.setWordWrap(True)
        root.addWidget(hint)

    def _apply_styles(self) -> None:
        self.setStyleSheet(build_app_stylesheet())

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
        self._refresh_bank_combo()

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

    def _refresh_bank_combo(self) -> None:
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

    def _current_bank(self) -> BankOption | None:
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
            "智能排款功能将在后续版本中提供，当前为预留入口。\n"
            "届时将在此更新「严格按账期支付金额」「本月应付款金额」等汇总。",
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
