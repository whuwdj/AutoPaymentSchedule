# -*- coding: utf-8 -*-
"""
设计令牌与 QSS：与项目内 figma-smart-payment-ui.html 对齐。
"""

from __future__ import annotations

FIGMA_DESIGN_URL = (
    "https://www.figma.com/design/aabHTb8OZnSlJBQPKrNhrZ/"
    "%E6%99%BA%E8%83%BD%E6%8E%92%E6%AC%BE%E7%95%8C%E9%9D%A2?node-id=0-1"
)


class Figma:
    BG_CANVAS = "#FFFFFF"
    BG_CARD = "#FFFFFF"
    BG_PAGE = "#FFFFFF"

    CARD_BORDER = "#E8E8E8"
    CARD_RADIUS = 6
    CARD_PAD_H = 12
    CARD_PAD_V = 12

    LABEL_PURPLE = "#722ED1"
    # 金额数字及相邻「元」等资金展示（与标签紫色区分）
    AMOUNT_TEXT = "#0052CC"
    SECTION_MICRO = "#9254DE"
    TEXT_MUTED = "#8C8C8C"

    SUFFIX_UNIT = "#1890FF"

    BORDER_INPUT = "#1890FF"
    BORDER_INPUT_FOCUS = "#40A9FF"
    BG_INPUT = "#FFFFFF"
    BG_COMBO = "#F5F5F5"

    PRIMARY = "#1677FF"
    # 主按钮轻微「凸起」：上浅下深 + 略深外轮廓（hover / pressed 略变）
    BTN_PRIMARY_FACE_TOP = "#4A9DFF"
    BTN_PRIMARY_FACE_BOTTOM = "#1268E6"
    BTN_PRIMARY_BORDER = "#0B52B8"
    BTN_PRIMARY_FACE_TOP_HOVER = "#5AA8FF"
    BTN_PRIMARY_FACE_BOTTOM_HOVER = "#1470ED"
    BTN_PRIMARY_BORDER_HOVER = "#0C58C4"
    BTN_PRIMARY_FACE_TOP_PRESSED = "#2C82E8"
    BTN_PRIMARY_FACE_BOTTOM_PRESSED = "#0E56C4"
    BTN_PRIMARY_BORDER_PRESSED = "#084494"

    RADIUS_INPUT = 4
    RADIUS_BUTTON = 4

    PAGE_MARGIN_H = 14
    PAGE_MARGIN_V = 14
    GAP_SECTION = 16
    GAP_INLINE = 10
    GAP_TIGHT = 6

    ROW_GAP = 12
    DATA_ROW_GAP = 48
    DATA_LABEL_TO_YUAN = 120

    BODY_PT = 12
    UPLOAD_BTN_PT = BODY_PT + 1
    LABEL_PT = 13
    ACTION_BTN_PT = 14
    MICRO_PT = 10

    LABEL_PURPLE_PX = LABEL_PT
    SECTION_TITLE_PT = LABEL_PT

    W_PCT_INPUT = 80
    # 百分比行「输入框 + 间距 + %」占位宽度，总额行「元」同列右对齐
    W_PCT_SUFFIX_SLOT = 16
    W_PCT_TAIL = W_PCT_INPUT + GAP_TIGHT + W_PCT_SUFFIX_SLOT
    W_INPUT_SHORT_MAX = 100
    W_AMOUNT_INPUT = 200

    W_COMBO_BANK = 400
    W_COMBO_PAY_METHOD = 360


def build_message_box_stylesheet() -> str:
    """独立于主窗口 QSS，挂到 QApplication，避免 Win11 深色系统下 QMessageBox 深底 + 深字不可读。"""
    t = Figma
    return f"""
    QMessageBox {{
        background-color: {t.BG_CARD};
        color: {t.LABEL_PURPLE};
    }}
    QMessageBox QLabel {{
        color: {t.LABEL_PURPLE};
        background-color: {t.BG_CARD};
    }}
    QMessageBox QPushButton {{
        color: {t.LABEL_PURPLE};
        background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
            stop:0 #FAFAFA, stop:1 #E8E8E8);
        border: 1px solid #CFCFCF;
        border-radius: {t.RADIUS_BUTTON}px;
        padding: 6px 18px;
        min-width: 72px;
        min-height: 24px;
    }}
    QMessageBox QPushButton:hover {{
        background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
            stop:0 #FCFCFC, stop:1 #EDEDED);
        border: 1px solid #C0C0C0;
    }}
    QMessageBox QPushButton:pressed {{
        background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
            stop:0 #E4E4E4, stop:1 #DADADA);
        border: 1px solid #B0B0B0;
    }}
    """


def build_app_stylesheet() -> str:
    t = Figma
    _g = (
        f"qlineargradient(x1:0, y1:0, x2:0, y2:1, "
        f"stop:0 {t.BTN_PRIMARY_FACE_TOP}, stop:1 {t.BTN_PRIMARY_FACE_BOTTOM})"
    )
    _gh = (
        f"qlineargradient(x1:0, y1:0, x2:0, y2:1, "
        f"stop:0 {t.BTN_PRIMARY_FACE_TOP_HOVER}, stop:1 {t.BTN_PRIMARY_FACE_BOTTOM_HOVER})"
    )
    _gp = (
        f"qlineargradient(x1:0, y1:0, x2:0, y2:1, "
        f"stop:0 {t.BTN_PRIMARY_FACE_TOP_PRESSED}, stop:1 {t.BTN_PRIMARY_FACE_BOTTOM_PRESSED})"
    )
    _ids_flat = (
        "figmaUploadBtn",
        "figmaActionBtn",
        "figmaBlue",
        "figmaBlueSmall",
    )
    _sel = ", ".join(f"QPushButton#{i}" for i in _ids_flat)
    _sel_h = ", ".join(f"QPushButton#{i}:hover" for i in _ids_flat)
    _sel_p = ", ".join(f"QPushButton#{i}:pressed" for i in _ids_flat)
    _sel_d = ", ".join(f"QPushButton#{i}:disabled" for i in _ids_flat)
    return f"""
    QWidget {{
        color: {t.LABEL_PURPLE};
        font-size: {t.BODY_PT}px;
    }}
    QWidget#centralRoot {{
        background-color: {t.BG_CANVAS};
    }}

    QFrame#elevatedCard {{
        background-color: {t.BG_CARD};
        border: 1px solid {t.CARD_BORDER};
        border-radius: {t.CARD_RADIUS}px;
    }}

    QLabel#figmaPurpleLabel {{
        color: {t.LABEL_PURPLE};
        font-size: {t.LABEL_PURPLE_PX}px;
    }}
    QLabel#figmaSectionTitle {{
        color: {t.LABEL_PURPLE};
        font-size: {t.SECTION_TITLE_PT}px;
    }}
    QLabel#figmaMicroTitle {{
        color: {t.LABEL_PURPLE};
        font-size: {t.MICRO_PT * 1.5}px;
        font-weight: 500;
    }}
    QLabel#suffixYuan {{
        color: {t.AMOUNT_TEXT};
        font-size: {t.LABEL_PT}px;
    }}
    QLabel#suffixUnit {{
        color: {t.LABEL_PURPLE};
        font-size: {t.LABEL_PT}px;
    }}
    QLabel#summaryPurple {{
        color: {t.LABEL_PURPLE};
        font-size: {t.BODY_PT}px;
    }}
    QLabel#summaryValue {{
        color: {t.AMOUNT_TEXT};
        font-size: {t.BODY_PT + 1}px;
        font-weight: 700;
        min-width: 2em;
    }}
    QLabel#footerHint {{
        color: {t.LABEL_PURPLE};
        font-size: {t.MICRO_PT}px;
    }}

    QLineEdit#figmaInput {{
        background: {t.BG_INPUT};
        border: 1px solid {t.BORDER_INPUT};
        border-radius: {t.RADIUS_INPUT}px;
        padding: 0 6px;
        min-height: 30px;
        max-height: 32px;
        color: {t.LABEL_PURPLE};
        font-size: {t.BODY_PT}px;
    }}
    QLineEdit#figmaInput:focus {{
        border: 1px solid {t.BORDER_INPUT_FOCUS};
    }}

    QLineEdit#figmaAmountInput {{
        background: {t.BG_INPUT};
        border: 1px solid {t.BORDER_INPUT};
        border-radius: {t.RADIUS_INPUT}px;
        padding: 0 6px;
        min-height: 30px;
        max-height: 32px;
        color: {t.AMOUNT_TEXT};
        font-size: {t.BODY_PT}px;
    }}
    QLineEdit#figmaAmountInput:focus {{
        border: 1px solid {t.BORDER_INPUT_FOCUS};
    }}

    QComboBox#figmaCombo {{
        background: {t.BG_COMBO};
        border: 1px solid {t.BORDER_INPUT};
        border-radius: {t.RADIUS_INPUT}px;
        padding: 8px 12px;
        min-height: 22px;
        color: {t.LABEL_PURPLE};
    }}
    QComboBox#figmaCombo:focus {{
        border: 1px solid {t.BORDER_INPUT_FOCUS};
    }}
    QComboBox#figmaCombo::drop-down {{
        border: none;
        width: 28px;
    }}
    QComboBox#figmaCombo QAbstractItemView {{
        background: {t.BG_COMBO};
        border: 1px solid #D9D9D9;
        color: {t.LABEL_PURPLE};
        selection-background-color: {t.PRIMARY};
        selection-color: #FFFFFF;
        outline: none;
    }}

    {_sel} {{
        color: #FFFFFF;
        background: {_g};
        border: 1px solid {t.BTN_PRIMARY_BORDER};
        border-radius: {t.RADIUS_BUTTON}px;
    }}
    {_sel_h} {{
        background: {_gh};
        border: 1px solid {t.BTN_PRIMARY_BORDER_HOVER};
    }}
    {_sel_p} {{
        background: {_gp};
        border: 1px solid {t.BTN_PRIMARY_BORDER_PRESSED};
    }}
    {_sel_d} {{
        background: #B5B5B5;
        color: #EDEDED;
        border: 1px solid #9E9E9E;
    }}

    QPushButton#figmaUploadBtn {{
        padding: 9px 17px;
        font-size: {t.UPLOAD_BTN_PT}px;
        min-height: 22px;
    }}
    QPushButton#figmaActionBtn {{
        padding: 18px 24px;
        font-size: {t.ACTION_BTN_PT}px;
        min-height: 22px;
    }}
    QPushButton#figmaBlue {{
        padding: 9px 17px;
        font-size: {t.BODY_PT}px;
        font-weight: 600;
        min-height: 22px;
    }}
    QPushButton#figmaBlueSmall {{
        padding: 6px 22px;
        font-weight: 600;
        min-height: 22px;
    }}

    QPushButton#figmaHero {{
        color: #FFFFFF;
        background: {_g};
        border: 1px solid {t.BTN_PRIMARY_BORDER};
        border-radius: 6px;
        padding: 20px 32px;
        font-weight: 700;
        font-size: 18px;
        min-height: 36px;
    }}
    QPushButton#figmaHero:hover {{
        background: {_gh};
        border: 1px solid {t.BTN_PRIMARY_BORDER_HOVER};
    }}
    QPushButton#figmaHero:pressed {{
        background: {_gp};
        border: 1px solid {t.BTN_PRIMARY_BORDER_PRESSED};
    }}
    QPushButton#figmaHero:disabled {{
        background: #B5B5B5;
        color: #EDEDED;
        border: 1px solid #9E9E9E;
    }}
    """
