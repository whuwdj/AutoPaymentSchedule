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
    CARD_RADIUS = 8
    CARD_PAD_H = 16
    CARD_PAD_V = 16

    LABEL_PURPLE = "#722ED1"
    SECTION_MICRO = "#9254DE"
    TEXT_MUTED = "#8C8C8C"

    SUFFIX_UNIT = "#1890FF"

    BORDER_INPUT = "#1890FF"
    BORDER_INPUT_FOCUS = "#40A9FF"
    BG_INPUT = "#FFFFFF"
    BG_COMBO = "#F5F5F5"

    PRIMARY = "#1677FF"

    RADIUS_INPUT = 4
    RADIUS_BUTTON = 4

    PAGE_MARGIN_H = 20
    PAGE_MARGIN_V = 20
    GAP_SECTION = 24
    GAP_INLINE = 12
    GAP_TIGHT = 8

    ROW_GAP = 16
    DATA_ROW_GAP = 60
    DATA_LABEL_TO_YUAN = 120

    BODY_PT = 16
    UPLOAD_BTN_PT = BODY_PT + 2
    LABEL_PT = 18
    ACTION_BTN_PT = 20
    MICRO_PT = 12

    LABEL_PURPLE_PX = LABEL_PT
    SECTION_TITLE_PT = LABEL_PT

    W_PCT_INPUT = 100
    # 百分比行「输入框 + 间距 + %」占位宽度，总额行「元」同列右对齐
    W_PCT_SUFFIX_SLOT = 18
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
        color: #262626;
    }}
    QMessageBox QLabel {{
        color: #262626;
        background-color: {t.BG_CARD};
    }}
    QMessageBox QPushButton {{
        background-color: {t.BG_COMBO};
        color: #262626;
        border: 1px solid {t.CARD_BORDER};
        border-radius: {t.RADIUS_BUTTON}px;
        padding: 6px 18px;
        min-width: 72px;
        min-height: 24px;
    }}
    """


def build_app_stylesheet() -> str:
    t = Figma
    return f"""
    QWidget {{
        color: #262626;
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
        color: {t.SECTION_MICRO};
        font-size: {t.MICRO_PT * 1.5}px;
        font-weight: 500;
    }}
    QLabel#suffixYuan, QLabel#suffixUnit {{
        color: {t.SUFFIX_UNIT};
        font-size: {t.LABEL_PT}px;
    }}
    QLabel#summaryPurple {{
        color: {t.LABEL_PURPLE};
        font-size: {t.BODY_PT}px;
    }}
    QLabel#summaryValue {{
        color: {t.LABEL_PURPLE};
        font-size: {t.BODY_PT + 1}px;
        font-weight: 700;
        min-width: 2em;
    }}
    QLabel#footerHint {{
        color: {t.TEXT_MUTED};
        font-size: {t.MICRO_PT}px;
    }}

    QLineEdit#figmaInput {{
        background: {t.BG_INPUT};
        border: 1px solid {t.BORDER_INPUT};
        border-radius: {t.RADIUS_INPUT}px;
        padding: 0 6px;
        min-height: 30px;
        max-height: 32px;
        color: #262626;
        font-size: {t.BODY_PT}px;
    }}
    QLineEdit#figmaInput:focus {{
        border: 1px solid {t.BORDER_INPUT_FOCUS};
    }}

    QComboBox#figmaCombo {{
        background: {t.BG_COMBO};
        border: 1px solid {t.BORDER_INPUT};
        border-radius: {t.RADIUS_INPUT}px;
        padding: 8px 12px;
        min-height: 22px;
        color: #262626;
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
        selection-background-color: {t.PRIMARY};
        selection-color: #FFFFFF;
        outline: none;
    }}

    QPushButton#figmaUploadBtn {{
        background-color: {t.PRIMARY};
        color: #FFFFFF;
        border: none;
        border-radius: {t.RADIUS_BUTTON}px;
        padding: 9px 17px;
        font-size: {t.UPLOAD_BTN_PT}px;
        min-height: 22px;
    }}

    QPushButton#figmaActionBtn {{
        background-color: {t.PRIMARY};
        color: #FFFFFF;
        border: none;
        border-radius: {t.RADIUS_BUTTON}px;
        padding: 18px 24px;
        font-size: {t.ACTION_BTN_PT}px;
        min-height: 22px;
    }}

    QPushButton#figmaBlue {{
        background-color: {t.PRIMARY};
        color: #FFFFFF;
        border: none;
        border-radius: {t.RADIUS_BUTTON}px;
        padding: 9px 17px;
        font-size: {t.BODY_PT}px;
        font-weight: 600;
        min-height: 22px;
    }}

    QPushButton#figmaBlueSmall {{
        background-color: {t.PRIMARY};
        color: #FFFFFF;
        border: none;
        border-radius: {t.RADIUS_BUTTON}px;
        padding: 6px 22px;
        font-weight: 600;
        min-height: 22px;
    }}

    QPushButton#figmaHero {{
        background-color: {t.PRIMARY};
        color: #FFFFFF;
        border: none;
        border-radius: 6px;
        padding: 20px 32px;
        font-weight: 700;
        font-size: 18px;
        min-height: 36px;
    }}
    """
