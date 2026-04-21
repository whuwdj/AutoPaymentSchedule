# -*- coding: utf-8 -*-
"""
Figma「智能排款界面」设计令牌（与 Dev Mode 色值不一致时可在此微调）。

https://www.figma.com/design/aabHTb8OZnSlJBQPKrNhrZ/%E6%99%BA%E8%83%BD%E6%8E%92%E6%AC%BE%E7%95%8C%E9%9D%A2?node-id=0-1
"""

from __future__ import annotations

FIGMA_DESIGN_URL = (
    "https://www.figma.com/design/aabHTb8OZnSlJBQPKrNhrZ/"
    "%E6%99%BA%E8%83%BD%E6%8E%92%E6%AC%BE%E7%95%8C%E9%9D%A2?node-id=0-1"
)


class Figma:
    # 整页白底（稿面为纯白）
    BG_PAGE = "#FFFFFF"

    # 标签 / 说明
    LABEL_PURPLE = "#722ED1"
    SECTION_MICRO = "#9254DE"
    TEXT_MUTED = "#8C8C8C"

    # 单位与强调（稿中为蓝色「元」）
    SUFFIX_BLUE = "#1890FF"

    # 输入框：浅青描边
    BORDER_INPUT = "#13C2C2"
    BORDER_INPUT_FOCUS = "#08979C"
    BG_INPUT = "#FFFFFF"
    BG_COMBO = "#F5F5F5"

    # 主按钮（亮蓝）
    PRIMARY = "#1890FF"
    PRIMARY_HOVER = "#40A9FF"
    PRIMARY_PRESSED = "#096DD9"
    TEXT_ON_PRIMARY = "#FFFFFF"

    RADIUS_INPUT = 4
    RADIUS_BUTTON = 4
    RADIUS_HERO = 6

    PAGE_MARGIN_H = 28
    PAGE_MARGIN_V = 24
    GAP_SECTION = 20
    GAP_INLINE = 12
    GAP_TIGHT = 8

    BODY_PT = 14
    MICRO_PT = 12
    HERO_PT = 18

    # 表单紫字标签 / 小节标题（相对正文为 2 倍字号，对应稿图框选区域）
    LABEL_PURPLE_PX = BODY_PT * 1.5
    SECTION_MICRO_PX = MICRO_PT * 1.5

    # 与「可支付总额」同宽的短输入框：总额、百分比、银行区「可支付金额」三处共用（像素可调）
    W_INPUT_SHORT_MAX = 200

    # 银行 / 付款方式 下拉宽度（相对原先 min 200、180 各为 2 倍）
    W_COMBO_BANK = 400
    W_COMBO_PAY_METHOD = 360


def build_app_stylesheet() -> str:
    t = Figma
    return f"""
    QWidget {{
        background-color: {t.BG_PAGE};
        color: #262626;
        font-size: {t.BODY_PT}px;
    }}
    QWidget#centralRoot {{
        background-color: {t.BG_PAGE};
    }}

    QLabel#figmaPurpleLabel {{
        color: {t.LABEL_PURPLE};
        font-size: {t.LABEL_PURPLE_PX}px;
    }}
    QLabel#figmaMicroTitle {{
        color: {t.SECTION_MICRO};
        font-size: {t.SECTION_MICRO_PX}px;
        font-weight: 500;
    }}
    QLabel#suffixYuan {{
        color: {t.SUFFIX_BLUE};
        font-size: {t.BODY_PT}px;
        font-weight: 500;
    }}
    QLabel#summaryPurple {{
        color: {t.LABEL_PURPLE};
        font-size: {t.BODY_PT}px;
    }}
    QLabel#summaryValue {{
        color: {t.LABEL_PURPLE};
        font-size: {t.BODY_PT}px;
        font-weight: 600;
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
        padding: 8px 12px;
        min-height: 22px;
        color: #262626;
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
        selection-color: {t.TEXT_ON_PRIMARY};
        outline: none;
    }}

    QPushButton#figmaBlue {{
        background-color: {t.PRIMARY};
        color: {t.TEXT_ON_PRIMARY};
        border: none;
        border-radius: {t.RADIUS_BUTTON}px;
        padding: 8px 20px;
        font-weight: 600;
        min-height: 24px;
    }}
    QPushButton#figmaBlue:hover {{
        background-color: {t.PRIMARY_HOVER};
    }}
    QPushButton#figmaBlue:pressed {{
        background-color: {t.PRIMARY_PRESSED};
    }}

    QPushButton#figmaBlueSmall {{
        background-color: {t.PRIMARY};
        color: {t.TEXT_ON_PRIMARY};
        border: none;
        border-radius: {t.RADIUS_BUTTON}px;
        padding: 6px 22px;
        font-weight: 600;
        min-height: 22px;
    }}
    QPushButton#figmaBlueSmall:hover {{
        background-color: {t.PRIMARY_HOVER};
    }}
    QPushButton#figmaBlueSmall:pressed {{
        background-color: {t.PRIMARY_PRESSED};
    }}

    QPushButton#figmaHero {{
        background-color: {t.PRIMARY};
        color: {t.TEXT_ON_PRIMARY};
        border: none;
        border-radius: {t.RADIUS_HERO}px;
        padding: 20px 32px;
        font-weight: 700;
        font-size: {t.HERO_PT}px;
        min-height: 36px;
    }}
    QPushButton#figmaHero:hover {{
        background-color: {t.PRIMARY_HOVER};
    }}
    QPushButton#figmaHero:pressed {{
        background-color: {t.PRIMARY_PRESSED};
    }}
    """
