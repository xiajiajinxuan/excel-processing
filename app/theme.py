# -*- coding: utf-8 -*-
"""
UI 主题与样式：瑞士现代 + 生产力工具配色。
主色：青绿 #0D9488；CTA：橙色 #F97316；背景：浅青白 #F0FDFA；文字：深青 #134E4A。
"""

COLORS = {
    "bg": "#F0FDFA",
    "bg_panel": "#FFFFFF",
    "surface": "#FFFFFF",
    "primary": "#0D9488",
    "primary_hover": "#0F766E",
    "primary_light": "#CCFBF1",
    "cta": "#F97316",
    "cta_hover": "#EA580C",
    "cta_light": "#FFEDD5",
    "text": "#134E4A",
    "text_secondary": "#0F766E",
    "text_muted": "#5EEAD4",
    "border": "#99F6E4",
    "border_focus": "#0D9488",
    "radius_sm": "6",
    "radius_md": "8",
    "radius_lg": "10",
}

FONT_FAMILY = '"Segoe UI", "PingFang SC", "Microsoft YaHei UI", sans-serif'


def _button_style_primary():
    return f"""
    QPushButton {{
        background-color: {COLORS['cta']};
        color: #FFFFFF;
        border: none;
        border-radius: {COLORS['radius_md']}px;
        padding: 10px 22px;
        font-size: 14px;
        font-weight: bold;
        font-family: {FONT_FAMILY};
    }}
    QPushButton:hover {{ background-color: {COLORS['cta_hover']}; }}
    QPushButton:pressed {{ background-color: {COLORS['cta_hover']}; }}
    QPushButton:disabled {{ background-color: #A8A29E; color: #E7E5E4; }}
"""


def _button_style_secondary():
    return f"""
    QPushButton {{
        background-color: transparent;
        color: {COLORS['primary']};
        border: 2px solid {COLORS['primary']};
        border-radius: {COLORS['radius_md']}px;
        padding: 8px 20px;
        font-size: 13px;
        font-family: {FONT_FAMILY};
    }}
    QPushButton:hover {{
        background-color: {COLORS['primary_light']};
        border-color: {COLORS['primary_hover']};
        color: {COLORS['primary_hover']};
    }}
    QPushButton:pressed {{ background-color: #99F6E4; }}
"""


def _panel_style():
    return f"""
    QGroupBox {{
        font-family: {FONT_FAMILY};
        font-size: 14px;
        font-weight: bold;
        color: {COLORS['text']};
        border: 1px solid {COLORS['border']};
        border-radius: {COLORS['radius_lg']}px;
        margin-top: 12px;
        padding-top: 16px;
        padding-left: 14px;
        padding-right: 14px;
        padding-bottom: 14px;
        background-color: {COLORS['bg_panel']};
    }}
    QGroupBox::title {{
        subcontrol-origin: margin;
        subcontrol-position: top left;
        left: 14px;
        padding: 0 8px;
        background-color: {COLORS['bg_panel']};
        color: {COLORS['primary']};
    }}
"""


BUTTON_STYLE_PRIMARY = _button_style_primary()
BUTTON_STYLE_SECONDARY = _button_style_secondary()
BUTTON_STYLE_SUCCESS = BUTTON_STYLE_SECONDARY
PANEL_STYLE = _panel_style()


def input_style(colors=None, font_family=None):
    """输入控件（QLineEdit、QSpinBox、QComboBox）统一样式。"""
    colors = colors or COLORS
    font_family = font_family or FONT_FAMILY
    return f"""
        QLineEdit, QSpinBox, QComboBox {{
            background: {colors['surface']};
            border: 1px solid {colors['border']};
            border-radius: 6px;
            padding: 6px 10px;
            font-size: 13px;
            font-family: {font_family};
            color: {colors['text']};
            min-height: 20px;
        }}
        QLineEdit:focus, QSpinBox:focus, QComboBox:focus {{ border-color: {colors['border_focus']}; }}
        QComboBox::drop-down {{ border: none; }}
    """


def app_global_stylesheet():
    """主窗口与菜单栏等全局样式（供 QApplication.setStyleSheet 使用）。"""
    return f"""
        QMainWindow {{ background-color: {COLORS['bg']}; }}
        QMenuBar {{
            background: {COLORS['surface']};
            color: {COLORS['text']};
            border-bottom: 1px solid {COLORS['border']};
            padding: 4px 0;
            font-family: {FONT_FAMILY};
        }}
        QMenuBar::item:selected {{ background: {COLORS['primary_light']}; color: {COLORS['primary']}; }}
        QMenu {{
            background: {COLORS['surface']};
            border: 1px solid {COLORS['border']};
            border-radius: {COLORS['radius_sm']}px;
            padding: 6px;
        }}
        QMenu::item:selected {{ background: {COLORS['primary_light']}; }}
    """
