# -*- coding: utf-8 -*-
"""Excel 数据处理工具 - 程序入口。"""

import os
import sys

# 确保项目根在路径中
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from PyQt6.QtWidgets import QApplication

from app.theme import app_global_stylesheet
from app.main_window import ExcelProcessingApp


def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    app.setStyleSheet(app_global_stylesheet())
    win = ExcelProcessingApp()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
