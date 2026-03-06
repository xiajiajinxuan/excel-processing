# -*- coding: utf-8 -*-
"""Excel 数据处理工具 - PyQt6 主窗口。"""

import os
import shutil
import subprocess
import sys
import tempfile
import threading
import traceback
import webbrowser
from datetime import datetime
from pathlib import Path

import pandas as pd
import yaml
from PyQt6.QtWidgets import (
    QApplication,
    QDialog,
    QFileDialog,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMessageBox,
    QMenu,
    QMenuBar,
    QPushButton,
    QSplitter,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)
from PyQt6.QtCore import Qt, QTimer
from PyQt6.QtGui import QAction, QColor

from version import __version__
from core.update_checker import check_update, download_file
from core.remote_rules import run_remote_rules_dialog

from app.theme import (
    COLORS,
    FONT_FAMILY,
    BUTTON_STYLE_PRIMARY,
    BUTTON_STYLE_SECONDARY,
    BUTTON_STYLE_SUCCESS,
    PANEL_STYLE,
)
from app.config_loader import load_config as load_config_data, get_project_paths
from app.processor import run_rule as processor_run_rule, write_result_to_excel as processor_write_result
from app.config_editor import ConfigEditorDialog


class ExcelProcessingApp(QMainWindow):
    """Excel 数据处理工具 - PyQt6 主窗口"""

    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"Excel数据处理工具 v{__version__}")
        self.setMinimumSize(640, 480)
        self.resize(820, 640)
        self.setStyleSheet(f"background-color: {COLORS['bg']};")

        paths = get_project_paths()
        self.config_dir = paths["config_dir"]
        self.config_file = paths["config_file"]
        self.rules_dir = paths["rules_dir"]
        self.templates_dir = paths["templates_dir"]
        self.output_dir = paths["output_dir"]
        self.templates_dir.mkdir(exist_ok=True)
        self.output_dir.mkdir(exist_ok=True)
        self.rules_dir.mkdir(exist_ok=True)
        self.config_dir.mkdir(parents=True, exist_ok=True)

        self.rule_ids = []
        self.available_rules = []
        self.current_rule_id = None
        self._download_dialog = None

        self.load_config()
        self._setup_ui()
        self._setup_menu()
        self.load_rules()

    def _setup_ui(self):
        """构建主界面：分组面板 + 可调节日志区，无步骤条"""
        central = QWidget()
        central.setStyleSheet(f"background-color: {COLORS['bg']};")
        self.setCentralWidget(central)
        root_layout = QVBoxLayout(central)
        root_layout.setSpacing(16)
        root_layout.setContentsMargins(24, 20, 24, 20)

        title = QLabel("Excel 数据处理工具")
        title.setStyleSheet(
            f"color: {COLORS['text']}; font-size: 22px; font-weight: bold; "
            f"font-family: {FONT_FAMILY}; background: transparent;"
        )
        root_layout.addWidget(title)

        workflow = QGroupBox("处理流程")
        workflow.setStyleSheet(PANEL_STYLE)
        wf_layout = QVBoxLayout(workflow)
        wf_layout.setSpacing(14)
        wf_layout.setContentsMargins(4, 8, 4, 4)

        file_row = QHBoxLayout()
        file_row.setSpacing(10)
        lbl_file = QLabel("文件")
        lbl_file.setStyleSheet(
            f"color: {COLORS['text_secondary']}; font-size: 13px; min-width: 52px; "
            f"font-family: {FONT_FAMILY}; background: transparent;"
        )
        self.file_path_edit = QLineEdit()
        self.file_path_edit.setPlaceholderText("选择要处理的 .xlsx 文件…")
        self.file_path_edit.setMinimumHeight(36)
        self.file_path_edit.setStyleSheet(f"""
            QLineEdit {{
                background: {COLORS['surface']};
                border: 1px solid {COLORS['border']};
                border-radius: {COLORS['radius_md']}px;
                padding: 8px 12px;
                font-size: 13px;
                font-family: {FONT_FAMILY};
                color: {COLORS['text']};
            }}
            QLineEdit:focus {{ border-color: {COLORS['border_focus']}; }}
        """)
        btn_browse = QPushButton("浏览…")
        btn_browse.setStyleSheet(BUTTON_STYLE_SECONDARY)
        btn_browse.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_browse.setFixedHeight(36)
        btn_browse.clicked.connect(self.browse_file)
        file_row.addWidget(lbl_file)
        file_row.addWidget(self.file_path_edit)
        file_row.addWidget(btn_browse)
        wf_layout.addLayout(file_row)

        rule_row = QHBoxLayout()
        rule_row.setSpacing(10)
        lbl_rule = QLabel("规则")
        lbl_rule.setStyleSheet(
            f"color: {COLORS['text_secondary']}; font-size: 13px; min-width: 52px; "
            f"font-family: {FONT_FAMILY}; background: transparent;"
        )
        self.rule_display = QLineEdit()
        self.rule_display.setReadOnly(True)
        self.rule_display.setPlaceholderText("请选择处理规则…")
        self.rule_display.setMinimumHeight(36)
        self.rule_display.setStyleSheet(f"""
            QLineEdit {{
                background: {COLORS['surface']};
                border: 1px solid {COLORS['border']};
                border-radius: {COLORS['radius_md']}px;
                padding: 8px 12px;
                font-size: 13px;
                font-family: {FONT_FAMILY};
                color: {COLORS['text']};
            }}
            QLineEdit:focus {{ border-color: {COLORS['border_focus']}; }}
        """)
        btn_pick_rule = QPushButton("选择规则…")
        btn_pick_rule.setStyleSheet(BUTTON_STYLE_SECONDARY)
        btn_pick_rule.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_pick_rule.setFixedHeight(36)
        btn_pick_rule.clicked.connect(self.open_rule_picker)
        rule_row.addWidget(lbl_rule)
        rule_row.addWidget(self.rule_display)
        rule_row.addWidget(btn_pick_rule)
        wf_layout.addLayout(rule_row)

        btn_row = QHBoxLayout()
        btn_row.setSpacing(10)
        btn_download = QPushButton("下载模板")
        btn_download.setStyleSheet(BUTTON_STYLE_SUCCESS)
        btn_download.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_download.setFixedHeight(36)
        btn_download.clicked.connect(self.download_template)
        self.btn_process = QPushButton("处理数据")
        self.btn_process.setStyleSheet(BUTTON_STYLE_PRIMARY)
        self.btn_process.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_process.setFixedHeight(36)
        self.btn_process.clicked.connect(self.process_data)
        btn_row.addWidget(btn_download)
        btn_row.addWidget(self.btn_process)
        btn_row.addStretch()
        wf_layout.addLayout(btn_row)

        log_group = QGroupBox("运行日志")
        log_group.setStyleSheet(PANEL_STYLE)
        log_layout = QVBoxLayout(log_group)
        log_layout.setContentsMargins(4, 8, 4, 4)
        self.result_text = QTextEdit()
        self.result_text.setReadOnly(True)
        self.result_text.setMinimumHeight(140)
        self.result_text.setStyleSheet(f"""
            QTextEdit {{
                background: #FAFAFA;
                border: 1px solid {COLORS['border']};
                border-radius: {COLORS['radius_md']}px;
                padding: 12px;
                font-family: "Cascadia Code", Consolas, "Microsoft YaHei UI", monospace;
                font-size: 12px;
                color: {COLORS['text']};
            }}
        """)
        log_layout.addWidget(self.result_text)

        splitter = QSplitter(Qt.Orientation.Vertical)
        splitter.addWidget(workflow)
        splitter.addWidget(log_group)
        splitter.setStretchFactor(0, 0)
        splitter.setStretchFactor(1, 1)
        splitter.setSizes([200, 400])
        root_layout.addWidget(splitter, 1)

    def _setup_menu(self):
        menubar = self.menuBar()
        rule_menu = menubar.addMenu("规则")
        act_remote = QAction("从远程获取规则", self)
        act_remote.triggered.connect(self.on_remote_rules)
        rule_menu.addAction(act_remote)
        settings_menu = menubar.addMenu("设置")
        act_edit_config = QAction("编辑配置文件…", self)
        act_edit_config.triggered.connect(self.on_edit_config)
        settings_menu.addAction(act_edit_config)
        act_show_config_dir = QAction("打开配置目录", self)
        act_show_config_dir.triggered.connect(self.on_show_config_dir)
        settings_menu.addAction(act_show_config_dir)
        help_menu = menubar.addMenu("帮助")
        act_update = QAction("检查更新", self)
        act_update.triggered.connect(self.on_check_update)
        help_menu.addAction(act_update)
        act_log_dir = QAction("打开日志目录", self)
        act_log_dir.triggered.connect(self.on_open_log_dir)
        help_menu.addAction(act_log_dir)
        act_about = QAction("关于", self)
        act_about.triggered.connect(self.show_about)
        help_menu.addAction(act_about)

    def load_config(self):
        self.config = load_config_data()
        if not self.config_file.exists():
            self.save_config()

    def save_config(self):
        try:
            self.config_file.parent.mkdir(parents=True, exist_ok=True)
            with open(self.config_file, "w", encoding="utf-8") as f:
                yaml.dump(self.config, f, default_flow_style=False, allow_unicode=True)
        except Exception as e:
            QMessageBox.critical(self, "错误", f"保存配置文件时出错: {str(e)}")

    def on_edit_config(self):
        dialog = ConfigEditorDialog(
            self, self.config_file, self.config,
            COLORS, FONT_FAMILY, BUTTON_STYLE_PRIMARY, BUTTON_STYLE_SECONDARY
        )
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.load_config()
            self.update_rule_list()

    def on_show_config_dir(self):
        self.config_dir.mkdir(parents=True, exist_ok=True)
        path = self.config_dir.resolve()
        if sys.platform == "win32":
            os.startfile(str(path))
        else:
            subprocess.run(["xdg-open", str(path)], check=False)

    def get_rule_display_name(self, rule_id):
        return self.config.get("rules", {}).get(rule_id, {}).get("display_name", rule_id)

    def get_rule_template(self, rule_id):
        return self.config.get("rules", {}).get(rule_id, {}).get("template", "")

    def get_rule_by_template(self, template_name):
        for rule_id, rule_info in self.config.get("rules", {}).items():
            if rule_info.get("template") == template_name:
                return rule_id
        return None

    def on_remote_rules(self):
        self.load_config()
        styles = {
            "COLORS": COLORS,
            "FONT_FAMILY": FONT_FAMILY,
            "BUTTON_STYLE_PRIMARY": BUTTON_STYLE_PRIMARY,
            "BUTTON_STYLE_SECONDARY": BUTTON_STYLE_SECONDARY,
        }
        run_remote_rules_dialog(
            self,
            get_config=lambda: self.config,
            save_config=self.save_config,
            refresh_rule_list=self.update_rule_list,
            styles=styles,
            rules_dir=self.rules_dir,
            templates_dir=self.templates_dir,
        )

    def on_open_log_dir(self):
        log_dir = Path(self.config.get("log", {}).get("dir", "output")).resolve()
        log_dir.mkdir(parents=True, exist_ok=True)
        if sys.platform == "win32":
            os.startfile(str(log_dir))
        else:
            subprocess.run(["xdg-open", str(log_dir)], check=False)

    def on_check_update(self):
        result = check_update(self.config, __version__)
        if result.get("error"):
            QMessageBox.information(self, "检查更新", result["error"])
            return
        if not result.get("has_new"):
            QMessageBox.information(self, "检查更新", f"当前已是最新版本（v{result['current']}）。")
            return
        msg = (
            f"发现新版本 v{result['latest']}（当前 v{result['current']}）。\n\n"
            f"更新说明：\n{result.get('release_notes', '')[:500]}\n\n"
            "是否立即更新？更新将下载新版本并替换当前程序后重启。"
        )
        if QMessageBox.question(self, "发现新版本", msg, QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No, QMessageBox.StandardButton.Yes) == QMessageBox.StandardButton.Yes:
            self.do_update_and_restart(result.get("download_url", ""))

    def do_update_and_restart(self, download_url):
        if not download_url or not download_url.strip():
            QMessageBox.information(self, "更新", "暂无可用下载链接，请前往发布页手动下载。")
            return
        current_exe = sys.executable
        current_dir = os.path.dirname(current_exe)
        exe_name = os.path.basename(current_exe)
        if not self._is_dir_writable(current_dir):
            QMessageBox.warning(
                self, "更新",
                "当前程序所在目录无写入权限，无法自动替换。请将程序安装到有写权限的目录（如用户目录），或手动下载安装。",
            )
            try:
                webbrowser.open(download_url)
            except Exception:
                pass
            return
        temp_dir = tempfile.gettempdir()
        new_exe_path = os.path.join(temp_dir, "Excel数据处理工具_new.exe")
        if QMessageBox.question(
            self, "更新",
            "将在后台下载新版本，下载完成后会提示您是否更新。您可以继续使用程序。\n\n是否开始下载？",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.Yes,
        ) != QMessageBox.StandardButton.Yes:
            return
        self._download_dialog = QDialog(self)
        self._download_dialog.setWindowTitle("更新")
        self._download_dialog.setFixedSize(360, 120)
        self._download_dialog.setStyleSheet(f"""
            QDialog {{ background: {COLORS['surface']}; border-radius: {COLORS['radius_lg']}px; }}
            QLabel {{ color: {COLORS['text']}; font-size: 13px; font-family: {FONT_FAMILY}; }}
        """)
        layout = QVBoxLayout(self._download_dialog)
        layout.setContentsMargins(24, 24, 24, 24)
        layout.addWidget(QLabel("正在后台下载新版本，请稍候…\n（可继续使用程序）"))
        self._download_dialog.show()

        def do_download():
            success = download_file(download_url, new_exe_path)
            QTimer.singleShot(0, lambda: self._on_download_done(success, new_exe_path, current_exe, exe_name, temp_dir, download_url))
        threading.Thread(target=do_download, daemon=True).start()

    def _on_download_done(self, success, new_exe_path, current_exe, exe_name, temp_dir, download_url):
        if self._download_dialog and self._download_dialog.isVisible():
            self._download_dialog.close()
            self._download_dialog = None
        if not success:
            QMessageBox.critical(self, "更新", "下载失败，请检查网络或稍后重试。")
            try:
                webbrowser.open(download_url)
            except Exception:
                pass
            return
        if QMessageBox.question(self, "更新", "新版本已下载完成，是否立即退出并更新？", QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No, QMessageBox.StandardButton.Yes) != QMessageBox.StandardButton.Yes:
            return
        bat_path = os.path.join(temp_dir, "excel_tool_update.bat")
        bat_content = f'''@echo off
:wait
tasklist /fi "imagename eq {exe_name}" 2>nul | find /i "{exe_name}" >nul && (timeout /t 2 /nobreak >nul & goto wait)
copy /y "{new_exe_path}" "{current_exe}"
start "" "{current_exe}"
del /f /q "{new_exe_path}" 2>nul
del /f /q "%~f0" 2>nul
'''
        try:
            with open(bat_path, "w", encoding="gbk") as f:
                f.write(bat_content)
        except (OSError, UnicodeEncodeError):
            QMessageBox.critical(self, "更新", "无法创建更新脚本，请手动下载安装。")
            try:
                webbrowser.open(download_url)
            except Exception:
                pass
            return
        creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0x08000000) if sys.platform == "win32" else 0
        try:
            subprocess.Popen(["cmd", "/c", bat_path], creationflags=creationflags, cwd=temp_dir)
        except Exception:
            QMessageBox.critical(self, "更新", "无法启动更新流程，请手动下载安装。")
            try:
                webbrowser.open(download_url)
            except Exception:
                pass
            return
        QMessageBox.information(self, "更新", "正在退出并更新，请稍候程序将自动重新启动。")
        QApplication.quit()
        sys.exit(0)

    def _is_dir_writable(self, dir_path):
        try:
            test_file = os.path.join(dir_path, ".write_test_tmp")
            with open(test_file, "w"):
                pass
            os.remove(test_file)
            return True
        except OSError:
            return False

    def update_rule_list(self):
        if not self.rules_dir.exists():
            self.rules_dir.mkdir(exist_ok=True)
        rule_files = [f.stem for f in self.rules_dir.glob("*.py") if f.stem != "__init__"]
        self.rule_ids = []
        for rule_id in rule_files:
            self.rule_ids.append(rule_id)
            if rule_id not in self.config.get("rules", {}):
                if "rules" not in self.config:
                    self.config["rules"] = {}
                self.config["rules"][rule_id] = {"display_name": rule_id, "template": f"{rule_id}_template.xlsx"}
                self.save_config()
        if self.current_rule_id in self.rule_ids:
            self.set_current_rule(self.current_rule_id)
        elif self.rule_ids:
            self.set_current_rule(self.rule_ids[0])
        else:
            self.current_rule_id = None
            self.rule_display.clear()

    def set_current_rule(self, rule_id: str | None):
        if not rule_id or rule_id not in self.rule_ids:
            self.current_rule_id = None
            self.rule_display.clear()
            return
        self.current_rule_id = rule_id
        self.rule_display.setText(self.get_rule_display_name(rule_id))

    def browse_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "选择Excel文件", "", "Excel文件 (*.xlsx);;所有文件 (*.*)")
        if path:
            self.file_path_edit.setText(path)
            file_name = os.path.basename(path)
            rule_id = self.get_rule_by_template(file_name)
            if rule_id and rule_id in self.rule_ids:
                self.set_current_rule(rule_id)

    def download_template(self):
        if not self.current_rule_id:
            QMessageBox.critical(self, "错误", "请先选择一个处理规则")
            return
        rule_id = self.current_rule_id
        template_name = self.get_rule_template(rule_id)
        if not template_name:
            QMessageBox.critical(self, "错误", f"规则 '{rule_id}' 没有对应的模板")
            return
        template_path = self.templates_dir / template_name
        if not template_path.exists():
            QMessageBox.critical(self, "错误", f"模板文件 '{template_name}' 不存在")
            return
        save_path, _ = QFileDialog.getSaveFileName(self, "保存模板文件", template_name, "Excel文件 (*.xlsx);;所有文件 (*.*)")
        if not save_path:
            return
        try:
            shutil.copy2(template_path, save_path)
            QMessageBox.information(self, "成功", f"模板已保存到: {save_path}")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"保存模板时出错: {str(e)}")

    def process_data(self):
        file_path = self.file_path_edit.text().strip()
        if not file_path:
            QMessageBox.critical(self, "错误", "请选择Excel文件")
            return
        if not self.current_rule_id:
            QMessageBox.critical(self, "错误", "请选择处理规则")
            return
        rule_id = self.current_rule_id
        self.btn_process.setText("处理中")
        self.btn_process.setEnabled(False)
        QApplication.processEvents()
        self.result_text.clear()
        self.log("——— 开始新任务 ———", live=True, timestamp=True)
        try:
            self.log("正在读取文件…", live=True)
            result, elapsed = processor_run_rule(rule_id, file_path, self.rules_dir)
            if result is None:
                raise RuntimeError(elapsed)
            self.log("已读取，正在执行规则…", live=True)
            self.log("规则执行完成，正在写入 Excel…", live=True)
            self.display_result(result, file_path, elapsed)
        except Exception as e:
            error_full = f"处理数据时出错: {str(e)}\n{traceback.format_exc()}"
            QMessageBox.critical(self, "错误", error_full)
            self.log(f"处理数据时出错: {str(e)}", level="error", live=True)
        finally:
            self.btn_process.setText("处理数据")
            self.btn_process.setEnabled(True)

    def open_rule_picker(self):
        if not self.rule_ids:
            QMessageBox.information(self, "提示", "当前没有可用的处理规则，请先在 rules 目录中添加规则。")
            return
        dlg = QDialog(self)
        dlg.setWindowTitle("选择处理规则")
        dlg.setMinimumSize(360, 360)
        dlg.setStyleSheet(
            f"QDialog {{ background: {COLORS['surface']}; }} "
            f"QLabel {{ color: {COLORS['text']}; font-family: {FONT_FAMILY}; }} "
        )
        layout = QVBoxLayout(dlg)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(10)
        layout.addWidget(QLabel("请选择要使用的处理规则："))
        search_edit = QLineEdit()
        search_edit.setPlaceholderText("输入关键字筛选规则（名称 / ID）…")
        layout.addWidget(search_edit)
        list_widget = QListWidget()
        for rid in self.rule_ids:
            name = self.get_rule_display_name(rid)
            item = QListWidgetItem(f"{name}  ({rid})")
            item.setData(Qt.ItemDataRole.UserRole, rid)
            list_widget.addItem(item)
            if rid == self.current_rule_id:
                list_widget.setCurrentItem(item)
        layout.addWidget(list_widget)
        btn_row = QHBoxLayout()
        btn_ok = QPushButton("确定")
        btn_ok.setStyleSheet(BUTTON_STYLE_PRIMARY)
        btn_cancel = QPushButton("取消")
        btn_cancel.setStyleSheet(BUTTON_STYLE_SECONDARY)
        btn_row.addStretch()
        btn_row.addWidget(btn_cancel)
        btn_row.addWidget(btn_ok)
        layout.addLayout(btn_row)

        def apply_filter(text: str):
            keyword = (text or "").strip().lower()
            for i in range(list_widget.count()):
                item = list_widget.item(i)
                visible = not keyword or keyword in item.text().lower()
                item.setHidden(not visible)
        search_edit.textChanged.connect(apply_filter)

        def accept_selection():
            item = list_widget.currentItem()
            if not item:
                QMessageBox.information(dlg, "提示", "请先选择一个规则。")
                return
            rid = item.data(Qt.ItemDataRole.UserRole)
            self.set_current_rule(rid)
            dlg.accept()
        btn_ok.clicked.connect(accept_selection)
        btn_cancel.clicked.connect(dlg.reject)
        list_widget.itemDoubleClicked.connect(lambda _item: accept_selection())
        dlg.exec()

    def display_result(self, result, file_path, elapsed=None):
        if isinstance(result, dict) and "error" in result:
            self.log(f"错误: {result['error']}", level="error", live=True)
            return
        self.log("处理完成！", live=True)
        try:
            output_file = processor_write_result(file_path, result, self.output_dir)
            if elapsed is not None:
                self.log(f"处理耗时: {elapsed:.2f} 秒", live=True)
            self.log(f"结果已保存到: {output_file}", live=True)
            if QMessageBox.question(self, "处理完成", f"结果已保存到: {output_file}\n是否打开文件？", QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No, QMessageBox.StandardButton.Yes) == QMessageBox.StandardButton.Yes:
                os.startfile(str(output_file))
        except Exception as e:
            error_full = f"保存结果时出错: {str(e)}\n{traceback.format_exc()}"
            QMessageBox.critical(self, "错误", error_full)
            self.log(f"保存结果时出错: {str(e)}", level="error", live=True)

    def load_rules(self):
        self.available_rules = []
        if not self.rules_dir.exists():
            self.rules_dir.mkdir(exist_ok=True)
            self.create_example_rule()
        for file in self.rules_dir.glob("*.py"):
            if file.name not in ("__init__.py",) and "__pycache__" not in file.name:
                self.available_rules.append(file.stem)
        self.update_rule_list()
        if self.available_rules:
            self.log(f"已加载 {len(self.available_rules)} 个处理规则")
        else:
            self.log("未找到处理规则，请在 rules 目录下添加规则文件")

    def create_example_rule(self):
        init_path = self.rules_dir / "__init__.py"
        if not init_path.exists():
            init_path.write_text("# 规则包初始化文件\n", encoding="utf-8")
        example_path = self.rules_dir / "example_rule.py"
        if not example_path.exists():
            example_path.write_text("""# 示例处理规则
import pandas as pd

def process(data_df, **kwargs):
    result_df = data_df.copy()
    result_df['处理状态'] = '已处理'
    return result_df

def get_rule_info():
    return {"name": "示例规则", "description": "示例", "version": "1.0", "author": "系统"}
""", encoding="utf-8")
        if not self.templates_dir.exists():
            self.templates_dir.mkdir(exist_ok=True)
        example_template_path = self.templates_dir / "example_template.xlsx"
        if not example_template_path.exists():
            df = pd.DataFrame({"姓名": ["张三", "李四", "王五"], "年龄": [25, 30, 35], "部门": ["技术部", "市场部", "人事部"]})
            df.to_excel(example_template_path, index=False)
        self.log("已创建示例规则文件和模板")

    def _get_log_path(self) -> Path:
        return Path(self.config.get("log", {}).get("dir", "output")) / "app.log"

    def _write_log_file(self, message: str, level: str = "info") -> None:
        if not self.config.get("log", {}).get("to_file", False):
            return
        log_path = self._get_log_path()
        try:
            log_path.parent.mkdir(parents=True, exist_ok=True)
            ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            prefix = ""
            if level == "warning":
                prefix = "[警告] "
            elif level == "error":
                prefix = "[错误] "
            with open(log_path, "a", encoding="utf-8") as f:
                f.write(f"[{ts}] {prefix}{message}\n")
        except Exception:
            pass

    def log(self, message: str, level: str = "info", live: bool = False, timestamp: bool | None = None) -> None:
        if timestamp is None:
            timestamp = True  # 统一带时间戳 [YYYY-MM-DD HH:MM:SS]
        ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S") if timestamp else ""
        prefix = f"[{ts}] " if ts else ""
        if level == "warning":
            prefix += "[警告] "
        elif level == "error":
            prefix += "[错误] "
        color = QColor(COLORS["text"])
        if level == "error":
            color = QColor("#c0392b")
        elif level == "warning":
            color = QColor("#d97706")
        self.result_text.setTextColor(color)
        self.result_text.append(f"{prefix}{message}")
        self.result_text.setTextColor(QColor(COLORS["text"]))
        if live:
            QApplication.processEvents()
        scrollbar = self.result_text.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())
        self._write_log_file(message, level=level)

    def show_about(self):
        QMessageBox.about(
            self, "关于",
            f"Excel数据处理工具\n\n版本：v{__version__}\n\n基于 Python 与 PyQt6 的桌面应用，支持插件式规则与模板管理。",
        )
