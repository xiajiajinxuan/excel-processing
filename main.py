# -*- coding: utf-8 -*-
import os
import importlib
import pandas as pd
from pathlib import Path
import sys
import traceback
import shutil
import yaml
import subprocess
import tempfile
import threading
import webbrowser

from version import __version__
from update_checker import check_update, download_file
from remote_rules import run_remote_rules_dialog

# 添加当前目录到系统路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from PyQt6.QtWidgets import (
    QApplication,
    QMainWindow,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QGroupBox,
    QLabel,
    QLineEdit,
    QPushButton,
    QComboBox,
    QTextEdit,
    QFileDialog,
    QMessageBox,
    QMenuBar,
    QMenu,
    QDialog,
    QFrame,
    QSplitter,
)
from PyQt6.QtCore import Qt, QTimer
from PyQt6.QtGui import QAction

# ---------- 新 UI 设计系统（瑞士现代 + 生产力工具配色）----------
# 主色：青绿 #0D9488；CTA：橙色 #F97316；背景：浅青白 #F0FDFA；文字：深青 #134E4A
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

BUTTON_STYLE_PRIMARY = f"""
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

BUTTON_STYLE_SECONDARY = f"""
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

BUTTON_STYLE_SUCCESS = BUTTON_STYLE_SECONDARY

PANEL_STYLE = f"""
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


class ExcelProcessingApp(QMainWindow):
    """Excel 数据处理工具 - PyQt6 主窗口"""

    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"Excel数据处理工具 v{__version__}")
        self.setMinimumSize(640, 480)
        self.resize(820, 640)
        self.setStyleSheet(f"background-color: {COLORS['bg']};")

        self.templates_dir = Path("templates")
        self.output_dir = Path("output")
        self.rules_dir = Path("rules")
        self.templates_dir.mkdir(exist_ok=True)
        self.output_dir.mkdir(exist_ok=True)
        self.rules_dir.mkdir(exist_ok=True)

        self.rule_ids = []
        self.available_rules = []
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

        # ---------- 顶部标题 ----------
        title = QLabel("Excel 数据处理工具")
        title.setStyleSheet(
            f"color: {COLORS['text']}; font-size: 22px; font-weight: bold; "
            f"font-family: {FONT_FAMILY}; background: transparent;"
        )
        root_layout.addWidget(title)

        # ---------- 分组：处理流程 ----------
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
        self.rule_combo = QComboBox()
        self.rule_combo.setMinimumHeight(36)
        self.rule_combo.setStyleSheet(f"""
            QComboBox {{
                background: {COLORS['surface']};
                border: 1px solid {COLORS['border']};
                border-radius: {COLORS['radius_md']}px;
                padding: 8px 12px;
                font-size: 13px;
                font-family: {FONT_FAMILY};
                color: {COLORS['text']};
            }}
            QComboBox:focus {{ border-color: {COLORS['border_focus']}; }}
            QComboBox::drop-down {{ border: none; }}
        """)
        rule_row.addWidget(lbl_rule)
        rule_row.addWidget(self.rule_combo)
        rule_row.addStretch()
        wf_layout.addLayout(rule_row)

        btn_row = QHBoxLayout()
        btn_row.setSpacing(10)
        btn_download = QPushButton("下载模板")
        btn_download.setStyleSheet(BUTTON_STYLE_SUCCESS)
        btn_download.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_download.setFixedHeight(36)
        btn_download.clicked.connect(self.download_template)
        btn_process = QPushButton("处理数据")
        btn_process.setStyleSheet(BUTTON_STYLE_PRIMARY)
        btn_process.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_process.setFixedHeight(36)
        btn_process.clicked.connect(self.process_data)
        btn_row.addWidget(btn_download)
        btn_row.addWidget(btn_process)
        btn_row.addStretch()
        wf_layout.addLayout(btn_row)

        # ---------- 分组：运行日志（可调节高度）----------
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
        """菜单栏：规则 -> 从远程获取规则；帮助 -> 检查更新、关于"""
        menubar = self.menuBar()

        rule_menu = menubar.addMenu("规则")
        act_remote = QAction("从远程获取规则", self)
        act_remote.triggered.connect(self.on_remote_rules)
        rule_menu.addAction(act_remote)

        help_menu = menubar.addMenu("帮助")
        act_update = QAction("检查更新", self)
        act_update.triggered.connect(self.on_check_update)
        help_menu.addAction(act_update)
        act_about = QAction("关于", self)
        act_about.triggered.connect(self.show_about)
        help_menu.addAction(act_about)

    def load_config(self):
        """加载 YAML 配置"""
        config_path = Path("config.yaml")
        if not config_path.exists():
            self.config = {
                "rules": {
                    "example_rule": {"display_name": "示例规则", "template": "example_template.xlsx"}
                },
                "default_rule": "example_rule",
            }
            self.save_config()
        else:
            try:
                with open(config_path, "r", encoding="utf-8") as f:
                    self.config = yaml.safe_load(f)
            except Exception as e:
                QMessageBox.critical(self, "错误", f"读取配置文件时出错: {str(e)}")
                self.config = {
                    "rules": {
                        "example_rule": {"display_name": "示例规则", "template": "example_template.xlsx"}
                    },
                    "default_rule": "example_rule",
                }

    def save_config(self):
        """保存 YAML 配置"""
        config_path = Path("config.yaml")
        try:
            with open(config_path, "w", encoding="utf-8") as f:
                yaml.dump(self.config, f, default_flow_style=False, allow_unicode=True)
        except Exception as e:
            QMessageBox.critical(self, "错误", f"保存配置文件时出错: {str(e)}")

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
        """打开「从远程获取规则」对话框"""
        # 重新加载一次配置，确保用户手动修改 config.yaml 后能立即生效
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

    def on_check_update(self):
        """检查更新"""
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
        """下载新 exe 并提示重启"""
        if not download_url or not download_url.strip():
            QMessageBox.information(self, "更新", "暂无可用下载链接，请前往发布页手动下载。")
            return
        current_exe = sys.executable
        current_dir = os.path.dirname(current_exe)
        exe_name = os.path.basename(current_exe)
        if not self._is_dir_writable(current_dir):
            QMessageBox.warning(
                self,
                "更新",
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
            self,
            "更新",
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
        """刷新规则下拉列表"""
        rules_dir = Path("rules")
        if not rules_dir.exists():
            rules_dir.mkdir(exist_ok=True)
        rule_files = [f.stem for f in rules_dir.glob("*.py") if f.stem != "__init__"]
        rule_display_names = []
        self.rule_ids = []
        for rule_id in rule_files:
            rule_display_names.append(self.get_rule_display_name(rule_id))
            self.rule_ids.append(rule_id)
            if rule_id not in self.config.get("rules", {}):
                if "rules" not in self.config:
                    self.config["rules"] = {}
                self.config["rules"][rule_id] = {"display_name": rule_id, "template": f"{rule_id}_template.xlsx"}
                self.save_config()
        self.rule_combo.clear()
        self.rule_combo.addItems(rule_display_names)
        if rule_display_names:
            self.rule_combo.setCurrentIndex(0)

    def browse_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "选择Excel文件", "", "Excel文件 (*.xlsx);;所有文件 (*.*)")
        if path:
            self.file_path_edit.setText(path)
            file_name = os.path.basename(path)
            rule_id = self.get_rule_by_template(file_name)
            if rule_id and rule_id in self.rule_ids:
                idx = self.rule_ids.index(rule_id)
                self.rule_combo.setCurrentIndex(idx)

    def download_template(self):
        idx = self.rule_combo.currentIndex()
        if idx < 0:
            QMessageBox.critical(self, "错误", "请先选择一个处理规则")
            return
        rule_id = self.rule_ids[idx]
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
        idx = self.rule_combo.currentIndex()
        if idx < 0:
            QMessageBox.critical(self, "错误", "请选择处理规则")
            return
        rule_id = self.rule_ids[idx]
        try:
            rule_module = importlib.import_module(f"rules.{rule_id}")
            data_df = pd.read_excel(file_path)
            result = rule_module.process(data_df, excel_file=file_path)
            self.display_result(result, file_path)
        except Exception as e:
            error_message = f"处理数据时出错: {str(e)}\n{traceback.format_exc()}"
            QMessageBox.critical(self, "错误", error_message)
            self.result_text.clear()
            self.result_text.setPlainText(error_message)

    def display_result(self, result, file_path):
        self.result_text.clear()
        if isinstance(result, dict) and "error" in result:
            self.result_text.setPlainText(f"错误: {result['error']}\n")
            return
        self.result_text.append("处理完成！\n")
        file_name = os.path.basename(file_path)
        base_name = os.path.splitext(file_name)[0]
        output_file = self.output_dir / f"{base_name}_processed.xlsx"
        try:
            self.output_dir.mkdir(exist_ok=True)
            original_dfs = {}
            with pd.ExcelFile(file_path) as xls:
                for sheet_name in xls.sheet_names:
                    original_dfs[sheet_name] = pd.read_excel(xls, sheet_name=sheet_name)
            if isinstance(result, dict) and "deduction_record" in result and "日期" in result["deduction_record"].columns:
                result["deduction_record"]["日期"] = pd.to_datetime(result["deduction_record"]["日期"]).dt.strftime("%Y-%m-%d")
            with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
                for sheet_name, df in original_dfs.items():
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
                    self.result_text.append(f"已保留原始工作表: {sheet_name}\n")
                sheet_mapping = {"deduction_record": "扣缴记录", "monthly_summary": "月度汇总"}
                if isinstance(result, dict):
                    for key, value in result.items():
                        if isinstance(value, pd.DataFrame) and not value.empty and key != "error":
                            sheet_name = sheet_mapping.get(key, key)
                            value.to_excel(writer, sheet_name=sheet_name, index=False)
                            self.result_text.append(f"已创建工作表: {sheet_name}，包含 {len(value)} 行数据\n")
                elif isinstance(result, pd.DataFrame):
                    result.to_excel(writer, sheet_name="结果", index=False)
                    self.result_text.append(f"已创建工作表: 结果，包含 {len(result)} 行数据\n")
            self.result_text.append(f"\n结果已保存到: {output_file}")
            if QMessageBox.question(self, "处理完成", f"结果已保存到: {output_file}\n是否打开文件？", QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No, QMessageBox.StandardButton.Yes) == QMessageBox.StandardButton.Yes:
                os.startfile(output_file)
        except Exception as e:
            error_message = f"保存结果时出错: {str(e)}\n{traceback.format_exc()}"
            QMessageBox.critical(self, "错误", error_message)
            self.result_text.setPlainText(error_message)

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

    def log(self, message):
        from datetime import datetime
        ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.result_text.append(f"[{ts}] {message}\n")
        scrollbar = self.result_text.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())

    def show_about(self):
        QMessageBox.about(
            self,
            "关于",
            f"Excel数据处理工具\n\n版本：v{__version__}\n\n基于 Python 与 PyQt6 的桌面应用，支持插件式规则与模板管理。",
        )


def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    # 全局 Pro 样式：菜单栏、消息框等与主界面一致
    app.setStyleSheet(f"""
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
    """)
    win = ExcelProcessingApp()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
