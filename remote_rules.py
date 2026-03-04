# -*- coding: utf-8 -*-
"""
远程规则模块：从静态托管拉取规则清单，用户选择后下载到本地 rules/ 并合并 config。
使用标准库 urllib，不新增第三方依赖。
"""

import json
import urllib.error
import urllib.request
import ssl
from pathlib import Path

from update_checker import download_file as _download_file


def fetch_manifest(manifest_url, timeout=15):
    """
    拉取并解析远程规则清单。

    :param manifest_url: 清单 JSON 的 URL
    :param timeout: 请求超时（秒）
    :return: (data_dict, error_message)。成功时 data_dict 为 {"base_url": str, "rules": [...]}，error_message 为 None；
             失败时 data_dict 为 None，error_message 为 str。
    """
    if not manifest_url or not str(manifest_url).strip():
        return None, "未配置远程规则清单地址"
    url = str(manifest_url).strip()
    try:
        ctx = ssl.create_default_context()
        req = urllib.request.Request(url, headers={"Accept": "application/json"})
        with urllib.request.urlopen(req, timeout=timeout, context=ctx) as resp:
            if resp.status != 200:
                return None, f"获取清单失败（HTTP {resp.status}）"
            raw = resp.read().decode("utf-8")
    except urllib.error.URLError as e:
        return None, f"网络错误：{e.reason}"
    except urllib.error.HTTPError as e:
        return None, f"获取清单失败（HTTP {e.code}）"
    except (OSError, UnicodeDecodeError) as e:
        return None, f"读取清单失败：{e}"

    try:
        data = json.loads(raw)
    except json.JSONDecodeError as e:
        return None, f"清单格式错误：{e}"

    if not isinstance(data, dict):
        return None, "清单格式错误：根节点应为对象"
    base_url = (data.get("base_url") or "").strip().rstrip("/")
    rules = data.get("rules")
    if not isinstance(rules, list):
        return None, "清单格式错误：缺少 rules 数组"
    for i, r in enumerate(rules):
        if not isinstance(r, dict):
            return None, f"清单格式错误：rules[{i}] 应为对象"
        if not (r.get("rule_id") and isinstance(r.get("files"), list)):
            return None, f"清单格式错误：rules[{i}] 缺少 rule_id 或 files"
    return {"base_url": base_url, "rules": rules}, None


def download_rule(base_url, rule_entry, rules_dir, templates_dir, timeout, on_file_exists):
    """
    按清单项下载一条规则的所有文件到 rules_dir / templates_dir。

    :param base_url: 规则文件根 URL（无末尾斜杠）
    :param rule_entry: 清单中的规则对象，含 rule_id, files (列表，每项含 path, target)
    :param rules_dir: 本地 rules 目录 Path
    :param templates_dir: 本地 templates 目录 Path
    :param timeout: 下载超时（秒）
    :param on_file_exists: 回调 (dest_path: Path) -> "overwrite" | "skip" | "cancel"
    :return: (success: bool, message: str)。success 为 True 表示至少成功写入一个文件；message 为错误或说明。
    """
    rule_id = rule_entry.get("rule_id") or ""
    files = rule_entry.get("files") or []
    if not files:
        return False, "该规则没有可下载文件"

    rules_dir = Path(rules_dir)
    templates_dir = Path(templates_dir)
    rules_dir.mkdir(parents=True, exist_ok=True)
    templates_dir.mkdir(parents=True, exist_ok=True)

    written = 0
    for item in files:
        path = (item.get("path") or "").strip()
        target = (item.get("target") or "rules").strip().lower()
        if not path:
            continue
        if target == "templates":
            dest_dir = templates_dir
        else:
            dest_dir = rules_dir
        dest_path = dest_dir / path
        if dest_path.is_dir():
            continue
        dest_path.parent.mkdir(parents=True, exist_ok=True)

        file_url = f"{base_url}/{path}" if base_url else path
        if dest_path.exists():
            choice = on_file_exists(dest_path)
            if choice == "cancel":
                return written > 0, "用户取消"
            if choice == "skip":
                continue
            # overwrite
        if _download_file(file_url, str(dest_path), timeout=timeout):
            written += 1
        else:
            return written > 0, f"下载失败：{path}"
    return True, "" if written else "没有写入任何文件"


def merge_rule_to_config(config, rule_id, display_name, template=""):
    """
    将一条规则合并进 config 的 rules 项（原地修改 config）。

    :param config: 主配置字典（含 "rules" 键）
    :param rule_id: 规则 ID
    :param display_name: 显示名称
    :param template: 模板文件名（可选）
    """
    if "rules" not in config:
        config["rules"] = {}
    if rule_id not in config["rules"]:
        config["rules"][rule_id] = {}
    entry = config["rules"][rule_id]
    if display_name:
        entry["display_name"] = display_name
    if template:
        entry["template"] = template
    if "template" not in entry:
        entry["template"] = f"{rule_id}_template.xlsx"


def _get_local_rule_ids(rules_dir):
    """返回本地已存在的规则 ID 集合（rules/*.py 的 stem，排除 __init__）。"""
    rules_dir = Path(rules_dir)
    if not rules_dir.exists():
        return set()
    return {f.stem for f in rules_dir.glob("*.py") if f.stem != "__init__"}


def run_remote_rules_dialog(parent, get_config, save_config, refresh_rule_list, styles, rules_dir, templates_dir):
    """
    打开「从远程获取规则」对话框。

    :param parent: 父窗口
    :param get_config: 无参，返回当前 config 字典
    :param save_config: 无参，保存 config 到文件
    :param refresh_rule_list: 无参，刷新主窗口规则列表
    :param styles: 样式字典，至少含 COLORS、FONT_FAMILY、BUTTON_STYLE_PRIMARY、BUTTON_STYLE_SECONDARY 等
    :param rules_dir: 本地 rules 目录（Path 或 str）
    :param templates_dir: 本地 templates 目录（Path 或 str）
    """
    from PyQt6.QtWidgets import (
        QDialog,
        QVBoxLayout,
        QHBoxLayout,
        QTableWidget,
        QTableWidgetItem,
        QHeaderView,
        QPushButton,
        QLabel,
        QMessageBox,
        QAbstractItemView,
        QCheckBox,
        QWidget,
        QApplication,
        QLineEdit,
    )
    from PyQt6.QtCore import Qt

    COLORS = styles.get("COLORS", {})
    FONT_FAMILY = styles.get("FONT_FAMILY", "sans-serif")
    BTN_PRIMARY = styles.get("BUTTON_STYLE_PRIMARY", "")
    BTN_SECONDARY = styles.get("BUTTON_STYLE_SECONDARY", "")

    rules_dir = Path(rules_dir)
    templates_dir = Path(templates_dir)

    class RemoteRulesDialog(QDialog):
        def __init__(self):
            super().__init__(parent)
            self.setWindowTitle("从远程获取规则")
            self.setMinimumSize(560, 400)
            self.resize(620, 440)
            self._get_config = get_config
            self._save_config = save_config
            self._refresh_rule_list = refresh_rule_list
            self._rules_dir = rules_dir
            self._templates_dir = templates_dir
            self._manifest_data = None
            self._all_rules = []
            self._local_ids = set()
            self._setup_ui()

        def _setup_ui(self):
            self.setStyleSheet(
                f"QDialog {{ background: {COLORS.get('bg', '#fff')}; }} "
                f"QLabel {{ color: {COLORS.get('text', '#333')}; font-family: {FONT_FAMILY}; }} "
                f"QTableWidget {{ background: {COLORS.get('surface', '#fff')}; border: 1px solid {COLORS.get('border', '#ddd')}; }} "
            )
            layout = QVBoxLayout(self)
            layout.setSpacing(12)
            layout.setContentsMargins(20, 20, 20, 20)

            self._status_label = QLabel("点击「刷新清单」获取远程规则列表")
            self._status_label.setStyleSheet(f"color: {COLORS.get('text_secondary', '#666')}; font-size: 13px;")
            layout.addWidget(self._status_label)

            # 关键字搜索框：支持按规则名称 / 说明 / ID 过滤
            self._search_edit = QLineEdit()
            self._search_edit.setPlaceholderText("输入关键字筛选规则（名称 / 说明 / ID）…")
            self._search_edit.setClearButtonEnabled(True)
            self._search_edit.textChanged.connect(self._on_search_changed)
            self._search_edit.setMinimumHeight(30)
            layout.addWidget(self._search_edit)

            self._table = QTableWidget(0, 4)
            self._table.setHorizontalHeaderLabels(["选择", "规则名称", "说明", "状态"])
            self._table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
            self._table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
            self._table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
            self._table.setColumnWidth(0, 50)
            self._table.setColumnWidth(1, 140)
            self._table.setColumnWidth(3, 80)
            layout.addWidget(self._table)

            btn_row = QHBoxLayout()
            btn_refresh = QPushButton("刷新清单")
            btn_refresh.setStyleSheet(BTN_SECONDARY)
            btn_refresh.clicked.connect(self._on_refresh)
            btn_download = QPushButton("下载选中")
            btn_download.setStyleSheet(BTN_PRIMARY)
            btn_download.clicked.connect(self._on_download)
            btn_close = QPushButton("关闭")
            btn_close.setStyleSheet(BTN_SECONDARY)
            btn_close.clicked.connect(self.accept)
            btn_row.addWidget(btn_refresh)
            btn_row.addWidget(btn_download)
            btn_row.addStretch()
            btn_row.addWidget(btn_close)
            layout.addLayout(btn_row)

        def _get_remote_config(self):
            cfg = self._get_config()
            remote = (cfg or {}).get("rules_remote") or {}
            return {
                "manifest_url": (remote.get("manifest_url") or "").strip(),
                "timeout": int(remote.get("timeout") or 15),
            }

        def _on_refresh(self):
            rcfg = self._get_remote_config()
            if not rcfg["manifest_url"]:
                QMessageBox.information(
                    self,
                    "提示",
                    "请在 config.yaml 中配置 rules_remote.manifest_url 后重试。",
                )
                return
            self._status_label.setText("正在获取清单…")
            QApplication.processEvents()
            data, err = fetch_manifest(rcfg["manifest_url"], rcfg["timeout"])
            if err:
                self._status_label.setText("")
                QMessageBox.warning(self, "获取清单失败", err)
                return
            self._manifest_data = data
            self._all_rules = list((data.get("rules") or []))
            self._local_ids = _get_local_rule_ids(self._rules_dir)
            self._fill_table()
            self._status_label.setText(f"已加载 {len(self._all_rules)} 条远程规则")

        def _on_search_changed(self, text: str):
            """根据搜索关键字过滤规则列表。"""
            keyword = (text or "").strip().lower()
            if not self._manifest_data:
                return
            if not keyword:
                self._fill_table()
                return
            filtered = []
            for r in self._all_rules:
                rule_id = (r.get("rule_id") or "").lower()
                name = (r.get("display_name") or "").lower()
                desc = (r.get("description") or "").lower()
                if keyword in rule_id or keyword in name or keyword in desc:
                    filtered.append(r)
            self._fill_table(filtered)
            self._status_label.setText(
                f"已加载 {len(self._all_rules)} 条远程规则，当前筛选出 {len(filtered)} 条"
            )

        def _fill_table(self, rules=None):
            if rules is None:
                rules = self._all_rules or ((self._manifest_data or {}).get("rules") or [])
            self._table.setRowCount(len(rules))
            for row, r in enumerate(rules):
                rule_id = r.get("rule_id") or ""
                display_name = r.get("display_name") or rule_id
                desc = (r.get("description") or "")[:80]
                installed = rule_id in self._local_ids
                status = "已安装" if installed else "未安装"

                check = QCheckBox()
                check.setChecked(not installed)
                cell_widget = QWidget()
                cell_layout = QHBoxLayout(cell_widget)
                cell_layout.setContentsMargins(4, 0, 4, 0)
                cell_layout.addWidget(check)
                cell_layout.addStretch()
                self._table.setCellWidget(row, 0, cell_widget)

                self._table.setItem(row, 1, QTableWidgetItem(display_name))
                self._table.setItem(row, 2, QTableWidgetItem(desc))
                self._table.setItem(row, 3, QTableWidgetItem(status))

                self._table.setRowHeight(row, 36)
                setattr(check, "_rule_entry", r)

        def _on_download(self):
            if not self._manifest_data:
                QMessageBox.information(self, "提示", "请先点击「刷新清单」。")
                return
            base_url = self._manifest_data.get("base_url") or ""
            rules_dir = self._rules_dir
            templates_dir = self._templates_dir
            timeout = self._get_remote_config().get("timeout", 15)
            config = self._get_config()

            def on_file_exists(dest_path):
                r = QMessageBox.question(
                    self,
                    "文件已存在",
                    f"文件已存在：{dest_path.name}\n是否覆盖？",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No | QMessageBox.StandardButton.Cancel,
                    QMessageBox.StandardButton.Yes,
                )
                if r == QMessageBox.StandardButton.Yes:
                    return "overwrite"
                if r == QMessageBox.StandardButton.No:
                    return "skip"
                return "cancel"

            selected = []
            for row in range(self._table.rowCount()):
                w = self._table.cellWidget(row, 0)
                if w:
                    cb = w.findChild(QCheckBox)
                    if cb and cb.isChecked():
                        entry = getattr(cb, "_rule_entry", None)
                        if entry:
                            selected.append(entry)
            if not selected:
                QMessageBox.information(self, "提示", "请至少勾选一条规则。")
                return

            ok_count = 0
            for entry in selected:
                rule_id = entry.get("rule_id") or ""
                success, msg = download_rule(
                    base_url, entry, rules_dir, templates_dir, timeout, on_file_exists
                )
                if success:
                    merge_rule_to_config(
                        config,
                        rule_id,
                        entry.get("display_name") or rule_id,
                        entry.get("template") or "",
                    )
                    ok_count += 1
                if msg and not success and "用户取消" not in msg:
                    QMessageBox.warning(self, "下载规则", f"规则 {rule_id}：{msg}")

            if ok_count > 0:
                self._save_config()
                self._refresh_rule_list()
                QMessageBox.information(self, "完成", f"已成功安装 {ok_count} 个规则，规则列表已刷新。")
                self._local_ids = _get_local_rule_ids(self._rules_dir)
                self._fill_table()

    dlg = RemoteRulesDialog()
    dlg.exec()
