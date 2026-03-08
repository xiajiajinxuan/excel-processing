# -*- coding: utf-8 -*-
"""
远程规则模块：从静态托管或本地目录拉取规则清单，用户选择后下载/复制到本地 rules/ 并合并 config。
使用标准库 urllib、pathlib，不新增第三方依赖。
"""

import json
import shutil
import urllib.error
import urllib.request
import ssl
from pathlib import Path

from app.processor import list_rule_ids as _list_rule_ids
from core.update_checker import download_file as _download_file


def _is_remote_url(s):
    """判断是否为远程 URL（http/https）。"""
    s = (s or "").strip().lower()
    return s.startswith("http://") or s.startswith("https://")


def _expand_rule_files(rule):
    """
    若规则有 rule_id 但无 files 或 files 为空，按约定生成 files：
    - rule_id/rule_id.py -> rules
    - 若有 template：rule_id/doc/template/<template> -> templates
    返回新 rule 的拷贝（含 files），原 rule 不变。
    """
    r = dict(rule)
    rule_id = (r.get("rule_id") or "").strip()
    if not rule_id:
        return r
    files = r.get("files")
    if isinstance(files, list) and len(files) > 0:
        return r
    expanded = [{"path": f"{rule_id}/{rule_id}.py", "target": "rules"}]
    template = (r.get("template") or "").strip()
    if template:
        expanded.append({"path": f"{rule_id}/doc/template/{template}", "target": "templates"})
    r["files"] = expanded
    return r


def fetch_manifest(manifest_url, timeout=15, source="remote"):
    """
    拉取并解析规则清单（支持远程 URL 或本地路径）。

    :param manifest_url: 清单 JSON 的 URL 或本地目录/文件路径
    :param timeout: 请求超时（秒），仅远程有效
    :param source: "remote" 或 "local"，未传时根据 manifest_url 是否以 http(s) 开头推断
    :return: (data_dict, error_message)。成功时 data_dict 为 {"base_url": str, "rules": [...]}，error_message 为 None；
             失败时 data_dict 为 None，error_message 为 str。
    """
    if not manifest_url or not str(manifest_url).strip():
        return None, "未配置规则清单地址"
    raw_url = str(manifest_url).strip()
    use_local = source == "local" or (source != "remote" and not _is_remote_url(raw_url))

    if use_local:
        # 本地：从文件系统读取 JSON
        p = Path(raw_url)
        if not p.exists():
            return None, f"本地路径不存在：{p}"
        if p.is_dir():
            manifest_path = p / "rules_manifest.json"
            if not manifest_path.exists():
                return None, f"目录下未找到 rules_manifest.json：{p}"
        else:
            manifest_path = p
        try:
            with open(manifest_path, "r", encoding="utf-8") as f:
                raw_content = f.read()
        except (OSError, UnicodeDecodeError) as e:
            return None, f"读取清单文件失败：{e}"
        data = None
        try:
            data = json.loads(raw_content)
        except json.JSONDecodeError as e:
            # 修复 base_url 中未转义的反斜杠（Windows 路径在 JSON 中需写为 \\）
            if "Invalid" in str(e.msg) and "escape" in str(e.msg).lower():
                idx = raw_content.find('"base_url": "')
                if idx != -1:
                    start = idx + len('"base_url": "')
                    end = start
                    while end < len(raw_content):
                        c = raw_content[end]
                        if c == "\\":
                            end += 1
                            if end < len(raw_content):
                                end += 1
                            continue
                        if c == '"':
                            break
                        end += 1
                    if end <= len(raw_content) and (end == len(raw_content) or raw_content[end] == '"'):
                        old_val = raw_content[start:end]
                        new_val = old_val.replace("\\", "\\\\")
                        raw_content = raw_content[:start] + new_val + raw_content[end:]
                        try:
                            data = json.loads(raw_content)
                        except json.JSONDecodeError:
                            pass
            if data is None:
                return None, f"清单格式错误：{e}"
        manifest_dir = str(manifest_path.parent.resolve())
        base_url = (data.get("base_url") or "").strip().rstrip("/") or manifest_dir
        if not base_url or _is_remote_url(base_url):
            base_url = manifest_dir
        rules = data.get("rules")
        if not isinstance(rules, list):
            return None, "清单格式错误：缺少 rules 数组"
        for i, r in enumerate(rules):
            if not isinstance(r, dict):
                return None, f"清单格式错误：rules[{i}] 应为对象"
            if not (r.get("rule_id") or "").strip():
                return None, f"清单格式错误：rules[{i}] 缺少 rule_id"
            if "files" in r and not isinstance(r.get("files"), list):
                return None, f"清单格式错误：rules[{i}] 的 files 应为数组"
        rules = [_expand_rule_files(r) for r in rules]
        return {"base_url": base_url, "rules": rules}, None

    # 远程：HTTP(S) 拉取
    try:
        ctx = ssl.create_default_context()
        req = urllib.request.Request(raw_url, headers={"Accept": "application/json"})
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
        if not (r.get("rule_id") or "").strip():
            return None, f"清单格式错误：rules[{i}] 缺少 rule_id"
        if "files" in r and not isinstance(r.get("files"), list):
            return None, f"清单格式错误：rules[{i}] 的 files 应为数组"
    rules = [_expand_rule_files(r) for r in rules]
    return {"base_url": base_url, "rules": rules}, None


def download_rule(base_url, rule_entry, rules_dir, templates_dir, timeout, on_file_exists, source="remote"):
    """
    按清单项下载或复制一条规则的所有文件到 rules_dir / templates_dir。

    :param base_url: 规则文件根 URL 或本地目录（无末尾斜杠）
    :param rule_entry: 清单中的规则对象，含 rule_id, files (列表，每项含 path, target)
    :param rules_dir: 本地 rules 目录 Path
    :param templates_dir: 本地 templates 目录 Path
    :param timeout: 下载超时（秒），仅远程有效
    :param on_file_exists: 回调 (dest_path: Path) -> "overwrite" | "skip" | "cancel"
    :param source: "remote" 或 "local"，未传时根据 base_url 是否以 http(s) 开头推断
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

    use_local = source == "local" or (source != "remote" and base_url and not _is_remote_url(base_url))
    base_path = Path(base_url) if use_local and base_url else None

    written = 0
    missing_local = []
    for item in files:
        path = (item.get("path") or "").strip()
        target = (item.get("target") or "rules").strip().lower()
        if not path:
            continue
        if target == "templates" and "doc/template" in path.replace("\\", "/"):
            dest_dir = rules_dir / rule_id / "doc" / "template"
            dest_path = dest_dir / Path(path).name
        elif target == "templates":
            dest_dir = templates_dir
            dest_path = dest_dir / path
        else:
            dest_dir = rules_dir
            dest_path = dest_dir / path
        if dest_path.is_dir():
            continue
        dest_path.parent.mkdir(parents=True, exist_ok=True)

        if dest_path.exists():
            choice = on_file_exists(dest_path)
            if choice == "cancel":
                return written > 0, "用户取消"
            if choice == "skip":
                continue

        if use_local and base_path is not None:
            src_path = base_path / path
            if not src_path.exists() and target in ("rules", "templates"):
                src_path = base_path / target / path
            if not src_path.exists() and target == "templates":
                src_path = base_path / "rules" / path
            if not src_path.exists():
                missing_local.append(path)
                continue
            try:
                shutil.copy2(src_path, dest_path)
                written += 1
            except OSError as e:
                return written > 0, f"复制失败：{path}（{e}）"
        else:
            file_url = f"{base_url}/{path}" if base_url else path
            if _download_file(file_url, str(dest_path), timeout=timeout):
                written += 1
            else:
                return written > 0, f"下载失败：{path}"
    if written == 0 and missing_local:
        return False, f"本地文件不存在：{missing_local[0]}"
    if missing_local:
        return True, "以下文件未找到：" + "、".join(missing_local)
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
    """返回本地已存在的规则 ID 集合（仅子目录形式 rules/<name>/<name>.py）。"""
    return set(_list_rule_ids(Path(rules_dir)))


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
            src = (remote.get("source") or "remote").strip().lower()
            if src not in ("remote", "local"):
                src = "remote"
            return {
                "manifest_url": (remote.get("manifest_url") or "").strip(),
                "source": src,
                "timeout": int(remote.get("timeout") or 15),
            }

        def _on_refresh(self):
            rcfg = self._get_remote_config()
            if not rcfg["manifest_url"]:
                QMessageBox.information(
                    self,
                    "提示",
                    "请在「设置」->「编辑配置文件」中配置 rules_remote.manifest_url 后重试。",
                )
                return
            self._status_label.setText("正在获取清单…")
            QApplication.processEvents()
            data, err = fetch_manifest(rcfg["manifest_url"], rcfg["timeout"], source=rcfg["source"])
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
            rcfg = self._get_remote_config()
            timeout = rcfg.get("timeout", 15)
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
                    base_url, entry, rules_dir, templates_dir, timeout, on_file_exists,
                    source=rcfg["source"],
                )
                if success:
                    merge_rule_to_config(
                        config,
                        rule_id,
                        entry.get("display_name") or rule_id,
                        entry.get("template") or "",
                    )
                    ok_count += 1
                if msg and "用户取消" not in msg:
                    QMessageBox.warning(self, "下载规则", f"规则 {rule_id}：{msg}")

            if ok_count > 0:
                self._save_config()
                self._refresh_rule_list()
                QMessageBox.information(self, "完成", f"已成功安装 {ok_count} 个规则，规则列表已刷新。")
                self._local_ids = _get_local_rule_ids(self._rules_dir)
                self._fill_table()

    dlg = RemoteRulesDialog()
    dlg.exec()
