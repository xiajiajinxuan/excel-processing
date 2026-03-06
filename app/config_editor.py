# -*- coding: utf-8 -*-
"""表单式配置编辑对话框，无需接触 YAML 文本。"""

from pathlib import Path

import yaml
from PyQt6.QtWidgets import (
    QCheckBox,
    QComboBox,
    QDialog,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QSpinBox,
    QTableWidget,
    QTableWidgetItem,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)
from PyQt6.QtCore import Qt

from app.theme import (
    COLORS,
    FONT_FAMILY,
    BUTTON_STYLE_PRIMARY,
    BUTTON_STYLE_SECONDARY,
    input_style,
)


class ConfigEditorDialog(QDialog):
    """表单式配置编辑对话框。"""

    def __init__(self, parent, config_file, config_data, colors=None, font_family=None, btn_primary_style=None, btn_secondary_style=None):
        super().__init__(parent)
        self.config_file = Path(config_file)
        self.config = dict(config_data) if config_data else {}
        self.colors = colors or COLORS
        self.font_family = font_family or FONT_FAMILY
        self.btn_primary_style = btn_primary_style or BUTTON_STYLE_PRIMARY
        self.btn_secondary_style = btn_secondary_style or BUTTON_STYLE_SECONDARY
        self.input_style = input_style(self.colors, self.font_family)
        self.setWindowTitle("设置")
        self.setMinimumSize(520, 440)
        self.resize(620, 520)
        self.setStyleSheet(f"background-color: {self.colors['bg']};")
        self._build_ui()
        self._load_to_form()

    def _make_label(self, text):
        lbl = QLabel(text)
        lbl.setStyleSheet(
            f"color: {self.colors['text_secondary']}; font-size: 13px; font-family: {self.font_family}; min-width: 90px;"
        )
        return lbl

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(16)
        layout.setContentsMargins(20, 16, 20, 16)

        tabs = QTabWidget()
        tabs.setStyleSheet(f"""
            QTabWidget::pane {{ border: 1px solid {self.colors['border']}; border-radius: 8px; background: {self.colors['bg_panel']}; }}
            QTabBar::tab {{ padding: 8px 16px; font-family: {self.font_family}; color: {self.colors['text_secondary']}; }}
            QTabBar::tab:selected {{ color: {self.colors['primary']}; font-weight: bold; }}
        """)

        # ---------- 常规 ----------
        general = QWidget()
        general_layout = QVBoxLayout(general)
        row1 = QHBoxLayout()
        row1.addWidget(self._make_label("默认规则"))
        self.default_rule_combo = QComboBox()
        self.default_rule_combo.setStyleSheet(self.input_style)
        self.default_rule_combo.setEditable(False)
        row1.addWidget(self.default_rule_combo, 1)
        general_layout.addLayout(row1)
        general_layout.addStretch()
        tabs.addTab(general, "常规")

        # ---------- 日志 ----------
        log_w = QWidget()
        log_layout = QVBoxLayout(log_w)
        log_row1 = QHBoxLayout()
        self.log_to_file_check = QCheckBox("将日志写入文件")
        self.log_to_file_check.setStyleSheet(f"font-family: {self.font_family}; color: {self.colors['text']};")
        log_row1.addWidget(self.log_to_file_check)
        log_layout.addLayout(log_row1)
        log_row2 = QHBoxLayout()
        log_row2.addWidget(self._make_label("日志目录"))
        self.log_dir_edit = QLineEdit()
        self.log_dir_edit.setPlaceholderText("例如：output")
        self.log_dir_edit.setStyleSheet(self.input_style)
        log_row2.addWidget(self.log_dir_edit, 1)
        log_layout.addLayout(log_row2)
        log_layout.addStretch()
        tabs.addTab(log_w, "日志")

        # ---------- 规则列表 ----------
        rules_w = QWidget()
        rules_layout = QVBoxLayout(rules_w)
        rules_layout.addWidget(QLabel("规则显示名称与模板文件名（可从「规则」菜单远程获取新规则）："))
        rules_label = QLabel("规则列表")
        rules_label.setStyleSheet(f"color: {self.colors['text']}; font-family: {self.font_family};")
        rules_layout.addWidget(rules_label)
        self.rules_table = QTableWidget(0, 3)
        self.rules_table.setHorizontalHeaderLabels(["规则 ID", "显示名称", "模板文件名"])
        self.rules_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self.rules_table.setStyleSheet(f"""
            QTableWidget {{ background: {self.colors['surface']}; border: 1px solid {self.colors['border']}; border-radius: 6px; }}
            QTableWidget::item {{ padding: 4px; font-family: {self.font_family}; }}
        """)
        rules_layout.addWidget(self.rules_table)
        tabs.addTab(rules_w, "规则列表")

        # ---------- 远程规则 ----------
        remote_w = QWidget()
        remote_layout = QVBoxLayout(remote_w)
        r1 = QHBoxLayout()
        r1.addWidget(self._make_label("清单地址"))
        self.manifest_url_edit = QLineEdit()
        self.manifest_url_edit.setPlaceholderText("rules_manifest.json 的 URL")
        self.manifest_url_edit.setStyleSheet(self.input_style)
        r1.addWidget(self.manifest_url_edit, 1)
        remote_layout.addLayout(r1)
        r2 = QHBoxLayout()
        r2.addWidget(self._make_label("超时(秒)"))
        self.timeout_spin = QSpinBox()
        self.timeout_spin.setRange(1, 120)
        self.timeout_spin.setStyleSheet(self.input_style)
        r2.addWidget(self.timeout_spin)
        remote_layout.addStretch()
        tabs.addTab(remote_w, "远程规则")

        # ---------- 更新检查 ----------
        update_w = QWidget()
        update_layout = QVBoxLayout(update_w)
        u1 = QHBoxLayout()
        self.update_enabled_check = QCheckBox("启用检查更新")
        self.update_enabled_check.setStyleSheet(f"font-family: {self.font_family}; color: {self.colors['text']};")
        u1.addWidget(self.update_enabled_check)
        update_layout.addLayout(u1)
        u2 = QHBoxLayout()
        u2.addWidget(self._make_label("GitHub 用户名"))
        self.update_owner_edit = QLineEdit()
        self.update_owner_edit.setStyleSheet(self.input_style)
        u2.addWidget(self.update_owner_edit, 1)
        update_layout.addLayout(u2)
        u3 = QHBoxLayout()
        u3.addWidget(self._make_label("仓库名"))
        self.update_repo_edit = QLineEdit()
        self.update_repo_edit.setStyleSheet(self.input_style)
        u3.addWidget(self.update_repo_edit, 1)
        update_layout.addLayout(u3)
        u4 = QHBoxLayout()
        u4.addWidget(self._make_label("来源"))
        self.update_source_edit = QLineEdit()
        self.update_source_edit.setPlaceholderText("例如：github")
        self.update_source_edit.setStyleSheet(self.input_style)
        u4.addWidget(self.update_source_edit, 1)
        update_layout.addLayout(u4)
        update_layout.addStretch()
        tabs.addTab(update_w, "更新检查")

        layout.addWidget(tabs)

        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        cancel_btn = QPushButton("取消")
        cancel_btn.setStyleSheet(self.btn_secondary_style)
        cancel_btn.clicked.connect(self.reject)
        save_btn = QPushButton("保存")
        save_btn.setStyleSheet(self.btn_primary_style)
        save_btn.clicked.connect(self._on_save)
        btn_layout.addWidget(cancel_btn)
        btn_layout.addWidget(save_btn)
        layout.addLayout(btn_layout)

    def _load_to_form(self):
        """从 self.config 填到表单"""
        rules = self.config.get("rules") or {}
        rule_ids = list(rules.keys())
        self.default_rule_combo.clear()
        self.default_rule_combo.addItems(rule_ids)
        default = self.config.get("default_rule") or (rule_ids[0] if rule_ids else "")
        idx = self.default_rule_combo.findText(default)
        if idx >= 0:
            self.default_rule_combo.setCurrentIndex(idx)
        else:
            self.default_rule_combo.setCurrentIndex(0)

        log_cfg = self.config.get("log") or {}
        self.log_to_file_check.setChecked(bool(log_cfg.get("to_file", False)))
        self.log_dir_edit.setText(str(log_cfg.get("dir", "output")).strip())

        self.rules_table.setRowCount(len(rules))
        for i, (rid, info) in enumerate(rules.items()):
            id_item = QTableWidgetItem(rid)
            id_item.setFlags(id_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self.rules_table.setItem(i, 0, id_item)
            self.rules_table.setItem(i, 1, QTableWidgetItem(info.get("display_name", "")))
            self.rules_table.setItem(i, 2, QTableWidgetItem(info.get("template", "")))
        self.rules_table.setColumnWidth(0, 160)

        remote = self.config.get("rules_remote") or {}
        self.manifest_url_edit.setText(str(remote.get("manifest_url", "")).strip())
        self.timeout_spin.setValue(int(remote.get("timeout") or 15))

        upd = self.config.get("update") or {}
        self.update_enabled_check.setChecked(bool(upd.get("enabled", True)))
        self.update_owner_edit.setText(str(upd.get("owner", "")).strip())
        self.update_repo_edit.setText(str(upd.get("repo", "")).strip())
        self.update_source_edit.setText(str(upd.get("source", "github")).strip())

    def _form_to_config(self):
        """从表单写回 self.config"""
        self.config["default_rule"] = self.default_rule_combo.currentText().strip() or None

        self.config.setdefault("log", {})
        self.config["log"]["to_file"] = self.log_to_file_check.isChecked()
        self.config["log"]["dir"] = self.log_dir_edit.text().strip() or "output"

        rules = {}
        for i in range(self.rules_table.rowCount()):
            rid_item = self.rules_table.item(i, 0)
            name_item = self.rules_table.item(i, 1)
            tpl_item = self.rules_table.item(i, 2)
            if rid_item and rid_item.text().strip():
                rules[rid_item.text().strip()] = {
                    "display_name": (name_item.text() if name_item else "").strip(),
                    "template": (tpl_item.text() if tpl_item else "").strip(),
                }
        self.config["rules"] = rules

        self.config.setdefault("rules_remote", {})
        self.config["rules_remote"]["manifest_url"] = self.manifest_url_edit.text().strip() or None
        self.config["rules_remote"]["timeout"] = self.timeout_spin.value()

        self.config.setdefault("update", {})
        self.config["update"]["enabled"] = self.update_enabled_check.isChecked()
        self.config["update"]["owner"] = self.update_owner_edit.text().strip() or ""
        self.config["update"]["repo"] = self.update_repo_edit.text().strip() or ""
        self.config["update"]["source"] = self.update_source_edit.text().strip() or "github"
        self.config["update"]["tag_prefix"] = self.config["update"].get("tag_prefix", "")

    def _on_save(self):
        self._form_to_config()
        try:
            self.config_file.parent.mkdir(parents=True, exist_ok=True)
            with open(self.config_file, "w", encoding="utf-8") as f:
                yaml.dump(self.config, f, default_flow_style=False, allow_unicode=True)
        except Exception as e:
            QMessageBox.critical(self, "错误", f"保存失败：{str(e)}")
            return
        QMessageBox.information(self, "保存成功", "配置已保存，规则列表已刷新。")
        self.accept()
