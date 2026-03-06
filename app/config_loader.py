# -*- coding: utf-8 -*-
"""配置与路径集中管理：从 config 目录读取 YAML，提供规则/模板/输出等路径。"""

import shutil
from pathlib import Path
from typing import Any

import yaml

# 相对于项目根目录（运行时的当前工作目录）
CONFIG_DIR = Path("config")
CONFIG_FILE = CONFIG_DIR / "config.yaml"

_DEFAULT_CONFIG = {
    "rules": {
        "example_rule": {"display_name": "示例规则", "template": "example_template.xlsx"}
    },
    "default_rule": "example_rule",
    "log": {"to_file": False, "dir": "output"},
}


def load_config(base_path: Path | None = None) -> dict[str, Any]:
    """
    加载 YAML 配置。若存在根目录 config.yaml 则迁移到 config/config.yaml；
    若配置文件不存在则返回默认配置（不写入文件）。
    """
    base = base_path or Path(".")
    config_file = base / CONFIG_FILE
    root_legacy = base / "config.yaml"

    if not config_file.exists() and root_legacy.exists():
        try:
            config_file.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy(root_legacy, config_file)
        except Exception:
            pass

    if not config_file.exists():
        return dict(_DEFAULT_CONFIG)

    try:
        with open(config_file, "r", encoding="utf-8") as f:
            data = yaml.safe_load(f)
    except Exception:
        return dict(_DEFAULT_CONFIG)

    if not isinstance(data, dict):
        return dict(_DEFAULT_CONFIG)

    # 兼容旧配置：确保存在 log 配置块
    log_cfg = data.setdefault("log", {})
    log_cfg.setdefault("to_file", False)
    log_cfg.setdefault("dir", "output")
    return data


def get_project_paths(base_path: Path | None = None) -> dict[str, Path]:
    """返回规则、模板、输出等目录路径（相对于 base_path，默认当前目录）。"""
    base = base_path or Path(".")
    return {
        "config_dir": base / CONFIG_DIR,
        "config_file": base / CONFIG_FILE,
        "rules_dir": base / "rules",
        "templates_dir": base / "templates",
        "output_dir": base / "output",
    }


def resolve_config_file(base_path: Path | None = None) -> Path:
    """返回配置文件绝对路径（用于保存等）。"""
    base = base_path or Path(".")
    return (base / CONFIG_FILE).resolve()
