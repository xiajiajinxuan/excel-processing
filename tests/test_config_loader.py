# -*- coding: utf-8 -*-
"""app.config_loader 模块测试。"""

import tempfile
import unittest
from pathlib import Path

import yaml

# 将项目根加入 path
import sys
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from app.config_loader import load_config, get_project_paths, CONFIG_DIR, CONFIG_FILE


class TestConfigLoader(unittest.TestCase):
    def test_load_config_returns_default_when_no_file(self):
        """当配置文件不存在时，应返回默认配置。"""
        with tempfile.TemporaryDirectory() as tmp:
            base = Path(tmp)
            config = load_config(base)
            self.assertIsInstance(config, dict)
            self.assertIn("rules", config)
            self.assertIn("default_rule", config)
            self.assertIn("log", config)
            self.assertFalse(config["log"].get("to_file"))
            self.assertEqual(config["log"].get("dir"), "output")

    def test_load_config_reads_existing_yaml(self):
        """当 config/config.yaml 存在时，应正确解析并返回。"""
        with tempfile.TemporaryDirectory() as tmp:
            base = Path(tmp)
            config_dir = base / CONFIG_DIR
            config_dir.mkdir(parents=True, exist_ok=True)
            config_file = base / CONFIG_FILE
            data = {
                "default_rule": "my_rule",
                "rules": {"my_rule": {"display_name": "我的规则", "template": "t.xlsx"}},
                "log": {"to_file": True, "dir": "logs"},
            }
            with open(config_file, "w", encoding="utf-8") as f:
                yaml.dump(data, f, default_flow_style=False, allow_unicode=True)
            config = load_config(base)
            self.assertEqual(config["default_rule"], "my_rule")
            self.assertEqual(config["rules"]["my_rule"]["display_name"], "我的规则")
            self.assertTrue(config["log"]["to_file"])
            self.assertEqual(config["log"]["dir"], "logs")

    def test_get_project_paths(self):
        """get_project_paths 应返回各目录路径。"""
        with tempfile.TemporaryDirectory() as tmp:
            base = Path(tmp)
            paths = get_project_paths(base)
            self.assertEqual(paths["config_dir"], base / "config")
            self.assertEqual(paths["config_file"], base / "config" / "config.yaml")
            self.assertEqual(paths["rules_dir"], base / "rules")
            self.assertEqual(paths["templates_dir"], base / "templates")
            self.assertEqual(paths["output_dir"], base / "output")


if __name__ == "__main__":
    unittest.main()
