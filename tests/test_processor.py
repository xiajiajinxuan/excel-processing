# -*- coding: utf-8 -*-
"""app.processor 模块测试。"""

import tempfile
import unittest
from pathlib import Path

import pandas as pd

import sys
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from app.processor import list_rule_ids, write_result_to_excel, run_rule


class TestProcessor(unittest.TestCase):
    def test_list_rule_ids_empty_dir(self):
        """空目录应返回空列表。"""
        with tempfile.TemporaryDirectory() as tmp:
            rules_dir = Path(tmp)
            self.assertEqual(list_rule_ids(rules_dir), [])

    def test_list_rule_ids_ignores_init(self):
        """应忽略 __init__.py。"""
        with tempfile.TemporaryDirectory() as tmp:
            rules_dir = Path(tmp)
            (rules_dir / "__init__.py").write_text("", encoding="utf-8")
            (rules_dir / "foo_rule.py").write_text("# rule", encoding="utf-8")
            ids = list_rule_ids(rules_dir)
            self.assertIn("foo_rule", ids)
            self.assertNotIn("__init__", ids)

    def test_write_result_to_excel_dataframe(self):
        """DataFrame 结果应写入「结果」工作表并保留原表。"""
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            excel_path = tmp_path / "source.xlsx"
            df_orig = pd.DataFrame({"A": [1, 2], "B": [3, 4]})
            df_orig.to_excel(excel_path, index=False, sheet_name="原始")
            result_df = pd.DataFrame({"X": [10], "Y": [20]})
            output_dir = tmp_path / "out"
            out_path = write_result_to_excel(str(excel_path), result_df, output_dir)
            self.assertEqual(out_path, output_dir / "source_processed.xlsx")
            self.assertTrue(out_path.exists())
            with pd.ExcelFile(out_path) as xls:
                sheets = xls.sheet_names
                self.assertIn("原始", sheets)
                self.assertIn("结果", sheets)
                result = pd.read_excel(xls, sheet_name="结果")
                self.assertEqual(list(result["X"]), [10])
                self.assertEqual(list(result["Y"]), [20])

    def test_write_result_to_excel_dict_with_deduction_record(self):
        """dict 结果带 deduction_record 时应写入对应工作表。"""
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            excel_path = tmp_path / "canteen.xlsx"
            pd.DataFrame({"姓名": ["张三"]}).to_excel(excel_path, index=False)
            from datetime import date
            result = {
                "deduction_record": pd.DataFrame({"日期": [date(2025, 1, 1)], "金额": [10]}),
            }
            output_dir = tmp_path / "out"
            out_path = write_result_to_excel(str(excel_path), result, output_dir)
            self.assertTrue(out_path.exists())
            with pd.ExcelFile(out_path) as xls:
                self.assertIn("扣缴记录", xls.sheet_names)

    def test_run_rule_missing_module(self):
        """不存在的规则模块应返回 (None, error_message)。"""
        with tempfile.TemporaryDirectory() as tmp:
            rules_dir = Path(tmp)
            result, err = run_rule("nonexistent_rule_xyz", "", rules_dir)
            self.assertIsNone(result)
            self.assertTrue("nonexistent_rule_xyz" in err or "加载" in err or "Error" in err or "error" in err)


if __name__ == "__main__":
    unittest.main()
