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

    def test_list_rule_ids_subdir_only(self):
        """仅识别子目录形式 <name>/<name>.py，不识别顶层 .py。"""
        with tempfile.TemporaryDirectory() as tmp:
            rules_dir = Path(tmp)
            (rules_dir / "foo_rule" / "foo_rule.py").parent.mkdir(parents=True, exist_ok=True)
            (rules_dir / "foo_rule" / "foo_rule.py").write_text("# rule", encoding="utf-8")
            ids = list_rule_ids(rules_dir)
            self.assertIn("foo_rule", ids)
            (rules_dir / "flat_rule.py").write_text("# flat", encoding="utf-8")
            ids2 = list_rule_ids(rules_dir)
            self.assertIn("foo_rule", ids2)
            self.assertNotIn("flat_rule", ids2)

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

    def test_write_result_to_excel_dict_with_mapping(self):
        """dict 结果应默认使用 key 作为表名，可通过映射重命名。"""
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            excel_path = tmp_path / "canteen.xlsx"
            pd.DataFrame({"姓名": ["张三"]}).to_excel(excel_path, index=False)
            from datetime import date
            result = {
                "deduction_record": pd.DataFrame({"日期": [date(2025, 1, 1)], "金额": [10]}),
            }
            output_dir = tmp_path / "out"

            # 默认行为：使用 key 作为工作表名称
            out_path = write_result_to_excel(str(excel_path), result, output_dir)
            self.assertTrue(out_path.exists())
            with pd.ExcelFile(out_path) as xls:
                self.assertIn("deduction_record", xls.sheet_names)

            # 通过 sheet_mapping 映射为自定义表名
            out_path2 = write_result_to_excel(
                str(excel_path),
                result,
                output_dir,
                sheet_mapping={"deduction_record": "扣缴记录"},
            )
            self.assertTrue(out_path2.exists())
            with pd.ExcelFile(out_path2) as xls:
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
