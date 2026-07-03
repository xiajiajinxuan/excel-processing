# -*- coding: utf-8 -*-
"""maturity_balance 规则测试。"""

import unittest
from pathlib import Path

import pandas as pd

import sys

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from app.processor import run_rule


class TestMaturityBalanceRule(unittest.TestCase):
    def setUp(self) -> None:
        self.rules_dir = Path(__file__).resolve().parent.parent / "rules"

    def test_maturity_balance_basic_cumsum_no_gl_date(self):
        """无【GL日期】列时，应保持原行顺序累计。"""
        from tempfile import TemporaryDirectory

        with TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            excel_path = tmp_path / "source.xlsx"

            df = pd.DataFrame(
                {
                    "本币借": [100, 50, 0],
                    "本币贷": [0, 20, 10],
                }
            )
            df.to_excel(excel_path, index=False)

            result, err_or_elapsed = run_rule(
                "maturity_balance",
                str(excel_path),
                self.rules_dir,
            )
            # run_rule 在成功时返回 (result, elapsed_seconds)
            self.assertIsNotNone(result)
            self.assertFalse(isinstance(err_or_elapsed, str))

            self.assertIn("本币到期余额", result.columns)
            expected = [100 - 0, 100 - 0 + 50 - 20, 100 - 0 + 50 - 20 + 0 - 10]
            self.assertEqual(list(result["本币到期余额"]), expected)

    def test_maturity_balance_with_default_gl_date_sort(self):
        """存在【GL日期】且未指定 sort_by 时，应按【GL日期】排序后累计。"""
        from tempfile import TemporaryDirectory
        from datetime import date

        with TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            excel_path = tmp_path / "source_gl.xlsx"

            df = pd.DataFrame(
                {
                    "GL日期": [date(2025, 1, 3), date(2025, 1, 1), date(2025, 1, 2)],
                    "本币借": [0, 100, 50],
                    "本币贷": [10, 0, 20],
                }
            )
            df.to_excel(excel_path, index=False)

            result, err_or_elapsed = run_rule(
                "maturity_balance",
                str(excel_path),
                self.rules_dir,
            )
            self.assertIsNotNone(result)
            self.assertFalse(isinstance(err_or_elapsed, str))
            self.assertIn("本币到期余额", result.columns)

            # 按 GL日期 升序后的借贷差额：行顺序为 1 日、2 日、3 日
            # 1 日: 100 - 0 = 100
            # 2 日: 50 - 20 = 30 -> 累计 130
            # 3 日: 0 - 10 = -10 -> 累计 120
            expected = [100, 130, 120]
            self.assertEqual(list(result["本币到期余额"]), expected)

    def test_maturity_balance_with_sort_by(self):
        """验证按排序字段排序后再进行累计。"""
        from tempfile import TemporaryDirectory
        from datetime import date

        with TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            excel_path = tmp_path / "source_sorted.xlsx"

            df = pd.DataFrame(
                {
                    "日期": [date(2025, 1, 3), date(2025, 1, 1), date(2025, 1, 2)],
                    "本币借": [0, 100, 50],
                    "本币贷": [10, 0, 20],
                }
            )
            df.to_excel(excel_path, index=False)

            # 直接调规则模块更方便传递 sort_by，这里使用 importlib 加载
            import importlib.util

            rule_py = self.rules_dir / "maturity_balance" / "maturity_balance.py"
            spec = importlib.util.spec_from_file_location("maturity_balance_rule", rule_py)
            module = importlib.util.module_from_spec(spec)
            assert spec.loader is not None
            spec.loader.exec_module(module)

            result = module.process(pd.read_excel(excel_path), sort_by="日期", excel_file=str(excel_path))
            self.assertIsInstance(result, pd.DataFrame)
            self.assertIn("本币到期余额", result.columns)

            # 按日期升序后的借贷差额：行顺序为 1 日、2 日、3 日
            # 1 日: 100 - 0 = 100
            # 2 日: 50 - 20 = 30 -> 累计 130
            # 3 日: 0 - 10 = -10 -> 累计 120
            expected = [100, 130, 120]
            self.assertEqual(list(result["本币到期余额"]), expected)

    def test_maturity_balance_missing_columns(self):
        """缺少必要列时应返回 error。"""
        import importlib.util

        rule_py = self.rules_dir / "maturity_balance" / "maturity_balance.py"
        spec = importlib.util.spec_from_file_location("maturity_balance_rule2", rule_py)
        module = importlib.util.module_from_spec(spec)
        assert spec.loader is not None
        spec.loader.exec_module(module)

        df = pd.DataFrame({"本币借": [100]})
        result = module.process(df, excel_file="dummy.xlsx")
        self.assertIsInstance(result, dict)
        self.assertIn("error", result)


if __name__ == "__main__":
    unittest.main()

