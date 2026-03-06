# -*- coding: utf-8 -*-
"""规则执行与结果写入：加载规则模块、执行 process、写入 Excel。与 GUI 解耦，便于单测。"""

import importlib
import time
from pathlib import Path
from typing import Any

import pandas as pd


def run_rule(
    rule_id: str,
    file_path: str,
    rules_dir: Path,
) -> tuple[Any, float | None] | tuple[None, str]:
    """
    执行指定规则处理 Excel 文件。
    :return: 成功时 (result, elapsed_seconds)，失败时 (None, error_message)。
    """
    try:
        rule_module = importlib.import_module(f"rules.{rule_id}")
    except Exception as e:
        return None, f"加载规则模块失败: {e}"

    if not getattr(rule_module, "process", None):
        return None, "规则模块缺少 process 函数"

    try:
        data_df = pd.read_excel(file_path)
    except Exception as e:
        return None, f"读取 Excel 失败: {e}"

    start = time.perf_counter()
    try:
        result = rule_module.process(data_df, excel_file=file_path)
    except Exception as e:
        return None, f"规则执行出错: {e}"
    elapsed = time.perf_counter() - start
    return result, elapsed


def write_result_to_excel(
    file_path: str,
    result: Any,
    output_dir: Path,
) -> Path:
    """
    将处理结果写入 Excel：保留原表，新增结果工作表。
    :return: 输出文件路径。
    """
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    file_name = Path(file_path).name
    base_name = Path(file_name).stem
    output_file = output_dir / f"{base_name}_processed.xlsx"

    original_dfs = {}
    with pd.ExcelFile(file_path) as xls:
        for sheet_name in xls.sheet_names:
            original_dfs[sheet_name] = pd.read_excel(xls, sheet_name=sheet_name)

    if isinstance(result, dict) and "deduction_record" in result and "日期" in result["deduction_record"].columns:
        result["deduction_record"]["日期"] = (
            pd.to_datetime(result["deduction_record"]["日期"]).dt.strftime("%Y-%m-%d")
        )

    sheet_mapping = {"deduction_record": "扣缴记录", "monthly_summary": "月度汇总"}
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        for sheet_name, df in original_dfs.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)
        if isinstance(result, dict):
            for key, value in result.items():
                if isinstance(value, pd.DataFrame) and not value.empty and key != "error":
                    sheet_name = sheet_mapping.get(key, key)
                    value.to_excel(writer, sheet_name=sheet_name, index=False)
        elif isinstance(result, pd.DataFrame):
            result.to_excel(writer, sheet_name="结果", index=False)

    return output_file


def list_rule_ids(rules_dir: Path) -> list[str]:
    """列出 rules 目录下所有规则模块名（不含 __init__）。"""
    rules_dir = Path(rules_dir)
    if not rules_dir.exists():
        return []
    return [f.stem for f in rules_dir.glob("*.py") if f.stem != "__init__"]
