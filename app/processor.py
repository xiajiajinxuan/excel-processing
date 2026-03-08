# -*- coding: utf-8 -*-
"""规则执行与结果写入：加载规则模块、执行 process、写入 Excel。与 GUI 解耦，便于单测。"""

import importlib.util
import sys
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
    规则模块从 rules_dir / rule_id / rule_id.py 加载。
    :return: 成功时 (result, elapsed_seconds)，失败时 (None, error_message)。
    """
    rules_dir = Path(rules_dir)
    rule_py = rules_dir / rule_id / f"{rule_id}.py"
    if not rule_py.exists():
        return None, f"规则模块不存在: {rule_py}"

    try:
        spec = importlib.util.spec_from_file_location(
            f"_rule_{rule_id}", rule_py, submodule_search_locations=[str(rules_dir / rule_id)]
        )
        if spec is None or spec.loader is None:
            return None, "加载规则模块失败: 无法创建 spec"
        rule_module = importlib.util.module_from_spec(spec)
        sys.modules[spec.name] = rule_module
        spec.loader.exec_module(rule_module)
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
    """列出 rules 目录下所有规则 ID（仅子目录形式：rules_dir/<name>/<name>.py）。"""
    rules_dir = Path(rules_dir)
    if not rules_dir.exists():
        return []
    ids = []
    for sub in rules_dir.iterdir():
        if sub.is_dir() and not sub.name.startswith("."):
            py_file = sub / f"{sub.name}.py"
            if py_file.exists():
                ids.append(sub.name)
    return sorted(ids)


def get_default_template_for_rule(rules_dir: Path, rule_id: str) -> str | None:
    """返回规则 doc/template 目录下第一个 .xlsx 文件名，若无则返回 None。"""
    template_dir = Path(rules_dir) / rule_id / "doc" / "template"
    if not template_dir.is_dir():
        return None
    for p in sorted(template_dir.iterdir()):
        if p.is_file() and p.suffix.lower() == ".xlsx":
            return p.name
    return None
