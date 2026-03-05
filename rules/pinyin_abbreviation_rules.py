# -*- coding: utf-8 -*-
"""
拼音首写转换规则：根据【字段名称】列生成【拼音简称】列（全小写，英文保留、中文取拼音首字母），并标记重复。
"""
import pandas as pd
import re

try:
    import pypinyin
except ImportError:
    pypinyin = None

# 用于判断整段是否包含中文（CJK 统一汉字范围），仅在有中文时调用 pypinyin
RE_HAS_CJK = re.compile(r'[\u4e00-\u9fff\u3400-\u4dbf]')

# CJK 码点范围（避免逐字用正则，用整数比较更快）
_CJK_RANGES = ((0x3400, 0x4DBF), (0x4E00, 0x9FFF))


def _is_chinese(char: str) -> bool:
    """判断单个字符是否为中文（用码点比较，比正则快）。"""
    if len(char) != 1:
        return False
    o = ord(char)
    return any(lo <= o <= hi for lo, hi in _CJK_RANGES)


def _segment_to_abbr(segment: str) -> str:
    """
    将一段字符串转为拼音简称：中文取拼音首字母，英文/数字转小写并保留，其他字符忽略。
    整段一次性调用 pypinyin，避免逐字调用。
    """
    if not segment or not isinstance(segment, str):
        return ''
    # 无中文时直接处理，不调用 pypinyin
    if not RE_HAS_CJK.search(segment):
        return ''.join(c.lower() for c in segment if c.isalnum())
    # 整段一次转拼音首字母，再按规则拼接
    py_list = pypinyin.lazy_pinyin(segment, style=pypinyin.Style.FIRST_LETTER)
    result = []
    for i, char in enumerate(segment):
        if _is_chinese(char):
            s = py_list[i].lower() if i < len(py_list) else char
            result.append(s)
        elif char.isalnum():
            result.append(char.lower())
    return ''.join(result)


def _field_name_to_abbr(name) -> str:
    """
    将字段名称转为拼音简称：按下划线分段，每段分别转换后再用下划线连接，结果全小写。
    """
    if name is None or (isinstance(name, float) and pd.isna(name)):
        return ''
    s = str(name).strip()
    if not s:
        return ''
    segments = s.split('_')
    return '_'.join(_segment_to_abbr(seg) for seg in segments)


def process(data_df, **kwargs):
    """
    读取【字段名称】列，新建【拼音简称】列，为每行生成对应的全小写拼音简称。

    参数:
        data_df (pandas.DataFrame): 至少包含【字段名称】列的 DataFrame
        kwargs: 保留给主程序使用（如 excel_file）

    返回:
        pandas.DataFrame: 原表所有列 + 【拼音简称】列 + 【重复】列（拼音简称重复时为「是」），供主程序写入「结果」工作表。
    """
    if data_df is None or not isinstance(data_df, pd.DataFrame):
        raise ValueError("未提供有效的 DataFrame")

    required_col = '字段名称'
    if required_col not in data_df.columns:
        raise ValueError(f"缺少必需列：{required_col}")

    if pypinyin is None:
        raise ValueError("请先安装 pypinyin：pip install pypinyin -i https://mirrors.aliyun.com/pypi/simple/")

    df = data_df.copy()
    # 仅对唯一的【字段名称】计算简称再映射，重复行不再重复计算
    unique_names = df['字段名称'].dropna().unique()
    abbr_map = {name: _field_name_to_abbr(name) for name in unique_names}
    df['拼音简称'] = df['字段名称'].map(abbr_map).fillna('')
    # 标记重复：若该行拼音简称在表中出现多于一次，则【重复】为「是」
    dup_mask = df['拼音简称'].duplicated(keep=False)
    df['重复'] = dup_mask.map({True: '是', False: '否'})
    return df


def get_rule_info():
    """返回规则的描述信息。"""
    return {
        "name": "拼音首写转换规则",
        "description": "读取【字段名称】列，新建【拼音简称】列，生成全小写拼音首字母简称；英文/数字保留并小写，中文取拼音首字母，按下划线分段转换。",
        "version": "1.0",
        "author": "系统",
    }
