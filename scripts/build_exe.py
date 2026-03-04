#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Excel 数据处理工具 - exe 打包脚本

使用 PyInstaller 根据 excel_tool.spec 打包为单文件 exe。
建议在项目虚拟环境中执行：python scripts/build_exe.py
"""

import os
import shutil
import subprocess
import sys
from datetime import datetime
from pathlib import Path


# 脚本所在目录与项目根目录
SCRIPT_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = SCRIPT_DIR.parent
SPEC_FILE = PROJECT_ROOT / "excel_tool.spec"
DIST_DIR = PROJECT_ROOT / "dist"
BUILD_DIR = PROJECT_ROOT / "build"
EXE_NAME = "Excel数据处理工具.exe"


def ensure_project_root():
    """确保当前工作目录为项目根目录。"""
    os.chdir(PROJECT_ROOT)
    return PROJECT_ROOT


def check_pyinstaller():
    """检查 PyInstaller 是否可用。"""
    try:
        result = subprocess.run(
            [sys.executable, "-m", "PyInstaller", "--version"],
            capture_output=True,
            text=True,
            check=False,
            cwd=PROJECT_ROOT,
        )
        if result.returncode != 0:
            return False, "PyInstaller 未安装或无法运行"
        return True, (result.stdout or result.stderr or "").strip()
    except Exception as e:
        return False, str(e)


def clean_build_dirs():
    """清理 dist 和 build 目录。"""
    removed = []
    for d in (DIST_DIR, BUILD_DIR):
        if d.is_dir():
            shutil.rmtree(d, ignore_errors=True)
            removed.append(d.name)
    return removed


def run_pyinstaller(clean=True):
    """执行 PyInstaller 打包。"""
    if not SPEC_FILE.is_file():
        print(f"[错误] 未找到 spec 文件: {SPEC_FILE}")
        return False

    cmd = [sys.executable, "-m", "PyInstaller", str(SPEC_FILE)]
    if clean:
        cmd.append("--clean")

    print("[打包] 执行: " + " ".join(cmd))
    result = subprocess.run(cmd, cwd=PROJECT_ROOT)
    return result.returncode == 0


def print_result(success):
    """打印打包结果与输出文件信息。"""
    exe_path = DIST_DIR / EXE_NAME
    if success and exe_path.is_file():
        size_mb = exe_path.stat().st_size / (1024 * 1024)
        mtime = exe_path.stat().st_mtime
        mtime_str = datetime.fromtimestamp(mtime).strftime("%Y-%m-%d %H:%M:%S")
        print(f"输出文件: {exe_path}")
        print(f"文件大小: {size_mb:.2f} MB")
        print(f"构建时间: {mtime_str}")
    elif success:
        print(f"[警告] 未找到输出文件: {exe_path}")
    else:
        print("[错误] 打包失败，请查看上方错误信息。")


def main():
    """主流程：检查环境、清理、打包、输出结果。"""
    print("========================================")
    print("Excel 数据处理工具 - exe 打包")
    print("========================================")
    print(f"项目目录: {PROJECT_ROOT}")
    print()

    ensure_project_root()

    ok, msg = check_pyinstaller()
    if not ok:
        print(f"[错误] {msg}")
        print("请先安装: pip install pyinstaller")
        return 1
    print(f"[环境] PyInstaller: {msg}")
    print()

    removed = clean_build_dirs()
    if removed:
        print(f"[清理] 已删除: {', '.join(removed)}")
    print()

    success = run_pyinstaller(clean=True)
    print()
    print("========================================")
    if success:
        print("打包完成！")
    else:
        print("打包失败。")
    print("========================================")
    print()
    print_result(success)
    return 0 if success else 1


if __name__ == "__main__":
    sys.exit(main())
