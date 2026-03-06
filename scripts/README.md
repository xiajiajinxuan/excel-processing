# 脚本目录

构建、打包与运维相关脚本。

## 文件说明

| 文件 | 说明 |
|------|------|
| **build_exe.py** | 使用 PyInstaller 按 `excel_tool.spec` 打包为单文件 exe。建议在项目根目录执行：`python scripts/build_exe.py` 或先激活虚拟环境后执行。 |
| **build.bat** | Windows 批处理一键打包（调用 PyInstaller）。 |
| **build.ps1** | PowerShell 一键打包脚本。 |

## 使用说明

- 打包前请确保已安装依赖：`pip install -r requirements.txt`
- 打包产物输出到项目根目录的 `dist/`，可执行文件名为「Excel数据处理工具.exe」。
