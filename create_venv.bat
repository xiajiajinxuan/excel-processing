@echo off
chcp 65001 > nul
echo 正在创建Python虚拟环境...

:: 创建虚拟环境
python -m venv venv

:: 检查是否创建成功
if not exist venv\Scripts\activate.bat (
    echo 创建虚拟环境失败！请确保已安装Python并且可以使用venv模块。
    pause
    exit /b 1
)

echo 虚拟环境创建成功！

:: 激活虚拟环境
call venv\Scripts\activate.bat

:: 安装依赖
echo 正在安装依赖...
pip install pyinstaller==5.9.0

echo.
echo 虚拟环境设置完成！您现在可以运行应用程序：
echo python main.py
echo.
echo 要退出虚拟环境，请输入 'deactivate'

:: 保持命令窗口打开
cmd /k 