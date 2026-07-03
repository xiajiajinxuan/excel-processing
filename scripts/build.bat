@echo off
chcp 65001 >nul
echo ========================================
echo Excel数据处理工具 - 打包脚本
echo ========================================
echo.

REM 切换到项目根目录（脚本在scripts目录下）
cd /d "%~dp0\.."

REM 优先使用项目虚拟环境
set "PYTHON=python"
if exist "venv\Scripts\python.exe" (
    set "PYTHON=venv\Scripts\python.exe"
    echo [环境] 使用项目虚拟环境: venv\Scripts\python.exe
)

REM 检查Python是否安装
%PYTHON% --version >nul 2>&1
if errorlevel 1 (
    echo [错误] 未找到Python，请先安装Python
    pause
    exit /b 1
)

REM 检查PyInstaller是否安装
%PYTHON% -m pip show pyinstaller >nul 2>&1
if errorlevel 1 (
    echo [警告] PyInstaller未安装，正在安装...
    %PYTHON% -m pip install pyinstaller -i https://mirrors.aliyun.com/pypi/simple/
    if errorlevel 1 (
        echo [错误] PyInstaller安装失败
        pause
        exit /b 1
    )
)

REM 检查必要的依赖
echo [1/3] 检查依赖...
%PYTHON% -m pip show PyQt6 >nul 2>&1
if errorlevel 1 (
    echo [错误] PyQt6 未安装，请先执行: %PYTHON% -m pip install -r requirements.txt
    pause
    exit /b 1
)

%PYTHON% -m pip show pandas >nul 2>&1
if errorlevel 1 (
    echo [警告] pandas未安装，正在安装...
    %PYTHON% -m pip install pandas -i https://mirrors.aliyun.com/pypi/simple/
)

%PYTHON% -m pip show openpyxl >nul 2>&1
if errorlevel 1 (
    echo [警告] openpyxl未安装，正在安装...
    %PYTHON% -m pip install openpyxl -i https://mirrors.aliyun.com/pypi/simple/
)

REM 清理旧的构建文件
echo [2/3] 清理旧的构建文件...
if exist "dist" (
    rmdir /s /q "dist"
    echo    - 已删除 dist 目录
)
if exist "build" (
    rmdir /s /q "build"
    echo    - 已删除 build 目录
)

REM 执行打包
echo [3/3] 开始打包...
echo.
%PYTHON% -m PyInstaller excel_tool.spec --clean

if errorlevel 1 (
    echo.
    echo [错误] 打包失败，请检查错误信息
    pause
    exit /b 1
)

echo.
echo ========================================
echo 打包完成！
echo ========================================
echo.
echo 输出目录: dist\Excel数据处理工具\
echo 可执行文件: dist\Excel数据处理工具\Excel数据处理工具.exe
if exist "dist\Excel数据处理工具\Excel数据处理工具.exe" (
    for %%A in ("dist\Excel数据处理工具\Excel数据处理工具.exe") do (
        echo exe 大小: %%~zA 字节
    )
)
echo.
pause

