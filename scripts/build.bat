@echo off
chcp 65001 >nul
echo ========================================
echo Excel数据处理工具 - 打包脚本
echo ========================================
echo.

REM 切换到项目根目录（脚本在scripts目录下）
cd /d "%~dp0\.."

REM 检查Python是否安装
python --version >nul 2>&1
if errorlevel 1 (
    echo [错误] 未找到Python，请先安装Python
    pause
    exit /b 1
)

REM 检查PyInstaller是否安装
python -m pip show pyinstaller >nul 2>&1
if errorlevel 1 (
    echo [警告] PyInstaller未安装，正在安装...
    python -m pip install pyinstaller -i https://mirrors.aliyun.com/pypi/simple/
    if errorlevel 1 (
        echo [错误] PyInstaller安装失败
        pause
        exit /b 1
    )
)

REM 检查必要的依赖
echo [1/3] 检查依赖...
python -m pip show pandas >nul 2>&1
if errorlevel 1 (
    echo [警告] pandas未安装，正在安装...
    python -m pip install pandas -i https://mirrors.aliyun.com/pypi/simple/
)

python -m pip show openpyxl >nul 2>&1
if errorlevel 1 (
    echo [警告] openpyxl未安装，正在安装...
    python -m pip install openpyxl -i https://mirrors.aliyun.com/pypi/simple/
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
python -m PyInstaller excel_tool.spec --clean

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
echo 输出文件位置: dist\Excel数据处理工具.exe
if exist "dist\Excel数据处理工具.exe" (
    for %%A in ("dist\Excel数据处理工具.exe") do (
        echo 文件大小: %%~zA 字节
    )
)
echo.
pause

