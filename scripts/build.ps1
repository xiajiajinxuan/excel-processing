# Excel数据处理工具 - 打包脚本 (PowerShell版本)
# 编码: UTF-8

$ErrorActionPreference = "Stop"

# 切换到项目根目录（脚本在scripts目录下）
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectRoot = Split-Path -Parent $scriptDir
Set-Location $projectRoot

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Excel数据处理工具 - 打包脚本" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "项目目录: $projectRoot" -ForegroundColor Gray
Write-Host ""

# 检查Python是否安装
try {
    $pythonVersion = python --version 2>&1
    Write-Host "[1/4] 检查Python环境..." -ForegroundColor Yellow
    Write-Host "   Python版本: $pythonVersion" -ForegroundColor Green
} catch {
    Write-Host "[错误] 未找到Python，请先安装Python" -ForegroundColor Red
    Read-Host "按Enter键退出"
    exit 1
}

# 检查并安装PyInstaller
Write-Host "[2/4] 检查PyInstaller..." -ForegroundColor Yellow
$pyinstallerInstalled = python -m pip show pyinstaller 2>&1
if ($LASTEXITCODE -ne 0) {
    Write-Host "   PyInstaller未安装，正在安装..." -ForegroundColor Yellow
    python -m pip install pyinstaller -i https://mirrors.aliyun.com/pypi/simple/
    if ($LASTEXITCODE -ne 0) {
        Write-Host "[错误] PyInstaller安装失败" -ForegroundColor Red
        Read-Host "按Enter键退出"
        exit 1
    }
    Write-Host "   PyInstaller安装成功" -ForegroundColor Green
} else {
    Write-Host "   PyInstaller已安装" -ForegroundColor Green
}

# 检查并安装必要的依赖
Write-Host "[3/4] 检查依赖..." -ForegroundColor Yellow
$dependencies = @("pandas", "openpyxl", "numpy", "PyYAML")

foreach ($dep in $dependencies) {
    $installed = python -m pip show $dep 2>&1
    if ($LASTEXITCODE -ne 0) {
        Write-Host "   $dep 未安装，正在安装..." -ForegroundColor Yellow
        python -m pip install $dep -i https://mirrors.aliyun.com/pypi/simple/
        if ($LASTEXITCODE -eq 0) {
            Write-Host "   $dep 安装成功" -ForegroundColor Green
        } else {
            Write-Host "   [警告] $dep 安装失败，但继续打包..." -ForegroundColor Yellow
        }
    } else {
        Write-Host "   $dep 已安装" -ForegroundColor Green
    }
}

# 清理旧的构建文件
Write-Host "[4/4] 清理旧的构建文件..." -ForegroundColor Yellow
if (Test-Path "dist") {
    Remove-Item -Recurse -Force "dist"
    Write-Host "   - 已删除 dist 目录" -ForegroundColor Green
}
if (Test-Path "build") {
    Remove-Item -Recurse -Force "build"
    Write-Host "   - 已删除 build 目录" -ForegroundColor Green
}

# 执行打包
Write-Host ""
Write-Host "开始打包..." -ForegroundColor Cyan
Write-Host ""

python -m PyInstaller excel_tool.spec --clean

if ($LASTEXITCODE -ne 0) {
    Write-Host ""
    Write-Host "[错误] 打包失败，请检查错误信息" -ForegroundColor Red
    Read-Host "按Enter键退出"
    exit 1
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "打包完成！" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

$exePath = "dist\Excel数据处理工具.exe"
if (Test-Path $exePath) {
    $fileInfo = Get-Item $exePath
    $fileSizeMB = [math]::Round($fileInfo.Length / 1MB, 2)
    Write-Host "输出文件位置: $exePath" -ForegroundColor Green
    Write-Host "文件大小: $fileSizeMB MB ($($fileInfo.Length) 字节)" -ForegroundColor Green
    Write-Host "构建时间: $($fileInfo.LastWriteTime)" -ForegroundColor Green
} else {
    Write-Host "[警告] 未找到输出文件" -ForegroundColor Yellow
}

Write-Host ""
Read-Host "按Enter键退出"

