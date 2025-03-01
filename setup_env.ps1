# PowerShell脚本 - 创建Python虚拟环境
Write-Host "正在创建Python虚拟环境..." -ForegroundColor Cyan

# 创建虚拟环境
python -m venv venv

# 检查是否创建成功
if (-not (Test-Path "venv\Scripts\Activate.ps1")) {
    Write-Host "创建虚拟环境失败！请确保已安装Python并且可以使用venv模块。" -ForegroundColor Red
    Read-Host "按Enter键退出"
    exit 1
}

Write-Host "虚拟环境创建成功！" -ForegroundColor Green

# 激活虚拟环境
& .\venv\Scripts\Activate.ps1

# 安装依赖
Write-Host "正在安装依赖..." -ForegroundColor Cyan
pip install pyinstaller==5.9.0

Write-Host ""
Write-Host "虚拟环境设置完成！您现在可以运行应用程序：" -ForegroundColor Green
Write-Host "python main.py" -ForegroundColor Yellow
Write-Host ""
Write-Host "要退出虚拟环境，请输入 'deactivate'" -ForegroundColor Cyan

# 保持PowerShell会话打开
Write-Host "按Enter键退出此脚本，但保持虚拟环境激活" -ForegroundColor Gray
Read-Host 