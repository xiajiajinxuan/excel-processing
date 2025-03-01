#!/bin/bash

echo "正在创建Python虚拟环境..."

# 创建虚拟环境
python3 -m venv venv

# 检查是否创建成功
if [ ! -f "venv/bin/activate" ]; then
    echo "创建虚拟环境失败！请确保已安装Python并且可以使用venv模块。"
    exit 1
fi

echo "虚拟环境创建成功！"

# 激活虚拟环境
source venv/bin/activate

# 安装依赖
echo "正在安装依赖..."
pip install -r requirements.txt

echo ""
echo "虚拟环境设置完成！您现在可以运行应用程序："
echo "python main.py"
echo ""
echo "要退出虚拟环境，请输入 'deactivate'"

# 保持终端打开
exec $SHELL 