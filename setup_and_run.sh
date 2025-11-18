#!/bin/bash

echo "🚀 九宫格潜力展示系统 - 完整安装启动器"
echo "=========================================="

# 检查Python命令
if command -v python3 &> /dev/null; then
    PYTHON_CMD="python3"
elif command -v python &> /dev/null; then
    PYTHON_CMD="python"
else
    echo "❌ 错误: 找不到Python命令"
    echo "请先安装Python 3.7+"
    exit 1
fi

echo "📦 使用Python命令: $PYTHON_CMD"

# 检查streamlit_app.py是否存在
if [ ! -f "streamlit_app.py" ]; then
    echo "❌ 错误: 找不到 streamlit_app.py 文件"
    echo "请确保在项目根目录运行此脚本"
    exit 1
fi

# 创建虚拟环境
if [ ! -d "venv" ]; then
    echo "🔧 创建虚拟环境..."
    $PYTHON_CMD -m venv venv
    if [ $? -ne 0 ]; then
        echo "❌ 创建虚拟环境失败"
        echo "请确保已安装python3-venv: sudo apt install python3-venv (Ubuntu/Debian)"
        exit 1
    fi
fi

# 激活虚拟环境
echo "🔄 激活虚拟环境..."
source venv/bin/activate

# 升级pip
echo "⬆️  升级pip..."
pip install --upgrade pip

# 安装依赖
echo "📦 安装依赖..."
pip install -r requirements.txt

# 启动应用
echo "🚀 启动应用..."
echo "💡 提示: 按 Ctrl+C 停止应用"
echo "🌐 应用将在浏览器中自动打开"
echo "----------------------------------"

streamlit run streamlit_app.py