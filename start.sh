#!/bin/bash

echo "🚀 九宫格潜力展示系统启动器"
echo "=================================="

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

# 安装依赖
echo "📦 检查并安装依赖..."
$PYTHON_CMD -m pip install -r requirements.txt --user

# 启动应用
echo "🚀 启动应用..."
echo "💡 提示: 按 Ctrl+C 停止应用"
echo "🌐 应用将在浏览器中自动打开"
echo "----------------------------------"

$PYTHON_CMD -m streamlit run streamlit_app.py