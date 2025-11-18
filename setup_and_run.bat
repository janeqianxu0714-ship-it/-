@echo off
chcp 65001 >nul
echo 🚀 九宫格潜力展示系统 - 完整安装启动器
echo ==========================================

REM 检查Python命令
python --version >nul 2>&1
if %errorlevel% == 0 (
    set PYTHON_CMD=python
    goto :found_python
)

python3 --version >nul 2>&1
if %errorlevel% == 0 (
    set PYTHON_CMD=python3
    goto :found_python
)

py --version >nul 2>&1
if %errorlevel% == 0 (
    set PYTHON_CMD=py
    goto :found_python
)

echo ❌ 错误: 找不到Python命令
echo 请先安装Python 3.7+
pause
exit /b 1

:found_python
echo 📦 使用Python命令: %PYTHON_CMD%

REM 检查streamlit_app.py是否存在
if not exist "streamlit_app.py" (
    echo ❌ 错误: 找不到 streamlit_app.py 文件
    echo 请确保在项目根目录运行此脚本
    pause
    exit /b 1
)

REM 创建虚拟环境
if not exist "venv" (
    echo 🔧 创建虚拟环境...
    %PYTHON_CMD% -m venv venv
    if %errorlevel% neq 0 (
        echo ❌ 创建虚拟环境失败
        echo 请确保Python安装正确
        pause
        exit /b 1
    )
)

REM 激活虚拟环境
echo 🔄 激活虚拟环境...
call venv\Scripts\activate.bat

REM 升级pip
echo ⬆️  升级pip...
python -m pip install --upgrade pip

REM 安装依赖
echo 📦 安装依赖...
pip install -r requirements.txt

REM 启动应用
echo 🚀 启动应用...
echo 💡 提示: 按 Ctrl+C 停止应用
echo 🌐 应用将在浏览器中自动打开
echo ----------------------------------

streamlit run streamlit_app.py

pause