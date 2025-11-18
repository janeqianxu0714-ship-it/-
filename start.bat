@echo off
chcp 65001 >nul
echo 🚀 九宫格潜力展示系统启动器
echo ==================================

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

REM 安装依赖
echo 📦 检查并安装依赖...
%PYTHON_CMD% -m pip install -r requirements.txt --user

REM 启动应用
echo 🚀 启动应用...
echo 💡 提示: 按 Ctrl+C 停止应用
echo 🌐 应用将在浏览器中自动打开
echo ----------------------------------

%PYTHON_CMD% -m streamlit run streamlit_app.py

pause