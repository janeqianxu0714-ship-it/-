#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
九宫格潜力展示系统 - 快速启动脚本
"""

import subprocess
import sys
import os
import webbrowser
import time
import threading

def check_dependencies():
    """检查依赖是否安装"""
    try:
        import streamlit
        import pandas
        import openpyxl
        return True
    except ImportError as e:
        print(f"❌ 缺少依赖: {e}")
        print("📦 正在安装依赖...")
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", "-r", "requirements.txt"])
            print("✅ 依赖安装完成!")
            return True
        except subprocess.CalledProcessError:
            print("❌ 依赖安装失败，请手动运行: pip install -r requirements.txt")
            return False

def open_browser(url, delay=3):
    """延迟打开浏览器"""
    time.sleep(delay)
    try:
        webbrowser.open(url)
        print(f"🌐 浏览器已打开: {url}")
    except Exception as e:
        print(f"⚠️  无法自动打开浏览器: {e}")
        print(f"请手动访问: {url}")

def main():
    """主函数"""
    print("🚀 九宫格潜力展示系统启动器")
    print("=" * 50)
    
    # 检查streamlit_app.py是否存在
    if not os.path.exists("streamlit_app.py"):
        print("❌ 错误: 找不到 streamlit_app.py 文件")
        print("请确保在项目根目录运行此脚本")
        input("按回车键退出...")
        return
    
    # 检查依赖
    if not check_dependencies():
        input("按回车键退出...")
        return
    
    print("📊 正在启动应用...")
    
    # 在后台线程中打开浏览器
    browser_thread = threading.Thread(target=open_browser, args=("http://localhost:8501", 3))
    browser_thread.daemon = True
    browser_thread.start()
    
    try:
        # 启动Streamlit应用
        cmd = [sys.executable, "-m", "streamlit", "run", "streamlit_app.py", "--server.headless", "true"]
        print("✅ 应用启动成功!")
        print("💡 提示: 按 Ctrl+C 停止应用")
        print("-" * 50)
        
        subprocess.run(cmd)
        
    except KeyboardInterrupt:
        print("\n👋 应用已停止")
    except FileNotFoundError:
        print("❌ 错误: 找不到 streamlit 命令")
        print("请先安装 streamlit: pip install streamlit")
        input("按回车键退出...")
    except Exception as e:
        print(f"❌ 启动失败: {e}")
        input("按回车键退出...")

if __name__ == "__main__":
    main()