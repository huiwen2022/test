#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
快速打包腳本 - 簡化版手動打包工具
直接執行即可打包你的Python程式
"""

import os
import shutil
import zipfile
from datetime import datetime
# import sys
# import io
# sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

def quick_package():
    """快速打包函數"""
    app_name = "MyPythonApp"
    
    print(f"🚀 開始打包 {app_name}...")
    
    # 1. 清理並創建打包目錄
    build_dir = f"portable_{app_name}"
    if os.path.exists(build_dir):
        shutil.rmtree(build_dir)
    os.makedirs(build_dir)
    print(f"✅ 創建打包目錄: {build_dir}")
    
    # 2. 複製Python檔案
    python_files = [
        'main.py',
        'form_app.py', 
        'form_validation.py',
        'excel_handler.py'
    ]
    
    print("📁 複製Python檔案...")
    for file in python_files:
        if os.path.exists(file):
            shutil.copy2(file, build_dir)
            print(f"   ✓ {file}")
        else:
            print(f"   ⚠️  找不到 {file}")
    
    # 3. 複製libs資料夾
    if os.path.exists('libs'):
        shutil.copytree('libs', os.path.join(build_dir, 'libs'))
        print("📦 複製libs資料夾")
    else:
        print("⚠️  找不到libs資料夾")
    
    # 4. 創建Windows啟動檔
    batch_content = f'''@echo off
title {app_name}
cd /d "%~dp0"

echo 正在檢查Python環境...
python --version >nul 2>&1
if errorlevel 1 (
    echo ❌ 找不到Python！請先安裝Python 3.7+
    echo 📥 下載網址: https://www.python.org/downloads/
    pause
    exit /b 1
)

echo ✅ Python環境正常
echo 🚀 啟動程式...

set PYTHONPATH=%cd%\\libs;%PYTHONPATH%
python main.py

if errorlevel 1 (
    echo ❌ 程式執行失敗！
    echo 💡 請執行 python check_env.py 檢查環境
)

pause
'''
    
    with open(os.path.join(build_dir, f'{app_name}.bat'), 'w', encoding='utf-8') as f:
        f.write(batch_content)
    print("🖥️  創建Windows啟動器")
    
    # 5. 創建環境檢查腳本
    check_script = '''#!/usr/bin/env python3
import sys
import os

print("=== 環境檢查 ===")
print(f"Python版本: {sys.version}")

# 檢查路徑設置
current_dir = os.path.dirname(os.path.abspath(__file__))
libs_path = os.path.join(current_dir, 'libs')
if libs_path not in sys.path:
    sys.path.insert(0, libs_path)

# 檢查模組
modules = ['openpyxl', 'xlsxwriter', 'tkinter']
for module in modules:
    try:
        __import__(module)
        print(f"✅ {module}")
    except ImportError as e:
        print(f"❌ {module}: {e}")

# 檢查檔案
files = ['main.py', 'form_app.py', 'form_validation.py', 'excel_handler.py']
for file in files:
    if os.path.exists(file):
        print(f"✅ {file}")
    else:
        print(f"❌ {file}")

input("\\n按Enter鍵退出...")
'''
    
    with open(os.path.join(build_dir, 'check_env.py'), 'w', encoding='utf-8') as f:
        f.write(check_script)
    print("🔍 創建環境檢查工具")
    
    # 6. 創建說明檔案
    readme = f'''# {app_name} 使用說明

## 🚀 快速開始
1. 確保電腦已安裝Python 3.7+
2. 雙擊 {app_name}.bat 啟動程式

## 🔧 疑難排解
如果程式無法啟動：
1. 執行 python check_env.py 檢查環境
2. 確認Python已加入系統PATH

## 📥 Python下載
https://www.python.org/downloads/
(安裝時請勾選 "Add Python to PATH")

---
打包時間: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}
'''
    
    with open(os.path.join(build_dir, 'README.txt'), 'w', encoding='utf-8') as f:
        f.write(readme)
    print("📝 創建使用說明")
    
    # 7. 創建ZIP壓縮包
    zip_name = f"{app_name}_portable.zip"
    with zipfile.ZipFile(zip_name, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root, dirs, files in os.walk(build_dir):
            for file in files:
                file_path = os.path.join(root, file)
                arcname = os.path.relpath(file_path, build_dir)
                zipf.write(file_path, f"{app_name}/{arcname}")
    
    print(f"📦 創建ZIP壓縮包: {zip_name}")
    
    # 8. 顯示結果
    zip_size = os.path.getsize(zip_name) / 1024 / 1024  # MB
    print("\n" + "="*50)
    print("🎉 打包完成！")
    print("="*50)
    print(f"📁 打包目錄: {build_dir}")
    print(f"📦 壓縮檔案: {zip_name} ({zip_size:.1f} MB)")
    print("\n🔧 使用方法:")
    print("1. 將ZIP檔案傳送到目標電腦")
    print("2. 解壓縮ZIP檔案") 
    print("3. 雙擊執行 .bat 檔案")
    print("\n💡 注意事項:")
    print("- 目標電腦需要安裝Python 3.7+")
    print("- 如有問題請執行check_env.py檢查")
    print("="*50)

if __name__ == "__main__":
    try:
        quick_package()
    except KeyboardInterrupt:
        print("\n❌ 使用者取消操作")
    except Exception as e:
        print(f"\n❌ 發生錯誤: {e}")
        input("按Enter鍵退出...")