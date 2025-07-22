@echo off
chcp 65001 >nul
cd /d %~dp0

where pythonw >nul 2>nul
if %errorlevel% neq 0 (
    msg * [錯誤] 未偵測到 Python，請先安裝 Python 並加入環境變數。
    exit /b
)

:: 有 pythonw，執行 GUI 隱藏版本
start "" pythonw main2.py
