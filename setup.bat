@echo off
chcp 65001 >nul
title Excel to PPT Generator - 安裝設定

echo.
echo ╔══════════════════════════════════════════════════════════╗
echo ║         📊 Excel to PPT Generator - 初始設定            ║
echo ╚══════════════════════════════════════════════════════════╝
echo.

cd /d "%~dp0"

:: 檢查 Python
python --version 2>nul
if errorlevel 1 (
    echo [錯誤] 找不到 Python，請先安裝 Python 3.10+
    echo        下載: https://www.python.org/downloads/
    pause
    exit /b 1
)

echo ✓ Python 已安裝
echo.

:: 建立虛擬環境
if not exist "venv" (
    echo [1/3] 建立虛擬環境...
    python -m venv venv
    echo ✓ 虛擬環境建立完成
) else (
    echo ✓ 虛擬環境已存在
)

:: 啟動虛擬環境
call venv\Scripts\activate.bat

:: 安裝依賴
echo.
echo [2/3] 安裝依賴套件 (可能需要幾分鐘)...
pip install -r requirements.txt -q

echo ✓ 依賴套件安裝完成
echo.

:: 建立必要資料夾
echo [3/3] 建立資料夾...
if not exist "uploads" mkdir uploads
if not exist "outputs" mkdir outputs
if not exist "temp_charts" mkdir temp_charts

echo ✓ 資料夾建立完成
echo.

echo ══════════════════════════════════════════════════════════
echo   🎉 安裝完成！
echo.
echo   執行 start_server.bat 啟動服務
echo ══════════════════════════════════════════════════════════
echo.
pause




