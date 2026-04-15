@echo off
chcp 65001 >nul
title Excel to PPT Generator

echo.
echo ╔══════════════════════════════════════════════════════════╗
echo ║         📊 Excel to PPT Generator v6.0                  ║
echo ╠══════════════════════════════════════════════════════════╣
echo ║  啟動中... 請稍候                                        ║
echo ╚══════════════════════════════════════════════════════════╝
echo.

:: 切換到專案目錄
cd /d "%~dp0"

:: 檢查虛擬環境是否存在
if not exist "venv\Scripts\activate.bat" (
    echo [錯誤] 找不到虛擬環境，請先執行 setup.bat 安裝
    pause
    exit /b 1
)

:: 啟動虛擬環境
call venv\Scripts\activate.bat

:: 檢查依賴套件
python -c "import fastapi" 2>nul
if errorlevel 1 (
    echo [資訊] 正在安裝依賴套件...
    pip install -r requirements.txt -q
)

echo.
echo ✓ 環境準備完成
echo.

:: 防呆：8000 已被占用時，瀏覽器常會開到別的程式 → 預設中止啟動
:: 進階：若你確定要共用連接埠，可先 set SKIP_PORT_CHECK=1 再執行本腳本
if not defined SKIP_PORT_CHECK (
    netstat -ano | findstr ":8000" | findstr "LISTENING" >nul
    if not errorlevel 1 (
        echo [錯誤] 連接埠 8000 已有程式在監聽，為避免開錯網頁已中止啟動。
        echo        請先執行 stop_server.bat，或關閉占用 8000 的程式後再試。
        echo        ^(進階：set SKIP_PORT_CHECK=1 可略過此檢查^)
        echo.
        pause
        exit /b 1
    )
)

echo ══════════════════════════════════════════════════════════
echo   🌐 服務啟動中...
echo   📍 網址: http://127.0.0.1:8000  ^(請勿用 localhost，避免 IPv6 誤連^)
echo   🛑 關閉: 按 Ctrl+C 或直接關閉此視窗
echo ══════════════════════════════════════════════════════════
echo.

:: 背景輪詢 /api/health 成功後才開瀏覽器（不依賴固定延遲；且驗證是本專案 API）
start "Open browser when ready" /MIN python "%~dp0scripts\wait_and_open_browser.py" 127.0.0.1 8000

:: 啟動 FastAPI 服務（只綁 IPv4，與上方網址一致）
python -m uvicorn app.main:app --host 127.0.0.1 --port 8000

:: 如果服務結束，暫停顯示
echo.
echo 服務已停止
pause