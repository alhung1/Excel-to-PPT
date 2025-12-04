@echo off
chcp 65001 >nul
echo.
echo 正在停止 Excel to PPT Generator 服務...
echo.

:: 尋找並終止 uvicorn 程序
taskkill /f /im python.exe /fi "WINDOWTITLE eq Excel to PPT Generator" 2>nul

:: 尋找佔用 8000 port 的程序
for /f "tokens=5" %%a in ('netstat -ano ^| findstr :8000 ^| findstr LISTENING') do (
    echo 終止 PID: %%a
    taskkill /f /pid %%a 2>nul
)

echo.
echo ✓ 服務已停止
timeout /t 2 >nul

