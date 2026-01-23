@echo off
chcp 65001 >nul
title Editable Charts Test - 可編輯圖表測試

echo ============================================================
echo   Editable Charts Test - 可編輯圖表測試
echo ============================================================
echo.
echo 此工具會測試將 Excel 圖表插入 PowerPoint 的三種方式:
echo   1. Embedded (嵌入式) - 圖表可編輯，資料存在 PPT 中
echo   2. Linked (連結式)   - 圖表可編輯，連結到 Excel 來源
echo   3. Image (圖片)      - 靜態圖片，無法編輯 (目前的方式)
echo.
echo ============================================================
echo.

:: Check if a file was dragged onto the batch file
if "%~1"=="" (
    echo 請輸入 Excel 檔案路徑，或直接將 Excel 檔案拖放到此視窗:
    echo.
    set /p EXCEL_FILE="Excel 檔案路徑: "
) else (
    set "EXCEL_FILE=%~1"
)

:: Remove quotes if present
set "EXCEL_FILE=%EXCEL_FILE:"=%"

:: Check if file exists
if not exist "%EXCEL_FILE%" (
    echo.
    echo [錯誤] 找不到檔案: %EXCEL_FILE%
    echo.
    pause
    exit /b 1
)

echo.
echo 使用 Excel 檔案: %EXCEL_FILE%
echo.

:: Find Python in venv
set "PYTHON_EXE=..\venv\Scripts\python.exe"
if not exist "%PYTHON_EXE%" (
    echo [錯誤] 找不到 Python 虛擬環境
    echo 請先執行 setup.bat 安裝
    pause
    exit /b 1
)

echo 開始測試...
echo.

:: Run the test
"%PYTHON_EXE%" test_editable_chart.py --excel "%EXCEL_FILE%"

echo.
echo ============================================================
echo.
echo 測試完成！請檢查產生的 PowerPoint 檔案:
echo.
echo   test_embedded_chart.pptx - 嵌入式 (可編輯)
echo   test_linked_chart.pptx   - 連結式 (可編輯)
echo   test_image_chart.pptx    - 圖片 (不可編輯)
echo.
echo 驗證方式:
echo   1. 開啟 test_embedded_chart.pptx 或 test_linked_chart.pptx
echo   2. 點擊圖表
echo   3. 右鍵應該會看到「編輯資料」選項
echo   4. 雙擊可以編輯圖表內容
echo.
echo ============================================================

pause
