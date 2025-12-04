@echo off
chcp 65001 >nul
echo.
echo 正在建立桌面捷徑...

:: 取得桌面路徑
set "DESKTOP=%USERPROFILE%\Desktop"

:: 建立 VBS 腳本來建立捷徑
echo Set oWS = WScript.CreateObject("WScript.Shell") > "%TEMP%\CreateShortcut.vbs"
echo sLinkFile = "%DESKTOP%\Excel to PPT Generator.lnk" >> "%TEMP%\CreateShortcut.vbs"
echo Set oLink = oWS.CreateShortcut(sLinkFile) >> "%TEMP%\CreateShortcut.vbs"
echo oLink.TargetPath = "%~dp0start_server.bat" >> "%TEMP%\CreateShortcut.vbs"
echo oLink.WorkingDirectory = "%~dp0" >> "%TEMP%\CreateShortcut.vbs"
echo oLink.Description = "Excel to PPT Generator" >> "%TEMP%\CreateShortcut.vbs"
echo oLink.IconLocation = "%SystemRoot%\System32\SHELL32.dll,21" >> "%TEMP%\CreateShortcut.vbs"
echo oLink.Save >> "%TEMP%\CreateShortcut.vbs"

:: 執行 VBS 腳本
cscript //nologo "%TEMP%\CreateShortcut.vbs"

:: 刪除暫存 VBS
del "%TEMP%\CreateShortcut.vbs"

echo.
echo ✓ 桌面捷徑已建立: "Excel to PPT Generator"
echo.
timeout /t 3 >nul

