@echo off
echo ========================================
echo   Word WebDAV Server - Cai dat
echo ========================================

:: Tao shortcut trong Startup folder
set STARTUP=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup
set CURRENT=%~dp0

echo Dang tao auto-start...

:: Tao VBS script de tao shortcut
echo Set oWS = WScript.CreateObject("WScript.Shell") > "%TEMP%\CreateShortcut.vbs"
echo sLinkFile = "%STARTUP%\WordWebDAV.lnk" >> "%TEMP%\CreateShortcut.vbs"
echo Set oLink = oWS.CreateShortcut(sLinkFile) >> "%TEMP%\CreateShortcut.vbs"
echo oLink.TargetPath = "%CURRENT%start-hidden.vbs" >> "%TEMP%\CreateShortcut.vbs"
echo oLink.WorkingDirectory = "%CURRENT%" >> "%TEMP%\CreateShortcut.vbs"
echo oLink.Description = "Word WebDAV Server" >> "%TEMP%\CreateShortcut.vbs"
echo oLink.Save >> "%TEMP%\CreateShortcut.vbs"

cscript //nologo "%TEMP%\CreateShortcut.vbs"
del "%TEMP%\CreateShortcut.vbs"

echo.
echo ========================================
echo   Cai dat thanh cong!
echo ========================================
echo.
echo   - App se tu dong chay khi Windows khoi dong
echo   - Dang khoi dong app...
echo.

:: Khoi dong app ngay
start "" "%CURRENT%start-hidden.vbs"

echo   App da chay! Kiem tra http://localhost:1900
echo.
pause
