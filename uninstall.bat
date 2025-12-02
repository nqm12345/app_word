@echo off
echo ========================================
echo   Word WebDAV Server - Go bo
echo ========================================

:: Tat process
echo Dang tat app...
taskkill /F /IM node.exe 2>nul

:: Xoa shortcut
set STARTUP=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup
del "%STARTUP%\WordWebDAV.lnk" 2>nul

echo.
echo ========================================
echo   Da go bo thanh cong!
echo ========================================
echo.
pause
