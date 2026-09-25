@echo off
setlocal
chcp 65001 >nul
set "DOTNET_CLI_UI_LANGUAGE=vi-VN"
cd /d "%~dp0"
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0Parse-LocalXml.ps1" -Interactive
set "exitCode=%ERRORLEVEL%"
echo.
if not "%exitCode%"=="0" echo Chuong trinh ket thuc voi ma loi %exitCode%.
echo Nhan phim bat ky de dong cua so...
pause >nul
exit /b %exitCode%
