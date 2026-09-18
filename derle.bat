@echo off
setlocal
title Fatura Yonetim Sistemi - Derleme
cd /d "%~dp0"

rem derle.ps1 cagrilir; argumanlar (ornegin -Temiz) aynen iletilir.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0derle.ps1" %*
set KOD=%ERRORLEVEL%

echo.
if "%KOD%"=="0" goto basarili
echo Derleme BASARISIZ (cikis kodu: %KOD%)
goto son

:basarili
echo Derleme tamamlandi: %~dp0dist\FaturaYonetimSistemi.exe

:son
echo.
pause
exit /b %KOD%
