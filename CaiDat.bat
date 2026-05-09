@echo off
title Cai dat Word Cleaner Pro (Ban Online)
echo ====================================================
echo      CAI DAT WORD CLEANER PRO (VERCEL / ONLINE)
echo ====================================================
cd /d "%~dp0"

echo.
echo [*] Dang dang ky ung dung vao Microsoft Word...

powershell -Command "$wefPath = Join-Path $env:LOCALAPPDATA 'Microsoft\Office\16.0\Wef\Developer'; if (!(Test-Path $wefPath)) { New-Item -ItemType Directory -Path $wefPath -Force | Out-Null }; Copy-Item -Path '.\manifest.xml' -Destination (Join-Path $wefPath 'WordCleanerVN.xml') -Force; Write-Host '=> Da cai dat xong!'"

echo.
echo ====================================================
echo                 HOAN TAT CAI DAT!
echo ====================================================
echo Ban chi can tat mo lai Word, ung dung se xuat hien tren thanh Home.
echo ====================================================
pause
