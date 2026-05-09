@echo off
:: Kiem tra quyen Administrator
openfiles >nul 2>&1
if '%errorlevel%' NEQ '0' (
    echo Vui long doi... Dang yeu cau quyen Quan tri vien (Administrator)...
    powershell Start-Process -FilePath "%0" -Verb RunAs
    exit /b
)

echo ===================================================
echo   CAI DAT TIEN ICH GIAO VIEN (thaydat.edu.vn)
echo ===================================================
echo Dang thiet lap chia se thu muc...

powershell -ExecutionPolicy Bypass -Command " ^
$folderPath = 'C:\TienIchGiaoVien'; ^
if (!(Test-Path $folderPath)) { New-Item -ItemType Directory -Force -Path $folderPath | Out-Null }; ^
$currentDir = '%~dp0'; ^
$manifestSource = Join-Path $currentDir 'manifest.xml'; ^
if (Test-Path $manifestSource) { Copy-Item $manifestSource -Destination $folderPath -Force } else { ^
    try { Invoke-WebRequest -Uri 'https://word-td.vercel.app/manifest.xml' -OutFile \"$folderPath\manifest.xml\" } catch { Write-Host 'Loi: Khong tim thay manifest.xml' -ForegroundColor Red; Pause; Exit } ^
}; ^
$shareName = 'TienIchGiaoVienShare'; ^
if (Get-SmbShare -Name $shareName -ErrorAction SilentlyContinue) { Remove-SmbShare -Name $shareName -Force }; ^
New-SmbShare -Name $shareName -Path $folderPath -FullAccess 'Everyone' | Out-Null; ^
$uncPath = '\\' + $env:COMPUTERNAME + '\' + $shareName; ^
$registryPath = 'HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs'; ^
if (!(Test-Path $registryPath)) { New-Item -Path $registryPath -Force | Out-Null }; ^
$catalogKey = \"$registryPath\TienIch\"; ^
if (!(Test-Path $catalogKey)) { New-Item -Path $catalogKey -Force | Out-Null }; ^
Set-ItemProperty -Path $catalogKey -Name 'Flags' -Value 1 -Type DWord; ^
Set-ItemProperty -Path $catalogKey -Name 'Id' -Value 'TienIchGiaoVien' -Type String; ^
Set-ItemProperty -Path $catalogKey -Name 'Url' -Value $uncPath -Type String; ^
Write-Host ''; Write-Host 'THANH CONG! Vui long tat va mo lai Word.' -ForegroundColor Green; ^
Write-Host 'Vao tab Insert -^> Get Add-ins -^> SHARED FOLDER de kich hoat.' -ForegroundColor Yellow; ^
"
pause
