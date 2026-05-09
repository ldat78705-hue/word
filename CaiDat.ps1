$ErrorActionPreference = "SilentlyContinue"

Write-Host "Dang chuan bi cai dat Word Cleaner Pro..."
Stop-Process -Name WINWORD -Force | Out-Null

# 1. Copy manifest vao thu muc an toan cua he thong
$targetDir = Join-Path $env:LOCALAPPDATA "WordCleanerPro"
if (!(Test-Path $targetDir)) {
    New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
}

$manifestSource = Join-Path $PSScriptRoot "manifest.xml"
$manifestTarget = Join-Path $targetDir "manifest.xml"
Copy-Item -Path $manifestSource -Destination $manifestTarget -Force | Out-Null

# 2. Don dep moi vet tich cu de chong xung dot
Write-Host "Dang lam sach he thong Word..."
Remove-Item -Path "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs" -Recurse -Force | Out-Null
Remove-Item -Path "HKCU:\Software\Microsoft\Office\16.0\WEF\Developer" -Recurse -Force | Out-Null
Remove-Item -Path "HKCU:\Software\Microsoft\Office\16.0\WEF\Cache" -Recurse -Force | Out-Null
Remove-Item -Path "$env:LOCALAPPDATA\Microsoft\Office\16.0\Wef" -Recurse -Force | Out-Null

# 3. Dang ky chay tu dong 1-click vao thanh Ribbon (Sideloading)
Write-Host "Dang tich hop vao thanh cong cu..."
$registryPath = "HKCU:\Software\Microsoft\Office\16.0\WEF\Developer"
New-Item -Path $registryPath -Force | Out-Null

# Day la GUID da duoc tao rieng cho phien ban Vercel nay
$guid = "c0079ced-e355-4102-938b-e03b7baa45df"
Set-ItemProperty -Path $registryPath -Name $guid -Value $manifestTarget

Write-Host "=> DA CAI DAT XONG THANH CONG!"
Write-Host "=> Ung dung da tu dong ghim vao Word. Ban khong can thao tac gi them!"
