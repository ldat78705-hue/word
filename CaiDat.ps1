$ErrorActionPreference = "SilentlyContinue"

Write-Host "Dang chuan bi cai dat Word Cleaner Pro..."
Stop-Process -Name WINWORD -Force | Out-Null

# Tao thu muc dung chung o Public de dam bao quyen truy cap
$targetDir = "C:\Users\Public\WordCleanerPro"
if (!(Test-Path $targetDir)) {
    New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
}

$manifestSource = Join-Path $PSScriptRoot "manifest.xml"
$manifestTarget = Join-Path $targetDir "manifest.xml"
Copy-Item -Path $manifestSource -Destination $manifestTarget -Force | Out-Null

# Xoa cac the tich cu
Remove-Item -Path "HKCU:\Software\Microsoft\Office\16.0\WEF\Developer" -Recurse -Force | Out-Null
Remove-Item -Path "HKCU:\Software\Microsoft\Office\16.0\WEF\Cache" -Recurse -Force | Out-Null
Remove-Item -Path "$env:LOCALAPPDATA\Microsoft\Office\16.0\Wef" -Recurse -Force | Out-Null

# Cai dat qua Developer (Sideloading) - Day la cach duy nhat hien thi nut tren Ribbon
$registryPath = "HKCU:\Software\Microsoft\Office\16.0\WEF\Developer"
New-Item -Path $registryPath -Force | Out-Null
$guid = "e3d99f22-f555-43c7-a92c-09b55e22ccdd"
Set-ItemProperty -Path $registryPath -Name $guid -Value $manifestTarget

Write-Host "=> DA CAI DAT XONG TEP HE THONG!"
