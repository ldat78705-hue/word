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

# Cai dat qua Developer (Sideloading) cho cac may ho tro 1-click
$registryPath = "HKCU:\Software\Microsoft\Office\16.0\WEF\Developer"
New-Item -Path $registryPath -Force | Out-Null
$guid = "c0079ced-e355-4102-938b-e03b7baa45df"
Set-ItemProperty -Path $registryPath -Name $guid -Value $manifestTarget

# Cai dat qua Trusted Catalogs cho cac may bao mat cao (nhu Office LTSC)
$catalogPath = "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\WordCleaner"
New-Item -Path $catalogPath -Force | Out-Null
Set-ItemProperty -Path $catalogPath -Name "Id" -Value $targetDir
Set-ItemProperty -Path $catalogPath -Name "Flags" -Value 1

Write-Host "=> DA CAI DAT XONG TEP HE THONG!"
