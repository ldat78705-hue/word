<# :
@echo off
setlocal
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -Command "Invoke-Expression (Get-Content '%~f0' -Raw)"
exit /b
#>
$ErrorActionPreference = "Stop"

# 1. Kiểm tra quyền Admin
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Host "Dang yeu cau quyen quan tri (Admin)..." -ForegroundColor Yellow
    Start-Process -FilePath "cmd.exe" -ArgumentList "/c `"$($MyInvocation.MyCommand.Path)`"" -Verb RunAs
    Exit
}

Write-Host "======================================" -ForegroundColor Cyan
Write-Host "   CAI DAT TIEN ICH GIAO VIEN 1-CLICK " -ForegroundColor Cyan
Write-Host "======================================" -ForegroundColor Cyan

# 2. Xóa phiên bản cũ và dọn dẹp Cache của Word
Write-Host "[1/6] Don dep phien ban cu va Cache..." -ForegroundColor Yellow
$installDir = "C:\TienIchWord"
if (Test-Path $installDir) {
    try {
        Remove-Item -Path "$installDir\*" -Recurse -Force -ErrorAction SilentlyContinue
        Write-Host "-> [OK] Da xoa sach thu muc cai dat cu." -ForegroundColor Green
    } catch { }
} else {
    $null = New-Item -ItemType Directory -Force -Path $installDir
}

# Xóa cache WEF của tất cả người dùng để Word bắt buộc tải lại bản mới nhất
$usersDirs = Get-ChildItem "C:\Users" -Directory
foreach ($ud in $usersDirs) {
    $wefPath = "$($ud.FullName)\AppData\Local\Microsoft\Office\16.0\Wef"
    if (Test-Path $wefPath) {
        try {
            Remove-Item -Path "$wefPath\*" -Recurse -Force -ErrorAction SilentlyContinue
        } catch { }
    }
}
Write-Host "-> [OK] Da xoa bo nho dem (Cache) cua Add-in tren Word." -ForegroundColor Green


# 3. Kiểm tra và cài đặt WebView2 Runtime
Write-Host "[2/6] Kiem tra Microsoft Edge WebView2 Runtime..." -ForegroundColor Yellow
$wv2RegPath1 = "HKLM:\SOFTWARE\WOW6432Node\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}"
$wv2RegPath2 = "HKLM:\SOFTWARE\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}"
$wv2RegPath3 = "HKCU:\Software\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}"

$isWv2Installed = (Test-Path $wv2RegPath1) -or (Test-Path $wv2RegPath2) -or (Test-Path $wv2RegPath3)

if (-not $isWv2Installed) {
    Write-Host "-> Chua cai dat WebView2. Dang tai va cai dat tu dong (vui long doi it phut)..." -ForegroundColor Yellow
    try {
        $wv2Installer = "$env:TEMP\MicrosoftEdgeWebview2Setup.exe"
        Invoke-WebRequest -Uri "https://go.microsoft.com/fwlink/p/?LinkId=2124703" -OutFile $wv2Installer -UseBasicParsing
        Start-Process -FilePath $wv2Installer -ArgumentList "/silent /install" -Wait -NoNewWindow
        Write-Host "-> [OK] Da cai dat WebView2 thanh cong!" -ForegroundColor Green
    } catch {
        Write-Host "-> [!] Khong the tai/cai dat WebView2 tu dong. Add-in co the bi loi giao dien tren Word cu." -ForegroundColor Red
    }
} else {
    Write-Host "-> [OK] WebView2 da duoc cai dat, tuong thich tot voi Office.js." -ForegroundColor Green
}

# Kích hoạt Webview2 thay vì IE cho Office (nếu có thể)
try {
    $edgeOptPath = "HKCU:\Software\Microsoft\Office\16.0\WEF"
    if (-not (Test-Path $edgeOptPath)) {
        $null = New-Item -Path $edgeOptPath -Force
    }
    Set-ItemProperty -Path $edgeOptPath -Name "Win32WebView2" -Value 1 -Type DWord -ErrorAction SilentlyContinue
} catch {}


# 4. Tạo file manifest.xml
Write-Host "[3/6] Khoi tao Manifest moi..." -ForegroundColor Yellow
$manifestContent = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0" xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="TaskPaneApp">
  <Id>a1b2c3d4-1111-2222-3333-444455556666</Id>
  <Version>1.0.1.7</Version>
  <ProviderName>AddonWord Team</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Tien Ich"/>
  <Description DefaultValue="Lam sach van ban copy tu nguon khac trong Word."/>
  <IconUrl DefaultValue="https://word-td.vercel.app/icon.png"/>
  <HighResolutionIconUrl DefaultValue="https://word-td.vercel.app/icon.png"/>
  <SupportUrl DefaultValue="https://word-td.vercel.app/support.html"/>
  <Hosts>
    <Host Name="Document"/>
  </Hosts>
  <DefaultSettings>
    <SourceLocation DefaultValue="https://word-td.vercel.app/taskpane.html"/>
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>
  <VersionOverrides xsi:type="ov:VersionOverridesV1_0" xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides">
    <Hosts>
      <Host xsi:type="ov:Document">
        <DesktopFormFactor>
          <GetStarted>
            <Title resid="GetStarted.Title"/>
            <Description resid="GetStarted.Description"/>
            <LearnMoreUrl resid="GetStarted.LearnMoreUrl"/>
          </GetStarted>
          <FunctionFile resid="Taskpane.Url"/>
          <ExtensionPoint xsi:type="PrimaryCommandSurface">
            <OfficeTab id="TabHome">
              <Group id="Cleaner.Group">
                <Label resid="Cleaner.GroupLabel"/>
                <Icon>
                  <bt:Image size="16" resid="Icon.16x16"/>
                  <bt:Image size="32" resid="Icon.32x32"/>
                  <bt:Image size="80" resid="Icon.80x80"/>
                </Icon>
                <Control xsi:type="Button" id="Cleaner.TaskpaneButton">
                  <Label resid="Cleaner.ButtonLabel"/>
                  <Supertip>
                    <Title resid="Cleaner.ButtonLabel"/>
                    <Description resid="Cleaner.ButtonTooltip"/>
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Icon.16x16"/>
                    <bt:Image size="32" resid="Icon.32x32"/>
                    <bt:Image size="80" resid="Icon.80x80"/>
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <TaskpaneId>ButtonId1</TaskpaneId>
                    <SourceLocation resid="Taskpane.Url"/>
                  </Action>
                </Control>
              </Group>
            </OfficeTab>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>
    </Hosts>
    <Resources>
      <bt:Images>
        <bt:Image id="Icon.16x16" DefaultValue="https://word-td.vercel.app/icon.png"/>
        <bt:Image id="Icon.32x32" DefaultValue="https://word-td.vercel.app/icon.png"/>
        <bt:Image id="Icon.80x80" DefaultValue="https://word-td.vercel.app/icon.png"/>
      </bt:Images>
      <bt:Urls>
        <bt:Url id="GetStarted.LearnMoreUrl" DefaultValue="https://word-td.vercel.app/support.html"/>
        <bt:Url id="Taskpane.Url" DefaultValue="https://word-td.vercel.app/taskpane.html"/>
      </bt:Urls>
      <bt:ShortStrings>
        <!-- KHÔNG NÊN ĐỂ TIẾNG VIỆT CÓ DẤU Ở ĐÂY SẼ BỊ LỖI FONT -->
        <bt:String id="GetStarted.Title" DefaultValue="Tien Ich"/>
        <bt:String id="Cleaner.GroupLabel" DefaultValue="Tien Ich"/>
        <bt:String id="Cleaner.ButtonLabel" DefaultValue="Mo Tien Ich"/>
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="GetStarted.Description" DefaultValue="Mo bang cong cu de sua loi dinh dang khi copy van ban."/>
        <bt:String id="Cleaner.ButtonTooltip" DefaultValue="Lam sach khoang trang, xuong dong, doan trong trong van ban."/>
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>
"@
Set-Content -Path "$installDir\manifest.xml" -Value $manifestContent -Encoding UTF8
Write-Host "-> [OK] Da tao file manifest.xml" -ForegroundColor Green

# 4.5 Copy MathType Server và tạo Startup
Write-Host "[4/6] Cai dat Server MathType cuc bo..." -ForegroundColor Yellow
$currentDir = (Get-Location).Path
$serverSrc = "$currentDir\MathTypeServer.exe"
$serverDest = "$installDir\MathTypeServer.exe"

# Tat process cu neu dang chay
try { Stop-Process -Name "MathTypeServer" -Force -ErrorAction SilentlyContinue } catch { }

if (Test-Path $serverSrc) {
    try {
        Copy-Item -Path $serverSrc -Destination $serverDest -Force
        
        # Tao shortcut vao thu muc Startup de chay cung Windows
        $startupFolder = [Environment]::GetFolderPath("Startup")
        $shortcutPath = "$startupFolder\MathTypeServer.lnk"
        
        $WshShell = New-Object -comObject WScript.Shell
        $Shortcut = $WshShell.CreateShortcut($shortcutPath)
        $Shortcut.TargetPath = $serverDest
        $Shortcut.WindowStyle = 7 # Minimized
        $Shortcut.Save()
        
        # Chay luon server ngay bay gio (thong qua explorer de ha quyen Admin ve quyen User, giup bat duoc COM Word dang mo)
        Start-Process -FilePath "explorer.exe" -ArgumentList $serverDest
        Write-Host "-> [OK] Da cai dat va khoi dong MathType Server ngam." -ForegroundColor Green
    } catch {
        Write-Host "-> [!] Loi khi copy MathType Server. Vui long tat cac file Word va thu lai." -ForegroundColor Red
    }
} else {
    Write-Host "-> [!] Khong tim thay MathTypeServer.exe de cai dat." -ForegroundColor Yellow
}

# 5. Share thư mục
Write-Host "[5/6] Chia se thu muc mang..." -ForegroundColor Yellow
$shareName = "TienIchWord"
$shareExists = Get-SmbShare -Name $shareName -ErrorAction SilentlyContinue
if (-not $shareExists) {
    $null = New-SmbShare -Name $shareName -Path $installDir -ReadAccess "Everyone"
    Write-Host "-> [OK] Da chia se mang (Share) thu muc localhost\$shareName" -ForegroundColor Green
} else {
    Write-Host "-> [OK] Thu muc da duoc Share tu truoc" -ForegroundColor Green
}

# 6. Thêm Registry cho Trust Center
Write-Host "[6/6] Dang ky Add-in vao Word Trust Center..." -ForegroundColor Yellow
$successCount = 0

# Cách 1: Đăng ký cho toàn bộ Users
try {
    $users = Get-ChildItem "Registry::HKEY_USERS" | Where-Object { $_.Name -match "^HKEY_USERS\\[S\-0\-9\-]+$" -and $_.Name -notmatch "_Classes$" }
    foreach ($user in $users) {
        try {
            $regPath = "$($user.Name)\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\$shareName"
            if (-not (Test-Path "Registry::$regPath")) {
                $null = New-Item -Path "Registry::$regPath" -Force -ErrorAction Stop
            }
            Set-ItemProperty -Path "Registry::$regPath" -Name "Url" -Value "\\localhost\$shareName" -ErrorAction Stop
            Set-ItemProperty -Path "Registry::$regPath" -Name "Flags" -Value 1 -Type DWord -ErrorAction Stop
            $successCount++
        } catch { }
    }
} catch { }

# Cách 2: Đăng ký thẳng cho Current User (HKCU)
try {
    $hkcuPath = "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\$shareName"
    if (-not (Test-Path $hkcuPath)) {
        $null = New-Item -Path $hkcuPath -Force -ErrorAction Stop
    }
    Set-ItemProperty -Path $hkcuPath -Name "Url" -Value "\\localhost\$shareName" -ErrorAction Stop
    Set-ItemProperty -Path $hkcuPath -Name "Flags" -Value 1 -Type DWord -ErrorAction Stop
    
    $successCount++
} catch { }

if ($successCount -eq 0) {
    Write-Host "-> [!] Khong the dang ky vao Registry. Ban co the can them Share folder vao Trust Center thu cong." -ForegroundColor Red
} else {
    Write-Host "-> [OK] Da dang ky Trust Center thanh cong ($successCount profiles)!" -ForegroundColor Green
}

Write-Host "======================================" -ForegroundColor Cyan
Write-Host "CAI DAT HOAN TAT!" -ForegroundColor Yellow
Write-Host "Bay gio ban hay:" -ForegroundColor White
Write-Host "1. Tat va mo lai Microsoft Word." -ForegroundColor White
Write-Host "2. Vao the Insert (Chen) -> My Add-ins (Add-in cua toi) -> Chon tab SHARED FOLDER (Thu muc dung chung)." -ForegroundColor White
Write-Host "3. Chon 'Tien ich' va bam Add (Them)." -ForegroundColor White
Write-Host "======================================" -ForegroundColor Cyan
Write-Host "Nhan Enter de thoat..." -ForegroundColor Gray
Read-Host
