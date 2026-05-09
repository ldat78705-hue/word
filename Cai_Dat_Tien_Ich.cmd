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

# 2. Tạo thư mục chứa
$installDir = "C:\TienIchWord"
if (-not (Test-Path $installDir)) {
    New-Item -ItemType Directory -Force -Path $installDir | Out-Null
}
Write-Host "[OK] Da tao thu muc $installDir" -ForegroundColor Green

# 3. Tạo file manifest.xml
$manifestContent = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0" xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="TaskPaneApp">
  <Id>a1b2c3d4-1111-2222-3333-444455556666</Id>
  <Version>1.0.1.5</Version>
  <ProviderName>AddonWord Team</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Tiện ích"/>
  <Description DefaultValue="Công cụ hỗ trợ giáo viên trên Word."/>
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
        <bt:String id="GetStarted.Title" DefaultValue="Tiện ích"/>
        <bt:String id="Cleaner.GroupLabel" DefaultValue="Tiện ích"/>
        <bt:String id="Cleaner.ButtonLabel" DefaultValue="Mở Tiện ích"/>
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="GetStarted.Description" DefaultValue="Mo bang cong cu de sua loi dinh dang khi copy van ban."/>
        <bt:String id="Cleaner.ButtonTooltip" DefaultValue="Hỗ trợ xử lý văn bản và công thức."/>
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>
"@
Set-Content -Path "$installDir\manifest.xml" -Value $manifestContent -Encoding UTF8
Write-Host "[OK] Da tao file manifest.xml" -ForegroundColor Green

# 4. Share thư mục
$shareName = "TienIchWord"
$shareExists = Get-SmbShare -Name $shareName -ErrorAction SilentlyContinue
if (-not $shareExists) {
    New-SmbShare -Name $shareName -Path $installDir -ReadAccess "Everyone" | Out-Null
    Write-Host "[OK] Da chia se mang (Share) thu muc localhost" -ForegroundColor Green
} else {
    Write-Host "[OK] Thu muc da duoc Share tu truoc" -ForegroundColor Green
}

# 5. Thêm Registry
Write-Host "[OK] Dang dang ky Add-in vao Word Trust Center..." -ForegroundColor Yellow
$users = Get-ChildItem "Registry::HKEY_USERS" | Where-Object { $_.Name -match "^HKEY_USERS\\[S\-0\-9\-]+$" -and $_.Name -notmatch "_Classes$" }
foreach ($user in $users) {
    $regPath = "$($user.Name)\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\$shareName"
    if (-not (Test-Path "Registry::$regPath")) {
        New-Item -Path "Registry::$regPath" -Force | Out-Null
    }
    Set-ItemProperty -Path "Registry::$regPath" -Name "Url" -Value "\\localhost\$shareName"
    Set-ItemProperty -Path "Registry::$regPath" -Name "Flags" -Value 1 -Type DWord
}

Write-Host "[OK] Da dang ky thanh cong!" -ForegroundColor Green
Write-Host "======================================" -ForegroundColor Cyan
Write-Host "CAI DAT HOAN TAT!" -ForegroundColor Yellow
Write-Host "Bay gio ban hay:" -ForegroundColor White
Write-Host "1. Tat va mo lai Microsoft Word." -ForegroundColor White
Write-Host "2. Vao the Insert -> My Add-ins -> Chon tab SHARED FOLDER." -ForegroundColor White
Write-Host "3. Chon 'Tien ich' va bam Add (Them)." -ForegroundColor White
Write-Host "======================================" -ForegroundColor Cyan
Write-Host "Nhan Enter de thoat..." -ForegroundColor Gray
Read-Host
