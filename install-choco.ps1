# ============================================================
# Full Setup Script - Windows Server 2022
# Installs: Windows Terminal + Common Packages via Chocolatey
# Run as Administrator
# ============================================================

Write-Host "=====================================" -ForegroundColor Cyan
Write-Host "  Full Server Setup Script" -ForegroundColor Cyan
Write-Host "=====================================" -ForegroundColor Cyan

# ============================================================
Write-Host "`n[1/9] Setting up environment..." -ForegroundColor Cyan
Set-ExecutionPolicy Bypass -Scope Process -Force
[System.Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$tmp = $env:TEMP

# ============================================================
Write-Host "[2/9] Installing VCLibs (latest)..." -ForegroundColor Cyan
Invoke-WebRequest -Uri "https://aka.ms/Microsoft.VCLibs.x64.14.00.Desktop.appx" `
  -OutFile "$tmp\VCLibs.appx" -UseBasicParsing
Add-AppxPackage "$tmp\VCLibs.appx" -ErrorAction SilentlyContinue

# ============================================================
Write-Host "[3/9] Installing UI.Xaml 2.8..." -ForegroundColor Cyan
Invoke-WebRequest -Uri "https://github.com/microsoft/microsoft-ui-xaml/releases/download/v2.8.6/Microsoft.UI.Xaml.2.8.x64.appx" `
  -OutFile "$tmp\UIXaml.appx" -UseBasicParsing
Add-AppxPackage "$tmp\UIXaml.appx" -ErrorAction SilentlyContinue

# ============================================================
Write-Host "[4/9] Installing Windows App Runtime 1.8..." -ForegroundColor Cyan
Invoke-WebRequest -Uri "https://aka.ms/windowsappsdk/1.8/latest/windowsappruntimeinstall-x64.exe" `
  -OutFile "$tmp\WinAppRuntime.exe" -UseBasicParsing
Start-Process "$tmp\WinAppRuntime.exe" -ArgumentList "--quiet --force" -Wait

# ============================================================
Write-Host "[5/9] Installing Chocolatey..." -ForegroundColor Cyan
try {
    Invoke-Expression ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))
    Write-Host "  ✓ Chocolatey installed successfully!" -ForegroundColor Green
} catch {
    Write-Host "  ✗ Failed to install Chocolatey: $_" -ForegroundColor Red
    exit 1
}

# Refresh PATH so choco is available immediately
$env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User")

Write-Host "  Chocolatey version: $(choco --version)" -ForegroundColor Green

# ============================================================
Write-Host "[6/9] Installing Windows Terminal via Chocolatey..." -ForegroundColor Cyan
choco install microsoft-windows-terminal -y
if ($LASTEXITCODE -eq 0) {
    Write-Host "  ✓ Windows Terminal installed!" -ForegroundColor Green
} else {
    Write-Host "  ✗ Windows Terminal install failed" -ForegroundColor Red
}

# ============================================================
Write-Host "[7/9] Installing common packages..." -ForegroundColor Cyan
Write-Host "=====================================" -ForegroundColor Cyan

$packages = @(
    "qbittorrent",
    "python",
    "notepadplusplus",
    "googlechrome",
    "ffmpeg",
    "7zip"
)

foreach ($package in $packages) {
    Write-Host "  Installing $package..." -ForegroundColor Yellow
    choco install $package -y
    if ($LASTEXITCODE -eq 0) {
        Write-Host "  ✓ $package installed successfully!" -ForegroundColor Green
    } else {
        Write-Host "  ✗ Failed to install $package" -ForegroundColor Red
    }
}

# ============================================================
Write-Host "[8/9] Registering Windows Terminal and updating PATH..." -ForegroundColor Cyan

Get-AppxPackage *WindowsTerminal* -AllUsers | ForEach-Object {
    Add-AppxPackage -DisableDevelopmentMode -Register "$($_.InstallLocation)\AppxManifest.xml" -ErrorAction SilentlyContinue
}

$wtInstall = (Get-AppxPackage *WindowsTerminal* -AllUsers | Select-Object -First 1).InstallLocation
if ($wtInstall) {
    $currentPath = [Environment]::GetEnvironmentVariable("Path", "Machine")
    if ($currentPath -notlike "*WindowsTerminal*") {
        [Environment]::SetEnvironmentVariable("Path", $currentPath + ";$wtInstall", "Machine")
        $env:Path += ";$wtInstall"
        Write-Host "  ✓ PATH updated" -ForegroundColor Green
    } else {
        Write-Host "  ✓ PATH already configured" -ForegroundColor Yellow
    }
} else {
    Write-Host "  ✗ Could not find WindowsTerminal install location" -ForegroundColor Red
}

# ============================================================
Write-Host "[9/9] Verifying all installations..." -ForegroundColor Cyan
Write-Host "=====================================" -ForegroundColor Cyan

Write-Host "`n  Windows Terminal:" -ForegroundColor White
Get-AppxPackage *WindowsTerminal* -AllUsers | Select-Object Name, Version, Status | Format-Table -AutoSize

Write-Host "  Dependencies:" -ForegroundColor White
Get-AppxPackage *UI.Xaml.2.8* -AllUsers | Select-Object Name, Version, Status | Format-Table -AutoSize
Get-AppxPackage *WindowsAppRuntime* -AllUsers | Select-Object Name, Version, Status | Format-Table -AutoSize
Get-AppxPackage *VCLibs* -AllUsers | Select-Object Name, Version, Status | Format-Table -AutoSize

Write-Host "  Chocolatey packages:" -ForegroundColor White
choco list

# ============================================================
Write-Host "`n=====================================" -ForegroundColor Green
Write-Host "  Setup Complete!" -ForegroundColor Green
Write-Host "=====================================" -ForegroundColor Green
Write-Host "  Launching Windows Terminal..." -ForegroundColor Cyan
Start-Process wt.exe -ErrorAction SilentlyContinue
Write-Host "  If Terminal didn't open, reboot and run 'wt'" -ForegroundColor Yellow
