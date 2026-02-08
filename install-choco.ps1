# Complete Chocolatey Installation Script
# Run this in PowerShell as Administrator

Write-Host "=====================================" -ForegroundColor Cyan
Write-Host "Chocolatey Installation Script" -ForegroundColor Cyan
Write-Host "=====================================" -ForegroundColor Cyan
Write-Host ""

# Step 1: Fix TLS settings
Write-Host "[1/2] Configuring TLS settings..." -ForegroundColor Yellow
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Step 2: Install Chocolatey
Write-Host "[2/2] Installing Chocolatey..." -ForegroundColor Yellow
Set-ExecutionPolicy Bypass -Scope Process -Force
try {
    iex ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))
    Write-Host "✓ Chocolatey installed successfully!" -ForegroundColor Green
} catch {
    Write-Host "✗ Failed to install Chocolatey: $_" -ForegroundColor Red
    exit 1
}

Write-Host ""

# Refresh environment variables
$env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User")

# Verify Chocolatey installation
Write-Host "Verifying Chocolatey installation..." -ForegroundColor Yellow
choco --version
Write-Host ""

# Step 3: Install packages
Write-Host "=====================================" -ForegroundColor Cyan
Write-Host "Installing Packages" -ForegroundColor Cyan
Write-Host "=====================================" -ForegroundColor Cyan
Write-Host ""

$packages = @(
    "qbittorrent",
    "python",
    "notepadplusplus",
    "googlechrome",
    "ffmpeg",
    "7zip"
)

foreach ($package in $packages) {
    Write-Host "Installing $package..." -ForegroundColor Yellow
    choco install $package -y
    if ($LASTEXITCODE -eq 0) {
        Write-Host "✓ $package installed successfully!" -ForegroundColor Green
    } else {
        Write-Host "✗ Failed to install $package" -ForegroundColor Red
    }
    Write-Host ""
}

Write-Host "=====================================" -ForegroundColor Cyan
Write-Host "Installation Complete!" -ForegroundColor Green
Write-Host "=====================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Installed packages:" -ForegroundColor Cyan
choco list --local-only
