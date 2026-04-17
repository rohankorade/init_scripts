# ============================================================
# Setup-Server.ps1 — Windows Server 2022 Full Bootstrap
# Installs: Chocolatey, Scoop, winget, Windows Terminal,
#           Java (Temurin 21), Python + pip packages,
#           and common CLI tools
#
# Run as Administrator:
#   Set-ExecutionPolicy Bypass -Scope Process -Force
#   .\Setup-Server.ps1
# ============================================================

#Requires -RunAsAdministrator

[CmdletBinding()]
param(
    [switch]$SkipRebootCheck,
    [string]$LogPath = "$env:SystemDrive\Logs\ServerSetup"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ============================================================
# REGION: Logging & Helpers
# ============================================================

$script:LogFile  = $null
$script:Results  = [System.Collections.Generic.List[PSCustomObject]]::new()
$script:HasError = $false

function Initialize-Logging {
    if (-not (Test-Path $LogPath)) { New-Item -ItemType Directory -Path $LogPath -Force | Out-Null }
    $timestamp       = Get-Date -Format "yyyyMMdd_HHmmss"
    $script:LogFile  = Join-Path $LogPath "setup_$timestamp.log"
    Start-Transcript -Path $script:LogFile -Append | Out-Null
    Write-Log "INFO" "Log started: $($script:LogFile)"
}

function Write-Log {
    param([string]$Level, [string]$Message)
    $ts   = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "[$ts] [$Level] $Message"
    switch ($Level) {
        "INFO"    { Write-Host $line -ForegroundColor Cyan   }
        "OK"      { Write-Host $line -ForegroundColor Green  }
        "WARN"    { Write-Host $line -ForegroundColor Yellow }
        "ERROR"   { Write-Host $line -ForegroundColor Red    }
        "SECTION" { Write-Host "`n$('='*60)`n  $Message`n$('='*60)" -ForegroundColor Magenta }
        default   { Write-Host $line }
    }
}

function Add-Result {
    param([string]$Step, [string]$Item, [bool]$Success, [string]$Detail = "")
    $script:Results.Add([PSCustomObject]@{
        Step    = $Step
        Item    = $Item
        Status  = if ($Success) { "OK" } else { "FAILED" }
        Detail  = $Detail
    })
    if (-not $Success) { $script:HasError = $true }
}

function Invoke-Step {
    param(
        [string]   $Name,
        [string]   $Item,
        [scriptblock]$Action,
        [bool]     $ContinueOnError = $true
    )
    Write-Log "INFO" "  → $Item"
    try {
        & $Action
        Write-Log "OK"   "  ✓ $Item"
        Add-Result -Step $Name -Item $Item -Success $true
    } catch {
        Write-Log "ERROR" "  ✗ $Item — $_"
        Add-Result -Step $Name -Item $Item -Success $false -Detail $_.Exception.Message
        if (-not $ContinueOnError) { throw }
    }
}

function Refresh-Env {
    # Reload PATH from registry so newly installed tools are immediately available
    $machinePath = [System.Environment]::GetEnvironmentVariable("Path", "Machine")
    $userPath    = [System.Environment]::GetEnvironmentVariable("Path", "User")
    $env:Path    = "$machinePath;$userPath"
    # Also re-import Chocolatey profile if present
    $chocoProfile = "$env:ChocolateyInstall\helpers\chocolateyProfile.psm1"
    if (Test-Path $chocoProfile) { Import-Module $chocoProfile -ErrorAction SilentlyContinue }
}

function Test-CommandExists {
    param([string]$Command)
    return [bool](Get-Command $Command -ErrorAction SilentlyContinue)
}

function Install-ChocoPackage {
    param([string]$Step, [string]$PackageId, [string[]]$ExtraArgs = @())
    Invoke-Step -Name $Step -Item "choco: $PackageId" -Action {
        $args = @("install", $PackageId, "-y", "--no-progress") + $ExtraArgs
        $result = & choco @args 2>&1
        if ($LASTEXITCODE -notin @(0, 3010)) {   # 3010 = reboot required, treat as OK
            throw "choco exited with code $LASTEXITCODE`n$result"
        }
    }
}

function Install-ScoopPackage {
    param([string]$Step, [string]$PackageId)
    Invoke-Step -Name $Step -Item "scoop: $PackageId" -Action {
        $result = scoop install $PackageId 2>&1
        if ($LASTEXITCODE -ne 0) { throw "scoop exited with code $LASTEXITCODE`n$result" }
    }
}

# ============================================================
# REGION: Pre-flight Checks
# ============================================================

function Invoke-Preflight {
    Write-Log "SECTION" "Pre-flight Checks"

    # OS check
    $os = (Get-WmiObject Win32_OperatingSystem).Caption
    Write-Log "INFO" "OS: $os"
    if ($os -notmatch "Windows") { throw "This script requires Windows." }

    # Architecture
    if ($env:PROCESSOR_ARCHITECTURE -ne "AMD64") {
        throw "Only x64 is supported. Current: $env:PROCESSOR_ARCHITECTURE"
    }

    # Admin guard (belt-and-suspenders beyond #Requires)
    $id = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $p  = New-Object System.Security.Principal.WindowsPrincipal($id)
    if (-not $p.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)) {
        throw "Must be run as Administrator."
    }

    # Execution policy
    Set-ExecutionPolicy Bypass -Scope Process -Force
    [System.Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    # Disk space (require at least 5 GB free on system drive)
    $drive = Get-PSDrive -Name ($env:SystemDrive -replace ":","") -ErrorAction SilentlyContinue
    if ($drive) {
        $freeGB = [math]::Round($drive.Free / 1GB, 1)
        if ($freeGB -lt 5) { Write-Log "WARN" "Low disk space: $freeGB GB free on $env:SystemDrive" }
        else                { Write-Log "OK"   "Disk space: $freeGB GB free" }
    }

    Write-Log "OK" "Pre-flight passed"
}

# ============================================================
# REGION: VCLibs / UI.Xaml / WinApp Runtime (for WinTerminal)
# ============================================================

function Install-WindowsTerminalDeps {
    Write-Log "SECTION" "Installing Windows Terminal Dependencies"
    $tmp = $env:TEMP

    Invoke-Step -Name "WinTermDeps" -Item "VCLibs x64 14.00" -Action {
        if (Get-AppxPackage "*VCLibs*x64*" -AllUsers -ErrorAction SilentlyContinue) {
            Write-Log "INFO" "    Already installed, skipping"
            return
        }
        $dest = "$tmp\VCLibs.appx"
        Invoke-WebRequest "https://aka.ms/Microsoft.VCLibs.x64.14.00.Desktop.appx" `
            -OutFile $dest -UseBasicParsing
        Add-AppxPackage $dest -ErrorAction SilentlyContinue
    }

    Invoke-Step -Name "WinTermDeps" -Item "Microsoft.UI.Xaml 2.8" -Action {
        if (Get-AppxPackage "*UI.Xaml.2.8*" -AllUsers -ErrorAction SilentlyContinue) {
            Write-Log "INFO" "    Already installed, skipping"
            return
        }
        $dest = "$tmp\UIXaml.appx"
        Invoke-WebRequest "https://github.com/microsoft/microsoft-ui-xaml/releases/download/v2.8.6/Microsoft.UI.Xaml.2.8.x64.appx" `
            -OutFile $dest -UseBasicParsing
        Add-AppxPackage $dest -ErrorAction SilentlyContinue
    }

    Invoke-Step -Name "WinTermDeps" -Item "Windows App Runtime 1.8" -Action {
        if (Get-AppxPackage "*WindowsAppRuntime*" -AllUsers -ErrorAction SilentlyContinue) {
            Write-Log "INFO" "    Already installed, skipping"
            return
        }
        $dest = "$tmp\WinAppRuntime.exe"
        Invoke-WebRequest "https://aka.ms/windowsappsdk/1.8/latest/windowsappruntimeinstall-x64.exe" `
            -OutFile $dest -UseBasicParsing
        $proc = Start-Process $dest -ArgumentList "--quiet --force" -Wait -PassThru
        if ($proc.ExitCode -notin @(0, 3010)) { throw "WinAppRuntime installer exited $($proc.ExitCode)" }
    }
}

# ============================================================
# REGION: Chocolatey
# ============================================================

function Install-Chocolatey {
    Write-Log "SECTION" "Installing Chocolatey"

    if (Test-CommandExists "choco") {
        Write-Log "OK" "Chocolatey already installed: $(choco --version)"
        Add-Result -Step "Chocolatey" -Item "choco" -Success $true -Detail "Pre-existing"
        return
    }

    Invoke-Step -Name "Chocolatey" -Item "Chocolatey package manager" -ContinueOnError $false -Action {
        Invoke-Expression ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))
        Refresh-Env
        if (-not (Test-CommandExists "choco")) { throw "choco not found after install" }
        Write-Log "OK" "    Chocolatey $(choco --version) installed"
    }

    # Disable confirmation prompts globally
    choco feature enable -n allowGlobalConfirmation --no-progress | Out-Null
}

# ============================================================
# REGION: Scoop
# ============================================================

function Install-Scoop {
    Write-Log "SECTION" "Installing Scoop"

    if (Test-CommandExists "scoop") {
        Write-Log "OK" "Scoop already installed"
        Add-Result -Step "Scoop" -Item "scoop" -Success $true -Detail "Pre-existing"
        return
    }

    Invoke-Step -Name "Scoop" -Item "Scoop package manager" -ContinueOnError $false -Action {
        # Scoop requires running as a non-admin user OR with SCOOP env set
        # On Server, run with the -RunAsAdmin workaround
        $env:SCOOP = "$env:USERPROFILE\scoop"
        Invoke-RestMethod -Uri "https://get.scoop.sh" | Invoke-Expression
        Refresh-Env
        if (-not (Test-CommandExists "scoop")) { throw "scoop not found after install" }
        Write-Log "OK" "    Scoop installed to $env:SCOOP"
    }

    # Add extras bucket (provides gifsicle and more)
    Invoke-Step -Name "Scoop" -Item "scoop bucket: extras" -Action {
        scoop bucket add extras 2>&1 | Out-Null
    }

    # Add java bucket for JDK installs
    Invoke-Step -Name "Scoop" -Item "scoop bucket: java" -Action {
        scoop bucket add java 2>&1 | Out-Null
    }
}

# ============================================================
# REGION: winget
# ============================================================

function Install-Winget {
    Write-Log "SECTION" "Installing winget"

    if (Test-CommandExists "winget") {
        Write-Log "OK" "winget already installed: $(winget --version)"
        Add-Result -Step "winget" -Item "winget" -Success $true -Detail "Pre-existing"
        return
    }

    Invoke-Step -Name "winget" -Item "winget (via Install-Script)" -Action {
        Install-Script -Name winget-install -Force -ErrorAction Stop
        winget-install -Force
        Refresh-Env
        if (-not (Test-CommandExists "winget")) {
            Write-Log "WARN" "    winget not immediately on PATH — may require reboot"
        } else {
            Write-Log "OK" "    winget $(winget --version)"
        }
    }
}

# ============================================================
# REGION: Windows Terminal
# ============================================================

function Install-WindowsTerminal {
    Write-Log "SECTION" "Installing Windows Terminal"

    $existing = Get-AppxPackage "*WindowsTerminal*" -AllUsers -ErrorAction SilentlyContinue
    if ($existing) {
        Write-Log "OK" "Windows Terminal already installed: $($existing.Version)"
        Add-Result -Step "WindowsTerminal" -Item "Windows Terminal" -Success $true -Detail $existing.Version
        return
    }

    Install-ChocoPackage -Step "WindowsTerminal" -PackageId "microsoft-windows-terminal"

    # Register + add to PATH
    Invoke-Step -Name "WindowsTerminal" -Item "Register AppxManifest" -Action {
        Get-AppxPackage "*WindowsTerminal*" -AllUsers | ForEach-Object {
            Add-AppxPackage -DisableDevelopmentMode `
                -Register "$($_.InstallLocation)\AppxManifest.xml" `
                -ErrorAction SilentlyContinue
        }
    }

    Invoke-Step -Name "WindowsTerminal" -Item "Add wt.exe to system PATH" -Action {
        $wtLocation = (Get-AppxPackage "*WindowsTerminal*" -AllUsers |
            Select-Object -First 1).InstallLocation
        if (-not $wtLocation) { throw "WindowsTerminal install location not found" }
        $cur = [Environment]::GetEnvironmentVariable("Path", "Machine")
        if ($cur -notlike "*WindowsTerminal*") {
            [Environment]::SetEnvironmentVariable("Path", "$cur;$wtLocation", "Machine")
            $env:Path += ";$wtLocation"
        }
    }
}

# ============================================================
# REGION: Java (Eclipse Temurin 21 LTS via Chocolatey)
# ============================================================

function Install-Java {
    Write-Log "SECTION" "Installing Java"

    if (Test-CommandExists "java") {
        $ver = & java -version 2>&1 | Select-Object -First 1
        Write-Log "OK" "Java already installed: $ver"
        Add-Result -Step "Java" -Item "java" -Success $true -Detail $ver
        return
    }

    # Temurin (formerly AdoptOpenJDK) — LTS 21
    Install-ChocoPackage -Step "Java" -PackageId "temurin21"
    Refresh-Env

    Invoke-Step -Name "Java" -Item "Verify java in PATH" -Action {
        if (-not (Test-CommandExists "java")) { throw "java not found after install" }
        Write-Log "OK" "    $(java -version 2>&1 | Select-Object -First 1)"
    }
}

# ============================================================
# REGION: Common Chocolatey Packages
# ============================================================

function Install-ChocoPackages {
    Write-Log "SECTION" "Installing Chocolatey Packages"

    $packages = [ordered]@{
        "python"                      = @()          # Python 3 (latest stable)
        "notepadplusplus"             = @()
        "googlechrome"                = @()
        "ffmpeg"                      = @()
        "7zip"                        = @()
        "qbittorrent"                 = @()
        "git"                         = @()          # Needed by Scoop & general dev
        "curl"                        = @()
        "wget"                        = @()
    }

    foreach ($pkg in $packages.Keys) {
        Install-ChocoPackage -Step "ChocoPackages" -PackageId $pkg -ExtraArgs $packages[$pkg]
    }

    Refresh-Env
}

# ============================================================
# REGION: Python pip Packages
# ============================================================

function Install-PipPackages {
    Write-Log "SECTION" "Installing Python pip Packages"

    # Ensure pip itself is up to date
    Invoke-Step -Name "pip" -Item "Upgrade pip" -Action {
        if (-not (Test-CommandExists "python")) { throw "python not found — ensure Chocolatey python installed" }
        python -m pip install --upgrade pip --quiet
    }

    $pipPackages = @(
        # Date & time
        "python-dateutil",    # Provides dateutil.parser.parse  (covers "parsedate")
        "arrow",              # Friendlier datetime wrapper

        # Numerics / data
        "numpy",
        "pandas",

        # HTTP / networking
        "requests",
        "httpx",

        # CLI / dev utilities
        "rich",               # Beautiful terminal output
        "tqdm",               # Progress bars
        "pydantic",           # Data validation
        "python-dotenv",      # .env file loading
        "loguru",             # Better logging
        "click",              # CLI framework
        "colorama",           # Windows terminal colours
        "psutil",             # System/process info
        "pywin32",            # Windows API bindings
        "pyinstaller"         # Package scripts as .exe
    )

    foreach ($pkg in $pipPackages) {
        Invoke-Step -Name "pip" -Item "pip: $pkg" -Action {
            $result = python -m pip install $pkg --quiet 2>&1
            if ($LASTEXITCODE -ne 0) { throw "pip exited $LASTEXITCODE`n$result" }
        }
    }

    # Confirm key packages importable
    Invoke-Step -Name "pip" -Item "Verify core imports" -Action {
        $check = python -c "import numpy, dateutil, arrow, pandas; print('OK')" 2>&1
        if ($check -ne "OK") { throw "Import check failed: $check" }
    }
}

# ============================================================
# REGION: Scoop Packages
# ============================================================

function Install-ScoopPackages {
    Write-Log "SECTION" "Installing Scoop Packages"

    $scoopPackages = @(
        "gifsicle",     # GIF optimizer (only reliably available via Scoop extras)
        "jq",           # JSON processor
        "fzf",          # Fuzzy finder
        "ripgrep",      # Fast grep
        "fd",           # Fast find replacement
        "bat",          # Better cat
        "lsd",          # Better ls
        "zoxide",       # Smarter cd
        "delta"         # Better git diff
    )

    foreach ($pkg in $scoopPackages) {
        Install-ScoopPackage -Step "ScoopPackages" -PackageId $pkg
    }
}

# ============================================================
# REGION: Final Verification
# ============================================================

function Invoke-Verification {
    Write-Log "SECTION" "Verification"

    $checks = [ordered]@{
        "choco"   = { choco --version }
        "scoop"   = { scoop --version }
        "python"  = { python --version }
        "pip"     = { python -m pip --version }
        "java"    = { java -version 2>&1 | Select-Object -First 1 }
        "git"     = { git --version }
        "ffmpeg"  = { ffmpeg -version 2>&1 | Select-Object -First 1 }
        "7z"      = { 7z i 2>&1 | Select-Object -First 1 }
        "gifsicle"= { gifsicle --version 2>&1 | Select-Object -First 1 }
        "jq"      = { jq --version }
        "winget"  = { if (Test-CommandExists "winget") { winget --version } else { "not in PATH (may need reboot)" } }
    }

    foreach ($cmd in $checks.Keys) {
        try {
            $ver = & $checks[$cmd]
            Write-Log "OK" "  ✓ $cmd — $ver"
        } catch {
            Write-Log "WARN" "  ? $cmd — not found or error"
        }
    }
}

# ============================================================
# REGION: Summary Report
# ============================================================

function Show-Summary {
    Write-Log "SECTION" "Setup Summary"

    $grouped = $script:Results | Group-Object Step
    foreach ($group in $grouped) {
        $ok     = ($group.Group | Where-Object Status -eq "OK").Count
        $failed = ($group.Group | Where-Object Status -eq "FAILED").Count
        $colour = if ($failed -gt 0) { "Yellow" } else { "Green" }
        Write-Host "  [$($group.Name)]  OK: $ok  FAILED: $failed" -ForegroundColor $colour
    }

    $totalOk     = ($script:Results | Where-Object Status -eq "OK").Count
    $totalFailed = ($script:Results | Where-Object Status -eq "FAILED").Count

    Write-Host ""
    Write-Host "  Total: $totalOk succeeded, $totalFailed failed" -ForegroundColor $(if ($totalFailed -gt 0) { "Yellow" } else { "Green" })
    Write-Host ""

    if ($totalFailed -gt 0) {
        Write-Log "WARN" "Failed steps:"
        $script:Results | Where-Object Status -eq "FAILED" | ForEach-Object {
            Write-Log "WARN" "  [$($_.Step)] $($_.Item) — $($_.Detail)"
        }
    }

    Write-Log "INFO" "Full log: $($script:LogFile)"

    if ($script:HasError) {
        Write-Log "WARN" "Setup completed with errors. Review log above."
        exit 1
    } else {
        Write-Log "OK"   "Setup completed successfully!"
        Write-Log "INFO" "Launching Windows Terminal..."
        Start-Process "wt.exe" -ErrorAction SilentlyContinue
        Write-Log "INFO" "If Terminal did not open, reboot and run: wt"
    }
}

# ============================================================
# MAIN
# ============================================================

try {
    Initialize-Logging
    Write-Log "SECTION" "Server Setup — $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"

    Invoke-Preflight
    Install-WindowsTerminalDeps
    Install-Chocolatey
    Install-Scoop
    Install-Winget
    Install-WindowsTerminal
    Install-Java
    Install-ChocoPackages
    Install-PipPackages
    Install-ScoopPackages
    Invoke-Verification
    Show-Summary

} catch {
    Write-Log "ERROR" "Fatal error: $_"
    Write-Log "ERROR" $_.ScriptStackTrace
    if ($script:LogFile) { Stop-Transcript -ErrorAction SilentlyContinue }
    exit 1
} finally {
    Stop-Transcript -ErrorAction SilentlyContinue
}
