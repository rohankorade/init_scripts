# ==============================================================================
# Setup-TailscaleRDP.ps1
#
# PURPOSE  : Install Tailscale via winget, authenticate headlessly with a
#            pre-auth key, configure the service for fully unattended operation,
#            then lock down RDP so it only accepts connections from the
#            Tailscale subnet.
#
# USAGE    : Run as Administrator in an elevated PowerShell session.
#
#   .\Setup-TailscaleRDP.ps1 -AuthKey "tskey-auth-XXXXXXXXXXXXXXXXX"
#
# PARAMETERS
#   -AuthKey         (Required) Pre-auth key from admin.tailscale.com
#   -TailscaleSubnet (Optional) CIDR allowed for RDP. Default: 100.64.0.0/10
#   -SkipFirewall    (Optional) Skip the RDP firewall lockdown step
#   -LogPath         (Optional) Transcript log path. Default: script directory
#
# SAFE TO RE-RUN : Every step checks current state before acting.
#
# FIREWALL NOTE  : Windows Firewall BLOCK rules always beat ALLOW rules.
#                  This script does NOT add a block rule. Instead it:
#                    1. Disables all existing inbound RDP allow rules
#                    2. Adds ONE allow rule scoped to the Tailscale subnet
#                    3. Relies on Windows Firewall's default inbound DENY
#                       to block everything else - which is the correct approach.
# ==============================================================================

#Requires -RunAsAdministrator
#Requires -Version 5.1

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [ValidatePattern('^tskey-[a-zA-Z0-9_\-]+$')]
    [string]$AuthKey,

    [Parameter(Mandatory = $false)]
    [ValidatePattern('^\d{1,3}(\.\d{1,3}){3}/\d{1,2}$')]
    [string]$TailscaleSubnet = "100.64.0.0/10",

    [Parameter(Mandatory = $false)]
    [switch]$SkipFirewall,

    [Parameter(Mandatory = $false)]
    [string]$LogPath = (Join-Path $PSScriptRoot "Setup-TailscaleRDP_$(Get-Date -Format 'yyyyMMdd_HHmmss').log")
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# -- Constants -----------------------------------------------------------------

$TAILSCALE_DEFAULT_PATH = "C:\Program Files\Tailscale\tailscale.exe"
$TAILSCALE_SERVICE_NAME = "Tailscale"
$RDP_ALLOW_RULE_NAME    = "RDP - Allow Tailscale Subnet Only"

# -- Transcript logging --------------------------------------------------------

try {
    Start-Transcript -Path $LogPath -Append | Out-Null
    Write-Host "Logging transcript to: $LogPath" -ForegroundColor DarkGray
} catch {
    Write-Warning "Could not start transcript logging: $_"
}

# -- Helper functions ----------------------------------------------------------

function Write-Step { param($msg) Write-Host "`n[ .. ] $msg" -ForegroundColor Cyan   }
function Write-OK   { param($msg) Write-Host "[ OK ] $msg"   -ForegroundColor Green  }
function Write-Warn { param($msg) Write-Host "[ !! ] $msg"   -ForegroundColor Yellow }
function Write-Fail { param($msg) Write-Host "[ XX ] $msg"   -ForegroundColor Red    }

function Invoke-NativeCommand {
    <#
    .SYNOPSIS
        Runs an external command and throws on unexpected exit codes.
    #>
    param(
        [string]   $Executable,
        [string[]] $Arguments        = @(),
        [int[]]    $AllowedExitCodes = @(0)
    )
    & $Executable @Arguments
    if ($LASTEXITCODE -notin $AllowedExitCodes) {
        throw "'$Executable $($Arguments -join ' ')' exited with code $LASTEXITCODE"
    }
}

function Resolve-TailscaleBinary {
    <#
    .SYNOPSIS
        Returns path to tailscale.exe - checks PATH first, then known install location.
    #>
    $cmd = Get-Command "tailscale" -ErrorAction SilentlyContinue
    if ($cmd) { return $cmd.Source }
    if (Test-Path $TAILSCALE_DEFAULT_PATH) { return $TAILSCALE_DEFAULT_PATH }
    return $null
}

function Get-TailscaleStatus {
    <#
    .SYNOPSIS
        Returns parsed Tailscale JSON status object, or $null on failure.
    #>
    try {
        $raw = & tailscale status --json 2>&1
        return ($raw | ConvertFrom-Json)
    } catch {
        return $null
    }
}

function Get-TailscaleIP {
    $status = Get-TailscaleStatus
    if (-not $status) { return $null }
    return ($status.TailscaleIPs | Where-Object { $_ -match '^100\.' } | Select-Object -First 1)
}

function Wait-ForService {
    <#
    .SYNOPSIS
        Polls a Windows service until it reaches the desired status or times out.
    #>
    param(
        [string] $ServiceName,
        [string] $DesiredStatus  = "Running",
        [int]    $TimeoutSeconds = 30
    )
    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    while ((Get-Date) -lt $deadline) {
        $svc = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
        if ($svc -and $svc.Status -eq $DesiredStatus) { return $true }
        Start-Sleep -Seconds 2
    }
    return $false
}

function Wait-TailscaleRunning {
    <#
    .SYNOPSIS
        Polls tailscale status until BackendState = Running or timeout.
    #>
    param ([int]$TimeoutSeconds = 90)
    Write-Step "Waiting for Tailscale to reach 'Running' state (timeout: ${TimeoutSeconds}s)..."
    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    while ((Get-Date) -lt $deadline) {
        $status = Get-TailscaleStatus
        if ($status -and $status.BackendState -eq "Running") {
            Write-OK "Tailscale backend is Running."
            return $true
        }
        $stateLabel = if ($status) { $status.BackendState } else { "no response" }
        $remaining  = [int](($deadline - (Get-Date)).TotalSeconds)
        Write-Host "   ...state: $stateLabel  (${remaining}s remaining)" -ForegroundColor DarkGray
        Start-Sleep -Seconds 3
    }
    return $false
}

function Stop-WithError {
    <#
    .SYNOPSIS
        Prints a failure message, stops the transcript, and exits with code 1.
    #>
    param([string]$Message)
    Write-Fail $Message
    try { Stop-Transcript | Out-Null } catch { }
    exit 1
}

# -- STEP 0 : Execution policy guard ------------------------------------------

Write-Step "Checking PowerShell execution policy..."
$policy = Get-ExecutionPolicy -Scope Process
if ($policy -eq "Restricted") {
    Stop-WithError "Execution policy is 'Restricted'. Run: Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass"
}
Write-OK "Execution policy OK: $policy"

# -- STEP 1 : Administrator confirmation ---------------------------------------
# (#Requires -RunAsAdministrator already enforces this; this is belt-and-suspenders
#  and surfaces a clear diagnostic message rather than a cryptic PS error.)

Write-Step "Confirming Administrator privileges..."
$identity  = [Security.Principal.WindowsIdentity]::GetCurrent()
$principal = [Security.Principal.WindowsPrincipal]$identity
if (-not $principal.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)) {
    Stop-WithError "Not running as Administrator. Re-launch PowerShell as Administrator."
}
Write-OK "Running as Administrator ($($identity.Name))."

# -- STEP 2 : Confirm winget is available --------------------------------------

Write-Step "Checking winget availability..."
if (-not (Get-Command "winget" -ErrorAction SilentlyContinue)) {
    Stop-WithError "winget not found. Install 'App Installer' from: https://aka.ms/getwinget"
}
$wingetVer = & winget --version 2>&1
Write-OK "winget found: $wingetVer"

# -- STEP 3 : Install Tailscale (idempotent) -----------------------------------

Write-Step "Checking if Tailscale is already installed..."
$tailscaleBin = Resolve-TailscaleBinary

if ($tailscaleBin) {
    Write-OK "Tailscale already installed at: $tailscaleBin"
} else {
    Write-Step "Tailscale not found - installing via winget..."

    # winget exit codes:
    #   0             = success
    #   -1978335189   = package already installed (0x8A150015) - treat as success
    try {
        Invoke-NativeCommand "winget" @(
            "install",
            "--id",    "Tailscale.Tailscale",
            "--exact",
            "--silent",
            "--scope", "machine",
            "--accept-package-agreements",
            "--accept-source-agreements"
        ) -AllowedExitCodes @(0, -1978335189)
    } catch {
        Stop-WithError "winget install failed: $_"
    }

    Write-OK "winget install completed."

    # Refresh PATH in this session so tailscale.exe becomes reachable
    $env:PATH = [System.Environment]::GetEnvironmentVariable("PATH", "Machine") + ";" +
                [System.Environment]::GetEnvironmentVariable("PATH", "User")

    # Re-resolve (PATH + known fallback path)
    $tailscaleBin = Resolve-TailscaleBinary
    if (-not $tailscaleBin) {
        Stop-WithError "tailscale.exe not found after install. Expected: $TAILSCALE_DEFAULT_PATH"
    }
    Write-OK "Tailscale binary confirmed at: $tailscaleBin"

    # If the binary was found via fallback path but is not in PATH yet, add it
    if (-not (Get-Command "tailscale" -ErrorAction SilentlyContinue)) {
        $tailscaleDir = Split-Path $tailscaleBin
        $env:PATH     = "$env:PATH;$tailscaleDir"
        Write-Warn "Added '$tailscaleDir' to session PATH."
    }
}

# -- STEP 4 : Ensure Tailscale service exists and is running -------------------

Write-Step "Checking Tailscale Windows service ($TAILSCALE_SERVICE_NAME)..."
$svc = Get-Service -Name $TAILSCALE_SERVICE_NAME -ErrorAction SilentlyContinue

if (-not $svc) {
    Stop-WithError "Service '$TAILSCALE_SERVICE_NAME' not found. Installation may be incomplete."
}
Write-OK "Service found. Current status: $($svc.Status)"

if ($svc.Status -ne "Running") {
    Write-Warn "Service is not running - starting it now..."
    Start-Service -Name $TAILSCALE_SERVICE_NAME
    $started = Wait-ForService -ServiceName $TAILSCALE_SERVICE_NAME -DesiredStatus "Running" -TimeoutSeconds 30
    if (-not $started) {
        Stop-WithError "Service did not reach 'Running' within 30s. Check Event Viewer (Application + System)."
    }
    Write-OK "Service started successfully."
} else {
    Write-OK "Service is already Running."
}

# -- STEP 5 : Configure service for unattended / headless operation -------------
#
#  Three things configured:
#
#  * Startup type = Automatic (Delayed)
#    Survives reboots with no user login required.
#    Delayed start gives the NIC / DNS stack time to initialise before
#    Tailscale attempts its first connection.
#
#  * Failure recovery
#    Auto-restarts on any crash with exponential back-off.
#    Counter resets after 1 hour of clean uptime.
#
#  * failureflag = 1
#    Extends recovery to any non-zero exit (clean-but-failed shutdowns),
#    not just hard crashes.

Write-Step "Configuring service for unattended headless operation..."

# 5a. Automatic (Delayed Start)
Write-Host "   Setting startup: Automatic (Delayed)..." -ForegroundColor DarkGray
& sc.exe config $TAILSCALE_SERVICE_NAME start= delayed-auto | Out-Null
if ($LASTEXITCODE -ne 0) {
    Write-Warn "sc.exe delayed-auto failed (exit $LASTEXITCODE) - falling back to plain Automatic."
    Set-Service -Name $TAILSCALE_SERVICE_NAME -StartupType Automatic
} else {
    Write-OK "Startup type: Automatic (Delayed)."
}

# 5b. Failure recovery: restart at 5s / 15s / 60s; reset counter after 1hr
Write-Host "   Configuring failure recovery..." -ForegroundColor DarkGray
& sc.exe failure $TAILSCALE_SERVICE_NAME `
    reset= 3600 `
    actions= restart/5000/restart/15000/restart/60000 | Out-Null
if ($LASTEXITCODE -ne 0) {
    Write-Warn "Could not set failure recovery actions (exit $LASTEXITCODE). Non-fatal - continuing."
} else {
    Write-OK "Failure recovery: restart at 5s / 15s / 60s."
}

# 5c. failureflag: trigger recovery on non-crash non-zero exits too
Write-Host "   Enabling failure actions flag..." -ForegroundColor DarkGray
& sc.exe failureflag $TAILSCALE_SERVICE_NAME 1 | Out-Null
if ($LASTEXITCODE -ne 0) {
    Write-Warn "Could not set failureflag (exit $LASTEXITCODE). Non-fatal - continuing."
} else {
    Write-OK "Failure actions flag enabled."
}

# 5d. Confirm
$svcConfig = (& sc.exe qc $TAILSCALE_SERVICE_NAME 2>&1) -join " "
if ($svcConfig -match "DELAYED") {
    Write-OK "Confirmed: service is Automatic (Delayed Start)."
} else {
    Write-Warn "Could not confirm delayed-start. Verify with: sc.exe qc $TAILSCALE_SERVICE_NAME"
}

# -- STEP 6 : Authenticate with auth key ---------------------------------------

Write-Step "Checking current Tailscale authentication state..."
$currentStatus = Get-TailscaleStatus
$backendState  = if ($currentStatus) { $currentStatus.BackendState } else { "Unknown" }
Write-Host "   Backend state: $backendState" -ForegroundColor DarkGray

if ($backendState -eq "Running") {
    $existingIP = Get-TailscaleIP
    Write-OK "Already authenticated and running. Tailscale IP: $existingIP"
    Write-Warn "Skipping login. To force re-auth: tailscale logout && re-run this script."
} else {
    Write-Step "Authenticating with provided auth key..."
    try {
        # Auth key is passed as a separate argument - avoids any shell quoting issues
        Invoke-NativeCommand "tailscale" @(
            "login",
            "--authkey", $AuthKey,
            "--unattended"
        ) -AllowedExitCodes @(0)
    } catch {
        Stop-WithError "tailscale login failed: $_"
    }
    Write-OK "Login command accepted."

    $connected = Wait-TailscaleRunning -TimeoutSeconds 90
    if (-not $connected) {
        $lastStatus = Get-TailscaleStatus
        $lastState  = if ($lastStatus) { $lastStatus.BackendState } else { "no response" }
        Stop-WithError "Tailscale did not reach 'Running' within 90s. Last state: $lastState. Check: auth key validity, DNS, access to controlplane.tailscale.com."
    }
}

# -- STEP 7 : Confirm Tailscale IP - hard gate before touching firewall ---------

Write-Step "Confirming Tailscale IP assignment..."
$myTailscaleIP = Get-TailscaleIP

if (-not $myTailscaleIP) {
    Stop-WithError "No Tailscale IP (100.x.x.x) detected. Firewall will NOT be modified. Run: tailscale status"
}
Write-OK "Tailscale IP: $myTailscaleIP"

# -- STEP 8 : Harden RDP firewall rules ---------------------------------------
#
#  WHY NO BLOCK RULE:
#  Windows Firewall processes rules in this fixed order:
#    (1) Authenticated bypass  (2) BLOCK  (3) ALLOW  (4) Default policy
#  A "block all on 3389" + "allow Tailscale on 3389" would have the block
#  win - Tailscale RDP would also be blocked. The correct pattern is:
#    * Remove/disable all existing RDP allow rules
#    * Add ONE allow rule for the Tailscale subnet
#    * Default inbound = Block handles everything else
#
#  This is the standard Windows Firewall least-privilege pattern.

if ($SkipFirewall) {
    Write-Warn "-SkipFirewall set - skipping all firewall changes."
} else {
    Write-Step "Hardening RDP firewall - restricting port 3389 to Tailscale subnet only..."

    # 8a. Ensure Windows Firewall is enabled on all profiles
    Write-Host "   Verifying Windows Firewall is enabled..." -ForegroundColor DarkGray
    $fwProfiles      = Get-NetFirewallProfile
    $disabledProfiles = $fwProfiles | Where-Object { $_.Enabled -eq $false }
    if ($disabledProfiles) {
        Write-Warn "Firewall disabled on: $($disabledProfiles.Name -join ', ') - enabling now..."
        Set-NetFirewallProfile -Profile Domain,Public,Private -Enabled True
        Write-OK "Windows Firewall enabled on all profiles."
    } else {
        Write-OK "Windows Firewall is enabled on all profiles."
    }

    # 8b. Ensure default inbound policy is Block on all profiles
    Write-Host "   Verifying default inbound policy is Block..." -ForegroundColor DarkGray
    $fwProfiles | ForEach-Object {
        if ($_.DefaultInboundAction -ne "Block") {
            Write-Warn "Profile '$($_.Name)' DefaultInboundAction = $($_.DefaultInboundAction) - setting to Block..."
            Set-NetFirewallProfile -Name $_.Name -DefaultInboundAction Block
        }
    }
    Write-OK "Default inbound policy is Block on all profiles."

    # 8c. Disable ALL existing inbound RDP allow rules (built-in + any legacy custom)
    Write-Host "   Disabling all existing inbound RDP allow rules on port 3389..." -ForegroundColor DarkGray
    $allInboundAllowRules = Get-NetFirewallRule -ErrorAction SilentlyContinue |
        Where-Object {
            $_.Direction -eq "Inbound" -and
            $_.Action    -eq "Allow"   -and
            $_.Enabled   -eq $true     -and
            $_.DisplayName -ne $RDP_ALLOW_RULE_NAME   # don't kill our own rule on re-run
        } | ForEach-Object {
            $portFilter = $_ | Get-NetFirewallPortFilter -ErrorAction SilentlyContinue
            if ($portFilter -and ($portFilter.LocalPort -eq "3389" -or $portFilter.LocalPort -contains "3389")) {
                $_
            }
        }

    if ($allInboundAllowRules) {
        foreach ($rule in $allInboundAllowRules) {
            Write-Warn "  Disabling: '$($rule.DisplayName)'"
            Disable-NetFirewallRule -InputObject $rule
        }
        Write-OK "All pre-existing inbound RDP allow rules disabled."
    } else {
        Write-OK "No pre-existing inbound RDP allow rules to disable."
    }

    # 8d. Remove stale copy of our own rule (safe re-run)
    $staleRule = Get-NetFirewallRule -DisplayName $RDP_ALLOW_RULE_NAME -ErrorAction SilentlyContinue
    if ($staleRule) {
        Write-Warn "Removing stale rule from previous run: '$RDP_ALLOW_RULE_NAME'"
        Remove-NetFirewallRule -DisplayName $RDP_ALLOW_RULE_NAME
    }

    # 8e. Create the single Tailscale-scoped allow rule
    Write-Step "Creating allow rule: RDP from $TailscaleSubnet only..."
    New-NetFirewallRule `
        -DisplayName   $RDP_ALLOW_RULE_NAME `
        -Direction     Inbound `
        -Protocol      TCP `
        -LocalPort     3389 `
        -RemoteAddress $TailscaleSubnet `
        -Action        Allow `
        -Profile       Any `
        -Enabled       True `
        -Description   "Permits inbound RDP only from the Tailscale CGNAT range ($TailscaleSubnet). All other source IPs are denied by the default inbound Block policy. Managed by Setup-TailscaleRDP.ps1." `
    | Out-Null
    Write-OK "Allow rule created: '$RDP_ALLOW_RULE_NAME'"

    # 8f. Verify the rule is in place and inspect its address filter
    Write-Step "Verifying final firewall configuration..."
    $verifyRule = Get-NetFirewallRule -DisplayName $RDP_ALLOW_RULE_NAME -ErrorAction SilentlyContinue
    if (-not $verifyRule) {
        Stop-WithError "Allow rule not found after creation. Review Windows Firewall manually."
    }
    $addrFilter = $verifyRule | Get-NetFirewallAddressFilter
    $portFilter = $verifyRule | Get-NetFirewallPortFilter
    Write-OK "Rule verified:"
    Write-OK "  Name          : $($verifyRule.DisplayName)"
    Write-OK "  Enabled       : $($verifyRule.Enabled)"
    Write-OK "  Action        : $($verifyRule.Action)"
    Write-OK "  RemoteAddress : $($addrFilter.RemoteAddress)"
    Write-OK "  LocalPort     : $($portFilter.LocalPort)"
}

# -- DONE ----------------------------------------------------------------------

Write-Host ""
Write-Host "================================================================" -ForegroundColor Green
Write-Host "   ALL STEPS COMPLETED SUCCESSFULLY                             " -ForegroundColor Green
Write-Host "================================================================" -ForegroundColor Green
Write-Host ""
Write-Host "  Tailscale IP        : $myTailscaleIP"
if (-not $SkipFirewall) {
    Write-Host "  RDP restricted to   : $TailscaleSubnet (Tailscale subnet only)"
    Write-Host "  Inbound default     : Block  (all other source IPs denied)"
}
Write-Host "  Service startup     : Automatic (Delayed) + auto-restart on failure"
Write-Host "  Log file            : $LogPath"
Write-Host ""
Write-Host "  HOW TO CONNECT VIA RDP:" -ForegroundColor White
Write-Host "    1. Ensure your client machine is also connected to Tailscale"
Write-Host "    2. Open RDP and connect to: $myTailscaleIP"
Write-Host ""
Write-Host "  TO UNDO FIREWALL CHANGES:" -ForegroundColor DarkGray
Write-Host "    Remove-NetFirewallRule -DisplayName '$RDP_ALLOW_RULE_NAME'" -ForegroundColor DarkGray
Write-Host "    Enable-NetFirewallRule -DisplayGroup 'Remote Desktop'"      -ForegroundColor DarkGray
Write-Host ""

try { Stop-Transcript | Out-Null } catch { }
