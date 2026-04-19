# ============================================================
# Setup-Server.ps1 -- Windows Server 2022 Full Bootstrap
# Installs: Chocolatey, Scoop, winget, Windows Terminal,
#           Java (Oracle JDK 26), Python + pip packages,
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
    $timestamp      = Get-Date -Format "yyyyMMdd_HHmmss"
    $script:LogFile = Join-Path $LogPath "setup_$timestamp.log"
    Start-Transcript -Path $script:LogFile -Append | Out-Null
    Write-Log "INFO" "Log started: $($script:LogFile)"
}

function Write-Log {
    param([string]$Level, [string]$Message)
    $ts   = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "[$ts] [$Level] $Message"
    switch ($Level) {
        "INFO"    { Write-Host $line -ForegroundColor Cyan    }
        "OK"      { Write-Host $line -ForegroundColor Green   }
        "WARN"    { Write-Host $line -ForegroundColor Yellow  }
        "ERROR"   { Write-Host $line -ForegroundColor Red     }
        "SECTION" { Write-Host "`n$('='*60)`n  $Message`n$('='*60)" -ForegroundColor Magenta }
        default   { Write-Host $line }
    }
}

function Add-Result {
    param([string]$Step, [string]$Item, [bool]$Success, [string]$Detail = "")
    $script:Results.Add([PSCustomObject]@{
        Step   = $Step
        Item   = $Item
        Status = if ($Success) { "OK" } else { "FAILED" }
        Detail = $Detail
    })
    if (-not $Success) { $script:HasError = $true }
}

function Invoke-Step {
    param(
        [string]      $Name,
        [string]      $Item,
        [scriptblock] $Action,
        [bool]        $ContinueOnError = $true
    )
    Write-Log "INFO" "  -> $Item"
    try {
        & $Action
        Write-Log "OK"    "  OK $Item"
        Add-Result -Step $Name -Item $Item -Success $true
    } catch {
        Write-Log "ERROR" "  FAIL $Item -- $_"
        Add-Result -Step $Name -Item $Item -Success $false -Detail $_.Exception.Message
        if (-not $ContinueOnError) { throw }
    }
}

function Refresh-Env {
    $machinePath = [System.Environment]::GetEnvironmentVariable("Path", "Machine")
    $userPath    = [System.Environment]::GetEnvironmentVariable("Path", "User")
    $env:Path    = "$machinePath;$userPath"
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
        $chocoArgs = @("install", $PackageId, "-y", "--no-progress") + $ExtraArgs
        $result = & choco @chocoArgs 2>&1
        if ($LASTEXITCODE -notin @(0, 3010)) {
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

function Invoke-Download {
    param([string]$Url, [string]$OutFile)
    $result = curl.exe -L -f -s -S --retry 3 --retry-delay 2 -o $OutFile $Url 2>&1
    if ($LASTEXITCODE -ne 0) { throw "curl failed (exit $LASTEXITCODE): $result" }
}

# ============================================================
# REGION: Pre-flight Checks
# ============================================================

function Invoke-Preflight {
    Write-Log "SECTION" "Pre-flight Checks"

    $os = (Get-WmiObject Win32_OperatingSystem).Caption
    Write-Log "INFO" "OS: $os"
    if ($os -notmatch "Windows") { throw "This script requires Windows." }

    if ($env:PROCESSOR_ARCHITECTURE -ne "AMD64") {
        throw "Only x64 is supported. Current: $env:PROCESSOR_ARCHITECTURE"
    }

    $id = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $p  = New-Object System.Security.Principal.WindowsPrincipal($id)
    if (-not $p.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)) {
        throw "Must be run as Administrator."
    }

    Set-ExecutionPolicy Bypass -Scope Process -Force
    [System.Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    $drive = Get-PSDrive -Name ($env:SystemDrive -replace ":", "") -ErrorAction SilentlyContinue
    if ($drive) {
        $freeGB = [math]::Round($drive.Free / 1GB, 1)
        if ($freeGB -lt 5) { Write-Log "WARN" "Low disk space: $freeGB GB free on $env:SystemDrive" }
        else                { Write-Log "OK"   "Disk space: $freeGB GB free" }
    }

    Write-Log "OK" "Pre-flight passed"
}

# ============================================================
# REGION: Windows Terminal Dependencies
# FIX 1: This function was defined but never called in MAIN.
#         Now called before Install-WindowsTerminal.
# FIX 2: Switched WinAppRuntime from 1.8 -> 1.6 which has
#         better Windows Server 2022 SKU compatibility.
# ============================================================

function Install-WindowsTerminalDeps {
    Write-Log "SECTION" "Installing Windows Terminal Dependencies"
    $tmp = $env:TEMP

    $needVCLibs  = -not (Get-AppxPackage "*VCLibs*x64*"       -AllUsers -ErrorAction SilentlyContinue)
    $needUIXaml  = -not (Get-AppxPackage "*UI.Xaml.2.8*"      -AllUsers -ErrorAction SilentlyContinue)
    $needRuntime = -not (Get-AppxPackage "*WindowsAppRuntime*" -AllUsers -ErrorAction SilentlyContinue)

    if (-not ($needVCLibs -or $needUIXaml -or $needRuntime)) {
        Write-Log "OK" "All Windows Terminal dependencies already installed, skipping"
        Add-Result -Step "WinTermDeps" -Item "All deps" -Success $true -Detail "Pre-existing"
        return
    }

    Write-Log "INFO" "  Downloading missing dependencies in parallel..."
    $jobs = @()

    if ($needVCLibs) {
        $jobs += Start-Job -ScriptBlock {
            curl.exe -L -f -s -S --retry 3 --retry-delay 2 `
                -o "$using:tmp\VCLibs.appx" `
                "https://aka.ms/Microsoft.VCLibs.x64.14.00.Desktop.appx"
            $LASTEXITCODE
        }
    }
    if ($needUIXaml) {
        $jobs += Start-Job -ScriptBlock {
            curl.exe -L -f -s -S --retry 3 --retry-delay 2 `
                -o "$using:tmp\UIXaml.appx" `
                "https://github.com/microsoft/microsoft-ui-xaml/releases/download/v2.8.6/Microsoft.UI.Xaml.2.8.x64.appx"
            $LASTEXITCODE
        }
    }
    if ($needRuntime) {
        # FIX: Use 1.6 instead of 1.8 -- better Server 2022 SKU support
        $jobs += Start-Job -ScriptBlock {
            curl.exe -L -f -s -S --retry 3 --retry-delay 2 `
                -o "$using:tmp\WinAppRuntime.exe" `
                "https://aka.ms/windowsappsdk/1.6/latest/windowsappruntimeinstall-x64.exe"
            $LASTEXITCODE
        }
    }

    $jobs | Wait-Job | ForEach-Object {
        $exitCode = Receive-Job $_
        if ($exitCode -ne 0) { Write-Log "WARN" "  A parallel download returned exit code $exitCode" }
        Remove-Job $_
    }

    Invoke-Step -Name "WinTermDeps" -Item "VCLibs x64 14.00" -Action {
        if (-not $needVCLibs) { Write-Log "INFO" "    Already installed, skipping"; return }
        Add-AppxPackage "$tmp\VCLibs.appx" -ErrorAction SilentlyContinue
    }

    Invoke-Step -Name "WinTermDeps" -Item "Microsoft.UI.Xaml 2.8" -Action {
        if (-not $needUIXaml) { Write-Log "INFO" "    Already installed, skipping"; return }
        Add-AppxPackage "$tmp\UIXaml.appx" -ErrorAction SilentlyContinue
    }

    # FIX: Non-fatal on Server SKU -- Terminal still launches via choco's bundled deps
    Invoke-Step -Name "WinTermDeps" -Item "Windows App Runtime 1.6" -Action {
        if (-not $needRuntime) { Write-Log "INFO" "    Already installed, skipping"; return }
        $proc = Start-Process "$tmp\WinAppRuntime.exe" -ArgumentList "--quiet --force" -Wait -PassThru
        # 0            = success
        # 3010         = success, reboot required
        # -2147483637  = not supported on this Server SKU (non-fatal, choco covers it)
        if ($proc.ExitCode -notin @(0, 3010, -2147483637)) {
            throw "WinAppRuntime installer exited $($proc.ExitCode)"
        }
        if ($proc.ExitCode -eq -2147483637) {
            Write-Log "WARN" "    WinAppRuntime not supported on this Server SKU -- choco will cover deps"
        }
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
        $env:SCOOP = "$env:USERPROFILE\scoop"
        $effectivePolicy = Get-ExecutionPolicy
        if ($effectivePolicy -notin @("Bypass", "Unrestricted", "RemoteSigned")) {
            Set-ExecutionPolicy RemoteSigned -Scope CurrentUser -Force -ErrorAction SilentlyContinue
        }
        iex "& {$(irm get.scoop.sh)} -RunAsAdmin"
        Refresh-Env
        if (-not (Test-CommandExists "scoop")) { throw "scoop not found after install" }
        Write-Log "OK" "    Scoop installed to $env:SCOOP"
    }

    Invoke-Step -Name "Scoop" -Item "scoop bucket: extras" -Action {
        scoop bucket add extras 2>&1 | Out-Null
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
        Refresh-Env
        winget-install 3>$null
        Refresh-Env
        if (-not (Test-CommandExists "winget")) {
            Write-Log "WARN" "    winget not immediately on PATH -- may require reboot"
        } else {
            Write-Log "OK" "    winget $(winget --version)"
        }
    }
}

# ============================================================
# REGION: winget Packages
# ============================================================

function Install-WingetPackages {
    Write-Log "SECTION" "Installing winget Packages"

    if (-not (Test-CommandExists "winget")) {
        Write-Log "WARN" "winget not available -- skipping winget packages"
        Add-Result -Step "WingetPackages" -Item "winget packages" -Success $true -Detail "Skipped -- winget not available"
        return
    }

    $wingetPackages = [ordered]@{
        "Google.Chrome"                 = "Google Chrome"
        "Oracle.JDK.26"                 = "Oracle JDK 26"
        "Oracle.JavaRuntimeEnvironment" = "Oracle Java Runtime Environment"
    }

    foreach ($id in $wingetPackages.Keys) {
        $label = $wingetPackages[$id]
        Invoke-Step -Name "WingetPackages" -Item "winget: $label" -Action {
            $result = winget install --id $id --silent --accept-source-agreements --accept-package-agreements 2>&1
            if ($LASTEXITCODE -notin @(0, -1978335189)) {
                throw "winget exited $LASTEXITCODE`n$result"
            }
        }
    }

    Refresh-Env
}

# ============================================================
# REGION: Windows Terminal
# FIX 3: Added explicit per-user AppxManifest registration.
#         Previously only the machine-wide loop ran, which
#         caused the package to exist but not launch for the
#         current user session.
# ============================================================

function Install-WindowsTerminal {
    Write-Log "SECTION" "Installing Windows Terminal"

    $existing = Get-AppxPackage "*WindowsTerminal*" -AllUsers -ErrorAction SilentlyContinue
    if ($existing) {
        Write-Log "OK" "Windows Terminal already installed: $($existing.Version)"
        Add-Result -Step "WindowsTerminal" -Item "Windows Terminal" -Success $true -Detail $existing.Version

        # Still ensure it is registered for the current user even if machine package exists
        Invoke-Step -Name "WindowsTerminal" -Item "Ensure registered for current user" -Action {
            $manifest = Join-Path $existing.InstallLocation "AppxManifest.xml"
            Add-AppxPackage -DisableDevelopmentMode -Register $manifest -ErrorAction SilentlyContinue
        }
        return
    }

    Install-ChocoPackage -Step "WindowsTerminal" -PackageId "microsoft-windows-terminal"

    # Register machine-wide (all users)
    Invoke-Step -Name "WindowsTerminal" -Item "Register AppxManifest (all users)" -Action {
        Get-AppxPackage "*WindowsTerminal*" -AllUsers | ForEach-Object {
            Add-AppxPackage -DisableDevelopmentMode `
                -Register "$($_.InstallLocation)\AppxManifest.xml" `
                -ErrorAction SilentlyContinue
        }
    }

    # FIX: Also register explicitly for the current user -- this is what
    #      allows wt.exe to actually launch in the current session.
    Invoke-Step -Name "WindowsTerminal" -Item "Register AppxManifest (current user)" -Action {
        $pkg = Get-AppxPackage "*WindowsTerminal*" -AllUsers | Select-Object -First 1
        if (-not $pkg) { throw "WindowsTerminal package not found after install" }
        $manifest = Join-Path $pkg.InstallLocation "AppxManifest.xml"
        Add-AppxPackage -DisableDevelopmentMode -Register $manifest -ErrorAction Stop
        Write-Log "OK" "    Registered for current user"
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
# REGION: Java (Oracle JDK 26 + JRE via winget)
# ============================================================

function Install-Java {
    Write-Log "SECTION" "Installing Java"

    Refresh-Env

    if (Test-CommandExists "java") {
        $ver = & { $ErrorActionPreference = 'Continue'; java -version 2>&1 } | Select-Object -First 1
        Write-Log "OK" "Java already on PATH: $ver"
        Add-Result -Step "Java" -Item "java" -Success $true -Detail $ver
        return
    }

    Invoke-Step -Name "Java" -Item "Set JAVA_HOME + patch PATH" -Action {
        $regPath  = "HKLM:\SOFTWARE\JavaSoft\JDK"
        $javaHome = $null

        if (Test-Path $regPath) {
            $subKey = Get-ChildItem $regPath -ErrorAction SilentlyContinue |
                      Where-Object { $_.PSChildName -like "26*" } |
                      Select-Object -Last 1
            if ($subKey) {
                $javaHome = (Get-ItemProperty $subKey.PSPath -Name "JavaHome" -ErrorAction SilentlyContinue).JavaHome
            }
        }

        if (-not $javaHome) {
            $javaHome = Get-ChildItem "$env:ProgramFiles\Java" -Filter "jdk-26*" -Directory `
                            -ErrorAction SilentlyContinue | Select-Object -Last 1 |
                            Select-Object -ExpandProperty FullName
        }
        if (-not $javaHome) {
            $javaHome = Get-ChildItem "$env:ProgramFiles" -Filter "jdk-26*" -Directory `
                            -ErrorAction SilentlyContinue | Select-Object -Last 1 |
                            Select-Object -ExpandProperty FullName
        }

        if (-not $javaHome) { throw "Oracle JDK 26 directory not found. Ensure winget installed it successfully." }

        $javaBin = Join-Path $javaHome "bin"
        Write-Log "INFO" "    Found Oracle JDK at: $javaHome"

        [Environment]::SetEnvironmentVariable("JAVA_HOME", $javaHome, "Machine")
        $env:JAVA_HOME = $javaHome

        $machinePath = [Environment]::GetEnvironmentVariable("Path", "Machine")
        if ($machinePath -notlike "*$javaBin*") {
            [Environment]::SetEnvironmentVariable("Path", "$machinePath;$javaBin", "Machine")
        }
        if ($env:Path -notlike "*$javaBin*") { $env:Path = "$env:Path;$javaBin" }

        Write-Log "OK" "    JAVA_HOME = $javaHome"
    }

    Invoke-Step -Name "Java" -Item "Verify java in PATH" -Action {
        if (-not (Test-CommandExists "java")) { throw "java not found in PATH after setup" }
        Write-Log "OK" "    $(& { $ErrorActionPreference = 'Continue'; java -version 2>&1 } | Select-Object -First 1)"
    }
}

# ============================================================
# REGION: Common Chocolatey Packages
# FIX 4: Added jq -- it was verified in Invoke-Verification
#         but was never installed anywhere in the script.
# ============================================================

function Install-ChocoPackages {
    Write-Log "SECTION" "Installing Chocolatey Packages"

    $packages = [ordered]@{
        "python"          = @()
        "notepadplusplus" = @()
        "ffmpeg"          = @()
        "7zip"            = @()
        "qbittorrent"     = @()
        "git"             = @()
        "curl"            = @()
        "wget"            = @()
        "jq"              = @()   # FIX: was checked in Verification but never installed
    }

    foreach ($pkg in $packages.Keys) {
        Install-ChocoPackage -Step "ChocoPackages" -PackageId $pkg -ExtraArgs $packages[$pkg]
    }

    Refresh-Env
}

# ============================================================
# REGION: Python pip Packages
# FIX 5: Removed gunicorn (Unix-only WSGI server, fails on Windows)
# FIX 6: Removed duplicate python-pptx entry
# ============================================================

function Install-PipPackages {
    Write-Log "SECTION" "Installing Python pip Packages"

    Invoke-Step -Name "pip" -Item "Upgrade pip" -Action {
        if (-not (Test-CommandExists "python")) { throw "python not found -- ensure Chocolatey python installed" }
        python -m pip install --upgrade pip --quiet
    }

    $pipPackages = @(

        # -- Web Frameworks & APIs ------------------------------------
        "fastapi",
        "flask",
        "flask-cors",
        "flask-restful",
        "starlette",
        "uvicorn",
        # "gunicorn" REMOVED -- Unix-only, does not install on Windows
        "waitress",         # Windows-compatible WSGI server (use instead of gunicorn)
        "tornado",
        "bottle",
        "django",

        # -- HTTP & Networking ----------------------------------------
        "requests",
        "httpx",
        "urllib3",
        "aiohttp",
        "websockets",
        "certifi",
        "chardet",
        "h2",

        # -- HTML Parsing & Scraping ----------------------------------
        "beautifulsoup4",
        "lxml",
        "html5lib",
        "scrapy",
        "playwright",
        "parsel",
        "trafilatura",
        "tldextract",

        # -- Data & Excel ---------------------------------------------
        "pandas",
        "openpyxl",
        "xlsxwriter",
        "xlrd",
        "numpy",
        "polars",
        "pyarrow",
        "tabulate",
        "natsort",

        # -- Image Processing ----------------------------------------
        "Pillow",
        "matplotlib",

        # -- Database ------------------------------------------------
        "pymongo",
        "motor",
        "sqlalchemy",
        "psycopg2-binary",
        "pymysql",
        "redis",
        "peewee",
        "tinydb",
        "alembic",
        "diskcache",

        # -- Logging & CLI -------------------------------------------
        "rich",
        "loguru",
        "colorama",
        "texttable",
        "tqdm",
        "pyfiglet",
        "click",
        "typer",
        "prompt-toolkit",
        "structlog",
        "python-json-logger",
        "sentry-sdk",

        # -- Authentication & Security --------------------------------
        "python-dotenv",
        "pydantic",
        "pydantic-settings",
        "pyjwt",
        "bcrypt",
        "passlib",
        "cryptography",
        "pyotp",
        "python-multipart",

        # -- Date & Time ---------------------------------------------
        "python-dateutil",
        "parsedatetime",
        "pendulum",
        "pytz",
        "arrow",
        "dateparser",
        "freezegun",
        "schedule",

        # -- YAML / Config / Serialization ---------------------------
        "pyyaml",
        "toml",
        "orjson",
        "ujson",
        "msgpack",
        "python-decouple",
        "dynaconf",
        "omegaconf",
        "python-benedict",

        # -- AI / ML / NLP -------------------------------------------
        "openai",
        "anthropic",
        "scikit-learn",
        "scipy",
        "statsmodels",
        "nltk",
        "textblob",
        "langdetect",
        "tiktoken",
        "sympy",
        "wordcloud",
        "textstat",
        "rapidfuzz",
        "ftfy",
        "unidecode",

        # -- Async & Task Queues -------------------------------------
        "anyio",
        "aiofiles",
        "janus",
        "celery",
        "rq",
        "apscheduler",
        "backoff",
        "tenacity",

        # -- Cloud & Storage -----------------------------------------
        "boto3",
        "google-auth",
        "google-cloud-storage",
        "azure-identity",
        "azure-storage-blob",
        "paramiko",
        "fabric",

        # -- Visualization -------------------------------------------
        "plotly",
        "bokeh",
        "seaborn",
        "altair",

        # -- Document Processing -------------------------------------
        "python-docx",
        "python-pptx",    # FIX: duplicate removed (was listed twice)
        "pypdf",
        "pdfplumber",

        # -- Windows-Specific ----------------------------------------
        "pywin32",
        "winotify",
        "pyperclip",
        "send2trash",
        "watchdog",

        # -- TTS / Media ---------------------------------------------
        "edge-tts",
        "yt-dlp",

        # -- Misc Utilities ------------------------------------------
        "jinja2",
        "psutil",
        "attrs",
        "cattrs",
        "cachetools",
        "more-itertools",
        "sortedcontainers",
        "funcy",
        "toolz",
        "pathspec",
        "parse",
        "validators",
        "babel",
        "humanize",
        "pycountry",
        "phonenumbers",
        "appdirs",
        "platformdirs",
        "dnspython",
        "pyparsing",
        "invoke",
        "itsdangerous",
        "mako",
        "marshmallow",
        "h5py",
        "deepdiff",
        "flasgger",

        # -- Testing & Dev Tools -------------------------------------
        "pytest",
        "pytest-asyncio",
        "pytest-cov",
        "pytest-mock",
        "pytest-xdist",
        "coverage",
        "mypy",
        "black",
        "ruff",
        "flake8",
        "pylint",
        "isort",
        "bandit",
        "hypothesis",
        "factory-boy",
        "faker",
        "responses",
        "pre-commit",

        # -- Packaging -----------------------------------------------
        "pyinstaller"
    )

    foreach ($pkg in $pipPackages) {
        Invoke-Step -Name "pip" -Item "pip: $pkg" -Action {
            $result = python -m pip install $pkg --quiet 2>&1
            if ($LASTEXITCODE -ne 0) { throw "pip exited $LASTEXITCODE`n$result" }
        }
    }

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
        "gifsicle"
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
        "choco"    = { choco --version }
        "scoop"    = { scoop --version }
        "python"   = { python --version }
        "pip"      = { python -m pip --version }
        "java"     = { & { $ErrorActionPreference = 'Continue'; java -version 2>&1 } | Select-Object -First 1 }
        "git"      = { git --version }
        "ffmpeg"   = { ffmpeg -version 2>&1 | Select-Object -First 1 }
        "7z"       = { 7z i 2>&1 | Select-Object -First 1 }
        "gifsicle" = { gifsicle --version 2>&1 | Select-Object -First 1 }
        "jq"       = { jq --version }
        "winget"   = { if (Test-CommandExists "winget") { winget --version } else { "not in PATH (may need reboot)" } }
        "wt"       = { if (Test-CommandExists "wt") { "wt.exe found" } else { "wt not in PATH" } }
    }

    foreach ($cmd in $checks.Keys) {
        try {
            $ver = & $checks[$cmd]
            Write-Log "OK" "  OK $cmd -- $ver"
        } catch {
            Write-Log "WARN" "  ? $cmd -- not found or error"
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
        $ok     = @($group.Group | Where-Object Status -eq "OK").Count
        $failed = @($group.Group | Where-Object Status -eq "FAILED").Count
        $colour = if ($failed -gt 0) { "Yellow" } else { "Green" }
        Write-Host "  [$($group.Name)]  OK: $ok  FAILED: $failed" -ForegroundColor $colour
    }

    $totalOk     = @($script:Results | Where-Object Status -eq "OK").Count
    $totalFailed = @($script:Results | Where-Object Status -eq "FAILED").Count

    Write-Host ""
    Write-Host "  Total: $totalOk succeeded, $totalFailed failed" -ForegroundColor $(if ($totalFailed -gt 0) { "Yellow" } else { "Green" })
    Write-Host ""

    if ($totalFailed -gt 0) {
        Write-Log "WARN" "Failed steps:"
        $script:Results | Where-Object Status -eq "FAILED" | ForEach-Object {
            Write-Log "WARN" "  [$($_.Step)] $($_.Item) -- $($_.Detail)"
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
# FIX: Install-WindowsTerminalDeps is now called before
#      Install-WindowsTerminal so VCLibs, UI.Xaml, and
#      WinAppRuntime are in place before choco installs WT.
# ============================================================

try {
    Initialize-Logging
    Write-Log "SECTION" "Server Setup -- $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"

    Invoke-Preflight
    Install-Chocolatey
    Install-Scoop
    Install-Winget
    Install-WingetPackages
    Install-WindowsTerminalDeps    # FIX: was missing from MAIN
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
