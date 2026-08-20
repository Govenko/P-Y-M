# Installs the parser apps: creates .venv, installs dependencies, creates desktop shortcuts.
$ErrorActionPreference = "Stop"
$env:PYTHONUTF8 = "1"
$env:PYTHONIOENCODING = "utf-8"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $root

Write-Host "=== Parsers setup (Avito / Yandex Maps / HH.ru / Rusprofile) ==="
Write-Host "Folder: $root"
Write-Host ""

# --- 1. Find Python 3.10+ ---
$pythonCmd = $null
$pythonArgs = @()
foreach ($candidate in @(@("python", @()), @("py", @("-3")))) {
    $name = $candidate[0]
    $args = $candidate[1]
    if (Get-Command $name -ErrorAction SilentlyContinue) {
        try {
            $version = (& $name @args --version) 2>$null
        } catch { continue }
        if ("$version" -match "Python 3\.(\d+)") {
            if ([int]$Matches[1] -ge 10) {
                $pythonCmd = $name
                $pythonArgs = $args
                Write-Host "Python found: $version"
                break
            }
        }
    }
}
if (-not $pythonCmd) {
    Write-Host ""
    Write-Host "ERROR: Python 3.10+ not found." -ForegroundColor Red
    Write-Host "Install it from https://www.python.org/downloads/windows/"
    Write-Host "IMPORTANT: check 'Add python.exe to PATH' during installation,"
    Write-Host "then run install.bat again."
    exit 1
}

# --- 2. Create virtual environment ---
$venv = Join-Path $root ".venv"
if (-not (Test-Path (Join-Path $venv "Scripts\python.exe"))) {
    Write-Host "Creating virtual environment..."
    & $pythonCmd @pythonArgs -m venv $venv
}
$py = Join-Path $venv "Scripts\python.exe"

# --- 3. Install dependencies ---
Write-Host "Upgrading pip..."
& $py -m pip install --upgrade pip --quiet
Write-Host "Installing dependencies (playwright, openpyxl, httpx, scrapy)..."
& $py -m pip install -r (Join-Path $root "requirements.txt")

# --- 4. Smoke test: imports ---
Write-Host "Checking installation..."
& $py -c "import playwright, openpyxl, httpx, scrapy, tkinter; print('Dependencies OK')"
if ($LASTEXITCODE -ne 0) { throw "Dependency check failed." }

# --- 5. Desktop shortcuts ---
Write-Host "Creating desktop shortcuts..."
& powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $root "create_shortcuts.ps1")

# --- 6. Chrome check (needed by Avito and Yandex Maps and HH parsers) ---
$chromeFound = $false
foreach ($chrome in @(
    "$env:ProgramFiles\Google\Chrome\Application\chrome.exe",
    "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe",
    "$env:LOCALAPPDATA\Google\Chrome\Application\chrome.exe")) {
    if (Test-Path $chrome) { $chromeFound = $true; break }
}
Write-Host ""
if ($chromeFound) {
    Write-Host "Google Chrome: found." -ForegroundColor Green
} else {
    Write-Host "WARNING: Google Chrome not found." -ForegroundColor Yellow
    Write-Host "Install Chrome from https://www.google.com/chrome/ - the parsers need it."
}

Write-Host ""
Write-Host "=== Done! ===" -ForegroundColor Green
Write-Host "Desktop shortcuts created:"
Write-Host "  - Avito Parser"
Write-Host "  - Yandex Maps Lead Parser"
Write-Host "  - HH.ru Parser"
Write-Host "  - Rusprofile Parser"

