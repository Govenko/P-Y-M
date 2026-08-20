$ErrorActionPreference = "Stop"
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $scriptDir

$pythonw = Join-Path $scriptDir ".venv\Scripts\pythonw.exe"
$python = Join-Path $scriptDir ".venv\Scripts\python.exe"
$app = Join-Path $scriptDir "hh_parser_app.py"

if (Test-Path $pythonw) {
    Start-Process -FilePath $pythonw -ArgumentList @($app) -WorkingDirectory $scriptDir
} elseif (Test-Path $python) {
    Start-Process -FilePath $python -ArgumentList @($app) -WorkingDirectory $scriptDir
} else {
    throw "Python environment not found. Run install.bat first."
}
