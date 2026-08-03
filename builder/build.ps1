$ErrorActionPreference = "Stop"

$builder = Split-Path -Parent $MyInvocation.MyCommand.Path
$root    = Split-Path -Parent $builder
$out     = Join-Path $root "OfficeSetupWizard.exe"

Set-Location $builder

python -m pip install --upgrade pip pyinstaller

python -m PyInstaller --noconfirm --clean office_setup_wizard.spec

Copy-Item (Join-Path $builder "dist\OfficeSetupWizard.exe") $out -Force
Write-Host ""
Write-Host "OK -> $out"
