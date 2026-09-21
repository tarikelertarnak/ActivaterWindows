$ErrorActionPreference = "Stop"

$builder = Split-Path -Parent $MyInvocation.MyCommand.Path
$root    = Split-Path -Parent $builder
$out     = Join-Path $root "OfficeSetupWizard.exe"

Set-Location $builder

Write-Host ""
Write-Host "=== Office Setup Wizard - Build v2.0.0 ===" -ForegroundColor Cyan
Write-Host ""

python -m pip install --upgrade pip pyinstaller

python -m PyInstaller --noconfirm --clean office_setup_wizard.spec

Copy-Item (Join-Path $builder "dist\OfficeSetupWizard.exe") $out -Force

$size = [math]::Round((Get-Item $out).Length / 1MB, 1)
Write-Host ""
Write-Host "OK -> $out ($size MB)" -ForegroundColor Green
Write-Host ""
