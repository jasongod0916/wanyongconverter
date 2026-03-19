$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

python -m pip install -r requirements.txt
python -m pip install pyinstaller

python -m PyInstaller `
  --noconfirm `
  --clean `
  .\SuperConverter.spec

Write-Host ""
Write-Host "Build complete: dist\\SuperConverter.exe"
