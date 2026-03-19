$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

$displayName = ([char]0x842C) + ([char]0x7528) + ([char]0x8F49) + ([char]0x6A94) + ([char]0x738B)
$internalName = "WanyongConverter"
$sourceToolsRoot = Join-Path $PSScriptRoot "tools"
$sourcePandocExe = Join-Path $sourceToolsRoot "pandoc\pandoc.exe"
if (-not (Test-Path $sourcePandocExe)) {
  throw "Missing required portable tool: tools\pandoc\pandoc.exe"
}

$portableRoot = Join-Path $PSScriptRoot ("portable\" + $internalName)
$distRoot = Join-Path $PSScriptRoot ("dist\" + $internalName)
$distExe = Join-Path $distRoot ($internalName + ".exe")
$ffmpegTargetDir = Join-Path $portableRoot "tools\ffmpeg"
$pandocTargetDir = Join-Path $portableRoot "tools\pandoc"
python -m PyInstaller `
  --noconfirm `
  --clean `
  --windowed `
  --onedir `
  --name $internalName `
  --exclude-module IPython `
  --exclude-module matplotlib `
  --exclude-module moviepy `
  --exclude-module onnxruntime `
  --exclude-module pandas `
  --exclude-module pygame `
  --exclude-module scipy `
  --exclude-module sklearn `
  --exclude-module tensorflow `
  --exclude-module torch `
  --exclude-module torchaudio `
  --exclude-module torchvision `
  --exclude-module transformers `
  --exclude-module ultralytics `
  launch_super_converter.pyw

if (Test-Path $portableRoot) {
  Remove-Item $portableRoot -Recurse -Force
}

New-Item -ItemType Directory -Force -Path $portableRoot | Out-Null
New-Item -ItemType Directory -Force -Path $ffmpegTargetDir | Out-Null
New-Item -ItemType Directory -Force -Path $pandocTargetDir | Out-Null

Copy-Item (Join-Path $distRoot "*") $portableRoot -Recurse -Force
Copy-Item (Join-Path $PSScriptRoot "README.md") (Join-Path $portableRoot "README.md") -Force
Copy-Item (Join-Path $sourceToolsRoot "pandoc\*") $pandocTargetDir -Recurse -Force
$ffmpegExe = python -c "import imageio_ffmpeg; print(imageio_ffmpeg.get_ffmpeg_exe())"
if ($LASTEXITCODE -ne 0 -or -not (Test-Path $ffmpegExe)) {
  throw "Embedded ffmpeg was not available during portable build."
}
Copy-Item $ffmpegExe (Join-Path $ffmpegTargetDir "ffmpeg.exe") -Force

$portableText = @"
$displayName

This is a fully bundled portable build.

Included tools
- tools\ffmpeg\ffmpeg.exe
- tools\pandoc\pandoc.exe

Usage
1. Run $internalName.exe
2. Move the whole folder to another Windows PC if needed
3. Delete the whole folder to remove the app
"@
$portableText | Set-Content (Join-Path $portableRoot "PORTABLE.txt") -Encoding UTF8

Write-Host ""
Write-Host ("Portable build complete: " + $portableRoot)
