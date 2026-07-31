$ErrorActionPreference = 'Stop'

$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $projectRoot
Add-Type -AssemblyName System.IO.Compression.FileSystem

$pyinstaller = Join-Path $projectRoot '.venv\Scripts\pyinstaller.exe'
if (-not (Test-Path $pyinstaller)) {
    throw 'No se encontro PyInstaller en .venv. Instale dependencias antes de compilar.'
}

$distDir = Join-Path $projectRoot 'dist'
$releaseDir = Join-Path $distDir 'release'

if (Test-Path $releaseDir) {
    Remove-Item $releaseDir -Recurse -Force
}

New-Item -ItemType Directory -Path $releaseDir | Out-Null

& $pyinstaller `
    --noconsole `
    --clean `
    --onefile `
    --name 'GeneradorTestimonios' `
    --add-data 'assets;assets' `
    --add-data 'firma.png;.' `
    main.py

& $pyinstaller `
    --noconsole `
    --clean `
    --onedir `
    --name 'GeneradorTestimoniosLauncher' `
    launcher.py

$appExe = Join-Path $distDir 'GeneradorTestimonios.exe'
$launcherDir = Join-Path $distDir 'GeneradorTestimoniosLauncher'

if (-not (Test-Path $appExe)) {
    throw 'No se genero GeneradorTestimonios.exe'
}

if (-not (Test-Path $launcherDir)) {
    throw 'No se genero la carpeta GeneradorTestimoniosLauncher'
}

$appZip = Join-Path $releaseDir 'GeneradorTestimonios-win64.zip'
$launcherZip = Join-Path $releaseDir 'GeneradorTestimoniosLauncher-win64.zip'

if (Test-Path $appZip) {
    Remove-Item $appZip -Force
}

if (Test-Path $launcherZip) {
    Remove-Item $launcherZip -Force
}

$tempAppDir = Join-Path $releaseDir 'app_zip_tmp'
if (Test-Path $tempAppDir) {
    Remove-Item $tempAppDir -Recurse -Force
}
New-Item -ItemType Directory -Path $tempAppDir | Out-Null
Copy-Item $appExe -Destination (Join-Path $tempAppDir 'GeneradorTestimonios.exe')

[System.IO.Compression.ZipFile]::CreateFromDirectory($tempAppDir, $appZip)
[System.IO.Compression.ZipFile]::CreateFromDirectory($launcherDir, $launcherZip)

Remove-Item $tempAppDir -Recurse -Force

$hashes = @()
$hashes += ((Get-FileHash $appZip -Algorithm SHA256).Hash.ToLower() + '  ' + [System.IO.Path]::GetFileName($appZip))
$hashes += ((Get-FileHash $launcherZip -Algorithm SHA256).Hash.ToLower() + '  ' + [System.IO.Path]::GetFileName($launcherZip))

$checksumsPath = Join-Path $releaseDir 'checksums.txt'
$hashes | Set-Content -Path $checksumsPath -Encoding UTF8

Write-Host ''
Write-Host 'Assets listos para GitHub Release:' -ForegroundColor Green
Write-Host "- $appZip"
Write-Host "- $launcherZip"
Write-Host "- $checksumsPath"
