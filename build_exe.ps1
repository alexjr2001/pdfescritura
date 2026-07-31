$ErrorActionPreference = 'Stop'

$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $projectRoot

$pyinstaller = Join-Path $projectRoot '.venv\Scripts\pyinstaller.exe'
if (-not (Test-Path $pyinstaller)) {
    throw 'No se encontro PyInstaller en .venv. Instale dependencias antes de compilar.'
}

& $pyinstaller `
    --noconsole `
    --onefile `
    --name 'GeneradorTestimonios' `
    --add-data 'assets;assets' `
    --add-data 'firma.png;.' `
    main.py

& $pyinstaller `
    --noconsole `
    --onefile `
    --name 'GeneradorTestimoniosLauncher' `
    launcher.py
