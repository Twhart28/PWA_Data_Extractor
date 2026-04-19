$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location -LiteralPath $repoRoot

$python = Join-Path $repoRoot ".venv\\Scripts\\python.exe"
$pyinstaller = Join-Path $repoRoot ".venv\\Scripts\\pyinstaller.exe"

if (-not (Test-Path $python)) {
    throw "Missing Python runtime at $python"
}

if (-not (Test-Path $pyinstaller)) {
    throw "Missing PyInstaller at $pyinstaller"
}

Write-Host "Cleaning previous build outputs..."
Remove-Item -Recurse -Force -ErrorAction SilentlyContinue build, dist, release

Write-Host "Running syntax check..."
& $python -m py_compile app.py backend.py pwa_extractor.py

Write-Host "Building portable executable..."
& $pyinstaller .\pwa_extractor.spec --noconfirm

$iscc = Get-Command ISCC.exe -ErrorAction SilentlyContinue
if (-not $iscc) {
    $commonPaths = @(
        "C:\\Program Files (x86)\\Inno Setup 6\\ISCC.exe",
        "C:\\Program Files\\Inno Setup 6\\ISCC.exe"
    )
    foreach ($candidate in $commonPaths) {
        if (Test-Path $candidate) {
            $iscc = @{ Source = $candidate }
            break
        }
    }
}

if (-not $iscc) {
    throw "Inno Setup compiler (ISCC.exe) not found. Install Inno Setup, then rerun build_release.ps1."
}

Write-Host "Building installer..."
& $iscc.Source .\pwa_extractor_installer.iss

Write-Host ""
Write-Host "Build complete."
Write-Host "Portable EXE: $repoRoot\\dist\\pwa_extractor.exe"
Write-Host "Installer:    $repoRoot\\release\\PWA_Data_Extractor_Setup.exe"
