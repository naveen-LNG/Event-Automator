# Assembles pybackend_dist/python: the official python.org embeddable
# interpreter (PSF-signed, avoids AV/EDR false positives that unsigned
# PyInstaller-compiled exes trigger) + openpyxl/et_xmlfile + our scripts.
$ErrorActionPreference = "Stop"

$desktopRoot = Split-Path -Parent $PSScriptRoot
$backendSrc = Join-Path $desktopRoot "pybackend-src"
$dist = Join-Path $desktopRoot "pybackend_dist\python"
$pyVersion = "3.14.6"
$zipUrl = "https://www.python.org/ftp/python/$pyVersion/python-$pyVersion-embed-amd64.zip"
$tmpZip = Join-Path $env:TEMP "python-embed-$pyVersion.zip"

if (Test-Path (Join-Path $desktopRoot "pybackend_dist")) {
    Remove-Item (Join-Path $desktopRoot "pybackend_dist") -Recurse -Force
}
New-Item -ItemType Directory -Force -Path $dist | Out-Null

Write-Host "Downloading embeddable Python $pyVersion..."
Invoke-WebRequest -Uri $zipUrl -OutFile $tmpZip
Expand-Archive -Path $tmpZip -DestinationPath $dist -Force

$sig = Get-AuthenticodeSignature (Join-Path $dist "python.exe")
if ($sig.Status -ne "Valid") {
    throw "python.exe signature check failed: $($sig.StatusMessage) - refusing to bundle an unverified interpreter."
}
Write-Host "python.exe signature verified: $($sig.SignerCertificate.Subject)"

# Enable Lib\site-packages (disabled by default in embeddable distributions)
$pthFile = Get-ChildItem $dist -Filter "python3*._pth" | Select-Object -First 1
Add-Content -Path $pthFile.FullName -Value "Lib\site-packages"

$sitePkgs = Join-Path $dist "Lib\site-packages"
New-Item -ItemType Directory -Force -Path $sitePkgs | Out-Null

$pkgDir = Join-Path $env:TEMP "pyembed_pkgs_$pyVersion"
if (Test-Path $pkgDir) { Remove-Item $pkgDir -Recurse -Force }
New-Item -ItemType Directory -Force -Path $pkgDir | Out-Null

Write-Host "Fetching openpyxl + et_xmlfile (pure-Python wheels, no compiler needed)..."
python -m pip download --no-deps --only-binary=:all: --python-version 314 --platform win_amd64 -d $pkgDir openpyxl et_xmlfile

Get-ChildItem $pkgDir -Filter "*.whl" | ForEach-Object {
    $asZip = [System.IO.Path]::ChangeExtension($_.FullName, ".zip")
    Copy-Item $_.FullName $asZip -Force
    Expand-Archive -Path $asZip -DestinationPath $sitePkgs -Force
}

# Scripts live flat next to python.exe: the embeddable interpreter's ._pth
# file makes sys.path fully explicit, so sibling-module imports (e.g.
# pybackend_cli.py importing event_pipeline.py) only resolve if both sit in
# the "." directory rather than a subfolder.
Copy-Item (Join-Path $backendSrc "*.py") -Destination $dist -Force

Write-Host "pybackend_dist assembled at $dist"
