# Fast dev-mode sync: copies backend .py files from source (tools/EventAutomation/
# PythonBackend) into pybackend_dist/python, WITHOUT re-downloading the Python
# interpreter or pip packages. Runs automatically before every "npm start" (see
# package.json's "prestart") so the bundled backend "npm start" actually runs can
# never silently drift from source during development - that drift is exactly what
# caused a confirmed fix to appear "not working" once before this script existed.
#
# Run "npm run build:py" instead of this script when:
#   - pybackend_dist doesn't exist yet (first-time setup - this script defers to it
#     automatically in that case, see below)
#   - You changed backend Python dependencies (openpyxl/et_xmlfile/anthropic versions)
#   - You need to refresh the bundled Python interpreter itself

$ErrorActionPreference = "Stop"

$desktopRoot = Split-Path -Parent $PSScriptRoot
$backendSrc = Join-Path (Split-Path -Parent $desktopRoot) "EventAutomation\PythonBackend"
$dist = Join-Path $desktopRoot "pybackend_dist\python"

if (-not (Test-Path $dist)) {
    Write-Host "pybackend_dist not found - running full build:py (first-time setup, this downloads Python + packages)..."
    & (Join-Path $PSScriptRoot "build-pybackend.ps1")
    exit $LASTEXITCODE
}

$srcFiles = Get-ChildItem $backendSrc -Filter "*.py"
$updated = @()

foreach ($f in $srcFiles) {
    $destFile = Join-Path $dist $f.Name
    $needsCopy = $true
    if (Test-Path $destFile) {
        $srcHash = (Get-FileHash $f.FullName -Algorithm SHA256).Hash
        $destHash = (Get-FileHash $destFile -Algorithm SHA256).Hash
        $needsCopy = $srcHash -ne $destHash
    }
    if ($needsCopy) {
        Copy-Item $f.FullName $destFile -Force
        $updated += $f.Name
    }
}

if ($updated.Count -gt 0) {
    Write-Host "pybackend sync: updated $($updated.Count) backend script(s) from source -> $($updated -join ', ')"
} else {
    Write-Host "pybackend sync: bundle already matches source, nothing to update."
}
