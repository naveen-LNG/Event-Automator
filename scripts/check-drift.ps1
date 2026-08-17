# Compares the live working tree (C:\SkipBo\tools\EventAutomationDesktop and
# the sibling PythonBackend folder) against what's actually committed in this
# git repo. Run this before every push - the repo has drifted out of sync
# with the live tool multiple times because changes land directly on the live
# tree (no git tracking there) while this repo is a separate checkout that
# only gets updated when someone remembers to copy files across.
#
# Exit code 0 = clean, nothing to sync. Exit code 1 = drift found, listed below.
#
# NOTE: the live paths below are hardcoded to this machine's known layout
# (same assumption make-icon.py already makes). On a machine without that
# exact layout, this script skips with a warning instead of failing - it's a
# convenience check for this specific dev setup, not a portable CI gate.

$ErrorActionPreference = "Stop"
$repoRoot = Split-Path -Parent $PSScriptRoot

$liveDesktop = "C:\SkipBo\tools\EventAutomationDesktop"
$liveBackend = "C:\SkipBo\tools\EventAutomation\PythonBackend"

if (-not (Test-Path $liveDesktop) -or -not (Test-Path $liveBackend)) {
    Write-Warning "Live source tree not found at the expected path(s) - skipping drift check (this is expected on any machine other than the primary dev box)."
    exit 0
}

# Files/dirs that are EXPECTED to differ, with the reason why - don't flag these.
# (build-pybackend.ps1 / sync-pybackend.ps1 resolve pybackend-src via a
# repo-relative path here vs. a sibling-folder path in the live copy, because
# the live copy sits inside the larger non-git SkipBo/Unity project tree and
# this repo doesn't.)
$expectedDiffs = @(
    "scripts\build-pybackend.ps1",
    "scripts\sync-pybackend.ps1"
)

$drift = @()

function Compare-File($relPath, $liveFull, $repoFull) {
    if (-not (Test-Path $liveFull)) { return }  # live-only additions handled by caller
    if (-not (Test-Path $repoFull)) {
        $script:drift += "MISSING FROM REPO: $relPath"
        return
    }
    $liveHash = (Get-FileHash $liveFull -Algorithm SHA256).Hash
    $repoHash = (Get-FileHash $repoFull -Algorithm SHA256).Hash
    if ($liveHash -ne $repoHash) {
        if ($expectedDiffs -contains $relPath) { return }
        $script:drift += "OUT OF SYNC: $relPath"
    }
}

# --- Desktop app core files ---
foreach ($f in @("main.js", "preload.js", "package.json")) {
    Compare-File $f (Join-Path $liveDesktop $f) (Join-Path $repoRoot $f)
}

# --- renderer/ (every file) ---
Get-ChildItem (Join-Path $liveDesktop "renderer") -File -Recurse | ForEach-Object {
    $rel = "renderer\" + $_.FullName.Substring((Join-Path $liveDesktop "renderer").Length + 1)
    Compare-File $rel $_.FullName (Join-Path $repoRoot $rel)
}

# --- build/ (icon assets) ---
$liveBuildDir = Join-Path $liveDesktop "build"
if (Test-Path $liveBuildDir) {
    Get-ChildItem $liveBuildDir -File | ForEach-Object {
        $rel = "build\$($_.Name)"
        Compare-File $rel $_.FullName (Join-Path $repoRoot $rel)
    }
}

# --- scripts/ ---
Get-ChildItem (Join-Path $liveDesktop "scripts") -File | ForEach-Object {
    $rel = "scripts\$($_.Name)"
    Compare-File $rel $_.FullName (Join-Path $repoRoot $rel)
}

# --- ideationData/Events/TEMPLATES (the only tracked part of ideationData) ---
$liveTemplates = Join-Path $liveDesktop "ideationData\Events\TEMPLATES"
if (Test-Path $liveTemplates) {
    Get-ChildItem $liveTemplates -File | ForEach-Object {
        $rel = "ideationData\Events\TEMPLATES\$($_.Name)"
        Compare-File $rel $_.FullName (Join-Path $repoRoot $rel)
    }
}

# --- Backend .py source (excludes Unity .meta/log/debug/cache artifacts) ---
Get-ChildItem $liveBackend -Filter "*.py" | ForEach-Object {
    $rel = "pybackend-src\$($_.Name)"
    Compare-File $rel $_.FullName (Join-Path $repoRoot $rel)
}
$repoBackendSrc = Join-Path $repoRoot "pybackend-src"
if (Test-Path $repoBackendSrc) {
    $liveNames = (Get-ChildItem $liveBackend -Filter "*.py").Name
    Get-ChildItem $repoBackendSrc -Filter "*.py" | Where-Object { $liveNames -notcontains $_.Name } | ForEach-Object {
        $drift += "IN REPO BUT DELETED FROM LIVE SOURCE: pybackend-src\$($_.Name)"
    }
}

if ($drift.Count -eq 0) {
    Write-Host "No drift detected - repo matches the live tool." -ForegroundColor Green
    exit 0
} else {
    Write-Host "DRIFT DETECTED - the following files differ between the live tool and this repo:" -ForegroundColor Red
    $drift | ForEach-Object { Write-Host "  - $_" -ForegroundColor Yellow }
    Write-Host ""
    Write-Host "Sync the live files into this repo before pushing." -ForegroundColor Red
    exit 1
}
