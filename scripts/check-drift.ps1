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

# build-pybackend.ps1 / sync-pybackend.ps1 legitimately differ on exactly ONE
# line: they resolve the backend source via a repo-relative path here vs. a
# sibling-folder path in the live copy (the live copy sits inside the larger
# non-git SkipBo/Unity tree; this repo does not).
#
# These used to be skipped wholesale, which hid every OTHER change to them -
# that is how an added pip dependency (pypdf) sat unnoticed in the live build
# script. Now only the backendSrc line is normalised away, so any real change
# still gets flagged.
$normalisedFiles = @(
    "scripts\build-pybackend.ps1",
    "scripts\sync-pybackend.ps1"
)

function Get-NormalisedText($path) {
    (Get-Content $path) |
        ForEach-Object { if ($_ -match '^\s*\$backendSrc\s*=') { '<BACKEND_SRC_PATH>' } else { $_ } } |
        Out-String
}

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
        if ($normalisedFiles -contains $relPath) {
            # differs only on the known path line? then it is fine
            if ((Get-NormalisedText $liveFull) -eq (Get-NormalisedText $repoFull)) { return }
            $script:drift += "OUT OF SYNC (beyond the expected path line): $relPath"
            return
        }
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

# --- tests/ ---
$liveTests = Join-Path $liveDesktop "tests"
if (Test-Path $liveTests) {
    Get-ChildItem $liveTests -File | ForEach-Object {
        $rel = "tests\$($_.Name)"
        Compare-File $rel $_.FullName (Join-Path $repoRoot $rel)
    }
}

# --- ideationData/Events/TEMPLATES (the only tracked part of ideationData) ---
$liveTemplates = Join-Path $liveDesktop "ideationData\Events\TEMPLATES"
if (Test-Path $liveTemplates) {
    Get-ChildItem $liveTemplates -File | ForEach-Object {
        $rel = "ideationData\Events\TEMPLATES\$($_.Name)"
        Compare-File $rel $_.FullName (Join-Path $repoRoot $rel)
    }
}

# --- .claude/ (project-local Claude Code skill for driving the app in tests) ---
$liveClaude = Join-Path $liveDesktop ".claude"
if (Test-Path $liveClaude) {
    Get-ChildItem $liveClaude -File -Recurse | ForEach-Object {
        $rel = ".claude\" + $_.FullName.Substring($liveClaude.Length + 1)
        Compare-File $rel $_.FullName (Join-Path $repoRoot $rel)
    }
}

# --- ideationData/StyleReference (art_reference.py reads these specs + images) ---
$liveStyle = Join-Path $liveDesktop "ideationData\StyleReference"
if (Test-Path $liveStyle) {
    Get-ChildItem $liveStyle -File -Recurse | ForEach-Object {
        $rel = "ideationData\StyleReference\" + $_.FullName.Substring($liveStyle.Length + 1)
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

# --- Catch-all: any top-level content in the live tool that this script does
# not already classify. Without this, a brand-new folder is invisible to the
# drift check entirely - which is exactly how ideationData\Roadmap\Roadmap.xlsx
# (required by the Sprint Board tab) shipped absent for months and only turned
# up when the tab failed on someone else's PC. Everything must be either
# compared above or deliberately listed here with a reason.
$knownUntracked = @{
    "node_modules"     = "installed dependency, never committed"
    "release"          = "build output"
    "pybackend_dist"   = "build output (assembled interpreter)"
    "package-lock.json" = "not currently tracked in this repo"
    ".git"             = "git metadata"
    ".gitignore"       = "compared implicitly - repo-managed"
    "README.md"        = "repo-managed docs"
    "githooks"         = "repo-managed hooks"
    "pybackend-src"    = "mirror of the authoritative PythonBackend folder"
}
$classifiedTracked = @("main.js", "preload.js", "package.json", "renderer", "build", "scripts", "tests", "ideationData", ".claude")
Get-ChildItem $liveDesktop -Force | ForEach-Object {
    $n = $_.Name
    if ($classifiedTracked -contains $n) { return }
    if ($knownUntracked.ContainsKey($n)) { return }
    $drift += "UNCLASSIFIED (in live tool, not checked or ignored by this script): $n"
}

# ideationData is only partly tracked, so classify its subfolders explicitly too.
$ideationKnown = @{
    "Events" = "only Events\TEMPLATES is tracked; per-event output is generated"
    "Roadmap" = "Feishu roadmap export - tracked (repo is private); compared below"
    "StyleReference" = "art render specs + reference images read by art_reference.py; compared above"
}
$liveIdeation = Join-Path $liveDesktop "ideationData"
if (Test-Path $liveIdeation) {
    Get-ChildItem $liveIdeation -Force | ForEach-Object {
        if (-not $ideationKnown.ContainsKey($_.Name)) {
            $drift += "UNCLASSIFIED under ideationData (decide: track or ignore): $($_.Name)"
        }
    }
    # The Sprint Board tab hard-requires this file at runtime, and it ships in
    # the build, so an out-of-sync copy here means shipping a stale roadmap.
    $rel = "ideationData\Roadmap\Roadmap.xlsx"
    $liveRoadmap = Join-Path $liveIdeation "Roadmap\Roadmap.xlsx"
    if (Test-Path $liveRoadmap) {
        Compare-File $rel $liveRoadmap (Join-Path $repoRoot $rel)
    } else {
        Write-Warning "$rel is absent from the live tool - the Sprint Board tab will not work in builds made from this tree."
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
