# Assembles pybackend_dist/python: the official python.org embeddable
# interpreter (PSF-signed, avoids AV/EDR false positives that unsigned
# PyInstaller-compiled exes trigger) + openpyxl/et_xmlfile + our scripts.
$ErrorActionPreference = "Stop"

$desktopRoot = Split-Path -Parent $PSScriptRoot
$backendSrc = Join-Path (Split-Path -Parent $desktopRoot) "EventAutomation\PythonBackend"
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

# Dependency notes:
#   openpyxl/et_xmlfile - Excel read/write (all config tabs)
#   anthropic           - Ideation + Art Requirement generation
#   pypdf               - reads finalized ideation/art-requirement PDFs
#   pillow, numpy       - art reference + generated-image processing
#   pandas              - NOT used by our code. Bundled so the Level Builder's
#                         conf_fill_ex.py can run on the bundled interpreter
#                         instead of requiring every machine to have a system
#                         Python with pandas installed. chapter_core's
#                         _resolve_interpreter prefers bundled and falls back
#                         to system Python if a script needs something else.
#                         PINNED <3 on purpose: conf_fill_ex.py is not our code
#                         and predates pandas 3, whose breaking changes (default
#                         copy-on-write, removed APIs) could change its output
#                         silently. 2.3.x has a cp314 wheel, so the pin costs
#                         nothing.
#   pywin32             - Merge to Trunk tab (xlsx_merge.py) writes into trunk workbooks
#                         through Excel COM so formatting/formula caches survive; the
#                         dry run works without it. Needs the ._pth lines added below.
#   rapidocr_onnxruntime- text_detect.py (catches baked-in text in generated art).
#                         Pulls onnxruntime + opencv + shapely: ~90MB of wheels,
#                         ~230MB installed. That is the bulk of the bundle size.
#                         Do NOT trim these: text_detect/bleed_check/icon_sheet
#                         import them at module load, so the Art tabs break
#                         outright without them.
#
# Full dependency resolution (not --no-deps) so pip pulls httpx/pydantic/etc.
# alongside anthropic itself. pydantic-core is the one compiled (non-pure-Python)
# wheel in this tree; verified cp314-win_amd64 wheels exist on PyPI as of this
# writing (pydantic-core 2.46.4). If a future transitive dep drops win_amd64/py314
# wheels, this download step fails loudly here rather than at runtime.
Write-Host "Fetching backend dependencies (openpyxl, et_xmlfile, anthropic + tree, pypdf)..."
python -m pip download --only-binary=:all: --python-version 314 --platform win_amd64 -d $pkgDir openpyxl et_xmlfile anthropic pypdf pillow numpy "pandas<3" rapidocr_onnxruntime pywin32

Get-ChildItem $pkgDir -Filter "*.whl" | ForEach-Object {
    $asZip = [System.IO.Path]::ChangeExtension($_.FullName, ".zip")
    Copy-Item $_.FullName $asZip -Force
    Expand-Archive -Path $asZip -DestinationPath $sitePkgs -Force
}

# pywin32 keeps its modules in three sub-folders that its own pywin32.pth would
# normally register - but the embeddable interpreter never runs site.py, so the
# ._pth has to list them explicitly or `import win32com` fails at runtime.
Add-Content -Path $pthFile.FullName -Value "Lib\site-packages\win32"
Add-Content -Path $pthFile.FullName -Value "Lib\site-packages\win32\lib"
Add-Content -Path $pthFile.FullName -Value "Lib\site-packages\Pythonwin"

# Scripts live flat next to python.exe: the embeddable interpreter's ._pth
# file makes sys.path fully explicit, so sibling-module imports (e.g.
# pybackend_cli.py importing event_pipeline.py) only resolve if both sit in
# the "." directory rather than a subfolder.
Copy-Item (Join-Path $backendSrc "*.py") -Destination $dist -Force

Write-Host "pybackend_dist assembled at $dist"
