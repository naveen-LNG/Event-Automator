"""Every third-party package the bundle is supposed to carry is actually in it.

This exists because a release shipped without pandas. build-pybackend.ps1 had
pandas in its pip list and was correct; the bundle on disk was simply stale,
because packaging was run as `electron-builder` directly rather than
`npm run dist` (which rebuilds the interpreter first). Nothing caught it, so
the Chapter Update step that shells out to the Level Builder's conf_fill_ex.py
died with "bundled Python lacks pandas" on a machine that had just installed.

Run with the BUNDLED interpreter (npm test does), so it checks the artifact
that actually ships rather than whatever the dev box has installed.
"""
import importlib
import sys

# (import name, why it is bundled) - keep in step with the pip list in
# scripts/build-pybackend.ps1.
REQUIRED = [
    ("openpyxl", "Excel read/write - every config tab"),
    ("et_xmlfile", "openpyxl's writer"),
    ("anthropic", "Ideation + Art Requirement generation"),
    ("pypdf", "reads finalized ideation / art-requirement PDFs"),
    ("PIL", "art reference + generated-image processing"),
    ("numpy", "art reference + generated-image processing"),
    ("pandas", "the Level Builder's conf_fill_ex.py (Chapter Update step 2)"),
    ("rapidocr_onnxruntime", "text_detect.py - baked-in text in generated art"),
]

def main():
    print(f"Checking bundled dependencies with {sys.executable}")
    missing = []
    for name, why in REQUIRED:
        try:
            importlib.import_module(name)
        except Exception as exc:
            missing.append((name, why, exc))
            print(f"MISSING  {name:22s} {why}")
        else:
            print(f"OK       {name:22s} {why}")
    if missing:
        print("")
        print(f"{len(missing)} bundled dependency/dependencies missing from "
              f"pybackend_dist.")
        print("The interpreter bundle is stale or was built with a different "
              "pip list. Run `npm run build:py` (or package with `npm run dist`, "
              "which does it for you) and re-test.")
        sys.exit(1)
    print("")
    print("BUNDLE DEPS OK")

if __name__ == "__main__":
    main()
