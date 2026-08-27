Small fix to how Chapter Update reports dependency problems.

## The problem

On another machine, Chapter Update step 1 (Magic Builder) failed with a raw Python traceback ending in `ModuleNotFoundError: No module named 'pandas'`, then `conf_fill_ex.py exited with code 1`. Nothing said what to install or where.

## Why it happens

`conf_fill_ex.py`, `gen_captures*.py` and `save_map_data.py` are **SkipBo Level Builder / Map Editor scripts, not part of Event Automator**. They need `pandas`, which the app's bundled interpreter deliberately does not carry, so the app borrows a system Python to run them. On that machine a system Python existed but had no `pandas` installed.

So this was never an app packaging fault - it is a missing prerequisite on that PC.

## What changed

When one of those scripts fails on a missing module, the app now says exactly what to do, for example:

```
conf_fill_ex.py needs the 'pandas' package, which is missing from the
Python it was run with.

Interpreter used: C:\Python313\python.exe

Install it there, then re-run this step:
  "C:\Python313\python.exe" -m pip install pandas
```

The interpreter path is spelled out because PATH can resolve to a different Python than you expect - installing into the wrong one is the easy mistake here.

All three steps that borrow a system Python now also log which interpreter they used, so this is diagnosable from the log alone.

Failures that are *not* a missing dependency report exactly as before.

## If you hit the original error

Run the pip command the app now prints. Nothing else in this release changes - if Chapter Update already worked for you, there is no reason to update.

## Download

- **Event Automator Setup 1.6.1.exe** - installer, adds a Start Menu shortcut
- **Event Automator 1.6.1.exe** - portable, no install
