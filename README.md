# Event Automator (Standalone)

A standalone Windows desktop tool that automates the repetitive parts of creating
and reopening SkipBo game events — Excel config cloning, localization, icons,
convert.json/convert_layout.lua registration, and more. No Unity Editor required.

## For everyone — just run it

1. Go to [Releases](../../releases/latest) and download **`Event Automator Setup.exe`**
   (installer, adds a Start Menu shortcut) or **`Event Automator.exe`** (portable, no
   install — just run it from wherever you put it).
2. Double-click it. That's it — no terminal, no Python install, nothing else needed.
   The Python backend is bundled inside the app.
3. First launch: open **Settings** and point the three branch paths (trunk / Misc /
   Feature) at your local checkout, if the auto-detected defaults aren't right for
   your machine.

## What it does

- **Main Events**: pick an event type (Merge/Sim/Blocks/Sort/Dessert), pick a source
  event to clone from (defaults to that type's latest), set a start date + duration,
  and run the full pipeline — proto config, local service registration, Excel
  cloning (items, shop, icons, localization, quiz, answer_challenge, asset refs,
  store, guide, sys.xlsx deep links), and final `coder_convert`.
- **Minigame Configs / Pet Battlepass / Holiday Battlepass**: dedicated tabs for
  those recurring config updates, independent of the main event pipeline.
- **Settings**: branch/workspace selection, Unity executable path (only needed for
  the one step — ProtoToLua — that's genuinely Unity-coupled; it's run headlessly),
  and a manual "Update Cache from SVN" action for the source-event history cache.

## For developers — building from source

Requires Node.js and Python 3 (with `pip` available) only for *building* — end
users never need either.

```bash
npm install
npm run dist
```

`npm run build:py` downloads the official python.org **embeddable** interpreter
(PSF-signed `python.exe` — its signature is verified during the build, and the
build fails if that check doesn't pass), drops `openpyxl`/`et_xmlfile` into its
`Lib/site-packages`, and copies `pybackend-src/*.py` alongside it — no compiler,
no PyInstaller. `npm run dist` runs that, then packages the Electron app with
`electron-builder`. Output lands in `release/` — both a portable `.exe` and an
NSIS installer.

We deliberately don't compile the backend into a custom executable (e.g. via
PyInstaller): unsigned, freshly-compiled bootloaders routinely trip antivirus/EDR
heuristics (SentinelOne, Defender) as generic-packer false positives. Shipping
the officially-signed interpreter instead avoids that class of problem entirely.

For local development without a full rebuild each time:

```bash
npm run build:py   # once, or whenever pybackend-src/ changes
npm start           # launches the app against the local pybackend_dist/ build
```

## Sprint Board — one file you must supply yourself

The **Sprint Board** tab reads the studio roadmap exported from Feishu, expected at:

```
ideationData/Roadmap/Roadmap.xlsx
```

That file is **deliberately not committed to this repo** — it contains unreleased
event names and launch dates, and this repo is public. It is therefore not
included in the installer or portable build either. Without it the Sprint Board
tab reports "No roadmap found" and every other tab works normally.

To enable it, get the current `Roadmap.xlsx` export through an internal channel
(shared drive / Feishu export) and drop it at that path:

- **Installed / portable build:** `<app folder>/resources/ideationData/Roadmap/Roadmap.xlsx`
- **Running from source:** `ideationData/Roadmap/Roadmap.xlsx` in this repo

Either the Strict or the ordinary OOXML export flavour works — the parser handles
both, so no save-as setting to remember.

## Keeping this repo in sync with the live tool

The live tool's actual working copy is not this git checkout — it lives
elsewhere on the primary dev machine (alongside the larger, non-git SkipBo/Unity
project tree the Python backend is shared with) and gets copied into this repo
before each push. That gap has caused real drift more than once: a fix landing
on the live tool without ever reaching GitHub.

`npm install` sets `core.hooksPath` to `githooks/` automatically, which enables
a `pre-push` hook that runs `scripts/check-drift.ps1` — it hashes every file
this repo tracks against the live tool's copy and **blocks the push** if
anything differs. Run it manually any time with:

```bash
npm run check:drift
```

If it reports drift, copy the listed files from the live tool into this repo
before pushing again.

## Repository layout

- `main.js` / `preload.js` — Electron main process and the IPC bridge to the renderer.
- `renderer/` — the UI (HTML/CSS/JS, no framework).
- `pybackend-src/` — the Python backend (Excel/SVN automation logic), run at
  runtime by the bundled python.org embeddable interpreter — end users never
  install Python themselves.
- `scripts/build-pybackend.ps1` — assembles the bundled interpreter + dependencies.
- `scripts/check-drift.ps1` / `githooks/pre-push` — the sync-drift guard above.
