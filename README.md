# Event Automator (Standalone)

A standalone Windows desktop tool that automates the repetitive parts of creating
and reopening SkipBo game events — Excel config cloning, localization, icons,
convert.json/convert_layout.lua registration, and more. No Unity Editor required.

## For everyone — just run it

1. Go to [Releases](../../releases/latest) and download **`Event-Automator-Setup-X.exe`**
   (installer, adds a Start Menu shortcut, updates itself). There is no portable
   build any more.
2. Double-click it. That's it — no terminal, no Python install, nothing else needed.
   The Python backend is bundled inside the app.
3. First launch: open **Settings** and point the three branch paths (trunk / Misc /
   Feature) at your local checkout, if the auto-detected defaults aren't right for
   your machine.

## Updates

From v1.31.0, the tool checks this repo's latest release on every launch. When a newer
version exists it asks to download it; once downloaded, **Restart now** installs it
(or it installs the next time the tool closes). Copies older than v1.31.0 have no
checker: install the latest version by hand once.

To publish an update, the release must carry **all four** of these files, under
these exact names: `latest.yml`, `Event-Automator-Setup-X.exe` and
`Event-Automator-Setup-X.exe.blockmap` from `release/` (or installed copies won't see
the update), and `event-automator.mcpb` (the app's "extension not set up" popup links to
the latest release's copy). Bump `version` in `package.json` first; the checker compares
that number.

## Claude Desktop extension (.mcpb)

Localization > Translate and the Art Requirement tools run through Claude Desktop, via the
`event-automator.mcpb` extension. On launch the app checks Claude Desktop's extensions
folder and `claude_desktop_config.json`; if neither has Event Automator, it shows a
popup with a download button (**Don't show again** hides it for good).

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
`electron-builder`. Output lands in `release/` — an NSIS installer.

We deliberately don't compile the backend into a custom executable (e.g. via
PyInstaller): unsigned, freshly-compiled bootloaders routinely trip antivirus/EDR
heuristics (SentinelOne, Defender) as generic-packer false positives. Shipping
the officially-signed interpreter instead avoids that class of problem entirely.

For local development without a full rebuild each time:

```bash
npm run build:py   # once, or whenever pybackend-src/ changes
npm start           # launches the app against the local pybackend_dist/ build
```

## Sprint Board and the roadmap

The **Sprint Board** tab joins the studio roadmap (the plan) against what is
actually configured in the design workbooks (the reality). The roadmap is the
Feishu export committed at `ideationData/Roadmap/Roadmap.xlsx`, and it ships
inside the installer, so the tab works with no setup.

Either the Strict or the ordinary OOXML export flavour works — the parser
handles both, so there is no save-as setting to remember.

To refresh the plan, re-export from Feishu over that file and commit it; then
rebuild so the new export is picked up by installed copies.

> **This repository is private, and must stay private.** The roadmap holds
> unreleased event names, launch dates and the monetisation schedule. If the
> repo is ever made public, remove `ideationData/Roadmap/` from both the working
> tree *and* the git history first, and re-add it to `.gitignore`.

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
