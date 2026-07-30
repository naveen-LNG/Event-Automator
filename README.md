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

Requires Node.js and Python 3 (with `pip install pyinstaller openpyxl`) only for
*building* — end users never need either.

```bash
npm install
npm run dist
```

This compiles the Python backend (`pybackend-src/`) into a bundled executable via
PyInstaller, then packages the Electron app with `electron-builder`. Output lands in
`release/` — both a portable `.exe` and an NSIS installer.

For local development without a full rebuild each time:

```bash
npm run build:py   # once, or whenever pybackend-src/ changes
npm start           # launches the app against the local pybackend_dist/ build
```

## Repository layout

- `main.js` / `preload.js` — Electron main process and the IPC bridge to the renderer.
- `renderer/` — the UI (HTML/CSS/JS, no framework).
- `pybackend-src/` — the Python backend (Excel/SVN automation logic), compiled at
  build time into a standalone executable so end users never need Python installed.
