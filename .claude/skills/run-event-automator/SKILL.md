---
name: run-event-automator
description: Launch and drive the Event Automator Electron app on Windows - click buttons, read the DOM, take screenshots - using a zero-dependency CDP driver and a stub LLM server. Use it to VERIFY behaviour in the running app - confirming a change to renderer/app.js, main.js, preload.js or style.css actually works, answering "is this control broken?", screenshotting a panel, checking an IPC channel is reachable from the renderer, or diagnosing a window that opens blank. Reach for it whenever someone says run the app, open it, click through something, or "don't just read the code", and before claiming any UI change is verified, since the test suites check contracts and never launch the app. Not for editing UI code, running the Python test suites, invoking pybackend_cli directly, or building and packaging the bundle.
---

# Running and driving the Event Automator

The test suites here check contracts - that a class used by `app.js` has a CSS
rule, that a handler is bound - and they never launch the app. That gap is real:
a cost line that ran two clicks behind shipped past a green suite and was only
found by clicking. So when a change touches the renderer, run it.

## Why this exists

The generic Electron recipe assumes Linux: xvfb, tmux, Playwright's `_electron`.
This box is Windows and the project has no Playwright, so the two things that
make driving possible are bundled here:

- `scripts/cdp.js` - talks the Chrome DevTools Protocol directly. Node 22 ships
  a global `WebSocket` and `fetch`, so this needs no dependencies at all. Adding
  Playwright just to tick a checkbox is not worth it.
- `scripts/stub_ollama.js` - stands in for the LLM. Most of the interesting UI
  is behind two model calls, and neither provider is reliably reachable.

## Launch

```bash
cd C:/SkipBo/tools/EventAutomationDesktop
npm start -- --remote-debugging-port=9222
```

Run it in the background - it does not exit. `npm start` runs `sync:py` first,
which copies the backend from source into `pybackend_dist`; see the sync note at
the bottom, it matters more than it looks.

Wait for the port rather than sleeping:

```bash
until curl -s http://127.0.0.1:9222/json >/dev/null 2>&1; do sleep 3; done
```

## Drive

```bash
node scripts/cdp.js targets              # list debuggable targets
node scripts/cdp.js eval '<expression>'  # evaluate in the renderer
node scripts/cdp.js click '<selector>'   # click the first match
node scripts/cdp.js text  '<selector>'   # textContent of the first match
node scripts/cdp.js shot  out.png        # screenshot the window
```

**Look at the screenshot.** A blank frame means the app did not really start,
and every later assertion is then meaningless.

Views are switched by `[data-view='...']`, e.g. `ideation-view`,
`settings-view`, `sprint-view`. Read the ids off the page rather than guessing:

```bash
node scripts/cdp.js eval "Array.from(document.querySelectorAll('[data-view]')).map(b=>b.dataset.view).join('\n')"
```

`renderer/app.js` is an IIFE, so its state (`artReqVariations`, render helpers)
is **not** reachable from `eval`. Do not plan around poking internals - drive the
DOM the way a user does, and use the stub below when you need particular data on
screen. `window.api` IS exposed by the preload bridge, which makes it a good way
to test an IPC channel on its own:

```bash
node scripts/cdp.js eval "(async()=>{ const r = await window.api.artReqPlan({variation: {...}, passes:['scene']}); return r.stdout; })()"
```

When you assert on something asynchronous, read it after the update rather than
after a fixed sleep, or you will measure your own timing rather than the app's.

## Getting past the LLM

Ideation themes and Art Requirement variations both require a model, and the
Anthropic key is usually empty while the LAN Ollama host is often unreachable.
`stub_ollama.js` reads the JSON Schema out of each request's `format` field and
returns a conforming object, so the renderer takes exactly the path it takes with
a real model.

```bash
STUB_PORT=11533 node scripts/stub_ollama.js         # background it
STUB_LEGACY=1  STUB_PORT=11533 node scripts/stub_ollama.js
STUB_DESSERT=1 STUB_PORT=11533 node scripts/stub_ollama.js
```

Three shapes, and they take different branches through the pass picker, so a
change to it is only really checked once you have seen all three:

| mode | what it returns | what it exercises |
|---|---|---|
| default | a Merge draft: 3 scenes, 24 buildings, 4 producers | every pass but recipes |
| `STUB_LEGACY=1` | a draft from before the schema carried buildings | how the UI degrades - only `scene` is live |
| `STUB_DESSERT=1` | a Dessert draft: 4 zones, 46 buildings, 39 recipes, 11 ingredients | recipes and ingredients live, chains and producers dark |

The counts in the Dessert fixture are the shipped Dessert2608 ones (24 common +
15 secret recipes, 6 + 5 ingredients), so the picker's numbers can be checked
against something real rather than against themselves.

Note that a `location.reload()` clears the ideation state, so re-drive the whole
flow after one rather than expecting the picker to still be there.

Point the app at it through Settings, which auto-saves on `change` (there is no
save button):

```js
const el = document.querySelector('#ollama-base-url-input');
el.value = 'http://127.0.0.1:11533';
el.dispatchEvent(new Event('input',  {bubbles:true}));
el.dispatchEvent(new Event('change', {bubbles:true}));
```

Then the flow is: Ideation → **Generate Themes** → **Select as Final** →
**Generate Art Requirement Variations**. The Art Requirement card stays hidden
until a theme is finalised.

## Leave the machine as you found it

This drives a real app against a real config file, so clean up deliberately.

**Back up the config before touching Settings**, and restore it at the end:

```bash
cp "$APPDATA/event-automation-desktop/config.json" ./config.backup.json
```

Restore **after** stopping Electron, or the app rewrites it as it exits.
Then diff against the backup rather than assuming - changing one Settings field
persists *others* that were previously absent and running on their defaults, so
"I only changed the Ollama URL" is not a safe assumption:

```bash
taskkill //IM electron.exe //F
cp ./config.backup.json "$APPDATA/event-automation-desktop/config.json"
```

Stop the stub by port, since it has no window to match on:

```bash
for pid in $(netstat -ano | grep ":11533" | grep LISTENING | awk '{print $5}' | sort -u); do
  taskkill //PID $pid //F
done
```

Confirm both are gone before reporting. A stray Electron holding port 9222 makes
the next run attach to a stale window and quietly test the old code.

## The sync trap

The backend lives in two places: `tools/EventAutomation/PythonBackend` is the
source of truth, and `pybackend_dist/python` is a deployed copy. The embedded
interpreter puts **its own directory first on `sys.path`**, so anything invoking
the CLI gets the deployed copy, not your edit.

The Python test suites `sys.path.insert` the source, so they can be green while
the app runs stale code. If a backend change seems to have no effect, this is
almost always why:

```bash
npm run sync:py
```

`npm start` runs it automatically, but a CLI invocation between edits does not.
