// UI CONTRACT TEST
// The renderer can be restyled and restructured freely, but app.js reaches into
// the DOM by ID and class. This test pins that contract so a redesign that drops
// a hook fails here instead of at runtime in front of a designer.
//
//   node tests/ui_contract_test.js
const fs = require('fs'), path = require('path');

const R = (f) => fs.readFileSync(path.join(__dirname, '..', 'renderer', f), 'utf-8');
const js = R('app.js'), html = R('index.html'), css = R('style.css');
const main = fs.readFileSync(path.join(__dirname, '..', 'main.js'), 'utf-8');
const preload = fs.readFileSync(path.join(__dirname, '..', 'preload.js'), 'utf-8');

// Backend .py files live in two places depending on checkout shape: the git
// repo keeps a mirror at pybackend-src/, while the primary dev machine has the
// authoritative copy in a sibling ../../EventAutomation/PythonBackend (shared
// with the Unity tool). Resolve whichever exists so this test runs from a
// plain clone as well as the dev tree.
const BACKEND_DIR = [
  path.join(__dirname, '..', 'pybackend-src'),
  path.join(__dirname, '..', '..', 'EventAutomation', 'PythonBackend'),
].find((d) => fs.existsSync(path.join(d, 'pybackend_cli.py')));
if (!BACKEND_DIR) { console.error('FAIL: cannot locate the backend .py sources'); process.exit(1); }
const backendFile = (name) => fs.readFileSync(path.join(BACKEND_DIR, name), 'utf-8');

let failures = 0;
function check(ok, msg, detail) {
  if (ok) { console.log('OK  ' + msg); }
  else { console.error('FAIL ' + msg + (detail ? '\n     ' + detail : '')); failures++; }
}

// ---- 1. every ID app.js queries must exist in the markup -------------------
const sels = new Set();
for (const m of js.matchAll(/\$\$?\(\s*["`']([^"`']+)["`']\s*\)/g)) sels.add(m[1]);
const wantIds = [...sels].filter(s => /^#[\w-]+$/.test(s)).map(s => s.slice(1)).sort();
const haveIds = new Set([...html.matchAll(/\bid="([^"]+)"/g)].map(m => m[1]));
const missingIds = wantIds.filter(i => !haveIds.has(i));
check(missingIds.length === 0,
  `all ${wantIds.length} element IDs referenced by app.js exist in index.html`,
  missingIds.length ? 'missing: ' + missingIds.join(', ') : '');

// ---- 2. IDs built by template literal (#ideation-feedback-${i}) ------------
// These are produced by app.js itself, so only the static prefix matters:
// the container they get injected into must exist.
for (const containerId of ['ideation-results', 'artreq-results']) {
  check(haveIds.has(containerId), `dynamic-content host #${containerId} exists`);
}

// ---- 3. navigation contract ------------------------------------------------
const navOrder = [...html.matchAll(/class="[^"]*menu-btn[^"]*"[^>]*data-view="([^"]+)"/g)].map(m => m[1]);
const EXPECTED_NAV = [
  'sprint-view', 'ideation-view', 'main-events-view',
  'minigames-view', 'battlepass-view', 'holiday-bp-view', 'decoration-shop-view',
  'streak-challenge-view', 'weekly-gift-view', 'full-screen-gift-view',
  'trivia-challenge-view', 'endless-treasure-view', 'chapter-view',
  'profile-view', 'settings-view',
];
check(JSON.stringify(navOrder) === JSON.stringify(EXPECTED_NAV),
  `sidebar lists all 15 destinations in order (Sprint -> Ideation -> Main Events -> 10 regular -> Profile -> Settings)`,
  'got: ' + JSON.stringify(navOrder));

// every data-view must have a matching <section class="... view">
const viewIds = [...html.matchAll(/<section[^>]*\bid="([^"]+)"[^>]*class="([^"]*)"/g)]
  .filter(m => /\bview\b/.test(m[2])).map(m => m[1]);
const orphanNav = navOrder.filter(v => !viewIds.includes(v));
check(orphanNav.length === 0, `every nav button has a matching <section class="view">`,
  orphanNav.length ? 'no pane for: ' + orphanNav.join(', ') : '');
// placeholder-view is reached programmatically -- switchView(id, title) retitles
// it for tabs that aren't built yet -- so it has no sidebar button by design.
const PROGRAMMATIC_VIEWS = new Set(['placeholder-view']);
const orphanView = viewIds.filter(v => !navOrder.includes(v) && !PROGRAMMATIC_VIEWS.has(v));
check(orphanView.length === 0, `every view pane is reachable (sidebar, or programmatic by design)`,
  orphanView.length ? 'unreachable: ' + orphanView.join(', ') : '');
check(viewIds.includes('placeholder-view') && /id="placeholder-title"/.test(html),
  'placeholder-view + #placeholder-title survive (switchView retitles them)');

// ---- 4. the collapsible group app.js wires up -------------------------------
check(/id="nav-group-regular"/.test(html), 'collapsible group #nav-group-regular present');
check(/id="nav-group-regular-items"/.test(html), '#nav-group-regular-items present');
check(/id="nav-group-regular-count"/.test(html), '#nav-group-regular-count present');
check(/class="[^"]*\bnav-group-header\b/.test(html),
  '.nav-group-header exists (app.js binds the collapse click to it)');
// the 9 regular-update buttons must live INSIDE the group container
const groupBlock = html.slice(html.indexOf('id="nav-group-regular-items"'));
const groupEnd = groupBlock.indexOf('</div>', groupBlock.indexOf('data-view="chapter-view"'));
const inGroup = [...groupBlock.slice(0, groupEnd).matchAll(/data-view="([^"]+)"/g)].map(m => m[1]);
check(inGroup.length === 10 && inGroup[0] === 'minigames-view' && inGroup[9] === 'chapter-view',
  'all 10 Regular Updates buttons nest inside the group container',
  'found ' + inGroup.length + ': ' + inGroup.join(', '));

// ---- 5. classes app.js emits must have styling ------------------------------
// Behaviour-only hooks are queried by JS but intentionally carry no rule.
const BEHAVIOUR_ONLY = new Set([
  'sp-go', 'ideation-select-btn', 'ideation-regen-btn',
  'artreq-select-btn', 'artreq-regen-btn', 'lesson-retract-btn',
  'artreq-images-btn', 'zoneimg-regen-btn',
  'status-', 'three',
]);
const used = new Set();
for (const m of js.matchAll(/class="([^"$]*)"/g)) m[1].split(/\s+/).filter(Boolean).forEach(c => used.add(c));
for (const m of js.matchAll(/className\s*=\s*[`"']([^`"'$]*)/g)) m[1].split(/\s+/).filter(Boolean).forEach(c => used.add(c));
for (const m of js.matchAll(/classList\.(?:add|remove|toggle|contains)\(\s*[`"']([\w-]+)/g)) used.add(m[1]);
for (const m of html.matchAll(/class="([^"]*)"/g)) m[1].split(/\s+/).filter(Boolean).forEach(c => used.add(c));
const styled = new Set([...css.matchAll(/\.([\w-]+)/g)].map(m => m[1]));
const unstyled = [...used].filter(c => !styled.has(c) && !BEHAVIOUR_ONLY.has(c)).sort();
check(unstyled.length === 0,
  `all ${used.size} classes used by the app have a rule in style.css`,
  unstyled.length ? 'unstyled: ' + unstyled.join(' ') : '');

// ---- 6. the states app.js toggles must be styled ----------------------------
// Derived from what app.js actually toggles, so this can't drift into checking
// for a class nobody uses.
const toggled = new Set([...js.matchAll(/classList\.(?:add|remove|toggle)\(\s*[`"']([\w-]+)/g)].map(m => m[1]));
for (const cls of [...toggled].sort()) {
  check(new RegExp('\\.' + cls + '\\b').test(css), `state class .${cls} (toggled by app.js) is styled`);
}
// switchView flips inline display; the stylesheet must not fight it with
// a !important display rule on .view
check(!/\.view[^{]*\{[^}]*display[^;]*!important/.test(css),
  '.view has no !important display rule that would break switchView()');

// ---- 6b. reference-image feature is wired end to end ------------------------
// Four separate files have to agree for this to work at all, and a break in
// any one of them is silent at build time.
const mainJs = fs.readFileSync(path.join(__dirname, '..', 'main.js'), 'utf-8');
const preloadJs = fs.readFileSync(path.join(__dirname, '..', 'preload.js'), 'utf-8');

for (const id of ['comfy-url-input', 'comfy-test-btn', 'comfy-test-result']) {
  check(haveIds.has(id), `Settings exposes #${id}`);
}
const CHANNELS = [['comfy:check', 'comfyCheck'],
                  ['artreq:images', 'artReqImages'],
                  ['artreq:regenImage', 'artReqRegenImage']];
for (const [channel, bridge] of CHANNELS) {
  check(mainJs.includes(`ipcMain.handle('${channel}'`), `main.js handles '${channel}'`);
  check(preloadJs.includes(`'${channel}'`) && preloadJs.includes(bridge + ':'),
    `preload.js bridges '${channel}' as api.${bridge}`);
  check(js.includes('window.api.' + bridge + '('), `app.js calls api.${bridge}()`);
}
check(/comfyUrl:/.test(mainJs), 'main.js gives comfyUrl a config default');
// The per-image feedback box and the apply-to-all toggle are the whole point
// of the feature; a refactor that drops either leaves it silently useless.
check(/zoneimg-fb-\$\{idx\}-\$\{z\.zone_number\}/.test(js),
  'each zone image gets its own feedback textarea');
check(/zoneimg-all-\$\{idx\}-\$\{z\.zone_number\}/.test(js),
  'each zone image gets an apply-to-all checkbox');
check(/applyToAll/.test(js) && /apply_to_all/.test(mainJs),
  'the apply-to-all flag reaches the backend payload');

// ---- 6c. Chapter Update is wired end to end --------------------------------
// Same four-file agreement as above, and just as silent when it breaks.
for (const [channel, bridge] of [['chapter:get', 'chapterGet'], ['chapter:run', 'chapterRun']]) {
  check(mainJs.includes(`ipcMain.handle('${channel}'`), `main.js handles '${channel}'`);
  check(preloadJs.includes(`'${channel}'`) && preloadJs.includes(bridge + ':'),
    `preload.js bridges '${channel}' as api.${bridge}`);
  check(js.includes('window.api.' + bridge + '('), `app.js calls api.${bridge}()`);
}
// The dispatcher has to know both script ids or the spawn fails at runtime only.
const cliPy = backendFile('pybackend_cli.py');
for (const id of ['get_chapter_info', 'update_chapter']) {
  check(cliPy.includes(`"${id}"`), `pybackend_cli.py dispatches '${id}'`);
}
// QA Validation is the sign-off gate: it must exist and must run last, after
// every step whose output it audits.
const chapterPy = backendFile('chapter_update.py');
const stepKeys = [...chapterPy.matchAll(/\(\s*"(\w+)",\s*"[^"]+",\s*chapter_core\./g)].map(m => m[1]);
check(stepKeys[stepKeys.length - 1] === 'qa',
  'QA Validation is the last step in the chapter workflow',
  'order: ' + stepKeys.join(' -> '));
// Full 10-step parity with the standalone tool, in workflow order.
const EXPECTED_STEPS = ['magic','upload','validate','reward','art','asset','map','import','sim','balance','qa'];
check(JSON.stringify(stepKeys) === JSON.stringify(EXPECTED_STEPS),
  'all 10 chapter steps + QA are registered in workflow order',
  'got: ' + stepKeys.join(' -> '));
// tools/ sits outside the branch checkout, so these must be configurable.
for (const id of ['chapter-lvlpy-input', 'browse-chapter-lvlpy-btn',
                  'chapter-mapbuilder-input', 'browse-chapter-mapbuilder-btn']) {
  check(haveIds.has(id), `Settings exposes #${id}`);
}
for (const key of ['chapterLvlPyPath', 'chapterMapBuilderPath']) {
  check(new RegExp(key + ':').test(mainJs), `main.js gives ${key} a config default`);
  check(js.includes(key), `app.js reads/persists ${key}`);
}
check(/lvl_py:\s*cfg\.chapterLvlPyPath/.test(js) && /map_builder:\s*cfg\.chapterMapBuilderPath/.test(js),
  'the run payload forwards both tool paths to the backend');
check(/CHAPTER_STEP_DEFAULT\s*=\s*\[[^\]]*'qa'/.test(js),
  'QA Validation is ticked by default in the UI');

// The simulation folder is a machine-level tool path: browsable and persisted
// like unityExePath / robotCalcPath, not retyped per run.
for (const id of ['chapter-sim-path', 'chapter-sim-browse-btn', 'chapter-sim-reset-btn']) {
  check(haveIds.has(id), `Chapter tab exposes #${id}`);
}
check(/chapterSimPath:/.test(mainJs), 'main.js gives chapterSimPath a config default');
check(/saveConfig\(\{\s*chapterSimPath/.test(js), 'app.js persists the simulation folder');
check(/sim_path:\s*cfg\.chapterSimPath/.test(js),
  'the run payload sends the stored simulation folder, not a transient field');

// Chapter numbers must stay DETECTED, never typed: a text input for them would
// reintroduce exactly the internal-vs-display offset bug the backend removes.
check(!/id="chapter-(old|new)-input"/.test(html),
  'chapter numbers are read from the workbook, not entered by hand');

// ---- 6b. the Art Requirement can start from an EXISTING One-Pager ----------
// Before this, the Art Requirement card only ever appeared after generating
// throwaway themes and marking one Final in the same session. A designer who
// already had a signed-off One-Pager - their own from last week, or a
// colleague's in the shared folder - had no route to it at all.
check(/id="ideation-mode-new"/.test(html) && /id="ideation-mode-existing"/.test(html),
  'the Ideation view offers both starting points');
check(/id="ideation-new-card"/.test(html) && /id="ideation-existing-card"/.test(html),
  'each starting point has its own card to show or hide');
check(/id="ideation-onepager-list"/.test(html),
  'the existing-One-Pager card has somewhere to list what is on disk');
check(/ideationListOnePagers/.test(R('../preload.js')) && /ideationReadOnePager/.test(R('../preload.js')),
  'preload exposes listing and reading One-Pagers');
check(/ideation:listOnePagers/.test(main) && /ideation:readOnePager/.test(main),
  'main.js handles both');
check(main.includes('startsWith(path.resolve(root)'),
  'reading a One-Pager is confined to the configured ideation folder - a path from '
  + 'the renderer is not a licence to read anywhere on disk');
check(/ideation_doc_text/.test(main),
  'the document text is forwarded to the backend payload');

const artreq = backendFile('art_requirement_generator.py');
check(artreq.includes('payload.get("ideation_doc_text")'),
  'the backend reads a One-Pager document as an alternative source');
check(artreq.includes('if not ideation_theme.get("theme_name") and not ideation_doc_text:'),
  'and only errors when NEITHER a theme nor a document was supplied');
check(/BEGIN ONE-PAGER/.test(artreq),
  'the document is fenced in the prompt so the model can tell brief from instruction');
check(/do not invent a different theme/.test(artreq),
  'and is framed as already signed off, not as raw inspiration');

// Exactly one source may be armed. Switching mode has to clear the other, or the
// Art Requirement is built from whichever the click handler happens to check first.
check(js.includes('ideationPickedOnePager = null;') && js.includes('applyIdeationMode'),
  'switching starting point clears whatever the other route had armed');
check(js.includes('const usingDoc = !!ideationPickedOnePager;'),
  'the generate handler picks its source explicitly rather than by fallback');

// ---- 6c. presence: who else has this event open ----------------------------
// Two designers on the same shared folder could both work an event and neither
// would know. Advisory only - nothing blocks, nothing has to be released.
check(/id="ideation-presence"/.test(html), 'the Ideation view has a presence banner');
check(/role="status"/.test(html) && /aria-live="polite"/.test(html),
  'and it announces itself to a screen reader without stealing focus');
check(main.includes("ipcMain.handle('presence:announce'")
   && main.includes("ipcMain.handle('presence:list'")
   && main.includes("ipcMain.handle('presence:clear'"),
  'main.js handles check-in, read and check-out');
check(/presenceAnnounce/.test(preload) && /presenceList/.test(preload) && /presenceClear/.test(preload),
  'preload exposes all three');

// The TTL is the whole safety story: a crashed app or a slept laptop leaves a
// record behind, and it has to expire on its own. A lock needing manual release
// is the failure mode this design exists to avoid.
check(/PRESENCE_TTL_MS/.test(main), 'a record expires on its own');
check(main.includes('(now - rec.lastSeen) > PRESENCE_TTL_MS'),
  'and a stale record is filtered out on read, so nobody ever clears a lock by hand');
check(main.includes('rec.authorId === cfg.authorId'),
  'you are never shown as working alongside yourself');
check(/before-quit/.test(main),
  'leaving cleanly checks out, so the banner is right sooner than the TTL alone');

// Heartbeat must be well inside the timeout or a slow OneDrive sync drops a
// working colleague off the banner.
const beat = /PRESENCE_BEAT_MS = (\d+) \* 1000/.exec(js);
check(!!beat, 'the renderer checks in on a timer');
if (beat) {
  check(Number(beat[1]) * 1000 * 4 <= 5 * 60 * 1000,
    'and does so at least four times inside the 5-minute timeout',
    'beat ' + beat[1] + 's');
}
check(/viewId === 'ideation-view'\) startPresence\(\); else stopPresence\(\)/.test(js),
  'presence follows the view - you are not reported on an event you navigated away from');
check(/presenceAnnouncedType !== e\.target\.value/.test(js),
  'switching event type checks out of the old one first');

// ---- 7. stylesheet is actually linked, and cache-busted ---------------------
const link = /href="style\.css(\?v=(\d+))?"/.exec(html);
check(!!link, 'index.html links style.css');
check(!!(link && link[2]), 'stylesheet link carries a ?v= cache-buster',
  'Electron caches aggressively; bump ?v= when style.css changes');

console.log('');
if (failures) { console.error(`${failures} CONTRACT CHECK(S) FAILED`); process.exit(1); }
console.log('UI CONTRACT INTACT');
