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
for (const containerId of ['ideation-results', 'artreq-results', 'loc-results']) {
  check(haveIds.has(containerId), `dynamic-content host #${containerId} exists`);
}

// ---- 3. navigation contract ------------------------------------------------
const navOrder = [...html.matchAll(/class="[^"]*menu-btn[^"]*"[^>]*data-view="([^"]+)"/g)].map(m => m[1]);
const EXPECTED_NAV = [
  'sprint-view', 'ideation-view', 'localization-view', 'main-events-view',
  'minigames-view', 'battlepass-view', 'holiday-bp-view', 'decoration-shop-view',
  'streak-challenge-view', 'weekly-gift-view', 'full-screen-gift-view',
  'trivia-challenge-view', 'endless-treasure-view', 'chapter-view',
  'xlsx-merge-view', 'profile-view', 'settings-view',
];
check(JSON.stringify(navOrder) === JSON.stringify(EXPECTED_NAV),
  `sidebar lists all 17 destinations in order (Sprint -> Ideation -> Localization -> Main Events -> 10 regular -> Merge to Trunk -> Profile -> Settings)`,
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
  'artreq-open-btn', 'artreq-docx-btn', 'artreq-reveal-btn',
  // Pre-existing gap, not introduced with the Word export: these two mark which
  // headings are feedback sections and which disclosure is the document body.
  // Both are read by JS and neither carries, or wants, a visual rule.
  'artreq-doc-disclosure', 'artreq-doc-section',
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

// ---- 6d. Merge to Trunk is wired end to end --------------------------------
for (const [channel, bridge] of [['xlsxmerge:defaults', 'xlsxMergeDefaults'],
                                 ['xlsxmerge:run', 'xlsxMergeRun'],
                                 ['xlsxmerge:reveal', 'xlsxMergeReveal']]) {
  check(mainJs.includes(`ipcMain.handle('${channel}'`), `main.js handles '${channel}'`);
  check(preloadJs.includes(`'${channel}'`) && preloadJs.includes(bridge + ':'),
    `preload.js bridges '${channel}' as api.${bridge}`);
  check(js.includes('window.api.' + bridge + '('), `app.js calls api.${bridge}()`);
}
check(cliPy.includes('"xlsx_merge"'), `pybackend_cli.py dispatches 'xlsx_merge'`);
check(fs.existsSync(path.join(BACKEND_DIR, 'xlsx_merge.py')), 'xlsx_merge.py lives in the backend source folder');
// Apply must never be the first click: it is gated on a dry run with identical settings.
check(/xmDryRunKey/.test(js) && /id="xm-apply-btn"[^>]*\bdisabled\b/.test(html),
  'Apply starts disabled and is gated on a matching dry run');
// pywin32 is what lets --apply drive Excel from the bundled interpreter. The wheel
// unpacks fine without these lines and `import win32com` fails only at runtime.
const buildPs1 = fs.readFileSync(path.join(__dirname, '..', 'scripts', 'build-pybackend.ps1'), 'utf-8');
check(/pip download[^\n]*\bpywin32\b/.test(buildPs1), 'build-pybackend.ps1 downloads pywin32');
check(/Lib\\site-packages\\win32\\lib/.test(buildPs1), 'build-pybackend.ps1 adds the pywin32 folders to the ._pth');

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

// ---- 6b2. uploading a One-Pager you already have ---------------------------
// Listing only what was already filed left a hole: a One-Pager that arrived by
// email or sits on a desktop had to be hand-copied into a path most people
// would have to be told, which is the step where it lands in the wrong folder.
check(/id="ideation-upload-onepager-btn"/.test(html), 'the picker card has an upload button');
check(main.includes("ipcMain.handle('ideation:uploadOnePager'"), 'main.js handles the upload');
check(/ideationUploadOnePager/.test(preload), 'preload exposes it');

// One list of extensions, shared by the picker, the reader and the upload, so
// they cannot disagree about what is usable.
// The rules live in lib/onepager.js so tests/onepager_test.js can reach them -
// the upload starts with a native dialog no driver can click, so anything
// downstream of it was previously only checkable by reading it.
check(main.includes("require('./lib/onepager')"),
  'the format rules come from the shared module, not a second copy in main.js');

// Word and PDF are converted rather than refused. The shipped One-Pagers ARE
// PDFs, so refusing them was the wrong answer to the right observation that
// they belong in the list.
check(main.includes('kind.convertible'),
  'reading a One-Pager converts a Word or PDF first rather than refusing it');
check(main.includes("runPythonAndScrub('doc_convert'"),
  'via the doc_convert backend, not a second parser in JS');
check(main.includes('mtimeMs >= fs.statSync(resolved).mtimeMs'),
  'and the conversion is cached until the original is edited, so re-uploading a '
  + 'corrected brief reconverts without anyone knowing a cache exists');
check(main.includes('generated.has(n)'),
  'a generated .md is hidden from the picker - it and its original are one '
  + 'document, and listing both invites picking the stale one');
const cli = backendFile('pybackend_cli.py');
check(cli.includes("script == \"doc_convert\""), 'the CLI exposes doc_convert');
const dc = backendFile('doc_convert.py');
check(dc.includes('import zipfile') && dc.includes('xml.etree'),
  '.docx is read with stdlib zipfile + ElementTree - no new dependency');
check(dc.includes('from pypdf import PdfReader'),
  '.pdf goes through pypdf, already vendored for reading shipped requirement PDFs');
check(dc.includes('No text layer in this PDF'),
  'a scanned PDF is reported rather than passed on as an empty brief');

// Images were dropped by the first version. In the shipped Merge2608 brief the
// actual zone contents exist ONLY inside mood-board pictures, so that default
// lost real content with no sign it had happened.
check(dc.includes('_write_images') && dc.includes('.assets'),
  'images are extracted into a folder beside the Markdown, not dropped');
check(dc.includes('OCR_SKIP_IF_REPEATED_OVER'),
  'and page furniture is not OCR-ed - the watermark would otherwise forward an '
  + 'email address to the model along with the brief');
check(dc.includes('import text_detect'),
  'OCR reuses the engine already vendored for checking generated art');
check(dc.includes('Text read from this image'),
  'text found inside an image is labelled as such, so it can be weighed differently');
check(dc.includes('appears %d times'),
  'a repeated image is referenced once with a count rather than once per occurrence');
check(main.includes('convertStats'), 'main.js forwards what happened to the images');
check(js.includes('image${st.images === 1'),
  'and the renderer reports it - silence about a picture-borne brief is the bug '
  + 'this path exists to fix');
check(dc.includes('_heading_level'), 'Word heading styles become Markdown levels');
check(dc.includes('_table_md'), 'and Word tables become pipe tables, not flattened prose');
check(main.includes('onePager.nextFreeName('),
  'and so does the name-clash rule, which is the one that could overwrite a colleague');
const opLib = fs.readFileSync(path.join(__dirname, '..', 'lib', 'onepager.js'), 'utf-8');
check(/READABLE = \['md', 'txt'\]/.test(opLib), 'readable formats are named once, in that module');
check(/CONVERTIBLE = \['docx', 'pdf'\]/.test(opLib),
  'and the convertible ones beside them, so STORABLE is derived rather than restated');
check(/STORABLE = READABLE\.concat\(CONVERTIBLE\)/.test(opLib),
  'STORABLE is the union of the two - a format cannot be accepted without saying '
  + 'which way it is handled');
const readerUses = main.includes('onePager.kindOf(resolved)');
const listUses = main.includes('onePager.kindOf(n).storable');
check(readerUses && listUses,
  'the reader and the picker both go through kindOf rather than their own regex',
  'reader ' + readerUses + ', list ' + listUses);
check(!/\.\(md\|txt\|pdf\)/.test(main),
  'and the old hardcoded extension regex is gone');

// A shared folder means the file already there may be a colleague's.
check(main.includes('showMessageBox') && main.includes("'Replace', 'Keep both', 'Cancel'"),
  'a name clash asks before overwriting - the folder is shared, so it may not be yours');
check(main.includes('Keep both'), 'and offers to file yours alongside instead');

// There is no "filed but useless" outcome any more, so uploading always selects.
check(!js.includes('if (!res.readable)'),
  'the filed-but-unusable branch is gone - every accepted format converts');
check(js.includes("doc.convertible ? 'Converting"),
  'a conversion says so on the button rather than leaving it dead for a few seconds');
check(js.includes('if (res.converted)'),
  'and names the .md it produced, which is the file the model was actually given');
check(js.includes('await pickOnePager(idx)'),
  'a usable upload is selected straight away - uploading is already an act of choosing');

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


// ---- Art Requirement documents on disk, and Word export ----------------------
// The tab used to show only what the current session generated, so a document
// finalised yesterday - or filed by a colleague - was invisible. This checks the
// listing exists end to end, because the failure mode is silent: a renamed
// channel leaves a button that looks fine and does nothing.
check(/id="artreq-docs-card"/.test(html) && /id="artreq-docs-list"/.test(html),
  'the Art Requirement tab has a card listing documents already on disk');
check(main.includes("ipcMain.handle('artreq:listDocs'")
   && main.includes("ipcMain.handle('artreq:readDoc'")
   && main.includes("ipcMain.handle('artreq:exportDocx'")
   && main.includes("ipcMain.handle('artreq:revealDoc'"),
  'main.js handles list, read, Word export and reveal');
check(/artReqListDocs/.test(preload) && /artReqReadDoc/.test(preload)
   && /artReqExportDocx/.test(preload) && /artReqRevealDoc/.test(preload),
  'preload exposes all four');
check(/artReqListDocs/.test(js) && /artReqExportDocx/.test(js),
  'and the renderer actually calls them');

// Every path the renderer hands to the main process is resolved and checked
// against the configured folder first. A path arriving from the renderer is not
// a licence to read, convert or reveal anything on disk.
check(/function insideArtReq/.test(main)
   && /startsWith\(base \+ path\.sep\)/.test(main),
  'those paths are confined to the Art Requirements folder before use');

// The export is a Python call, so the channel needs a dispatch entry - without
// it the button reports a bare non-zero exit and nobody can tell why.
check(backendFile('pybackend_cli.py').includes('md_to_docx')
   && backendFile('pybackend_cli.py').includes('doc_convert.docx_main()'),
  'pybackend_cli dispatches md_to_docx to the converter');

// Exporting off the file on disk rather than the variation held in memory is
// what makes what a reviewer opens in Word identical to what was saved.
check(/artReqSavedPath\) \|\| artReqDraftPaths\[idx\]/.test(js),
  'a generated variation exports the document already written to disk');

// ---- generating art for a document that was opened, not generated ------------
// The Reference Images controls used to exist only on a freshly generated
// variation card. Open a document filed last week and there was no way to render
// its art at all. The image pass needs the STRUCTURED document - blunt_steps,
// character_scale and render_aspect never appear in the rendered Markdown - so a
// .json sidecar is written on every save and read back when a document is opened.
check(/ipcMain\.handle\('artreq:docData'/.test(main)
   && /ipcMain\.handle\('artreq:saveDoc'/.test(main),
  'main.js can read a document\'s structured data and save it back');
check(/artReqDocData/.test(preload) && /artReqSaveDoc/.test(preload),
  'preload exposes both');
check(/openedDocArtHtml/.test(js) && /generateOpenedDocImages/.test(js),
  'the renderer offers an art panel on an opened document');
check(/passPickerHtml\(OPENED_DOC_IDX/.test(js),
  'and reuses the same pass picker the generated cards use, so the cost line is the same one');
check(/artReqSaveDoc\(\{[\s\S]{0,200}?variation: artReqOpenedDoc/.test(js),
  'renders are written back into the document they came from');

// Every save site must write the sidecar, or a document saved by one path gets
// the art button and one saved by another silently does not.
const argPy = backendFile('art_requirement_generator.py');
check((argPy.match(/_write_doc_data\(/g) || []).length >= 4,
  'all three save paths write the sidecar (plus its definition)',
  'found ' + (argPy.match(/_write_doc_data\(/g) || []).length);
check(/def save_doc_main/.test(argPy)
   && backendFile('pybackend_cli.py').includes('artreq_save_doc'),
  'and artreq_save_doc is dispatched');

// An in-place update must not reassign the document. Adding art to a colleague's
// document should leave their name and its review state alone.
check(/found_by = re\.search/.test(argPy) && /found_status = re\.search/.test(argPy),
  'saving in place carries the existing By: and Status: forward');
check(!/author_name: cfg\.authorName,\s*\}\), 'utf-8'\);\s*return parseJsonResult\(await runPythonAndScrub\('artreq_save_doc'/.test(main),
  'and the IPC does not override them with the current user');

// A document with no sidecar must say so rather than offering a button that
// cannot work, or parsing the Markdown back and being wrong where nobody checks.
check(/has no structured data beside it/.test(main),
  'a document without structured data explains why art cannot be generated from it');

// ---- the documents panel shows this event, once each ------------------------
// It listed every document in the folder - a Blocks folder holds every Blocks
// event ever run - and listed a .md and its exported .docx as separate rows even
// though the .md row already carries a "Show .docx" button.
check(/function eventNumberOf/.test(main),
  'main.js derives an event number from a filename so documents can be matched to a One-Pager');
check(/onePagerName, showAll/.test(main),
  'and listDocs accepts the picked One-Pager plus an escape hatch');
check(/mdStems/.test(main) && /ext === '\.docx'/.test(main),
  'a .docx exported from a listed .md is not listed a second time');
check(/onePagerName: ideationPickedOnePager \? ideationPickedOnePager\.name/.test(js),
  'the renderer passes the picked One-Pager to the listing');
check(/artreq-showall-link/.test(js) && /function toggleArtReqDocFilter/.test(js),
  'the filter is reversible from the panel itself');
check(/Showing event/.test(js) && /hiding/.test(js),
  'and says which event it is showing and how much it is holding back',
  'a filter that hides rows silently is the next bug report');
check(/artReqShowAllDocs = false;\s*\n\s*loadArtReqDocs\(\);/.test(js),
  'picking a One-Pager re-filters the panel to that event');
check(/loadArtReqDocs\(\);\s*\n\s*\}/.test(js.slice(js.indexOf('function clearPickedOnePager'))),
  'and clearing it refreshes the panel too');
// The event number must survive both spellings the studio uses: the One-Pager
// says "Block2609", the shipped Art Requirements say "Blocks2604".
check(/\\d\{4\}/.test(main.slice(main.indexOf('function eventNumberOf'),
                                 main.indexOf('function eventNumberOf') + 400)),
  'the match is on the four-digit number, not the type word, so Block2609 and Blocks2609 agree');

// ---- a document records the One-Pager it was written from -------------------
// The panel used to match documents to a One-Pager by scraping a four-digit
// number out of both filenames. That worked only because the naming happened to
// be consistent, and would have failed silently the first time someone named a
// file without its event number. The link is data now.
const argPy2 = backendFile('art_requirement_generator.py');
check(/\("Source One-Pager", doc\.get\("source_one_pager"\)/.test(argPy2),
  'the rendered header carries the source One-Pager',
  'a document gets exported and mailed around; the sidecar can be lost, the header travels');
check(/v\["source_one_pager"\] = source_name/.test(argPy2),
  'generation stamps it from the file the tool actually read');
check(/new_variation\["source_one_pager"\] = prior_source/.test(argPy2),
  'and a regeneration keeps it, since a rewrite is the same document');
check(!/"source_one_pager"/.test(JSON.stringify(
        require('child_process').execFileSync(
          path.join(__dirname, '..', 'pybackend_dist', 'python', 'python.exe'),
          [path.join(__dirname, '..', 'pybackend_dist', 'python', 'pybackend_cli.py'),
           'artreq_schemas'], { encoding: 'utf-8', maxBuffer: 32 * 1024 * 1024 }))),
  'but it is NOT in the generation schema - the model has no idea what the file is called',
  'asking a model for a filename invites it to invent one');

check(/function onePagerStem/.test(main),
  'the match ignores the extension, because a One-Pager is picked as .pdf and read as .md');
check(/onePagerStem\(source\) === onePagerStem\(onePagerName\)/.test(main),
  'and prefers the recorded source over the filename number');
check(/matchedBySource/.test(main) && /matched by recorded source/.test(js),
  'the panel says which basis it matched on, so a silent fallback is visible');
check(/source_one_pager: \{/.test(fs.readFileSync(path.join(__dirname, '..', 'mcp', 'server.js'), 'utf-8')),
  'the MCP save tool can record it too');
// ---- 7. stylesheet is actually linked, and cache-busted ---------------------
const link = /href="style\.css(\?v=(\d+))?"/.exec(html);
check(!!link, 'index.html links style.css');
check(!!(link && link[2]), 'stylesheet link carries a ?v= cache-buster',
  'Electron caches aggressively; bump ?v= when style.css changes');


// ---- 8. every local module main.js pulls in is actually PACKAGED -----------
// lib/onepager.js was required at the TOP of main.js while build.files listed
// only main.js, preload.js and renderer/** - so it would have been absent from
// the asar and the packaged app could not have started at all. Every test
// passed the whole time, because tests run against the source tree where the
// file is sitting right there. This checks the build manifest, not the disk.
{
  const appRoot = path.join(__dirname, '..');
  const pkg = JSON.parse(fs.readFileSync(path.join(appRoot, 'package.json'), 'utf-8'));
  const globs = ((pkg.build && pkg.build.files) || []).filter(g => !g.startsWith('!'));
  // Globs here are simple (main.js, renderer/**/*, lib/**/*), so match by
  // literal segments around the wildcards rather than building a regex - the
  // escaping that needs is more of a hazard than the matching is worth.
  const covered = (rel) => globs.some(g => {
    const parts = g.split(/\*+\/?\*?/);
    if (parts.length === 1) return g === rel;
    if (!rel.startsWith(parts[0])) return false;
    let at = parts[0].length;
    for (const part of parts.slice(1)) {
      if (!part) continue;
      const found = rel.indexOf(part, at);
      if (found < 0) return false;
      at = found + part.length;
    }
    return true;
  });
  for (const entry of ['main.js', 'preload.js']) {
    const src = fs.readFileSync(path.join(appRoot, entry), 'utf-8');
    for (const m of src.matchAll(/require\(['"](\.\/[^'"]+)['"]\)/g)) {
      let rel = m[1].slice(2);
      if (!/\.[a-z]+$/i.test(rel)) rel += '.js';
      check(covered(rel), `${entry} requires ${m[1]}, and build.files packages it`,
        `add a glob covering "${rel}" to build.files - without it the packaged app cannot load`);
    }
  }
}

// ---- 9. nothing mkdirs under an unusable branch root ------------------------
// With no branch selected, path.join('', 'develop', ...) is RELATIVE and every
// resolveLogDir caller then mkdirSync({recursive})'d it inside the app's own
// folder - a stray develop/client/Skipbo/... with an empty extra_args.json,
// found only because the drift check refused to push past it.
{
  // Run the real function, lifted out of main.js, rather than trusting its text.
  const m = /function branchRootProblem\([^)]*\)\s*\{[\s\S]*?\n\}/.exec(main);
  check(!!m, 'main.js defines branchRootProblem');
  if (m) {
    const fn = new Function('path', 'fs', `${m[0]}; return branchRootProblem;`)(path, fs);
    const os = require('os');
    check(!!fn(''), 'an empty branch root is refused');
    check(!!fn('   '), 'a whitespace-only branch root is refused');
    check(!!fn('develop'), 'a relative branch root is refused - it would resolve inside the app');
    check(!!fn(path.join(os.tmpdir(), 'no-such-branch-' + Date.now())), 'a branch folder that does not exist is refused');
    check(fn(os.tmpdir()) === null, 'an existing absolute branch folder is accepted');
  }

  // And every handler that builds the log dir asks first. A new caller added
  // without the guard fails here instead of writing into the app folder.
  const lines = main.split('\n');
  lines.forEach((ln, i) => {
    if (!/=\s*resolveLogDir\(/.test(ln)) return;
    const before = lines.slice(Math.max(0, i - 4), i).join('\n');
    check(/branchRootProblem\(/.test(before),
      `main.js:${i + 1} checks branchRootProblem before resolveLogDir`,
      'call branchRootProblem(branchRoot) and return its message before building the log dir');
  });
}

console.log('');
if (failures) { console.error(`${failures} CONTRACT CHECK(S) FAILED`); process.exit(1); }
console.log('UI CONTRACT INTACT');
