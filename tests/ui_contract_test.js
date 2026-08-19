// UI CONTRACT TEST
// The renderer can be restyled and restructured freely, but app.js reaches into
// the DOM by ID and class. This test pins that contract so a redesign that drops
// a hook fails here instead of at runtime in front of a designer.
//
//   node tests/ui_contract_test.js
const fs = require('fs'), path = require('path');

const R = (f) => fs.readFileSync(path.join(__dirname, '..', 'renderer', f), 'utf-8');
const js = R('app.js'), html = R('index.html'), css = R('style.css');

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
  'trivia-challenge-view', 'endless-treasure-view',
  'profile-view', 'settings-view',
];
check(JSON.stringify(navOrder) === JSON.stringify(EXPECTED_NAV),
  `sidebar lists all 14 destinations in order (Sprint -> Ideation -> Main Events -> 9 regular -> Profile -> Settings)`,
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
const groupEnd = groupBlock.indexOf('</div>', groupBlock.indexOf('data-view="endless-treasure-view"'));
const inGroup = [...groupBlock.slice(0, groupEnd).matchAll(/data-view="([^"]+)"/g)].map(m => m[1]);
check(inGroup.length === 9 && inGroup[0] === 'minigames-view' && inGroup[8] === 'endless-treasure-view',
  'all 9 Regular Updates buttons nest inside the group container',
  'found ' + inGroup.length + ': ' + inGroup.join(', '));

// ---- 5. classes app.js emits must have styling ------------------------------
// Behaviour-only hooks are queried by JS but intentionally carry no rule.
const BEHAVIOUR_ONLY = new Set([
  'sp-go', 'ideation-select-btn', 'ideation-regen-btn',
  'artreq-select-btn', 'artreq-regen-btn', 'lesson-retract-btn',
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

// ---- 7. stylesheet is actually linked, and cache-busted ---------------------
const link = /href="style\.css(\?v=(\d+))?"/.exec(html);
check(!!link, 'index.html links style.css');
check(!!(link && link[2]), 'stylesheet link carries a ?v= cache-buster',
  'Electron caches aggressively; bump ?v= when style.css changes');

console.log('');
if (failures) { console.error(`${failures} CONTRACT CHECK(S) FAILED`); process.exit(1); }
console.log('UI CONTRACT INTACT');
