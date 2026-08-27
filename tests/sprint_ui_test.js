// Drives the REAL renderer/app.js sprint board against the REAL backend
// output (captured from python), so the UI is verified against production
// data shapes rather than hand-written fixtures.
const fs = require('fs'), path = require('path'), vm = require('vm'), cp = require('child_process');

function assert(c, m) { if (!c) { console.error('FAIL: ' + m); process.exitCode = 1; throw new Error(m); } }

// --- get genuine backend output ---
const PY = path.resolve(__dirname, '..', 'pybackend_dist', 'python');
const payload = path.join(PY, '_uitest_payload.json');
fs.writeFileSync(payload, JSON.stringify({
  roadmap_path: path.resolve(__dirname, '..', 'ideationData', 'Roadmap', 'Roadmap.xlsx'),
  design_path: String.raw`C:\SkipBo\branches\Misc\design\DesignData\WorkSpcae_Design\design`,
  year_hint: 2026
}));
const out = cp.execFileSync(path.join(PY, 'python.exe'),
  ['pybackend_cli.py', 'sprint_board', '_uitest_payload.json'],
  { cwd: PY, encoding: 'utf-8', maxBuffer: 8 * 1024 * 1024 });
fs.unlinkSync(payload);
const REAL = JSON.parse(out);
assert(!REAL.error, 'backend returned an error: ' + REAL.error);
console.log(`OK  backend produced real data: ${REAL.sprints.length} sprints, current="${REAL.current_sprint}"`);

// --- DOM stub ---
const els = new Map();
function mk(sel) {
  const el = {
    _sel: sel, value: '', textContent: '', innerHTML: '', style: {}, dataset: {}, _handlers: {}, _attrs: {},
    classList: { _s: new Set(), add(c){this._s.add(c);}, remove(c){this._s.delete(c);},
      contains(c){return this._s.has(c);}, toggle(c){ if(this._s.has(c)){this._s.delete(c);return false;} this._s.add(c); return true; } },
    addEventListener(e,f){ (this._handlers[e]=this._handlers[e]||[]).push(f); },
    fire(e,a){ return Promise.all((this._handlers[e]||[]).map(f=>f(a||{target:this}))); },
    click(){ return this.fire('click'); }, setAttribute(k,v){this._attrs[k]=v;}, getAttribute(k){return this._attrs[k];},
    querySelector(){return null;}, querySelectorAll(){return [];},
    scrollTop: 0, scrollHeight: 0, children: { length: 0 }, firstChild: null,
    appendChild(){}, remove(){}, focus(){}, closest(){return null;}, getContext:()=>({drawImage(){}}), toDataURL:()=>'data:,'
  };
  els.set(sel, el); return el;
}
const get = (s) => els.get(s) || mk(s);
const navButtons = ['sprint-view','ideation-view','main-events-view','minigames-view','battlepass-view',
  'holiday-bp-view','decoration-shop-view','streak-challenge-view','weekly-gift-view','full-screen-gift-view',
  'trivia-challenge-view','endless-treasure-view','profile-view','settings-view'].map(v => {
    const b = mk(`btn:${v}`); b.dataset.view = v; return b; });
const views = {}; navButtons.forEach(b => views[`#${b.dataset.view}`] = mk(`#${b.dataset.view}`));

const document = {
  querySelector: (s) => {
    const m = /^\.menu-btn\[data-view="([^"]+)"\]$/.exec(s);
    if (m) return navButtons.find(b => b.dataset.view === m[1]) || null;
    if (s === '#nav-group-regular' || s === '#nav-group-regular-items') { const e = get(s); e.querySelectorAll = () => []; e.querySelector = () => (s==='#nav-group-regular'? get('.nav-group-header'):null); return e; }
    return get(s);
  },
  querySelectorAll: (s) => {
    if (s === '.menu-btn') return navButtons;
    if (s === '.view') return Object.values(views);
    if (s === '.sp-go') return document.__spgo || [];
    return [];
  },
  __spgo: [], createElement: () => mk('created'), addEventListener() {}
};

const api = new Proxy({
  onLogLine: () => {},
  getConfig: async () => ({ cfg: { branchPaths:{trunk:'C:\\SkipBo\\branches\\Misc'}, activeBranch:'trunk',
      activeWorkspace:'WorkSpcae_Design', anthropicApiKey:'', ollamaBaseUrl:'', ollamaModel:'q',
      unityExePath:'', robotCalcPath:'', ideationDataPath:'', authorId:'AzureAD\\N', authorName:'N',
      avatarColor:'', avatarImage:'' },
    branchKeys:['trunk'], workspaces:['WorkSpcae_Design'], defaults:{trunk:''}, pipelineSteps:[] }),
  ideationDataFolderInfo: async () => ({ root:'', isDefault:true, exists:false, templatesFound:false }),
  resolveDesignPath: async () => String.raw`C:\SkipBo\branches\Misc\design\DesignData\WorkSpcae_Design\design`,
  sprintBoard: async ({ selectedSprint }) => {
    if (!selectedSprint) return { success: true, data: REAL };
    const p2 = path.join(PY, '_uitest_p2.json');
    fs.writeFileSync(p2, JSON.stringify({
      roadmap_path: path.resolve(__dirname, '..', 'ideationData', 'Roadmap', 'Roadmap.xlsx'),
      design_path: String.raw`C:\SkipBo\branches\Misc\design\DesignData\WorkSpcae_Design\design`,
      year_hint: 2026, selected_sprint: selectedSprint }));
    const o = cp.execFileSync(path.join(PY, 'python.exe'), ['pybackend_cli.py','sprint_board','_uitest_p2.json'],
      { cwd: PY, encoding: 'utf-8', maxBuffer: 8*1024*1024 });
    fs.unlinkSync(p2);
    return { success: true, data: JSON.parse(o) };
  },
}, { get: (t,k) => (k in t ? t[k] : async () => ({ success:true, stdout:'{}', data:{} })) });

const sandbox = { document, window:{api}, console, setTimeout, clearTimeout, Date, JSON, Math, Number,
  Array, Object, String, Boolean, isNaN, parseInt, parseFloat, FileReader:function(){}, Image:function(){} };
sandbox.window.document = document;
vm.createContext(sandbox);
vm.runInContext(fs.readFileSync(path.join(__dirname,'..','renderer','app.js'),'utf-8'), sandbox, {filename:'app.js'});

(async () => {
  await new Promise(r => setTimeout(r, 400));

  // 1. Lands on the sprint board with it marked active
  assert(get('#sprint-view').style.display === 'flex', 'sprint view is not the landing view');
  const sprintBtn = navButtons.find(b => b.dataset.view === 'sprint-view');
  assert(sprintBtn.classList.contains('active-btn'), 'Sprint Board not highlighted in the sidebar on startup');
  console.log('OK  app lands on the Sprint Board with the sidebar entry highlighted');

  // 2. Picker built from the roadmap, defaulting to today's sprint
  const sel = get('#sprint-select');
  assert((sel.innerHTML.match(/<option/g)||[]).length === REAL.sprints.length,
    'picker option count != roadmap sprint count');
  assert(sel.innerHTML.includes('← current'), 'current sprint not marked in the picker');
  assert(sel.innerHTML.includes('store submission'), 'Store Submission milestones not labelled');
  console.log(`OK  picker lists all ${REAL.sprints.length} columns, marks the current one and the milestones`);

  // 3. Detail renders a known sprint with its config match.
  //    Pinned explicitly rather than relying on whichever sprint is "current":
  //    the roadmap's sprints are 14-day windows, so a test that asserts on
  //    today's sprint passes for a fortnight and then fails on a calendar
  //    rollover with nothing having changed in the code.
  sel.value = 'Aug Event 1';
  await sel.fire('change', { target: sel });
  await new Promise(r => setTimeout(r, 400));
  const detail = get('#sprint-detail').innerHTML;
  assert(detail.includes('Five-Star Picnic'), 'primary event title missing from detail');
  assert(detail.includes('Dessert2608'), 'joined config alias not shown');
  assert(detail.includes('Alice Themed Event'), 'overlay campaign not surfaced');
  assert(detail.includes('All configured'), `expected Aug Event 1 fully configured, got: ${detail.match(/sp-count-pill[^>]*>([^<]*)/)}`);
  console.log('OK  detail shows event, joined config alias, overlay, and an all-clear pill');

  // 4a. "Chapter 82 + Star Chest" is now an actionable tab, not a read-only
  //     line: it renders as a Chapter Update card wired to the chapter view,
  //     and is marked configured because 章节82 exists in pve_chapter.xlsx.
  assert(detail.includes('Chapter Update'), 'Chapter Update card missing from the board');
  assert(/data-view="chapter-view"/.test(detail), 'Chapter Update card has no Open button to the tab');
  assert(detail.includes('82'), 'the chapter number the roadmap asked for is not shown');
  console.log('OK  "Chapter 82 + Star Chest" is an actionable Chapter Update card');

  // 4b. Lines this tool genuinely cannot action are still surfaced read-only,
  //     so claiming the chapter line did not swallow the rest of the scope.
  assert(detail.includes('Event Pack theme refresh'), 'roadmap lines without a tool tab are not shown');
  console.log("OK  roadmap lines this tool can't action (Event Pack theme …) still listed");

  // 5. Switching to the drifted sprint surfaces the drift loudly
  sel.value = 'Sep Event 1';
  await sel.fire('change', { target: sel });
  await new Promise(r => setTimeout(r, 400));
  const d2 = get('#sprint-detail').innerHTML;
  assert(d2.includes('sp-drift'), 'unconfigured sprint did not raise a drift banner');
  assert(/No event configured/.test(d2), 'missing-event drift text absent');
  assert(/to configure<\/span>/.test(d2), 'TODO count pill missing');
  assert(d2.includes('sp-todo'), 'no tab flagged as needing configuration');
  console.log('OK  "Sep Event 1" (planned, never configured) raises a drift banner + TODO cards');

  // 6. Date-mismatch drift, not just missing-event
  sel.value = 'Sep Event 2';
  await sel.fire('change', { target: sel });
  await new Promise(r => setTimeout(r, 400));
  const d3 = get('#sprint-detail').innerHTML;
  assert(d3.includes('Sim2609'), 'matched config alias missing on the mismatch case');
  assert(/Dates disagree/.test(d3), 'date-mismatch drift not explained');
  console.log('OK  "Sep Event 2" flags the 09-28 vs 09-30 date disagreement');

  // 7. Escaping holds for sheet-authored text
  assert(!/<script>/i.test(d3), 'roadmap text rendered unescaped');
  console.log('OK  roadmap-authored text is HTML-escaped');

  // 8. Clicking a nav button must LEAVE it highlighted. switchView() clears
  //    every active-btn, so a handler that marks before calling it ends up
  //    stripping its own highlight and the sidebar goes blank on every click.
  for (const target of ['settings-view', 'minigames-view', 'sprint-view']) {
    const btn = navButtons.find(b => b.dataset.view === target);
    await btn.fire('click');
    await new Promise(r => setTimeout(r, 60));
    const lit = navButtons.filter(b => b.classList.contains('active-btn')).map(b => b.dataset.view);
    assert(lit.length === 1 && lit[0] === target,
      `after clicking ${target} the sidebar should highlight exactly it, got [${lit}]`);
    assert(get(`#${target}`).style.display === 'flex', `${target} pane did not open on click`);
  }
  console.log('OK  clicking a destination leaves exactly that one highlighted');

  console.log('\nALL SPRINT UI TESTS PASSED');
})();
