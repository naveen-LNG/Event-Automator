const { app, BrowserWindow, ipcMain, dialog, shell } = require('electron');
const path = require('path');
const fs = require('fs');
const os = require('os');
const { spawn, execFile } = require('child_process');

// ---- Static config ----

// Backend runs on the official python.org embeddable interpreter (PSF-signed
// python.exe) bundled alongside our scripts + openpyxl, rather than a custom
// PyInstaller-compiled exe — unsigned PyInstaller bootloaders routinely trip
// AV/EDR heuristics (SentinelOne, Defender) as generic-packer false positives.
// No system Python required on the end user's machine either way.
// Dev mode: built locally into pybackend_dist/python/ (npm run build:py).
// Packaged mode: electron-builder ships it under resources/pybackend/ (extraResources).
const PYBACKEND_DIR = app.isPackaged
  ? path.join(process.resourcesPath, 'pybackend')
  : path.join(__dirname, 'pybackend_dist', 'python');
const PYTHON_EXE = path.join(PYBACKEND_DIR, 'python.exe');
const PYBACKEND_SCRIPT = path.join(PYBACKEND_DIR, 'pybackend_cli.py');
const PYBACKEND_CWD = PYBACKEND_DIR;

const BRANCH_KEYS = ['trunk', 'Misc', 'Feature'];
const WORKSPACES = ['WorkSpcae_Design', 'WorkSpcae_NE', 'WorkSpcae_FC', 'WorkSpcae3'];

// Ideation reference data lives alongside the app, not per-branch — it's
// cross-branch design knowledge (past events' art requirements, feedback,
// one-pagers), same as the existing PDFs already checked in here. Shipped
// via extraResources (like pybackend) rather than the asar-packed `files`
// list, since app.asar is a read-only virtual archive — the app writes
// feedback/final-doc files here at runtime, which asar can't support.
const DEFAULT_IDEATION_DATA_ROOT = app.isPackaged
  ? path.join(process.resourcesPath, 'ideationData')
  : path.join(__dirname, 'ideationData');

// Point this at a synced folder (OneDrive for Business / a network share) to
// share ideation work across the team — every designer's app writes into the
// same tree, and the file-per-entry storage in ideation_store.py is what
// keeps concurrent writes from colliding there. Empty = the local default
// above, which is also the correct setting for working solo.
function ideationDataRoot() {
  const configured = (loadConfig().ideationDataPath || '').trim();
  return configured || DEFAULT_IDEATION_DATA_ROOT;
}

function ideationEventsRoot() {
  return path.join(ideationDataRoot(), 'Events');
}

// Identity is attribution, not authentication: it answers "who chose this
// variation / who left this note", and it cannot be enforced by a local app
// with no server. Rather than a login screen nobody can verify anyway, take
// the account Windows already authenticated — these are AzureAD-joined
// machines, so USERDOMAIN\USERNAME is a real corporate identity. The display
// name stays editable in Settings; the id is what's stamped on records.
function detectIdentity() {
  let username = '';
  try {
    username = os.userInfo().username || '';
  } catch (e) {
    username = process.env.USERNAME || '';
  }
  const domain = process.env.USERDOMAIN || '';
  return {
    authorId: domain ? `${domain}\\${username}` : (username || 'unknown'),
    authorName: username || 'Unknown User'
  };
}

const PIPELINE_STEPS = [
  '1. Configure storagedata.proto',
  '2. Convert proto files to Lua (ProtoToLua)',
  '3. Add active save file in local service',
  '4. Register in LocalServiceMgr.lua',
  '5. Clone WorkSpace3 configuration to active workspace',
  '6. Clone Event Excel configuration tables',
  '7. Copy script to app_client directory',
  '8. Modify convert_layout.lua & convert.json',
  '9. Update bi.xlsx (BI tracking)',
  '10. Update events.xlsx (Event registration)',
  '11. Update event_shop.xlsx (Shop config)',
  '12. Update item.xlsx (Item definitions)',
  '13. Update icon.xlsx (Icon references)',
  '14. Update localization & answer_challenge (stage rewards)',
  '15. Update asset_ref.xlsx (Asset references)',
  '16. Update store.xlsx (Store config)',
  '17. Update card.xlsx (Card config)',
  '18. Update pack.xlsx (Pack config)',
  '19. Update guide.xlsx (Guide config)',
  '20. Update sys.xlsx (DeepLink registration)',
  '21. Execute coder_convert script',
  '22. Register event in EventConfig.lua, BIDef.lua, UIManager.lua, PanelConfig.lua & GuideRefuiMgr.lua',
  '23. Copy & rename event code folders (Core<Event> + UI<Event>Panel)',
  '24. Copy & rename ResourcesAb event data (ResourcesAb/Events/<Event>)',
  '25. Copy & rename event-ID resources (TabEventBtn, CommonEventImg, PropIcon, EventAchieve)'
];

// Guessed repo root: this app lives at <repoRoot>/tools/EventAutomationDesktop
const GUESSED_REPO_ROOT = path.resolve(__dirname, '..', '..');
const DEFAULT_BRANCH_SUBPATH = {
  trunk: 'trunk',
  Misc: path.join('branches', 'Misc'),
  Feature: path.join('branches', 'Feature')
};

function defaultBranchPath(branchKey) {
  return path.join(GUESSED_REPO_ROOT, DEFAULT_BRANCH_SUBPATH[branchKey]);
}

// ---- Config persistence (per-user, in userData) ----

function configPath() {
  return path.join(app.getPath('userData'), 'config.json');
}

function loadConfig() {
  const p = configPath();
  let cfg = {
    branchPaths: {},
    activeBranch: 'trunk',
    activeWorkspace: 'WorkSpcae_Design',
    anthropicApiKey: '',
    ollamaBaseUrl: '',
    ollamaModel: 'qwen3:32b',
    robotCalcPath: '',
    // Blank = derive from the active branch (<branch>/develop/tools/RunCardEx2).
    // Only set when the simulation output lives somewhere else on this machine.
    chapterSimPath: '',
    // tools/ sits OUTSIDE the branch checkout and its position relative to a
    // branch differs (trunk/ is one level down, branches/Misc/ two), so these
    // cannot be derived from the branch path the way design/ and develop/ are.
    // Blank = guess <repo>/tools/...; Browse when the checkout is laid out
    // differently on this machine.
    chapterLvlPyPath: '',
    chapterMapBuilderPath: '',
    // Blank = the local default folder next to the app. Point at a synced
    // folder to share ideation work with the team.
    ideationDataPath: '',
    comfyUrl: 'https://skipbo.comfyui.liquidnitro.games',
    // Auto-filled from the OS on first launch (see detectIdentity).
    authorId: '',
    authorName: '',
    // Avatar is mirrored into the shared folder by profile:save; the local
    // copy is what renders instantly at startup before that read completes.
    avatarColor: '',
    avatarImage: ''
  };
  if (fs.existsSync(p)) {
    try {
      const onDisk = JSON.parse(fs.readFileSync(p, 'utf-8'));
      cfg = Object.assign(cfg, onDisk);
      cfg.branchPaths = Object.assign({}, onDisk.branchPaths || {});
    } catch (e) {
      // Corrupt config; fall back to defaults rather than crash the app.
    }
  }

  // Auto-fill a branch path only if unset AND the guessed default actually exists on
  // this machine — otherwise leave it blank so the UI shows the default as a placeholder
  // hint and the user must Browse to their real checkout.
  for (const key of BRANCH_KEYS) {
    if (!cfg.branchPaths[key]) {
      const guess = defaultBranchPath(key);
      if (fs.existsSync(guess)) {
        cfg.branchPaths[key] = guess;
      }
    }
  }

  // First launch: adopt the Windows-authenticated account rather than
  // prompting for a login the app couldn't verify anyway. Only fills blanks,
  // so an edited display name is never overwritten.
  if (!cfg.authorId || !cfg.authorName) {
    const detected = detectIdentity();
    cfg.authorId = cfg.authorId || detected.authorId;
    cfg.authorName = cfg.authorName || detected.authorName;
  }
  return cfg;
}

function saveConfig(partial) {
  const current = loadConfig();
  const merged = Object.assign({}, current, partial);
  if (partial.branchPaths) {
    merged.branchPaths = Object.assign({}, current.branchPaths, partial.branchPaths);
  }
  fs.writeFileSync(configPath(), JSON.stringify(merged, null, 2), 'utf-8');
  return merged;
}

// ---- Path resolution helpers (mirrors EventAutomationWindow.cs logic) ----

function resolveDesignPath(branchRoot, workspace) {
  return path.join(branchRoot, 'design', 'DesignData', workspace, 'design');
}

function resolveUnityProjectPath(branchRoot) {
  return path.join(branchRoot, 'develop', 'client', 'Skipbo');
}

// Direct port of ProtoTool.ProtoToLua (Assets/Editor/ABBuilder/ProtoTool.cs): each
// develop/protocol/<name>.proto becomes PB/<name>.lua containing `return [[<proto text>]]`.
// That's the whole conversion — no reason to boot a 22GB Unity project headlessly for it
// (which took >10 min and refused to start whenever the Editor had the project open).
// Verified byte-identical to the Unity-generated files for every untouched proto in
// trunk/Misc/Feature. Differences from the C# original, both deliberate:
//  - only *.proto files: the C# reads EVERY file in the folder, so step 1's
//    storagedata.proto.bak would have produced a junk PB/storagedata.proto.lua asset;
//  - no TortoiseProc SVN update: it pops a GUI dialog, and step 1 has just edited
//    storagedata.proto locally, so updating here risks a conflict mid-pipeline.
// useBackups: read <name>.proto.bak when present (step 1's pre-edit copy) — used by Undo.
function protoToLua(branchRoot, useBackups) {
  const protoDir = path.join(branchRoot, 'develop', 'protocol');
  const pbDir = path.join(resolveUnityProjectPath(branchRoot), 'Assets', 'Lua', 'Game', 'Module', 'Protocol', 'PB');
  if (!fs.existsSync(protoDir)) return { success: false, output: `Protocol folder not found: ${protoDir}` };
  if (!fs.existsSync(pbDir)) return { success: false, output: `PB output folder not found: ${pbDir}` };

  // Build every output first so a bad file aborts before anything is written.
  const outputs = [];
  for (const file of fs.readdirSync(protoDir).filter(f => f.endsWith('.proto'))) {
    const src = path.join(protoDir, file);
    const bak = src + '.bak';
    const readFrom = useBackups && fs.existsSync(bak) ? bak : src;
    let text = fs.readFileSync(readFrom, 'utf8');
    if (text.charCodeAt(0) === 0xFEFF) text = text.slice(1); // C# File.ReadAllText drops the BOM
    // Lua long string `[[...]]` ends at the first `]]`; one inside the proto would silently
    // truncate the embedded schema. The C# original has the same hole — fail loudly instead.
    if (text.includes(']]')) return { success: false, output: `${file} contains "]]", which would break the Lua long string. Nothing written.` };
    outputs.push({
      readFrom,
      out: path.join(pbDir, file.replace(/\.proto$/, '.lua')),
      content: Buffer.from('return [[' + text + ']]', 'utf8') // UTF-8, no BOM, line endings kept
    });
  }

  const lines = [];
  let changed = 0;
  for (const { readFrom, out, content } of outputs) {
    if (fs.existsSync(out) && fs.readFileSync(out).equals(content)) {
      lines.push(`  unchanged  ${path.basename(out)}`);
      continue;
    }
    fs.writeFileSync(out, content);
    changed++;
    lines.push(`  written    ${path.basename(out)}  (from ${path.basename(readFrom)}, ${content.length} bytes)`);
  }
  lines.unshift(`ProtoToLua: ${changed} file(s) written to ${pbDir}`);
  return { success: true, output: lines.join('\n') };
}

function resolveLogDir(branchRoot) {
  // Must match the path event_pipeline.py itself derives from proj_root,
  // so undo_log.json / extra_args.json land in the same place the script reads them from.
  return path.join(branchRoot, 'develop', 'client', 'Skipbo', 'Assets', 'Editor', 'EventAutomation', 'PythonBackend');
}

// Every caller of resolveLogDir mkdirs the result, so a missing branch is not a
// harmless "not found". path.join('', 'develop', ...) is RELATIVE, it resolves
// against the app's own folder, and mkdirSync({recursive}) quietly creates
// develop/client/Skipbo/... inside EventAutomationDesktop - which is exactly how
// a stray extra_args.json full of empty values turned up there and blocked a
// push on the drift check. A branch folder that does not exist (an unplugged
// drive, a stale path in config) would do the same on whatever drive it names.
// Returns the reason to refuse, or null when the branch root is usable.
function branchRootProblem(branchRoot) {
  if (!branchRoot || !String(branchRoot).trim()) return 'No branch is selected - pick one in the top bar first.';
  if (!path.isAbsolute(branchRoot)) return `Branch path is not absolute, refusing to write relative to the app: ${branchRoot}`;
  if (!fs.existsSync(branchRoot)) return `Branch folder not found: ${branchRoot}`;
  return null;
}

// ---- Window ----

function createWindow() {
  const win = new BrowserWindow({
    width: 1300,
    height: 860,
    minWidth: 480,
    minHeight: 600,
    title: `Event Automator v${app.getVersion()}`,
    icon: path.join(__dirname, 'build', 'icon.png'),
    // Painted before the renderer's first frame — must match --void in
    // style.css or launch flashes the old blue theme for a beat.
    backgroundColor: '#04050c',
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false
    }
  });
  // Keep the version in the title; the page's <title> would replace it.
  win.on('page-title-updated', (e) => e.preventDefault());
  win.loadFile(path.join(__dirname, 'renderer', 'index.html'));
  return win;
}

// Update check against GitHub releases.
const RELEASES_URL = 'https://github.com/naveen-LNG/Event-Automator/releases/latest';
const BASE_TITLE = () => `Event Automator v${app.getVersion()}`;

// Progress goes to the in-app bar, the taskbar icon and the window title; null clears it.
function showUpdateProgress(win, p) {
  if (!win || win.isDestroyed()) return;
  win.webContents.send('update:progress', p);
  win.setProgressBar(p ? (p.percent >= 0 ? p.percent / 100 : 2) : -1);  // 2 = indeterminate
  win.setTitle(p && p.percent >= 0 ? `${BASE_TITLE()} - downloading v${p.version}: ${Math.floor(p.percent)}%` : BASE_TITLE());
}

function checkForUpdates(win) {
  if (!app.isPackaged) return;
  const { autoUpdater } = require('electron-updater');
  autoUpdater.autoDownload = false;
  let downloading = null;  // version being downloaded
  autoUpdater.on('error', (e) => {
    console.warn('Update check failed:', e && e.message);
    if (!downloading) return;  // a failed background check stays silent
    const version = downloading;
    downloading = null;
    showUpdateProgress(win, null);
    dialog.showMessageBox(win, {
      type: 'error', title: 'Update failed',
      message: `Downloading v${version} failed.`,
      detail: `${(e && e.message) || e}\n\nCheck your connection and restart the tool to try again, or download the installer from ${RELEASES_URL}`
    });
  });
  autoUpdater.on('download-progress', (p) => {
    if (downloading) showUpdateProgress(win, { version: downloading, percent: p.percent, transferred: p.transferred, total: p.total, speed: p.bytesPerSecond });
  });
  autoUpdater.on('update-available', async (info) => {
    const { response } = await dialog.showMessageBox(win, {
      type: 'info',
      title: 'Update available',
      message: `Event Automator v${info.version} is available (you have v${app.getVersion()}).`,
      detail: 'Download it now? It downloads in the background (about 1 GB, progress shows at the top of the window) and installs when you restart the tool.',
      buttons: ['Download', 'Later'],
      defaultId: 0, cancelId: 1
    });
    if (response !== 0) return;
    downloading = info.version;
    showUpdateProgress(win, { version: info.version, percent: 0, transferred: 0, total: 0, speed: 0 });
    autoUpdater.downloadUpdate().catch(() => {});  // failures arrive via 'error'
  });
  autoUpdater.on('update-downloaded', async (info) => {
    downloading = null;
    showUpdateProgress(win, null);
    const { response } = await dialog.showMessageBox(win, {
      type: 'info',
      title: 'Update ready',
      message: `v${info.version} is downloaded.`,
      detail: 'Restart now to install it, or it installs the next time you close the tool.',
      buttons: ['Restart now', 'Later'],
      defaultId: 0, cancelId: 1
    });
    if (response === 0) autoUpdater.quitAndInstall();
  });
  autoUpdater.checkForUpdates().catch(() => {});
}

// Claude Desktop needs the Event Automator extension (.mcpb) or an mcpServers entry
// for Localization > Translate and the Art Requirement tools.
function claudeExtensionConfigured() {
  const root = path.join(process.env.APPDATA || '', 'Claude');
  try {
    const ext = fs.readdirSync(path.join(root, 'Claude Extensions')).find((d) => d.endsWith('.event-automator'));
    if (ext) {
      let enabled = true;
      try {
        enabled = JSON.parse(fs.readFileSync(path.join(root, 'Claude Extensions Settings', ext + '.json'), 'utf-8')).isEnabled !== false;
      } catch (e) { /* no settings file = default enabled */ }
      if (enabled) return true;
    }
  } catch (e) { /* no Claude Desktop extensions folder */ }
  try {
    const servers = JSON.parse(fs.readFileSync(path.join(root, 'claude_desktop_config.json'), 'utf-8')).mcpServers || {};
    return Object.entries(servers).some(([k, v]) =>
      /event.?automator/i.test(k) || /EventAutomationDesktop/i.test(JSON.stringify(v)));
  } catch (e) { return false; }
}

async function warnIfClaudeExtensionMissing(win) {
  if (claudeExtensionConfigured() || loadConfig().hideMcpbWarning) return;
  const { response } = await dialog.showMessageBox(win, {
    type: 'warning',
    title: 'Claude Desktop extension not set up',
    message: 'The Event Automator extension (.mcpb) is not installed in Claude Desktop.',
    detail: 'Localization > Translate and Art Requirement writing go through Claude Desktop and will not work without it. ' +
            'Every other tab works as normal.\n\n' +
            'To set it up: download event-automator.mcpb, double-click it, click Install in Claude Desktop, then restart Claude Desktop. ' +
            'It needs your SkipBo checkout at C:\\SkipBo (or set EVENT_AUTOMATOR_MCP to your mcp\\server.js).',
    buttons: ['Download .mcpb', 'Remind me later', "Don't show again"],
    defaultId: 0, cancelId: 1
  });
  if (response === 0) shell.openExternal(RELEASES_URL + '/download/event-automator.mcpb');
  if (response === 2) saveConfig({ hideMcpbWarning: true });
}

app.whenReady().then(() => {
  const win = createWindow();
  checkForUpdates(win);
  warnIfClaudeExtensionMissing(win);
  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});

// ---- IPC: config & dialogs ----

ipcMain.handle('config:get', () => {
  const cfg = loadConfig();
  return {
    cfg,
    branchKeys: BRANCH_KEYS,
    workspaces: WORKSPACES,
    defaults: Object.fromEntries(BRANCH_KEYS.map(k => [k, defaultBranchPath(k)])),
    pipelineSteps: PIPELINE_STEPS
  };
});

ipcMain.handle('config:save', (evt, partial) => saveConfig(partial));

ipcMain.handle('dialog:browseFolder', async (evt, startPath) => {
  const win = BrowserWindow.getFocusedWindow();
  const result = await dialog.showOpenDialog(win, {
    properties: ['openDirectory'],
    defaultPath: startPath && fs.existsSync(startPath) ? startPath : undefined
  });
  if (result.canceled || result.filePaths.length === 0) return null;
  return result.filePaths[0];
});

ipcMain.handle('dialog:browseFile', async (evt, { startPath, filters }) => {
  const win = BrowserWindow.getFocusedWindow();
  const result = await dialog.showOpenDialog(win, {
    properties: ['openFile'],
    defaultPath: startPath && fs.existsSync(startPath) ? startPath : undefined,
    filters: filters || [{ name: 'Executable', extensions: ['exe'] }]
  });
  if (result.canceled || result.filePaths.length === 0) return null;
  return result.filePaths[0];
});

ipcMain.handle('path:validate', (evt, targetPath) => {
  return !!(targetPath && fs.existsSync(targetPath));
});

ipcMain.handle('path:resolveDesign', (evt, { branchRoot, workspace }) => {
  return resolveDesignPath(branchRoot, workspace);
});

// ---- IPC: pipeline execution ----

function runPython(scriptId, args) {
  return new Promise((resolve) => {
    execFile(PYTHON_EXE, [PYBACKEND_SCRIPT, scriptId, ...args], { cwd: PYBACKEND_CWD, maxBuffer: 20 * 1024 * 1024 }, (error, stdout, stderr) => {
      resolve({
        success: !error,
        code: error ? (error.code || 1) : 0,
        stdout: stdout || '',
        stderr: stderr || '',
        output: (stdout || '') + (stderr ? '\n' + stderr : '')
      });
    });
  });
}

ipcMain.handle('pipeline:runStep', async (evt, params) => {
  const {
    stepIndex, srcEvt, tgtEvt, isRevert,
    branchRoot, workspace,
    eventId, startTime, endTime, nearEndTime, closeTime, isReopen, srcStartTime
  } = params;

  const problem = branchRootProblem(branchRoot);
  if (problem) return { success: false, code: 1, output: problem };
  const logDir = resolveLogDir(branchRoot);
  fs.mkdirSync(logDir, { recursive: true });

  // Write extra_args.json exactly where event_pipeline.py expects to load it from.
  const extraArgs = {
    event_id: eventId || '',
    start_time: startTime || '',
    end_time: endTime || '',
    near_end_time: nearEndTime || '',
    close_time: closeTime || '',
    is_reopen: isReopen ? '1' : '0',
    // Source event's own start_time — lets the SVN-history lookup in step 7
    // (inject_convert_references) bracket its search by date instead of a slow
    // blind scan back from HEAD.
    src_start_time: srcStartTime || ''
  };
  fs.writeFileSync(path.join(logDir, 'extra_args.json'), JSON.stringify(extraArgs), 'utf-8');

  // Step index 1: ProtoToLua — done directly (see protoToLua above), no Unity launch.
  // Revert regenerates from step 1's storagedata.proto.bak while it still exists: Undo All
  // runs this before step 1's own revert, so the .bak is the true pre-pipeline proto.
  if (stepIndex === 1) {
    try {
      return protoToLua(branchRoot, isRevert);
    } catch (e) {
      return { success: false, output: `ProtoToLua failed: ${e.message}` };
    }
  }

  // Step index 20: coder_convert — a standalone batch file, no Unity involved.
  if (stepIndex === 20) {
    if (isRevert) {
      return { success: true, output: 'Reverting coder_convert is not supported directly. Skipping.' };
    }
    const batPath = path.join(branchRoot, 'design', 'DesignData', workspace, 'coder_convert_not_pause.bat');
    if (!fs.existsSync(batPath)) {
      return { success: false, output: `Error: coder_convert_not_pause.bat not found at ${batPath}` };
    }
    return new Promise((resolve) => {
      let settled = false;

      // Run in a real, visible console window instead of a detached/piped child.
      // Two direct benefits: (1) full native output — colors, formatting, everything —
      // exactly as coder_convert normally produces it, easiest possible diagnosis when
      // something fails. (2) convert_to_lua.bat's own `@pause` on failure (see the step-19
      // review) stops being a silent hang and becomes a real interactive prompt: the window
      // stays open showing the error, the user reads it and presses a key, done. The same
      // @pause we flagged as a liability when hidden is exactly the right behavior once visible.
      const title = `Coder Convert - ${workspace} (${srcEvt} -> ${tgtEvt})`;

      // convert_to_lua.bat shells out to bare `python` for a suffix-handling script — not
      // our bundled interpreter. Prepend our bundled Python's directory to PATH for this
      // window so that resolves too, keeping "no system Python required" true here as well.
      const childEnv = Object.assign({}, process.env, {
        PATH: `${PYBACKEND_DIR}${path.delimiter}${process.env.PATH || ''}`
      });

      // `start /wait`'s OWN exit code does not reliably reflect the launched command's exit
      // code (verified empirically — it came back 0 even when the inner script exited 1).
      // Route the real exit code through a temp file instead, written from inside the same
      // window. Two more gotchas found the same way: the .bat must be invoked with `call`
      // (without it, control never returns to the rest of the command line), and the file
      // write needs delayed expansion (`/v:on` + `!errorlevel!`) — `%errorlevel%` gets
      // substituted at parse time for the whole line, before the batch file has even run,
      // so it always reads as whatever errorlevel was before this command started.
      const exitCodeFile = path.join(app.getPath('temp'), `coder_convert_exitcode_${Date.now()}.txt`);
      const command = `start "${title}" /wait cmd.exe /v:on /c "call "${batPath}" & echo !errorlevel! > "${exitCodeFile}""`;

      evt.sender.send('pipeline:log-line',
        `[CoderConvert] Opened in a separate terminal window ("${title}") — watch that window for live output. ` +
        `This step completes automatically once it finishes (or once you dismiss a failure prompt there).`);

      const proc = spawn(command, { cwd: path.dirname(batPath), shell: true, windowsHide: false, env: childEnv });

      // Safety net only — the visible window is the primary mechanism now, not a substitute
      // for one. Generous duration since a real @pause is now a legitimate state a person
      // might be reading, not an invisible freeze; this just guards against walking away.
      const COOKED_TIMEOUT_MS = 45 * 60 * 1000;
      const timer = setTimeout(() => {
        if (settled) return;
        settled = true;
        const msg = `[CoderConvert] No response after ${COOKED_TIMEOUT_MS / 60000} minutes — closing the terminal window.`;
        evt.sender.send('pipeline:log-line', msg);
        if (proc.pid) {
          spawn('taskkill', ['/PID', String(proc.pid), '/T', '/F']); // kill the whole tree, not just the outer cmd
        }
        resolve({ success: false, code: 1, output: msg });
      }, COOKED_TIMEOUT_MS);

      proc.on('close', () => {
        if (settled) return;
        settled = true;
        clearTimeout(timer);

        let realCode = null;
        try {
          if (fs.existsSync(exitCodeFile)) {
            realCode = parseInt(fs.readFileSync(exitCodeFile, 'utf-8').trim(), 10);
            fs.unlinkSync(exitCodeFile);
          }
        } catch (e) { /* fall through to the "couldn't determine" branch below */ }

        if (realCode === null || Number.isNaN(realCode)) {
          const msg = `[CoderConvert] Terminal window closed, but couldn't read the real exit code — treating as failed so this doesn't silently pass.`;
          evt.sender.send('pipeline:log-line', msg);
          resolve({ success: false, code: 1, output: msg });
          return;
        }

        const msg = `[CoderConvert] Terminal window closed (exit code ${realCode}).`;
        evt.sender.send('pipeline:log-line', msg);
        resolve({ success: realCode === 0, code: realCode, output: msg });
      });
      proc.on('error', (err) => {
        if (settled) return;
        settled = true;
        clearTimeout(timer);
        resolve({ success: false, code: 1, output: `Failed to open coder_convert terminal: ${err.message}` });
      });
    });
  }

  // Every other step is a pure Python operation on the design/develop trees — no Unity needed.
  // Streamed (not buffered like runPython/execFile) because some steps — notably step 7's
  // SVN-history lookup for old source events — can legitimately take minutes; without live
  // output the UI would show nothing and look hung instead of just slow.
  return new Promise((resolve) => {
    let output = '';
    const proc = spawn(PYTHON_EXE, [
      PYBACKEND_SCRIPT, 'event_pipeline', String(stepIndex), srcEvt, tgtEvt, branchRoot, isRevert ? '1' : '0', workspace
    ], { cwd: PYBACKEND_CWD });

    proc.stdout.on('data', (d) => {
      output += d.toString();
      evt.sender.send('pipeline:log-line', d.toString());
    });
    proc.stderr.on('data', (d) => {
      output += d.toString();
      evt.sender.send('pipeline:log-line', d.toString());
    });
    proc.on('close', (code) => {
      resolve({ success: code === 0, code, output });
    });
    proc.on('error', (err) => {
      resolve({ success: false, code: 1, output: `Failed to launch Python: ${err.message}` });
    });
  });
});

// ---- IPC: Minigames / BattlePass / Holiday BP sub-tools (unchanged Python scripts) ----

function tempJsonFile(prefix) {
  return path.join(app.getPath('temp'), `${prefix}_${Date.now()}.json`);
}

ipcMain.handle('minigames:get', async (evt, { designPath }) => {
  return runPython('get_minigames_info', [designPath]);
});

ipcMain.handle('minigames:commit', async (evt, { designPath, payload }) => {
  const tmp = tempJsonFile('minigames_payload');
  fs.writeFileSync(tmp, JSON.stringify(payload), 'utf-8');
  return runPython('update_minigames', [designPath, tmp]);
});

ipcMain.handle('bp:get', async (evt, { designPath }) => {
  return runPython('get_bp_info', [designPath]);
});

ipcMain.handle('bp:commit', async (evt, { designPath, payload }) => {
  const tmp = tempJsonFile('bp_payload');
  fs.writeFileSync(tmp, JSON.stringify(payload), 'utf-8');
  return runPython('update_bp', [designPath, tmp]);
});

ipcMain.handle('hbp:get', async (evt, { designPath }) => {
  return runPython('get_hbp_info', [designPath]);
});

ipcMain.handle('hbp:commit', async (evt, { designPath, payload }) => {
  const tmp = tempJsonFile('hbp_payload');
  fs.writeFileSync(tmp, JSON.stringify(payload), 'utf-8');
  return runPython('update_hbp', [designPath, tmp]);
});

ipcMain.handle('decoshop:commit', async (evt, { designPath, payload }) => {
  const tmp = tempJsonFile('decoshop_payload');
  fs.writeFileSync(tmp, JSON.stringify(payload), 'utf-8');
  return runPython('update_decoration_shop', [designPath, tmp]);
});

ipcMain.handle('robotpvp:commit', async (evt, { designPath, payload }) => {
  const tmp = tempJsonFile('robotpvp_payload');
  fs.writeFileSync(tmp, JSON.stringify(payload), 'utf-8');
  return runPython('update_robot_pvp', [designPath, tmp]);
});

ipcMain.handle('residenttask:commit', async (evt, { designPath, payload }) => {
  const tmp = tempJsonFile('residenttask_payload');
  fs.writeFileSync(tmp, JSON.stringify(payload), 'utf-8');
  return runPython('update_resident_task', [designPath, tmp]);
});

ipcMain.handle('weeklygift:commit', async (evt, { designPath, payload }) => {
  const tmp = tempJsonFile('weeklygift_payload');
  fs.writeFileSync(tmp, JSON.stringify(payload), 'utf-8');
  return runPython('update_weekly_gift', [designPath, tmp]);
});

ipcMain.handle('fullscreengift:get', async (evt, { designPath }) => {
  return runPython('get_full_screen_gift_info', [designPath]);
});

ipcMain.handle('fullscreengift:commit', async (evt, { designPath, payload }) => {
  const tmp = tempJsonFile('fullscreengift_payload');
  fs.writeFileSync(tmp, JSON.stringify(payload), 'utf-8');
  return runPython('update_full_screen_gift', [designPath, tmp]);
});

ipcMain.handle('endlesstreasure:get', async (evt, { designPath }) => {
  return runPython('get_endless_treasure_info', [designPath]);
});

ipcMain.handle('endlesstreasure:commit', async (evt, { designPath, payload }) => {
  const tmp = tempJsonFile('endlesstreasure_payload');
  fs.writeFileSync(tmp, JSON.stringify(payload), 'utf-8');
  return runPython('update_endless_treasure', [designPath, tmp]);
});

// ---- IPC: Trivia Challenge (答题挑战) quiz cycle refresh ----

// Every configured branch's design dir (trunk/Misc/Feature/...), same workspace —
// so a genuinely new cycle picks the next batch nobody else has used yet, even
// though each branch is a separate checkout (possibly on a separate machine for a
// separate team). A branch with no path configured yet is just skipped.
// Every configured branch × workspace, labelled for the batch usage table. The next
// batch avoids anything used in any of them, not just this workspace.
function allQuizDesignDirs() {
  const cfg = loadConfig();
  const dirs = {};
  for (const key of BRANCH_KEYS) {
    const root = cfg.branchPaths[key];
    if (!root) continue;
    for (const ws of WORKSPACES) dirs[`${key} / ${ws}`] = resolveDesignPath(root, ws);
  }
  return dirs;
}

// questionsSource: '' = next unused batch, else a pool module like 'trivia_questions_batch_03'.
ipcMain.handle('trivia:get', async (evt, { branchRoot, workspace, questionsSource }) => {
  return runPython('get_trivia_challenge_info', [branchRoot, workspace, JSON.stringify(allQuizDesignDirs()), questionsSource || '']);
});

ipcMain.handle('trivia:run', async (evt, { branchRoot, workspace, questionsSource }) => {
  const problem = branchRootProblem(branchRoot);
  if (problem) return { success: false, code: 1, output: problem };
  const logDir = resolveLogDir(branchRoot);
  fs.mkdirSync(logDir, { recursive: true });

  return new Promise((resolve) => {
    let output = '';
    const proc = spawn(PYTHON_EXE, [
      PYBACKEND_SCRIPT, 'trivia_challenge_cycle', branchRoot, workspace, 'run', JSON.stringify(allQuizDesignDirs()), questionsSource || ''
    ], { cwd: PYBACKEND_CWD });

    proc.stdout.on('data', (d) => { output += d.toString(); evt.sender.send('pipeline:log-line', d.toString()); });
    proc.stderr.on('data', (d) => { output += d.toString(); evt.sender.send('pipeline:log-line', d.toString()); });
    proc.on('close', (code) => resolve({ success: code === 0, code, output }));
    proc.on('error', (err) => resolve({ success: false, code: 1, output: `Failed to launch Python: ${err.message}` }));
  });
});

// ---- IPC: Merge to Trunk (xlsx_merge) ----
// Row-level 3-way merge of a sprint workspace's xlsx changes into trunk/WorkSpcae3.
// The backend is the same xlsx_merge.py as tools/XlsxMerge; it streams progress
// and leaves report.md / conflicts.csv / remap_suggest.csv / summary.json in a
// per-run folder under the shared data dir (shared with the CLI so both see the
// same state.json, i.e. the same "last merged revision").
const XLSX_MERGE_DATA_DIR = path.join(process.env.LOCALAPPDATA || app.getPath('userData'), 'XlsxMerge');

function workspaceRoot(branchRoot, workspace) {
  return path.join(branchRoot, 'design', 'DesignData', workspace);
}

ipcMain.handle('xlsxmerge:defaults', () => {
  const cfg = loadConfig();
  const activeRoot = (cfg.branchPaths || {})[cfg.activeBranch] || '';
  const trunkRoot = (cfg.branchPaths || {}).trunk || defaultBranchPath('trunk');
  return {
    branchWs: activeRoot ? workspaceRoot(activeRoot, cfg.activeWorkspace) : '',
    trunkWs: workspaceRoot(trunkRoot, 'WorkSpcae3')
  };
});

// Minimal RFC4180 reader for the CSVs the backend writes (utf-8-sig, CRLF,
// quoted fields that may contain commas). Returns one object per row keyed
// by the header line.
function parseCsv(text) {
  text = (text || '').replace(/^﻿/, '');
  if (!text.trim()) return [];
  const rows = []; let row = [], field = '', quoted = false;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (quoted) {
      if (c === '"') { if (text[i + 1] === '"') { field += '"'; i++; } else quoted = false; }
      else field += c;
    } else if (c === '"') quoted = true;
    else if (c === ',') { row.push(field); field = ''; }
    else if (c === '\n' || c === '\r') {
      if (c === '\r' && text[i + 1] === '\n') i++;
      row.push(field); rows.push(row); row = []; field = '';
    } else field += c;
  }
  if (field !== '' || row.length) { row.push(field); rows.push(row); }
  const [hdr, ...body] = rows;
  if (!hdr) return [];
  return body.filter(r => r.length > 1 || r[0]).map(r => Object.fromEntries(hdr.map((h, i) => [h, r[i] ?? ''])));
}

ipcMain.handle('xlsxmerge:run', async (evt, p) => {
  const d = new Date();
  const pad = (n) => String(n).padStart(2, '0');
  const ts = `${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}_${pad(d.getHours())}${pad(d.getMinutes())}${pad(d.getSeconds())}`;
  const outDir = path.join(XLSX_MERGE_DATA_DIR, 'out', ts);
  fs.mkdirSync(outDir, { recursive: true });

  const args = ['--branch-ws', p.branchWs, '--trunk-ws', p.trunkWs, '--out', outDir, '--data-dir', XLSX_MERGE_DATA_DIR];
  if (p.base) args.push('--base', String(p.base));
  if (p.files) args.push('--files', p.files);
  if (p.onCollision && p.onCollision !== 'conflict') args.push('--on-collision', p.onCollision);
  if (p.scope) args.push('--scope', p.scope);
  if (p.scopeInclude) args.push('--scope-include', p.scopeInclude);
  if (p.preferBranch) args.push('--prefer-branch', p.preferBranch);
  if (p.preferTrunk) args.push('--prefer-trunk', p.preferTrunk);
  if (p.yieldTrunk) args.push('--yield-trunk', p.yieldTrunk);
  if (p.autoIds) args.push('--auto-ids');
  if (p.switchOff) args.push('--switch-off');
  if (p.applyDeletes) args.push('--apply-deletes');
  if (p.addMissing) args.push('--add-missing');
  if (p.noText) args.push('--no-text');
  if (p.apply) {
    args.push('--apply');
    if (p.unlockLocal) args.push('--unlock-local');
  }
  if (p.remapCsv && p.remapCsv.trim()) {
    const remapPath = path.join(outDir, 'remap.csv');
    fs.writeFileSync(remapPath, p.remapCsv, 'utf-8');
    args.push('--remap', remapPath);
  }

  // Streamed like the pipeline steps: the first run pulls ~45 base workbooks
  // out of SVN and an apply drives Excel for minutes; a silent pane reads as a hang.
  return new Promise((resolve) => {
    let output = '';
    const proc = spawn(PYTHON_EXE, [PYBACKEND_SCRIPT, 'xlsx_merge', ...args], {
      cwd: PYBACKEND_CWD, env: { ...process.env, PYTHONUTF8: '1' }
    });
    proc.stdout.on('data', (d) => { output += d.toString(); evt.sender.send('pipeline:log-line', d.toString()); });
    proc.stderr.on('data', (d) => { output += d.toString(); evt.sender.send('pipeline:log-line', d.toString()); });
    proc.on('close', (code) => {
      const readIf = (f) => { const fp = path.join(outDir, f); return fs.existsSync(fp) ? fs.readFileSync(fp, 'utf-8') : ''; };
      let summary = null;
      try { summary = JSON.parse(readIf('summary.json') || 'null'); } catch (e) { summary = null; }
      resolve({
        success: code === 0, code, output, outDir, summary,
        report: readIf('report.md'),
        conflicts: parseCsv(readIf('conflicts.csv')),
        suggest: parseCsv(readIf('remap_suggest.csv')),
        svnAfter: readIf('svn_after.txt')
      });
    });
    proc.on('error', (err) => resolve({ success: false, code: 1, output: `Failed to launch Python: ${err.message}`, outDir }));
  });
});

ipcMain.handle('xlsxmerge:reveal', async (evt, { outDir }) => {
  const inside = outDir && path.resolve(outDir).toLowerCase().startsWith(XLSX_MERGE_DATA_DIR.toLowerCase());
  if (!inside || !fs.existsSync(outDir)) return { success: false, output: 'That is not a merge report folder.' };
  const err = await shell.openPath(outDir);
  return { success: !err, output: err || '' };
});

// ---- IPC: Chapter Update (new PVE chapter roll-forward) ----

ipcMain.handle('chapter:get', async (evt, { designPath }) => {
  return runPython('get_chapter_info', [designPath]);
});

// Streamed rather than buffered: copying dynamic_map_NNNNN can take a while and
// a silent progress pane reads as a hang.
ipcMain.handle('chapter:run', async (evt, { branchRoot, payload }) => {
  const tmp = tempJsonFile('chapter_payload');
  // The backend cannot work out its own version: in an installed app
  // package.json only exists inside app.asar, so nothing next to
  // resources\pybackend says which build this is. It has to be told.
  //
  // This is here because two reports of the same missing-dependency failure
  // arrived with no way to tell whether the install was old or a fix had not
  // taken - the log named the interpreter but never the build.
  fs.writeFileSync(tmp, JSON.stringify({ ...payload, app_version: app.getVersion() }), 'utf-8');

  return new Promise((resolve) => {
    let output = '';
    const proc = spawn(PYTHON_EXE, [PYBACKEND_SCRIPT, 'update_chapter', branchRoot, tmp], { cwd: PYBACKEND_CWD });
    proc.stdout.on('data', (d) => { output += d.toString(); evt.sender.send('pipeline:log-line', d.toString()); });
    proc.stderr.on('data', (d) => { output += d.toString(); evt.sender.send('pipeline:log-line', d.toString()); });
    proc.on('close', (code) => resolve({ success: code === 0, code, output }));
    proc.on('error', (err) => resolve({ success: false, code: 1, output: `Failed to launch Python: ${err.message}` }));
  });
});

// ---- IPC: Event type / history lookup (Main Events tab source picker) ----

ipcMain.handle('eventtypes:get', async (evt, { designPath }) => {
  return runPython('get_events_by_type', [designPath]);
});

// ---- IPC: manual "Update local cache from SVN" (convert_layout.lua source-event data) ----

ipcMain.handle('convertLayoutCache:refresh', async (evt, { branchRoot, workspace, eventNames }) => {
  const problem = branchRootProblem(branchRoot);
  if (problem) return { success: false, code: 1, output: problem };
  const logDir = resolveLogDir(branchRoot);
  fs.mkdirSync(logDir, { recursive: true });

  // Streamed, same as pipeline steps — this can hit the slow SVN-history path for
  // events that aren't live and aren't cached yet, so silence would look hung.
  return new Promise((resolve) => {
    let output = '';
    const proc = spawn(PYTHON_EXE, [
      PYBACKEND_SCRIPT, 'refresh_convert_layout_cache', branchRoot, logDir, workspace, JSON.stringify(eventNames)
    ], { cwd: PYBACKEND_CWD });

    proc.stdout.on('data', (d) => { output += d.toString(); evt.sender.send('pipeline:log-line', d.toString()); });
    proc.stderr.on('data', (d) => { output += d.toString(); evt.sender.send('pipeline:log-line', d.toString()); });
    proc.on('close', (code) => resolve({ success: code === 0, code, output }));
    proc.on('error', (err) => resolve({ success: false, code: 1, output: `Failed to launch Python: ${err.message}` }));
  });
});

// ---- IPC: Ideation tab — 5-theme generation + feedback-driven regeneration ----
// Calls the real Anthropic API (metered Console key from Settings, not a
// claude.ai subscription — third-party OAuth passthrough of a Claude
// subscription is against Anthropic's terms and wouldn't work here anyway).

// Unlike the game-data payload temp files above (never cleaned up — harmless
// if left behind), these two carry a live Anthropic API key. Delete them as
// soon as the Python process has read them rather than leaving a plaintext
// credential sitting in the OS temp dir indefinitely.
async function runPythonAndScrub(scriptId, tmpPath) {
  try {
    return await runPython(scriptId, [tmpPath]);
  } finally {
    fs.unlink(tmpPath, () => {});
  }
}

// Claude always wins when a key is present; the Ollama fallback only kicks
// in when no Anthropic key is configured at all — never a per-run choice,
// so there's no ambiguity about which one produced a given draft.
function resolveIdeationProvider(cfg) {
  if (cfg.anthropicApiKey) return { provider: 'anthropic' };
  if (cfg.ollamaBaseUrl) return { provider: 'ollama' };
  return { provider: null };
}

ipcMain.handle('ideation:generate', async (evt, { eventType, durationDays, themeCount, designPath }) => {
  const cfg = loadConfig();
  const { provider } = resolveIdeationProvider(cfg);
  if (!provider) {
    return { success: false, output: 'Add an Anthropic API key or an Ollama Base URL in Settings first.' };
  }
  const count = Math.max(1, Math.min(10, parseInt(themeCount, 10) || 5));
  const tmp = tempJsonFile('ideation_generate');
  fs.writeFileSync(tmp, JSON.stringify({
    event_type: eventType,
    duration_days: durationDays,
    theme_count: count,
    design_path: designPath,
    provider,
    author_id: cfg.authorId,
    author_name: cfg.authorName,
    api_key: cfg.anthropicApiKey,
    ollama_base_url: cfg.ollamaBaseUrl,
    ollama_model: cfg.ollamaModel || 'qwen3:32b',
    events_root: ideationEventsRoot()
  }), 'utf-8');
  return runPythonAndScrub('ideation_generate', tmp);
});

ipcMain.handle('ideation:regenerate', async (evt, { eventType, durationDays, theme, feedback }) => {
  const cfg = loadConfig();
  const { provider } = resolveIdeationProvider(cfg);
  if (!provider) {
    return { success: false, output: 'Add an Anthropic API key or an Ollama Base URL in Settings first.' };
  }
  const tmp = tempJsonFile('ideation_regenerate');
  fs.writeFileSync(tmp, JSON.stringify({
    event_type: eventType,
    duration_days: durationDays,
    provider,
    author_id: cfg.authorId,
    author_name: cfg.authorName,
    api_key: cfg.anthropicApiKey,
    ollama_base_url: cfg.ollamaBaseUrl,
    ollama_model: cfg.ollamaModel || 'qwen3:32b',
    events_root: ideationEventsRoot(),
    theme,
    feedback
  }), 'utf-8');
  return runPythonAndScrub('ideation_regenerate', tmp);
});

// Settings "Test Connection" — confirms the configured Ollama server is
// reachable and has the configured model pulled, before the user ever
// starts an ideation run against it.
ipcMain.handle('ideation:testOllama', async (evt, { baseUrl, model }) => {
  if (!baseUrl) return { success: false, output: 'Enter an Ollama Base URL first.' };
  try {
    const res = await fetch(`${baseUrl.replace(/\/+$/, '')}/api/tags`, { signal: AbortSignal.timeout(5000) });
    if (!res.ok) return { success: false, output: `HTTP ${res.status} from ${baseUrl}` };
    const data = await res.json();
    const names = (data.models || []).map((m) => m.name);
    const hasModel = names.includes(model);
    if (!names.length) return { success: false, output: 'Reachable, but no models are pulled on that server.' };
    return {
      success: true,
      output: hasModel
        ? `Connected. "${model}" is available.`
        : `Connected, but "${model}" isn't pulled there yet. Available: ${names.join(', ')}`
    };
  } catch (e) {
    return { success: false, output: `Could not reach ${baseUrl}: ${e.message}` };
  }
});

// ---- PRESENCE -------------------------------------------------------------
// Two designers pointed at the same shared folder could both work an event and
// neither would know. This is the smallest honest fix: each machine drops a
// file saying who it is and when it last checked in, and everyone reads
// everyone else's.
//
// It is ADVISORY and deliberately so. Nothing is blocked and nothing has to be
// released, because a hard lock on a synced folder fails in the one way you
// cannot recover from in an office: a laptop sleeps or the app crashes, the lock
// survives, and someone has to be talked through deleting a file. A record that
// simply ages out has no such state to get stuck in.
//
// The numbers are set by OneDrive rather than by taste. A write here is not
// visible on another machine until the folder syncs, which is seconds at best
// and minutes on a slow link, so the timeout has to be generous enough that a
// working colleague does not flicker in and out of the banner. Thirty seconds
// between check-ins against a five-minute timeout means ten missed writes before
// someone disappears.
const PRESENCE_TTL_MS = 5 * 60 * 1000;

function presenceDir(eventType) {
  return path.join(ideationEventsRoot(), eventType, '.active');
}

// Windows forbids these in a filename, and an AzureAD id is DOMAIN\user.
function presenceFileName(authorId) {
  return (authorId || 'unknown').replace(/[\\/:*?"<>|]/g, '_') + '.json';
}

ipcMain.handle('presence:announce', async (evt, { eventType, detail }) => {
  try {
    const cfg = loadConfig();
    if (!eventType || !cfg.authorId) return { success: false, output: 'No event type or identity.' };
    const dir = presenceDir(eventType);
    fs.mkdirSync(dir, { recursive: true });
    const file = path.join(dir, presenceFileName(cfg.authorId));
    // startedAt survives a refresh so the banner can say how long someone has
    // been in, not merely that they pinged a moment ago.
    let startedAt = Date.now();
    try {
      const prev = JSON.parse(fs.readFileSync(file, 'utf8'));
      if (prev && prev.startedAt && (Date.now() - prev.lastSeen) < PRESENCE_TTL_MS) {
        startedAt = prev.startedAt;
      }
    } catch { /* first check-in, or a half-written file mid-sync */ }
    fs.writeFileSync(file, JSON.stringify({
      authorId: cfg.authorId,
      authorName: cfg.authorName || cfg.authorId,
      avatarColor: cfg.avatarColor || '',
      eventType,
      detail: detail || '',
      startedAt,
      lastSeen: Date.now(),
    }), 'utf8');
    return { success: true };
  } catch (e) {
    // A sync conflict or a locked file must never take the UI down with it -
    // presence is a courtesy, not a feature anyone is blocked on.
    return { success: false, output: e.message };
  }
});

ipcMain.handle('presence:list', async (evt, { eventType }) => {
  try {
    const cfg = loadConfig();
    const dir = presenceDir(eventType);
    if (!fs.existsSync(dir)) return { success: true, others: [] };
    const now = Date.now();
    const others = [];
    for (const name of fs.readdirSync(dir)) {
      if (!name.endsWith('.json')) continue;
      let rec;
      try { rec = JSON.parse(fs.readFileSync(path.join(dir, name), 'utf8')); }
      catch { continue; }            // mid-sync write, ignore it this round
      if (!rec || !rec.authorId) continue;
      if (rec.authorId === cfg.authorId) continue;          // that is you
      if (!rec.lastSeen || (now - rec.lastSeen) > PRESENCE_TTL_MS) continue;
      others.push({
        authorId: rec.authorId,
        authorName: rec.authorName || rec.authorId,
        avatarColor: rec.avatarColor || '',
        detail: rec.detail || '',
        sinceMs: now - (rec.startedAt || rec.lastSeen),
        lastSeenMs: now - rec.lastSeen,
      });
    }
    others.sort((a, b) => a.authorName.localeCompare(b.authorName));
    return { success: true, others };
  } catch (e) {
    return { success: false, output: e.message };
  }
});

ipcMain.handle('presence:clear', async (evt, { eventType }) => {
  try {
    const cfg = loadConfig();
    if (!eventType || !cfg.authorId) return { success: true };
    fs.unlinkSync(path.join(presenceDir(eventType), presenceFileName(cfg.authorId)));
    return { success: true };
  } catch (e) {
    // Already gone is the normal case when leaving a view you never announced.
    return { success: true, output: e.message };
  }
});

// Leaving without checking out is the common case - the window is closed, the
// machine sleeps, the app crashes. Clearing on the way out is a courtesy that
// makes the banner accurate sooner; the TTL is what actually guarantees it.
app.on('before-quit', () => {
  try {
    const cfg = loadConfig();
    if (!cfg.authorId) return;
    const root = ideationEventsRoot();
    if (!fs.existsSync(root)) return;
    for (const type of fs.readdirSync(root)) {
      const f = path.join(root, type, '.active', presenceFileName(cfg.authorId));
      try { if (fs.existsSync(f)) fs.unlinkSync(f); } catch { /* best effort */ }
    }
  } catch { /* never block the quit */ }
});

// What counts as a One-Pager, and how a name clash is resolved, both live in
// lib/onepager.js - the upload starts with a native dialog no driver can click,
// so the rules downstream of it were only checkable by reading them. Moved out
// so tests/onepager_test.js can call them directly.
const onePager = require('./lib/onepager');

// Every One-Pager already sitting in the shared folder for this event type,
// whichever machine wrote it. This is what lets someone who ALREADY has a
// signed-off One-Pager go straight to the Art Requirement step instead of being
// forced to generate throwaway themes first — including a One-Pager a colleague
// finalized, since the folder is shared.
//
// All four formats are usable. .md and .txt are read directly; .docx and .pdf
// are converted to Markdown on first use. That matters because the real shipped
// One-Pagers ARE PDFs — listing them as unusable was the wrong answer to the
// right observation that they belong in this list.
//
// A converted sibling is hidden rather than listed: "Merge2608 One-Pager.pdf"
// and its generated "Merge2608 One-Pager.md" are one document, and showing both
// invites picking the stale one after the original has been edited.
ipcMain.handle('ideation:listOnePagers', async (evt, { eventType }) => {
  try {
    const dir = path.join(ideationEventsRoot(), eventType, 'One-Pager Ideation');
    if (!fs.existsSync(dir)) return { success: true, docs: [] };
    const all = fs.readdirSync(dir);
    const generated = new Set(
      all.filter(n => onePager.kindOf(n).convertible)
         .map(n => path.basename(onePager.markdownSiblingOf(n))));
    const docs = all
      .filter(n => onePager.kindOf(n).storable && !generated.has(n))
      .map(n => {
        const p = path.join(dir, n);
        let st; try { st = fs.statSync(p); } catch { return null; }
        if (!st.isFile()) return null;
        return {
          name: n,
          path: p,
          modified: st.mtimeMs,
          bytes: st.size,
          ...onePager.kindOf(n),
        };
      })
      .filter(Boolean)
      .sort((a, b) => b.modified - a.modified);
    return { success: true, docs };
  } catch (e) {
    return { success: false, output: e.message };
  }
});

// Copy a One-Pager the designer already has into the shared folder for this
// event type. Listing only what was already filed left an obvious hole: a
// One-Pager that arrived by email, or sits on a desktop, had to be moved by hand
// into a path most people would have to be told - and being told a path is the
// step where it lands in the wrong folder.
//
// Word and PDF are accepted on equal terms with Markdown: doc_convert.py turns
// either into Markdown on first use, so what the designer has is what they can
// upload. Only the conversion is deferred - see ideation:readOnePager.
ipcMain.handle('ideation:uploadOnePager', async (evt, { eventType }) => {
  try {
    if (!eventType) return { success: false, output: 'Pick an event type first.' };
    const win = BrowserWindow.getFocusedWindow();
    const result = await dialog.showOpenDialog(win, {
      title: `Choose the ${eventType} One-Pager`,
      properties: ['openFile'],
      filters: [
        { name: 'One-Pager (Word, PDF, Markdown, text)', extensions: onePager.STORABLE },
        { name: 'All files', extensions: ['*'] },
      ],
    });
    if (result.canceled || !result.filePaths.length) return { success: false, canceled: true };

    const src = result.filePaths[0];
    const kind = onePager.kindOf(src);
    if (!kind.storable) {
      return { success: false, output: `A .${kind.ext} file cannot be used as a One-Pager. Choose a Word, PDF, Markdown or text file.` };
    }
    const ext = kind.ext;

    const dir = path.join(ideationEventsRoot(), eventType, 'One-Pager Ideation');
    fs.mkdirSync(dir, { recursive: true });
    let dest = path.join(dir, path.basename(src));

    // Never overwrite a colleague's One-Pager without being told to. The folder
    // is shared, so the file already there may not be one this machine put there.
    if (fs.existsSync(dest)) {
      const ans = await dialog.showMessageBox(win, {
        type: 'question',
        buttons: ['Replace', 'Keep both', 'Cancel'],
        defaultId: 1,
        cancelId: 2,
        title: 'A One-Pager with that name is already filed',
        message: `"${path.basename(src)}" already exists in the ${eventType} One-Pager folder.`,
        detail: 'The folder is shared, so this may be a colleague\'s. "Keep both" files yours alongside it.',
      });
      if (ans.response === 2) return { success: false, canceled: true };
      if (ans.response === 1) {
        const free = onePager.nextFreeName(
          path.basename(src, path.extname(src)), ext,
          (n) => fs.existsSync(path.join(dir, n)));
        dest = path.join(dir, free);
      }
    }

    fs.copyFileSync(src, dest);
    // Conversion is left to the read step rather than done here, so there is one
    // conversion path instead of two and a PDF already filed by someone else
    // behaves exactly like one just uploaded.
    return { success: true, name: path.basename(dest), path: dest,
             readable: kind.readable, convertible: kind.convertible,
             usable: kind.usable, ext };
  } catch (e) {
    return { success: false, output: e.message };
  }
});

ipcMain.handle('ideation:readOnePager', async (evt, { docPath }) => {
  try {
    const root = ideationEventsRoot();
    const resolved = path.resolve(docPath);
    // Only ever read out of the ideation folder the user has configured - a path
    // arriving from the renderer is not a licence to read anywhere on disk.
    if (!resolved.startsWith(path.resolve(root) + path.sep)) {
      return { success: false, output: 'That file is outside the ideation data folder.' };
    }
    const kind = onePager.kindOf(resolved);
    if (!kind.usable) {
      return { success: false, output: `A .${kind.ext} file is not a One-Pager format. `
        + 'Use Markdown, text, Word or PDF.' };
    }

    // A .docx or .pdf is converted to Markdown first, beside the original. The
    // result is cached: converting again only happens when the source has been
    // edited since, which is what makes re-uploading a corrected brief work
    // without anyone having to know a cache exists.
    let toRead = resolved;
    let converted = null;
    let convertStats = null;
    if (kind.convertible) {
      const md = onePager.markdownSiblingOf(resolved);
      const fresh = fs.existsSync(md)
        && fs.statSync(md).mtimeMs >= fs.statSync(resolved).mtimeMs;
      if (!fresh) {
        const tmp = tempJsonFile('doc_convert');
        fs.writeFileSync(tmp, JSON.stringify({ src: resolved, dest: md }), 'utf-8');
        const res = await runPythonAndScrub('doc_convert', tmp);
        if (!res.success) return { success: false, output: res.output };
        try {
          const parsed = JSON.parse(res.stdout);
          if (parsed.error) return { success: false, output: parsed.error };
          // What happened to the images, so the renderer can say so. A brief
          // whose content sits in pictures is the case this whole path exists
          // for, and silence about it is what the first version got wrong.
          convertStats = {
            images: parsed.images || 0,
            ocrImages: parsed.ocr_images || 0,
            ocrLines: parsed.ocr_lines || 0,
            skipped: parsed.skipped || 0,
          };
        } catch { /* a non-JSON line is fine as long as the file landed */ }
        if (!fs.existsSync(md)) {
          return { success: false, output: `Converting ${path.basename(resolved)} produced nothing.` };
        }
        converted = path.basename(md);
      }
      toRead = md;
    }

    const text = fs.readFileSync(toRead, 'utf8');
    if (!text.trim()) return { success: false, output: 'That One-Pager is empty.' };
    // `converted` is the .md just written, so the renderer can say what happened
    // rather than leaving a Word file to look as if it were read directly.
    // docPath is the file actually READ - the generated .md for a Word or PDF
    // brief - because that is where its extracted images live.
    return { success: true, text, name: path.basename(resolved), docPath: toRead,
             converted, convertStats };
  } catch (e) {
    return { success: false, output: e.message };
  }
});

// "Select as Final" — writes the theme as a Markdown 1-Pager into the same
// per-type "One-Pager Ideation" folder the real PDFs already live in (e.g.
// ideationData/Events/Merge/One-Pager Ideation/Merge2608 One-Pager Ideation.pdf).
// Pure main-process file write — no Python/Excel involved, the theme JSON
// already has everything needed.
ipcMain.handle('ideation:saveFinal', async (evt, { eventType, durationDays, theme }) => {
  try {
    const cfg = loadConfig();
    const dir = path.join(ideationEventsRoot(), eventType, 'One-Pager Ideation');
    fs.mkdirSync(dir, { recursive: true });
    const refName = (theme.internal_reference_name || 'UnnamedEvent').replace(/[\\/:*?"<>|]/g, '_');
    const filePath = path.join(dir, `${refName} One-Pager Ideation.md`);

    const af = theme.audience_fit || {};
    const vd = theme.visual_direction || {};
    const imp = theme.expected_impact || {};
    const eff = theme.effort_estimate || {};
    const scenes = (theme.scenes || []).map((s, i) => `${i + 1}. **${s.name}** — ${s.description}`).join('\n');
    const keyFeatures = (vd.key_features || []).map(f => `- ${f}`).join('\n');
    const uiDirection = (vd.ui_direction || []).map(f => `- ${f}`).join('\n');

    const md = `# Event Ideation 1-Pager — ${theme.theme_name}

**Internal Reference Name:** ${theme.internal_reference_name}
**Event Type:** ${eventType}
**Sprint Duration:** ${durationDays} days
**Generated:** ${new Date().toISOString().slice(0, 10)}
**Selected as Final by:** ${cfg.authorName || 'Unknown'}
**Status:** Selected as Final (draft — pending PO review, per the studio's Event Ideation Checklist)

## Theme & Reasoning

${theme.reasoning}

## Audience Fit

- Progression Completionist: ${af.progression_completionist}
- Casual Escape: ${af.casual_escape}
- Social/Collector: ${af.social_collector}

${af.notes || ''}

## Visual Direction

**Art Style:** ${vd.art_style}

**Color Palette:** ${vd.color_palette}

**Key Features:**
${keyFeatures}

**UI Direction:**
${uiDirection}

## Scene Progression

${scenes}

## New Features & Additions

${theme.new_features}

## Expected Impact (draft estimate — needs PO/Economy Designer review)

- Engagement: ${imp.engagement_note}
- Monetization: ${imp.monetization_note}

## Effort Estimation

| Area | Effort |
|---|---|
| Art | ${eff.art} |
| Engineering | ${eff.engineering} |
| Design | ${eff.design} |
| Overall | ${eff.overall} |

---
Generated by the Event Automator Ideation tab (Claude API). Treat as a first draft for PO review, not an approved spec.
`;
    fs.writeFileSync(filePath, md, 'utf-8');
    return { success: true, output: filePath };
  } catch (e) {
    return { success: false, output: e.message };
  }
});

// ---- IPC: Art Requirement tab — generated only from a finalized Ideation
// theme (payload includes the full theme JSON, not just its name, so the
// Art doc is grounded in the actual approved reasoning/visual direction).
// Rendering itself lives in Python (art_requirement_generator.py) so the
// same real-template-verified structure backs drafts, regenerations, and
// the final save — no second markdown template to keep in sync in JS. ----

ipcMain.handle('artreq:generate', async (evt, { eventType, ideationTheme, ideationDocText, ideationDocName, ideationDocPath, variationCount }) => {
  const cfg = loadConfig();
  const { provider } = resolveIdeationProvider(cfg);
  if (!provider) {
    return { success: false, output: 'Add an Anthropic API key or an Ollama Base URL in Settings first.' };
  }
  const count = Math.max(1, Math.min(10, parseInt(variationCount, 10) || 3));
  const tmp = tempJsonFile('artreq_generate');
  fs.writeFileSync(tmp, JSON.stringify({
    event_type: eventType,
    ideation_theme: ideationTheme,
    // Set instead of ideation_theme when the designer already had a One-Pager.
    // The backend takes either.
    ideation_doc_text: ideationDocText || '',
    ideation_doc_name: ideationDocName || '',
    // Lets the backend find the brief's extracted reference art and attach it.
    ideation_doc_path: ideationDocPath || '',
    variation_count: count,
    provider,
    author_id: cfg.authorId,
    author_name: cfg.authorName,
    api_key: cfg.anthropicApiKey,
    ollama_base_url: cfg.ollamaBaseUrl,
    ollama_model: cfg.ollamaModel || 'qwen3:32b',
    events_root: ideationEventsRoot()
  }), 'utf-8');
  return runPythonAndScrub('artreq_generate', tmp);
});

ipcMain.handle('artreq:regenerate', async (evt, { eventType, variation, variationIndex, feedback }) => {
  const cfg = loadConfig();
  const { provider } = resolveIdeationProvider(cfg);
  if (!provider) {
    return { success: false, output: 'Add an Anthropic API key or an Ollama Base URL in Settings first.' };
  }
  const tmp = tempJsonFile('artreq_regenerate');
  fs.writeFileSync(tmp, JSON.stringify({
    event_type: eventType,
    variation,
    variation_index: variationIndex,
    feedback,
    provider,
    author_id: cfg.authorId,
    author_name: cfg.authorName,
    api_key: cfg.anthropicApiKey,
    ollama_base_url: cfg.ollamaBaseUrl,
    ollama_model: cfg.ollamaModel || 'qwen3:32b',
    events_root: ideationEventsRoot()
  }), 'utf-8');
  return runPythonAndScrub('artreq_regenerate', tmp);
});

ipcMain.handle('artreq:saveFinal', async (evt, { eventType, refName, variation }) => {
  const cfg = loadConfig();
  const tmp = tempJsonFile('artreq_save_final');
  fs.writeFileSync(tmp, JSON.stringify({
    event_type: eventType,
    ref_name: refName,
    variation,
    author_name: cfg.authorName,
    events_root: ideationEventsRoot()
  }), 'utf-8');
  const res = await runPythonAndScrub('artreq_save_final', tmp);
  if (!res.success) return { success: false, output: res.output };
  try {
    const parsed = JSON.parse(res.stdout);
    if (parsed.error) return { success: false, output: parsed.error };
    return { success: true, output: parsed.path };
  } catch (e) {
    return { success: false, output: `Could not parse save result: ${e.message}` };
  }
});

// Renders the variation through the same Python writer the draft and final
// files use, so what a reviewer reads on screen is the document that lands on
// disk - not a second, drifting reimplementation of it in the renderer.
ipcMain.handle('artreq:preview', async (evt, { eventType, variation, isFinal }) => {
  const cfg = loadConfig();
  const tmp = tempJsonFile('artreq_preview');
  fs.writeFileSync(tmp, JSON.stringify({
    event_type: eventType,
    variation,
    is_final: !!isFinal,
    author_name: cfg.authorName,
    events_root: ideationEventsRoot()
  }), 'utf-8');
  return parseJsonResult(await runPythonAndScrub('artreq_preview', tmp), 'document preview');
});

// ---- IPC: Art Requirement documents already on disk ----
// The tab used to show only what the current session generated, which meant a
// document finalised yesterday - or one written straight into the shared folder,
// or filed there by a colleague on another machine - was invisible to everyone,
// including the person who put it there. Listing the folder is what makes this a
// view of the event's art requirements rather than a view of this session's
// scratch work.
//
// Drafts are listed alongside finals, labelled, because "which of these is the
// current one" is answered by the label and the date, not by the reader
// remembering which subfolder means what.

const ARTREQ_DOC_EXT = ['.md', '.docx', '.pdf'];

function artReqDirs(eventType) {
  const base = path.join(ideationEventsRoot(), eventType, 'Art Requirements');
  return [{ dir: base, folder: '' }, { dir: path.join(base, 'Drafts'), folder: 'Drafts' }];
}

// A path arriving from the renderer is not a licence to touch anywhere on disk,
// the same rule ideation:readOnePager follows. Resolve first, then check, so
// '..' segments cannot walk out of the folder.
function insideArtReq(docPath, eventType) {
  if (!docPath || !eventType) return null;
  const base = path.resolve(path.join(ideationEventsRoot(), eventType, 'Art Requirements'));
  const resolved = path.resolve(docPath);
  return resolved.startsWith(base + path.sep) ? resolved : null;
}

// The event number a filename is talking about: "Block2609 - Event One-Pager.md"
// and "Blocks2609 Harvest Fair - Variation 1.md" both yield "2609".
//
// Matching on the number rather than the type word is deliberate: the studio
// writes both "Block2609" and "Blocks2609" for the same event - the One-Pager
// uses the singular, the shipped Art Requirements the plural - and a match that
// insisted on the exact spelling would quietly filter everything out.
// A One-Pager is picked as the file the designer sees - often "X.pdf" - while
// what actually gets read, and therefore recorded, is the converted "X.md"
// beside it. Comparing the full names would never match. The stem does.
function onePagerStem(name) {
  const n = String(name || '').trim();
  if (!n) return '';
  return n.replace(/\.(md|txt|docx|pdf)$/i, '').toLowerCase();
}

function eventNumberOf(name) {
  const m = /(?:^|[^0-9])(\d{4})(?:[^0-9]|$)/.exec(String(name || ''));
  return m ? m[1] : null;
}

ipcMain.handle('artreq:listDocs', async (evt, { eventType, onePagerName, showAll }) => {
  try {
    if (!eventType) return { success: true, docs: [] };
    // With a One-Pager picked, the panel answers "what exists for THIS event",
    // not "what is in the folder". A Blocks folder accumulates every Blocks
    // event ever run, so an unfiltered list buries the documents someone is
    // working on under a dozen they are not.
    const wanted = showAll ? null : eventNumberOf(onePagerName);

    const all = [];
    for (const { dir, folder } of artReqDirs(eventType)) {
      if (!fs.existsSync(dir)) continue;
      for (const name of fs.readdirSync(dir)) {
        const ext = path.extname(name).toLowerCase();
        if (!ARTREQ_DOC_EXT.includes(ext)) continue;
        const p = path.join(dir, name);
        let st;
        try { st = fs.statSync(p); } catch { continue; }
        if (!st.isFile()) continue;
        all.push({ name, path: p, folder, ext, st });
      }
    }

    // A .docx exported from a listed .md is not a separate document - it is the
    // same document in the format Word opens, and the .md row already carries a
    // "Show .docx" button. Listing both doubled the panel and made people choose
    // between two names for one thing. A .docx with no .md beside it still
    // shows, because then it is the only copy there is.
    const mdStems = new Set(all.filter(f => f.ext === '.md')
                               .map(f => f.path.slice(0, -3).toLowerCase()));
    const visible = all.filter(f => !(f.ext === '.docx'
                                   && mdStems.has(f.path.slice(0, -5).toLowerCase())));
    const hiddenAsDocx = all.length - visible.length;

    const docs = [];
    let hiddenByEvent = 0;
    let matchedBySource = 0;
    for (const f of visible) {
      const num = eventNumberOf(f.name);
      const isMd = f.ext === '.md';
      // Provenance first, filename second. Every document saved by the tool now
      // records the One-Pager it was written from, so the match can be a fact
      // rather than a number scraped out of a filename. The number stays as the
      // fallback for documents that predate the field or came from elsewhere.
      let source = null;
      if (isMd) {
        const data = f.path.slice(0, -3) + '.json';
        try {
          if (fs.existsSync(data)) {
            source = JSON.parse(fs.readFileSync(data, 'utf-8')).source_one_pager || null;
          }
        } catch { /* an unreadable sidecar just means falling back to the name */ }
      }
      if (wanted) {
        const keep = source
          ? onePagerStem(source) === onePagerStem(onePagerName)
          : num === wanted;
        if (!keep) { hiddenByEvent++; continue; }
        if (source) matchedBySource++;
      }
      docs.push({
        name: f.name, path: f.path, folder: f.folder, ext: f.ext.slice(1),
        modified: f.st.mtimeMs, bytes: f.st.size,
        eventNumber: num,
        sourceOnePager: source,
        // Only Markdown can be reopened, exported or have its art generated. A
        // .pdf here is a finished artefact the production team produced, so it
        // is listed for context and otherwise left alone.
        readable: isMd,
        hasData: isMd && fs.existsSync(f.path.slice(0, -3) + '.json'),
        docxPath: isMd && fs.existsSync(f.path.slice(0, -3) + '.docx')
          ? f.path.slice(0, -3) + '.docx' : null,
      });
    }
    docs.sort((a, b) => b.modified - a.modified);
    return { success: true, docs, filteredTo: wanted, hiddenByEvent, hiddenAsDocx,
             matchedBySource };
  } catch (e) {
    return { success: false, output: e.message };
  }
});

ipcMain.handle('artreq:readDoc', async (evt, { eventType, docPath }) => {
  try {
    const resolved = insideArtReq(docPath, eventType);
    if (!resolved) return { success: false, output: 'That file is outside the Art Requirements folder.' };
    if (path.extname(resolved).toLowerCase() !== '.md') {
      return { success: false, output: 'Only a Markdown document can be opened here.' };
    }
    return { success: true, data: { markdown: fs.readFileSync(resolved, 'utf-8') } };
  } catch (e) {
    return { success: false, output: e.message };
  }
});

// Word export. The production team marks these up and sends them back, and they
// do that in Word - so a document that only exists as Markdown is a document
// they cannot review. doc_convert.markdown_to_docx builds the .docx from the
// file on disk, which means what gets exported is exactly what was saved, not a
// second rendering of the variation held in the renderer.
ipcMain.handle('artreq:exportDocx', async (evt, { eventType, docPath }) => {
  try {
    const resolved = insideArtReq(docPath, eventType);
    if (!resolved) return { success: false, output: 'That file is outside the Art Requirements folder.' };
    if (path.extname(resolved).toLowerCase() !== '.md') {
      return { success: false, output: 'Only a Markdown document can be exported to Word.' };
    }
    if (!fs.existsSync(resolved)) {
      return { success: false, output: 'That document is no longer on disk. Save it first.' };
    }
    const tmp = tempJsonFile('md_to_docx');
    fs.writeFileSync(tmp, JSON.stringify({ src: resolved }), 'utf-8');
    return parseJsonResult(await runPythonAndScrub('md_to_docx', tmp), 'Word export');
  } catch (e) {
    return { success: false, output: e.message };
  }
});

// Reveal the exported file in Explorer. Telling someone a path and leaving them
// to find it is where a shared network folder turns into "I can't see it".
// Generating art for a document that was OPENED rather than just generated.
//
// The image pass consumes the structured document, not the Markdown - it needs
// zones, construction steps, materials and the art direction as data, and
// several fields it relies on (blunt_steps, character_scale, render_aspect) are
// never rendered into the Markdown at all. So a .json sidecar is written beside
// every saved document, and these two channels read it back and save it again
// once the renders have added image paths.
//
// A document with no sidecar - written by hand, or by a version of the tool that
// predates this - simply does not offer the button. That is honest: the
// alternative is parsing the Markdown back into a document and being wrong in
// exactly the places nobody checks.

ipcMain.handle('artreq:docData', async (evt, { eventType, docPath }) => {
  try {
    const resolved = insideArtReq(docPath, eventType);
    if (!resolved) return { success: false, output: 'That file is outside the Art Requirements folder.' };
    const data = resolved.replace(/\.md$/i, '.json');
    if (!fs.existsSync(data)) {
      return { success: true, data: { variation: null,
        reason: 'This document has no structured data beside it, so its art cannot be '
              + 'generated from here. Documents saved by the tool from now on will have it.' } };
    }
    return { success: true, data: { variation: JSON.parse(fs.readFileSync(data, 'utf-8')) } };
  } catch (e) {
    return { success: false, output: e.message };
  }
});

// Save an opened document back in place after an image pass. Deliberately writes
// to the SAME path rather than filing a new draft: the renders belong to the
// document the reviewer already has open, and a second copy beside it is how a
// shared folder fills with near-identical files nobody can choose between.
ipcMain.handle('artreq:saveDoc', async (evt, { eventType, docPath, variation, status }) => {
  try {
    const resolved = insideArtReq(docPath, eventType);
    if (!resolved) return { success: false, output: 'That file is outside the Art Requirements folder.' };
    if (!fs.existsSync(resolved)) return { success: false, output: 'That document is no longer on disk.' };
    const tmp = tempJsonFile('artreq_save_doc');
    // No author_name and no default status: this is an in-place update of a
    // document that already has a header. Sending the current user's name would
    // reassign authorship of a colleague's document just for adding art, and
    // forcing 'Draft' would erase a review state this call knows nothing about.
    // Python carries both forward from the file unless a caller overrides them.
    fs.writeFileSync(tmp, JSON.stringify({
      md_path: resolved,
      event_type: eventType,
      variation,
      status: status || '',
    }), 'utf-8');
    return parseJsonResult(await runPythonAndScrub('artreq_save_doc', tmp), 'document save');
  } catch (e) {
    return { success: false, output: e.message };
  }
});

ipcMain.handle('artreq:revealDoc', async (evt, { eventType, docPath }) => {
  const resolved = insideArtReq(docPath, eventType);
  if (!resolved || !fs.existsSync(resolved)) {
    return { success: false, output: 'That file is not in the Art Requirements folder.' };
  }
  shell.showItemInFolder(resolved);
  return { success: true };
});

// ---- IPC: shared feedback thread + lesson management ----
// Read-only listing of who said what (and the lessons currently steering
// generation), plus retraction so a bad lesson stops poisoning every future
// run. Both are plain reads/writes of the shared entry files — no LLM call,
// so they're cheap enough to refresh whenever a card is rendered.

function parseJsonResult(res, label) {
  if (!res.success) return { success: false, output: res.output };
  try {
    const parsed = JSON.parse(res.stdout);
    if (parsed.error) return { success: false, output: parsed.error };
    return { success: true, data: parsed };
  } catch (e) {
    return { success: false, output: `Could not parse ${label}: ${e.message}` };
  }
}

ipcMain.handle('ideation:entries', async (evt, { eventType, docKind, refName }) => {
  const tmp = tempJsonFile('ideation_entries');
  fs.writeFileSync(tmp, JSON.stringify({
    event_type: eventType,
    doc_kind: docKind || null,
    ref_name: refName || null,
    events_root: ideationEventsRoot()
  }), 'utf-8');
  return parseJsonResult(await runPythonAndScrub('ideation_entries', tmp), 'feedback entries');
});

ipcMain.handle('ideation:retract', async (evt, { eventType, entryId, reason }) => {
  const cfg = loadConfig();
  const tmp = tempJsonFile('ideation_retract');
  fs.writeFileSync(tmp, JSON.stringify({
    event_type: eventType,
    entry_id: entryId,
    reason: reason || '',
    author_id: cfg.authorId,
    author_name: cfg.authorName,
    events_root: ideationEventsRoot()
  }), 'utf-8');
  return parseJsonResult(await runPythonAndScrub('ideation_retract', tmp), 'retraction result');
});

// ---- IPC: sprint board (landing page) ----
// Joins the Feishu roadmap export (the plan) against the design workbooks of
// the active branch (what's actually configured). Read-only — it never writes
// anything, so it's safe to refresh on every visit.

function roadmapPath() {
  return path.join(ideationDataRoot(), 'Roadmap', 'Roadmap.xlsx');
}

// ---- Reference images (ComfyUI / Nano Banana) ----
// These call the same backend the document flow uses, so the zone image paths
// stamped onto the variation come straight back to the renderer and end up
// embedded in the .md when the variation is saved.

ipcMain.handle('comfy:check', async () => {
  const cfg = loadConfig();
  const tmp = tempJsonFile('comfy_check');
  fs.writeFileSync(tmp, JSON.stringify({ comfy_url: cfg.comfyUrl || '' }), 'utf-8');
  return runPythonAndScrub('comfy_check', tmp);
});

// Passes beyond 'scene' are opt-in: a full Merge event is ~96 renders at roughly
// half a minute each, so the renderer asks artreq:plan for the cost first and the
// user chooses what to spend. Omitting passes keeps the original scene-only behaviour.
ipcMain.handle('artreq:plan', async (evt, { variation, passes }) => {
  const tmp = tempJsonFile('artreq_plan');
  fs.writeFileSync(tmp, JSON.stringify({ doc: variation, passes: passes || null }), 'utf-8');
  return runPythonAndScrub('artreq_plan', tmp);
});

ipcMain.handle('artreq:images', async (evt, { eventType, variation, variationIndex, passes }) => {
  const cfg = loadConfig();
  if (!(cfg.comfyUrl || '').trim()) {
    return { success: false, output: 'Set the ComfyUI URL in Settings first.' };
  }
  const { provider } = resolveIdeationProvider(cfg);
  const tmp = tempJsonFile('artreq_images');
  fs.writeFileSync(tmp, JSON.stringify({
    comfy_url: cfg.comfyUrl,
    event_type: eventType,
    doc: variation,
    index: variationIndex,
    provider,
    api_key: cfg.anthropicApiKey,
    ollama_base_url: cfg.ollamaBaseUrl,
    ollama_model: cfg.ollamaModel || 'qwen3:32b',
    ollama_vision_model: cfg.ollamaVisionModel || '',
    passes: passes && passes.length ? passes : ['scene'],
    events_root: ideationEventsRoot()
  }), 'utf-8');
  return runPythonAndScrub('artreq_images', tmp);
});

ipcMain.handle('artreq:regenImage', async (evt, { eventType, variation, variationIndex, zoneNumber, feedback, applyToAll }) => {
  const cfg = loadConfig();
  if (!(cfg.comfyUrl || '').trim()) {
    return { success: false, output: 'Set the ComfyUI URL in Settings first.' };
  }
  const tmp = tempJsonFile('artreq_regen_image');
  fs.writeFileSync(tmp, JSON.stringify({
    comfy_url: cfg.comfyUrl,
    event_type: eventType,
    doc: variation,
    index: variationIndex,
    zone_number: zoneNumber,
    feedback,
    apply_to_all: !!applyToAll,
    author_id: cfg.authorId,
    author_name: cfg.authorName,
    events_root: ideationEventsRoot()
  }), 'utf-8');
  return runPythonAndScrub('artreq_regen_image', tmp);
});

ipcMain.handle('sprint:board', async (evt, { designPath, selectedSprint }) => {
  const rp = roadmapPath();
  if (!fs.existsSync(rp)) {
    return { success: false, output: 'No roadmap found. Export the Feishu roadmap to ' + rp };
  }
  const tmp = tempJsonFile('sprint_board');
  fs.writeFileSync(tmp, JSON.stringify({
    roadmap_path: rp,
    design_path: designPath || '',
    year_hint: new Date().getFullYear(),
    selected_sprint: selectedSprint || null
  }), 'utf-8');
  return parseJsonResult(await runPythonAndScrub('sprint_board', tmp), 'sprint board');
});

// ---- IPC: profile (name + avatar) ----
// Saved into the shared folder, not just local config, so colleagues see
// your avatar next to your feedback. authorId always comes from the OS —
// only the display name and avatar are user-editable.

ipcMain.handle('profile:save', async (evt, { displayName, avatarColor, avatarImage }) => {
  const cfg = loadConfig();
  const name = (displayName || '').trim() || cfg.authorName;
  saveConfig({ authorName: name, avatarColor: avatarColor || '', avatarImage: avatarImage || '' });

  const tmp = tempJsonFile('profile_save');
  fs.writeFileSync(tmp, JSON.stringify({
    events_root: ideationEventsRoot(),
    author_id: cfg.authorId,
    display_name: name,
    avatar_color: avatarColor || '',
    avatar_image: avatarImage || ''
  }), 'utf-8');
  const res = parseJsonResult(await runPythonAndScrub('profile_save', tmp), 'profile');
  // A failed shared-folder write (offline share, bad path) must not lose the
  // local edit — config was already updated above, so report and carry on.
  if (!res.success) return { success: false, output: res.output, savedLocally: true };
  return res;
});

ipcMain.handle('profile:list', async () => {
  const tmp = tempJsonFile('profiles_read');
  fs.writeFileSync(tmp, JSON.stringify({ events_root: ideationEventsRoot() }), 'utf-8');
  return parseJsonResult(await runPythonAndScrub('profiles_read', tmp), 'profiles');
});

// ---- LOCALIZATION ---------------------------------------------------------
// Paste English, get the six shipped languages back. The memory it draws on
// lives beside the ideation data in the shared folder, because it is the same
// kind of thing: knowledge the whole design team accumulates, not per-branch
// config. loc_corrections_import.py populates it from the translator's
// delivered workbooks.
function localizationDataRoot() {
  return path.join(ideationDataRoot(), 'Localization');
}

// Read from WorkSpcae3, NOT from the workspace picked in the context bar.
// WorkSpcae3 is the master that pipeline step 5 clones into the others, so it is
// where the translator's corrections land first — against WorkSpcae_Design, 464
// of the same corrections read as unmerged purely because that clone is behind.
// The memory is reference knowledge, not per-workspace config, so it must not
// change under the user when they switch workspace to do unrelated work.
const MEMORY_WORKSPACE = 'WorkSpcae3';

function localizationDesignPath() {
  const cfg = loadConfig();
  const branch = (cfg.branchPaths || {})[cfg.activeBranch] || '';
  return branch ? resolveDesignPath(branch, MEMORY_WORKSPACE) : '';
}

// ---- Claude Desktop handoff -----------------------------------------------
// The panel cannot make Claude Desktop translate. An MCP server only answers
// the host, and no claude:// deep link carries a prompt (the app's own route
// table is claude://claude, resume, code/new, code/continue, code/needs-input,
// cowork/shared-artifact - none of them take text). So ONE human action is
// unavoidable: send one message over there.
//
// Everything else is automated. Translate queues a job, opens Claude Desktop,
// and starts watching for the answer; the model picks the job up with
// localization_pending and save_translations files the result beside it, which
// is what the renderer is polling for.
function pendingDir() {
  return path.join(localizationDataRoot(), 'pending');
}

ipcMain.handle('loc:queue', async (evt, { strings, languages }) => {
  const cfg = loadConfig();
  const designPath = localizationDesignPath();
  if (!designPath) {
    return { success: false, output: 'Set the branch path in Settings first.' };
  }
  // Second resolution in the id so two clicks in the same second cannot
  // collide - the same bug ideation_store had to fix for feedback entries.
  const jobId = `${new Date().toISOString().replace(/[-:.]/g, '').slice(0, 15)}Z_`
    + `${(cfg.authorName || 'unknown').replace(/[^A-Za-z0-9]+/g, '-')}_`
    + `${Math.random().toString(36).slice(2, 6)}`;
  try {
    fs.mkdirSync(pendingDir(), { recursive: true });
    fs.writeFileSync(path.join(pendingDir(), `${jobId}.json`), JSON.stringify({
      strings, languages, design_path: designPath,
      author_id: cfg.authorId, author_name: cfg.authorName,
      queued_at: new Date().toISOString()
    }, null, 2), 'utf-8');
  } catch (e) {
    return { success: false, output: `Could not queue the batch: ${e.message}` };
  }
  return { success: true, jobId };
});

// Polled by the renderer rather than pushed: the result is written by another
// process into a OneDrive folder, and fs.watch on a synced folder misses and
// double-fires often enough that a 2-second poll is both simpler and steadier.
ipcMain.handle('loc:queueResult', async (evt, { jobId }) => {
  const done = path.join(pendingDir(), `${jobId}.result.json`);
  try {
    if (!fs.existsSync(done)) {
      const stillQueued = fs.existsSync(path.join(pendingDir(), `${jobId}.json`));
      return { success: true, ready: false, queued: stillQueued };
    }
    return { success: true, ready: true, result: JSON.parse(fs.readFileSync(done, 'utf-8')) };
  } catch (e) {
    // A half-synced file reads as invalid JSON for a moment; that is not a
    // failure, it is "not ready yet".
    return { success: true, ready: false, queued: true };
  }
});

// Buffered stdout (the JSON result) but STREAMED stderr (progress lines), so a
// long batch can say where it is. runPythonAndScrub uses execFile, which only
// resolves at the end — fine for a 2-second call, useless for a 5-minute one.
function runPythonStreamed(scriptId, tmpPath, evt, channel) {
  return new Promise((resolve) => {
    let stdout = '';
    let stderr = '';
    let pending = '';
    const proc = spawn(PYTHON_EXE, [PYBACKEND_SCRIPT, scriptId, tmpPath],
      { cwd: PYBACKEND_CWD });

    proc.stdout.on('data', (d) => { stdout += d.toString(); });
    proc.stderr.on('data', (d) => {
      stderr += d.toString();
      // Line-buffer: a chunk boundary can land mid-line and a half-parsed
      // progress line would render as garbage.
      pending += d.toString();
      let nl;
      while ((nl = pending.indexOf('\n')) >= 0) {
        const line = pending.slice(0, nl).trim();
        pending = pending.slice(nl + 1);
        const m = /^PROGRESS (\S+) (\d+)\/(\d+) ?(.*)$/.exec(line);
        if (m && evt && !evt.sender.isDestroyed()) {
          evt.sender.send(channel, {
            stage: m[1], done: Number(m[2]), total: Number(m[3]), detail: m[4]
          });
        }
      }
    });
    proc.on('close', (code) => {
      fs.unlink(tmpPath, () => {});
      resolve({
        success: code === 0, code, stdout, stderr,
        output: stdout + (stderr ? '\n' + stderr : '')
      });
    });
    proc.on('error', (err) => {
      fs.unlink(tmpPath, () => {});
      resolve({ success: false, code: 1, stdout: '', stderr: String(err),
                output: `Failed to launch Python: ${err.message}` });
    });
  });
}

// Recovering a batch the panel stopped waiting for.
//
// This exists because of a real loss: a 124-string batch was queued, answered
// correctly 33 minutes later (all 744 cells, nothing failed), and never
// reached the designer — the panel had given up polling at 10 minutes and had
// no way to pick the result up afterwards. The work was sitting on disk the
// whole time. A finished result must be recoverable, including after the app
// has been closed and reopened.
ipcMain.handle('loc:latestResult', async (evt, { since }) => {
  const dir = pendingDir();
  try {
    if (!fs.existsSync(dir)) return { success: true, found: false };
    const done = fs.readdirSync(dir)
      .filter((n) => n.endsWith('.result.json'))
      .map((n) => ({ name: n, jobId: n.slice(0, -'.result.json'.length),
                     mtime: fs.statSync(path.join(dir, n)).mtimeMs }))
      .sort((a, b) => b.mtime - a.mtime);
    if (!done.length) return { success: true, found: false };
    const newest = done[0];
    if (since && newest.jobId === since) return { success: true, found: false };
    return {
      success: true, found: true, jobId: newest.jobId, mtime: newest.mtime,
      result: JSON.parse(fs.readFileSync(path.join(dir, newest.name), 'utf-8'))
    };
  } catch (e) {
    return { success: false, output: e.message };
  }
});

ipcMain.handle('loc:exportXlsx', async (evt, { rows, languages, suggestedName }) => {
  const picked = await dialog.showSaveDialog({
    title: 'Export translations to Excel',
    defaultPath: path.join(app.getPath('downloads'), suggestedName || 'translations.xlsx'),
    filters: [{ name: 'Excel Workbook', extensions: ['xlsx'] }]
  });
  if (picked.canceled || !picked.filePath) return { success: true, canceled: true };

  const tmp = tempJsonFile('loc_export');
  fs.writeFileSync(tmp, JSON.stringify({
    rows, languages, out_path: picked.filePath
  }), 'utf-8');
  const res = await runPythonAndScrub('loc_export_xlsx', tmp);
  if (res.success) {
    // Reveal rather than open: the designer usually wants to attach or move
    // the file, and opening Excel over the top of the app is in the way.
    shell.showItemInFolder(picked.filePath);
  }
  return res;
});

ipcMain.handle('loc:openClaudeDesktop', async () => {
  // claude://claude just brings the app up. There is no route that can carry
  // the prompt, which is exactly why the user still has to send one message.
  try {
    await shell.openExternal('claude://claude');
    return { success: true };
  } catch (e) {
    return { success: false, output: e.message };
  }
});

// No provider check here, deliberately: the brief is the whole prompt built
// from the memory, with no model involved. It is the route for a designer with
// no API key — they paste it into Claude Desktop and paste the table back.
ipcMain.handle('loc:brief', async (evt, { strings, languages }) => {
  const designPath = localizationDesignPath();
  if (!designPath) {
    return { success: false, output: 'Set the branch path in Settings first.' };
  }
  const tmp = tempJsonFile('loc_brief');
  fs.writeFileSync(tmp, JSON.stringify({
    strings,
    languages,
    design_path: designPath,
    corrections_root: localizationDataRoot(),
    cache_dir: path.join(app.getPath('userData'), 'loc-cache')
  }), 'utf-8');
  return runPythonAndScrub('loc_mcp_brief', tmp);
});

ipcMain.handle('loc:translate', async (evt, { strings, languages }) => {
  const cfg = loadConfig();
  const { provider } = resolveIdeationProvider(cfg);
  const designPath = localizationDesignPath();
  if (!designPath) {
    return { success: false, output: 'Set the branch path in Settings first.' };
  }
  // Missing provider is NOT an error here: the memory tiers need no model, so
  // the run goes ahead and reports what it could answer. Refusing outright
  // would throw away the approved and already-shipped text too, which is the
  // part that works without a key.
  const tmp = tempJsonFile('loc_translate');
  fs.writeFileSync(tmp, JSON.stringify({
    strings,
    languages,
    design_path: designPath,
    corrections_root: localizationDataRoot(),
    // Per-machine, NOT the synced folder: this is a ~6MB derived snapshot of
    // localization.xlsx that rebuilds itself in seconds, so syncing it would
    // cost every designer bandwidth for something none of them shares.
    cache_dir: path.join(app.getPath('userData'), 'loc-cache'),
    provider,
    skip_generation: !provider,
    api_key: cfg.anthropicApiKey,
    ollama_base_url: cfg.ollamaBaseUrl,
    ollama_model: cfg.ollamaModel || 'qwen3:32b'
  }), 'utf-8');
  return runPythonStreamed('loc_translate', tmp, evt, 'loc:progress');
});

// Settings: pick the shared ideationData folder (OneDrive / network share).
// Validated on selection so a typo'd or offline path surfaces immediately
// rather than as a confusing mid-generation failure.
ipcMain.handle('ideation:browseDataFolder', async () => {
  const result = await dialog.showOpenDialog({
    title: 'Select the shared ideationData folder',
    defaultPath: ideationDataRoot(),
    properties: ['openDirectory']
  });
  if (result.canceled || !result.filePaths.length) return null;
  return result.filePaths[0];
});

ipcMain.handle('ideation:dataFolderInfo', async () => {
  const root = ideationDataRoot();
  const events = path.join(root, 'Events');
  return {
    root,
    isDefault: !(loadConfig().ideationDataPath || '').trim(),
    exists: fs.existsSync(events),
    templatesFound: fs.existsSync(path.join(events, 'TEMPLATES', 'Event_Ideation_1Pager.md'))
  };
});

// ---- IPC: manual "Update all local caches from SVN" — convert_layout.lua,
// icon.xlsx/item.xlsx, and localization.xlsx source-event data in one pass ----

ipcMain.handle('caches:refreshAll', async (evt, { branchRoot, workspace, eventNames }) => {
  const problem = branchRootProblem(branchRoot);
  if (problem) return { success: false, code: 1, output: problem };
  const logDir = resolveLogDir(branchRoot);
  fs.mkdirSync(logDir, { recursive: true });

  return new Promise((resolve) => {
    let output = '';
    const proc = spawn(PYTHON_EXE, [
      PYBACKEND_SCRIPT, 'refresh_all_caches', branchRoot, logDir, workspace, JSON.stringify(eventNames)
    ], { cwd: PYBACKEND_CWD });

    proc.stdout.on('data', (d) => { output += d.toString(); evt.sender.send('pipeline:log-line', d.toString()); });
    proc.stderr.on('data', (d) => { output += d.toString(); evt.sender.send('pipeline:log-line', d.toString()); });
    proc.on('close', (code) => resolve({ success: code === 0, code, output }));
    proc.on('error', (err) => resolve({ success: false, code: 1, output: `Failed to launch Python: ${err.message}` }));
  });
});
