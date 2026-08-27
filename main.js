const { app, BrowserWindow, ipcMain, dialog } = require('electron');
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
  '2. Execute ProtoToLua in Unity',
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
  '23. Copy & rename event code folders (Core<Event> + UI<Event>Panel)'
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
    unityExePath: '',
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

function resolveLogDir(branchRoot) {
  // Must match the path event_pipeline.py itself derives from proj_root,
  // so undo_log.json / extra_args.json land in the same place the script reads them from.
  return path.join(branchRoot, 'develop', 'client', 'Skipbo', 'Assets', 'Editor', 'EventAutomation', 'PythonBackend');
}

// ---- Window ----

function createWindow() {
  const win = new BrowserWindow({
    width: 1300,
    height: 860,
    minWidth: 480,
    minHeight: 600,
    title: 'Event Automator (Standalone)',
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
  win.loadFile(path.join(__dirname, 'renderer', 'index.html'));
}

app.whenReady().then(() => {
  createWindow();
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
    branchRoot, workspace, unityExePath,
    eventId, startTime, endTime, nearEndTime, closeTime, isReopen, srcStartTime
  } = params;

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

  // Step index 1: ProtoToLua — the one genuinely Unity-coupled step.
  // Run Unity headlessly instead of requiring the editor to be open.
  if (stepIndex === 1) {
    if (isRevert) {
      return { success: true, output: 'Reverting ProtoToLua is not strictly trackable. Skipping.' };
    }
    if (!unityExePath || !fs.existsSync(unityExePath)) {
      return { success: false, output: 'Unity executable path is not configured (Settings > Unity Executable Path). Cannot run ProtoToLua headlessly.' };
    }
    const projectPath = resolveUnityProjectPath(branchRoot);
    if (!fs.existsSync(projectPath)) {
      return { success: false, output: `Unity project not found at ${projectPath}` };
    }
    return new Promise((resolve) => {
      const args = ['-batchmode', '-quit', '-nographics', '-projectPath', projectPath, '-executeMethod', 'ProtoTool.ProtoToLua', '-logFile', '-'];
      execFile(unityExePath, args, { maxBuffer: 20 * 1024 * 1024, timeout: 10 * 60 * 1000 }, (error, stdout, stderr) => {
        resolve({
          success: !error,
          output: `Launched Unity headless: ${unityExePath} ${args.join(' ')}\n` + (stdout || '') + (stderr ? '\n' + stderr : '')
        });
      });
    });
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
function allBranchDesignDirs(workspace) {
  const cfg = loadConfig();
  return BRANCH_KEYS
    .map((key) => cfg.branchPaths[key])
    .filter(Boolean)
    .map((root) => resolveDesignPath(root, workspace));
}

ipcMain.handle('trivia:get', async (evt, { branchRoot, workspace }) => {
  return runPython('get_trivia_challenge_info', [branchRoot, workspace, JSON.stringify(allBranchDesignDirs(workspace))]);
});

ipcMain.handle('trivia:run', async (evt, { branchRoot, workspace }) => {
  const logDir = resolveLogDir(branchRoot);
  fs.mkdirSync(logDir, { recursive: true });

  return new Promise((resolve) => {
    let output = '';
    const proc = spawn(PYTHON_EXE, [
      PYBACKEND_SCRIPT, 'trivia_challenge_cycle', branchRoot, workspace, 'run', JSON.stringify(allBranchDesignDirs(workspace))
    ], { cwd: PYBACKEND_CWD });

    proc.stdout.on('data', (d) => { output += d.toString(); evt.sender.send('pipeline:log-line', d.toString()); });
    proc.stderr.on('data', (d) => { output += d.toString(); evt.sender.send('pipeline:log-line', d.toString()); });
    proc.on('close', (code) => resolve({ success: code === 0, code, output }));
    proc.on('error', (err) => resolve({ success: false, code: 1, output: `Failed to launch Python: ${err.message}` }));
  });
});

// ---- IPC: Chapter Update (new PVE chapter roll-forward) ----

ipcMain.handle('chapter:get', async (evt, { designPath }) => {
  return runPython('get_chapter_info', [designPath]);
});

// Streamed rather than buffered: copying dynamic_map_NNNNN can take a while and
// a silent progress pane reads as a hang.
ipcMain.handle('chapter:run', async (evt, { branchRoot, payload }) => {
  const tmp = tempJsonFile('chapter_payload');
  fs.writeFileSync(tmp, JSON.stringify(payload), 'utf-8');

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

ipcMain.handle('artreq:generate', async (evt, { eventType, ideationTheme, variationCount }) => {
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
