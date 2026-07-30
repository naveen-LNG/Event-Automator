const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const fs = require('fs');
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
  '21. Execute coder_convert script'
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
    unityExePath: ''
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
    backgroundColor: '#0e2673',
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
      let output = '';
      const proc = spawn(batPath, [], { cwd: path.dirname(batPath), shell: true });
      proc.stdout.on('data', (d) => {
        output += d.toString();
        evt.sender.send('pipeline:log-line', `[CoderConvert] ${d.toString()}`);
      });
      proc.stderr.on('data', (d) => {
        output += d.toString();
        evt.sender.send('pipeline:log-line', `[CoderConvert Err] ${d.toString()}`);
      });
      proc.on('close', (code) => {
        resolve({ success: code === 0, code, output });
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
