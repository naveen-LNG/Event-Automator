(() => {
  let cfg = null;
  let branchKeys = [];
  let workspaces = [];
  let defaults = {};
  let pipelineSteps = [];
  let stepStatus = []; // 'pending' | 'running' | 'success' | 'error'
  let cachedMinigames = null;
  let cachedBP = null;
  let cachedHBP = null;
  let cachedEventTypes = null;
  let isExecuting = false;

  const $ = (sel) => document.querySelector(sel);
  const $$ = (sel) => Array.from(document.querySelectorAll(sel));

  // ---------------- Console log ----------------
  function log(msg, cls) {
    const view = $('#console-scrollview');
    const line = document.createElement('div');
    const ts = new Date().toLocaleTimeString();
    line.textContent = `[${ts}] ${msg}`;
    line.className = cls || 'console-log-line';
    view.appendChild(line);
    while (view.children.length > 400) view.removeChild(view.firstChild);
    view.scrollTop = view.scrollHeight;
  }
  window.api.onLogLine((line) => log(line, 'console-log-line'));

  // ---------------- Toasts (replaces blocking alert()) ----------------
  function toast(message, kind) {
    const stack = $('#toast-stack');
    const el = document.createElement('div');
    el.className = `toast ${kind || ''}`.trim();
    const title = kind === 'error' ? 'Error' : kind === 'success' ? 'Success' : 'Notice';
    el.innerHTML = `<strong>${title}</strong>${message}`;
    stack.appendChild(el);
    setTimeout(() => { el.style.opacity = '0'; el.style.transition = 'opacity 200ms'; setTimeout(() => el.remove(), 220); }, 4500);
  }

  // ---------------- Navigation ----------------
  let currentView = 'main-events-view';

  function switchView(viewId, title) {
    currentView = viewId;
    $$('.view').forEach(v => v.style.display = 'none');
    $(`#${viewId}`).style.display = 'flex';
    $$('.menu-btn').forEach(b => b.classList.remove('active-btn'));
    if (viewId === 'placeholder-view' && title) {
      $('#placeholder-title').textContent = title;
    }
    if (viewId === 'main-events-view') loadEventTypes();
    if (viewId === 'minigames-view') loadMinigames();
    if (viewId === 'battlepass-view') loadBP();
    if (viewId === 'holiday-bp-view') loadHBP();
  }

  // Re-fetch whichever tab is currently open — called after branch/workspace
  // changes so a visible tab never keeps showing data from the old context.
  async function refreshActiveViewData() {
    if (currentView === 'main-events-view') await loadEventTypes(true);
    else if (currentView === 'minigames-view') await loadMinigames(true);
    else if (currentView === 'battlepass-view') await loadBP(true);
    else if (currentView === 'holiday-bp-view') await loadHBP(true);
  }

  $$('.menu-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      $$('.menu-btn').forEach(b => b.classList.remove('active-btn'));
      btn.classList.add('active-btn');
      switchView(btn.dataset.view, btn.dataset.title);
    });
  });

  // ---------------- Context (branch / workspace) ----------------
  function activeBranchRoot() {
    return (cfg.branchPaths || {})[cfg.activeBranch] || '';
  }

  async function refreshContext() {
    const root = activeBranchRoot();
    const statusEl = $('#context-status');
    const pathEl = $('#context-resolved-path');
    if (!root) {
      pathEl.textContent = `(not set — default guess: ${defaults[cfg.activeBranch]})`;
      statusEl.textContent = 'Branch path not set';
      statusEl.classList.remove('ok');
      return false;
    }
    const design = await window.api.resolveDesignPath(root, cfg.activeWorkspace);
    pathEl.textContent = design;
    const exists = await window.api.validatePath(design);
    if (exists) {
      statusEl.textContent = 'Ready';
      statusEl.classList.add('ok');
    } else {
      statusEl.textContent = 'Path does not exist';
      statusEl.classList.remove('ok');
    }
    $('#execute-btn').disabled = !exists;
    return exists;
  }

  function populateBranchWorkspaceSelects() {
    for (const sel of [$('#active-branch-select'), $('#settings-branch-select')]) {
      sel.innerHTML = '';
      branchKeys.forEach(k => {
        const opt = document.createElement('option');
        opt.value = k; opt.textContent = k;
        sel.appendChild(opt);
      });
      sel.value = cfg.activeBranch;
    }
    for (const sel of [$('#active-workspace-select'), $('#settings-workspace-select')]) {
      sel.innerHTML = '';
      workspaces.forEach(w => {
        const opt = document.createElement('option');
        opt.value = w; opt.textContent = w;
        sel.appendChild(opt);
      });
      sel.value = cfg.activeWorkspace;
    }
  }

  async function setActiveBranch(branch) {
    cfg.activeBranch = branch;
    $('#active-branch-select').value = branch;
    $('#settings-branch-select').value = branch;
    await window.api.saveConfig({ activeBranch: branch });
    // Invalidate sub-tool caches — different branch/workspace means different Excel files.
    cachedMinigames = cachedBP = cachedHBP = cachedEventTypes = null;
    await refreshContext();
    await refreshSettingsResolvedPath();
    await refreshActiveViewData();
  }

  async function setActiveWorkspace(ws) {
    cfg.activeWorkspace = ws;
    $('#active-workspace-select').value = ws;
    $('#settings-workspace-select').value = ws;
    await window.api.saveConfig({ activeWorkspace: ws });
    cachedMinigames = cachedBP = cachedHBP = cachedEventTypes = null;
    buildChecklist();
    await refreshContext();
    await refreshSettingsResolvedPath();
    await refreshActiveViewData();
  }

  $('#active-branch-select').addEventListener('change', (e) => setActiveBranch(e.target.value));
  $('#settings-branch-select').addEventListener('change', (e) => setActiveBranch(e.target.value));
  $('#active-workspace-select').addEventListener('change', (e) => setActiveWorkspace(e.target.value));
  $('#settings-workspace-select').addEventListener('change', (e) => setActiveWorkspace(e.target.value));

  async function refreshSettingsResolvedPath() {
    const root = activeBranchRoot();
    if (!root) { $('#settings-resolved-path').textContent = '(branch path not set)'; return; }
    $('#settings-resolved-path').textContent = await window.api.resolveDesignPath(root, cfg.activeWorkspace);
  }

  // ---------------- Settings: branch paths ----------------
  function renderBranchPathRows() {
    const container = $('#branch-paths-container');
    container.innerHTML = '';
    branchKeys.forEach(key => {
      const row = document.createElement('div');
      row.className = 'branch-path-row';

      const label = document.createElement('div');
      label.className = 'branch-label';
      label.textContent = key;

      const input = document.createElement('input');
      input.type = 'text';
      input.className = 'text-input';
      input.placeholder = defaults[key] || '';
      input.value = (cfg.branchPaths || {})[key] || '';
      input.addEventListener('change', async () => {
        cfg.branchPaths[key] = input.value.trim();
        await window.api.saveConfig({ branchPaths: cfg.branchPaths });
        if (key === cfg.activeBranch) { await refreshContext(); await refreshSettingsResolvedPath(); }
      });

      const browseBtn = document.createElement('button');
      browseBtn.className = 'secondary-btn';
      browseBtn.textContent = 'Browse';
      browseBtn.addEventListener('click', async () => {
        const picked = await window.api.browseFolder(input.value || defaults[key]);
        if (picked) {
          input.value = picked;
          cfg.branchPaths[key] = picked;
          await window.api.saveConfig({ branchPaths: cfg.branchPaths });
          if (key === cfg.activeBranch) { await refreshContext(); await refreshSettingsResolvedPath(); }
        }
      });

      row.appendChild(label);
      row.appendChild(input);
      row.appendChild(browseBtn);
      container.appendChild(row);
    });
  }

  $('#browse-unity-exe-btn').addEventListener('click', async () => {
    const picked = await window.api.browseFile({ startPath: cfg.unityExePath, filters: [{ name: 'Unity Editor', extensions: ['exe'] }] });
    if (picked) {
      $('#unity-exe-input').value = picked;
      cfg.unityExePath = picked;
      await window.api.saveConfig({ unityExePath: picked });
    }
  });

  function getLatestEventNamesPerType() {
    if (!cachedEventTypes) return [];
    return Object.values(cachedEventTypes.events_by_type || {})
      .filter(list => list.length > 0)
      .map(list => list[list.length - 1].name);
  }

  $('#refresh-cache-btn').addEventListener('click', async () => {
    const btn = $('#refresh-cache-btn');
    btn.disabled = true;
    btn.textContent = 'Refreshing...';
    log('Refreshing convert_layout.lua source-event cache from SVN...');

    if (!cachedEventTypes) await loadEventTypes(true);
    const eventNames = getLatestEventNamesPerType();

    const res = await window.api.refreshConvertLayoutCache({
      branchRoot: activeBranchRoot(),
      workspace: cfg.activeWorkspace,
      eventNames
    });

    btn.disabled = false;
    btn.textContent = 'Update Cache from SVN';

    if (res.success) {
      toast('Cache refresh complete — see the Main Events console for what changed.', 'success');
    } else {
      log(`Cache refresh failed: ${res.output}`, 'console-log-error');
      toast('Cache refresh failed — check console.', 'error');
    }
  });

  // ---------------- Event type / source event picker ----------------
  function clearEventTypeSelects(message) {
    cachedEventTypes = null;
    const typeSel = $('#event-type-select');
    const srcSel = $('#source-event-select');
    typeSel.innerHTML = '';
    srcSel.innerHTML = '';
    const opt = document.createElement('option');
    opt.value = ''; opt.textContent = message;
    srcSel.appendChild(opt);
  }

  async function loadEventTypes(force) {
    if (cachedEventTypes && !force) { populateSourceEventOptions(); return; }
    const root = activeBranchRoot();
    if (!root) { clearEventTypeSelects('(branch path not set)'); return; }
    const designPath = await window.api.resolveDesignPath(root, cfg.activeWorkspace);
    const res = await window.api.eventTypesGet({ designPath });
    if (!res.success) {
      log(`Error loading event history: ${res.output}`, 'console-log-error');
      clearEventTypeSelects('(events.xlsx not found for this branch/workspace)');
      toast('Could not load event history for this branch/workspace — check the console for details.', 'error');
      return;
    }
    try {
      const data = JSON.parse(res.stdout);
      if (data.error) {
        log(`Error: ${data.error}`, 'console-log-error');
        clearEventTypeSelects('(events.xlsx not found for this branch/workspace)');
        toast('Could not load event history for this branch/workspace — check the console for details.', 'error');
        return;
      }
      cachedEventTypes = data;
      populateEventTypeSelect();
    } catch (e) {
      log(`Error parsing event history JSON: ${e.message}`, 'console-log-error');
      clearEventTypeSelects('(failed to load event history)');
    }
  }

  function populateEventTypeSelect() {
    const sel = $('#event-type-select');
    const prev = sel.value;
    sel.innerHTML = '';
    (cachedEventTypes.types || []).forEach(t => {
      const opt = document.createElement('option');
      opt.value = t; opt.textContent = t;
      sel.appendChild(opt);
    });
    sel.value = (prev && cachedEventTypes.types.includes(prev)) ? prev : (cachedEventTypes.types[0] || '');
    populateSourceEventOptions();
  }

  function populateSourceEventOptions() {
    if (!cachedEventTypes) return;
    const sel = $('#source-event-select');
    const type = $('#event-type-select').value;
    const list = (cachedEventTypes.events_by_type || {})[type] || [];
    sel.innerHTML = '';
    if (list.length === 0) {
      const opt = document.createElement('option');
      opt.value = ''; opt.textContent = '(no previous events found for this type)';
      sel.appendChild(opt);
      return;
    }
    list.forEach(ev => {
      const opt = document.createElement('option');
      opt.value = ev.name;
      opt.dataset.startTime = ev.start_time || '';
      opt.textContent = `${ev.alias} — ${ev.start_time || 'no date'}`;
      sel.appendChild(opt);
    });
    // Latest event is the last one in chronological order — pick it by default.
    sel.selectedIndex = list.length - 1;
  }

  $('#event-type-select').addEventListener('change', populateSourceEventOptions);

  // ---------------- Start Date + Duration -> Start/End/Near-End/Close times ----------------
  function pad2(n) { return String(n).padStart(2, '0'); }
  function fmtDate(utcMs) {
    const d = new Date(utcMs);
    return `${d.getUTCFullYear()}-${pad2(d.getUTCMonth() + 1)}-${pad2(d.getUTCDate())}`;
  }

  // Verified against events.xlsx (trunk/WorkSpcae3/design/events.xlsx, "events" sheet, cols H-K):
  // end date = start date + (duration - 1) days, at 23:59:59; near-end = same end date at 00:00:00;
  // close time = identical to end time. Checked against a 14-day case (Blocks2601: 2026-01-08 -> 2026-01-21)
  // and a 21-day case (Merge2512: 2025-12-18 -> 2026-01-07); both matched.
  function computeEventTimes(startDateStr, durationDays) {
    if (!startDateStr || !durationDays || durationDays < 1) return null;
    const [y, m, d] = startDateStr.split('-').map(Number);
    const startMs = Date.UTC(y, m - 1, d);
    const endMs = startMs + (durationDays - 1) * 86400000;
    const endDateStr = fmtDate(endMs);
    return {
      startTime: `${startDateStr} 00:00:00`,
      endTime: `${endDateStr} 23:59:59`,
      nearEndTime: `${endDateStr} 00:00:00`,
      closeTime: `${endDateStr} 23:59:59`
    };
  }

  function updateTimesPreview() {
    const startDateStr = $('#start-date-input').value;
    const duration = parseInt($('#duration-days-input').value, 10);
    const times = computeEventTimes(startDateStr, duration);
    if (!times) {
      ['#preview-start-time', '#preview-end-time', '#preview-near-end-time', '#preview-close-time']
        .forEach(sel => $(sel).textContent = '...');
      return;
    }
    $('#preview-start-time').textContent = times.startTime;
    $('#preview-end-time').textContent = times.endTime;
    $('#preview-near-end-time').textContent = times.nearEndTime;
    $('#preview-close-time').textContent = times.closeTime;
  }
  $('#start-date-input').addEventListener('input', updateTimesPreview);
  $('#duration-days-input').addEventListener('input', updateTimesPreview);

  // ---------------- Checklist ----------------

  // Step index 4 ("Clone WorkSpace3 configuration to active workspace") deletes and
  // re-clones the active workspace FROM WorkSpcae3 — if WorkSpcae3 is itself the active
  // workspace, source and destination are the same path (the Python backend now refuses
  // to run in that case). Rather than show a step that will only ever fail, hide it and
  // renumber the rest so the checklist never has a confusing gap or wrong step count.
  const WORKSPACE3_ONLY_STEP_INDEX = 4;

  function visibleStepIndices() {
    const indices = pipelineSteps.map((_, i) => i);
    if (cfg.activeWorkspace === 'WorkSpcae3') {
      return indices.filter(i => i !== WORKSPACE3_ONLY_STEP_INDEX);
    }
    return indices;
  }

  function displayNumberFor(trueIdx) {
    return visibleStepIndices().indexOf(trueIdx) + 1;
  }

  function buildChecklist() {
    stepStatus = pipelineSteps.map(() => 'pending');
    const container = $('#checklist-container');
    container.innerHTML = '';
    const visible = visibleStepIndices();

    visible.forEach((trueIdx, pos) => {
      const displayLabel = pipelineSteps[trueIdx].replace(/^\d+\.\s*/, `${pos + 1}. `);

      const item = document.createElement('div');
      item.className = 'checklist-item';
      item.dataset.idx = trueIdx;

      const icon = document.createElement('div');
      icon.className = 'status-icon status-pending';

      const text = document.createElement('div');
      text.className = 'checklist-item-text';
      text.textContent = displayLabel;

      const runBtn = document.createElement('button');
      runBtn.className = 'checklist-item-btn';
      runBtn.textContent = 'Run';
      runBtn.addEventListener('click', () => runSingleStep(trueIdx, false));

      const undoBtn = document.createElement('button');
      undoBtn.className = 'checklist-item-undo-btn';
      undoBtn.textContent = 'Undo';
      undoBtn.addEventListener('click', () => runSingleStep(trueIdx, true));

      item.appendChild(icon);
      item.appendChild(text);
      item.appendChild(runBtn);
      item.appendChild(undoBtn);
      container.appendChild(item);
    });

    if (cfg.activeWorkspace === 'WorkSpcae3') {
      const note = document.createElement('div');
      note.className = 'settings-hint';
      note.style.padding = '8px 4px 0';
      note.textContent = 'Workspace clone step is hidden — the active workspace is WorkSpcae3 itself, so there is nothing to clone from.';
      container.appendChild(note);
    }
  }

  function setStepStatus(idx, status) {
    stepStatus[idx] = status;
    const item = $(`.checklist-item[data-idx="${idx}"]`);
    if (!item) return;
    const icon = item.querySelector('.status-icon');
    icon.className = `status-icon status-${status}`;
  }

  function gatherEventNames() {
    const src = $('#source-event-select').value;
    const tgt = $('#target-event-input').value.trim();
    return { src, tgt };
  }

  async function runSingleStep(idx, isRevert) {
    const { src, tgt } = gatherEventNames();
    if (!src) { log('Error: No Source Event selected — pick an Event Type with at least one previous event.', 'console-log-error'); return false; }
    if (!tgt) { log('Error: Target Event Name cannot be empty.', 'console-log-error'); return false; }

    const startDateStr = $('#start-date-input').value;
    const duration = parseInt($('#duration-days-input').value, 10);
    const times = computeEventTimes(startDateStr, duration);
    if (!times) { log('Error: Start Date and Duration (days) are required.', 'console-log-error'); return false; }

    const ready = await refreshContext();
    if (!ready) { log('Error: Active branch/workspace design path does not exist. Check Settings.', 'console-log-error'); return false; }

    const displayNum = displayNumberFor(idx);
    log(`${isRevert ? 'Reverting' : 'Executing'} Step ${displayNum}: ${pipelineSteps[idx].replace(/^\d+\.\s*/, '')}...`);
    setStepStatus(idx, 'running');

    const params = {
      stepIndex: idx,
      srcEvt: src,
      tgtEvt: tgt,
      isRevert,
      branchRoot: activeBranchRoot(),
      workspace: cfg.activeWorkspace,
      unityExePath: cfg.unityExePath || '',
      eventId: $('#event-id-input').value.trim(),
      startTime: times.startTime,
      endTime: times.endTime,
      nearEndTime: times.nearEndTime,
      closeTime: times.closeTime,
      isReopen: false,
      srcStartTime: $('#source-event-select').selectedOptions[0]?.dataset.startTime || ''
    };

    const result = await window.api.runStep(params);
    // Every step except the Unity headless call (idx === 1) streams its output live via
    // pipeline:log-line as it runs — only dump the buffered result here for the one that doesn't,
    // otherwise everything would print twice (once streamed, once as a final blob).
    if (idx === 1 && result.output) log(result.output, result.success ? 'console-log-line' : 'console-log-error');

    if (result.success) {
      setStepStatus(idx, isRevert ? 'pending' : 'success');
      log(`Step ${displayNum} ${isRevert ? 'Reverted' : 'Succeeded'}.`, 'console-log-success');
      return true;
    } else {
      setStepStatus(idx, 'error');
      log(`Step ${displayNum} Failed!`, 'console-log-error');
      return false;
    }
  }

  $('#execute-btn').addEventListener('click', async () => {
    if (isExecuting) return;
    isExecuting = true;
    $('#execute-btn').disabled = true;
    $('#undo-all-btn').disabled = true;
    log('Execution started...');
    for (const idx of visibleStepIndices()) {
      const ok = await runSingleStep(idx, false);
      if (!ok) { log('Pipeline aborted due to step failure.', 'console-log-error'); break; }
    }
    isExecuting = false;
    $('#execute-btn').disabled = false;
    $('#undo-all-btn').disabled = false;
    log('Execution process finished.');
  });

  $('#undo-all-btn').addEventListener('click', async () => {
    if (isExecuting) return;
    isExecuting = true;
    $('#execute-btn').disabled = true;
    $('#undo-all-btn').disabled = true;
    log('Revert pipeline started...');
    const steps = visibleStepIndices();
    for (let i = steps.length - 1; i >= 0; i--) {
      await runSingleStep(steps[i], true);
    }
    isExecuting = false;
    $('#execute-btn').disabled = false;
    $('#undo-all-btn').disabled = false;
    log('Revert process finished.');
  });

  // ---------------- Minigames tab ----------------
  async function loadMinigames(force) {
    if (cachedMinigames && !force) return;
    cachedMinigames = null;
    $('#minigame-rows-container').innerHTML = '';
    const root = activeBranchRoot();
    if (!root) { log('Cannot load minigames: branch path not set.', 'console-log-error'); return; }
    const designPath = await window.api.resolveDesignPath(root, cfg.activeWorkspace);
    log('Loading Minigame configs from Excel...');
    const res = await window.api.minigamesGet({ designPath });
    if (!res.success) { log(`Error loading minigames: ${res.output}`, 'console-log-error'); toast('Could not load mini_mgr.xlsx for this branch/workspace.', 'error'); return; }
    try {
      const data = JSON.parse(res.stdout);
      if (data.error) { log(`Error: ${data.error}`, 'console-log-error'); toast('Could not load mini_mgr.xlsx for this branch/workspace.', 'error'); return; }
      cachedMinigames = data;
      log(`Loaded ${data.minigames.length} minigames. Max ID: ${data.max_id}`);
      addMinigameRow();
    } catch (e) {
      log(`Error parsing minigame JSON: ${e.message}`, 'console-log-error');
    }
  }

  function addMinigameRow() {
    if (!cachedMinigames) { log('Minigame configs still loading...', 'console-log-error'); return; }
    const container = $('#minigame-rows-container');
    const nextId = cachedMinigames.max_id + container.children.length + 1;

    const row = document.createElement('div');
    row.className = 'minigame-row';

    const idInput = document.createElement('input');
    idInput.className = 'mg-id'; idInput.value = nextId;

    const nameSelect = document.createElement('select');
    nameSelect.className = 'mg-name';
    cachedMinigames.minigames.forEach(m => {
      const opt = document.createElement('option');
      opt.value = m.name; opt.textContent = m.name;
      nameSelect.appendChild(opt);
    });

    const startInput = document.createElement('input');
    startInput.className = 'mg-time'; startInput.value = '2026-06-01 00:00:00';
    const endInput = document.createElement('input');
    endInput.className = 'mg-time'; endInput.value = '2026-06-03 23:59:59';

    const dwInput = document.createElement('input');
    dwInput.className = 'mg-time'; dwInput.placeholder = 'Double Wk ID'; dwInput.style.display = 'none';

    function syncDw() {
      const mg = cachedMinigames.minigames.find(m => m.name === nameSelect.value);
      dwInput.style.display = (mg && mg.requires_double_week) ? '' : 'none';
    }
    nameSelect.addEventListener('change', syncDw);
    syncDw();

    const removeBtn = document.createElement('button');
    removeBtn.className = 'mg-remove-btn';
    removeBtn.textContent = '✕';
    removeBtn.addEventListener('click', () => row.remove());

    row.appendChild(idInput);
    row.appendChild(nameSelect);
    row.appendChild(startInput);
    row.appendChild(endInput);
    row.appendChild(dwInput);
    row.appendChild(removeBtn);
    container.appendChild(row);
  }

  $('#add-minigame-row-btn').addEventListener('click', addMinigameRow);

  $('#minigame-execute-btn').addEventListener('click', async () => {
    const rows = [];
    $$('#minigame-rows-container .minigame-row').forEach(row => {
      const [idInput, nameSelect, startInput, endInput, dwInput] = row.children;
      const id = parseInt(idInput.value, 10);
      if (Number.isNaN(id)) return;
      const mg = cachedMinigames.minigames.find(m => m.name === nameSelect.value);
      rows.push({
        id,
        minigame: nameSelect.value,
        start_time: startInput.value,
        end_time: endInput.value,
        double_week_id: dwInput.style.display === 'none' ? '' : dwInput.value,
        discount_id: mg ? mg.discount_id : -1
      });
    });
    if (rows.length === 0) { log('No valid minigame rows to commit.', 'console-log-error'); return; }

    const root = activeBranchRoot();
    const designPath = await window.api.resolveDesignPath(root, cfg.activeWorkspace);
    const btn = $('#minigame-execute-btn');
    btn.disabled = true; btn.textContent = 'Committing...';
    const res = await window.api.minigamesCommit({ designPath, payload: { rows } });
    btn.disabled = false; btn.textContent = 'Commit to mini_mgr.xlsx';

    if (res.success) {
      log(`Success: ${res.stdout}`, 'console-log-success');
      toast(`${rows.length} minigame run(s) appended to mini_mgr.xlsx.`, 'success');
      $('#minigame-rows-container').innerHTML = '';
      cachedMinigames = null;
    } else {
      log(`Commit failed: ${res.output}`, 'console-log-error');
      toast('Commit failed. Check that mini_mgr.xlsx is closed in Excel.<br>' + res.output, 'error');
    }
  });

  // ---------------- BattlePass tab ----------------
  async function loadBP(force) {
    if (cachedBP && !force) { updateBPPreview(); return; }
    cachedBP = null;
    const root = activeBranchRoot();
    if (!root) { log('Cannot load BattlePass: branch path not set.', 'console-log-error'); return; }
    const designPath = await window.api.resolveDesignPath(root, cfg.activeWorkspace);
    log('Loading BattlePass metadata...');
    const res = await window.api.bpGet({ designPath });
    if (!res.success) { log(`Error loading BP info: ${res.output}`, 'console-log-error'); toast('Could not load battle_pass.xlsx for this branch/workspace.', 'error'); return; }
    try {
      cachedBP = JSON.parse(res.stdout);
      if (cachedBP.error) { log(`Error: ${cachedBP.error}`, 'console-log-error'); toast('Could not load battle_pass.xlsx for this branch/workspace.', 'error'); cachedBP = null; return; }
      log(`BP metadata loaded. Last Cycle: ${cachedBP.last_cycle_id}`);
      updateBPPreview();
    } catch (e) {
      log(`Error parsing BP JSON: ${e.message}`, 'console-log-error');
    }
  }

  function nextSequentialYYMM(lastAlias) {
    const m = /(\d{4})/.exec(lastAlias || '');
    if (!m) return '2601';
    let yy = parseInt(m[1].substring(0, 2), 10);
    let mm = parseInt(m[1].substring(2, 4), 10);
    mm++;
    if (mm > 12) { mm = 1; yy++; }
    return `${String(yy).padStart(2, '0')}${String(mm).padStart(2, '0')}`;
  }

  function updateBPPreview() {
    if (!cachedBP) return;
    $('#bp-prev-cycle').textContent = cachedBP.last_cycle_id + 1;
    $('#bp-prev-plan').textContent = cachedBP.last_scheme_id + 1;
    $('#bp-prev-item').textContent = cachedBP.last_item_id + 1;
    $('#bp-prev-yymm').textContent = nextSequentialYYMM(cachedBP.last_recharge_alias);

    const start = new Date($('#bp-start-time').value);
    const end = new Date($('#bp-end-time').value);
    if (!isNaN(start) && !isNaN(end)) {
      const days = Math.ceil((end - start) / 86400000);
      const mode = days >= 30 ? 900 : 800;
      const template = mode === 900 ? cachedBP.last_900_cycle : cachedBP.last_800_cycle;
      $('#bp-point-mode').textContent = `${mode} Points (${days} days)`;
      $('#bp-prev-template').textContent = `BP ${template}`;
    } else {
      $('#bp-point-mode').textContent = 'Enter valid dates...';
      $('#bp-prev-template').textContent = '...';
    }
  }
  $('#bp-start-time').addEventListener('input', updateBPPreview);
  $('#bp-end-time').addEventListener('input', updateBPPreview);

  $('#bp-execute-btn').addEventListener('click', async () => {
    if (!cachedBP) return;
    const start = new Date($('#bp-start-time').value);
    const end = new Date($('#bp-end-time').value);
    if (isNaN(start) || isNaN(end)) { toast('Invalid date format. Please use yyyy-mm-dd hh:mm:ss', 'error'); return; }

    const days = Math.ceil((end - start) / 86400000);
    const mode = days >= 30 ? 900 : 800;
    const template = mode === 900 ? cachedBP.last_900_cycle : cachedBP.last_800_cycle;

    const payload = {
      cycle_id: cachedBP.last_cycle_id + 1,
      start_time: $('#bp-start-time').value,
      end_time: $('#bp-end-time').value,
      bg_limit: mode,
      template_cycle: template
    };

    const root = activeBranchRoot();
    const designPath = await window.api.resolveDesignPath(root, cfg.activeWorkspace);
    const btn = $('#bp-execute-btn');
    btn.disabled = true; btn.textContent = 'Committing...';
    const res = await window.api.bpCommit({ designPath, payload });
    btn.disabled = false; btn.textContent = 'Commit BattlePass Config';

    if (res.success) {
      toast('BattlePass configuration updated across all sheets!', 'success');
      cachedBP = null;
      loadBP();
    } else {
      toast('Commit failed: ' + res.output, 'error');
    }
  });

  // ---------------- Holiday BP tab ----------------
  async function loadHBP(force) {
    if (cachedHBP && !force) { updateHBPPreview(); return; }
    cachedHBP = null;
    const root = activeBranchRoot();
    if (!root) { log('Cannot load Holiday BP: branch path not set.', 'console-log-error'); return; }
    const designPath = await window.api.resolveDesignPath(root, cfg.activeWorkspace);
    log('Loading Holiday BattlePass metadata...');
    const res = await window.api.hbpGet({ designPath });
    if (!res.success) { log(`Error fetching HBP metadata: ${res.output}`, 'console-log-error'); toast('Could not load festivals_bp.xlsx for this branch/workspace.', 'error'); return; }
    try {
      cachedHBP = JSON.parse(res.stdout);
      log('Holiday BP metadata loaded.');
      updateHBPPreview();
    } catch (e) {
      log(`Error parsing HBP metadata: ${e.message}`, 'console-log-error');
    }
  }

  function updateHBPPreview() {
    if (!cachedHBP) return;
    $('#hbp-next-cycle').textContent = cachedHBP.last_cycle_id + 1;
    $('#hbp-next-recommend').textContent = cachedHBP.last_recommend_id + 1;
    $('#hbp-next-item').textContent = cachedHBP.last_item_id + 1;
    $('#hbp-next-scheme').textContent = cachedHBP.last_scheme_id + 1;
    const hid = $('#hbp-holiday-id').value.trim() || '[ID]';
    $('#hbp-suffix-preview').textContent = (cachedHBP.last_suffixes || []).map(s => hid + s).join(' ');
  }
  $('#hbp-holiday-id').addEventListener('input', updateHBPPreview);

  $('#hbp-execute-btn').addEventListener('click', async () => {
    const holidayId = $('#hbp-holiday-id').value.trim();
    const start = $('#hbp-start-time').value.trim();
    const end = $('#hbp-end-time').value.trim();
    if (!holidayId || !start || !end) { toast('Please fill in Holiday ID, Start Time and End Time.', 'error'); return; }

    const root = activeBranchRoot();
    const designPath = await window.api.resolveDesignPath(root, cfg.activeWorkspace);
    const btn = $('#hbp-execute-btn');
    btn.disabled = true;
    const res = await window.api.hbpCommit({
      designPath,
      payload: { holiday_id: holidayId, start_time: start, end_time: end }
    });
    btn.disabled = false;

    if (res.success) {
      log('Holiday BP Commit Successful!', 'console-log-success');
      toast('Holiday BattlePass configuration updated across all files.', 'success');
      cachedHBP = null;
      loadHBP();
    } else {
      log(`Holiday BP Commit Failed: ${res.output}`, 'console-log-error');
      toast('Commit failed: ' + res.output, 'error');
    }
  });

  // ---------------- Init ----------------
  async function init() {
    const data = await window.api.getConfig();
    cfg = data.cfg;
    branchKeys = data.branchKeys;
    workspaces = data.workspaces;
    defaults = data.defaults;
    pipelineSteps = data.pipelineSteps;

    $('#unity-exe-input').value = cfg.unityExePath || '';

    populateBranchWorkspaceSelects();
    renderBranchPathRows();
    buildChecklist();
    await refreshContext();
    await refreshSettingsResolvedPath();
    await loadEventTypes();
    updateTimesPreview();

    log('Tool initialized. Waiting for input...');
  }

  init();
})();
