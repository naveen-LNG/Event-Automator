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
  let cachedFSG = null;
  let cachedTrivia = null;
  let cachedEndlessTreasure = null;
  let ideationThemes = null;
  let ideationDraftPaths = [];
  let ideationSelectedIndex = -1;
  let ideationSavedPath = '';
  let artReqVariations = null;
  let artReqDraftPaths = [];
  let artReqSelectedIndex = -1;
  let artReqSavedPath = '';
  // Shared across designers via the ideationData folder — refreshed after
  // every generate/regenerate/retract so a colleague's note shows up here.
  let ideationEntries = [];
  let ideationLessons = [];
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
  let currentView = 'sprint-view';

  // "Regular Updates" collapsible group. The count is read from the DOM
  // rather than hardcoded, so adding a destination to the group in the
  // markup can't leave a stale number behind.
  const navGroup = $('#nav-group-regular');
  const navGroupItems = $('#nav-group-regular-items');
  const navGroupHeader = navGroup && navGroup.querySelector('.nav-group-header');

  if (navGroup && navGroupHeader) {
    $('#nav-group-regular-count').textContent =
      navGroupItems.querySelectorAll('.menu-btn').length;

    navGroupHeader.addEventListener('click', () => {
      const collapsed = navGroup.classList.toggle('collapsed');
      navGroupHeader.setAttribute('aria-expanded', String(!collapsed));
    });
  }

  // switchView only swaps the panes; it deliberately clears every active-btn
  // and leaves re-marking to the click handler. Anything navigating
  // programmatically (startup, the sprint board's Open buttons) needs this
  // wrapper, or the sidebar ends up with nothing highlighted.
  function navigateTo(viewId) {
    switchView(viewId);
    const btn = document.querySelector(`.menu-btn[data-view="${viewId}"]`);
    if (btn) btn.classList.add('active-btn');
  }

  // Never let the active destination hide inside a collapsed group.
  function revealGroupFor(viewId) {
    if (!navGroup || !navGroupItems) return;
    const inGroup = !!navGroupItems.querySelector(`.menu-btn[data-view="${viewId}"]`);
    if (inGroup && navGroup.classList.contains('collapsed')) {
      navGroup.classList.remove('collapsed');
      navGroupHeader.setAttribute('aria-expanded', 'true');
    }
  }

  function switchView(viewId, title) {
    currentView = viewId;
    revealGroupFor(viewId);
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
    if (viewId === 'decoration-shop-view') updateDecoShopPreview();
    if (viewId === 'streak-challenge-view') updateStreakPreview();
    if (viewId === 'weekly-gift-view') updateWeeklyGiftPreview();
    if (viewId === 'full-screen-gift-view') loadFullScreenGift();
    if (viewId === 'trivia-challenge-view') loadTrivia();
    if (viewId === 'endless-treasure-view') loadEndlessTreasure();
    if (viewId === 'profile-view') refreshProfiles();
    if (viewId === 'sprint-view' && !sprintData) loadSprintBoard();
  }

  // Re-fetch whichever tab is currently open — called after branch/workspace
  // changes so a visible tab never keeps showing data from the old context.
  async function refreshActiveViewData() {
    if (currentView === 'main-events-view') await loadEventTypes(true);
    else if (currentView === 'minigames-view') await loadMinigames(true);
    else if (currentView === 'battlepass-view') await loadBP(true);
    else if (currentView === 'holiday-bp-view') await loadHBP(true);
    else if (currentView === 'full-screen-gift-view') await loadFullScreenGift(true);
    else if (currentView === 'trivia-challenge-view') await loadTrivia(true);
    else if (currentView === 'endless-treasure-view') await loadEndlessTreasure(true);
  }

  $$('.menu-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      // switchView() clears every active-btn itself, so the highlight has to
      // be applied AFTER it — marking first meant switchView immediately
      // stripped it and every click left the sidebar with nothing lit.
      switchView(btn.dataset.view, btn.dataset.title);
      btn.classList.add('active-btn');
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
    cachedMinigames = cachedBP = cachedHBP = cachedEventTypes = cachedFSG = null;
    await refreshContext();
    await refreshSettingsResolvedPath();
    await refreshActiveViewData();
  }

  async function setActiveWorkspace(ws) {
    cfg.activeWorkspace = ws;
    $('#active-workspace-select').value = ws;
    $('#settings-workspace-select').value = ws;
    await window.api.saveConfig({ activeWorkspace: ws });
    cachedMinigames = cachedBP = cachedHBP = cachedEventTypes = cachedFSG = null;
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

  // ---------------- Profile & avatars ----------------
  const AVATAR_COLORS = ['#7c5cfc', '#34d399', '#fbbf24', '#fb7185', '#38bdf8', '#f472b6', '#2dd4bf', '#fb923c'];
  let teamProfiles = {};       // author_id -> profile, from the shared folder
  let draftAvatarColor = '';
  let draftAvatarImage = '';

  function initialsOf(name) {
    const parts = String(name || '').trim().split(/\s+/).filter(Boolean);
    if (!parts.length) return '?';
    if (parts.length === 1) return parts[0].slice(0, 2).toUpperCase();
    return (parts[0][0] + parts[parts.length - 1][0]).toUpperCase();
  }

  // Deterministic fallback so a colleague who never set a colour still gets a
  // stable one rather than changing hue between renders.
  function autoColor(seed) {
    let h = 0;
    for (const ch of String(seed || '')) h = (h * 31 + ch.charCodeAt(0)) >>> 0;
    return AVATAR_COLORS[h % AVATAR_COLORS.length];
  }

  function avatarHtml(authorId, displayName, size) {
    const p = teamProfiles[authorId] || {};
    const name = p.display_name || displayName || authorId || '?';
    const img = p.avatar_image || (authorId === cfg.authorId ? draftAvatarImage : '');
    const color = p.avatar_color || (authorId === cfg.authorId ? draftAvatarColor : '') || autoColor(authorId || name);
    const px = size || 26;
    const style = `width:${px}px;height:${px}px;font-size:${Math.round(px * 0.4)}px;`;
    if (img) {
      return `<span class="avatar" style="${style}background-image:url('${escapeHtml(img)}');background-size:cover;background-position:center;" title="${escapeHtml(name)}"></span>`;
    }
    return `<span class="avatar" style="${style}background:${escapeHtml(color)};" title="${escapeHtml(name)}">${escapeHtml(initialsOf(name))}</span>`;
  }

  function renderProfileEditor() {
    const name = $('#author-name-input').value || cfg.authorName || '';
    const color = draftAvatarColor || autoColor(cfg.authorId);
    const preview = $('#profile-avatar-preview');
    if (draftAvatarImage) {
      preview.style.backgroundImage = `url('${draftAvatarImage}')`;
      preview.style.backgroundSize = 'cover';
      preview.style.backgroundPosition = 'center';
      preview.style.background = `url('${draftAvatarImage}') center/cover`;
      preview.textContent = '';
    } else {
      preview.style.backgroundImage = '';
      preview.style.background = color;
      preview.textContent = initialsOf(name);
    }
    $('#sidebar-avatar').innerHTML = avatarHtml(cfg.authorId, cfg.authorName, 22);
    $('#sidebar-profile-name').textContent = cfg.authorName || 'Profile';

    $('#avatar-swatches').innerHTML = AVATAR_COLORS.map(c =>
      `<button type="button" class="avatar-swatch${c === color ? ' selected' : ''}" data-color="${c}" style="background:${c};" title="${c}"></button>`
    ).join('');
    $$('.avatar-swatch').forEach(b => b.addEventListener('click', () => {
      draftAvatarColor = b.dataset.color;
      draftAvatarImage = '';   // a picked colour means "use initials"
      renderProfileEditor();
    }));
  }

  function renderTeamRoster() {
    const ids = Object.keys(teamProfiles);
    const el = $('#team-roster');
    if (!el) return;
    if (!ids.length) {
      el.innerHTML = '<p class="settings-hint">No profiles found yet — yours will appear here once you save it.</p>';
      return;
    }
    el.innerHTML = ids.map(id => {
      const p = teamProfiles[id];
      return `<div class="team-row">${avatarHtml(id, p.display_name, 30)}
        <div><div class="lesson-text">${escapeHtml(p.display_name)}${id === cfg.authorId ? ' <em>(you)</em>' : ''}</div>
        <div class="feedback-when mono">${escapeHtml(id)}</div></div></div>`;
    }).join('');
  }

  async function refreshProfiles() {
    const res = await window.api.profileList();
    if (res.success) {
      teamProfiles = res.data.profiles || {};
      renderTeamRoster();
      renderProfileEditor();
    }
  }

  $('#author-name-input').addEventListener('input', renderProfileEditor);

  // Downscale before storing: the image is inlined into the profile JSON that
  // syncs to everyone, so a 4MB camera photo would be 4MB in every teammate's
  // folder. 96px is plenty for a 30px avatar on a HiDPI screen.
  $('#avatar-upload-btn').addEventListener('click', () => $('#avatar-file-input').click());
  $('#avatar-file-input').addEventListener('change', (e) => {
    const file = e.target.files && e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => {
      const img = new Image();
      img.onload = () => {
        const S = 96;
        const canvas = document.createElement('canvas');
        canvas.width = S; canvas.height = S;
        const ctx = canvas.getContext('2d');
        // Cover-crop so non-square images aren't squashed.
        const scale = Math.max(S / img.width, S / img.height);
        const w = img.width * scale, h = img.height * scale;
        ctx.drawImage(img, (S - w) / 2, (S - h) / 2, w, h);
        draftAvatarImage = canvas.toDataURL('image/png');
        renderProfileEditor();
        log(`Avatar image loaded (${Math.round(draftAvatarImage.length / 1024)} KB after resize). Save Profile to share it.`);
      };
      img.onerror = () => toast('That file could not be read as an image.', 'error');
      img.src = reader.result;
    };
    reader.onerror = () => toast('Could not read that file.', 'error');
    reader.readAsDataURL(file);
    e.target.value = '';
  });

  $('#avatar-clear-btn').addEventListener('click', () => {
    draftAvatarImage = '';
    renderProfileEditor();
  });

  $('#profile-save-btn').addEventListener('click', async () => {
    const name = $('#author-name-input').value.trim();
    if (!name) { toast('Enter a display name.', 'error'); return; }
    const btn = $('#profile-save-btn');
    btn.disabled = true; btn.textContent = 'Saving...';
    const res = await window.api.profileSave({
      displayName: name,
      avatarColor: draftAvatarColor || autoColor(cfg.authorId),
      avatarImage: draftAvatarImage
    });
    btn.disabled = false; btn.textContent = 'Save Profile';

    cfg.authorName = name;
    cfg.avatarColor = draftAvatarColor;
    cfg.avatarImage = draftAvatarImage;

    if (res.success) {
      $('#profile-save-status').textContent = 'Saved and shared with the team.';
      log(`Profile saved as "${name}".`, 'console-log-success');
      toast('Profile saved.', 'success');
      await refreshProfiles();
    } else {
      $('#profile-save-status').textContent = 'Saved locally, but not shared — check the ideation data folder in Settings.';
      log(`Profile saved locally but not to the shared folder: ${res.output}`, 'console-log-error');
      toast('Saved locally only — shared folder unreachable.', 'error');
      renderProfileEditor();
    }
  });

  async function refreshIdeationDataStatus() {
    const info = await window.api.ideationDataFolderInfo();
    $('#ideation-data-path-input').value = info.isDefault ? '' : info.root;
    $('#ideation-data-path-input').placeholder = info.isDefault
      ? `(default — local only, not shared) ${info.root}`
      : '';
    const el = $('#ideation-data-status');
    if (!info.exists) {
      el.textContent = `⚠ No Events folder found at ${info.root} — generation will create it, but check the path is right.`;
    } else if (!info.templatesFound) {
      el.textContent = `⚠ Found the folder, but no TEMPLATES/Event_Ideation_1Pager.md — documents will be generated without the studio template to follow.`;
    } else {
      el.textContent = info.isDefault
        ? 'Using the local folder bundled with this app. Nothing is shared with other designers yet.'
        : `Shared folder in use. Templates found — this is the right shape of folder.`;
    }
  }

  $('#browse-ideation-data-btn').addEventListener('click', async () => {
    const picked = await window.api.ideationBrowseDataFolder();
    if (!picked) return;
    cfg.ideationDataPath = picked;
    await window.api.saveConfig({ ideationDataPath: picked });
    await refreshIdeationDataStatus();
    log(`Ideation data folder set to ${picked}`);
  });

  $('#reset-ideation-data-btn').addEventListener('click', async () => {
    cfg.ideationDataPath = '';
    await window.api.saveConfig({ ideationDataPath: '' });
    await refreshIdeationDataStatus();
    log('Ideation data folder reset to the local default.');
  });

  $('#anthropic-key-input').addEventListener('change', async (e) => {
    cfg.anthropicApiKey = e.target.value.trim();
    await window.api.saveConfig({ anthropicApiKey: cfg.anthropicApiKey });
  });
  $('#toggle-anthropic-key-btn').addEventListener('click', () => {
    const input = $('#anthropic-key-input');
    const btn = $('#toggle-anthropic-key-btn');
    const showing = input.type === 'text';
    input.type = showing ? 'password' : 'text';
    btn.textContent = showing ? 'Show' : 'Hide';
  });

  $('#ollama-base-url-input').addEventListener('change', async (e) => {
    cfg.ollamaBaseUrl = e.target.value.trim();
    await window.api.saveConfig({ ollamaBaseUrl: cfg.ollamaBaseUrl });
  });
  $('#ollama-model-input').addEventListener('change', async (e) => {
    cfg.ollamaModel = e.target.value.trim() || 'qwen3:32b';
    await window.api.saveConfig({ ollamaModel: cfg.ollamaModel });
  });
  $('#ollama-test-btn').addEventListener('click', async () => {
    const baseUrl = $('#ollama-base-url-input').value.trim();
    const model = $('#ollama-model-input').value.trim() || 'qwen3:32b';
    const resultEl = $('#ollama-test-result');
    const btn = $('#ollama-test-btn');
    btn.disabled = true; btn.textContent = 'Testing...';
    resultEl.textContent = '';
    resultEl.classList.remove('ok', 'error');

    const res = await window.api.ideationTestOllama({ baseUrl, model });
    btn.disabled = false; btn.textContent = 'Test Connection';
    resultEl.textContent = res.output;
    resultEl.classList.add(res.success ? 'ok' : 'error');
    log(`Ollama connection test: ${res.output}`, res.success ? 'console-log-line' : 'console-log-error');
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
    log('Refreshing all source-event caches (convert_layout.lua, icon.xlsx/item.xlsx, localization.xlsx) from SVN...');

    if (!cachedEventTypes) await loadEventTypes(true);
    const eventNames = getLatestEventNamesPerType();

    const res = await window.api.refreshAllCaches({
      branchRoot: activeBranchRoot(),
      workspace: cfg.activeWorkspace,
      eventNames
    });

    btn.disabled = false;
    btn.textContent = 'Update All Caches from SVN';

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

  // card.xlsx step is a backend no-op (no event info present) — hidden permanently
  // rather than deleted from PIPELINE_STEPS, since that array's positions are the
  // stepIndex values sent straight to the Python backend and to the Unity tool's
  // own mirrored step list; removing the entry would silently shift every step after it.
  const ALWAYS_HIDDEN_STEP_INDICES = [16];

  function visibleStepIndices() {
    let indices = pipelineSteps.map((_, i) => i).filter(i => !ALWAYS_HIDDEN_STEP_INDICES.includes(i));
    if (cfg.activeWorkspace === 'WorkSpcae3') {
      indices = indices.filter(i => i !== WORKSPACE3_ONLY_STEP_INDEX);
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
  let generatedMinigameRows = [];

  function renderMinigamePreview() {
    const container = $('#minigame-preview-container');
    if (!generatedMinigameRows.length) {
      container.innerHTML = '<p class="header-subtitle small" style="margin:0;">Generate a set above to preview the rows here.</p>';
      return;
    }
    const rowsHtml = generatedMinigameRows.map(r => `
      <tr>
        <td>${r.id}</td>
        <td>${r.minigame}</td>
        <td>${r.start_time}</td>
        <td>${r.end_time}</td>
        <td>${r.double_week_id || '-'}</td>
        <td>${r.discount_id !== -1 ? r.discount_id : '-'}</td>
      </tr>`).join('');
    container.innerHTML = `
      <table class="mg-preview-table">
        <thead><tr><th>ID</th><th>Minigame</th><th>Start</th><th>End</th><th>Double Wk ID</th><th>Discount ID</th></tr></thead>
        <tbody>${rowsHtml}</tbody>
      </table>`;
  }

  function setMinigameControlsEnabled(enabled) {
    $('#minigame-generate-set-btn').disabled = !enabled;
    $('#minigame-execute-btn').disabled = !enabled;
  }

  async function loadMinigames(force) {
    if (cachedMinigames && !force) return;
    cachedMinigames = null;
    generatedMinigameRows = [];
    renderMinigamePreview();
    setMinigameControlsEnabled(false);
    try {
      const root = activeBranchRoot();
      if (!root) { log('Cannot load minigames: branch path not set.', 'console-log-error'); return; }
      const designPath = await window.api.resolveDesignPath(root, cfg.activeWorkspace);
      log('Loading Minigame configs from Excel...');
      const res = await window.api.minigamesGet({ designPath });
      if (!res.success) { log(`Error loading minigames: ${res.output}`, 'console-log-error'); toast('Could not load mini_mgr.xlsx for this branch/workspace.', 'error'); return; }
      const data = JSON.parse(res.stdout);
      if (data.error) { log(`Error: ${data.error}`, 'console-log-error'); toast('Could not load mini_mgr.xlsx for this branch/workspace.', 'error'); return; }
      cachedMinigames = data;
      log(`Loaded ${data.minigames.length} minigames. Max ID: ${data.max_id}`);
      const preview = $('#minigame-next-set-preview');
      if (preview && data.next_set_template) preview.textContent = data.next_set_template.map(t => t.name).join(' → ');
    } catch (e) {
      log(`Error parsing minigame JSON: ${e.message}`, 'console-log-error');
    } finally {
      setMinigameControlsEnabled(!!cachedMinigames);
    }
  }

  function parseMgDateTime(str) {
    const [datePart, timePart] = str.trim().split(' ');
    const [y, m, d] = datePart.split('-').map(Number);
    const [hh, mm, ss] = (timePart || '00:00:00').split(':').map(Number);
    return new Date(y, m - 1, d, hh || 0, mm || 0, ss || 0);
  }
  function formatMgDateTime(dt) {
    const pad = n => String(n).padStart(2, '0');
    return `${dt.getFullYear()}-${pad(dt.getMonth() + 1)}-${pad(dt.getDate())} ${pad(dt.getHours())}:${pad(dt.getMinutes())}:${pad(dt.getSeconds())}`;
  }
  function addMgDays(dt, days) {
    const copy = new Date(dt);
    copy.setDate(copy.getDate() + days);
    return copy;
  }

  $('#minigame-generate-set-btn').addEventListener('click', () => {
    if (!cachedMinigames || !cachedMinigames.next_set_template) { log('Minigame configs still loading...', 'console-log-error'); return; }

    const startInput = $('#minigame-set-start-input');
    const cursorStart = parseMgDateTime(startInput.value);
    if (isNaN(cursorStart.getTime())) { toast('Invalid start time. Use YYYY-MM-DD HH:MM:SS.', 'error'); return; }

    const gapDays = cachedMinigames.gap_days || 1;
    let lastQuizDw = cachedMinigames.last_quiz_dw;
    let nextId = cachedMinigames.max_id + 1;
    let cursor = cursorStart;
    const rows = [];

    cachedMinigames.next_set_template.forEach(item => {
      const start = cursor;
      const end = addMgDays(start, item.duration_days - 1);
      end.setHours(23, 59, 59);

      let dw = '';
      if (item.name === '答题挑战') {
        const yyyymm = `${start.getFullYear()}${String(start.getMonth() + 1).padStart(2, '0')}`;
        let counter = 1;
        if (lastQuizDw && lastQuizDw.slice(0, 6) === yyyymm) counter = parseInt(lastQuizDw.slice(6), 10) + 1;
        dw = `${yyyymm}${String(counter).padStart(2, '0')}`;
        lastQuizDw = dw;
      }

      const mg = cachedMinigames.minigames.find(m => m.name === item.name);
      rows.push({
        id: nextId++,
        minigame: item.name,
        start_time: formatMgDateTime(start),
        end_time: formatMgDateTime(end),
        double_week_id: dw,
        discount_id: mg ? mg.discount_id : -1
      });

      cursor = addMgDays(end, gapDays);
      cursor.setHours(0, 0, 0, 0);
    });

    generatedMinigameRows = rows;
    renderMinigamePreview();
    toast(`Generated ${rows.length} rows for the next set.`, 'success');
  });

  $('#minigame-execute-btn').addEventListener('click', async () => {
    const rows = generatedMinigameRows;
    if (rows.length === 0) { log('No generated minigame set to commit.', 'console-log-error'); return; }

    const root = activeBranchRoot();
    const designPath = await window.api.resolveDesignPath(root, cfg.activeWorkspace);
    const btn = $('#minigame-execute-btn');
    btn.disabled = true; btn.textContent = 'Committing...';
    const res = await window.api.minigamesCommit({ designPath, payload: { rows } });
    btn.disabled = false; btn.textContent = 'Commit to mini_mgr.xlsx';

    if (res.success) {
      log(`Success: ${res.stdout}`, 'console-log-success');
      toast(`${rows.length} minigame run(s) appended to mini_mgr.xlsx.`, 'success');
      generatedMinigameRows = [];
      renderMinigamePreview();
      cachedMinigames = null;
      await loadMinigames(true);
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

  // ---------------- Decoration Shop tab ----------------
  // Shared by the Alias Preview card and the PVP Bots Card Back Update section below it —
  // both need the exact same frame1/frame2/cardback aliases this cycle is about to create.
  function computeDecoShopAliases() {
    const eventId = ($('#event-id-input').value || '').trim();
    // Event ID is 8-digit yyyymmNN (e.g. 20260501) — alias uses YYMM, i.e. chars 2-6, not the leading century.
    const id4 = eventId.slice(2, 6) || '';
    const name = ($('#decoshop-event-name').value || '').trim();
    const mode1 = $('#decoshop-mode-1').value;
    const mode2 = $('#decoshop-mode-2').value;
    const eventId2 = ($('#decoshop-event-id-2').value || '').trim();
    const id4_2 = eventId2.slice(2, 6) || '';
    const name2 = ($('#decoshop-event-name-2').value || '').trim();
    return {
      frame1Alias: `${id4 || '[ID]'}${name || '[EventName]'}${mode1}`,
      frame2Alias: `${id4_2 || '[ID2]'}${name2 || '[EventName2]'}${mode2}`,
      cardbackAlias: `${id4 || '[ID]'}${name || '[EventName]'}`,
      valid: !!(id4 && name && id4_2 && name2)
    };
  }

  function updateDecoShopPreview() {
    const eventId = ($('#event-id-input').value || '').trim();
    $('#decoshop-event-id-display').value = eventId;
    const { frame1Alias, frame2Alias, cardbackAlias } = computeDecoShopAliases();
    $('#decoshop-preview-frame1').textContent = frame1Alias;
    $('#decoshop-preview-frame2').textContent = frame2Alias;
    $('#decoshop-preview-cardback').textContent = cardbackAlias;
  }
  ['#decoshop-event-name', '#decoshop-mode-1', '#decoshop-mode-2', '#decoshop-event-id-2', '#decoshop-event-name-2'].forEach(sel => {
    $(sel).addEventListener('input', updateDecoShopPreview);
    $(sel).addEventListener('change', updateDecoShopPreview);
  });
  $('#event-id-input').addEventListener('input', () => {
    if (currentView === 'decoration-shop-view') updateDecoShopPreview();
  });

  $('#decoshop-execute-btn').addEventListener('click', async () => {
    const eventId = ($('#event-id-input').value || '').trim();
    const eventName = ($('#decoshop-event-name').value || '').trim();
    const mode1 = $('#decoshop-mode-1').value;
    const mode2 = $('#decoshop-mode-2').value;
    const eventId2 = ($('#decoshop-event-id-2').value || '').trim();
    const eventName2 = ($('#decoshop-event-name-2').value || '').trim();
    const priceFrame = ($('#decoshop-price-frame').value || '').trim();
    const priceCardback = ($('#decoshop-price-cardback').value || '').trim();

    if (!eventId || eventId.length < 4) { toast('Please fill in a valid Event ID on the Main Events tab first.', 'error'); return; }
    if (!eventName) { toast('Please fill in Event Name (Game Mode 1).', 'error'); return; }
    if (!eventId2 || eventId2.length < 4) { toast('Please fill in a valid Event ID (Game Mode 2).', 'error'); return; }
    if (!eventName2) { toast('Please fill in Event Name (Game Mode 2).', 'error'); return; }
    if (mode1 === mode2) { toast('Game Mode 1 and Game Mode 2 must be different.', 'error'); return; }

    const root = activeBranchRoot();
    const designPath = await window.api.resolveDesignPath(root, cfg.activeWorkspace);
    const btn = $('#decoshop-execute-btn');
    btn.disabled = true;
    const res = await window.api.decoShopCommit({
      designPath,
      payload: {
        event_id: eventId,
        event_name: eventName,
        game_mode_1: mode1,
        game_mode_2: mode2,
        event_id_2: eventId2,
        event_name_2: eventName2,
        price_frame: priceFrame || '4',
        price_cardback: priceCardback || '10'
      }
    });
    btn.disabled = false;

    if (res.success) {
      log('Decoration Shop Commit Successful!', 'console-log-success');
      toast('Decoration Shop configuration updated across all files.', 'success');
    } else {
      log(`Decoration Shop Commit Failed: ${res.output}`, 'console-log-error');
      toast('Commit failed: ' + res.output, 'error');
    }
  });

  // ---------------- PVP Bots Card Back Update (part of the Decoration Shop tab) ----------------
  function updateRobotPvpPathDisplay() {
    $('#robotpvp-calc-path').value = cfg.robotCalcPath || '';
  }

  $('#robotpvp-browse-btn').addEventListener('click', async () => {
    const picked = await window.api.browseFile({
      startPath: cfg.robotCalcPath,
      filters: [{ name: 'Excel Workbook', extensions: ['xlsx'] }]
    });
    if (picked) {
      cfg.robotCalcPath = picked;
      updateRobotPvpPathDisplay();
      await window.api.saveConfig({ robotCalcPath: picked });
    }
  });

  $('#robotpvp-execute-btn').addEventListener('click', async () => {
    if (!cfg.robotCalcPath) { toast('Please browse for Robot_Calc.xlsx first.', 'error'); return; }
    const { frame1Alias, frame2Alias, cardbackAlias, valid } = computeDecoShopAliases();
    if (!valid) { toast('Fill in the Decoration Shop event fields above first — the same aliases are reused here.', 'error'); return; }

    const frameWeight = ($('#robotpvp-frame-weight').value || '').trim() || '30';
    const cardbackWeight = ($('#robotpvp-cardback-weight').value || '').trim() || '30';

    const root = activeBranchRoot();
    const designPath = await window.api.resolveDesignPath(root, cfg.activeWorkspace);
    const btn = $('#robotpvp-execute-btn');
    btn.disabled = true;
    const res = await window.api.robotPvpCommit({
      designPath,
      payload: {
        robot_calc_path: cfg.robotCalcPath,
        frame_aliases: [frame1Alias, frame2Alias],
        cardback_alias: cardbackAlias,
        frame_weight: frameWeight,
        cardback_weight: cardbackWeight
      }
    });
    btn.disabled = false;

    if (res.output) log(res.output, res.success ? 'console-log-line' : 'console-log-error');

    if (res.success) {
      log('PVP Bots Card Back Update Successful!', 'console-log-success');
      toast('robot.xlsx re-rolled and Robot_Calc.xlsx weight tables extended.', 'success');
    } else {
      log(`PVP Bots Card Back Update Failed: ${res.output}`, 'console-log-error');
      toast('Commit failed: ' + res.output, 'error');
    }
  });

  // ---------------- Streak Challenge tab ----------------
  // Start Date + Duration are reused directly from the Main Events tab (not re-entered here) —
  // same rationale as Event ID: one source of truth for "this event's" dates.
  function updateStreakPreview() {
    const eventId = ($('#event-id-input').value || '').trim();
    $('#streak-event-id-display').value = eventId;
    const startDateStr = $('#start-date-input').value;
    const duration = parseInt($('#duration-days-input').value, 10);
    $('#streak-start-date-display').value = startDateStr || '';
    $('#streak-duration-display').value = Number.isFinite(duration) ? duration : '';
    const times = computeEventTimes(startDateStr, duration);
    if (!times) {
      $('#streak-preview-end-date').textContent = '...';
      $('#streak-preview-rows').textContent = '...';
      return;
    }
    $('#streak-preview-end-date').textContent = times.endTime;
    $('#streak-preview-rows').textContent = Math.ceil(duration / 7);
  }
  ['#event-id-input', '#start-date-input', '#duration-days-input'].forEach(sel => {
    $(sel).addEventListener('input', () => {
      if (currentView === 'streak-challenge-view') updateStreakPreview();
    });
  });

  $('#streak-execute-btn').addEventListener('click', async () => {
    const eventId = ($('#event-id-input').value || '').trim();
    const eventName = ($('#streak-event-name').value || '').trim();
    const startDateStr = $('#start-date-input').value;
    const duration = parseInt($('#duration-days-input').value, 10);
    const times = computeEventTimes(startDateStr, duration);

    if (!eventId || eventId.length < 4) { toast('Please fill in a valid Event ID on the Main Events tab first.', 'error'); return; }
    if (!eventName) { toast('Please fill in Event Name.', 'error'); return; }
    if (!times) { toast('Please fill in Start Date and Duration (days) on the Main Events tab first.', 'error'); return; }

    const root = activeBranchRoot();
    const designPath = await window.api.resolveDesignPath(root, cfg.activeWorkspace);
    const btn = $('#streak-execute-btn');
    btn.disabled = true;
    const res = await window.api.residentTaskCommit({
      designPath,
      payload: {
        event_id: eventId,
        event_name: eventName,
        start_time: times.startTime,
        end_time: times.endTime
      }
    });
    btn.disabled = false;

    // Always show the Python script's own log output — "success" can also mean an
    // idempotent no-op (activity ID already present, nothing written), which the
    // script reports as success too; without this the console stayed silent either way.
    if (res.output) log(res.output, res.success ? 'console-log-line' : 'console-log-error');

    if (res.success) {
      if (/already present/i.test(res.output || '')) {
        log('Streak Challenge Commit: no changes made (activity ID already present).', 'console-log-line');
        toast('No changes made — that Event ID already has a cycle in resident_task.xlsx.', 'error');
      } else {
        log('Streak Challenge Commit Successful!', 'console-log-success');
        toast('resident_task.xlsx updated with the new cycle.', 'success');
      }
    } else {
      log(`Streak Challenge Commit Failed: ${res.output}`, 'console-log-error');
      toast('Commit failed: ' + res.output, 'error');
    }
  });

  // ---------------- Weekly Gift tab ----------------
  // Start Date + Duration are reused directly from the Main Events tab, same as Streak
  // Challenge. weekly_gift is a fixed weekly grid the event "claims" rows from by date
  // overlap (ceil(duration/7) rows) — it does not clone anything per event.
  function updateWeeklyGiftPreview() {
    const eventId = ($('#event-id-input').value || '').trim();
    $('#weeklygift-event-id-display').value = eventId;
    const startDateStr = $('#start-date-input').value;
    const duration = parseInt($('#duration-days-input').value, 10);
    $('#weeklygift-start-date-display').value = startDateStr || '';
    $('#weeklygift-duration-display').value = Number.isFinite(duration) ? duration : '';
    const times = computeEventTimes(startDateStr, duration);
    if (!times) {
      $('#weeklygift-preview-end-date').textContent = '...';
      $('#weeklygift-preview-rows').textContent = '...';
      return;
    }
    $('#weeklygift-preview-end-date').textContent = times.endTime;
    $('#weeklygift-preview-rows').textContent = Math.ceil(duration / 7);
  }
  ['#event-id-input', '#start-date-input', '#duration-days-input'].forEach(sel => {
    $(sel).addEventListener('input', () => {
      if (currentView === 'weekly-gift-view') updateWeeklyGiftPreview();
    });
  });

  $('#weeklygift-execute-btn').addEventListener('click', async () => {
    const eventId = ($('#event-id-input').value || '').trim();
    const startDateStr = $('#start-date-input').value;
    const duration = parseInt($('#duration-days-input').value, 10);
    const times = computeEventTimes(startDateStr, duration);

    if (!eventId || eventId.length < 4) { toast('Please fill in a valid Event ID on the Main Events tab first.', 'error'); return; }
    if (!times) { toast('Please fill in Start Date and Duration (days) on the Main Events tab first.', 'error'); return; }

    const root = activeBranchRoot();
    const designPath = await window.api.resolveDesignPath(root, cfg.activeWorkspace);
    const btn = $('#weeklygift-execute-btn');
    btn.disabled = true;
    const res = await window.api.weeklyGiftCommit({
      designPath,
      payload: {
        event_id: eventId,
        start_time: times.startTime,
        end_time: times.endTime
      }
    });
    btn.disabled = false;

    if (res.output) log(res.output, res.success ? 'console-log-line' : 'console-log-error');

    if (res.success) {
      if (/No changes needed/i.test(res.output || '')) {
        log('Weekly Gift Commit: no changes made (rows already enabled).', 'console-log-line');
        toast('No changes made — those weekly_gift rows are already enabled.', 'error');
      } else {
        log('Weekly Gift Commit Successful!', 'console-log-success');
        toast('store.xlsx (weekly_gift) updated for this event.', 'success');
      }
    } else {
      log(`Weekly Gift Commit Failed: ${res.output}`, 'console-log-error');
      toast('Commit failed: ' + res.output, 'error');
    }
  });

  // ---------------- Full Screen Gift tab ----------------
  // recommend_gift (store.xlsx) slot-16 rows recur on a ~14-day cadence, each cloned
  // verbatim (recharge list, formula, purchase count, etc.) from the most recent
  // instance — verified identical across the whole existing series. Only
  // ID/Alias/Start/End are regenerated per new instance, and switch is always forced
  // to FALSE regardless of the template (the sheet's own default).
  async function loadFullScreenGift(force) {
    if (cachedFSG && !force) { updateFSGPreview(); return; }
    cachedFSG = null;
    const root = activeBranchRoot();
    if (!root) { log('Cannot load Full Screen Gift info: branch path not set.', 'console-log-error'); return; }
    const designPath = await window.api.resolveDesignPath(root, cfg.activeWorkspace);
    log('Loading Full Screen Gift metadata from store.xlsx...');
    const res = await window.api.fullScreenGiftGet({ designPath });
    if (!res.success) { log(`Error loading Full Screen Gift info: ${res.output}`, 'console-log-error'); toast('Could not load store.xlsx for this branch/workspace.', 'error'); return; }
    try {
      cachedFSG = JSON.parse(res.stdout);
      if (cachedFSG.error) { log(`Error: ${cachedFSG.error}`, 'console-log-error'); toast(cachedFSG.error, 'error'); cachedFSG = null; return; }
      log(`Full Screen Gift metadata loaded. Next ID: ${cachedFSG.next_id}`);
      if (!$('#fsg-start-date').value && cachedFSG.suggested_next_start) {
        $('#fsg-start-date').value = cachedFSG.suggested_next_start;
      }
      updateFSGPreview();
    } catch (e) {
      log(`Error parsing Full Screen Gift JSON: ${e.message}`, 'console-log-error');
    }
  }

  function updateFSGPreview() {
    if (!cachedFSG) return;
    $('#fsg-preview-id').textContent = cachedFSG.next_id;
    const startDateStr = $('#fsg-start-date').value.trim();
    const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(startDateStr);
    if (!m) {
      $('#fsg-preview-alias').textContent = '...';
      $('#fsg-preview-start').textContent = '...';
      $('#fsg-preview-end').textContent = '...';
      return;
    }
    const [, y, mo, d] = m;
    const yy = y.slice(2);
    $('#fsg-preview-alias').textContent = `全屏礼包${yy}${mo}${d}`;
    const startMs = Date.UTC(Number(y), Number(mo) - 1, Number(d));
    const endMs = startMs + (cachedFSG.window_days - 1) * 86400000;
    $('#fsg-preview-start').textContent = `${startDateStr} 00:00:00`;
    $('#fsg-preview-end').textContent = `${fmtDate(endMs)} 23:59:59`;
  }
  $('#fsg-start-date').addEventListener('input', updateFSGPreview);

  $('#fsg-execute-btn').addEventListener('click', async () => {
    if (!cachedFSG) { toast('Full Screen Gift metadata still loading...', 'error'); return; }
    const startDateStr = $('#fsg-start-date').value.trim();
    if (!/^\d{4}-\d{2}-\d{2}$/.test(startDateStr)) { toast('Please enter a valid Start Date (yyyy-mm-dd).', 'error'); return; }

    const root = activeBranchRoot();
    const designPath = await window.api.resolveDesignPath(root, cfg.activeWorkspace);
    const btn = $('#fsg-execute-btn');
    btn.disabled = true; btn.textContent = 'Committing...';
    const res = await window.api.fullScreenGiftCommit({ designPath, payload: { start_date: startDateStr } });
    btn.disabled = false; btn.textContent = 'Commit Full Screen Gift';

    if (res.output) log(res.output, res.success ? 'console-log-line' : 'console-log-error');

    if (res.success) {
      if (/already exists/i.test(res.output || '')) {
        log('Full Screen Gift Commit: no changes made (row already exists for that date).', 'console-log-line');
        toast('No changes made — a Full Screen Gift row already exists for that date.', 'error');
      } else {
        log('Full Screen Gift Commit Successful!', 'console-log-success');
        toast('store.xlsx (recommend_gift) updated with the new Full Screen Gift.', 'success');
        cachedFSG = null;
        await loadFullScreenGift(true);
      }
    } else {
      log(`Full Screen Gift Commit Failed: ${res.output}`, 'console-log-error');
      toast('Commit failed: ' + res.output, 'error');
    }
  });

  // ---------------- Endless Treasure tab ----------------
  // 折扣链数据 event-chain (类型=2) triplets alternate between two fixed
  // stage-ID templates every ~window (verified against SVN history: the two
  // templates' reward content has never diverged). New triplet clones the
  // OTHER slot's most recent triplet verbatim; only ID + 说明 change. See the
  // static analysis cards above this form for the full mechanism.
  async function loadEndlessTreasure(force) {
    if (cachedEndlessTreasure && !force) { updateEndlessTreasurePreview(); return; }
    cachedEndlessTreasure = null;
    const root = activeBranchRoot();
    if (!root) { log('Cannot load Endless Treasure info: branch path not set.', 'console-log-error'); return; }
    const designPath = await window.api.resolveDesignPath(root, cfg.activeWorkspace);
    log('Loading Endless Treasure metadata from endless_treasure.xlsx...');
    const res = await window.api.endlessTreasureGet({ designPath });
    if (!res.success) { log(`Error loading Endless Treasure info: ${res.output}`, 'console-log-error'); toast('Could not load endless_treasure.xlsx for this branch/workspace.', 'error'); return; }
    try {
      cachedEndlessTreasure = JSON.parse(res.stdout);
      if (cachedEndlessTreasure.error) { log(`Error: ${cachedEndlessTreasure.error}`, 'console-log-error'); toast(cachedEndlessTreasure.error, 'error'); cachedEndlessTreasure = null; return; }
      log(`Endless Treasure metadata loaded. Next chain IDs: ${cachedEndlessTreasure.next_chain_ids}`);
      if (!$('#et-start-date').value && cachedEndlessTreasure.suggested_next_start) {
        $('#et-start-date').value = cachedEndlessTreasure.suggested_next_start;
      }
      updateEndlessTreasurePreview();
    } catch (e) {
      log(`Error parsing Endless Treasure JSON: ${e.message}`, 'console-log-error');
    }
  }

  function updateEndlessTreasurePreview() {
    if (!cachedEndlessTreasure) return;
    $('#et-preview-rec-id').textContent = cachedEndlessTreasure.next_rec_id;
    $('#et-preview-chain-ids').textContent = cachedEndlessTreasure.next_chain_ids.join(', ');
    const startDateStr = $('#et-start-date').value.trim();
    const eventName = $('#et-event-name').value.trim();
    const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(startDateStr);
    if (!m || !eventName) {
      $('#et-preview-start').textContent = '...';
      $('#et-preview-end').textContent = '...';
      $('#et-preview-theme').textContent = '...';
      return;
    }
    const [, y, mo, d] = m;
    const startMs = Date.UTC(Number(y), Number(mo) - 1, Number(d));
    const endMs = startMs + (cachedEndlessTreasure.window_days - 1) * 86400000;
    $('#et-preview-start').textContent = `${startDateStr} 00:00:00`;
    $('#et-preview-end').textContent = `${fmtDate(endMs)} 23:59:59`;
    $('#et-preview-theme').textContent = `${y}${mo}${eventName.replace(/[^a-zA-Z0-9]/g, '').toLowerCase()}`;
  }
  $('#et-start-date').addEventListener('input', updateEndlessTreasurePreview);
  $('#et-event-name').addEventListener('input', updateEndlessTreasurePreview);
  $('#et-refresh-btn').addEventListener('click', () => loadEndlessTreasure(true));

  $('#et-execute-btn').addEventListener('click', async () => {
    if (!cachedEndlessTreasure) { toast('Endless Treasure metadata still loading...', 'error'); return; }
    const startDateStr = $('#et-start-date').value.trim();
    const eventName = $('#et-event-name').value.trim();
    if (!/^\d{4}-\d{2}-\d{2}$/.test(startDateStr)) { toast('Please enter a valid Start Date (yyyy-mm-dd).', 'error'); return; }
    if (!eventName) { toast('Please enter an Event Name.', 'error'); return; }

    const root = activeBranchRoot();
    const designPath = await window.api.resolveDesignPath(root, cfg.activeWorkspace);
    const btn = $('#et-execute-btn');
    btn.disabled = true; btn.textContent = 'Committing...';
    const res = await window.api.endlessTreasureCommit({ designPath, payload: { start_date: startDateStr, event_name: eventName } });
    btn.disabled = false; btn.textContent = 'Commit Endless Treasure';

    if (res.output) log(res.output, res.success ? 'console-log-line' : 'console-log-error');

    if (res.success) {
      if (/already exists/i.test(res.output || '')) {
        log('Endless Treasure Commit: no changes made (a window already exists for that date).', 'console-log-line');
        toast('No changes made — a recommendation row already exists for that date.', 'error');
      } else {
        log('Endless Treasure Commit Successful!', 'console-log-success');
        toast('endless_treasure.xlsx updated with the new chain window.', 'success');
        cachedEndlessTreasure = null;
        await loadEndlessTreasure(true);
      }
    } else {
      log(`Endless Treasure Commit Failed: ${res.output}`, 'console-log-error');
      toast('Commit failed: ' + res.output, 'error');
    }
  });

  // ---------------- Trivia Challenge tab ----------------
  async function loadTrivia(force) {
    if (cachedTrivia && !force) { updateTriviaPreview(); return; }
    cachedTrivia = null;
    $('#trivia-preview-prefix').textContent = '...';
    $('#trivia-preview-batch').textContent = '...';
    $('#trivia-preview-count').textContent = '...';
    $('#trivia-preview-tiers').textContent = '...';
    $('#trivia-preview-status').textContent = '...';
    const root = activeBranchRoot();
    if (!root) { log('Cannot load Trivia Challenge info: branch path not set.', 'console-log-error'); return; }
    log('Loading Trivia Challenge cycle info from mini_mgr.xlsx...');
    const res = await window.api.triviaGet({ branchRoot: root, workspace: cfg.activeWorkspace });
    if (!res.success) { log(`Error loading Trivia Challenge info: ${res.output}`, 'console-log-error'); toast('Could not compute the upcoming 答题挑战 cycle for this branch/workspace.', 'error'); return; }
    try {
      cachedTrivia = JSON.parse(res.stdout);
      if (cachedTrivia.error) { log(`Error: ${cachedTrivia.error}`, 'console-log-error'); toast(cachedTrivia.error, 'error'); cachedTrivia = null; return; }
      log(`Trivia Challenge cycle info loaded. Prefix: ${cachedTrivia.prefix}`);
      updateTriviaPreview();
    } catch (e) {
      log(`Error parsing Trivia Challenge JSON: ${e.message}`, 'console-log-error');
    }
  }

  function updateTriviaPreview() {
    if (!cachedTrivia) return;
    $('#trivia-preview-prefix').textContent = cachedTrivia.prefix;
    $('#trivia-preview-batch').textContent = cachedTrivia.batch_number != null ? `#${cachedTrivia.batch_number}` : '...';
    $('#trivia-preview-count').textContent = cachedTrivia.question_count;
    const t = cachedTrivia.tier_counts || {};
    $('#trivia-preview-tiers').textContent = `${t['1'] ?? 0} / ${t['2'] ?? 0} / ${t['3'] ?? 0}`;
    $('#trivia-preview-status').textContent = cachedTrivia.already_current ? 'Already current' : 'Needs refresh';
  }

  $('#trivia-refresh-btn').addEventListener('click', () => loadTrivia(true));

  $('#trivia-execute-btn').addEventListener('click', async () => {
    if (!cachedTrivia) { toast('Trivia Challenge info still loading...', 'error'); return; }
    if (cachedTrivia.already_current) {
      toast('Nothing to do — the quiz already matches the upcoming cycle.');
      return;
    }
    const root = activeBranchRoot();
    const btn = $('#trivia-execute-btn');
    btn.disabled = true; btn.textContent = 'Running...';
    const res = await window.api.triviaRun({ branchRoot: root, workspace: cfg.activeWorkspace });
    btn.disabled = false; btn.textContent = 'Run Trivia Challenge Cycle';

    if (res.success) {
      log('Trivia Challenge Cycle Successful!', 'console-log-success');
      toast('localization_quiz.xlsx and answer_challenge.xlsx updated for the new cycle.', 'success');
      await loadTrivia(true);
    } else {
      log(`Trivia Challenge Cycle Failed: ${res.output}`, 'console-log-error');
      toast('Run failed: ' + res.output, 'error');
    }
  });

  // ---------------- Ideation tab ----------------
  // Pure generate-on-demand: no "get" call on tab switch, no cache to
  // invalidate on branch change — each run calls the Claude API fresh with
  // the currently-selected event type/duration plus this branch's recent
  // event history for context.
  function escapeHtml(s) {
    return String(s == null ? '' : s)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  function fitPillClass(rating) {
    if (rating === 'Green') return 'fit-pill fit-green';
    if (rating === 'Yellow') return 'fit-pill fit-yellow';
    if (rating === 'Red') return 'fit-pill fit-red';
    return 'fit-pill';
  }

  function renderIdeationCard(theme, idx) {
    const t = theme;
    const selected = idx === ideationSelectedIndex;
    const scenes = (t.scenes || []).map(s =>
      `<div class="preview-cell"><span class="preview-label">${escapeHtml(s.name)}</span><span class="preview-value">${escapeHtml(s.description)}</span></div>`
    ).join('');
    const keyFeatures = (t.visual_direction && t.visual_direction.key_features || []).map(f => `<li>${escapeHtml(f)}</li>`).join('');
    const uiDirection = (t.visual_direction && t.visual_direction.ui_direction || []).map(f => `<li>${escapeHtml(f)}</li>`).join('');
    const af = t.audience_fit || {};
    const eff = t.effort_estimate || {};
    const imp = t.expected_impact || {};

    return `
      <div class="stacked-card ideation-card${selected ? ' ideation-selected' : ''}" data-idx="${idx}">
        <div class="ideation-card-header">
          <div>
            <h2 class="card-title">${escapeHtml(t.theme_name)}</h2>
            <span class="preview-label mono">${escapeHtml(t.internal_reference_name)}</span>
          </div>
          <button class="secondary-btn small ideation-select-btn" data-idx="${idx}">${selected ? 'Selected ✓' : 'Select as Final'}</button>
        </div>
        <p class="settings-hint">${escapeHtml(t.reasoning)}</p>
        ${ideationDraftPaths[idx] ? `<p class="settings-hint ideation-saved-path"><strong>Draft document:</strong> <span class="mono">${escapeHtml(ideationDraftPaths[idx])}</span> — open this to review or mark up directly.</p>` : ''}
        ${selected && ideationSavedPath ? `<p class="settings-hint ideation-saved-path"><strong>Saved as final:</strong> <span class="mono">${escapeHtml(ideationSavedPath)}</span></p>` : ''}

        <div class="preview-grid three">
          <div class="preview-cell"><span class="preview-label">Progression Completionist</span><span class="${fitPillClass(af.progression_completionist)}">${escapeHtml(af.progression_completionist)}</span></div>
          <div class="preview-cell"><span class="preview-label">Casual Escape</span><span class="${fitPillClass(af.casual_escape)}">${escapeHtml(af.casual_escape)}</span></div>
          <div class="preview-cell"><span class="preview-label">Social/Collector</span><span class="${fitPillClass(af.social_collector)}">${escapeHtml(af.social_collector)}</span></div>
        </div>
        ${af.notes ? `<p class="settings-hint">${escapeHtml(af.notes)}</p>` : ''}

        <details class="info-disclosure">
          <summary>
            <span class="info-disclosure-icon" aria-hidden="true"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9"/><path d="M12 8h.01M11 12h1v4h1"/></svg></span>
            <span>Visual direction &amp; 3-scene progression</span>
            <svg class="info-disclosure-chevron" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M6 9l6 6 6-6"/></svg>
          </summary>
          <div class="info-disclosure-body">
            <p class="settings-hint"><strong>Art style:</strong> ${escapeHtml(t.visual_direction && t.visual_direction.art_style)}</p>
            <p class="settings-hint"><strong>Palette:</strong> ${escapeHtml(t.visual_direction && t.visual_direction.color_palette)}</p>
            <p class="settings-hint"><strong>Key features:</strong></p><ul class="ideation-list">${keyFeatures}</ul>
            <p class="settings-hint"><strong>UI direction:</strong></p><ul class="ideation-list">${uiDirection}</ul>
            <div class="preview-grid three">${scenes}</div>
          </div>
        </details>

        <details class="info-disclosure">
          <summary>
            <span class="info-disclosure-icon" aria-hidden="true"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9"/><path d="M12 8h.01M11 12h1v4h1"/></svg></span>
            <span>New features, impact &amp; effort</span>
            <svg class="info-disclosure-chevron" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M6 9l6 6 6-6"/></svg>
          </summary>
          <div class="info-disclosure-body">
            <p class="settings-hint"><strong>New features:</strong> ${escapeHtml(t.new_features)}</p>
            <p class="settings-hint"><strong>Engagement (draft estimate):</strong> ${escapeHtml(imp.engagement_note)}</p>
            <p class="settings-hint"><strong>Monetization (draft estimate):</strong> ${escapeHtml(imp.monetization_note)}</p>
            <div class="preview-grid four">
              <div class="preview-cell"><span class="preview-label">Art</span><span class="preview-value">${escapeHtml(eff.art)}</span></div>
              <div class="preview-cell"><span class="preview-label">Engineering</span><span class="preview-value">${escapeHtml(eff.engineering)}</span></div>
              <div class="preview-cell"><span class="preview-label">Design</span><span class="preview-value">${escapeHtml(eff.design)}</span></div>
              <div class="preview-cell"><span class="preview-label">Overall</span><span class="preview-value">${escapeHtml(eff.overall)}</span></div>
            </div>
          </div>
        </details>

        ${threadHtml('ideation', t.internal_reference_name)}

        <div class="board-divider"></div>
        <div class="form-row">
          <label class="form-label" for="ideation-feedback-${idx}">Feedback</label>
          <textarea id="ideation-feedback-${idx}" class="text-input ideation-feedback" rows="2" placeholder="e.g. Make the colors more vibrant; reference our anniversary event in the props"></textarea>
        </div>
        <div class="execution-footer" style="justify-content:flex-end;">
          <button class="primary-btn small ideation-regen-btn" data-idx="${idx}">Submit Feedback &amp; Regenerate</button>
        </div>
      </div>
    `;
  }

  // Art Requirement docs are keyed to the finalized ideation theme's ref
  // code (e.g. Merge2609), so their feedback thread hangs off the same event.
  function artReqRefName() {
    if (ideationSelectedIndex === -1 || !ideationThemes) return '';
    return ideationThemes[ideationSelectedIndex].internal_reference_name || '';
  }

  // ---- Shared feedback thread + lessons (both pipelines) ----
  function fmtWhen(iso) {
    if (!iso) return '';
    const d = new Date(iso);
    return isNaN(d) ? iso : d.toLocaleString();
  }

  async function refreshIdeationEntries(eventType) {
    const res = await window.api.ideationEntries({ eventType });
    if (!res.success) {
      log(`Could not read shared feedback: ${res.output}`, 'console-log-error');
      ideationEntries = []; ideationLessons = [];
      return;
    }
    ideationEntries = res.data.entries || [];
    ideationLessons = res.data.lessons || [];
    if (res.data.profiles) teamProfiles = res.data.profiles;
  }

  // The feedback thread for one specific document, so a reviewer sees what
  // colleagues already said before adding their own note.
  function threadHtml(docKind, refName) {
    const items = ideationEntries.filter(e => e.doc_kind === docKind && e.ref_name === refName);
    if (!items.length) return '';
    const rows = items.map(e => `
      <div class="feedback-entry${e.retracted ? ' feedback-retracted' : ''}">
        <div class="feedback-entry-head">
          ${avatarHtml(e.author_id, e.author_name, 22)}
          <strong>${escapeHtml(e.author_name)}</strong>
          <span class="feedback-when">${escapeHtml(fmtWhen(e.created_at))}</span>
        </div>
        <div class="feedback-body">${escapeHtml(e.feedback)}</div>
        ${e.extracted_lesson ? `<div class="feedback-lesson">Lesson: ${escapeHtml(e.extracted_lesson)}${e.retracted ? ` <em>(retracted by ${escapeHtml(e.retracted_by || 'someone')})</em>` : ''}</div>` : ''}
      </div>`).join('');
    return `
      <details class="info-disclosure" open>
        <summary>
          <span class="info-disclosure-icon" aria-hidden="true"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"/></svg></span>
          <span>Feedback from the team (${items.length})</span>
          <svg class="info-disclosure-chevron" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M6 9l6 6 6-6"/></svg>
        </summary>
        <div class="info-disclosure-body">${rows}</div>
      </details>`;
  }

  function renderLessonsPanel() {
    const panel = $('#ideation-lessons-panel');
    if (!panel) return;
    if (!ideationLessons.length) {
      panel.innerHTML = '';
      return;
    }
    const rows = ideationLessons.map(l => `
      <div class="lesson-row">
        ${avatarHtml(l.author_id, l.author_name, 26)}
        <div>
          <div class="lesson-text">${escapeHtml(l.lesson)}</div>
          <div class="feedback-when">from ${escapeHtml(l.author_name)} · ${escapeHtml(fmtWhen(l.created_at))} · ${escapeHtml(l.doc_kind === 'artreq' ? 'Art Requirement' : 'Ideation')}</div>
        </div>
        <button class="secondary-btn small lesson-retract-btn" data-entry="${escapeHtml(l.entry_id)}">Retract</button>
      </div>`).join('');
    panel.innerHTML = `
      <div class="stacked-card">
        <h2 class="card-title">Active Lessons (${ideationLessons.length})</h2>
        <p class="settings-hint">Folded into every future generation for this event type — from everyone sharing this folder, not just you. Retract one that's steering results the wrong way; the feedback stays in the thread, only its influence on generation stops.</p>
        ${rows}
      </div>`;
    $$('.lesson-retract-btn').forEach(btn => {
      btn.addEventListener('click', () => retractLesson(btn.dataset.entry));
    });
  }

  async function retractLesson(entryId) {
    const eventType = $('#ideation-event-type').value;
    const res = await window.api.ideationRetract({ eventType, entryId });
    if (!res.success) {
      log(`Retract failed: ${res.output}`, 'console-log-error');
      toast('Could not retract that lesson: ' + res.output, 'error');
      return;
    }
    log('Lesson retracted — it will no longer steer future generations.', 'console-log-success');
    toast('Lesson retracted.', 'success');
    await refreshIdeationEntries(eventType);
    renderLessonsPanel();
    renderIdeationResults();
    renderArtReqResults();
  }

  function renderIdeationResults() {
    const container = $('#ideation-results');
    if (!ideationThemes || !ideationThemes.length) { container.innerHTML = ''; return; }
    container.innerHTML = ideationThemes.map((t, i) => renderIdeationCard(t, i)).join('');

    $$('.ideation-select-btn').forEach(btn => {
      btn.addEventListener('click', () => selectIdeationTheme(Number(btn.dataset.idx)));
    });
    $$('.ideation-regen-btn').forEach(btn => {
      btn.addEventListener('click', () => regenerateIdeationTheme(Number(btn.dataset.idx)));
    });
  }

  async function selectIdeationTheme(idx) {
    const eventType = $('#ideation-event-type').value;
    const durationDays = parseInt($('#ideation-duration').value, 10) || 14;
    const theme = ideationThemes[idx];

    const btn = document.querySelector(`.ideation-select-btn[data-idx="${idx}"]`);
    if (btn) { btn.disabled = true; btn.textContent = 'Saving...'; }
    log(`Saving "${theme.theme_name}" as the final 1-Pager...`);

    const res = await window.api.ideationSaveFinal({ eventType, durationDays, theme });

    if (res.success) {
      ideationSelectedIndex = idx;
      ideationSavedPath = res.output;
      log(`Saved to ${res.output}`, 'console-log-success');
      toast('Saved as the final 1-Pager for this event.', 'success');
      renderIdeationResults();
      $('#artreq-trigger-card').style.display = 'block';
      artReqVariations = null;
      artReqDraftPaths = [];
      artReqSelectedIndex = -1;
      artReqSavedPath = '';
      $('#artreq-results').innerHTML = '';
    } else {
      log(`Could not save the final 1-Pager: ${res.output}`, 'console-log-error');
      toast('Save failed: ' + res.output, 'error');
      if (btn) { btn.disabled = false; btn.textContent = 'Select as Final'; }
    }
  }

  function clampThemeCount() {
    const input = $('#ideation-theme-count');
    let n = parseInt(input.value, 10);
    if (!Number.isFinite(n)) n = 5;
    n = Math.max(1, Math.min(10, n));
    input.value = n;
    return n;
  }
  $('#ideation-theme-count').addEventListener('change', clampThemeCount);

  // Lessons and feedback are per event type — reload when the picker changes
  // so you're never looking at another type's history.
  $('#ideation-event-type').addEventListener('change', async (e) => {
    await refreshIdeationEntries(e.target.value);
    renderLessonsPanel();
    renderIdeationResults();
    renderArtReqResults();
  });

  $('#ideation-generate-btn').addEventListener('click', async () => {
    const eventType = $('#ideation-event-type').value;
    const durationDays = parseInt($('#ideation-duration').value, 10) || 14;
    const themeCount = clampThemeCount();
    const root = activeBranchRoot();
    if (!root) { toast('Set a branch path in Settings first.', 'error'); return; }
    if (!cfg.anthropicApiKey && !cfg.ollamaBaseUrl) { toast('Add an Anthropic API key or an Ollama Base URL in Settings first.', 'error'); return; }

    const designPath = await window.api.resolveDesignPath(root, cfg.activeWorkspace);
    const btn = $('#ideation-generate-btn');
    const providerLabel = cfg.anthropicApiKey ? 'Claude' : `local Ollama (${cfg.ollamaModel || 'qwen3:32b'})`;
    btn.disabled = true; btn.textContent = `Generating ${themeCount} theme${themeCount === 1 ? '' : 's'}...`;
    log(`Generating ${themeCount} ${eventType} ideation theme${themeCount === 1 ? '' : 's'} (${durationDays}-day sprint) via ${providerLabel}...`);

    const res = await window.api.ideationGenerate({ eventType, durationDays, themeCount, designPath });
    btn.disabled = false; btn.textContent = 'Generate Themes';

    if (!res.success) { log(`Ideation generation failed: ${res.output}`, 'console-log-error'); toast('Generation failed: ' + res.output, 'error'); return; }
    try {
      const parsed = JSON.parse(res.stdout);
      if (parsed.error) { log(`Error: ${parsed.error}`, 'console-log-error'); toast(parsed.error, 'error'); return; }
      await refreshIdeationEntries(eventType);
      renderLessonsPanel();
      ideationThemes = parsed.themes;
      ideationDraftPaths = parsed.draft_paths || [];
      ideationSelectedIndex = -1;
      ideationSavedPath = '';
      $('#artreq-trigger-card').style.display = 'none';
      $('#artreq-results').innerHTML = '';
      artReqVariations = null;
      artReqDraftPaths = [];
      artReqSelectedIndex = -1;
      artReqSavedPath = '';
      renderIdeationResults();
      log(`Generated ${ideationThemes.length} theme draft(s) via ${parsed.provider === 'ollama' ? 'local Ollama' : 'Claude'}. Draft documents written under ideationData/Events/${eventType}/One-Pager Ideation/Drafts/.`, 'console-log-success');
      toast(`${ideationThemes.length} theme(s) generated — draft documents saved to disk.`, 'success');
    } catch (e) {
      log(`Error parsing ideation response: ${e.message}`, 'console-log-error');
      toast('Could not parse the generated themes.', 'error');
    }
  });

  async function regenerateIdeationTheme(idx) {
    const feedback = $(`#ideation-feedback-${idx}`).value.trim();
    if (!feedback) { toast('Enter feedback before regenerating.', 'error'); return; }
    const eventType = $('#ideation-event-type').value;
    const durationDays = parseInt($('#ideation-duration').value, 10) || 14;
    const theme = ideationThemes[idx];

    const btn = document.querySelector(`.ideation-regen-btn[data-idx="${idx}"]`);
    if (btn) { btn.disabled = true; btn.textContent = 'Regenerating...'; }
    log(`Regenerating theme "${theme.theme_name}" with feedback...`);

    const res = await window.api.ideationRegenerate({ eventType, durationDays, theme, feedback });

    if (!res.success) {
      log(`Regeneration failed: ${res.output}`, 'console-log-error');
      toast('Regeneration failed: ' + res.output, 'error');
      if (btn) { btn.disabled = false; btn.textContent = 'Submit Feedback & Regenerate'; }
      return;
    }
    try {
      const parsed = JSON.parse(res.stdout);
      if (parsed.error) { log(`Error: ${parsed.error}`, 'console-log-error'); toast(parsed.error, 'error'); return; }
      ideationThemes[idx] = parsed.theme;
      if (parsed.draft_path) ideationDraftPaths[idx] = parsed.draft_path;
      await refreshIdeationEntries(eventType);
      renderLessonsPanel();
      renderIdeationResults();
      log(`Theme "${parsed.theme.theme_name}" regenerated. Draft document updated on disk. Feedback saved for future runs.`, 'console-log-success');
      toast('Theme regenerated — draft document + feedback updated.', 'success');
    } catch (e) {
      log(`Error parsing regeneration response: ${e.message}`, 'console-log-error');
      toast('Could not parse the regenerated theme.', 'error');
    } finally {
      if (btn) { btn.disabled = false; btn.textContent = 'Submit Feedback & Regenerate'; }
    }
  }

  // ---------------- Art Requirement tab ----------------
  // Only reachable once an Ideation theme has been saved as Final (see
  // selectIdeationTheme above, which reveals #artreq-trigger-card). Same
  // generate/regenerate/select-final shape as Ideation, one step later in
  // the studio's real pipeline.
  function renderArtReqCard(v, idx) {
    const selected = idx === artReqSelectedIndex;
    const go = v.gameplay_overview || {};
    const zonesHtml = (v.zones || []).map(z => {
      const steps = (z.construction_steps || []).map(s => `<li>${escapeHtml(s)}</li>`).join('');
      return `
        <details class="info-disclosure">
          <summary>
            <span class="info-disclosure-icon" aria-hidden="true"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9"/><path d="M12 8h.01M11 12h1v4h1"/></svg></span>
            <span>Zone ${escapeHtml(z.zone_number)} — ${escapeHtml(z.zone_theme)}</span>
            <svg class="info-disclosure-chevron" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M6 9l6 6 6-6"/></svg>
          </summary>
          <div class="info-disclosure-body">
            <p class="settings-hint"><strong>Palette/Mood:</strong> ${escapeHtml(z.palette_or_mood)}</p>
            <p class="settings-hint"><strong>Construction/Reveal Steps:</strong></p>
            <ol class="ideation-list">${steps}</ol>
            <p class="settings-hint"><strong>Reference Notes:</strong> ${escapeHtml(z.reference_notes)}</p>
          </div>
        </details>`;
    }).join('');
    const charsRows = (v.characters || []).map(c =>
      `<tr><td>${escapeHtml(c.name)}</td><td>${escapeHtml(c.role)}</td><td>${escapeHtml(c.design_notes)}</td><td>${escapeHtml((c.scene_actions || []).join(', '))}</td></tr>`
    ).join('');

    return `
      <div class="stacked-card ideation-card${selected ? ' ideation-selected' : ''}" data-idx="${idx}">
        <div class="ideation-card-header">
          <div><h2 class="card-title">${escapeHtml(v.title)}</h2></div>
          <button class="secondary-btn small artreq-select-btn" data-idx="${idx}">${selected ? 'Selected ✓' : 'Select as Final'}</button>
        </div>
        ${artReqDraftPaths[idx] ? `<p class="settings-hint ideation-saved-path"><strong>Draft document:</strong> <span class="mono">${escapeHtml(artReqDraftPaths[idx])}</span></p>` : ''}
        ${selected && artReqSavedPath ? `<p class="settings-hint ideation-saved-path"><strong>Saved as final:</strong> <span class="mono">${escapeHtml(artReqSavedPath)}</span></p>` : ''}

        <div class="preview-grid three">
          <div class="preview-cell"><span class="preview-label">Duration</span><span class="preview-value">${escapeHtml(go.duration_summary)}</span></div>
          <div class="preview-cell"><span class="preview-label">Core Loop</span><span class="preview-value">${escapeHtml(go.loop_summary)}</span></div>
          <div class="preview-cell"><span class="preview-label">Resource Planning</span><span class="preview-value">${escapeHtml(go.resource_planning)}</span></div>
        </div>
        <p class="settings-hint">${escapeHtml(go.narrative_packaging)}</p>

        ${zonesHtml}

        <details class="info-disclosure">
          <summary>
            <span class="info-disclosure-icon" aria-hidden="true"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9"/><path d="M12 8h.01M11 12h1v4h1"/></svg></span>
            <span>Gameplay-specific art requirement</span>
            <svg class="info-disclosure-chevron" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M6 9l6 6 6-6"/></svg>
          </summary>
          <div class="info-disclosure-body"><p class="settings-hint">${escapeHtml(v.gameplay_specific_content)}</p></div>
        </details>

        <details class="info-disclosure">
          <summary>
            <span class="info-disclosure-icon" aria-hidden="true"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9"/><path d="M12 8h.01M11 12h1v4h1"/></svg></span>
            <span>Character requirements</span>
            <svg class="info-disclosure-chevron" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M6 9l6 6 6-6"/></svg>
          </summary>
          <div class="info-disclosure-body">
            <table class="mg-preview-table">
              <thead><tr><th>Name</th><th>Role</th><th>Design Notes</th><th>Scene Actions</th></tr></thead>
              <tbody>${charsRows}</tbody>
            </table>
          </div>
        </details>

        <details class="info-disclosure">
          <summary>
            <span class="info-disclosure-icon" aria-hidden="true"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9"/><path d="M12 8h.01M11 12h1v4h1"/></svg></span>
            <span>Art direction &amp; UI icon refresh</span>
            <svg class="info-disclosure-chevron" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M6 9l6 6 6-6"/></svg>
          </summary>
          <div class="info-disclosure-body">
            <p class="settings-hint"><strong>Lean into:</strong></p>
            <ul class="ideation-list">${((v.art_direction_guidance || {}).lean_into || []).map(x => `<li>${escapeHtml(x)}</li>`).join('')}</ul>
            <p class="settings-hint"><strong>Avoid:</strong></p>
            <ul class="ideation-list">${((v.art_direction_guidance || {}).avoid || []).map(x => `<li>${escapeHtml(x)}</li>`).join('')}</ul>
            <p class="settings-hint"><strong>UI icon refresh:</strong> ${escapeHtml(v.ui_icon_refresh_notes)}</p>
          </div>
        </details>

        ${threadHtml('artreq', artReqRefName())}

        <div class="board-divider"></div>
        <div class="form-row">
          <label class="form-label" for="artreq-feedback-${idx}">Feedback</label>
          <textarea id="artreq-feedback-${idx}" class="text-input ideation-feedback" rows="2" placeholder="e.g. Icons need a consistent outline width; the gift NPC doesn't match the theme"></textarea>
        </div>
        <div class="execution-footer" style="justify-content:flex-end;">
          <button class="primary-btn small artreq-regen-btn" data-idx="${idx}">Submit Feedback &amp; Regenerate</button>
        </div>
      </div>
    `;
  }

  function renderArtReqResults() {
    const container = $('#artreq-results');
    if (!artReqVariations || !artReqVariations.length) { container.innerHTML = ''; return; }
    container.innerHTML = artReqVariations.map((v, i) => renderArtReqCard(v, i)).join('');
    $$('.artreq-select-btn').forEach(btn => btn.addEventListener('click', () => selectArtReqVariation(Number(btn.dataset.idx))));
    $$('.artreq-regen-btn').forEach(btn => btn.addEventListener('click', () => regenerateArtReqVariation(Number(btn.dataset.idx))));
  }

  $('#artreq-generate-btn').addEventListener('click', async () => {
    if (ideationSelectedIndex === -1) { toast('Select an Ideation theme as Final first.', 'error'); return; }
    const eventType = $('#ideation-event-type').value;
    const ideationTheme = ideationThemes[ideationSelectedIndex];
    let variationCount = parseInt($('#artreq-variation-count').value, 10);
    if (!Number.isFinite(variationCount)) variationCount = 3;
    variationCount = Math.max(1, Math.min(10, variationCount));
    $('#artreq-variation-count').value = variationCount;

    const btn = $('#artreq-generate-btn');
    btn.disabled = true; btn.textContent = `Generating ${variationCount} variation${variationCount === 1 ? '' : 's'}...`;
    log(`Generating ${variationCount} Art Requirement variation(s) for "${ideationTheme.theme_name}"...`);

    const res = await window.api.artReqGenerate({ eventType, ideationTheme, variationCount });
    btn.disabled = false; btn.textContent = 'Generate Art Requirement Variations';

    if (!res.success) { log(`Art Requirement generation failed: ${res.output}`, 'console-log-error'); toast('Generation failed: ' + res.output, 'error'); return; }
    try {
      const parsed = JSON.parse(res.stdout);
      if (parsed.error) { log(`Error: ${parsed.error}`, 'console-log-error'); toast(parsed.error, 'error'); return; }
      artReqVariations = parsed.variations;
      artReqDraftPaths = parsed.draft_paths || [];
      artReqSelectedIndex = -1;
      artReqSavedPath = '';
      renderArtReqResults();
      log(`Generated ${artReqVariations.length} Art Requirement variation(s) via ${parsed.provider === 'ollama' ? 'local Ollama' : 'Claude'}.`, 'console-log-success');
      toast(`${artReqVariations.length} variation(s) generated.`, 'success');
    } catch (e) {
      log(`Error parsing Art Requirement response: ${e.message}`, 'console-log-error');
      toast('Could not parse the generated variations.', 'error');
    }
  });

  async function regenerateArtReqVariation(idx) {
    const feedback = $(`#artreq-feedback-${idx}`).value.trim();
    if (!feedback) { toast('Enter feedback before regenerating.', 'error'); return; }
    const eventType = $('#ideation-event-type').value;
    const variation = artReqVariations[idx];

    const btn = document.querySelector(`.artreq-regen-btn[data-idx="${idx}"]`);
    if (btn) { btn.disabled = true; btn.textContent = 'Regenerating...'; }
    log(`Regenerating Art Requirement variation "${variation.title}" with feedback...`);

    const res = await window.api.artReqRegenerate({ eventType, variation, variationIndex: idx, feedback });

    if (!res.success) {
      log(`Regeneration failed: ${res.output}`, 'console-log-error');
      toast('Regeneration failed: ' + res.output, 'error');
      if (btn) { btn.disabled = false; btn.textContent = 'Submit Feedback & Regenerate'; }
      return;
    }
    try {
      const parsed = JSON.parse(res.stdout);
      if (parsed.error) { log(`Error: ${parsed.error}`, 'console-log-error'); toast(parsed.error, 'error'); return; }
      artReqVariations[idx] = parsed.variation;
      if (parsed.draft_path) artReqDraftPaths[idx] = parsed.draft_path;
      await refreshIdeationEntries(eventType);
      renderLessonsPanel();
      renderArtReqResults();
      log(`Variation "${parsed.variation.title}" regenerated. Feedback saved for future runs.`, 'console-log-success');
      toast('Variation regenerated — feedback stored for next time.', 'success');
    } catch (e) {
      log(`Error parsing regeneration response: ${e.message}`, 'console-log-error');
      toast('Could not parse the regenerated variation.', 'error');
    } finally {
      if (btn) { btn.disabled = false; btn.textContent = 'Submit Feedback & Regenerate'; }
    }
  }

  async function selectArtReqVariation(idx) {
    const eventType = $('#ideation-event-type').value;
    const variation = artReqVariations[idx];
    const refName = ideationThemes[ideationSelectedIndex].internal_reference_name;

    const btn = document.querySelector(`.artreq-select-btn[data-idx="${idx}"]`);
    if (btn) { btn.disabled = true; btn.textContent = 'Saving...'; }
    log(`Saving Art Requirement variation "${variation.title}" as final...`);

    const res = await window.api.artReqSaveFinal({ eventType, refName, variation });

    if (res.success) {
      artReqSelectedIndex = idx;
      artReqSavedPath = res.output;
      log(`Saved to ${res.output}`, 'console-log-success');
      toast('Saved as the final Art Requirement for this event.', 'success');
      renderArtReqResults();
    } else {
      log(`Could not save the final Art Requirement: ${res.output}`, 'console-log-error');
      toast('Save failed: ' + res.output, 'error');
      if (btn) { btn.disabled = false; btn.textContent = 'Select as Final'; }
    }
  }

  // ---------------- Sprint Board (landing page) ----------------
  // The roadmap is the plan; the design workbooks are the truth. This view
  // only ever reads — every fix is a jump into the tab that owns it.
  let sprintData = null;

  const SPRINT_TAB_VIEW = {
    minigames: 'minigames-view', trivia: 'trivia-challenge-view',
    battlepass: 'battlepass-view', holiday_bp: 'holiday-bp-view',
    decoration_shop: 'decoration-shop-view', streak_challenge: 'streak-challenge-view',
    weekly_gift: 'weekly-gift-view', full_screen_gift: 'full-screen-gift-view',
    endless_treasure: 'endless-treasure-view'
  };
  const STATUS_META = {
    done:   { cls: 'sp-done',   label: 'Configured' },
    todo:   { cls: 'sp-todo',   label: 'Needs configuring' },
    extra:  { cls: 'sp-extra',  label: 'Configured (not in roadmap)' },
    manual: { cls: 'sp-manual', label: 'Manual check' },
    'n/a':  { cls: 'sp-na',     label: 'Not this sprint' }
  };

  function fmtRange(a, b) {
    if (!a) return '—';
    const d = (s) => s.slice(5).replace('-', '/');
    return b && b !== a ? `${d(a)} – ${d(b)}` : d(a);
  }

  async function loadSprintBoard(selectedSprint) {
    const root = activeBranchRoot();
    const designPath = root ? await window.api.resolveDesignPath(root, cfg.activeWorkspace) : '';
    const res = await window.api.sprintBoard({ designPath, selectedSprint });
    if (!res.success) {
      $('#sprint-summary').textContent = res.output;
      $('#sprint-detail').innerHTML = '';
      log(`Sprint board: ${res.output}`, 'console-log-error');
      return;
    }
    sprintData = res.data;
    renderSprintPicker();
    renderSprintDetail();
  }

  function renderSprintPicker() {
    const sel = $('#sprint-select');
    sel.innerHTML = sprintData.sprints.map(s => {
      const tag = s.is_milestone ? ' — store submission' : '';
      const cur = s.name === sprintData.current_sprint ? '  ← current' : '';
      return `<option value="${escapeHtml(s.name)}"${s.name === sprintData.selected ? ' selected' : ''}>${escapeHtml(s.name)} · ${escapeHtml(fmtRange(s.start, s.end))}${tag}${cur}</option>`;
    }).join('');
    $('#sprint-summary').textContent = sprintData.current_sprint
      ? `Roadmap has ${sprintData.sprints.length} columns. Today falls in "${sprintData.current_sprint}".`
      : `Roadmap has ${sprintData.sprints.length} columns. No sprint covers today's date.`;
  }

  function renderSprintDetail() {
    const d = sprintData.detail, sp = d.sprint, el = $('#sprint-detail');
    if (sp.is_milestone) {
      el.innerHTML = `<div class="stacked-card">
        <h2 class="card-title">${escapeHtml(sp.name)}</h2>
        <p class="settings-hint">Store submission milestone — no LiveOps configuration here.</p>
        ${(sp.events || []).map(e => `<p class="settings-hint">• ${escapeHtml(e.title || e.text)}</p>`).join('')}
      </div>`;
      return;
    }

    const ev = d.event_config;
    let evHtml;
    if (!ev) {
      evHtml = `<p class="sp-drift">No event configured for this window in the active branch — the whole sprint is unstarted here.</p>`;
    } else if (!ev.exact) {
      evHtml = `<p class="sp-drift">Configured as <strong>${escapeHtml(ev.alias)}</strong> ${escapeHtml(fmtRange(ev.start, ev.end))} — the roadmap plans ${escapeHtml(fmtRange(sp.start, sp.end))}. Dates disagree; worth checking which is right.</p>`;
    } else {
      evHtml = `<p class="settings-hint">Configured as <strong>${escapeHtml(ev.alias)}</strong>, dates match the roadmap exactly.</p>`;
    }

    const todo = (d.tabs || []).filter(t => t.status === 'todo').length;
    const cards = (d.tabs || []).map(t => {
      const m = STATUS_META[t.status] || STATUS_META['n/a'];
      const rows = (t.rows || []).slice(0, 3)
        .map(r => `<div class="sp-row">${escapeHtml(r.name || '(unnamed)')} <span class="feedback-when">${escapeHtml(fmtRange(r.start, r.end))}${r.switch ? ' · ' + escapeHtml(r.switch) : ''}</span></div>`).join('');
      return `<div class="sp-card ${m.cls}">
        <div class="sp-card-head">
          <span class="sp-dot" aria-hidden="true"></span>
          <strong>${escapeHtml(t.label)}</strong>
          <button class="secondary-btn small sp-go" data-view="${SPRINT_TAB_VIEW[t.tab]}">Open</button>
        </div>
        <div class="sp-status">${escapeHtml(m.label)}${t.count ? ` · ${t.count} row${t.count === 1 ? '' : 's'}` : ''}</div>
        ${rows}
      </div>`;
    }).join('');

    el.innerHTML = `
      <div class="stacked-card">
        <div class="ideation-card-header">
          <div>
            <h2 class="card-title">${escapeHtml(sp.events && sp.events[0] ? sp.events[0].title : sp.name)}</h2>
            <span class="preview-label mono">${escapeHtml(sp.name)} · ${escapeHtml(fmtRange(sp.start, sp.end))} · ${escapeHtml(sp.method)}</span>
          </div>
          <span class="sp-count-pill${todo ? ' sp-count-warn' : ''}">${todo ? `${todo} to configure` : 'All configured'}</span>
        </div>
        ${evHtml}
        ${(sp.overlays || []).length ? `<p class="settings-hint">Also running across this sprint: ${sp.overlays.map(o => escapeHtml(o.title) + ' (' + escapeHtml(fmtRange(o.start, o.end)) + ')').join(', ')}</p>` : ''}
      </div>
      <div class="sp-grid">${cards}</div>
      ${(d.other_updates || []).length ? `<div class="stacked-card">
        <h2 class="card-title">Also in this sprint</h2>
        <p class="settings-hint">Roadmap items this tool doesn't configure — tracked here so the sprint's full scope is visible.</p>
        ${d.other_updates.map(o => `<div class="sp-row">• ${escapeHtml(o)}</div>`).join('')}
      </div>` : ''}`;

    $$('.sp-go').forEach(b => b.addEventListener('click', () => {
      if (b.dataset.view) navigateTo(b.dataset.view);
    }));
  }

  $('#sprint-select').addEventListener('change', (e) => loadSprintBoard(e.target.value));
  $('#sprint-refresh-btn').addEventListener('click', () => loadSprintBoard($('#sprint-select').value));

  // ---------------- Init ----------------
  async function init() {
    const data = await window.api.getConfig();
    cfg = data.cfg;
    branchKeys = data.branchKeys;
    workspaces = data.workspaces;
    defaults = data.defaults;
    pipelineSteps = data.pipelineSteps;

    $('#unity-exe-input').value = cfg.unityExePath || '';
    updateRobotPvpPathDisplay();
    $('#anthropic-key-input').value = cfg.anthropicApiKey || '';
    $('#ollama-base-url-input').value = cfg.ollamaBaseUrl || '';
    $('#ollama-model-input').value = cfg.ollamaModel || 'qwen3:32b';
    $('#author-name-input').value = cfg.authorName || '';
    $('#author-id-display').textContent = cfg.authorId || '';
    draftAvatarColor = cfg.avatarColor || '';
    draftAvatarImage = cfg.avatarImage || '';
    renderProfileEditor();
    await refreshIdeationDataStatus();

    populateBranchWorkspaceSelects();
    renderBranchPathRows();
    buildChecklist();
    await refreshContext();
    await refreshSettingsResolvedPath();
    await loadEventTypes();
    updateTimesPreview();

    navigateTo('sprint-view');
    log('Tool initialized. Waiting for input...');
  }

  init();
})();
