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
  let cachedChapter = null;
  let ideationThemes = null;
  let ideationDraftPaths = [];
  let ideationSelectedIndex = -1;
  let ideationSavedPath = '';
  let artReqVariations = null;
  let artReqDraftPaths = [];
  // Per-variation rendered Markdown from the backend. undefined = not fetched
  // yet (card shows "Loading…"), null = the render failed.
  let artReqPreviews = [];
  let artReqSelectedIndex = -1;
  let artReqSavedPath = '';
  // A One-Pager the user already had, picked off disk instead of generated here.
  // { name, path, text } once chosen. The Art Requirement step accepts either
  // this or a theme selected as Final, never both at once.
  let ideationOnePagers = [];
  let ideationPickedOnePager = null;
  // Presence: which event type this machine has currently announced, so a type
  // switch can check out of the old one before checking in to the new.
  let presenceTimer = null;
  let presenceAnnouncedType = '';
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
    // Presence is scoped to the Ideation view. Announcing while someone is off
    // in Settings would report them as working on an event they are not on.
    if (viewId === 'ideation-view') startPresence(); else stopPresence();
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
    if (viewId === 'chapter-view') loadChapter();
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
    else if (currentView === 'chapter-view') await loadChapter(true);
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

  // Chapter Update tool paths. Folder pickers, persisted like unityExePath.
  for (const [btn, input, key] of [
    ['#browse-chapter-lvlpy-btn', '#chapter-lvlpy-input', 'chapterLvlPyPath'],
    ['#browse-chapter-mapbuilder-btn', '#chapter-mapbuilder-input', 'chapterMapBuilderPath'],
  ]) {
    $(btn).addEventListener('click', async () => {
      const picked = await window.api.browseFolder(cfg[key] || '');
      if (picked) {
        $(input).value = picked;
        cfg[key] = picked;
        await window.api.saveConfig({ [key]: picked });
        toast('Saved. Blank the field to go back to the guessed layout.', 'success');
      }
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

  $('#comfy-url-input').addEventListener('change', async (e) => {
    cfg.comfyUrl = e.target.value.trim();
    await window.api.saveConfig({ comfyUrl: cfg.comfyUrl });
  });
  $('#comfy-test-btn').addEventListener('click', async () => {
    const resultEl = $('#comfy-test-result');
    const btn = $('#comfy-test-btn');
    btn.disabled = true; btn.textContent = 'Testing...';
    resultEl.textContent = '';
    resultEl.classList.remove('ok', 'error');

    const res = await window.api.comfyCheck();
    btn.disabled = false; btn.textContent = 'Test Connection';
    let info = null;
    try { info = JSON.parse(res.stdout); } catch (e) { /* fall through to raw output */ }
    // 'wedged' is reported as its own state on purpose: the server answers
    // status requests in milliseconds while refusing all work, so a plain
    // reachable/unreachable readout would call a dead worker healthy.
    const okState = info && info.state === 'ok';
    resultEl.textContent = info ? info.detail : (res.output || 'No response.');
    resultEl.classList.add(okState ? 'ok' : 'error');
    log(`ComfyUI check: ${info ? info.state : 'error'} — ${info ? info.detail : res.output}`,
        okState ? 'console-log-line' : 'console-log-error');
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

  // ---------------- Chapter Update tab ----------------
  // Chapter numbers are DETECTED, never typed: pve_chapter.xlsx's chapter sheet
  // carries the internal map id (col A, 88000) and the display number (col C,
  // 78) side by side, so the offset between them is a fact from the data rather
  // than a setting that can drift. Everything downstream keys off that.
  const CHAPTER_STEP_DEFAULT = ['validate', 'reward', 'art', 'asset', 'qa'];

  async function loadChapter(force) {
    if (cachedChapter && !force) { renderChapter(); return; }
    cachedChapter = null;
    const root = activeBranchRoot();
    if (!root) { log('Cannot load Chapter info: branch path not set.', 'console-log-error'); return; }
    const designPath = await window.api.resolveDesignPath(root, cfg.activeWorkspace);
    log('Reading current chapter from pve_chapter.xlsx...');
    const res = await window.api.chapterGet({ designPath });
    if (!res.success) {
      log(`Error loading Chapter info: ${res.output}`, 'console-log-error');
      toast('Could not read pve_chapter.xlsx for this branch/workspace.', 'error');
      return;
    }
    try {
      cachedChapter = JSON.parse(res.stdout);
      if (cachedChapter.error) {
        log(`Error: ${cachedChapter.error}`, 'console-log-error');
        toast(cachedChapter.error, 'error');
        cachedChapter = null;
        return;
      }
      log(`Chapter info loaded. Current chapter ${cachedChapter.display_old} -> new ${cachedChapter.display_new} (map id ${cachedChapter.new_map_id}).`);
      renderChapter();
    } catch (e) {
      log(`Error parsing Chapter JSON: ${e.message}`, 'console-log-error');
    }
  }

  function renderChapter() {
    updateChapterSimDisplay();
    if (!cachedChapter) return;
    const c = cachedChapter;
    $('#chapter-cur').textContent = `章节${c.display_old} (internal ${c.old_chapter})`;
    $('#chapter-new').textContent = `章节${c.display_new} (internal ${c.new_chapter})`;
    $('#chapter-mapid').textContent = c.new_map_id;
    $('#chapter-range').textContent = `${c.level_range[0]} – ${c.level_range[1]}`;

    // Rebuilt from the backend's own step list so the two can't drift apart.
    // Level Balancing stays unticked by default: it needs a finished simulation
    // run, which won't exist at the point the config steps are done. QA
    // Validation is ticked — it only reads, and it is the sign-off gate.
    // Steps with an external prerequisite say so inline, and Magic Builder is
    // disabled outright when there is no system Python for conf_fill_ex.py —
    // better than letting someone tick it and fail two minutes in.
    $('#chapter-steps').innerHTML = c.steps.map(s => {
      const noPy = s.key === 'magic' && c.has_system_python === false;
      const off = noPy ? ' pass-option-off' : '';
      const need = noPy ? 'No system Python found — install Python 3.9+' : (s.needs || '');
      return `
      <label class="pass-option${off}" title="${escapeHtml(need)}">
        <input type="checkbox" class="chapter-step" value="${s.key}"
          ${CHAPTER_STEP_DEFAULT.includes(s.key) ? 'checked' : ''} ${noPy ? 'disabled' : ''} />
        <span>${escapeHtml(s.label)}</span>
        ${need ? `<span class="pass-count">${escapeHtml(need.split('(')[0].trim().slice(0, 34))}</span>` : ''}
      </label>`;
    }).join('');
  }

  // The simulation folder is a machine-level tool path, so it persists in
  // config like robotCalcPath / unityExePath rather than being retyped per run.
  // Blank means "derive from the active branch", which is right on every normal
  // checkout — the box shows the resolved path either way so it is never a
  // guess what Level Balancing will read.
  function chapterSimDefault() {
    const root = activeBranchRoot();
    // Built with an explicit separator: an unescaped \d in a template
    // literal is silently swallowed by JS.
    return root ? [root, 'develop', 'tools', 'RunCardEx2'].join('\\') : '';
  }

  function updateChapterSimDisplay() {
    const el = $('#chapter-sim-path');
    if (cfg.chapterSimPath) {
      el.value = cfg.chapterSimPath;
    } else {
      const d = chapterSimDefault();
      el.value = d ? `${d}   (branch default)` : '';
    }
  }

  $('#chapter-sim-browse-btn').addEventListener('click', async () => {
    const picked = await window.api.browseFolder(cfg.chapterSimPath || chapterSimDefault());
    if (picked) {
      cfg.chapterSimPath = picked;
      updateChapterSimDisplay();
      await window.api.saveConfig({ chapterSimPath: picked });
      toast('Simulation folder saved.', 'success');
    }
  });

  $('#chapter-sim-reset-btn').addEventListener('click', async () => {
    cfg.chapterSimPath = '';
    updateChapterSimDisplay();
    await window.api.saveConfig({ chapterSimPath: '' });
    toast('Reverted to the active branch default.', 'success');
  });

  $('#chapter-refresh-btn').addEventListener('click', () => loadChapter(true));

  $('#chapter-run-btn').addEventListener('click', async () => {
    if (!cachedChapter) { toast('Chapter metadata still loading...', 'error'); return; }
    const steps = $$('#chapter-steps .chapter-step').filter(b => b.checked).map(b => b.value);
    if (!steps.length) { toast('Tick at least one step to run.', 'error'); return; }

    const c = cachedChapter;
    const ok = confirm(
      `Roll chapter ${c.display_old} forward to ${c.display_new}?\n\n` +
      `Steps: ${steps.join(', ')}\n` +
      `Workspace: ${cfg.activeWorkspace}\n\n` +
      `This edits the design workbooks in place. Make sure they are not open in ` +
      `Excel and that you can svn revert if something looks wrong.`
    );
    if (!ok) return;

    const root = activeBranchRoot();
    const btn = $('#chapter-run-btn');
    btn.disabled = true; btn.textContent = 'Running...';
    const res = await window.api.chapterRun({
      branchRoot: root,
      payload: {
        workspace: cfg.activeWorkspace,
        old_chapter: c.old_chapter,
        display_offset: c.display_offset,
        steps,
        sim_path: cfg.chapterSimPath || '',
        lvl_py: cfg.chapterLvlPyPath || '',
        map_builder: cfg.chapterMapBuilderPath || '',
        scene_file: ''
      }
    });
    btn.disabled = false; btn.textContent = 'Run Chapter Update';

    if (res.success) {
      log(`Chapter Update complete — 章节${c.display_new} configured.`, 'console-log-success');
      toast(`Chapter ${c.display_new} configured. Review the workbooks before committing.`, 'success');
      cachedChapter = null;
      await loadChapter(true);
    } else {
      log(`Chapter Update failed (exit ${res.code}).`, 'console-log-error');
      toast('Chapter Update failed — see the output pane.', 'error');
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
      artReqPreviews = [];
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
    // The One-Pager list is per type too, and a document picked for Merge must
    // not still be armed after switching to Blocks.
    clearPickedOnePager();
    if (ideationMode() === 'existing') await loadOnePagers();
    await loadArtReqDocs();
    // Check out of the type just left before checking in to the new one, or the
    // old record lingers for its full timeout saying you are still on Merge.
    if (presenceAnnouncedType && presenceAnnouncedType !== e.target.value) {
      await window.api.presenceClear({ eventType: presenceAnnouncedType });
      presenceAnnouncedType = '';
    }
    await presenceBeat();
  });

  // ---- presence: who else has this event open ------------------------------
  // Checking in every 30s against main.js's 5-minute timeout, so ten writes can
  // be lost to a slow OneDrive sync before a working colleague drops off the
  // banner. Nothing here blocks anything - it answers "is someone else already
  // on this?", which is the question that actually costs a duplicated afternoon.
  const PRESENCE_BEAT_MS = 30 * 1000;

  function presenceDetail() {
    if (ideationPickedOnePager) return ideationPickedOnePager.name;
    if (ideationSavedPath) return 'finalised a One-Pager';
    if (ideationThemes) return 'generating themes';
    return '';
  }

  async function presenceBeat() {
    const eventType = $('#ideation-event-type').value;
    if (!eventType) return;
    await window.api.presenceAnnounce({ eventType, detail: presenceDetail() });
    presenceAnnouncedType = eventType;
    await renderPresence(eventType);
  }

  async function renderPresence(eventType) {
    const bar = $('#ideation-presence');
    if (!bar) return;
    const res = await window.api.presenceList({ eventType });
    const others = (res && res.others) || [];
    if (!others.length) { bar.style.display = 'none'; bar.innerHTML = ''; return; }
    const mins = (ms) => {
      const m = Math.floor(ms / 60000);
      if (m < 1) return 'just now';
      if (m === 1) return 'for a minute';
      if (m < 60) return `for ${m} minutes`;
      const h = Math.floor(m / 60);
      return h === 1 ? 'for over an hour' : `for over ${h} hours`;
    };
    const who = others.map(o => {
      const dot = o.avatarColor
        ? `<span class="presence-dot" style="background:${escapeHtml(o.avatarColor)}"></span>`
        : '<span class="presence-dot"></span>';
      const detail = o.detail ? ` <span class="presence-detail">— ${escapeHtml(o.detail)}</span>` : '';
      return `<span class="presence-who">${dot}<span class="presence-name">${escapeHtml(o.authorName)}</span>${detail}`
        + ` <span class="presence-since">${escapeHtml(mins(o.sinceMs))}</span></span>`;
    }).join('');
    const lead = others.length === 1 ? 'is also working on' : 'are also working on';
    bar.innerHTML = `${who} <span class="presence-detail">${lead} ${escapeHtml(eventType)}.</span>`;
    bar.style.display = 'flex';
  }

  async function startPresence() {
    if (presenceTimer) return;
    await presenceBeat();
    presenceTimer = setInterval(presenceBeat, PRESENCE_BEAT_MS);
  }

  async function stopPresence() {
    if (presenceTimer) { clearInterval(presenceTimer); presenceTimer = null; }
    const bar = $('#ideation-presence');
    if (bar) { bar.style.display = 'none'; bar.innerHTML = ''; }
    if (presenceAnnouncedType) {
      const was = presenceAnnouncedType;
      presenceAnnouncedType = '';
      await window.api.presenceClear({ eventType: was });
    }
  }

  // ---- starting point: generate themes, or use a One-Pager you already have --

  function ideationMode() {
    return $('#ideation-mode-existing').checked ? 'existing' : 'new';
  }

  function clearPickedOnePager() {
    ideationPickedOnePager = null;
    renderOnePagerList();
  }

  function applyIdeationMode() {
    const existing = ideationMode() === 'existing';
    $('#ideation-new-card').style.display = existing ? 'none' : '';
    $('#ideation-existing-card').style.display = existing ? '' : 'none';
    // Results from the other route are stale the moment you switch, and leaving
    // them on screen makes it ambiguous which source the Art Requirement will
    // actually be built from.
    $('#ideation-results').innerHTML = '';
    $('#artreq-results').innerHTML = '';
    ideationThemes = null;
    ideationDraftPaths = [];
    ideationSelectedIndex = -1;
    ideationSavedPath = '';
    artReqVariations = null;
    artReqDraftPaths = [];
    artReqPreviews = [];
    artReqSelectedIndex = -1;
    artReqSavedPath = '';
    ideationPickedOnePager = null;
    $('#artreq-trigger-card').style.display = 'none';
    if (existing) loadOnePagers();
  }

  $$('input[name="ideation-mode"]').forEach(r => {
    r.addEventListener('change', applyIdeationMode);
  });
  $('#ideation-refresh-onepagers-btn').addEventListener('click', () => loadOnePagers());
  $('#ideation-upload-onepager-btn').addEventListener('click', uploadOnePager);

  async function uploadOnePager() {
    const eventType = $('#ideation-event-type').value;
    const btn = $('#ideation-upload-onepager-btn');
    btn.disabled = true;
    const res = await window.api.ideationUploadOnePager({ eventType });
    btn.disabled = false;
    if (res.canceled) return;
    if (!res.success) {
      log(`Upload failed: ${res.output}`, 'console-log-error');
      toast(res.output, 'error');
      return;
    }
    await loadOnePagers();
    log(`Filed "${res.name}" into ${eventType}/One-Pager Ideation.`, 'console-log-success');
    // Every accepted format is usable - Word and PDF are converted on the way
    // through - so uploading always selects. There is no "filed but useless"
    // outcome to report any more.
    const idx = ideationOnePagers.findIndex(d => d.path === res.path);
    if (idx >= 0) await pickOnePager(idx);
    else toast('Filed — pick it from the list below.', 'success');
  }

  async function loadOnePagers() {
    const eventType = $('#ideation-event-type').value;
    const box = $('#ideation-onepager-list');
    box.innerHTML = '<div class="onepager-empty">Looking…</div>';
    const res = await window.api.ideationListOnePagers({ eventType });
    if (!res.success) {
      box.innerHTML = '';
      log(`Could not list One-Pagers: ${res.output}`, 'console-log-error');
      toast('Could not read the One-Pager folder: ' + res.output, 'error');
      return;
    }
    ideationOnePagers = res.docs || [];
    renderOnePagerList();
  }

  function renderOnePagerList() {
    const box = $('#ideation-onepager-list');
    if (!box) return;
    const eventType = $('#ideation-event-type').value;
    if (!ideationOnePagers.length) {
      box.innerHTML = `<div class="onepager-empty">No One-Pager found for ${escapeHtml(eventType)}. `
        + `Save one into its <code>One-Pager Ideation</code> folder, or switch to "Start from scratch".</div>`;
      return;
    }
    box.innerHTML = ideationOnePagers.map((d, i) => {
      const picked = ideationPickedOnePager && ideationPickedOnePager.path === d.path;
      const when = new Date(d.modified).toLocaleString();
      const cls = ['onepager-row', picked ? 'is-picked' : '', d.readable ? '' : 'is-unreadable']
        .filter(Boolean).join(' ');
      // Every format is usable now; a Word or PDF brief is converted to Markdown
      // on first use. The label says so, because "Use this" on a PDF looks like
      // a bug until you know the conversion happens.
      const ext = (d.name.split('.').pop() || '').toUpperCase();
      const label = picked ? 'Picked' : (d.convertible ? `Convert & use` : 'Use this');
      const note = d.convertible ? `<span class="onepager-meta">${escapeHtml(ext)}</span>` : '';
      const action = `${note}<button class="secondary-btn small onepager-pick-btn" data-idx="${i}">${label}</button>`;
      return `<div class="${cls}">
          <span class="onepager-name">${escapeHtml(d.name)}</span>
          <span class="onepager-meta">${escapeHtml(when)}</span>
          ${action}
        </div>`;
    }).join('');
    $$('.onepager-pick-btn').forEach(b => {
      b.addEventListener('click', () => pickOnePager(Number(b.dataset.idx)));
    });
  }

  async function pickOnePager(idx) {
    const doc = ideationOnePagers[idx];
    if (!doc) return;
    // A Word or PDF brief is converted on the way through, which spawns Python
    // and takes a moment - so say what is happening rather than leaving a dead
    // button. Conversion is cached beside the original, so this is once per edit.
    const btn = document.querySelector(`.onepager-pick-btn[data-idx="${idx}"]`);
    if (btn) { btn.disabled = true; btn.textContent = doc.convertible ? 'Converting…' : 'Reading…'; }
    if (doc.convertible) log(`Converting ${doc.name} to Markdown…`);

    const res = await window.api.ideationReadOnePager({ docPath: doc.path });
    if (!res.success) {
      renderOnePagerList();                 // restore the button
      log(`Could not use ${doc.name}: ${res.output}`, 'console-log-error');
      toast(res.output, 'error');
      return;
    }
    ideationPickedOnePager = { name: doc.name, path: doc.path, text: res.text,
                               // the .md actually read - a converted brief's
                               // reference art sits beside it
                               docPath: res.docPath || doc.path };
    renderOnePagerList();
    $('#artreq-trigger-card').style.display = 'block';
    artReqVariations = null;
    artReqDraftPaths = [];
    artReqPreviews = [];
    artReqSelectedIndex = -1;
    artReqSavedPath = '';
    $('#artreq-results').innerHTML = '';
    if (res.converted) {
      // Naming the .md matters: it is the file the model was actually given, and
      // it is editable if the conversion got something wrong.
      const st = res.convertStats || {};
      let imgNote = '';
      if (st.images) {
        // The images are the part worth reporting. In the shipped Merge2608
        // brief the zone contents exist only inside mood-board pictures, so
        // "read text from N of them" is the line that says whether the brief
        // actually made it through.
        imgNote = `, ${st.images} image${st.images === 1 ? '' : 's'} extracted`;
        if (st.ocrLines) imgNote += ` (read text from ${st.ocrImages} of them)`;
      }
      log(`Converted to ${res.converted} — ${res.text.length.toLocaleString()} characters`
        + `${imgNote}. That file is what the model is given, so edit it if the `
        + `conversion missed something.`, 'console-log-success');
      toast(`Converted ${doc.name} to Markdown${st.images ? ` + ${st.images} images` : ''} `
        + `— generate variations below.`, 'success');
    } else {
      log(`Using "${doc.name}" as the approved One-Pager.`, 'console-log-success');
      toast('One-Pager selected — generate Art Requirement variations below.', 'success');
    }
  }

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
      artReqPreviews = [];
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

    // The real document, straight from the Python writer (fetched into
    // artReqPreviews by loadArtReqPreview). Until it arrives, say so rather
    // than falling back to a second hand-built rendering of the same doc -
    // that divergence is exactly what this replaced.
    const md = artReqPreviews[idx];
    let docHtml, sections = [];
    if (md === undefined) {
      docHtml = `<p class="settings-hint">Loading the document…</p>`;
    } else if (md === null) {
      docHtml = `<p class="settings-hint">Could not render this document — see the log.</p>`;
    } else {
      const rendered = renderMarkdownDoc(md, idx);
      docHtml = rendered.html;
      sections = rendered.sections;
    }

    const sectionFeedbackHtml = sections.length ? `
      <details class="info-disclosure" open>
        <summary>
          <span class="info-disclosure-icon" aria-hidden="true"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"/></svg></span>
          <span>Feedback by section</span>
          <svg class="info-disclosure-chevron" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M6 9l6 6 6-6"/></svg>
        </summary>
        <div class="info-disclosure-body">
          <p class="settings-hint">A note here is tagged with its section, so the rewrite knows exactly what it refers to. Leave the rest blank.</p>
          ${sections.map(s => `
            <div class="artreq-section-fb-row">
              <label class="artreq-section-fb-label" for="artreq-sfb-${idx}-${escapeHtml(s.id)}">
                <a href="#${escapeHtml(s.id)}" class="artreq-section-jump">${escapeHtml(s.title)}</a>
              </label>
              <textarea id="artreq-sfb-${idx}-${escapeHtml(s.id)}" class="text-input ideation-feedback artreq-section-fb"
                data-idx="${idx}" data-section="${escapeHtml(s.title)}" rows="1"
                placeholder="Notes on this section…"></textarea>
            </div>`).join('')}
        </div>
      </details>` : '';

    return `
      <div class="stacked-card ideation-card${selected ? ' ideation-selected' : ''}" data-idx="${idx}">
        <div class="ideation-card-header">
          <div><h2 class="card-title">${escapeHtml(v.title)}</h2></div>
          <button class="secondary-btn small artreq-select-btn" data-idx="${idx}">${selected ? 'Selected ✓' : 'Select as Final'}</button>
        </div>
        ${artReqDraftPaths[idx] ? `<p class="settings-hint ideation-saved-path"><strong>Draft document:</strong> <span class="mono">${escapeHtml(artReqDraftPaths[idx])}</span></p>` : ''}
        ${selected && artReqSavedPath ? `<p class="settings-hint ideation-saved-path"><strong>Saved as final:</strong> <span class="mono">${escapeHtml(artReqSavedPath)}</span></p>` : ''}
        ${artReqExportRow(idx, selected)}

        <details class="info-disclosure artreq-doc-disclosure" open>
          <summary>
            <span class="info-disclosure-icon" aria-hidden="true"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><path d="M14 2v6h6M8 13h8M8 17h8"/></svg></span>
            <span>Document preview — exactly what gets written to the .md</span>
            <svg class="info-disclosure-chevron" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M6 9l6 6 6-6"/></svg>
          </summary>
          <div class="info-disclosure-body">
            <div class="artreq-doc">${docHtml}</div>
          </div>
        </details>

        ${sectionFeedbackHtml}

        ${threadHtml('artreq', artReqRefName())}

        <div class="board-divider"></div>
        <h3 class="card-title">Reference Images</h3>
        ${selected ? renderZoneImages(v, idx) : `<p class="settings-hint">Select this variation as final before generating reference images — avoids spending render time on a draft that might still change.</p>`}

        <div class="board-divider"></div>
        <div class="form-row">
          <label class="form-label" for="artreq-feedback-${idx}">Document Feedback</label>
          <textarea id="artreq-feedback-${idx}" class="text-input ideation-feedback" rows="2" placeholder="Changes to the WRITTEN document. Image notes go under each image above."></textarea>
        </div>
        <div class="execution-footer" style="justify-content:flex-end;">
          <button class="primary-btn small artreq-regen-btn" data-idx="${idx}">Submit Feedback &amp; Regenerate</button>
        </div>
      </div>
    `;
  }

  // ---- Document preview (the real .md, rendered) ----
  // Deliberately a small renderer for exactly what _render_markdown emits -
  // ATX headings, pipe tables, -/1. lists, bold/italic, inline images and
  // code - rather than a general Markdown engine. Everything is escaped
  // before any markup is added, so document text can never inject HTML.
  function mdInline(s) {
    return escapeHtml(s)
      .replace(/!\[([^\]]*)\]\(([^)]+)\)/g, (m, alt, src) =>
        `<img class="artreq-doc-img" src="${fileUrl(decodeURIComponent(src))}" alt="${alt}" />`)
      .replace(/\*\*([^*]+)\*\*/g, '<strong>$1</strong>')
      .replace(/(^|[^*])\*([^*]+)\*/g, '$1<em>$2</em>')
      .replace(/_([^_]+)_/g, '<em>$1</em>')
      .replace(/`([^`]+)`/g, '<code>$1</code>');
  }

  function mdTableRow(line) {
    const cells = line.trim().replace(/^\||\|$/g, '').split('|');
    return cells.map(c => c.trim().replace(/\\\|/g, '|'));
  }

  // Returns { html, sections:[{id,title}] } - the section list drives the
  // per-section feedback controls, so they always match the document's real
  // headings instead of a hardcoded list that can fall out of step with it.
  function renderMarkdownDoc(md, idx) {
    const lines = String(md || '').split('\n');
    const out = [];
    const sections = [];
    let i = 0;
    let listOpen = null;

    const closeList = () => { if (listOpen) { out.push(`</${listOpen}>`); listOpen = null; } };

    while (i < lines.length) {
      const line = lines[i];

      const isTable = /^\s*\|/.test(line) && /^\s*\|?[\s:-]*-[\s:|-]*$/.test(lines[i + 1] || '');
      if (isTable) {
        closeList();
        const headers = mdTableRow(line);
        i += 2;
        const body = [];
        while (i < lines.length && /^\s*\|/.test(lines[i])) { body.push(mdTableRow(lines[i])); i++; }
        out.push(`<div class="artreq-doc-tablewrap"><table class="mg-preview-table"><thead><tr>${
          headers.map(h => `<th>${mdInline(h)}</th>`).join('')}</tr></thead><tbody>${
          body.map(r => `<tr>${r.map(c => `<td>${mdInline(c)}</td>`).join('')}</tr>`).join('')
        }</tbody></table></div>`);
        continue;
      }

      const h = /^(#{1,6})\s+(.*)$/.exec(line);
      if (h) {
        closeList();
        const level = h[1].length;
        const text = h[2].trim();
        // Only the numbered template sections get a feedback control; the
        // title and deeper sub-headings would just add noise.
        const isSection = level === 2;
        if (isSection) {
          const id = `artreq-sec-${idx}-${sections.length}`;
          sections.push({ id, title: text });
          out.push(`<h${level} class="artreq-doc-h artreq-doc-section" id="${id}">${mdInline(text)}</h${level}>`);
        } else {
          out.push(`<h${level} class="artreq-doc-h">${mdInline(text)}</h${level}>`);
        }
        i++; continue;
      }

      const li = /^\s*([-*]|\d+\.)\s+(.*)$/.exec(line);
      if (li) {
        const want = /^\d/.test(li[1]) ? 'ol' : 'ul';
        if (listOpen !== want) { closeList(); out.push(`<${want} class="artreq-doc-list">`); listOpen = want; }
        out.push(`<li>${mdInline(li[2])}</li>`);
        i++; continue;
      }

      if (/^\s*---+\s*$/.test(line)) { closeList(); out.push('<hr class="artreq-doc-hr" />'); i++; continue; }

      if (!line.trim()) { closeList(); i++; continue; }

      closeList();
      out.push(`<p class="artreq-doc-p">${mdInline(line)}</p>`);
      i++;
    }
    closeList();
    return { html: out.join('\n'), sections };
  }

  // Reference images live under each variation. Image feedback is deliberately
  // separate from document feedback: a note about the art should regenerate the
  // art, not rewrite the spec, and the two accumulate as different lesson kinds.
  function renderZoneImages(v, idx) {
    const zones = v.zones || [];
    const withImages = zones.filter(z => z.image_path);
    if (!withImages.length) {
      return `
        <p class="settings-hint">No reference images yet. Generating renders one image per zone
        through ComfyUI, grounded in the approved render spec and this event type's own shipped
        art documents.</p>
        ${passPickerHtml(idx, artReqVariations[idx])}
        <div class="execution-footer" style="justify-content:flex-end;">
          <button class="secondary-btn small artreq-images-btn" data-idx="${idx}">Generate Reference Images</button>
        </div>`;
    }
    const cards = withImages.map(z => `
      <div class="zone-image">
        <img class="zone-image-thumb" src="${fileUrl(z.image_path)}" alt="Zone ${escapeHtml(String(z.zone_number))}" />
        <div class="zone-image-body">
          <div class="zone-image-head"><strong>Zone ${escapeHtml(String(z.zone_number))}</strong>
            <span class="feedback-when mono">${escapeHtml(z.zone_theme || '')}</span></div>
          <textarea id="zoneimg-fb-${idx}-${z.zone_number}" class="text-input ideation-feedback" rows="2"
            placeholder="e.g. the press should be teal, not red"></textarea>
          <div class="zone-image-actions">
            <label class="zone-image-all">
              <input type="checkbox" id="zoneimg-all-${idx}-${z.zone_number}" />
              Apply to all images
            </label>
            <button class="secondary-btn small zoneimg-regen-btn"
              data-idx="${idx}" data-zone="${escapeHtml(String(z.zone_number))}">Regenerate</button>
          </div>
        </div>
      </div>`).join('');
    return `<div class="zone-image-grid">${cards}</div>
      <div class="execution-footer" style="justify-content:flex-end;">
        <button class="secondary-btn small artreq-images-btn" data-idx="${idx}">Regenerate All From Scratch</button>
      </div>`;
  }

  // file:// URL for a local render. Electron will not load a bare Windows path
  // from an <img src>, and the path can contain spaces and non-ASCII.
  function fileUrl(p) {
    if (!p) return '';
    return 'file:///' + String(p).replace(/\\/g, '/').split('/').map(encodeURIComponent).join('/');
  }


  // Which asset passes to render. Scene is on by default and is the only one an
  // older document can satisfy; the rest need the buildings/producers the newer
  // schema carries, so each is disabled until the document actually has them.
  function passPickerHtml(idx, doc) {
    const zones = (doc && doc.zones) || [];
    const nb = zones.reduce((a, z) => a + ((z.buildings || []).length), 0);
    const np = ((doc && doc.producers) || []).length;
    // Merge carries producers and no recipes; Dessert the reverse. Each pass is
    // gated on what the document actually has, so a type is never offered the
    // other's asset classes.
    const ns = zones.reduce((a, z) => a +
      (z.buildings || []).reduce((b, x) => b + Math.max(0, (Number(x.max_level) || 1) - 1), 0), 0);
    const nr = ((doc && doc.recipes) || []).length;
    // A Blocks scene ships as a build sequence: derelict + empty + one frame per
    // construction step + characters. Counted here so the cost line is honest -
    // five scenes of six steps is 45 renders, not five.
    // Only Blocks delivers a scene as a build sequence; other types use
    // construction_steps as free-text reveal notes.
    const evType = (document.querySelector('#ideation-event-type') || {}).value || '';
    const np2 = evType !== 'Blocks' ? 0 : zones.reduce((a, z) => {
      const st = (z.construction_steps || []).filter(x => String(x).trim());
      return a + (st.length >= 2 ? st.length + 2 + Math.min(2, ((doc && doc.characters) || []).length) : 0);
    }, 0);
    // Sort goods live on the ZONE - one themed catalogue per floor.
    const ng = zones.reduce((a, z) => a + ((z.goods || []).length), 0);
    const nm = ((doc && doc.construction_materials) || []).length;
    const nc = ((doc && doc.characters) || []).length;
    const ni = ((doc && doc.ingredients) || []).length;
    const ex = (doc && doc.event_extras) || {};
    const nx = (String(ex.hot_air_balloon || '').trim() ? 1 : 0)
      + (String(ex.achievement_badge || '').trim() ? 1 : 0)
      + ((ex.task_icons || []).filter(t => String(t).trim()).length);
    const rows = [
      ['scene', 'Scene overviews', zones.length, true],
      ['building', 'Building sheets', nb, nb > 0],
      // Gated on the stage count, not the building count: a Dessert document has
      // 46 buildings and every one of them is max_level 1, so gating on nb left
      // a live checkbox that would render nothing.
      ['stage', 'Upgrade stages', ns, ns > 0],
      ['icon', 'Chain item icons', ((doc && doc.producers) || []).reduce((a, pr) =>
        a + (pr.chains || []).reduce((b, c) => b + ((c.items || []).length), 0), 0), np > 0],
      ['producer', 'Producer icons', np * 2, np > 0],
      ['recipe', 'Recipes', nr, nr > 0],
      ['ingredient', 'Ingredients', ni, ni > 0],
      ['progression', 'Scene build sequences', np2, np2 > 0],
      ['goods', 'Goods icons', ng, ng > 0],
      ['material', 'Construction materials', nm, nm > 0],
      ['character', 'Characters', nc, nc > 0],
      ['extras', 'Balloon, badge, task icons', nx, nx > 0]
    ];
    const boxes = rows.map(([key, label, count, enabled]) => `
      <label class="pass-option${enabled ? '' : ' pass-option-off'}">
        <input type="checkbox" class="artreq-pass" data-idx="${idx}" value="${key}"
          ${key === 'scene' ? 'checked' : ''} ${enabled ? '' : 'disabled'} />
        <span>${label}</span><span class="pass-count">${count}</span>
      </label>`).join('');
    return `
      <div class="pass-picker" data-idx="${idx}">
        ${boxes}
        <div class="pass-total" id="artreq-plan-${idx}">Scene overviews only — ${zones.length} render(s).</div>
      </div>`;
  }

  // Seconds per render, measured across this session's runs. Mirrors the same
  // constant in artreq_images.plan_main; kept here so the cost line needs no
  // round trip.
  const SECONDS_PER_RENDER = 32;

  // Deliberately synchronous. This used to call artreq:plan on every tick, and
  // driving it in the real app showed the cost line running TWO clicks behind:
  // each tick spawned a Python process taking ~2s, and nothing guarded against
  // responses arriving out of order, so a stale reply could land last and stick.
  // The counts are already on the page - passPickerHtml rendered them into the
  // .pass-count pills - so summing the ticked ones is exact, instant, and has no
  // race to guard. artreq:plan remains the authoritative planner and is logged
  // once before a run actually starts.
  function refreshPlan(idx) {
    const el = $(`#artreq-plan-${idx}`);
    if (!el) return;
    let total = 0, chosen = 0;
    $$(`.pass-picker[data-idx="${idx}"] .pass-option`).forEach(o => {
      const cb = o.querySelector('input');
      if (cb && cb.checked) {
        total += Number(o.querySelector('.pass-count').textContent) || 0;
        chosen++;
      }
    });
    if (!chosen) { el.textContent = 'Nothing selected.'; return; }
    const mins = Math.round(total * SECONDS_PER_RENDER / 60);
    el.textContent = `${total} render(s) — roughly ${mins} minute${mins === 1 ? '' : 's'}.`;
  }

  // Images now only ever run on the already-finalized variation (the
  // Reference Images section is hidden until Select as Final), so every
  // successful render/regeneration needs to update the .md that was already
  // written to disk - save_final_main() just re-renders whatever variation
  // object it's given, so calling it again with the image-populated
  // variation is exactly how an already-finalized doc picks up its images.
  async function resaveArtReqFinal(idx) {
    const refName = artReqRefName();
    if (!refName) return;
    const res = await window.api.artReqSaveFinal({
      eventType: $('#ideation-event-type').value,
      refName,
      variation: artReqVariations[idx]
    });
    if (res.success) {
      artReqSavedPath = res.output;
      log(`Final document updated with the new image(s): ${res.output}`, 'console-log-success');
    } else {
      log(`Images were generated, but updating the saved final document failed: ${res.output}`, 'console-log-error');
      toast('Images saved, but the final document could not be updated — see the log.', 'error');
    }
  }

  async function generateZoneImages(idx) {
    const btn = $(`.artreq-images-btn[data-idx="${idx}"]`);
    if (btn) { btn.disabled = true; btn.textContent = 'Rendering…'; }
    log(`Generating reference images for variation ${idx + 1}…`);
    const passes = $$(`.artreq-pass[data-idx="${idx}"]`).filter(c => c.checked).map(c => c.value);
    // The authoritative count, straight from the backend, recorded before a
    // long run starts so the log says what was actually asked for.
    try {
      const pl = JSON.parse((await window.api.artReqPlan({
        variation: artReqVariations[idx], passes: passes.length ? passes : ['scene'] })).stdout);
      if (pl.ok) log(`Plan: ${pl.passes.map(p => `${p.pass} ${p.count}`).join(', ')} — ${pl.total} render(s).`);
    } catch (e) { /* the plan is informational; never block the run on it */ }
    const res = await window.api.artReqImages({
      eventType: $('#ideation-event-type').value,
      variation: artReqVariations[idx],
      variationIndex: idx,
      passes: passes.length ? passes : ['scene']
    });
    if (btn) { btn.disabled = false; btn.textContent = 'Generate Reference Images'; }
    if (!res.success) {
      log(`Image generation failed: ${res.output}`, 'console-log-error');
      toast('Image generation failed: ' + res.output, 'error');
      return;
    }
    try {
      const parsed = JSON.parse(res.stdout);
      if (!parsed.ok) {
        log(`Image generation failed: ${parsed.error}`, 'console-log-error');
        toast(parsed.error || 'Image generation failed.', 'error');
        return;
      }
      artReqVariations[idx] = parsed.doc || artReqVariations[idx];
      reportImageReview(parsed.review);
      Object.entries(parsed.extra_reviews || {}).forEach(([label, rev]) => {
        (rev.findings || []).forEach(f => log(
          `[${label}] ${f.level.toUpperCase()} ${f.frame}: ${f.message}`,
          f.level === 'error' ? 'console-log-error' : 'console-log-warn'));
      });
      log(`Rendered ${parsed.generated}/${parsed.requested} image(s).`, 'console-log-success');
      toast(`Rendered ${parsed.generated} reference image(s).`, 'success');
      await resaveArtReqFinal(idx);
      renderArtReqResults();
      loadArtReqPreview(idx);
    } catch (e) {
      log(`Could not parse the image response: ${e.message}`, 'console-log-error');
      toast('Could not parse the image response.', 'error');
    }
  }

  async function regenerateZoneImage(idx, zoneNumber) {
    const fb = $(`#zoneimg-fb-${idx}-${zoneNumber}`);
    const allBox = $(`#zoneimg-all-${idx}-${zoneNumber}`);
    const feedback = fb ? fb.value.trim() : '';
    if (!feedback) { toast('Add a note describing the change first.', 'error'); return; }
    const applyToAll = !!(allBox && allBox.checked);
    const btn = $(`.zoneimg-regen-btn[data-idx="${idx}"][data-zone="${zoneNumber}"]`);
    if (btn) { btn.disabled = true; btn.textContent = 'Revising…'; }
    log(`Revising zone ${zoneNumber} image${applyToAll ? ' (applying to all zones)' : ''}…`);
    const res = await window.api.artReqRegenImage({
      eventType: $('#ideation-event-type').value,
      variation: artReqVariations[idx],
      variationIndex: idx,
      zoneNumber,
      feedback,
      applyToAll
    });
    if (btn) { btn.disabled = false; btn.textContent = 'Regenerate'; }
    if (!res.success) {
      log(`Image regeneration failed: ${res.output}`, 'console-log-error');
      toast('Image regeneration failed: ' + res.output, 'error');
      return;
    }
    try {
      const parsed = JSON.parse(res.stdout);
      if (!parsed.ok) {
        log(`Image regeneration failed: ${parsed.error}`, 'console-log-error');
        toast(parsed.error || 'Image regeneration failed.', 'error');
        return;
      }
      artReqVariations[idx] = parsed.doc || artReqVariations[idx];
      reportImageReview(parsed.review);
      await refreshIdeationEntries($('#ideation-event-type').value);
      renderLessonsPanel();
      await resaveArtReqFinal(idx);
      renderArtReqResults();
      loadArtReqPreview(idx);
      log(`Revised ${parsed.regenerated} image(s). The note is stored and will condition future renders.`,
          'console-log-success');
      toast(`Revised ${parsed.regenerated} image(s) — note stored for next time.`, 'success');
    } catch (e) {
      log(`Could not parse the regeneration response: ${e.message}`, 'console-log-error');
      toast('Could not parse the regeneration response.', 'error');
    }
  }

  // The automated check runs backend-side on every render; surface whatever it
  // found rather than letting a blank or duplicated image pass silently.
  function reportImageReview(review) {
    if (!review) return;
    (review.findings || []).forEach(f => {
      log(`Image review ${f.level}: ${f.frame} — ${f.message}`,
          f.level === 'error' ? 'console-log-error' : 'console-log-line');
    });
    if (review.ok === false) toast('Image review found problems — see the log.', 'error');
  }

  // Fetch (or re-fetch) the rendered document for one variation, then repaint.
  // Called after anything that changes the doc - generation, regeneration,
  // images, select-as-final (which changes the status line in the header).
  async function loadArtReqPreview(idx) {
    const res = await window.api.artReqPreview({
      eventType: $('#ideation-event-type').value,
      variation: artReqVariations[idx],
      isFinal: idx === artReqSelectedIndex
    });
    if (res.success && res.data && res.data.markdown !== undefined) {
      artReqPreviews[idx] = res.data.markdown;
    } else {
      artReqPreviews[idx] = null;
      log(`Could not render the document preview for variation ${idx + 1}: ${res.output}`, 'console-log-error');
    }
    renderArtReqResults();
  }

  function loadAllArtReqPreviews() {
    if (!artReqVariations) return;
    artReqVariations.forEach((_, i) => { loadArtReqPreview(i); });
  }

  // Cards are rebuilt wholesale on every render, which destroys any feedback
  // the reviewer has typed but not yet submitted. Previews arrive
  // asynchronously and re-render, so without this a note typed while a
  // preview was still loading would silently vanish mid-sentence. Capture
  // and restore around the rebuild, keyed by variation + section.
  function captureArtReqFeedbackDrafts() {
    const drafts = {};
    $$('.artreq-section-fb').forEach(el => {
      if (el.value.trim()) drafts[`${el.dataset.idx}::${el.dataset.section}`] = el.value;
    });
    (artReqVariations || []).forEach((_, i) => {
      const g = $(`#artreq-feedback-${i}`);
      if (g && g.value.trim()) drafts[`${i}::__general__`] = g.value;
    });
    return drafts;
  }

  function restoreArtReqFeedbackDrafts(drafts) {
    $$('.artreq-section-fb').forEach(el => {
      const v = drafts[`${el.dataset.idx}::${el.dataset.section}`];
      if (v !== undefined) el.value = v;
    });
    (artReqVariations || []).forEach((_, i) => {
      const v = drafts[`${i}::__general__`];
      const g = $(`#artreq-feedback-${i}`);
      if (g && v !== undefined) g.value = v;
    });
  }


  // ---- Art Requirement documents already on disk ---------------------------
  // Until this existed the tab could only show what the current session had
  // generated. A document finalised yesterday, or one a colleague filed from
  // another machine, or one written straight into the shared folder, was
  // invisible - so the honest answer to "is there already an Art Requirement for
  // this event?" was "open Explorer and look". This is that answer, in the tab.
  let artReqDocs = [];
  let artReqOpenedDocPath = '';

  async function loadArtReqDocs() {
    const eventType = $('#ideation-event-type').value;
    const card = $('#artreq-docs-card');
    if (!card) return;
    if (!eventType) { card.style.display = 'none'; return; }
    const res = await window.api.artReqListDocs({ eventType });
    if (!res || !res.success) {
      card.style.display = 'none';
      log(`Could not list Art Requirement documents: ${res && res.output}`, 'console-log-error');
      return;
    }
    artReqDocs = res.docs || [];
    card.style.display = artReqDocs.length ? '' : 'none';
    renderArtReqDocList();
  }

  function renderArtReqDocList() {
    const box = $('#artreq-docs-list');
    if (!box) return;
    box.innerHTML = artReqDocs.map((d, i) => {
      const when = new Date(d.modified).toLocaleString();
      const size = d.bytes < 1024 * 1024
        ? `${Math.round(d.bytes / 1024)} KB`
        : `${(d.bytes / 1024 / 1024).toFixed(1)} MB`;
      const opened = artReqOpenedDocPath === d.path;
      const cls = ['onepager-row', opened ? 'is-picked' : '', d.readable ? '' : 'is-unreadable']
        .filter(Boolean).join(' ');
      // A draft and a final can carry near-identical names, so the folder a
      // document came from is shown rather than left to be inferred.
      const where = d.folder ? `<span class="onepager-meta">${escapeHtml(d.folder)}</span>` : '';
      const type = `<span class="onepager-meta">${escapeHtml(d.ext.toUpperCase())}</span>`;
      let actions;
      if (d.readable) {
        actions = `<button class="secondary-btn small artreq-open-btn" data-idx="${i}">${opened ? 'Close' : 'Open'}</button>`
          + `<button class="secondary-btn small artreq-docx-btn" data-path="${escapeHtml(d.path)}">`
          + `${d.docxPath ? 'Re-export Word' : 'Export to Word'}</button>`
          + (d.docxPath
            ? `<button class="secondary-btn small artreq-reveal-btn" data-path="${escapeHtml(d.docxPath)}">Show .docx</button>`
            : '');
      } else {
        actions = `<button class="secondary-btn small artreq-reveal-btn" data-path="${escapeHtml(d.path)}">Show in folder</button>`;
      }
      return `<div class="${cls}">
          <span class="onepager-name">${escapeHtml(d.name)}</span>
          ${where}${type}
          <span class="onepager-meta">${escapeHtml(size)}</span>
          <span class="onepager-meta">${escapeHtml(when)}</span>
          ${actions}
        </div>`;
    }).join('');
    // Scoped to this list on purpose. A generated variation card carries an
    // .artreq-docx-btn of its own, and a document-wide query here would bind
    // those a second time - one click, two exports.
    box.querySelectorAll('.artreq-open-btn').forEach(b =>
      b.addEventListener('click', () => toggleArtReqDoc(Number(b.dataset.idx))));
    box.querySelectorAll('.artreq-docx-btn').forEach(b =>
      b.addEventListener('click', () => exportArtReqDocx(b.dataset.path, b)));
    box.querySelectorAll('.artreq-reveal-btn').forEach(b =>
      b.addEventListener('click', () => revealArtReqDoc(b.dataset.path)));
  }

  async function toggleArtReqDoc(i) {
    const doc = artReqDocs[i];
    const view = $('#artreq-doc-view');
    if (!doc || !view) return;
    if (artReqOpenedDocPath === doc.path) {
      artReqOpenedDocPath = '';
      view.innerHTML = '';
      renderArtReqDocList();
      return;
    }
    const eventType = $('#ideation-event-type').value;
    view.innerHTML = `<p class="settings-hint">Opening ${escapeHtml(doc.name)}…</p>`;
    const res = await window.api.artReqReadDoc({ eventType, docPath: doc.path });
    if (!res || !res.success) {
      view.innerHTML = '';
      toast('Could not open that document: ' + (res && res.output), 'error');
      return;
    }
    artReqOpenedDocPath = doc.path;
    // The same markdown renderer the generated variations use, so a document
    // written last month reads exactly like one written a minute ago. The index
    // is negative so its section anchors cannot collide with a generated card's,
    // which are keyed by variation index from zero.
    view.innerHTML = renderMarkdownDoc(res.data.markdown, -1 - i).html;
    renderArtReqDocList();
  }

  // The Word export on a generated variation works off the file already written
  // to disk - the draft, or the final once it has been saved - never off the
  // variation held in memory, so what lands in Word is exactly the document a
  // reviewer would open from the folder.
  function artReqExportRow(idx, selected) {
    const p = (selected && artReqSavedPath) || artReqDraftPaths[idx] || '';
    if (!p) return '';
    return `<div class="execution-footer" style="justify-content:flex-start;">
          <button class="secondary-btn small artreq-docx-btn" data-path="${escapeHtml(p)}">Export to Word</button>
        </div>`;
  }

  async function exportArtReqDocx(docPath, btn) {
    if (!docPath) return;
    const eventType = $('#ideation-event-type').value;
    const label = btn ? btn.textContent : '';
    if (btn) { btn.disabled = true; btn.textContent = 'Exporting…'; }
    const res = await window.api.artReqExportDocx({ eventType, docPath });
    if (btn) { btn.disabled = false; btn.textContent = label; }
    if (!res || !res.success) {
      log(`Word export failed: ${res && res.output}`, 'console-log-error');
      toast('Word export failed: ' + (res && res.output), 'error');
      return;
    }
    log(`Exported to ${res.data.docx_path}`, 'console-log-success');
    toast('Exported to Word — it is beside the Markdown in the same folder.', 'success');
    await loadArtReqDocs();
  }

  async function revealArtReqDoc(docPath) {
    if (!docPath) return;
    const eventType = $('#ideation-event-type').value;
    const res = await window.api.artReqRevealDoc({ eventType, docPath });
    if (!res || !res.success) toast((res && res.output) || 'Could not show that file.', 'error');
  }

  $('#artreq-refresh-docs-btn').addEventListener('click', () => loadArtReqDocs());
  // The select defaults to a real event type, so list its folder on load
  // rather than waiting for someone to touch the dropdown.
  loadArtReqDocs();
  function renderArtReqResults() {
    const container = $('#artreq-results');
    if (!artReqVariations || !artReqVariations.length) { container.innerHTML = ''; return; }
    const drafts = captureArtReqFeedbackDrafts();
    container.innerHTML = artReqVariations.map((v, i) => renderArtReqCard(v, i)).join('');
    $$('.artreq-select-btn').forEach(btn => btn.addEventListener('click', () => selectArtReqVariation(Number(btn.dataset.idx))));
    container.querySelectorAll('.artreq-docx-btn').forEach(btn => btn.addEventListener('click', () => exportArtReqDocx(btn.dataset.path, btn)));
    $$('.artreq-regen-btn').forEach(btn => btn.addEventListener('click', () => regenerateArtReqVariation(Number(btn.dataset.idx))));
    $$('.artreq-images-btn').forEach(btn => btn.addEventListener('click', () => generateZoneImages(Number(btn.dataset.idx))));
    $$('.artreq-pass').forEach(cb => cb.addEventListener('change', () => refreshPlan(Number(cb.dataset.idx))));
    $$('.zoneimg-regen-btn').forEach(btn => btn.addEventListener('click', () => regenerateZoneImage(Number(btn.dataset.idx), btn.dataset.zone)));
    restoreArtReqFeedbackDrafts(drafts);
  }

  $('#artreq-generate-btn').addEventListener('click', async () => {
    // Two legitimate sources now: a theme generated and marked Final in this
    // session, or a One-Pager that already existed. Exactly one is armed at a
    // time - switching modes clears the other.
    const usingDoc = !!ideationPickedOnePager;
    if (!usingDoc && ideationSelectedIndex === -1) {
      toast('Select an Ideation theme as Final, or pick an existing One-Pager.', 'error');
      return;
    }
    const eventType = $('#ideation-event-type').value;
    const ideationTheme = usingDoc ? null : ideationThemes[ideationSelectedIndex];
    const sourceLabel = usingDoc ? ideationPickedOnePager.name : ideationTheme.theme_name;
    let variationCount = parseInt($('#artreq-variation-count').value, 10);
    if (!Number.isFinite(variationCount)) variationCount = 3;
    variationCount = Math.max(1, Math.min(10, variationCount));
    $('#artreq-variation-count').value = variationCount;

    const btn = $('#artreq-generate-btn');
    btn.disabled = true; btn.textContent = `Generating ${variationCount} variation${variationCount === 1 ? '' : 's'}...`;
    log(`Generating ${variationCount} Art Requirement variation(s) for "${sourceLabel}"...`);

    const res = await window.api.artReqGenerate({
      eventType, variationCount,
      ideationTheme: ideationTheme || undefined,
      ideationDocText: usingDoc ? ideationPickedOnePager.text : undefined,
      ideationDocName: usingDoc ? ideationPickedOnePager.name : undefined,
      ideationDocPath: usingDoc ? ideationPickedOnePager.docPath : undefined,
    });
    btn.disabled = false; btn.textContent = 'Generate Art Requirement Variations';

    if (!res.success) { log(`Art Requirement generation failed: ${res.output}`, 'console-log-error'); toast('Generation failed: ' + res.output, 'error'); return; }
    try {
      const parsed = JSON.parse(res.stdout);
      if (parsed.error) { log(`Error: ${parsed.error}`, 'console-log-error'); toast(parsed.error, 'error'); return; }
      artReqVariations = parsed.variations;
      artReqDraftPaths = parsed.draft_paths || [];
      artReqPreviews = [];
      artReqSelectedIndex = -1;
      artReqSavedPath = '';
      renderArtReqResults();
      loadAllArtReqPreviews();
      // Say when reference art went with the request. It changes what comes
      // back, and it is billed on every variation, so it should never be
      // invisible.
      const refs = parsed.reference_images || 0;
      const refNote = refs
        ? ` with ${refs} reference image${refs === 1 ? '' : 's'} from the One-Pager attached`
        : '';
      log(`Generated ${artReqVariations.length} Art Requirement variation(s) via `
        + `${parsed.provider === 'ollama' ? 'local Ollama' : 'Claude'}${refNote}.`,
        'console-log-success');
      toast(`${artReqVariations.length} variation(s) generated.`, 'success');
    } catch (e) {
      log(`Error parsing Art Requirement response: ${e.message}`, 'console-log-error');
      toast('Could not parse the generated variations.', 'error');
    }
  });

  // Per-section notes and the general box are merged into the one feedback
  // string the backend already takes, each prefixed with the section it was
  // written against. Nothing downstream changes - the store, the lesson
  // extraction and the regeneration prompt all keep working as they are, and
  // the model gets told which section a note is about instead of having to
  // infer it.
  function collectArtReqFeedback(idx) {
    const parts = [];
    $$(`.artreq-section-fb[data-idx="${idx}"]`).forEach(el => {
      const v = el.value.trim();
      if (v) parts.push(`[${el.dataset.section}] ${v}`);
    });
    const general = ($(`#artreq-feedback-${idx}`) || {}).value || '';
    if (general.trim()) parts.push(general.trim());
    return parts.join('\n');
  }

  async function regenerateArtReqVariation(idx) {
    const feedback = collectArtReqFeedback(idx);
    if (!feedback) { toast('Add feedback in a section or the general box before regenerating.', 'error'); return; }
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
      // This feedback has been applied and stored - clear the boxes before the
      // rebuild so the draft-preservation above doesn't restore notes the
      // reviewer has already submitted onto the freshly rewritten document.
      $$(`.artreq-section-fb[data-idx="${idx}"]`).forEach(el => { el.value = ''; });
      const genBox = $(`#artreq-feedback-${idx}`);
      if (genBox) genBox.value = '';
      await refreshIdeationEntries(eventType);
      renderLessonsPanel();
      renderArtReqResults();
      loadArtReqPreview(idx);
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
      // The status line in the document header changes on finalisation, so
      // re-render the preview rather than leaving it showing "Draft".
      loadArtReqPreview(idx);
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
    endless_treasure: 'endless-treasure-view', chapter: 'chapter-view'
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
        .map(r => {
          const meta = [r.start ? fmtRange(r.start, r.end) : '', r.switch || ''].filter(Boolean).join(' · ');
          return `<div class="sp-row">${escapeHtml(r.name || '(unnamed)')}${meta ? ` <span class="feedback-when">${escapeHtml(meta)}</span>` : ''}</div>`;
        }).join('');
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
    $('#chapter-lvlpy-input').value = cfg.chapterLvlPyPath || '';
    $('#chapter-mapbuilder-input').value = cfg.chapterMapBuilderPath || '';
    updateRobotPvpPathDisplay();
    $('#anthropic-key-input').value = cfg.anthropicApiKey || '';
    $('#comfy-url-input').value = cfg.comfyUrl || '';
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
