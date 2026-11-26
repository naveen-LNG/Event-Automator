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

  // ---------------- Update download bar (fed by main's auto-updater) ----------------
  window.api.onUpdateProgress((p) => {
    let bar = document.getElementById('update-bar');
    if (!p) { if (bar) bar.remove(); return; }
    if (!bar) {
      bar = document.createElement('div');
      bar.id = 'update-bar';
      bar.style.cssText = 'position:fixed;top:0;left:0;right:0;z-index:9999;padding:6px 14px;background:var(--card);' +
        'border-bottom:1px solid var(--line);color:var(--fg);font:12px var(--font-sans);display:flex;align-items:center;gap:12px;';
      bar.innerHTML = '<span id="update-bar-text"></span><div style="flex:1;height:6px;background:var(--line);border-radius:3px;overflow:hidden;">' +
        '<div id="update-bar-fill" style="height:100%;width:0;background:var(--good);transition:width 300ms;"></div></div>';
      document.body.appendChild(bar);
    }
    const mb = (b) => (b / 1048576).toFixed(0);
    const pct = Math.floor(p.percent || 0);
    bar.querySelector('#update-bar-text').textContent = p.total
      ? `Downloading update v${p.version}: ${pct}% (${mb(p.transferred)} / ${mb(p.total)} MB, ${(p.speed / 1048576).toFixed(1)} MB/s)`
      : `Downloading update v${p.version}: starting...`;
    bar.querySelector('#update-bar-fill').style.width = `${pct}%`;
  });

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
    if (viewId === 'xlsx-merge-view') loadXlsxMergeDefaults();
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
    else if (currentView === 'xlsx-merge-view') await loadXlsxMergeDefaults(true);
  }

  $$('.menu-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      // switchView() clears every active-btn itself, so the highlight has to
      // be applied AFTER it — marking first meant switchView immediately
      // stripped it and every click left the sidebar with nothing lit.
      switchView(btn.dataset.view, btn.dataset.title);
      if (btn.dataset.view === 'localization-view') locOfferFinishedBatch();
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

  // Chapter Update tool paths. Folder pickers, persisted in config.
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

  // AB-test switch columns found in 活动数据 (e.g. '控制开关 $LNMN2').
  const mgAbColumns = () => (cachedMinigames && cachedMinigames.ab_columns) || [];

  function mgEl(tag, props, field, i) {
    const el = Object.assign(document.createElement(tag), props);
    el.dataset.f = field; el.dataset.i = i;
    return el;
  }
  function mgSelect(options, value, field, i) {
    const s = mgEl('select', { className: 'text-input mg-cell' }, field, i);
    for (const [v, label] of options) s.add(new Option(label, v));
    s.value = value;
    return s;
  }
  const MG_BOOL = [['true', 'TRUE'], ['false', 'FALSE']];
  const MG_AB = [['', '-'], ...MG_BOOL];

  function renderMinigamePreview() {
    const container = $('#minigame-preview-container');
    container.textContent = '';
    if (generatedMinigameRows.length) {
      const ab = mgAbColumns();
      const table = document.createElement('table');
      table.className = 'mg-preview-table mg-edit-table';
      const head = table.createTHead().insertRow();
      for (const h of ['ID', 'Minigame', 'Start', 'End', 'Double Wk ID', 'Discount ID', '控制开关', ...ab, ''])
        head.appendChild(Object.assign(document.createElement('th'), { textContent: h }));
      const body = table.createTBody();
      const names = (cachedMinigames ? cachedMinigames.minigames : []).map(m => [m.name, m.name]);
      generatedMinigameRows.forEach((r, i) => {
        if (!names.some(([n]) => n === r.minigame)) names.push([r.minigame, r.minigame]);
        const tr = body.insertRow();
        const cells = [
          mgEl('input', { className: 'mg-time mg-cell mg-id', type: 'number', value: r.id }, 'id', i),
          mgSelect(names, r.minigame, 'minigame', i),
          mgEl('input', { className: 'mg-time mg-cell', value: r.start_time }, 'start_time', i),
          mgEl('input', { className: 'mg-time mg-cell', value: r.end_time }, 'end_time', i),
          mgEl('input', { className: 'mg-time mg-cell', value: r.double_week_id || '' }, 'double_week_id', i),
          mgEl('input', { className: 'mg-time mg-cell mg-id', value: r.discount_id === -1 ? '' : r.discount_id }, 'discount_id', i),
          mgSelect(MG_BOOL, String(r.switch !== false), 'switch', i),
          ...ab.map(h => mgSelect(MG_AB, r.ab && r.ab[h] != null ? String(r.ab[h]) : '', 'ab:' + h, i)),
          mgEl('button', { className: 'secondary-btn small mg-del', textContent: '✕', title: 'Delete row' }, 'delete', i),
        ];
        for (const el of cells) tr.insertCell().appendChild(el);
      });
      const wrap = document.createElement('div');
      wrap.className = 'mg-edit-wrap';
      wrap.appendChild(table);
      container.appendChild(wrap);
    } else {
      container.appendChild(Object.assign(document.createElement('p'), {
        className: 'header-subtitle small', style: 'margin:0;',
        textContent: 'Generate a set above, or add rows one by one.' }));
    }
    const add = Object.assign(document.createElement('button'), { className: 'secondary-btn small', textContent: '+ Add row' });
    add.dataset.f = 'add';
    add.style.marginTop = '10px';
    add.disabled = !cachedMinigames;
    container.appendChild(add);
  }

  function newMinigameRow() {
    const last = generatedMinigameRows[generatedMinigameRows.length - 1];
    const ids = generatedMinigameRows.map(r => Number(r.id)).filter(n => !isNaN(n));
    let start = last ? addMgDays(parseMgDateTime(last.end_time), cachedMinigames.gap_days || 1)
                     : parseMgDateTime($('#minigame-set-start-input').value);
    if (isNaN(start.getTime())) start = new Date();
    start.setHours(0, 0, 0, 0);
    const end = new Date(start); end.setHours(23, 59, 59);
    const mg = cachedMinigames.minigames[0];
    return {
      id: Math.max(cachedMinigames.max_id, ...ids) + 1,
      minigame: mg ? mg.name : '',
      start_time: formatMgDateTime(start), end_time: formatMgDateTime(end),
      double_week_id: '', discount_id: mg ? mg.discount_id : -1,
      switch: true, ab: Object.fromEntries(mgAbColumns().map(h => [h, null])),
    };
  }

  // Edits write straight into generatedMinigameRows; only structural changes re-render,
  // so typing never loses focus.
  function onMinigameEdit(e) {
    const el = e.target, f = el.dataset.f;
    if (!f || f === 'add' || f === 'delete') return;
    const r = generatedMinigameRows[Number(el.dataset.i)];
    const v = el.value;
    el.classList.remove('mg-invalid');
    if (f.startsWith('ab:')) (r.ab ||= {})[f.slice(3)] = v === '' ? null : v === 'true';
    else if (f === 'switch') r.switch = v === 'true';
    else if (f === 'id') r.id = v === '' ? '' : Number(v);
    else if (f === 'discount_id') r.discount_id = v.trim() === '' ? -1 : (isNaN(Number(v)) ? v.trim() : Number(v));
    else if (f === 'minigame') {
      r.minigame = v;
      const mg = cachedMinigames.minigames.find(m => m.name === v);
      r.discount_id = mg ? mg.discount_id : -1;
      renderMinigamePreview();
    } else r[f] = v.trim();
  }
  $('#minigame-preview-container').addEventListener('input', onMinigameEdit);
  $('#minigame-preview-container').addEventListener('change', onMinigameEdit);
  $('#minigame-preview-container').addEventListener('click', (e) => {
    const f = e.target.dataset && e.target.dataset.f;
    if (f === 'add' && cachedMinigames) generatedMinigameRows.push(newMinigameRow());
    else if (f === 'delete') generatedMinigameRows.splice(Number(e.target.dataset.i), 1);
    else return;
    renderMinigamePreview();
  });

  // Returns a list of problems (empty = ok) and marks the offending cells.
  function validateMinigameRows() {
    const problems = [];
    const DT = /^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$/;
    const seen = new Set();
    const mark = (i, f) => {
      const el = $(`#minigame-preview-container [data-i="${i}"][data-f="${f}"]`);
      if (el) el.classList.add('mg-invalid');
    };
    generatedMinigameRows.forEach((r, i) => {
      const n = i + 1;
      if (!Number.isInteger(r.id) || r.id <= cachedMinigames.max_id) { problems.push(`Row ${n}: ID must be a whole number above ${cachedMinigames.max_id} (the sheet's last ID)`); mark(i, 'id'); }
      else if (seen.has(r.id)) { problems.push(`Row ${n}: ID ${r.id} is used twice`); mark(i, 'id'); }
      seen.add(r.id);
      if (!r.minigame) { problems.push(`Row ${n}: pick a minigame`); mark(i, 'minigame'); }
      for (const f of ['start_time', 'end_time'])
        if (!DT.test(String(r[f]).trim()) || isNaN(parseMgDateTime(r[f]).getTime())) { problems.push(`Row ${n}: ${f === 'start_time' ? 'Start' : 'End'} must be YYYY-MM-DD HH:MM:SS`); mark(i, f); }
      if (DT.test(r.start_time) && DT.test(r.end_time) && parseMgDateTime(r.end_time) < parseMgDateTime(r.start_time)) { problems.push(`Row ${n}: End is before Start`); mark(i, 'end_time'); }
    });
    return problems;
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
      renderMinigamePreview();
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
        discount_id: mg ? mg.discount_id : -1,
        switch: true,
        ab: Object.fromEntries(mgAbColumns().map(h => [h, null]))
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
    const problems = validateMinigameRows();
    if (problems.length) {
      problems.forEach(p => log(p, 'console-log-error'));
      toast(`Fix ${problems.length} problem(s) in the preview first:<br>${problems.slice(0, 4).map(escapeHtml).join('<br>')}`, 'error');
      return;
    }

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

  // Preview = read-only run of the script (PREVIEW_JSON line); the user edits the plan, and Commit
  // sends those edits back as overrides. Any change to the inputs drops the preview.
  let streakPlan = null;
  const STREAK_COLS = [['id', 'ID', 'number'], ['type', 'Type', 'number'], ['start', 'Start', 'text'], ['end', 'End', 'text'],
    ['tasks', '固定任务列表', 'text'], ['activity_id', '活动ID', 'text'], ['card_id', '集卡活动ID', 'text']];

  function streakInvalidate() {
    streakPlan = null;
    $('#streak-plan-card').style.display = 'none';
    $('#streak-execute-btn').disabled = true;
  }
  ['#event-id-input', '#start-date-input', '#duration-days-input', '#streak-event-name'].forEach(sel =>
    $(sel).addEventListener('input', streakInvalidate));

  function renderStreakPlan() {
    const wrap = $('#streak-plan-rows');
    wrap.textContent = '';
    const table = document.createElement('table');
    table.className = 'mg-preview-table mg-edit-table';
    const head = table.createTHead().insertRow();
    for (const [, h] of STREAK_COLS) head.appendChild(Object.assign(document.createElement('th'), { textContent: h }));
    head.appendChild(document.createElement('th'));
    const body = table.createTBody();
    streakPlan.rows.forEach((r, i) => {
      const tr = body.insertRow();
      for (const [f, , type] of STREAK_COLS) {
        const el = Object.assign(document.createElement('input'), { className: 'text-input mono mg-cell', type, value: r[f] ?? '' });
        el.dataset.i = i; el.dataset.f = f;
        if (f === 'tasks') el.style.minWidth = '260px'; else if (f === 'start' || f === 'end') el.style.minWidth = '152px';
        tr.insertCell().appendChild(el);
      }
      const del = Object.assign(document.createElement('button'), { className: 'secondary-btn small mg-del', textContent: '✕', title: 'Delete row' });
      del.dataset.f = 'delete'; del.dataset.i = i;
      tr.insertCell().appendChild(del);
    });
    const box = document.createElement('div');
    box.className = 'mg-edit-wrap';
    box.appendChild(table);
    wrap.appendChild(box);
    const add = Object.assign(document.createElement('button'), { className: 'secondary-btn small', textContent: '+ Add row' });
    add.dataset.f = 'add'; add.style.marginTop = '10px';
    wrap.appendChild(add);
  }

  $('#streak-plan-rows').addEventListener('input', (e) => {
    const f = e.target.dataset.f;
    if (!streakPlan || !f || f === 'add' || f === 'delete') return;
    e.target.classList.remove('mg-invalid');
    streakPlan.rows[Number(e.target.dataset.i)][f] = e.target.value;
  });
  $('#streak-plan-rows').addEventListener('click', (e) => {
    const f = e.target.dataset.f;
    if (!streakPlan) return;
    if (f === 'delete') streakPlan.rows.splice(Number(e.target.dataset.i), 1);
    else if (f === 'add') {
      const last = streakPlan.rows[streakPlan.rows.length - 1] || {};
      // Next weekly slot: day after the last End at 06:00, running 7 days
      const pad = (n) => String(n).padStart(2, '0');
      const fmt = (d, t) => `${d.getUTCFullYear()}-${pad(d.getUTCMonth() + 1)}-${pad(d.getUTCDate())} ${t}`;
      const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(String(last.end || ''));
      const next = m ? new Date(Date.UTC(+m[1], +m[2] - 1, +m[3] + 1)) : null;
      const endD = next ? new Date(next.getTime() + 6 * 86400000) : null;
      streakPlan.rows.push({ ...last, id: Number(last.id) + 1, type: Number(last.type) === 1 ? 2 : 1,
        start: next ? fmt(next, '06:00:00') : last.start, end: endD ? fmt(endD, '23:59:59') : last.end });
    } else return;
    renderStreakPlan();
  });

  // Normalises + validates the edited rows; returns { rows, problems }.
  function streakCollectRows() {
    const DT = /^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$/;
    const problems = [];
    const rows = streakPlan.rows.map((r, i) => {
      const n = i + 1;
      const id = Number(r.id), type = Number(r.type);
      const start = String(r.start).trim(), end = String(r.end).trim(), tasks = String(r.tasks).trim();
      if (!Number.isInteger(id) || id <= 0) problems.push(`Row ${n}: ID must be a whole number`);
      if (type !== 1 && type !== 2) problems.push(`Row ${n}: Type must be 1 or 2`);
      if (!DT.test(start) || !DT.test(end)) problems.push(`Row ${n}: Start/End must look like 2026-09-17 06:00:00`);
      else if (start >= end) problems.push(`Row ${n}: Start must be before End`);
      if (!/^\d+(,\d+)*$/.test(tasks)) problems.push(`Row ${n}: 固定任务列表 must be comma-separated task IDs`);
      const act = String(r.activity_id).trim();
      const card = String(r.card_id ?? '').trim();
      return { id, type, start, end, tasks,
        activity_id: /^\d+$/.test(act) ? Number(act) : act,
        card_id: /^\d+$/.test(card) ? Number(card) : (card || null) };
    });
    if (!rows.length) problems.push('Add at least one 任务链 row');
    return { rows, problems };
  }

  function streakBasePayload() {
    const eventId = ($('#event-id-input').value || '').trim();
    const eventName = ($('#streak-event-name').value || '').trim();
    const times = computeEventTimes($('#start-date-input').value, parseInt($('#duration-days-input').value, 10));
    if (!eventId || eventId.length < 4) { toast('Please fill in a valid Event ID on the Main Events tab first.', 'error'); return null; }
    if (!eventName) { toast('Please fill in Event Name.', 'error'); return null; }
    if (!times) { toast('Please fill in Start Date and Duration (days) on the Main Events tab first.', 'error'); return null; }
    return { event_id: eventId, event_name: eventName, start_time: times.startTime, end_time: times.endTime };
  }

  $('#streak-preview-btn').addEventListener('click', async () => {
    const base = streakBasePayload();
    if (!base) return;
    const designPath = await window.api.resolveDesignPath(activeBranchRoot(), cfg.activeWorkspace);
    const btn = $('#streak-preview-btn');
    btn.disabled = true;
    const res = await window.api.residentTaskCommit({ designPath, payload: { ...base, preview: true } });
    btn.disabled = false;
    streakInvalidate();
    const out = res.output || '';
    if (out) log(out.replace(/PREVIEW_JSON:.*/g, 'PREVIEW_JSON: (plan shown below)'), res.success ? 'console-log-line' : 'console-log-error');
    if (!res.success) { toast('Preview failed: ' + out, 'error'); return; }
    if (/already present/i.test(out)) { toast('That Event ID already has a cycle in resident_task.xlsx — nothing to add.', 'error'); return; }
    const line = out.split('\n').find(l => l.startsWith('PREVIEW_JSON:'));
    if (!line) { toast('Preview returned no plan — see the console.', 'error'); return; }
    streakPlan = JSON.parse(line.slice('PREVIEW_JSON:'.length));
    $('#streak-plan-token').value = streakPlan.token_name || '';
    $('#streak-plan-pack').value = streakPlan.pack_prefix || '';
    $('#streak-plan-summary').textContent =
      `Chain continues after row ${streakPlan.last_row} (ends ${streakPlan.last_end}). ` +
      `Token will be updated on ${streakPlan.token_rows} 任务列表 row(s), card-pack prefix on ${streakPlan.pack_rows} row(s).`;
    renderStreakPlan();
    $('#streak-plan-card').style.display = '';
    $('#streak-execute-btn').disabled = false;
  });

  $('#streak-execute-btn').addEventListener('click', async () => {
    if (!streakPlan) { toast('Preview the changes first.', 'error'); return; }
    const base = streakBasePayload();
    if (!base) return;
    const { rows, problems } = streakCollectRows();
    if (problems.length) { toast(problems.join(' · '), 'error'); return; }

    const designPath = await window.api.resolveDesignPath(activeBranchRoot(), cfg.activeWorkspace);
    const btn = $('#streak-execute-btn');
    btn.disabled = true;
    const res = await window.api.residentTaskCommit({
      designPath,
      payload: { ...base, overrides: {
        rows,
        token_name: $('#streak-plan-token').value.trim(),
        pack_prefix: $('#streak-plan-pack').value.trim(),
      } }
    });

    // Always show the script's own output — "success" can also be an idempotent no-op.
    if (res.output) log(res.output, res.success ? 'console-log-line' : 'console-log-error');

    if (res.success) {
      if (/already present/i.test(res.output || '')) {
        log('Streak Challenge Commit: no changes made (activity ID already present).', 'console-log-line');
        toast('No changes made — that Event ID already has a cycle in resident_task.xlsx.', 'error');
        btn.disabled = false;
      } else {
        log('Streak Challenge Commit Successful!', 'console-log-success');
        toast('resident_task.xlsx updated with the new cycle.', 'success');
        streakInvalidate();
      }
    } else {
      log(`Streak Challenge Commit Failed: ${res.output}`, 'console-log-error');
      toast('Commit failed: ' + res.output, 'error');
      btn.disabled = false;
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
      if (!$('#fsg-duration-days').value) $('#fsg-duration-days').value = cachedFSG.window_days;
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
    const days = parseInt($('#fsg-duration-days').value, 10);
    if (!m || !(days >= 1)) {
      $('#fsg-preview-alias').textContent = m ? `全屏礼包${m[1].slice(2)}${m[2]}${m[3]}` : '...';
      $('#fsg-preview-start').textContent = '...';
      $('#fsg-preview-end').textContent = '...';
      return;
    }
    const [, y, mo, d] = m;
    const yy = y.slice(2);
    $('#fsg-preview-alias').textContent = `全屏礼包${yy}${mo}${d}`;
    const startMs = Date.UTC(Number(y), Number(mo) - 1, Number(d));
    const endMs = startMs + (days - 1) * 86400000;
    $('#fsg-preview-start').textContent = `${startDateStr} 00:00:00`;
    $('#fsg-preview-end').textContent = `${fmtDate(endMs)} 23:59:59`;
  }
  $('#fsg-start-date').addEventListener('input', updateFSGPreview);
  $('#fsg-duration-days').addEventListener('input', updateFSGPreview);

  $('#fsg-execute-btn').addEventListener('click', async () => {
    if (!cachedFSG) { toast('Full Screen Gift metadata still loading...', 'error'); return; }
    const startDateStr = $('#fsg-start-date').value.trim();
    if (!/^\d{4}-\d{2}-\d{2}$/.test(startDateStr)) { toast('Please enter a valid Start Date (yyyy-mm-dd).', 'error'); return; }
    const durationDays = parseInt($('#fsg-duration-days').value, 10);
    if (!(durationDays >= 1)) { toast('Please enter a Duration of at least 1 day.', 'error'); return; }

    const root = activeBranchRoot();
    const designPath = await window.api.resolveDesignPath(root, cfg.activeWorkspace);
    const btn = $('#fsg-execute-btn');
    btn.disabled = true; btn.textContent = 'Committing...';
    const res = await window.api.fullScreenGiftCommit({ designPath, payload: { start_date: startDateStr, duration_days: durationDays } });
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
  // config like robotCalcPath rather than being retyped per run.
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
    $('#trivia-preview-answers').textContent = '...';
    $('#trivia-preview-status').textContent = '...';
    const root = activeBranchRoot();
    if (!root) { log('Cannot load Trivia Challenge info: branch path not set.', 'console-log-error'); return; }
    log('Loading Trivia Challenge cycle info from mini_mgr.xlsx...');
    const res = await window.api.triviaGet({ branchRoot: root, workspace: cfg.activeWorkspace, questionsSource: triviaQuestionsSource() });
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
    const c = cachedTrivia.correct_counts || {};
    $('#trivia-preview-answers').textContent = `${c['1'] ?? 0} / ${c['2'] ?? 0}`;
    $('#trivia-preview-status').textContent = cachedTrivia.already_current ? 'Already current' : 'Needs refresh';
    renderTriviaBatches();
  }

  // '' = Auto (next unused); otherwise the pool module for the picked batch number.
  function triviaQuestionsSource() {
    const n = $('#trivia-batch-select').value;
    return n ? `trivia_questions_batch_${String(n).padStart(2, '0')}` : '';
  }

  function renderTriviaBatches() {
    const usage = cachedTrivia.batch_usage || [];
    const live = {}, past = {};
    for (const u of usage) {
      if (u.live_batch != null) (live[u.live_batch] ||= []).push(u.label);
      for (const [, b] of u.archived) (past[b] ||= new Set()).add(u.label);
    }
    const select = $('#trivia-batch-select');
    const picked = select.value;
    select.textContent = '';
    select.add(new Option(`Auto: next unused batch (#${cachedTrivia.next_unused_batch})`, ''));
    for (const n of cachedTrivia.available_batches || []) {
      const where = live[n] ? `live in ${live[n].join(', ')}` : past[n] ? `used earlier in ${[...past[n]].join(', ')}` : 'unused';
      select.add(new Option(`#${n}: ${where}`, String(n)));
    }
    select.value = picked;

    const body = $('#trivia-usage-body');
    body.textContent = '';
    for (const u of usage) {
      const tr = body.insertRow();
      const liveBatch = u.live_batch == null ? 'unknown' : `#${u.live_batch}${u.live_batch_inferred ? ' (matched by text, no marker)' : ''}`;
      const earlier = u.archived.map(([p, b]) => `#${b} (${p})`).join(', ') || '-';
      for (const text of [u.label, u.live_prefix || '-', liveBatch, earlier]) tr.insertCell().textContent = text;
    }
  }

  $('#trivia-batch-select').addEventListener('change', () => loadTrivia(true));
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
    const res = await window.api.triviaRun({ branchRoot: root, workspace: cfg.activeWorkspace, questionsSource: triviaQuestionsSource() });
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

  // ---------------- Merge to Trunk tab ----------------
  // Thin front-end over xlsx_merge.py (the same script as tools/XlsxMerge): the
  // backend does the 3-way work and leaves report.md / conflicts.csv /
  // remap_suggest.csv / summary.json in a per-run folder; this tab only shapes
  // the arguments and renders what came back. Apply is gated on a dry run with
  // identical settings, so nobody writes into trunk on a stale picture.
  let xmLast = null;        // result of the most recent run (dry or apply)
  let xmDryRunKey = '';     // settings fingerprint of the last successful dry run
  let xmRunning = false;
  let xmArmedUntil = 0;     // two-click confirm on Apply

  const XM_INPUT_IDS = ['xm-branch-ws', 'xm-trunk-ws', 'xm-base-rev', 'xm-files', 'xm-on-collision',
    'xm-scope', 'xm-scope-include', 'xm-prefer-branch', 'xm-prefer-trunk', 'xm-yield-trunk', 'xm-auto-ids', 'xm-switch-off',
    'xm-apply-deletes', 'xm-add-missing', 'xm-no-text', 'xm-unlock-local', 'xm-remap'];

  function xmSettings() {
    return {
      branchWs: $('#xm-branch-ws').value.trim(),
      trunkWs: $('#xm-trunk-ws').value.trim(),
      base: $('#xm-base-rev').value.trim(),
      files: $('#xm-files').value.trim(),
      onCollision: $('#xm-on-collision').value,
      scope: $('#xm-scope').value.trim(),
      scopeInclude: $('#xm-scope-include').value.trim(),
      preferBranch: $('#xm-prefer-branch').value.trim(),
      preferTrunk: $('#xm-prefer-trunk').value.trim(),
      yieldTrunk: $('#xm-yield-trunk').value.trim(),
      autoIds: $('#xm-auto-ids').checked,
      switchOff: $('#xm-switch-off').checked,
      applyDeletes: $('#xm-apply-deletes').checked,
      addMissing: $('#xm-add-missing').checked,
      noText: $('#xm-no-text').checked,
      unlockLocal: $('#xm-unlock-local').checked,
      remapCsv: $('#xm-remap').value
    };
  }
  // unlockLocal only matters at apply time, so it must not invalidate a dry run.
  const xmKey = (s) => JSON.stringify({ ...s, unlockLocal: undefined });

  async function loadXlsxMergeDefaults(force) {
    const d = await window.api.xlsxMergeDefaults();
    if (force || !$('#xm-branch-ws').value) $('#xm-branch-ws').value = d.branchWs || '';
    if (!$('#xm-trunk-ws').value) $('#xm-trunk-ws').value = d.trunkWs || '';
    xmRefreshApplyGate();
  }

  function xmRefreshApplyGate() {
    const s = xmSettings();
    const stale = !!xmDryRunKey && xmKey(s) !== xmDryRunKey;
    $('#xm-apply-btn').disabled = xmRunning || !xmDryRunKey || stale || !s.branchWs || !s.trunkWs;
    if (stale) $('#xm-apply-hint').textContent = 'Settings changed since the last dry run — run it again before applying.';
  }
  XM_INPUT_IDS.forEach(id => {
    $('#' + id).addEventListener('input', xmRefreshApplyGate);
    $('#' + id).addEventListener('change', xmRefreshApplyGate);
  });

  $('#xm-browse-branch-btn').addEventListener('click', async () => {
    const p = await window.api.browseFolder($('#xm-branch-ws').value);
    if (p) { $('#xm-branch-ws').value = p; xmRefreshApplyGate(); }
  });
  $('#xm-browse-trunk-btn').addEventListener('click', async () => {
    const p = await window.api.browseFolder($('#xm-trunk-ws').value);
    if (p) { $('#xm-trunk-ws').value = p; xmRefreshApplyGate(); }
  });
  $('#xm-clear-remap-btn').addEventListener('click', () => { $('#xm-remap').value = ''; xmRefreshApplyGate(); });
  $('#xm-load-collisions-btn').addEventListener('click', () => {
    const suggest = (xmLast && xmLast.suggest) || [];
    if (!suggest.length) { toast('The last dry run reported no ID collisions.'); return; }
    const existing = $('#xm-remap').value.split('\n').map(l => l.trim()).filter(Boolean);
    // Rows the tool could resolve itself (same Alias already in trunk) arrive with new_id prefilled;
    // "same row, content differs" rows are not ID clashes and are left out (use the collision policy).
    const wanted = suggest.filter(r => !(r.note || '').startsWith('same row'));
    const fresh = wanted.map(r => `${r.xlsx},${r.sheet},${r.old_id},${r.new_id || ''}`)
      .filter(l => !existing.some(e => e.startsWith(l.slice(0, l.lastIndexOf(',') + 1))));
    $('#xm-remap').value = [...existing, ...fresh].join('\n');
    xmRefreshApplyGate();
    const filled = wanted.filter(r => r.new_id).length;
    toast(`${fresh.length} colliding IDs added — ${filled} prefilled from trunk's Alias match, ${fresh.length - filled} need a new_id after the trailing comma.`);
  });

  async function xmRun(apply) {
    const s = xmSettings();
    if (!s.branchWs || !s.trunkWs) { toast('Set both the branch and trunk workspace paths first.', 'error'); return; }
    if (s.branchWs.toLowerCase() === s.trunkWs.toLowerCase()) { toast('Source and target are the same folder.', 'error'); return; }
    if (apply) {
      const now = Date.now();
      if (now > xmArmedUntil) {
        xmArmedUntil = now + 6000;
        $('#xm-apply-btn').textContent = 'Click again to write into trunk';
        setTimeout(() => { if (Date.now() >= xmArmedUntil && !xmRunning) $('#xm-apply-btn').textContent = 'Apply to Trunk'; }, 6100);
        return;
      }
      xmArmedUntil = 0;
    }
    xmRunning = true;
    const dryBtn = $('#xm-dryrun-btn'), applyBtn = $('#xm-apply-btn');
    dryBtn.disabled = true; applyBtn.disabled = true;
    dryBtn.textContent = apply ? 'Dry Run (report only)' : 'Running…';
    applyBtn.textContent = apply ? 'Applying…' : 'Apply to Trunk';
    log(`Merge to Trunk — ${apply ? 'APPLY' : 'dry run'}: ${s.branchWs} → ${s.trunkWs}`);

    const res = await window.api.xlsxMergeRun({ ...s, apply });

    xmRunning = false;
    dryBtn.disabled = false; dryBtn.textContent = 'Dry Run (report only)'; applyBtn.textContent = 'Apply to Trunk';
    xmLast = res;
    xmRender(res, apply);
    if (res.success) {
      if (apply) {
        xmDryRunKey = '';
        $('#xm-apply-hint').textContent = 'Applied. Review trunk with svn status / svn diff, run coder_convert, then lock and commit.';
        log('Merge to Trunk: apply finished and verified.', 'console-log-success');
        toast('Merged into trunk — review before committing.', 'success');
      } else {
        xmDryRunKey = xmKey(s);
        $('#xm-apply-hint').textContent = 'Dry run complete with these settings — Apply is unlocked. It writes through Excel, verifies, and never commits.';
        log('Merge to Trunk: dry run finished.', 'console-log-success');
        toast('Dry run complete — check the result card.', 'success');
      }
    } else {
      xmDryRunKey = '';
      log(`Merge to Trunk ${apply ? 'apply' : 'dry run'} failed (exit ${res.code}).`, 'console-log-error');
      toast('Merge failed — see the output pane and the report.', 'error');
    }
    xmRefreshApplyGate();
  }
  $('#xm-dryrun-btn').addEventListener('click', () => xmRun(false));
  $('#xm-apply-btn').addEventListener('click', () => xmRun(true));
  $('#xm-reveal-btn').addEventListener('click', async () => {
    if (!xmLast || !xmLast.outDir) return;
    const r = await window.api.xlsxMergeReveal({ outDir: xmLast.outDir });
    if (!r.success) toast(r.output || 'Could not open the report folder.', 'error');
  });
  $('#xm-conflict-filter').addEventListener('change', () => xmRenderConflicts());

  function xmRender(res, apply) {
    $('#xm-result-card').classList.remove('xm-hidden');
    const sm = res.summary || {};
    const t = sm.totals || {};
    $('#xm-result-mode').textContent = apply ? 'apply' : 'dry run';
    $('#xm-res-revs').textContent = sm.base_rev ? `r${sm.base_rev} → r${sm.branch_rev}` : (res.success ? '…' : 'failed');
    $('#xm-res-add').textContent = sm.totals ? (t.add ?? 0) : '…';
    $('#xm-res-mod').textContent = sm.totals ? (t.mod ?? 0) : '…';
    $('#xm-res-new').textContent = sm.totals ? `${(sm.new_files || []).length} / ${t.new_sheets ?? 0} / ${t.new_cols ?? 0}` : '…';
    $('#xm-res-conflicts').textContent = sm.n_conflicts ?? (res.conflicts ? res.conflicts.length : '…');
    $('#xm-res-collisions').textContent = sm.n_collisions ?? (res.suggest ? res.suggest.length : '…');
    let verify = 'n/a (dry run)';
    if (apply) {
      if (Array.isArray(sm.verify_problems)) verify = sm.verify_problems.length ? `${sm.verify_problems.length} problems` : 'OK';
      else verify = res.success ? 'not run' : 'failed';
    }
    $('#xm-res-verify').textContent = verify;
    $('#xm-res-lint').textContent = sm.n_lint == null ? '…' : (sm.n_lint === 0 ? 'clean' : `${sm.n_lint} findings`);
    $('#xm-report').textContent = res.report || res.output || '(no report written)';
    if (res.svnAfter) { $('#xm-svn-after-wrap').classList.remove('xm-hidden'); $('#xm-svn-after').textContent = res.svnAfter; }
    else $('#xm-svn-after-wrap').classList.add('xm-hidden');
    $('#xm-load-collisions-btn').disabled = !(res.suggest && res.suggest.length);

    const sel = $('#xm-conflict-filter');
    const keep = sel.value;
    const conflicts = res.conflicts || [];
    const kinds = [...new Set(conflicts.map(r => r.kind))].sort();
    sel.innerHTML = '';
    const all = document.createElement('option'); all.value = ''; all.textContent = `All conflict kinds (${conflicts.length})`; sel.appendChild(all);
    kinds.forEach(k => {
      const o = document.createElement('option'); o.value = k;
      o.textContent = `${k} (${conflicts.filter(r => r.kind === k).length})`; sel.appendChild(o);
    });
    sel.value = kinds.includes(keep) ? keep : '';
    xmRenderConflicts();
  }

  function xmRenderConflicts() {
    const body = $('#xm-conflicts-body'); body.innerHTML = '';
    const kind = $('#xm-conflict-filter').value;
    const rows = ((xmLast && xmLast.conflicts) || []).filter(r => !kind || r.kind === kind);
    const LIMIT = 300;
    rows.slice(0, LIMIT).forEach(r => {
      const tr = document.createElement('tr');
      [r.xlsx, r.sheet, r.kind, r.key, r.detail].forEach((v, i) => {
        const td = document.createElement('td');
        if (i === 2) { const b = document.createElement('span'); b.className = 'xm-kind'; b.textContent = v; td.appendChild(b); }
        else td.textContent = v;
        tr.appendChild(td);
      });
      body.appendChild(tr);
    });
    $('#xm-conflicts-more').textContent = rows.length > LIMIT
      ? `Showing ${LIMIT} of ${rows.length} — the full list is in conflicts.csv (Open report folder).`
      : (rows.length ? '' : 'No conflicts.');
  }

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
    // The documents panel is filtered by the picked One-Pager, so it has to be
    // told when there stops being one - otherwise it keeps showing one event's
    // documents while the picker says nothing is selected.
    loadArtReqDocs();
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
    // Picking a One-Pager decides which event the documents panel is about.
    artReqShowAllDocs = false;
    loadArtReqDocs();
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
  // How many shipped panels carry art of their own, per event type. Mirrors
  // BLOCKS_UI_PANELS in art_requirement_generator.py - the backend is the
  // authority and plan_passes re-derives it; this is only for the cost line.
  const PANEL_COUNTS = { Blocks: 18 };

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
    // Panels come from the shipped-build table the backend owns, not from the
    // document - a document does not get to choose which panels the game has.
    const npn = PANEL_COUNTS[$('#ideation-event-type').value] || 0;
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
      ['extras', 'Balloon, badge, task icons', nx, nx > 0],
      // A re-skin, not an invention: the previous panel's own art is attached as
      // the reference, so the count is the number of shipped panels that HAVE
      // art of their own. Panels built from the shared chrome have nothing to
      // hold onto and are not rendered.
      ['panel', 'UI panel re-skins', npn, npn > 0]
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
  let artReqShowAllDocs = false;
  let artReqDocFilter = { event: null, hiddenByEvent: 0, hiddenAsDocx: 0 };

  async function loadArtReqDocs() {
    const eventType = $('#ideation-event-type').value;
    const card = $('#artreq-docs-card');
    if (!card) return;
    if (!eventType) { card.style.display = 'none'; return; }
    const res = await window.api.artReqListDocs({
      eventType,
      onePagerName: ideationPickedOnePager ? ideationPickedOnePager.name : '',
      showAll: artReqShowAllDocs,
    });
    if (!res || !res.success) {
      card.style.display = 'none';
      log(`Could not list Art Requirement documents: ${res && res.output}`, 'console-log-error');
      return;
    }
    artReqDocs = res.docs || [];
    artReqDocFilter = { event: res.filteredTo || null,
                        hiddenByEvent: res.hiddenByEvent || 0,
                        hiddenAsDocx: res.hiddenAsDocx || 0,
                        matchedBySource: res.matchedBySource || 0 };
    // Keep the card up when a filter is the reason it is empty, or the panel
    // vanishes and looks like the folder is empty rather than filtered.
    card.style.display = (artReqDocs.length || artReqDocFilter.hiddenByEvent) ? '' : 'none';
    renderArtReqDocList();
  }

  function renderArtReqDocList() {
    const box = $('#artreq-docs-list');
    if (!box) return;
    const f = artReqDocFilter;
    let banner = '';
    if (f.event) {
      const held = [];
      if (f.hiddenByEvent) held.push(`${f.hiddenByEvent} from other events`);
      if (f.hiddenAsDocx) held.push(`${f.hiddenAsDocx} exported .docx`);
      // Say HOW the match was made. "Matched by recorded source" means the
      // documents state which brief they came from; the filename fallback is a
      // guess that happens to be right, and the difference matters when one of
      // them stops working.
      const how = f.matchedBySource
        ? `${f.matchedBySource} matched by recorded source`
        : 'matched by filename';
      banner = `<div class="onepager-empty">Showing event <strong>${escapeHtml(f.event)}</strong>`
        + ` <span class="onepager-meta">(${escapeHtml(how)})</span>`
        + (held.length ? ` — hiding ${held.join(' and ')}.` : '.')
        + ` <a href="#" id="artreq-showall-link">Show everything</a></div>`;
    } else if (f.hiddenAsDocx) {
      banner = `<div class="onepager-empty">Hiding ${f.hiddenAsDocx} exported .docx —`
        + ` each one is reachable from its Markdown row.`
        + (artReqShowAllDocs ? ` <a href="#" id="artreq-showall-link">Filter to this event</a>` : '')
        + `</div>`;
    }
    if (!artReqDocs.length) {
      box.innerHTML = banner + `<div class="onepager-empty">No Art Requirement filed for this`
        + ` event yet.</div>`;
      const link0 = box.querySelector('#artreq-showall-link');
      if (link0) link0.addEventListener('click', toggleArtReqDocFilter);
      return;
    }
    box.innerHTML = banner + artReqDocs.map((d, i) => {
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
    const link = box.querySelector('#artreq-showall-link');
    if (link) link.addEventListener('click', toggleArtReqDocFilter);
  }

  function toggleArtReqDocFilter(e) {
    if (e) e.preventDefault();
    artReqShowAllDocs = !artReqShowAllDocs;
    loadArtReqDocs();
  }


  // ---- generating art for a document that was opened, not just generated ----
  // The Reference Images controls used to exist only on a freshly generated
  // variation card, keyed to artReqVariations[idx]. Open a document filed last
  // week and there was no way to render its art at all - the work had to be
  // regenerated from scratch to get a button back.
  //
  // The structured document is now written as a .json sidecar on every save, so
  // an opened document can drive the same pass picker, the same cost line and
  // the same backend call. The index -1 keeps its DOM ids clear of the generated
  // cards, which are numbered from zero.
  const OPENED_DOC_IDX = -1;
  let artReqOpenedDoc = null;          // the structured variation, when it exists

  function openedDocArtHtml(info) {
    if (!info || !info.variation) {
      return `<p class="settings-hint">${escapeHtml((info && info.reason)
        || 'No structured data beside this document, so its art cannot be generated here.')}</p>`;
    }
    return `
      <div class="board-divider"></div>
      <h3 class="card-title">Reference Art for this document</h3>
      <p class="settings-hint">Renders straight from the document on disk and writes the
        images back into it, so the file the Art team opens is the one carrying the
        references. Pick the passes first — the cost line is the real count.</p>
      ${passPickerHtml(OPENED_DOC_IDX, info.variation)}
      <div class="execution-footer" style="justify-content:flex-end;">
        <button id="artreq-doc-images-btn" class="primary-btn">Generate Reference Images</button>
      </div>`;
  }

  async function generateOpenedDocImages() {
    if (!artReqOpenedDoc || !artReqOpenedDocPath) return;
    const eventType = $('#ideation-event-type').value;
    // Scoped to the view: the button is created when a document is opened, so
    // it is not in index.html and must not be queried as though it were.
    const view = $('#artreq-doc-view');
    const btn = view && view.querySelector('#artreq-doc-images-btn');
    const passes = $$(`.artreq-pass[data-idx="${OPENED_DOC_IDX}"]`)
      .filter(c => c.checked).map(c => c.value);
    const chosen = passes.length ? passes : ['scene'];
    if (btn) { btn.disabled = true; btn.textContent = 'Rendering…'; }

    // The authoritative count, from the backend, logged before a long run starts
    // so the log says what was actually asked for rather than what was intended.
    try {
      const pl = JSON.parse((await window.api.artReqPlan({
        variation: artReqOpenedDoc, passes: chosen })).stdout);
      if (pl.ok) log(`Plan: ${pl.passes.map(p => `${p.pass} ${p.count}`).join(', ')} — ${pl.total} render(s).`);
    } catch (e) { /* informational only; never block the run on it */ }

    const res = await window.api.artReqImages({
      eventType, variation: artReqOpenedDoc, variationIndex: 0, passes: chosen });
    if (btn) { btn.disabled = false; btn.textContent = 'Generate Reference Images'; }
    if (!res.success) {
      log(`Image generation failed: ${res.output}`, 'console-log-error');
      toast('Image generation failed: ' + res.output, 'error');
      return;
    }
    let parsed;
    try { parsed = JSON.parse(res.stdout); }
    catch (e) {
      log(`Could not parse the image response: ${e.message}`, 'console-log-error');
      toast('Could not parse the image response.', 'error');
      return;
    }
    if (!parsed.ok) {
      log(`Image generation failed: ${parsed.error}`, 'console-log-error');
      toast(parsed.error || 'Image generation failed.', 'error');
      return;
    }
    artReqOpenedDoc = parsed.doc || artReqOpenedDoc;
    reportImageReview(parsed.review);
    log(`Rendered ${parsed.generated}/${parsed.requested} image(s).`, 'console-log-success');

    // Write the document back where it came from. Without this the renders sit
    // on disk with nothing referencing them, which looks identical to failure.
    const saved = await window.api.artReqSaveDoc({
      eventType, docPath: artReqOpenedDocPath, variation: artReqOpenedDoc });
    if (!saved.success) {
      log(`Rendered, but saving the document failed: ${saved.output}`, 'console-log-error');
      toast('Images rendered, but the document could not be updated: ' + saved.output, 'error');
      return;
    }
    log(`Updated ${saved.data.path} — ${saved.data.zones_with_images} zone(s) now carry an image.`,
        'console-log-success');
    toast(`Rendered ${parsed.generated} image(s) into the document.`, 'success');

    // Reopen so the reader sees the images that were just written.
    const path = artReqOpenedDocPath;
    artReqOpenedDocPath = '';
    const i = artReqDocs.findIndex(d => d.path === path);
    if (i >= 0) await toggleArtReqDoc(i);
    await loadArtReqDocs();
  }
  async function toggleArtReqDoc(i) {
    const doc = artReqDocs[i];
    const view = $('#artreq-doc-view');
    if (!doc || !view) return;
    if (artReqOpenedDocPath === doc.path) {
      artReqOpenedDocPath = '';
      artReqOpenedDoc = null;
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
    const dataRes = await window.api.artReqDocData({ eventType, docPath: doc.path });
    const info = (dataRes && dataRes.success) ? dataRes.data : null;
    artReqOpenedDoc = (info && info.variation) || null;
    view.innerHTML = renderMarkdownDoc(res.data.markdown, -1 - i).html
                   + openedDocArtHtml(info);
    const gen = view.querySelector('#artreq-doc-images-btn');
    if (gen) gen.addEventListener('click', generateOpenedDocImages);
    view.querySelectorAll('.artreq-pass').forEach(cb =>
      cb.addEventListener('change', () => refreshPlan(OPENED_DOC_IDX)));
    refreshPlan(OPENED_DOC_IDX);
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
  // ---------------- Localization ----------------
  // Paste English, get the six shipped languages back. This panel deliberately
  // does NOT write into localization.xlsx — the designer copies rows out and
  // pastes them where they belong, which keeps step 14 the only code in this
  // app that edits that workbook. Column order in the copied TSV matches the
  // workbook's own B..H order so a paste lands in the right columns.
  const LOC_LANGS = [['DE', 'German'], ['ES', 'Spanish'], ['NL', 'Dutch'],
                     ['SV', 'Swedish'], ['FR', 'French'], ['DA', 'Danish']];
  let locResult = null;

  // One row per string, added and removed individually. A single textarea was
  // fine for typing and bad for editing: changing string 12 of 30 meant
  // finding the right line in a wall of text. Rows keep paste working though -
  // see the paste handler, which splits a pasted block back out into rows
  // rather than dumping 30 lines into one field.
  function locRowHtml(value) {
    return '<div class="loc-row">'
      + `<textarea class="text-input mono loc-row-input" spellcheck="false" rows="1" `
      + `placeholder="English string">${escapeHtml(value || '')}</textarea>`
      + '<button class="loc-row-del" type="button" title="Remove this row" '
      + 'aria-label="Remove this row">&minus;</button></div>';
  }

  // Grow to fit rather than scroll: a two-line string that only shows its first
  // line is how you end up translating half of it without noticing.
  function locAutoGrow(el) {
    if (!el) return;
    el.style.height = 'auto';
    el.style.height = Math.min(el.scrollHeight + 2, 260) + 'px';
  }

  function renderLocRows(values) {
    const list = (values && values.length) ? values : [''];
    $('#loc-rows').innerHTML = list.map(locRowHtml).join('');
    $$('.loc-row-input').forEach(locAutoGrow);
  }

  function locRowValues() {
    return $$('.loc-row-input').map((el) => el.value);
  }

  function locStrings() {
    return locRowValues().map((s) => s.trim()).filter(Boolean);
  }

  function locAddRow(after) {
    const row = document.createElement('div');
    row.className = 'loc-row';
    row.innerHTML = locRowHtml('').replace(/^<div class="loc-row">|<\/div>$/g, '');
    if (after && after.parentNode) after.insertAdjacentElement('afterend', row);
    else $('#loc-rows').appendChild(row);
    const input = row.querySelector('.loc-row-input');
    if (input) input.focus();
    return row;
  }

  $('#loc-add-row').addEventListener('click', () => locAddRow());

  $('#loc-rows').addEventListener('click', (evt) => {
    const del = evt.target.closest('.loc-row-del');
    if (!del) return;
    const rows = $$('.loc-row');
    // Never leave the panel with nowhere to type: the last row empties instead
    // of disappearing.
    if (rows.length <= 1) {
      const only = rows[0] && rows[0].querySelector('.loc-row-input');
      if (only) { only.value = ''; only.focus(); }
      return;
    }
    del.closest('.loc-row').remove();
  });

  // A row is a textarea, so Enter inserts a newline — a single string can span
  // lines, which 29 of the English strings in localization.xlsx genuinely do.
  // Ctrl+Enter is the shortcut for the next row.
  $('#loc-rows').addEventListener('keydown', (evt) => {
    if (!evt.target.classList.contains('loc-row-input')) return;
    if (evt.key === 'Enter' && (evt.ctrlKey || evt.metaKey)) {
      evt.preventDefault();
      locAddRow(evt.target.closest('.loc-row'));
    }
  });

  $('#loc-rows').addEventListener('input', (evt) => {
    if (evt.target.classList.contains('loc-row-input')) locAutoGrow(evt.target);
  });

  // Paste is deliberately literal — it does NOT split lines into rows.
  //
  // It used to. That is wrong for the case that matters: a multi-line string
  // pasted in would be torn into fragments and each fragment translated on its
  // own, out of context, with the damage invisible in a table that looks
  // perfectly normal. Pasting a bulk list into one row is the opposite kind of
  // mistake — obvious on sight and one click to fix with "Split lines". Given
  // one silent failure and one loud one, the default takes the loud one.

  $('#loc-split-rows').addEventListener('click', () => {
    const before = locRowValues();
    const values = [];
    before.forEach((value) => {
      const parts = value.split(/\r?\n/).map((s) => s.trim()).filter(Boolean);
      if (parts.length > 1) values.push(...parts);
      else values.push(value);
    });
    const gained = values.length - before.length;
    if (gained <= 0) { toast('Nothing to split — no row has more than one line.'); return; }
    renderLocRows(values);
    toast(`Split into ${values.length} rows.`, 'success');
  });

  function renderLocLangPicker() {
    $('#loc-lang-picker').innerHTML = LOC_LANGS.map(([code, name]) => (
      `<label class="loc-lang-chip" title="${escapeHtml(name)}">` +
      `<input type="checkbox" class="loc-lang-box" value="${code}" checked />` +
      `<span>${code}</span></label>`
    )).join('');
  }

  // Unticking everything is a slip, not a request for nothing — fall back to
  // all six rather than silently translating into no languages at all.
  function selectedLocLangs() {
    const picked = $$('.loc-lang-box').filter((b) => b.checked).map((b) => b.value);
    return picked.length ? picked : LOC_LANGS.map(([code]) => code);
  }

  async function copyText(text, label) {
    try {
      await navigator.clipboard.writeText(text);
      toast(`${label} copied to the clipboard.`, 'success');
      return;
    } catch (e) {
      // Clipboard API can be refused depending on focus; the textarea route
      // still works and is the difference between the button doing something
      // and doing nothing.
      const scratch = document.createElement('textarea');
      scratch.value = text;
      document.body.appendChild(scratch);
      scratch.select();
      const ok = document.execCommand('copy');
      scratch.remove();
      toast(ok ? `${label} copied to the clipboard.` : 'Could not copy — select the table and copy manually.',
        ok ? 'success' : 'error');
    }
  }

  // A field carrying a real line break, a tab or a quote has to be quoted or
  // the paste tears one row into several. Excel reads the doubled-quote form.
  function tsvField(value) {
    const text = String(value == null ? '' : value);
    return /[\t\r\n"]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text;
  }

  function locTsv() {
    if (!locResult) return '';
    const langs = locResult.languages;
    const header = ['English', ...langs].join('\t');
    const body = locResult.rows.map((row) => [
      row.en,
      ...langs.map((code) => (row.targets[code] || {}).text || ''),
    ].map(tsvField).join('\t'));
    return [header, ...body].join('\r\n');
  }

  function renderLocResults() {
    const host = $('#loc-results');
    if (!locResult) { host.innerHTML = ''; return; }
    const langs = locResult.languages;
    const counts = locResult.counts || {};
    const summary = [
      `${locResult.rows.length} string(s)`,
      `${counts.approved || 0} reused from translator-approved text`,
      `${counts.memory || 0} reused from what we already ship`,
      `${counts.generated || 0} newly translated`,
    ];
    if (counts.failed) summary.push(`${counts.failed} not translated`);

    const head = ['<th>English</th>', ...langs.map((c) => `<th>${c}</th>`)].join('');
    const body = locResult.rows.map((row) => {
      const cells = langs.map((code) => {
        const cell = row.targets[code] || { text: '', source: 'failed', warnings: [] };
        const warns = (cell.warnings || []).map(
          (w) => `<div class="loc-warn">${escapeHtml(w)}</div>`).join('');
        const badgeClass = `loc-src loc-src-${cell.source}`;
        const note = cell.note ? ` title="${escapeHtml(cell.note)}"` : '';
        return `<td class="loc-cell"><div class="loc-cell-text">${escapeHtml(cell.text) || '—'}</div>` +
          `<span class="${badgeClass}"${note}>${cell.source}</span>${warns}</td>`;
      }).join('');
      // The English lands in column B of localization.xlsx like every other
      // column, so a character the convert cannot carry has to show here too -
      // the per-language warnings above never look at it.
      const enWarns = (row.warnings || []).map(
        (w) => `<div class="loc-warn">${escapeHtml(w)}</div>`).join('');
      return `<tr><td class="loc-en-cell">${escapeHtml(row.en)}${enWarns}</td>${cells}</tr>`;
    }).join('');

    // One banner for a batch-level failure. Without it the only signal is a
    // column of FAILED badges, which says that nothing worked but not why or
    // what to do about it.
    const banner = locResult.provider_error
      ? `<div class="loc-banner"><strong>No model was reachable</strong>` +
        `<span class="loc-banner-reason">${escapeHtml(locResult.provider_error)}</span>` +
        `<span>${escapeHtml(locResult.provider_hint || '')}</span></div>`
      : '';

    host.innerHTML =
      `<div class="stacked-card"><h2 class="card-title">Translations</h2>` +
      banner +
      `<div class="loc-summary">${escapeHtml(summary.join('  ·  '))}</div>` +
      `<div class="loc-actions">` +
      `<button class="secondary-btn small" data-loc-copy="1">Copy all as TSV</button>` +
      `<button class="secondary-btn small" data-loc-xlsx="1">Export to Excel Sheet</button>` +
      `</div>` +
      `<div class="loc-table-wrap"><table class="loc-table">` +
      `<thead><tr>${head}</tr></thead><tbody>${body}</tbody></table></div></div>`;

    host.querySelector('[data-loc-copy]').addEventListener('click', () => {
      copyText(locTsv(), `${locResult.rows.length} row(s)`);
    });

    host.querySelector('[data-loc-xlsx]').addEventListener('click', async (evt) => {
      const btn = evt.currentTarget;
      btn.disabled = true; btn.textContent = 'Exporting...';
      const stamp = new Date().toISOString().slice(0, 10);
      const res = await window.api.locExportXlsx({
        rows: locResult.rows,
        languages: locResult.languages,
        suggestedName: `Skip-Bo translations ${stamp}.xlsx`,
      });
      btn.disabled = false; btn.textContent = 'Export to Excel Sheet';
      if (res.canceled) return;
      if (!res.success) {
        log(`Excel export failed: ${res.output}`, 'console-log-error');
        toast('Excel export failed - see the console.', 'error');
        return;
      }
      let parsed;
      try { parsed = JSON.parse(res.stdout); }
      catch (e) { toast('Could not read the export result.', 'error'); return; }
      if (parsed.error) { toast(parsed.error, 'error'); return; }
      log(`Exported ${parsed.rows} row(s) to ${parsed.path}. ${parsed.note}`
        + (parsed.flagged ? ` ${parsed.flagged} cell(s) carry checks - see the `
          + `Provenance sheet.` : ''), 'console-log-success');
      toast('Exported to Excel.', 'success');
    });
  }

  // The no-key route. Same prompt the MCP server hands Claude Desktop, only
  // delivered by clipboard, so a designer without an Anthropic key still gets
  // the memory, the neighbours and the translator's corrections working for
  // them instead of translating from a blank page.
  $('#loc-brief-btn').addEventListener('click', async () => {
    const strings = locStrings();
    if (!strings.length) { toast('Paste at least one English string first.', 'error'); return; }

    const btn = $('#loc-brief-btn');
    btn.disabled = true; btn.textContent = 'Building...';
    const res = await window.api.locBrief({ strings, languages: selectedLocLangs() });
    btn.disabled = false; btn.textContent = 'Copy brief for Claude';

    if (!res.success) {
      log(`Could not build the brief: ${res.output}`, 'console-log-error');
      toast('Could not build the brief — see the console.', 'error');
      return;
    }
    let parsed;
    try { parsed = JSON.parse(res.stdout); }
    catch (e) { toast('Could not parse the brief.', 'error'); return; }
    if (parsed.error) { toast(parsed.error, 'error'); return; }

    await copyText(parsed.text, 'Brief');
    const c = parsed.counts || {};
    log(`Brief built for ${c.strings} string(s): ${c.settled} already settled by memory, `
      + `${c.gaps} needing translation. Paste it into Claude Desktop, then paste the `
      + `table back — or ask Claude Desktop to run localization_brief itself.`,
      'console-log-success');
  });

  // Queue the batch, bring Claude Desktop up, and watch for the answer. The
  // trigger phrase goes on the clipboard because no claude:// route can carry
  // it — the app's routes are claude://claude, resume, code/*, and
  // cowork/shared-artifact, none of which take text.
  // A queued batch is answered by a person working in Claude Desktop, so the
  // wait is theirs, not a machine's. A real 124-string batch came back after
  // 33 minutes; the old ceiling of 300 polls x 2s gave up at 10 and the
  // finished work — all 744 cells — was never shown. Poll for two hours, and
  // slow down rather than hammering the synced folder.
  const LOC_POLL_LIMIT_MS = 2 * 60 * 60 * 1000;
  function locPollDelay(elapsedMs) {
    if (elapsedMs < 60000) return 2000;          // first minute: responsive
    if (elapsedMs < 10 * 60000) return 5000;
    return 15000;
  }
  let locPollTimer = null;
  // Persisted, not just in-memory: otherwise every restart re-offers a batch
  // the designer has already loaded, and a banner that cries wolf gets ignored
  // exactly when it finally matters.
  const LOC_SEEN_KEY = 'loc.lastLoadedJobId';
  let locLastJobId = (() => {
    try { return localStorage.getItem(LOC_SEEN_KEY) || null; } catch (e) { return null; }
  })();
  function locRememberJob(jobId) {
    locLastJobId = jobId;
    try { localStorage.setItem(LOC_SEEN_KEY, jobId); } catch (e) { /* private mode */ }
  }

  function locStopPolling() {
    if (locPollTimer) { clearTimeout(locPollTimer); locPollTimer = null; }
  }

  function locWaitingBanner(jobId, phrase) {
    const host = $('#loc-results');
    const existing = host.querySelector('.loc-banner-wait');
    const html =
      `<div class="loc-banner loc-banner-wait"><strong>Waiting for Claude Desktop</strong>` +
      `<span>Batch <code>${escapeHtml(jobId)}</code> is queued. Claude Desktop should be ` +
      `open — paste the line below (it is already on your clipboard) and send it. ` +
      `This table fills in by itself when it comes back.</span>` +
      `<span class="loc-banner-reason">${escapeHtml(phrase)}</span>` +
      `<span class="loc-wait-elapsed">Waiting…</span></div>`;
    if (existing) existing.outerHTML = html;
    else host.insertAdjacentHTML('afterbegin', html);
  }

  async function handOffToClaudeDesktop(strings, languages) {
    const btn = $('#loc-translate-btn');
    const queued = await window.api.locQueue({ strings, languages });
    if (!queued.success) {
      log(queued.output, 'console-log-error');
      toast(queued.output, 'error');
      btn.disabled = false; btn.textContent = 'Translate';
      return;
    }

    const phrase = 'Do the pending localization batch.';
    // Not awaited: the clipboard is a convenience, and navigator.clipboard can
    // reject or stall when the window is not focused. The phrase is shown in
    // the banner too, so a failed copy must not hold up the handoff.
    copyText(phrase, 'Trigger phrase');
    await window.api.locOpenClaudeDesktop();
    locWaitingBanner(queued.jobId, phrase);
    btn.disabled = true; btn.textContent = 'Waiting for Claude…';
    log(`Batch ${queued.jobId} queued for Claude Desktop. Send "${phrase}" there — `
      + `it will call localization_pending and the table will land back here.`,
      'console-log-success');

    const startedAt = Date.now();
    const poll = async () => {
      locPollTimer = null;
      const res = await window.api.locQueueResult({ jobId: queued.jobId });
      if (res.success && res.ready) {
        locResult = res.result;
        locRememberJob(queued.jobId);
        renderLocResults();
        btn.disabled = false; btn.textContent = 'Translate';
        const c = res.result.counts || {};
        const mins = Math.round((Date.now() - startedAt) / 60000);
        log(`Batch ${queued.jobId} came back after ${mins} min: `
          + `${c.generated || 0} newly translated, ${c.approved || 0} approved, `
          + `${c.memory || 0} already shipped.`, 'console-log-success');
        toast('Translations came back from Claude Desktop.', 'success');
        return;
      }
      const elapsed = Date.now() - startedAt;
      if (elapsed >= LOC_POLL_LIMIT_MS) {
        btn.disabled = false; btn.textContent = 'Translate';
        log(`Stopped watching batch ${queued.jobId} after 2 hours. Nothing is lost — `
          + `when it finishes, reopen this tab and the panel will offer it.`,
          'console-log-line');
        return;
      }
      const waiting = document.querySelector('.loc-banner-wait');
      if (waiting) {
        const mins = Math.floor(elapsed / 60000);
        const note = waiting.querySelector('.loc-wait-elapsed');
        if (note) note.textContent = mins < 1 ? 'Waiting…' : `Waiting ${mins} min…`;
      }
      locPollTimer = setTimeout(poll, locPollDelay(elapsed));
    };
    locPollTimer = setTimeout(poll, locPollDelay(0));
  }

  // Live progress from the backend. Without it a large batch shows
  // "Translating..." and nothing else, and slow is indistinguishable from hung.
  window.api.onLocProgress((p) => {
    const btn = $('#loc-translate-btn');
    if (!btn.disabled) return;                 // a stale line after we finished
    if (p.stage === 'generating' && p.total > 1) {
      btn.textContent = `Translating ${p.done}/${p.total}...`;
    } else {
      btn.textContent = 'Translating...';
    }
    $('#loc-memory-note').textContent = p.detail || '';
  });

  // Offer a batch that finished while nobody was watching.
  //
  // The panel used to stop polling after 10 minutes, and a result that landed
  // later was unreachable forever even though it was sitting on disk. Checking
  // on open means closing the app, or walking away, no longer loses the work.
  async function locOfferFinishedBatch() {
    const res = await window.api.locLatestResult({ since: locLastJobId });
    if (!res.success || !res.found) return;
    const rows = (res.result.rows || []).length;
    if (!rows) return;
    const when = new Date(res.mtime).toLocaleString();
    const host = $('#loc-results');
    if (host.querySelector('.loc-recover')) return;
    host.insertAdjacentHTML('afterbegin',
      `<div class="loc-banner loc-recover"><strong>A finished batch is waiting</strong>` +
      `<span>${rows} string(s), filed ${escapeHtml(when)}. It came back after this ` +
      `panel stopped watching.</span>` +
      `<span><button class="secondary-btn small" data-loc-load="1">Load it</button></span></div>`);
    host.querySelector('[data-loc-load]').addEventListener('click', () => {
      locResult = res.result;
      locRememberJob(res.jobId);
      renderLocResults();
      const banner = host.querySelector('.loc-recover');
      if (banner) banner.remove();
      log(`Loaded batch ${res.jobId}: ${rows} string(s).`, 'console-log-success');
    });
  }

  $('#loc-translate-btn').addEventListener('click', async () => {
    locStopPolling();
    const strings = locStrings();
    if (!strings.length) { toast('Paste at least one English string first.', 'error'); return; }

    if (!activeBranchRoot()) { toast('Set the branch path in Settings first.', 'error'); return; }

    // No designPath is sent on purpose: main.js pins the memory to WorkSpcae3
    // (the master the corrections land in) rather than following the workspace
    // picker, so switching workspace for unrelated work cannot change what this
    // panel considers already-approved.
    const btn = $('#loc-translate-btn');
    btn.disabled = true; btn.textContent = 'Translating...';
    const res = await window.api.locTranslate({
      strings, languages: selectedLocLangs(),
    });
    btn.disabled = false; btn.textContent = 'Translate';

    if (!res.success) {
      log(`Translation failed: ${res.output}`, 'console-log-error');
      toast('Translation failed — see the console.', 'error');
      return;
    }
    let parsed;
    try { parsed = JSON.parse(res.stdout); }
    catch (e) {
      log(`Could not parse the translation response: ${res.output}`, 'console-log-error');
      toast('Could not parse the translation response.', 'error');
      return;
    }
    if (parsed.error) {
      log(`Translation error: ${parsed.error}`, 'console-log-error');
      toast(parsed.error, 'error');
      return;
    }

    locResult = parsed;
    renderLocResults();

    const mem = parsed.memory || {};
    $('#loc-memory-note').textContent =
      `${(mem.approved_entries || 0).toLocaleString()} approved · ` +
      `${(mem.shipped_entries || 0).toLocaleString()} shipped · ` +
      `memory ${mem.snapshot === 'cache' ? 'from cache' : 'rebuilt'}`;
    if (!mem.approved_entries) {
      log('Note: no translator-approved memory found — run loc_corrections_import.py '
        + 'against the LOC Corrections folder to populate it.', 'console-log-line');
    }
    // No key and nothing reachable: hand the batch to Claude Desktop instead of
    // stopping at an error. One message over there is the only manual step.
    if (parsed.provider_error) {
      log(`No model reachable: ${parsed.provider_error}`, 'console-log-error');
      await handOffToClaudeDesktop(strings, selectedLocLangs());
      return;
    }
    const counts = parsed.counts || {};
    log(`Translated ${parsed.rows.length} string(s): ${counts.approved || 0} approved, `
      + `${counts.memory || 0} already shipped, ${counts.generated || 0} generated `
      + `via ${parsed.provider === 'ollama' ? 'local Ollama' : 'Claude'}.`, 'console-log-success');
  });

  async function init() {
    const data = await window.api.getConfig();
    cfg = data.cfg;
    branchKeys = data.branchKeys;
    workspaces = data.workspaces;
    defaults = data.defaults;
    pipelineSteps = data.pipelineSteps;

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
    renderLocLangPicker();
    renderLocRows();
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
