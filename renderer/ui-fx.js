// ui-fx.js — purely additive motion & HUD layer for the Event Automator.
//
// Contract: this file never touches app state, never calls window.api, and
// never changes any element app.js reads or writes. It only:
//   1. floats decorative card-suit glyphs behind the UI (an aria-hidden layer),
//   2. renders an XP-style "quest progress" bar for the Main Events pipeline,
//      derived by observing the .status-* classes app.js already sets,
//   3. lets the docked console be pinned open with a click on its title, and
//      auto-lights it while any pipeline step is running.
// If anything here throws, the app is unaffected — every entry point is wrapped.
(() => {
  'use strict';
  const reduceMotion = window.matchMedia && window.matchMedia('(prefers-reduced-motion: reduce)').matches;
  const $ = (s, r) => (r || document).querySelector(s);
  const $$ = (s, r) => Array.from((r || document).querySelectorAll(s));
  const safe = (fn) => { try { fn(); } catch (e) { console.warn('[ui-fx]', e); } };

  // ---- 1. drifting suit glyphs -------------------------------------------
  safe(() => {
    if (reduceMotion) return;
    const layer = document.createElement('div');
    layer.className = 'fx-suits';
    layer.setAttribute('aria-hidden', 'true');
    const suits = ['♠', '♥', '♦', '♣'];
    for (let i = 0; i < 16; i++) {
      const s = document.createElement('span');
      s.className = 'fx-suit';
      s.textContent = suits[i % 4];
      // Spread across the width; vary size, speed and phase so it never loops visibly.
      s.style.left = (3 + Math.random() * 94).toFixed(1) + '%';
      s.style.fontSize = (12 + Math.random() * 26).toFixed(0) + 'px';
      s.style.animationDuration = (26 + Math.random() * 30).toFixed(1) + 's';
      s.style.animationDelay = (-Math.random() * 40).toFixed(1) + 's';
      s.style.setProperty('--fx-drift', ((Math.random() - 0.5) * 120).toFixed(0) + 'px');
      layer.appendChild(s);
    }
    document.body.appendChild(layer);
  });

  // ---- 2. quest progress bar for the pipeline checklist -------------------
  safe(() => {
    const list = $('#checklist-container');
    if (!list) return;
    const card = list.closest('.stacked-card');
    if (!card) return;
    const title = $('.card-title', card);

    const wrap = document.createElement('div');
    wrap.className = 'quest-progress';
    wrap.setAttribute('aria-hidden', 'true');
    wrap.innerHTML =
      '<div class="quest-progress-head">' +
        '<span class="quest-progress-label">Quest progress</span>' +
        '<span class="quest-progress-count">0 / 0</span>' +
      '</div>' +
      '<div class="quest-progress-track"><div class="quest-progress-fill"></div><div class="quest-progress-ticks"></div></div>';
    if (title && title.nextSibling) card.insertBefore(wrap, title.nextSibling); else card.insertBefore(wrap, list);

    const fill = $('.quest-progress-fill', wrap);
    const count = $('.quest-progress-count', wrap);
    const ticks = $('.quest-progress-ticks', wrap);
    const consoleSection = $('.console-section');
    let timer = 0;

    const update = () => {
      timer = 0;
      const items = $$('.checklist-item', list);
      const total = items.length;
      const done = $$('.status-success', list).length;
      const running = $$('.status-running', list).length;
      const failed = $$('.status-error', list).length;
      const pct = total ? Math.round((done / total) * 100) : 0;
      fill.style.width = pct + '%';
      count.textContent = total ? `${done} / ${total}` : '—';
      wrap.dataset.state = failed ? 'error' : running ? 'running' : (total && done === total) ? 'complete' : 'idle';
      // One tick per step, so the bar reads as segments rather than a slider.
      if (ticks.childElementCount !== total) {
        ticks.innerHTML = '';
        for (let i = 1; i < total; i++) {
          const t = document.createElement('i');
          t.style.left = ((i / total) * 100).toFixed(2) + '%';
          ticks.appendChild(t);
        }
      }
      if (consoleSection) consoleSection.classList.toggle('console-live', running > 0);
    };
    const schedule = () => { if (!timer) timer = setTimeout(update, 16); };
    new MutationObserver(schedule).observe(list, { childList: true, subtree: true, attributes: true, attributeFilter: ['class'] });
    schedule();
  });

  // ---- 3. console pin ------------------------------------------------------
  safe(() => {
    const section = $('.console-section');
    const title = $('.console-title', section || document);
    if (!section || !title) return;
    title.setAttribute('role', 'button');
    title.setAttribute('tabindex', '0');
    title.title = 'Click to pin the log open';
    const toggle = () => section.classList.toggle('console-pinned');
    title.addEventListener('click', toggle);
    title.addEventListener('keydown', (e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); toggle(); } });
  });
})();
