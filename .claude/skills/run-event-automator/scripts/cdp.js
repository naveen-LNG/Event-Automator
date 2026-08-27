// Zero-dependency Chrome DevTools Protocol driver for the Electron app.
//
// The run skill's Electron recipe assumes Linux (xvfb + tmux + Playwright's
// _electron). This box is Windows and the project has no Playwright, so rather
// than add a heavy dev dependency just to click a checkbox, this talks CDP
// directly - Node 22 ships a global WebSocket and fetch, which is all it needs.
//
// Usage:  node cdp.js <command> [json-arg]
//   targets                  list debuggable targets
//   eval '<expr>'            evaluate in the renderer, print the result
//   shot <file>              screenshot the window
//   click '<selector>'       click the first match
//   text '<selector>'        textContent of the first match

const PORT = process.env.CDP_PORT || 9222;

async function pageTarget() {
  const res = await fetch(`http://127.0.0.1:${PORT}/json`);
  const list = await res.json();
  // The renderer is the 'page' target that is not devtools itself.
  const t = list.find(x => x.type === 'page' && !/devtools:/.test(x.url));
  if (!t) throw new Error('no renderer target yet. targets: ' + JSON.stringify(list.map(x => [x.type, x.url])));
  return t;
}

function connect(wsUrl) {
  return new Promise((resolve, reject) => {
    const ws = new WebSocket(wsUrl);
    let id = 0;
    const pending = new Map();
    ws.onmessage = (ev) => {
      const msg = JSON.parse(ev.data);
      if (msg.id && pending.has(msg.id)) {
        const { resolve: r, reject: j } = pending.get(msg.id);
        pending.delete(msg.id);
        msg.error ? j(new Error(JSON.stringify(msg.error))) : r(msg.result);
      }
    };
    ws.onerror = (e) => reject(new Error('ws error: ' + (e.message || e.type)));
    ws.onopen = () => resolve({
      send(method, params = {}) {
        return new Promise((r, j) => {
          const n = ++id;
          pending.set(n, { resolve: r, reject: j });
          ws.send(JSON.stringify({ id: n, method, params }));
        });
      },
      close() { ws.close(); }
    });
  });
}

async function evaluate(cli, expression) {
  const r = await cli.send('Runtime.evaluate', {
    expression, returnByValue: true, awaitPromise: true, userGesture: true
  });
  if (r.exceptionDetails) {
    throw new Error(r.exceptionDetails.exception?.description || JSON.stringify(r.exceptionDetails));
  }
  return r.result?.value;
}

(async () => {
  const [cmd, arg] = process.argv.slice(2);
  if (cmd === 'targets') {
    const res = await fetch(`http://127.0.0.1:${PORT}/json`);
    console.log(JSON.stringify(await res.json(), null, 2));
    return;
  }
  const t = await pageTarget();
  const cli = await connect(t.webSocketDebuggerUrl);
  try {
    if (cmd === 'eval') {
      console.log(JSON.stringify(await evaluate(cli, arg), null, 2));
    } else if (cmd === 'text') {
      console.log(await evaluate(cli, `(document.querySelector(${JSON.stringify(arg)})||{}).textContent || '<no match>'`));
    } else if (cmd === 'click') {
      console.log(await evaluate(cli,
        `(() => { const el = document.querySelector(${JSON.stringify(arg)});
           if (!el) return 'NO MATCH'; el.click(); return 'clicked: ' + (el.textContent||el.value||el.className).trim().slice(0,60); })()`));
    } else if (cmd === 'shot') {
      const { data } = await cli.send('Page.captureScreenshot', { format: 'png' });
      require('fs').writeFileSync(arg, Buffer.from(data, 'base64'));
      console.log('wrote ' + arg);
    } else {
      console.log('unknown command: ' + cmd);
    }
  } finally {
    cli.close();
  }
})().catch(e => { console.error('ERROR: ' + e.message); process.exit(1); });
