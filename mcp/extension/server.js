#!/usr/bin/env node
// Claude Desktop extension entry point: a shim, not a copy.
//
// Why a shim. The real server resolves several things RELATIVE to its own
// location - the bundled Python interpreter and CLI under pybackend_dist, and
// localization-app.html next to it. Copying it into the extension bundle would
// break all of those, and would also fork the file so a fix in the repo would
// silently not reach Claude Desktop. Pointing at the checkout keeps one copy.
//
// Consequence worth knowing: this bundle is machine-specific. It locates the
// checkout rather than carrying it, so a designer whose SkipBo lives elsewhere
// sets EVENT_AUTOMATOR_MCP to their own mcp/server.js path.

'use strict';

const fs = require('fs');
const path = require('path');

const RELATIVE = path.join('tools', 'EventAutomationDesktop', 'mcp', 'server.js');

function candidates() {
  const found = [];
  const fromEnv = (process.env.EVENT_AUTOMATOR_MCP || '').trim();
  if (fromEnv) found.push(fromEnv);
  // The usual checkout, then the other drives a checkout tends to land on.
  for (const root of ['C:/SkipBo', 'D:/SkipBo', 'E:/SkipBo',
                      path.join(process.env.USERPROFILE || '', 'SkipBo')]) {
    found.push(path.join(root, RELATIVE));
  }
  return found;
}

const target = candidates().find((p) => {
  try { return fs.statSync(p).isFile(); } catch (e) { return false; }
});

if (!target) {
  // stderr, not stdout: stdout is the JSON-RPC channel and anything else on it
  // corrupts the stream. Claude Desktop surfaces this in the server's log.
  process.stderr.write(
    'Event Automator: could not find the checkout. Looked for:\n'
    + candidates().map((p) => '  ' + p).join('\n')
    + '\nSet EVENT_AUTOMATOR_MCP to the full path of '
    + 'tools/EventAutomationDesktop/mcp/server.js and restart Claude Desktop.\n');
  process.exit(1);
}

require(target);
