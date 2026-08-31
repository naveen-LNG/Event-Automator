#!/usr/bin/env node
// Event Automator as an MCP server, so a designer can write an Art Requirement
// in Claude Desktop and have it land, rendered and exported, in the shared
// folder without an Anthropic API key and without leaving the conversation.
//
// WHY THIS EXISTS, and why it is not "paste the prompt into Claude":
//
// A Blocks Art Requirement once shipped carrying Merge's mechanics - a buildings
// table with 1-3 upgrade levels and A/B/C areas, none of which Blocks has. The
// fix was not a better prompt. It was making the schema unable to express the
// foreign mechanic: `buildings` capped at maxItems 0 for types that have none,
// and max_level deleted outright rather than discouraged. That works because the
// Anthropic API CONSTRAINS the decoder to the schema.
//
// A copy-paste flow throws that away and goes back to asking nicely, which is
// the exact failure mode that caused the bug. An MCP tool call does not: the
// host constrains a call to the tool's declared input schema. So this server
// declares ONE SAVE TOOL PER EVENT TYPE, each carrying that type's own schema
// straight from schema_for() in Python. The Blocks tool has no max_level in it,
// so it cannot be called with one.
//
// That costs context - five schemas is roughly 53k characters in the tool list -
// and it is a deliberate trade. The alternative, one generic tool validated
// after the fact, would let the model spend a whole document getting it wrong
// before being told. Server-side validation runs anyway as a second net, since
// a host that does not enforce input schemas would otherwise fail silently.
//
// Zero dependencies, like the rest of this codebase: MCP over stdio is JSON-RPC
// 2.0 with Content-Length framing, and Node has everything that needs.

'use strict';

const { execFile } = require('child_process');
const fs = require('fs');
const path = require('path');

const ROOT = path.join(__dirname, '..');
const PYTHON = path.join(ROOT, 'pybackend_dist', 'python', 'python.exe');
const CLI = path.join(ROOT, 'pybackend_dist', 'python', 'pybackend_cli.py');
const PROTOCOL_VERSION = '2024-11-05';

// ---------------------------------------------------------------------------
// Config: the shared ideationData folder, read from the desktop app's own
// config so the two always agree about where documents live. Falling back to
// the bundled copy would silently write into a folder nobody reads - which has
// already happened once in this project.
function loadConfig() {
  const p = path.join(process.env.APPDATA || '', 'event-automation-desktop', 'config.json');
  try { return JSON.parse(fs.readFileSync(p, 'utf-8')); } catch { return {}; }
}

function eventsRoot() {
  const cfg = loadConfig();
  const configured = (cfg.ideationDataPath || '').trim();
  const base = configured || path.join(ROOT, 'ideationData');
  return path.join(base, 'Events');
}

function authorName() {
  return (loadConfig().authorName || '').trim();
}

// A path from a tool call is not a licence to touch anywhere on disk - the same
// rule the Electron IPC handlers follow. Resolve first, then check, so '..'
// cannot walk out.
function insideEvents(p) {
  const base = path.resolve(eventsRoot());
  const resolved = path.resolve(p);
  return resolved === base || resolved.startsWith(base + path.sep) ? resolved : null;
}

function runPython(scriptId, payload) {
  return new Promise((resolve, reject) => {
    const tmp = path.join(
      process.env.TEMP || '.',
      `mcp_${scriptId}_${process.pid}_${Date.now()}.json`);
    fs.writeFileSync(tmp, JSON.stringify(payload), 'utf-8');
    execFile(PYTHON, [CLI, scriptId, tmp], { maxBuffer: 32 * 1024 * 1024 },
      (err, stdout, stderr) => {
        fs.unlink(tmp, () => {});
        if (err && !stdout) return reject(new Error(stderr || err.message));
        try {
          const parsed = JSON.parse(stdout);
          if (parsed.error) return reject(new Error(parsed.error));
          resolve(parsed);
        } catch (e) {
          reject(new Error(`Could not parse ${scriptId} output: ${e.message}`));
        }
      });
  });
}

function runPythonArgs(scriptId, args) {
  return new Promise((resolve, reject) => {
    execFile(PYTHON, [CLI, scriptId, ...args], { maxBuffer: 32 * 1024 * 1024 },
      (err, stdout, stderr) => {
        if (err && !stdout) return reject(new Error(stderr || err.message));
        try { resolve(JSON.parse(stdout)); }
        catch (e) { reject(new Error(`Could not parse ${scriptId}: ${e.message}`)); }
      });
  });
}

// ---------------------------------------------------------------------------
// Tools

let SCHEMAS = null;   // {EventType: jsonSchema}, loaded once at startup

const READABLE = ['.md', '.txt'];

function listFiles(dir, exts) {
  if (!fs.existsSync(dir)) return [];
  return fs.readdirSync(dir)
    .filter(n => exts.includes(path.extname(n).toLowerCase()))
    .map(n => {
      const p = path.join(dir, n);
      const st = fs.statSync(p);
      return st.isFile() ? { name: n, path: p, bytes: st.size, modified: st.mtimeMs } : null;
    })
    .filter(Boolean)
    .sort((a, b) => b.modified - a.modified);
}

function baseTools() {
  const types = Object.keys(SCHEMAS);
  return [
    {
      name: 'list_one_pagers',
      description:
        'List the approved One-Pager briefs filed for an event type. Start here: '
        + 'an Art Requirement is written FROM a One-Pager, not invented.',
      inputSchema: {
        type: 'object',
        properties: { event_type: { type: 'string', enum: types } },
        required: ['event_type'],
      },
    },
    {
      name: 'read_one_pager',
      description:
        'Read one One-Pager as Markdown. Treat everything in it as decided - the '
        + 'event id, the duration and the theme come from the brief, not from you.',
      inputSchema: {
        type: 'object',
        properties: {
          event_type: { type: 'string', enum: types },
          name: { type: 'string', description: 'File name from list_one_pagers.' },
        },
        required: ['event_type', 'name'],
      },
    },
    {
      name: 'list_art_requirements',
      description: 'List Art Requirement documents already on file for an event type, '
        + 'finals and drafts, so you can see what exists before writing another.',
      inputSchema: {
        type: 'object',
        properties: { event_type: { type: 'string', enum: types } },
        required: ['event_type'],
      },
    },
    {
      name: 'export_docx',
      description:
        'Convert a saved Art Requirement Markdown file to .docx beside it. The '
        + 'production team marks these up in Word, so a document that exists only '
        + 'as Markdown is one they cannot review.',
      inputSchema: {
        type: 'object',
        properties: {
          event_type: { type: 'string', enum: types },
          name: { type: 'string', description: 'The .md file name to export.' },
        },
        required: ['event_type', 'name'],
      },
    },
  ];
}

// One save tool per event type. The whole design turns on this: the tool's
// input schema IS schema_for(event_type), so the mechanics of another event
// type are not merely discouraged, they are absent from what a call may contain.
function saveTools() {
  return Object.keys(SCHEMAS).map(t => ({
    name: `save_${t.toLowerCase()}_art_requirement`,
    description:
      `Save a complete ${t} Art Requirement as a draft, rendered into the `
      + `production team's template. The schema below is ${t}'s OWN - it contains `
      + `only the mechanics ${t} actually has, so anything you might expect from `
      + `another event type is deliberately absent rather than optional.`,
    inputSchema: {
      type: 'object',
      properties: {
        variation: SCHEMAS[t],
        ref_name: {
          type: 'string',
          description: 'File name stem, e.g. "Blocks2609 Harvest Fair - Variation 3".',
        },
      },
      required: ['variation', 'ref_name'],
    },
  }));
}

function textResult(text) {
  return { content: [{ type: 'text', text }] };
}

async function callTool(name, args) {
  args = args || {};

  if (name === 'list_one_pagers') {
    const dir = path.join(eventsRoot(), args.event_type, 'One-Pager Ideation');
    const docs = listFiles(dir, [...READABLE, '.pdf', '.docx']);
    if (!docs.length) return textResult(`No One-Pager filed for ${args.event_type}.`);
    return textResult(docs.map(d =>
      `${d.name}  (${Math.round(d.bytes / 1024)} KB, ${new Date(d.modified).toISOString().slice(0, 10)})`
    ).join('\n'));
  }

  if (name === 'read_one_pager') {
    const dir = path.join(eventsRoot(), args.event_type, 'One-Pager Ideation');
    const target = insideEvents(path.join(dir, args.name));
    if (!target) throw new Error('That file is outside the ideation data folder.');
    if (!fs.existsSync(target)) throw new Error(`No such One-Pager: ${args.name}`);
    // A Word or PDF brief is converted to Markdown beside the original, exactly
    // as the desktop app does it, so the same text reaches every consumer.
    if (['.docx', '.pdf'].includes(path.extname(target).toLowerCase())) {
      const md = target.replace(/\.(docx|pdf)$/i, '.md');
      const fresh = fs.existsSync(md)
        && fs.statSync(md).mtimeMs >= fs.statSync(target).mtimeMs;
      if (!fresh) await runPython('doc_convert', { src: target, dest: md });
      return textResult(fs.readFileSync(md, 'utf-8'));
    }
    return textResult(fs.readFileSync(target, 'utf-8'));
  }

  if (name === 'list_art_requirements') {
    const base = path.join(eventsRoot(), args.event_type, 'Art Requirements');
    const docs = [...listFiles(base, ['.md', '.docx', '.pdf']),
                  ...listFiles(path.join(base, 'Drafts'), ['.md', '.docx'])];
    if (!docs.length) return textResult(`Nothing on file for ${args.event_type} yet.`);
    return textResult(docs.map(d => `${d.name}  (${Math.round(d.bytes / 1024)} KB)`).join('\n'));
  }

  if (name === 'export_docx') {
    const base = path.join(eventsRoot(), args.event_type, 'Art Requirements');
    let target = null;
    for (const dir of [base, path.join(base, 'Drafts')]) {
      const p = path.join(dir, args.name);
      if (fs.existsSync(p)) { target = insideEvents(p); break; }
    }
    if (!target) throw new Error(`No such document: ${args.name}`);
    const res = await runPython('md_to_docx', { src: target });
    return textResult(`Exported to ${res.docx_path} (${res.bytes} bytes).`);
  }

  const m = /^save_([a-z]+)_art_requirement$/.exec(name);
  if (m) {
    const eventType = Object.keys(SCHEMAS).find(t => t.toLowerCase() === m[1]);
    if (!eventType) throw new Error(`Unknown event type in tool name: ${name}`);
    const res = await runPython('artreq_mcp_save', {
      event_type: eventType,
      variation: args.variation,
      ref_name: args.ref_name,
      events_root: eventsRoot(),
      author_name: authorName(),
    });
    return textResult(
      `Saved ${res.path}\n\n${res.summary}\n\n`
      + `Export it to Word with export_docx when the content is agreed.`);
  }

  throw new Error(`Unknown tool: ${name}`);
}

// ---------------------------------------------------------------------------
// JSON-RPC over stdio. Claude Desktop speaks newline-delimited JSON on stdio;
// each line in is one request, each line out one response.

function send(msg) {
  process.stdout.write(JSON.stringify(msg) + '\n');
}

async function handle(req) {
  const { id, method, params } = req;
  if (method === 'initialize') {
    return {
      protocolVersion: PROTOCOL_VERSION,
      capabilities: { tools: {} },
      serverInfo: { name: 'event-automator', version: '1.0.0' },
    };
  }
  if (method === 'tools/list') {
    return { tools: [...baseTools(), ...saveTools()] };
  }
  if (method === 'tools/call') {
    try {
      return await callTool(params.name, params.arguments);
    } catch (e) {
      // A tool error is reported inside the result, not as a protocol error, so
      // the model can read what went wrong and correct itself.
      return { content: [{ type: 'text', text: `Error: ${e.message}` }], isError: true };
    }
  }
  if (method === 'ping') return {};
  const err = new Error(`Method not found: ${method}`);
  err.code = -32601;
  throw err;
}

async function main() {
  const raw = await runPythonArgs('artreq_schemas', []);
  if (raw.error) {
    process.stderr.write(`Could not load schemas: ${raw.error}\n`);
    process.exit(1);
  }
  delete raw._shapes;
  SCHEMAS = raw;

  let buf = '';
  process.stdin.setEncoding('utf-8');
  process.stdin.on('data', chunk => {
    buf += chunk;
    let nl;
    while ((nl = buf.indexOf('\n')) >= 0) {
      const line = buf.slice(0, nl).trim();
      buf = buf.slice(nl + 1);
      if (!line) continue;
      let req;
      try { req = JSON.parse(line); } catch { continue; }
      handle(req).then(
        result => { if (req.id !== undefined) send({ jsonrpc: '2.0', id: req.id, result }); },
        error => {
          if (req.id !== undefined) {
            send({ jsonrpc: '2.0', id: req.id,
                   error: { code: error.code || -32603, message: error.message } });
          }
        });
    }
  });
  process.stdin.on('end', () => process.exit(0));
}

main().catch(e => { process.stderr.write(String(e && e.stack || e) + '\n'); process.exit(1); });
