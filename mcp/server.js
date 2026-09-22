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

function dataRoot() {
  const cfg = loadConfig();
  const configured = (cfg.ideationDataPath || '').trim();
  return configured || path.join(ROOT, 'ideationData');
}

function eventsRoot() {
  return path.join(dataRoot(), 'Events');
}

function localizationRoot() {
  return path.join(dataRoot(), 'Localization');
}

// The localization memory is read from WorkSpcae3 and NOT from the workspace
// selected in the app. WorkSpcae3 is the master that pipeline step 5 clones
// into the others, so it is where the translator's corrections land first:
// against WorkSpcae_Design, 464 of the same corrections read as unmerged purely
// because that clone is behind. A stale clone would quietly teach this tool the
// wrong "before".
const MEMORY_WORKSPACE = 'WorkSpcae3';

function localizationDesignPath() {
  const cfg = loadConfig();
  const branch = (cfg.branchPaths || {})[cfg.activeBranch || 'trunk']
    || (cfg.branchPaths || {}).trunk || '';
  if (!branch) return '';
  return path.join(branch, 'design', 'DesignData', MEMORY_WORKSPACE, 'design');
}

function localizationCacheDir() {
  return path.join(process.env.APPDATA || ROOT, 'event-automation-desktop', 'loc-cache');
}

function authorName() {
  return (loadConfig().authorName || '').trim();
}

// A one-line record of what the host actually asked for, because otherwise
// questions like "did the panel render?" are unanswerable from outside the
// host. This build stopped writing per-server logs under %APPDATA%\Claude\logs
// (the Apify one goes stale while Apify is still installed), so there was no
// way to tell a server that never started from one the host ignored.
//
// The interesting line is a resources/read for the ui:// panel: the host only
// fetches that when it is actually going to render the MCP App.
const TRACE = path.join(process.env.APPDATA || ROOT, 'event-automation-desktop',
                        'mcp-trace.log');

function trace(line) {
  try {
    // Truncate rather than rotate: this is a diagnostic, and an unbounded log
    // in a folder nobody prunes is its own small bug.
    if (fs.existsSync(TRACE) && fs.statSync(TRACE).size > 256 * 1024) {
      fs.writeFileSync(TRACE, '');
    }
    fs.mkdirSync(path.dirname(TRACE), { recursive: true });
    fs.appendFileSync(TRACE, `${new Date().toISOString()} pid=${process.pid} ${line}\n`);
  } catch (e) {
    // Never let tracing break a tool call.
  }
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
    execFile(PYTHON, [CLI, scriptId, tmp], { maxBuffer: 128 * 1024 * 1024 },
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
    execFile(PYTHON, [CLI, scriptId, ...args], { maxBuffer: 128 * 1024 * 1024 },
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

// Same order and codes as localization.xlsx's B..H columns.
const LOC_LANGS = ['DE', 'ES', 'NL', 'SV', 'FR', 'DA'];

// A save_translations response echoes at most this many rows. Beyond it the
// response grew with the input until it exceeded the stdout buffer below, and
// the failure landed AFTER the files were written - work done, caller told it
// failed. Measured: 30,000 rows fine, 40,000 broke at exactly 33,554,432 bytes.
const MAX_ECHO_ROWS = 500;

// MCP Apps (io.modelcontextprotocol/ui): the localization panel rendered inside
// the conversation, so translating needs no Anthropic key — the host is the
// model. The mime type is fixed by the extension spec and must be exactly this.
const UI_MIME = 'text/html;profile=mcp-app';
const UI_APPS_EXTENSION = 'io.modelcontextprotocol/ui';
const LOC_APP_URI = 'ui://event-automator/localization-panel';
const LOC_APP_FILE = path.join(__dirname, 'localization-app.html');

// No connectDomains/resourceDomains on purpose: the panel is self-contained,
// so it needs no CSP allowances and there is no third-party origin to trust.
const LOC_APP_META = { ui: { prefersBorder: true } };

function uiResources() {
  return [{
    uri: LOC_APP_URI,
    name: 'localization_panel',
    description: 'Interactive panel: English in, six languages out, drawing on the '
      + 'translator-approved memory.',
    mimeType: UI_MIME,
    _meta: LOC_APP_META,
  }];
}

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
    {
      name: 'localization_brief',
      description:
        'START HERE for any translation work. Given English UI strings, returns '
        + 'what is ALREADY SETTLED (the translator\'s own approved text, or what the '
        + 'game already ships - copy these through unchanged, never retranslate '
        + 'them), which strings still need translating, the nearest strings we '
        + 'already ship in each language, and per language the corrections the '
        + 'professional translator made to our earlier attempts. Translate from '
        + 'this, not from scratch - the whole point is that a string the translator '
        + 'has already ruled on never gets a second opinion.',
      inputSchema: {
        type: 'object',
        properties: {
          strings: {
            type: 'array',
            items: { type: 'string' },
            minItems: 1,
            description: 'The English strings, one per entry.',
          },
          languages: {
            type: 'array',
            items: { type: 'string', enum: LOC_LANGS },
            description: 'Subset to work on. Omit for all six.',
          },
        },
        required: ['strings'],
      },
    },
    {
      name: 'save_translations',
      description:
        'File a completed translation table into the shared folder as .tsv and '
        + '.md, with the columns in localization.xlsx\'s own B..H order so a paste '
        + 'lands in the right place. Every language is REQUIRED per row - a row '
        + 'cannot be filed with Danish quietly left out. Typography is applied here '
        + 'rather than trusted to you: write plain straight quotes and this converts '
        + 'them to each language\'s own marks. Placeholders, length and casing are '
        + 'then checked and anything doubtful is reported back to you.',
      inputSchema: {
        type: 'object',
        properties: {
          rows: {
            type: 'array',
            minItems: 1,
            items: {
              type: 'object',
              properties: {
                en: { type: 'string', description: 'The English string, echoed exactly.' },
                DE: { type: 'string' }, ES: { type: 'string' }, NL: { type: 'string' },
                SV: { type: 'string' }, FR: { type: 'string' }, DA: { type: 'string' },
              },
              required: ['en', ...LOC_LANGS],
              additionalProperties: false,
            },
          },
          ref_name: {
            type: 'string',
            description: 'File name stem, e.g. "Blocks2610 event copy".',
          },
          job_id: {
            type: 'string',
            description: 'Pass this back VERBATIM when the batch came from '
              + 'localization_pending. It is what releases the designer\'s app '
              + 'from waiting and drops the table into it. Omit only for work '
              + 'that did not come from a queued job.',
          },
        },
        required: ['rows', 'ref_name'],
      },
      // Same panel as the entry tool: when this is called the host re-renders
      // the finished table back INTO the open panel instead of only printing it
      // in the chat, which is what closes the loop the user actually sees.
      _meta: { ui: { resourceUri: LOC_APP_URI, visibility: ['model', 'app'] } },
    },
    {
      name: 'localization_pending',
      description:
        'Pick up a translation batch that a designer queued from the Event '
        + 'Automator desktop app and is now waiting on. Call this when the user '
        + 'says anything like "do the pending localization batch", "translate the '
        + 'queued strings", or just "the batch". It returns the brief for the '
        + 'oldest unanswered job. Translate the gaps, then call save_translations '
        + 'passing back the SAME job_id — that is what puts the finished table '
        + 'into the app the designer is sitting in front of. Without the job_id '
        + 'they keep waiting.',
      inputSchema: {
        type: 'object',
        properties: {
          job_id: {
            type: 'string',
            description: 'A specific queued job. Omit to take the oldest waiting one.',
          },
        },
      },
    },
    {
      name: 'localization_panel',
      description:
        'Open the interactive Localization panel in the conversation. Use this '
        + 'whenever the user wants to translate UI strings and would rather work '
        + 'in a panel than in chat - they can paste English in, see what the '
        + 'translator-approved memory already answers, and send only the genuinely '
        + 'new strings to you for translation. Pass `strings` to prefill it.',
      inputSchema: {
        type: 'object',
        properties: {
          strings: {
            type: 'array',
            items: { type: 'string' },
            description: 'Optional English strings to prefill the panel with.',
          },
          languages: {
            type: 'array',
            items: { type: 'string', enum: LOC_LANGS },
            description: 'Subset to work on. Omit for all six.',
          },
        },
      },
      _meta: { ui: { resourceUri: LOC_APP_URI, visibility: ['model', 'app'] } },
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
        source_one_pager: {
          type: 'string',
          description:
            'The One-Pager file name this document was written from, exactly as '
            + 'list_one_pagers reported it. Recorded in the document so a reviewer '
            + 'can see which brief it came from and the tool can group documents '
            + 'by event. Pass it whenever you read a One-Pager first.',
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

  if (name === 'localization_brief') {
    const design = localizationDesignPath();
    if (!design) throw new Error('No branch path configured in the desktop app yet.');
    const res = await runPython('loc_mcp_brief', {
      strings: args.strings,
      languages: args.languages,
      design_path: design,
      corrections_root: localizationRoot(),
      cache_dir: localizationCacheDir(),
    });
    const mem = res.memory || {};
    // structuredContent as well as text: the MCP App renders the table from it,
    // while the text is what a chat-only conversation reads.
    return {
      content: [{ type: 'text', text:
        `${res.counts.strings} string(s): ${res.counts.settled} fully settled by memory, `
        + `${res.counts.gaps} needing work. Memory: ${mem.approved_entries} approved, `
        + `${mem.shipped_entries} shipped (${(mem.workbooks || []).join(', ')}).\n\n`
        + res.text }],
      structuredContent: res,
    };
  }

  if (name === 'localization_panel') {
    // The panel's own opening state. It renders from this immediately rather
    // than starting blank and making the designer click to see the memory.
    const design = localizationDesignPath();
    if (!design) throw new Error('No branch path configured in the desktop app yet.');
    const strings = Array.isArray(args.strings) ? args.strings.filter(Boolean) : [];
    if (!strings.length) {
      return {
        content: [{ type: 'text', text:
          'Localization panel opened. Paste the English strings into it — '
          + '"Check memory" answers with no model at all, and "Translate" sends '
          + 'only the genuinely new strings back to me.' }],
        structuredContent: { rows: [], languages: args.languages || LOC_LANGS, counts: {} },
      };
    }
    const res = await runPython('loc_mcp_brief', {
      strings,
      languages: args.languages,
      design_path: design,
      corrections_root: localizationRoot(),
      cache_dir: localizationCacheDir(),
    });
    return {
      content: [{ type: 'text', text:
        `Localization panel opened with ${res.counts.strings} string(s): `
        + `${res.counts.settled} fully answered by memory, ${res.counts.gaps} needing `
        + `translation. The panel will send me the gaps when the designer clicks `
        + `Translate.` }],
      structuredContent: res,
    };
  }

  if (name === 'localization_pending') {
    const res = await runPython('loc_pending', {
      job_id: args.job_id,
      corrections_root: localizationRoot(),
      cache_dir: localizationCacheDir(),
    });
    if (!res.job) return textResult(res.reason || 'No batches are waiting.');
    const job = res.job;
    return textResult(
      `Batch ${job.job_id}`
      + (job.requested_by ? ` queued by ${job.requested_by}` : '')
      + `: ${job.counts.strings} string(s), ${job.counts.settled} already answered `
      + `by memory, ${job.counts.gaps} needing translation.\n\n`
      + `When the table is complete, call save_translations with `
      + `job_id "${job.job_id}" so it lands back in their app.\n\n`
      + job.text);
  }

  if (name === 'save_translations') {
    const res = await runPython('loc_mcp_save', {
      rows: args.rows,
      ref_name: args.ref_name,
      job_id: args.job_id || '',
      author_name: authorName(),
      corrections_root: localizationRoot(),
      // So a row passed back unchanged is labelled as reused rather than as
      // something the model wrote.
      design_path: localizationDesignPath(),
      cache_dir: localizationCacheDir(),
    });
    const flagged = (res.flagged || []).map(
      f => `- ${f.en} / ${f.lang}: ${f.warning}`).join('\n');

    // Echo back at most MAX_ECHO_ROWS. The response used to carry every row,
    // so it grew with the input and eventually blew execFile's stdout buffer -
    // and it failed AFTER the .tsv/.md were written, so a 40,000-row call left
    // 57MB in the shared folder while reporting failure. The files and the
    // counts are the deliverable; echoing the whole table back is not.
    const total = (res.rows || []).length;
    const structured = total > MAX_ECHO_ROWS
      ? Object.assign({}, res, {
          rows: res.rows.slice(0, MAX_ECHO_ROWS),
          rows_echoed: MAX_ECHO_ROWS,
          rows_total: total,
          rows_truncated: true,
        })
      : res;

    return {
      content: [{ type: 'text', text:
        `${res.summary}\n\nTSV: ${res.tsv_path}\nMarkdown: ${res.md_path}\n`
        + (total > MAX_ECHO_ROWS
            ? `\nAll ${total} rows were filed; only the first ${MAX_ECHO_ROWS} are `
              + `echoed back here. The files above hold the complete table.\n`
            : '')
        + (flagged ? `\nChecks to look at:\n${flagged}\n\nFix any of these and call `
                     + `save_translations again; each save is filed separately.`
                   : '\nNothing flagged.') }],
      structuredContent: structured,
    };
  }

  const m = /^save_([a-z]+)_art_requirement$/.exec(name);
  if (m) {
    const eventType = Object.keys(SCHEMAS).find(t => t.toLowerCase() === m[1]);
    if (!eventType) throw new Error(`Unknown event type in tool name: ${name}`);
    const res = await runPython('artreq_mcp_save', {
      event_type: eventType,
      variation: args.variation,
      ref_name: args.ref_name,
      source_one_pager: args.source_one_pager || '',
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
  if (method === 'resources/read') {
    trace(`resources/read ${(params || {}).uri}  <- the host is rendering the app`);
  } else if (method === 'tools/call') {
    trace(`tools/call ${(params || {}).name}`);
  } else if (method === 'initialize') {
    trace(`initialize clientProtocol=${(params || {}).protocolVersion} `
      + `client=${JSON.stringify((params || {}).clientInfo || {})} `
      + `clientCapabilities=${JSON.stringify((params || {}).capabilities || {})}`);
  } else {
    trace(method);
  }
  if (method === 'initialize') {
    // Echo the version the client asked for rather than pinning our own.
    //
    // This was a real bug, and it is why the localization panel never rendered:
    // we answered '2024-11-05' no matter what Claude Desktop requested, so the
    // whole session negotiated down to a protocol that predates extensions and
    // MCP Apps. The host then had nowhere to put a ui:// resource. Everything
    // this server does is version-stable, so honouring the client's version is
    // both correct per spec and what makes the extension reachable.
    const asked = (params || {}).protocolVersion;
    return {
      protocolVersion: (typeof asked === 'string' && asked) ? asked : PROTOCOL_VERSION,
      capabilities: {
        tools: {},
        // resources is what serves the ui:// panel; without it a host has no
        // way to fetch the HTML even though the tool points at it.
        resources: {},
        extensions: { [UI_APPS_EXTENSION]: { mimeTypes: [UI_MIME] } },
      },
      serverInfo: { name: 'event-automator', version: '1.0.0' },
    };
  }
  if (method === 'tools/list') {
    return { tools: [...baseTools(), ...saveTools()] };
  }
  if (method === 'resources/list') {
    return { resources: uiResources() };
  }
  if (method === 'resources/read') {
    if ((params || {}).uri !== LOC_APP_URI) {
      throw new Error(`No such resource: ${(params || {}).uri}`);
    }
    return {
      contents: [{
        uri: LOC_APP_URI,
        mimeType: UI_MIME,
        text: fs.readFileSync(LOC_APP_FILE, 'utf-8'),
        _meta: LOC_APP_META,
      }],
    };
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
