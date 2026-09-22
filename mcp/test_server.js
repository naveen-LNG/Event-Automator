// Drive the MCP server over stdio the way Claude Desktop will.
//
// The load-bearing check is the last one: a Blocks save carrying Merge's
// mechanics must be impossible. Two independent things have to hold for that -
// the Blocks tool's INPUT SCHEMA must not contain max_level (so a host that
// constrains tool calls cannot produce one), and the Python save must REJECT one
// anyway (so a host that does not constrain calls still cannot write the file).
// Testing only the schema would pass while the second net was missing.
//
// Run: node mcp/test_server.js

'use strict';
const { spawn } = require('child_process');
const path = require('path');
const fs = require('fs');

let fails = 0;
function ok(cond, msg, detail) {
  if (cond) { console.log('OK   ' + msg); }
  else { console.log('FAIL ' + msg + (detail ? '\n     ' + detail : '')); fails++; }
}

function rpc(child, pending, method, params) {
  const id = rpc.n = (rpc.n || 0) + 1;
  return new Promise((resolve, reject) => {
    pending.set(id, { resolve, reject });
    child.stdin.write(JSON.stringify({ jsonrpc: '2.0', id, method, params }) + '\n');
  });
}

async function main() {
  const child = spawn(process.execPath, [path.join(__dirname, 'server.js')],
    { stdio: ['pipe', 'pipe', 'pipe'] });
  const pending = new Map();
  let buf = '';
  child.stdout.setEncoding('utf-8');
  child.stdout.on('data', c => {
    buf += c;
    let nl;
    while ((nl = buf.indexOf('\n')) >= 0) {
      const line = buf.slice(0, nl).trim();
      buf = buf.slice(nl + 1);
      if (!line) continue;
      const msg = JSON.parse(line);
      const p = pending.get(msg.id);
      if (p) { pending.delete(msg.id); msg.error ? p.reject(new Error(msg.error.message)) : p.resolve(msg.result); }
    }
  });
  child.stderr.on('data', d => process.stderr.write('[server] ' + d));

  const init = await rpc(child, pending, 'initialize', {});
  ok(init.serverInfo && init.serverInfo.name === 'event-automator',
     'the server initialises and identifies itself');
  ok(!!init.capabilities.tools, 'and advertises tools');

  const { tools } = await rpc(child, pending, 'tools/list', {});
  const names = tools.map(t => t.name);
  ok(names.includes('list_one_pagers') && names.includes('read_one_pager')
     && names.includes('export_docx'),
     'the workflow tools are present', names.join(', '));

  const saves = names.filter(n => n.startsWith('save_'));
  ok(saves.length >= 5, 'one save tool per event type', saves.join(', '));

  // ---- the guarantee, at the schema layer --------------------------------
  const blocks = tools.find(t => t.name === 'save_blocks_art_requirement');
  const merge = tools.find(t => t.name === 'save_merge_art_requirement');
  ok(!!blocks && !!merge, 'Blocks and Merge each get their own save tool');

  const bs = JSON.stringify(blocks.inputSchema);
  const ms = JSON.stringify(merge.inputSchema);
  ok(!bs.includes('max_level'),
     "the Blocks tool's schema has no max_level at all - a host cannot emit one");
  ok(ms.includes('max_level'),
     'while Merge, which really has upgrade stages, keeps it');
  const bz = blocks.inputSchema.properties.variation.properties.zones.items.properties;
  ok(bz.buildings.maxItems === 0,
     'Blocks buildings are capped at zero, not merely discouraged');
  ok(blocks.inputSchema.properties.variation.properties.producers.maxItems === 0,
     "and Merge's producers cannot appear in a Blocks document either");
  ok(bs.length < ms.length,
     'the Blocks schema is smaller than Merge\'s, not just differently worded',
     `${bs.length} vs ${ms.length} chars`);

  // ---- the guarantee, at the save layer ----------------------------------
  // Hand the Blocks tool a document a non-enforcing host might let through.
  const smuggled = {
    title: '[Artistic Requirement] Blocks2609 Smuggled',
    gameplay_overview: { theme: 't', duration_summary: '16 days', narrative_packaging: 'n',
                         loop_summary: 'l', resource_planning: 'r', major_changes: 'None.' },
    zones: [{
      zone_number: 1, zone_theme: 'Z', palette_or_mood: 'warm', main_tone: 'WARM',
      scene_material: 'timber', before_state: 'bare', construction_steps: ['a', 'b'],
      buildings: [{ building_number: 1, name: 'Smuggled Hall', area: 'A',
                    silhouette: 's', key_features: ['f'], max_level: 3, signage: [] }],
      reference_notes: 'n',
    }],
    producers: [], recipes: [], ingredients: [], construction_materials: [],
    characters: [], event_motif: '', art_direction_guidance: { lean_into: [], avoid: [] },
    event_extras: {}, gameplay_specific_content: '', ui_icon_refresh_notes: '', notes: '',
  };
  const res = await rpc(child, pending, 'tools/call', {
    name: 'save_blocks_art_requirement',
    arguments: { variation: smuggled, ref_name: 'ZZ_MCP_SELFTEST_should_not_exist' },
  });
  const text = res.content[0].text;
  ok(res.isError === true, 'a Blocks save carrying Merge mechanics is refused');
  ok(/max_level|buildings/.test(text),
     'and the refusal names what was wrong so the model can fix it',
     text.slice(0, 200));

  // ---- the read path ------------------------------------------------------
  const list = await rpc(child, pending, 'tools/call',
    { name: 'list_one_pagers', arguments: { event_type: 'Blocks' } });
  ok(/One-Pager/.test(list.content[0].text),
     'list_one_pagers finds the shared folder', list.content[0].text.slice(0, 120));

  const bad = await rpc(child, pending, 'tools/call',
    { name: 'read_one_pager', arguments: { event_type: 'Blocks', name: '../../../../etc/passwd' } });
  ok(bad.isError === true && /outside/.test(bad.content[0].text),
     'a traversing path is refused rather than read', bad.content[0].text.slice(0, 120));

  // ---- a big save must not blow the response ------------------------------
  // save_translations used to echo every row back, so the response grew with
  // the input until it exceeded execFile's stdout buffer at exactly 32MiB.
  // Worse, that failure landed AFTER the .tsv/.md were written: a 40,000-row
  // call left 57MB in the shared folder while telling the caller it failed.
  // The echo is capped now; this pins that so it cannot creep back.
  const LANGS = ['DE', 'ES', 'NL', 'SV', 'FR', 'DA'];
  const many = Array.from({ length: 600 }, (_, i) => {
    const row = { en: `MCP selftest row ${i}` };
    for (const l of LANGS) row[l] = `MCP selftest ${l} ${i}`;
    return row;
  });
  const big = await rpc(child, pending, 'tools/call', {
    name: 'save_translations',
    arguments: { rows: many, ref_name: 'ZZ_MCP_SELFTEST_echo_cap' },
  });
  const bigSc = big.structuredContent || {};
  ok(!big.isError, 'a 600-row save succeeds', (big.content[0] || {}).text);
  ok(bigSc.rows_total === 600, 'it reports all 600 rows filed', bigSc.rows_total);
  ok((bigSc.rows || []).length <= 500,
     `the response echoes at most 500 rows (got ${(bigSc.rows || []).length})`);
  ok(bigSc.rows_truncated === true, 'and says so rather than looking complete');
  // Leave nothing in the shared folder.
  for (const key of ['tsv_path', 'md_path']) {
    if (bigSc[key]) {
      try { fs.unlinkSync(bigSc[key]); }
      catch (e) { console.error('WARN could not remove ' + bigSc[key] + ': ' + e.message); }
    }
  }

  child.kill();
  console.log('');
  if (fails) { console.error(`${fails} MCP CHECK(S) FAILED`); process.exit(1); }
  console.log('MCP SERVER OK');
}

main().catch(e => { console.error(e); process.exit(1); });
