// A stand-in for the Ollama server, so the app's real UI flow can be driven
// without a reachable model.
//
// The app is dual-provider and its Ollama path is a single unauthenticated
// POST /api/chat carrying the JSON Schema it wants back in `format`. So this
// walks that schema and returns a conforming object - which means the renderer
// takes exactly the code path it takes with a real model, and nothing in the
// app knows the difference.
//
// The Merge art requirement gets a hand-built body rather than a generic one,
// because the pass picker's counts come from buildings and producers and a
// schema-minimal answer would leave them all empty.

const http = require('http');
const PORT = Number(process.env.STUB_PORT || 11434);

const LEVELS = [2, 2, 1, 2, 3, 2, 1, 2];   // 15 stages per scene, as Merge2608 ships
const MATERIALS = ['woven willow, wicker, rope and canvas',
                   'blown glass, sea glass and polished ceramic',
                   'glazed mosaic tile, terrazzo and polished marble'];

function building(n, lv) {
  return {
    building_number: n, name: `Building ${n}`, area: 'ABC'[(n - 1) % 3],
    silhouette: 'a woven form with a bound rim',
    key_features: ['basket weave', 'bound rim', 'rope join'],
    max_level: lv,
    signage: n === 2 ? ['WILLOW WORKS'] : (n === 4 ? ['BASKET'] : [])
  };
}

// A draft as it looked BEFORE the schema carried buildings and producers, so the
// pass picker's disabled branch can be seen in the real app: only `scene` has
// anything to render, and every other pass should come up dimmed and unclickable.
function legacyVariation(i) {
  return {
    title: `[Artistic Requirement] Merge2610 Legacy Draft (variation ${i})`,
    gameplay_overview: {
      theme: 'Legacy Draft', duration_summary: '14 days',
      narrative_packaging: 'A draft written before the schema change.',
      loop_summary: 'Merge + Build', resource_planning: '3 zones',
      major_changes: 'None - standard template.'
    },
    zones: [1, 2, 3].map(z => ({
      zone_number: z, zone_theme: `Legacy Zone ${z}`,
      palette_or_mood: 'honey willow',
      construction_steps: ['clear the ground', 'raise the frame', 'dress the front'],
      reference_notes: 'no buildings block on this draft'
    })),
    gameplay_specific_content: 'Free-text only, as older drafts were.',
    characters: [], art_direction_guidance: { lean_into: [], avoid: [] },
    ui_icon_refresh_notes: '', notes: 'Pre-schema fixture.'
  };
}

function mergeVariation(i) {
  return {
    title: `[Artistic Requirement] Merge2610 Artisan Harbour (variation ${i})`,
    gameplay_overview: {
      theme: 'Artisan Harbour', duration_summary: '14 days',
      narrative_packaging: 'A craft harbour of weavers, glassblowers and tilers.',
      loop_summary: 'Merge + Build', resource_planning: '3 scenes, 8 buildings each, 45 stages',
      major_changes: 'None - standard template.'
    },
    zones: [1, 2, 3].map(z => ({
      zone_number: z, zone_theme: ['Woven Willow Village', 'Blown Glass Quarter', 'Mosaic Terrace'][z - 1],
      palette_or_mood: ['honey willow and rope cream', 'sea green and frosted white',
                        'terracotta, saffron and lapis'][z - 1],
      main_tone: '', scene_material: MATERIALS[z - 1], construction_steps: [],
      buildings: LEVELS.map((lv, k) => building(k + 1, lv)),
      reference_notes: 'seaside craft harbour'
    })),
    recipes: [], ingredients: [], event_motif: '',   // Merge has none of these
    producers: [1, 2, 3, 4].map(g => ({
      producer_number: g, name: `Producer ${g}`, palette: 'honey willow',
      closed_description: 'a stout lidded creel, closed and strapped',
      open_description: 'the same creel open, its goods inside',
      chains: [1, 2].map(c => ({
        chain_number: c, family: c === 1 ? 'willow basketry' : 'rope and sail',
        palette: 'honey willow',
        items: ['a willow rod', 'a tied bundle', 'a woven base', 'a half basket', 'a lidded hamper']
      }))
    })),
    gameplay_specific_content: 'Producer and chain tables above.',
    characters: [{
      name: 'Harbour Weaver', role: 'Gift-giving/mechanic NPC', npc_id: '101',
      base_character: 'squirrel', costume: 'a canvas apron and a rope belt',
      design_notes: 'Re-costumed existing cast.', scene_actions: ['waves', 'weaves']
    }],
    event_extras: {
      hot_air_balloon: 'a balloon of woven willow panels banded with cream canvas',
      achievement_badge: 'a crest of a woven basket rim and crossed willow rods',
      task_icons: ['a brass harbour compass', 'a coiled rope knot', 'a sea-glass float']
    },
    art_direction_guidance: { lean_into: ['handmade craft', 'seaside light'], avoid: ['dark realism'] },
    ui_icon_refresh_notes: 'Re-skin the entry icon to the willow palette.',
    notes: 'Stub fixture for driving the UI.'
  };
}

// A Dessert draft, which is the OTHER shape the pass picker has to handle: no
// producers and no merge chains, but 39 recipes and 11 ingredients. The counts
// are the shipped Dessert2608 ones - 4 zones, 11/11/12/12 buildings, 24 common
// plus 15 secret recipes, 6 common plus 5 secret ingredients - so the picker's
// numbers can be checked against something real rather than against itself.
const ZONE_THEMES = ['Camping Picnic Meadow', 'Sweet Bakery Open Market',
                     'SkipBo Park Carnival', 'Anniversary DJ Club'];
const ZONE_PALETTES = ['forest green and autumn warm', 'blush pink and cream',
                       'amber and terracotta', 'deep violet and gold'];
const ZONE_BUILDINGS = [11, 11, 12, 12];       // 46 in all
const DISHES = ['Tart', 'Bun', 'Macaron', 'Pretzel', 'Parfait', 'Jelly'];

function dessertVariation(i) {
  const recipes = [];
  for (let z = 1; z <= 4; z++) {
    for (let k = 0; k < 6; k++) {
      recipes.push({
        recipe_number: recipes.length + 1, name: `${ZONE_THEMES[z - 1].split(' ')[0]} ${DISHES[k]}`,
        zone_number: z, description: 'a plated dish carrying the anniversary 5',
        palette: ZONE_PALETTES[z - 1], ingredients: ['Party Flour', 'Sweet Butter'],
        secret: false
      });
    }
  }
  for (let s = 1; s <= 15; s++) {
    recipes.push({
      recipe_number: recipes.length + 1, name: `Secret Dish ${s}`,
      zone_number: (s - 1) % 4 + 1, description: 'a research-unlocked dish',
      palette: ZONE_PALETTES[(s - 1) % 4], ingredients: ['Carnival Sugar'], secret: true
    });
  }
  const ingNames = ['Party Flour', 'Sweet Butter', 'Fresh Berry', 'Honey Drop',
                    'Carnival Sugar', 'Garden Herb',                       // 6 common
                    'Midnight Cocoa', 'Starlight Sugar', 'Velvet Cream',
                    'Candied Petal', 'Spiced Ember'];                      // 5 secret
  return {
    title: `[Artistic Requirement] Dessert2611 Picnic Park (variation ${i})`,
    gameplay_overview: {
      theme: 'Picnic Park', duration_summary: '14 days',
      narrative_packaging: 'Four colour-coded zones of a cooking park.',
      loop_summary: 'Cook + Build',
      resource_planning: '4 zones, 46 buildings, 39 recipes, 11 ingredients',
      major_changes: 'None - standard template.'
    },
    zones: [1, 2, 3, 4].map(z => ({
      zone_number: z, zone_theme: ZONE_THEMES[z - 1], palette_or_mood: ZONE_PALETTES[z - 1],
      main_tone: z === 4 ? 'COOL' : 'WARM', scene_material: '', construction_steps: [],
      buildings: Array.from({ length: ZONE_BUILDINGS[z - 1] }, (_, k) => ({
        building_number: k + 1, name: `Zone ${z} Building ${k + 1}`, area: '',
        silhouette: 'an isometric flat-shaded cut-out with a culinary hook',
        key_features: ['pastry display case', 'awning', 'chimney'],
        max_level: 1,
        signage: []          // Dessert buildings carry no text at all
      })),
      reference_notes: 'isometric plate, high chroma, no text'
    })),
    producers: [],           // Dessert has none - the icon and producer passes must stay dark
    recipes: recipes,
    ingredients: ingNames.map((n, k) => ({
      name: n, description: 'a raw component', palette: 'warm cream', secret: k >= 6
    })),
    event_motif: 'the 5th-anniversary numeral 5',
    gameplay_specific_content: 'Recipe and ingredient tables above.',
    characters: [{
      name: 'Picnic Host', role: 'Gift-giving/mechanic NPC', npc_id: '101',
      base_character: 'squirrel', costume: 'a gingham apron and a sun hat',
      design_notes: 'Re-costumed existing cast.', scene_actions: ['waves', 'serves']
    }],
    event_extras: { hot_air_balloon: '', achievement_badge: '', task_icons: [] },
    art_direction_guidance: { lean_into: ['high chroma', 'clean isometric'], avoid: ['text on buildings'] },
    ui_icon_refresh_notes: 'Re-skin the entry icon to the picnic palette.',
    notes: 'Stub fixture for driving the Dessert pass picker.'
  };
}

// Generic schema walker, for anything not hand-built above (the ideation themes).
function fake(schema, key = '') {
  if (!schema || typeof schema !== 'object') return null;
  if (schema.enum) return schema.enum[0];
  switch (schema.type) {
    case 'object': {
      const out = {};
      for (const [k, v] of Object.entries(schema.properties || {})) out[k] = fake(v, k);
      return out;
    }
    case 'array': {
      const n = schema.minItems || 1;
      return Array.from({ length: n }, () => fake(schema.items, key));
    }
    case 'integer': return schema.minimum ?? 1;
    case 'number': return schema.minimum ?? 1;
    case 'boolean': return true;
    default:
      if (key === 'internal_reference_name') return 'Merge2610';
      if (key === 'theme_name') return 'Artisan Harbour';
      return `stub ${key || 'value'}`;
  }
}

http.createServer((req, res) => {
  if (req.method !== 'POST') { res.writeHead(404).end(); return; }
  let body = '';
  req.on('data', c => body += c);
  req.on('end', () => {
    let payload = {};
    try { payload = JSON.parse(body); } catch (e) { /* fall through to a generic answer */ }
    const schema = payload.format || {};
    let content;
    if (schema.properties && schema.properties.variations) {
      const n = schema.properties.variations.minItems || 1;
      const make = process.env.STUB_LEGACY ? legacyVariation
        : process.env.STUB_DESSERT ? dessertVariation : mergeVariation;
      content = { variations: Array.from({ length: n }, (_, i) => make(i + 1)) };
    } else {
      content = fake(schema);
    }
    console.log(`[stub] ${req.url} -> ${schema.properties ? Object.keys(schema.properties).join(',') : 'no schema'}`);
    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({ message: { role: 'assistant', content: JSON.stringify(content) }, done: true }));
  });
}).listen(PORT, '127.0.0.1', () => console.log(`stub ollama on http://127.0.0.1:${PORT}`));
