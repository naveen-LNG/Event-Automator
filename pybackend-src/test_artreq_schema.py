"""Tests for the extended art requirement schema.

The schema now carries the structure the Merge documents actually specify -
buildings with upgrade levels, and producers with their two 5-item chains -
instead of hiding it in a free-text `construction_steps` list.

The counts asserted here are not invented. They come from the Merge2603 template
and from the shipped Merge2608 game data
(WorkSpcae3/convert/RawJson/merge2608_building.json): 25 buildings, 45 upgrade
stages, exactly 15 per scene, and 4 producers x 2 chains x 5 items.

The biggest regression risk is stored drafts written before these fields
existed, so that case is tested first.
"""
import os, sys, json, tempfile, shutil

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
HERE = os.path.dirname(os.path.abspath(__file__))
os.chdir(HERE); sys.path.insert(0, HERE)

import art_requirement_generator as arg

fails = []
def ok(c, m, d=""):
    print(("OK  " if c else "FAIL ") + m + (("\n     " + d) if (d and not c) else ""))
    if not c: fails.append(m)

# ---- backward compatibility: a doc written before this change ---------------
OLD_DOC = {
    "title": "[Artistic Requirement] Sort2609 Pet Sanctuary",
    "gameplay_overview": {"theme": "Pets", "duration_summary": "14 days",
                          "narrative_packaging": "n", "loop_summary": "l",
                          "resource_planning": "r", "major_changes": "None"},
    "zones": [{"zone_number": 1, "zone_theme": "Cat Room", "palette_or_mood": "pink",
               "construction_steps": ["a", "b"], "reference_notes": "n"}],
    "characters": [], "art_direction_guidance": {"lean_into": [], "avoid": []},
    "gameplay_specific_content": "x", "ui_icon_refresh_notes": "y", "notes": "z",
}
tmp = tempfile.mkdtemp()
try:
    md = arg._render_markdown(OLD_DOC, "Sort", "Draft", "Tester", md_dir=tmp)
    ok("Cat Room" in md, "a draft written before this change still renders")
    ok("Construction/Reveal Steps" in md,
       "and its construction_steps still appear")
    ok("Buildings (" not in md and "Producers (" not in md,
       "with no empty buildings or producers section bolted on",
       "absent data must render as nothing, not as an empty table")
finally:
    shutil.rmtree(tmp, ignore_errors=True)

# ---- the schema itself -------------------------------------------------------
S, Z, B = arg.ART_REQUIREMENT_SCHEMA, arg.ZONE_SCHEMA, arg.BUILDING_SCHEMA
C, P = arg.CHAIN_SCHEMA, arg.PRODUCER_SCHEMA

ok("producers" in S["properties"] and "producers" in S["required"],
   "producers is a required top-level section",
   "required, not optional, so a constrained decoder reliably emits it")
ok("buildings" in Z["properties"] and "buildings" in Z["required"],
   "buildings is a required part of a zone")
for f in ("main_tone", "scene_material"):
    ok(f in Z["properties"] and f in Z["required"], "a zone declares %s" % f)
ok(Z["properties"]["main_tone"]["enum"] == ["COOL", "WARM", ""],
   "main_tone is a closed set including empty, for types that do not commit")

ok(B["properties"]["max_level"]["minimum"] == 1
   and B["properties"]["max_level"]["maximum"] == 3,
   "max_level is bounded 1..3, as the shipped Merge2608 data is")
ok(B["properties"]["signage"]["type"] == "array",
   "signage is an array, so an EMPTY one can mean 'carries no text'",
   "a string could not express the difference between 'none' and 'unspecified'")
ok("EMPTY" in B["properties"]["signage"]["description"],
   "and the schema says so, because that distinction is load-bearing")

ok(C["properties"]["items"]["minItems"] == 5 and C["properties"]["items"]["maxItems"] == 5,
   "a chain is exactly 5 items - 每个生产链有5级产品")
ok(P["properties"]["chains"]["minItems"] == 2 and P["properties"]["chains"]["maxItems"] == 2,
   "a producer has exactly 2 chains - 每组生产器有2条生产链")
ok("family" in C["required"],
   "a chain must name its family, which is what keeps it category-pure")

ok(json.dumps(S) and S["additionalProperties"] is False,
   "the schema still serialises and stays closed")

# ---- a full Merge document renders ------------------------------------------
def bld(n, area, lv, signage):
    return {"building_number": n, "name": "Building %d" % n, "area": area,
            "silhouette": "a shape", "key_features": ["f1", "f2", "f3"],
            "max_level": lv, "signage": signage}

# 2 + 2 + 1 + 2 + 3 + 2 + 1 + 2 = 15 stages, matching the shipped per-scene total
LEVELS = [2, 2, 1, 2, 3, 2, 1, 2]
zones = []
for zi in range(1, 4):
    zones.append({
        "zone_number": zi, "zone_theme": "Scene %d" % zi, "palette_or_mood": "p",
        "main_tone": "", "scene_material": "woven willow", "construction_steps": [],
        "buildings": [bld(i, "ABC"[(i - 1) % 3], LEVELS[i - 1],
                          ["MARKET"] if i == 2 else [])
                      for i in range(1, 9)],
        "reference_notes": "n"})
producers = [{
    "producer_number": g, "name": "Producer %d" % g, "palette": "warm",
    "closed_description": "closed", "open_description": "open",
    "chains": [{"chain_number": c, "family": "fam%d" % c, "palette": "warm",
                "items": ["i%d" % k for k in range(1, 6)]} for c in (1, 2)]
} for g in range(1, 5)]
NEW_DOC = dict(OLD_DOC, zones=zones, producers=producers,
               title="[Artistic Requirement] Merge2610 Artisan Harbour")

total = sum(int(b["max_level"]) for z in zones for b in z["buildings"])
ok(total == 45, "the fixture totals 45 upgrade stages, as Merge2608 does", str(total))
per = [sum(int(b["max_level"]) for b in z["buildings"]) for z in zones]
ok(per == [15, 15, 15], "15 per scene", str(per))
n_items = sum(len(ch["items"]) for p in producers for ch in p["chains"])
ok(n_items == 40 and len(producers) * 2 == 8,
   "and 40 item icons across 8 producer icons", "%d items" % n_items)

tmp2 = tempfile.mkdtemp()
try:
    md = arg._render_markdown(NEW_DOC, "Merge", "Draft", "Tester", md_dir=tmp2)
    ok("Buildings (8) — 15 upgrade stages in this zone" in md,
       "each zone reports its building count and stage total", md[:0])
    ok("| Levels |" in md and "| Signage |" in md,
       "the buildings table carries Levels and Signage as columns")
    ok("— none —" in md,
       "a building with no signage renders as an explicit 'none', not a blank")
    ok("MARKET" in md, "and declared wording is carried through")
    ok("Scene Material" in md and "woven willow" in md,
       "the scene material reaches the document")
    ok("Producers (4)" in md and "8 producer icons" in md and "40 item icons" in md,
       "the producer section reports what has to be drawn")
    ok(md.count("| Chain | Family |") == 4, "one chain table per producer",
       "got %d" % md.count("| Chain | Family |"))
    ok("Main Tone" not in md,
       "an empty main_tone renders nothing rather than an empty row",
       "Merge is polychrome and leaves it empty")
    ok("Construction/Reveal Steps" not in md,
       "and empty construction_steps are omitted when buildings are used instead")
finally:
    shutil.rmtree(tmp2, ignore_errors=True)

# ---- the type hint tells the model the real numbers -------------------------
hint = arg.TYPE_GAMEPLAY_HINTS["Merge"]
for token in ("8 per scene", "15 PER SCENE", "exactly 4", "exactly 2 chains",
              "exactly 5 items", "scene_material", "max_level"):
    ok(token in hint, "the Merge hint states %r" % token)

# ---- the pass plan -----------------------------------------------------------
# The point of planning separately from generating: a full Merge event is ~96
# renders at roughly half a minute each, and nobody should learn that by
# starting the job.
import artreq_images as _ai

plan = _ai.plan_passes(NEW_DOC, _ai.PASSES)
counts = {e["pass"]: e["count"] for e in plan["passes"]}
ok(counts["scene"] == 3, "3 scene overviews", str(counts))
ok(counts["building"] == 24, "24 building sheets", str(counts))
ok(counts["stage"] == 21, "21 earlier stages (45 total minus the 24 finals)", str(counts))
ok(counts["icon"] == 40, "40 chain item icons", str(counts))
ok(counts["producer"] == 8, "8 producer icons", str(counts))
ok(plan["total"] == 96, "96 renders in total", str(plan["total"]))

ok(_ai.plan_passes(NEW_DOC, ["scene"])["total"] == 3,
   "asking for scene only plans only the scenes")
ok(_ai.plan_passes(OLD_DOC, _ai.PASSES)["total"] == 1,
   "a pre-schema draft plans just its one scene - nothing else is available",
   str(_ai.plan_passes(OLD_DOC, _ai.PASSES)))
ok(_ai.DEFAULT_PASSES == ("scene",),
   "and the default stays scene-only, so existing callers are unaffected")

# stage count must equal total stages minus the finals, whatever the levels are
tot_levels = sum(int(b["max_level"]) for z in NEW_DOC["zones"] for b in z["buildings"])
n_bld = sum(len(z["buildings"]) for z in NEW_DOC["zones"])
ok(counts["stage"] == tot_levels - n_bld,
   "the stage pass renders exactly the levels below each final",
   "%d stages, %d levels, %d buildings" % (counts["stage"], tot_levels, n_bld))

# ---- characters are RE-COSTUMED, not invented --------------------------------
# Merge2608 says it outright: "reuse existing guests as boarding travelers",
# "reuse existing guests in gala formalwear". The shipped Icon/Character set is
# the same SkipBo animals in different outfits per scene, so the art requirement
# is the COSTUME - the character underneath is already drawn. A schema that only
# had free-text design_notes could not say which cast member it meant.
CH = arg.CHARACTER_SCHEMA
for f in ("npc_id", "base_character", "costume"):
    ok(f in CH["properties"] and f in CH["required"], "a character declares %s" % f)
ok("re-costume" in CH["properties"]["base_character"]["description"].lower()
   or "RE-COSTUME" in CH["properties"]["base_character"]["description"],
   "and the schema states that the cast is reused, not invented")
ok("101" in CH["properties"]["npc_id"]["description"],
   "npc_id names the favourability slots the game data actually uses")

# ---- event extras ------------------------------------------------------------
EX = arg.EVENT_EXTRAS_SCHEMA
for f in ("hot_air_balloon", "achievement_badge", "task_icons"):
    ok(f in EX["properties"] and f in EX["required"], "event_extras declares %s" % f)
ok(EX["properties"]["task_icons"]["type"] == "array",
   "task_icons is a list - the template asks for 3-4 of them")
ok("event_extras" in S["properties"] and "event_extras" in S["required"],
   "and event_extras is a required top-level section")

_ex_doc = dict(NEW_DOC, event_extras={
    "hot_air_balloon": "a woven willow balloon with canvas panels",
    "achievement_badge": "a basket-weave crest",
    "task_icons": ["a brass compass", "a rope knot", "a sea-glass float"]})
ok(_ai.plan_passes(_ex_doc, ["extras"])["total"] == 5,
   "the extras pass plans balloon + badge + 3 task icons",
   str(_ai.plan_passes(_ex_doc, ["extras"])))
_empty = dict(NEW_DOC, event_extras={"hot_air_balloon": "", "achievement_badge": "",
                                     "task_icons": []})
ok(_ai.plan_passes(_empty, ["extras"])["total"] == 0,
   "and plans nothing when the type has no extras")
ok("extras" in _ai.PASSES, "extras is a real pass")

# every pass a document can carry is plannable
full = _ai.plan_passes(_ex_doc, _ai.PASSES)
ok(full["total"] == 101,
   "a full Merge event plans 101 renders - 96 plus the 5 extras", str(full["total"]))

# ---- Dessert: recipes and ingredients ----------------------------------------
# The mechanic-specific classes do not overlap. A Merge document has producers
# and chains; a Dessert document has recipes and ingredients. Each pass is gated
# on what the document carries, so neither type is ever asked for the other's
# assets - a Dessert event has no producers to render, and the 101-render Merge
# plan above must not grow just because these passes now exist.
ok(full["total"] == 101,
   "adding the recipe and ingredient passes does not change a Merge plan",
   str(full["total"]))

for f in ("recipes", "ingredients", "event_motif"):
    ok(f in S["properties"] and f in S["required"],
       "the schema declares %s as a required top-level field" % f)
RS = arg.RECIPE_SCHEMA
for f in ("recipe_number", "name", "zone_number", "description", "palette",
          "ingredients", "secret"):
    ok(f in RS["properties"] and f in RS["required"], "a recipe declares %s" % f)
ok(RS["properties"]["ingredients"]["type"] == "array",
   "a recipe lists the ingredients it needs")
IS_ = arg.INGREDIENT_SCHEMA
for f in ("name", "description", "palette", "secret"):
    ok(f in IS_["properties"] and f in IS_["required"], "an ingredient declares %s" % f)

# Shipped Dessert2608 shape: 24 common + 15 secret recipes, 6 + 5 ingredients.
_recipes = [{"recipe_number": n, "name": "Dish %d" % n, "zone_number": (n - 1) % 4 + 1,
             "description": "a plated dish", "palette": "amber",
             "ingredients": ["Party Flour"], "secret": n > 24}
            for n in range(1, 40)]
_ings = [{"name": "Ing %d" % n, "description": "a raw component", "palette": "amber",
          "secret": n > 6} for n in range(1, 12)]
_dessert = dict(NEW_DOC, producers=[], recipes=_recipes, ingredients=_ings,
                event_motif="the 5th-anniversary numeral 5")
dcounts = {e["pass"]: e["count"] for e in _ai.plan_passes(_dessert, _ai.PASSES)["passes"]}
ok(dcounts["recipe"] == 39, "39 recipes planned (24 common + 15 secret)", str(dcounts))
ok(dcounts["ingredient"] == 11, "11 ingredients planned (6 common + 5 secret)", str(dcounts))
ok(dcounts["icon"] == 0 and dcounts["producer"] == 0,
   "a Dessert document plans no chain or producer icons at all", str(dcounts))
ok(_ai.plan_passes(NEW_DOC, ["recipe", "ingredient"])["total"] == 0,
   "and a Merge document plans no recipes or ingredients",
   str(_ai.plan_passes(NEW_DOC, ["recipe", "ingredient"])))

# The geometry is measured off the shipped icons, not chosen - see icon_sheet.
import icon_sheet as _is
ok((_is.RECIPE_CANVAS, _is.INGREDIENT_CANVAS) == (342, 224),
   "recipe and ingredient canvases match the shipped Dessert2608 art",
   "%s" % ((_is.RECIPE_CANVAS, _is.INGREDIENT_CANVAS),))
ok(_is.INGREDIENT_FRACTION > _is.RECIPE_FRACTION,
   "an ingredient fills more of its canvas than a recipe does (0.90 vs 0.78)",
   "%s vs %s" % (_is.INGREDIENT_FRACTION, _is.RECIPE_FRACTION))
ok(_is.stroke_for(_is.CANVAS) == _is.STROKE_PX and _is.stroke_for(342) == 4,
   "the rim scales with the canvas, so a 342px recipe is not given a hairline",
   "180 -> %d, 342 -> %d" % (_is.stroke_for(180), _is.stroke_for(342)))

# The prompts have to say the two Dessert-specific things the shipped art shows:
# the motif is worked into every dish, and an ingredient is raw, not plated.
import asset_prompts as _ap
_rp = _ap.recipe_prompt(_recipes[0], "the 5th-anniversary numeral 5", "amber")
ok("5th-anniversary numeral 5" in _rp, "the recipe prompt carries the event motif")
ok("KEYLINE" in _rp.upper(),
   "and asks for the dark keyline a food icon has and the scenery does not")
ok("Party Flour" in _rp, "and names the ingredients the dish is made from")
ok("motif" not in _ap.recipe_prompt(_recipes[0], "", "amber").lower(),
   "an event with no motif gets no motif block, rather than an empty one")
_ip = _ap.ingredient_prompt(_ings[0])
ok("RAW COMPONENT" in _ip, "the ingredient prompt insists on a raw component")
ok("NOT plated" in _ip, "and names the failure it guards against")

dhint = arg.TYPE_GAMEPLAY_HINTS["Dessert"]
for token in ("exactly 4", "11, 11, 12, 12", "46", "24 common", "6 common and 5 secret",
              "event_motif", "ISOMETRIC"):
    ok(token in dhint, "the Dessert hint states %r" % token)
ok("no text at all" in dhint.lower(),
   "and that Dessert buildings carry no signage, unlike Merge")

tmp3 = tempfile.mkdtemp()
try:
    dmd = arg._render_markdown(_dessert, "Dessert", "Draft", "Tester", md_dir=tmp3)
    ok("Recipes (39)" in dmd and "Ingredients (11)" in dmd,
       "the document reports what has to be drawn")
    ok("#### Common recipes (24)" in dmd and "#### Secret recipes (15)" in dmd,
       "common and secret recipes are tabled separately")
    ok("the 5th-anniversary numeral 5" in dmd, "and the motif is stated once, up front")
    ok("Producers (" not in dmd, "a Dessert document renders no producer section")
    mmd = arg._render_markdown(NEW_DOC, "Merge", "Draft", "Tester", md_dir=tmp3)
    ok("Recipes (" not in mmd and "Ingredients (" not in mmd,
       "and a Merge document renders no recipe section")
finally:
    shutil.rmtree(tmp3, ignore_errors=True)

print("")
if fails:
    print("%d CHECK(S) FAILED" % len(fails)); sys.exit(1)
print("ALL ARTREQ SCHEMA TESTS PASSED")

