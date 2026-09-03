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
    # The label gained spaces when the document moved to the production team's
    # template ("Construction / Reveal Steps"), so this asserts the STEPS survive
    # rather than one spelling of the heading.
    ok("Reveal Steps" in md and "1. a" in md and "2. b" in md,
       "and its construction_steps still appear, numbered")
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

# ---- Blocks: construction materials and characters ---------------------------
# Blocks' mechanic-specific class is the construction material icon, the way
# recipes are Dessert's and chains are Merge's. Characters are shared across
# every type but had no pass until now - six sections of the Blocks style guide
# are about character construction and none of it was reachable.
ok("construction_materials" in S["properties"] and "construction_materials" in S["required"],
   "construction_materials is a required top-level field")
MS = arg.MATERIAL_SCHEMA
for f in ("name", "description", "palette"):
    ok(f in MS["properties"] and f in MS["required"], "a material declares %s" % f)

_blocks = dict(NEW_DOC, producers=[], recipes=[], ingredients=[],
               construction_materials=[{"name": "Driftwood Plank",
                                        "description": "a bundle of pale planks",
                                        "palette": "warm cream"} for _ in range(6)],
               characters=[{"name": "Cove Keeper", "role": "Chef/host", "npc_id": "101",
                            "base_character": "turtle", "costume": "a lantern-keeper's coat",
                            "design_notes": "", "scene_actions": ["waves"]}])
bc = {e["pass"]: e["count"] for e in _ai.plan_passes(_blocks, _ai.PASSES)["passes"]}
ok(bc["material"] == 6, "6 construction materials planned", str(bc))
ok(bc["character"] == 1, "and the characters pass counts them", str(bc))
ok(bc["recipe"] == 0 and bc["icon"] == 0 and bc["producer"] == 0,
   "a Blocks document plans no recipes, chains or producers", str(bc))
ok(_ai.plan_passes(NEW_DOC, ["material"])["total"] == 0,
   "and a Merge document plans no construction materials")

import asset_prompts as _ap2
mp = _ap2.material_prompt({"name": "Driftwood Plank", "description": "a bundle of pale planks",
                           "palette": "warm cream"})
ok("CONSTRUCTION MATERIAL" in mp, "the material prompt says what class it is")
ok("not a finished structure" in mp, "and rules out the failure it guards against")

cp = _ap2.character_prompt({"name": "Cove Keeper", "base_character": "turtle",
                            "costume": "a lantern-keeper's coat", "scene_actions": ["waves"]})
ok("turtle" in cp, "the character prompt names the existing cast member")
ok("already designed" in cp,
   "and states that the animal is reused and only the costume is new")
ok("40-55%" in cp, "it carries the head-to-body ratio the style guide fixes")
ok("THUMBNAIL SIZE" in cp, "and the rule that decides the face")
ok("NOT anime proportions" in cp, "and the proportions it must not use")
cp2 = _ap2.character_prompt({"name": "Newcomer", "costume": "a coat"})
ok("already designed" not in cp2,
   "a genuinely new character gets no reuse clause, rather than an empty one")

# ---- Blocks gets its own render spec and failure block ------------------------
import art_reference as _ar2
bf = _ar2.named_failures("Blocks")
ok("hyper-detailed textures" in bf, "the Blocks block carries the doc's own negatives")
ok("70%" in bf, "and its 70/20/10 detail ratio")
for clash, why in (("Matte only", "matte-only, which the doc contradicts"),
                   ("render lineless", "lineless, which the doc contradicts"),
                   ("valve wheels", "producer fittings Blocks does not have")):
    ok(clash not in bf, "and does NOT carry %s" % why)
# Sim is now the only type without a studio document of its own.
ok("Matte only" in _ar2.named_failures("Sim"),
   "while a type with no document of its own keeps the default block")
ok("empty rooms" not in _ar2.named_failures("Sort"),
   "Sort has its own block, and drops the one negative its document got wrong")
ok("Matte only" not in _ar2.named_failures("Merge"),
   "Merge no longer gets it either - Merge Art.docx asks for glossy, not matte")

# ---- the progression pass ----------------------------------------------------
# A Blocks scene is delivered as a build sequence, not a single picture, and the
# cost line has to say so: five scenes of six steps is 45 renders, not five.
_prog = dict(NEW_DOC, producers=[], recipes=[], ingredients=[],
             construction_materials=[],
             characters=[{"name": "Keeper", "role": "Chef/host", "npc_id": "101",
                          "base_character": "turtle", "costume": "a teal coat",
                          "design_notes": "", "scene_actions": []}])
for z in _prog["zones"]:
    z = z
_prog["zones"] = [dict(z, construction_steps=["hero lantern", "boat shed", "farm cart",
                                              "hamper", "jars", "bunting"])
                  for z in _prog["zones"]]
pc = {e["pass"]: e["count"] for e in _ai.plan_passes(_prog, _ai.PASSES, "Blocks")["passes"]}
ok("progression" in _ai.PASSES, "progression is a real pass")
ok(pc["progression"] == 3 * (6 + 2 + 1),
   "each scene plans derelict + empty + one frame per step + characters",
   "6 steps over 3 zones should be %d, got %s" % (3 * 9, pc["progression"]))
_noch = dict(_prog, characters=[])
ok({e["pass"]: e["count"] for e in _ai.plan_passes(_noch, ["progression"], "Blocks")["passes"]}
   ["progression"] == 3 * 8,
   "with no characters the character frame is not planned or charged")
_one = dict(_prog, zones=[dict(_prog["zones"][0], construction_steps=["only one"])])
ok(_ai.plan_passes(_one, ["progression"], "Blocks")["total"] == 0,
   "a zone with fewer than two steps has no sequence and plans nothing")
ok(_ai.plan_passes(NEW_DOC, ["progression"], "Merge")["total"] == 0,
   "and a Merge document, whose zones use buildings not steps, plans none")
ok(_ai.plan_passes(_prog, ["progression"], "Sort")["total"] == 0,
   "a non-Blocks type is never offered a sequence, however many steps it lists",
   "legacy Sort and Sim drafts use construction_steps as free-text reveal notes")
ok(_ai.PROGRESSION_TYPES == {"Blocks"}, "and Blocks is the only type that gets one")

# The prompts encode the rules that document paid for.
import asset_prompts as _ap3
ok("HERO" in _ap3.built_scene(["a lantern", "a shed"]),
   "the base prompt names the first step as the hero")
one = _ap3.remove_one("the wooden boat shed")
ok("ONLY change" in one, "a hero/structure removal is a single blunt change")
many = _ap3.remove_many(["the cart", "the jars", "the bunting"])
ok("the cart, the jars and the bunting" in many,
   "a bulk removal lists every item not yet built, cumulatively", many[:120])
ok("ALL of those" in many, "and insists on all of them")
der = _ap3.derelict_prompt({"zone_theme": "Cove"})
ok("colour drained" in der and "FROZEN" in der,
   "the before frame changes the season and the light, not the exposure")
ok("Nothing built, nothing placed" in der, "and it is empty of built objects")
ch = _ap3.add_characters([{"base_character": "turtle", "costume": "a teal coat"}])
ok("turtle" in ch and "ONLY change" in ch,
   "characters are one additive hop naming the existing cast")
ok("do NOT add anything" not in ch,
   "and the hold clause drops its no-adding rule for that one hop")
for p_ in (one, many, der):
    ok("No text anywhere" in p_, "every sequence hop still bans text")

# ---- art director feedback, encoded ------------------------------------------
# Direction on the first current-style regeneration, section 11 of the rules.
der = _ap3.derelict_prompt({"zone_theme": "Cove"})
ok("FROZEN" in der, "the before frame is FROZEN, not merely damaged")
ok("REMOVE EVERYTHING HUMANISED" in der, "and empty of everything humanised")
ok("NO DEBRIS OR RUINS" in der,
   "debris and ruins are excluded too - they are human traces, which was the wrong instinct",
   "the earlier version asked for broken boards and scattered litter")
ok("same fields, the same hills" in der, "while the natural landscape is kept")

# FROZEN IS NOT THE RULE. It was one suggestion for making a scene distinctive,
# and applying it to every scene reproduces the sameness it was meant to fix.
dv = _ap3.derelict_prompt({"zone_theme": "Fruit Shop", "interior": True,
                           "before_state": "shuttered and dusty, one cold shaft through a gap"})
ok("shuttered and dusty" in dv, "a zone's own before_state is used verbatim")
ok("frozen" not in dv.lower() and "winter" not in dv.lower(),
   "and nothing forces frost or winter onto a scene that did not ask for it")
ok("REMOVE EVERYTHING HUMANISED" in dv and "NO DEBRIS OR RUINS" in dv,
   "while the mandatory part - empty of humanised things - still holds")
ok("room shell" in dv, "an interior keeps its shell")
ok("landscape - fields, hills" in _ap3.derelict_prompt(
       {"zone_theme": "Yard", "before_state": "a winter-bare orchard"}),
   "and an exterior keeps its landscape")

# An interior has no fields to freeze - the shell is what gets kept instead.
din = _ap3.derelict_prompt({"zone_theme": "Fruit Shop", "interior": True})
ok("this room in deep winter" in din, "an interior zone freezes the ROOM, not the land")
ok("same walls, floor, ceiling" in din, "keeping the shell that makes it this location")
ok("REMOVE EVERYTHING THAT SHOWS IT IS USED" in din, "and stripping every trace of use")
ok("NO DEBRIS" in din, "clutter excluded here too, for the same reason")
ok("fields" not in din, "and it never mentions fields, which an interior does not have")

bs = _ap3.built_scene(["a hero", "a shed", "a cart", "a hamper"])
ok("do NOT cluster" in bs, "the base spreads its objects instead of clustering them")
ok("bottom centre" in bs and "LEFT" in bs and "RIGHT" in bs,
   "and gives each build step its own region, so the sequence jumps between areas")
ok("HUMANISED ACTIVITY" in bs, "the steps read as humanised activity")

CH = [{"name": "A", "base_character": "badger", "costume": "a waistcoat"},
      {"name": "B", "base_character": "fox", "costume": "a pinafore"}]
first = _ap3.add_characters(CH[:1], arriving=CH[:1])
second = _ap3.add_characters(CH, arriving=CH[1:])
ok("badger" in first and "fox" not in first,
   "the first character frame introduces one character only")
ok("fox" in second and "stays exactly where they are" in second,
   "and the second adds the other while holding the first in place", second[:200])
ok("adding 1 character" in first.lower(),
   "characters arrive one by one, not as a group", first[:120])
# They came in as small background figures first time. The style guide asks for
# character-centric framing with large foreground elements close to camera.
# Two scales, both kept, because they trade against each other: "beside"
# composites cleanly and stays faithful to the base, "foreground" reads far
# better but repaints the frame (67-95% of columns change).
ok("BESIDE the central hero" in first and "scene's own scale" in first,
   "the DEFAULT is the lateral, scene-scale placement that composites cleanly")
ok("THIRD OF THE FRAME" not in first,
   "and nothing forces the large near-camera framing onto it")
big = _ap3.add_characters(CH[:1], arriving=CH[:1], scale="foreground")
ok("THIRD OF THE FRAME'S HEIGHT" in big, "the foreground scale sizes them against the frame")
ok("close to camera" in big.lower(),
   "and stands them in the near ground rather than back beside the hero")
ok("small background figure is wrong" in big,
   "with the failure named, since that is what the first pass produced")
ok("hiding the hero is" in big,
   "while still not being allowed to hide the hero behind them")
ok(_ap3.CHARACTER_SCALES == ("beside", "foreground"), "both scales stay available")

_prog2 = dict(_prog, characters=CH)
ok({e["pass"]: e["count"] for e in _ai.plan_passes(_prog2, ["progression"], "Blocks")["passes"]}
   ["progression"] == 3 * (6 + 2 + 2),
   "and the plan charges one frame per arriving character")

# ---- the zone hints the sequence pass needs ----------------------------------
# Both were used by the pass before they were declared, so a generated document
# could not have carried either. interior changes the 'before' frame; blunt_steps
# forces a dedicated removal hop where a bulk one would silently fail.
for f in ("interior", "blunt_steps", "before_state"):
    ok(f in Z["properties"], "a zone can declare %s" % f)
    ok(f not in Z["required"],
       "and %s stays optional - it is a production hint, usually absent" % f)
ok(Z["properties"]["blunt_steps"]["items"]["type"] == "integer",
   "blunt_steps is a list of step numbers")
ok("hero" in Z["properties"]["blunt_steps"]["description"]
   and "flat ground element" in Z["properties"]["blunt_steps"]["description"],
   "and its description names the three classes a bulk removal cannot touch")
ok(Z["properties"]["interior"]["type"] == "boolean", "interior is a flag")
ok("variety" in Z["properties"]["before_state"]["description"]
   and "not the rule" in Z["properties"]["before_state"]["description"],
   "before_state's description says variety is the point and frozen is optional")
ok(Z["additionalProperties"] is False,
   "the zone schema stays closed, so a declared field is the only way to pass one")

# ---- Sort goods: the last mechanic-specific class to get a pass ---------------
# Merge has chains and producers, Dessert recipes and ingredients, Blocks
# construction materials - Sort's loop is literally sorting goods onto shelves
# and it had nothing. Goods live on the ZONE because the catalogue is per floor.
ok("goods" in Z["properties"], "a zone can declare its goods catalogue")
ok("goods" not in Z["required"], "and it stays optional, like the other zone hints")
gi = Z["properties"]["goods"]["items"]["properties"]
for f in ("name", "description", "silhouette", "palette"):
    ok(f in gi, "a good declares %s" % f)
_gd = Z["properties"]["goods"]["description"]
ok("REAL tool or product of THIS floor" in _gd,
   "the schema demands real items of the craft, not shapes with a theme painted on")
ok("no other floor gets a bottle" in _gd,
   "and forbids the same KIND of object twice across the whole event")
ok("share an OUTLINE" in _gd, "and the same outline twice across the whole event")
ok("repeats down every column" in _gd,
   "and says why per-floor rules were not enough - five floors spending one list "
   "of fifteen shapes gave the event fifteen objects painted five ways")
ok("enum" not in gi["silhouette"],
   "silhouette is free text, not an enum - a fixed list caps the catalogue at that "
   "many distinct objects")
ok("size_in_frame" not in gi,
   "there is no per-item size any more - shrinking individual goods is what left "
   "half the canvas empty")
ok("TWICE AS TALL AS IT IS WIDE" in gi["silhouette"]["description"],
   "the schema states the portrait rule measured off the shipped goods (0.48 aspect)")
ok("never by stretching it" in gi["silhouette"]["description"],
   "and that the height comes from the object, not from distorting it")

_sortdoc = dict(NEW_DOC, producers=[], recipes=[], ingredients=[],
                construction_materials=[],
                zones=[dict(z, goods=[{"name": "G%d" % i, "description": "a thing",
                                       "silhouette": "a shape",
                                       "palette": "blue"} for i in range(1, 16)])
                       for z in NEW_DOC["zones"]])
gc = {e["pass"]: e["count"] for e in _ai.plan_passes(_sortdoc, _ai.PASSES, "Sort")["passes"]}
ok("goods" in _ai.PASSES, "goods is a real pass")
ok(gc["goods"] == 45, "15 goods across 3 zones plans 45 renders", str(gc))
ok(_ai.plan_passes(NEW_DOC, ["goods"], "Merge")["total"] == 0,
   "and a document with no goods plans none")

import icon_sheet as _is2
ok(_is2.GOODS_CANVAS == (256, 512),
   "the goods canvas is the shipped 1:2 PORTRAIT, not a square", str(_is2.GOODS_CANVAS))
# 0.88 came from "longest content edge", which says an object touches ONE edge
# and nothing about the other. Re-measured across all 76 shipped icons: bbox is
# 0.86 of the width AND 0.90 of the height, because the object is itself 1:2.
ok(abs(_is2.GOODS_FRACTION - 0.90) < 0.01,
   "at the re-measured fill, matching the shipped 0.86 wide / 0.90 tall",
   str(_is2.GOODS_FRACTION))
ok(_is2.GOODS_FRACTION == _ap3.GOODS_FILL,
   "and the renderer and the geometry module agree on it")

gp = _ap3.goods_prompt({"name": "Ink Tin", "description": "a squat tin of ink",
                        "palette": "ink blue"})
ok("SHELF GOOD" in gp, "the goods prompt says what class it is")
# This used to assert the prompt demanded "taller than it is wide". That single
# line forced all 75 goods into one silhouette and is the reason a floor came
# back with four objects sharing a rounded rectangle. The canvas is 1:2; the
# OBJECT need not be, and a bowl drawn tall is the wrong bowl.
ok("TALLER THAN IT IS WIDE" not in gp,
   "the prompt no longer forces every good into one upright silhouette")
gp_bowl = _ap3.goods_prompt(
    {"name": "Clay Harp", "description": "a clay harp",
     "silhouette": "a tall narrow U-frame on a handle with a taut wire across its mouth",
     "palette": "blue"},
    siblings=["Amphora", "Kiln Post"])
ok("a tall narrow U-frame on a handle" in gp_bowl,
   "a good is asked for the outline that object actually has, verbatim")
ok("TWICE AS TALL AS IT IS WIDE" in gp_bowl,
   "and for the portrait proportions the shipped goods have")
ok("NEVER stretch, squash or distort" in gp_bowl,
   "and is told to get that from the object, not by distorting it")
ok("Amphora" in gp_bowl and "Kiln Post" in gp_bowl,
   "the rest of the floor's catalogue reaches the prompt, so a collision is avoidable")
ok("SHAPE" in gp_bowl.split("MUST NOT LOOK LIKE")[1],
   "and the instruction is to change the shape, not the colour")
# This used to assert the prompt ASKED for the keyline. It did, and icon_sheet
# painted one as well, so every good wore two - which is what made the icons read
# dark: 5th-percentile brightness 0.24 against a shipped 0.52. The generator now
# leaves the edge clean, like every other icon class, and add_stroke supplies the
# one keyline at a controlled 3px.
ok("NO TEXT, NO OUTLINE" in gp,
   "the generator leaves the edge clean - the keyline is added afterwards, once")
ok("No dark rim, no inked contour" in gp,
   "and is told not to shade a dark band round the silhouette either")
ok("lineless" not in gp.lower(), "so the floor rule does not leak into the icon")

print("")

# ---- no event type may render or accept another type's mechanics -------------
# A Blocks Art Requirement shipped with Merge's mechanics in it: every zone
# carried a buildings table with 1-3 upgrade levels and an A/B/C area column,
# none of which Blocks has. Three layers had to agree for that to happen, so all
# three are checked here - what the schema ACCEPTS, what the prompt ASKS for, and
# what the renderer PRINTS.
#
# The document below is deliberately stuffed with EVERY mechanic. The question is
# not "does Blocks look right when handed Blocks data" - it is "does Blocks stay
# Blocks when handed Merge's", which is the case that actually failed.

EVERY_MECHANIC = {
    "title": "[Artistic Requirement] X2610 Kitchen Sink",
    "gameplay_overview": {"theme": "t", "duration_summary": "14 days",
                          "narrative_packaging": "n", "loop_summary": "l",
                          "resource_planning": "r", "major_changes": "None."},
    "zones": [{
        "zone_number": 1, "zone_theme": "The Test Zone", "palette_or_mood": "warm",
        "main_tone": "WARM", "scene_material": "timber", "before_state": "bare",
        "construction_steps": ["step one", "step two"],
        "buildings": [{"building_number": 1, "name": "Test Hall", "area": "A",
                       "silhouette": "a shape", "key_features": ["f1", "f2"],
                       "max_level": 3, "signage": ["SIGN"]}],
        "goods": [{"name": "Jar", "silhouette": "round", "palette": "amber"}],
        "reference_notes": "notes",
    }],
    "producers": [], "recipes": [], "ingredients": [],
    "construction_materials": [{"name": "Beam", "description": "d", "palette": "brown"}],
    "characters": [], "event_motif": "a motif",
    "art_direction_guidance": {"lean_into": ["x"], "avoid": ["y"]},
    "event_extras": {}, "gameplay_specific_content": "content",
    "ui_icon_refresh_notes": "notes", "notes": "notes",
}

MERGE_ONLY_MARKERS = {
    "levels": ("| Levels |", "upgrade stages"),
    "buildings": ("Room / Building / Workshop Breakdown",),
    "areas": ("spread across areas A/B/C",),
}

for etype in ("Merge", "Dessert", "Blocks", "Sort", "Sim", "Hunt"):
    shape = arg.zone_shape(etype)
    md = arg._render_markdown(dict(EVERY_MECHANIC), etype, "Draft")
    for key, marks in MERGE_ONLY_MARKERS.items():
        present = any(m in md for m in marks)
        ok(present == shape[key],
           "%s renders %s only if its shape has them" % (etype, key),
           "rendered=%s shape=%s" % (present, shape[key]))

    s = arg.schema_for(etype)
    zp = s["properties"]["zones"]["items"]["properties"]
    ok((zp["buildings"].get("maxItems") != 0) == shape["buildings"],
       "%s's schema accepts buildings only if it has them" % etype)
    bi = zp["buildings"].get("items", {}).get("properties", {})
    ok(shape["levels"] or "max_level" not in bi,
       "%s's schema does not even offer max_level" % etype)
    for coll in arg.ALL_COLLECTIONS:
        allowed = coll in arg.TYPE_COLLECTIONS.get(etype, set())
        ok((s["properties"][coll].get("maxItems") != 0) == allowed,
           "%s may fill %s: %s" % (etype, coll, allowed))

# An unknown type must inherit nothing. The failure mode for a type nobody has
# calibrated should be a thin document that says so, not a confident Merge one.
unknown = arg.zone_shape("SomeNewMode")
ok(not unknown["buildings"] and not unknown["levels"] and not unknown["verified"],
   "an unrecognised event type inherits no mechanics at all")

# The item schema goes with the field. Leaving it in would keep shipping Merge's
# building vocabulary - "Merge2608 ships 45 stages across 25 buildings" - inside
# every Blocks prompt, which is the pressure that caused this in the first place.
blocks_schema = json.dumps(arg.schema_for("Blocks"))
ok("max_level" not in blocks_schema and "45 stages" not in blocks_schema,
   "a Blocks schema carries no trace of Merge's building vocabulary")
ok(len(blocks_schema) < len(json.dumps(arg.schema_for("Merge"))),
   "and is smaller than Merge's as a result, not merely differently worded")

# The hints have to agree with the table, or the prompt argues with the schema.
for etype, shape in arg.ZONE_SHAPES.items():
    hint = arg.TYPE_GAMEPLAY_HINTS[etype]
    if not shape["buildings"]:
        ok("NO buildings" in hint,
           "the %s hint says outright that it has no buildings" % etype)
    if not shape["levels"]:
        ok("upgrade level" in hint.lower() or "levels" in hint.lower(),
           "the %s hint addresses upgrade levels rather than leaving them open" % etype)

# Regeneration is bound by the same shape as the first draft - feedback on a
# Blocks document must not be able to hand back Merge buildings.
regen = arg._regenerate_schema("Blocks")
ok(regen["properties"]["variation"]["properties"]["zones"]["items"]
       ["properties"]["buildings"].get("maxItems") == 0,
   "a Blocks regeneration cannot reintroduce buildings either")

# ---- construction-step counts are enforced, not suggested --------------------
# Six Harvest Fair variations shipped at 6/6/6/6/6 = 30 when both shipped Blocks
# events run 7/7/8/8/8 = 38. Nothing caught it: construction_steps had no
# minItems or maxItems, so any count validated, and the hint led with the bottom
# of the band. Both halves are checked here, because fixing only the prose would
# leave the schema still accepting a three-step scene.

for etype in ("Blocks", "Sort", "Sim"):
    shape = arg.zone_shape(etype)
    cs = (arg.schema_for(etype)["properties"]["zones"]["items"]
             ["properties"]["construction_steps"])
    ok(cs.get("minItems") == shape["steps_min"] and cs.get("maxItems") == shape["steps_max"],
       "%s's step band reaches the schema as minItems/maxItems" % etype,
       "schema %s-%s vs shape %s-%s" % (cs.get("minItems"), cs.get("maxItems"),
                                        shape["steps_min"], shape["steps_max"]))
    ok(cs["minItems"] >= 3 and cs["maxItems"] > cs["minItems"],
       "and the %s band is a real range, not a fixed number" % etype)

# A type that builds through discrete buildings has no step band to enforce.
for etype in ("Merge", "Dessert"):
    cs = (arg.schema_for(etype)["properties"]["zones"]["items"]
             ["properties"]["construction_steps"])
    ok(cs.get("maxItems") == 0,
       "%s has no construction steps at all, so no band applies" % etype)

# The profile is the default the prompt asks for, and it must match the table -
# a hint that says one number while the schema enforces another is how the two
# drift apart silently.
blocks = arg.zone_shape("Blocks")
ok(blocks["shipped_profile"] == (7, 7, 8, 8, 8),
   "the Blocks shipped profile is 7/7/8/8/8, from Blocks2607 and blocks2604 data")
ok(sum(blocks["shipped_profile"]) == 38, "which totals the 38 Blocks2607 states")
hint = arg.TYPE_GAMEPLAY_HINTS["Blocks"]
ok("7, 7, 8, 8, 8" in hint and "38" in hint,
   "and the Blocks hint asks for that profile by name")
ok(hint.index("7, 7, 8, 8, 8") < hint.index("6-8"),
   "stated BEFORE the 6-8 band, so the floor does not read as the default")
ok("FLOOR" in hint, "and the band's bottom is called a floor outright")

for etype, shape in arg.ZONE_SHAPES.items():
    p = shape.get("shipped_profile")
    if p:
        ok(all(shape["steps_min"] <= n <= shape["steps_max"] for n in p),
           "%s's shipped profile sits inside its own band" % etype,
           "%s vs %s-%s" % (list(p), shape["steps_min"], shape["steps_max"]))

# A zone below the band must be unrepresentable, not merely discouraged.
cs = (arg.schema_for("Blocks")["properties"]["zones"]["items"]
         ["properties"]["construction_steps"])
ok(cs["minItems"] == 6,
   "a Blocks zone cannot be emitted with fewer than six steps")
ok(cs["maxItems"] == 8,
   "nor with more than eight")

# ---- the asset classes a Blocks event actually ships -------------------------
# Read off trunk/develop/client/Skipbo/Assets/ResourcesAb/Events/Blocks2607.
# Generated documents used to describe scenes, steps, materials and characters
# and stop there - no build-step icons, no block art, no board backgrounds, no
# panel list, and section 8 deferred the VFX entirely. All of it ships.

BLOCKS_DOC = {
    "title": "[Artistic Requirement] Blocks2609 Shipped Asset Check",
    "gameplay_overview": {"theme": "t", "duration_summary": "16 days",
                          "narrative_packaging": "n", "loop_summary": "l",
                          "resource_planning": "r", "major_changes": "None."},
    "zones": [{"zone_number": i, "zone_theme": "Z%d" % i, "palette_or_mood": "warm",
               "main_tone": "WARM", "scene_material": "timber",
               "before_state": "bare", "buildings": [], "goods": [],
               # The shipped profile, so the derived counts below are the real ones.
               "construction_steps": ["s%d" % n for n in range((7, 7, 8, 8, 8)[i - 1])],
               "reference_notes": "n"} for i in range(1, 6)],
    "producers": [], "recipes": [], "ingredients": [],
    "construction_materials": [{"name": "M%d" % i, "description": "d", "palette": "p"}
                               for i in range(7)],
    "characters": [{"name": "Gift", "role": "Gift NPC", "npc_id": "",
                    "base_character": "squirrel", "costume": "apron",
                    "design_notes": "d", "scene_actions": ["a"]}],
    "event_motif": "m", "art_direction_guidance": {"lean_into": [], "avoid": []},
    "event_extras": {}, "gameplay_specific_content": "", "ui_icon_refresh_notes": "",
    "notes": "",
}
blocks_md = arg._render_markdown(BLOCKS_DOC, "Blocks", "Draft")

ok("### UI Panels" in blocks_md,
   "the document has a UI Panels section of its own")
for col in ("Panel Name", "Previous UI Panel", "New Theme Reference"):
    ok(col in blocks_md, "and its columns include %s" % col)
ok(len(arg.BLOCKS_UI_PANELS) == 31,
   "all 31 shipped Blocks2607 panels are listed",
   "26 prefabs plus 5 that carry only a texture folder")
ok("UIBlocks2607Panel" in blocks_md and "UIBlocks2607UndoConfirmPopPanel" in blocks_md,
   "including both the biggest and the smallest")
ok("69 textures" in blocks_md,
   "with the texture count, so a re-skin estimate is not a guess",
   "Main event board is 69 textures; Share is none - they cannot cost the same")

# The panel list is data. A model asked for panel names produces plausible ones.
schema_blob = json.dumps(arg.schema_for("Blocks"))
# One panel name appears as an example in the ui_panels description, which is
# fine and useful. The LIST must not be there - that is what would let a model
# treat panel names as its own to choose.
in_schema = [a for _n, a, _t in arg.BLOCKS_UI_PANELS if a in schema_blob]
ok(len(in_schema) <= 1,
   "the panel LIST is not in the generation schema - the tool supplies it",
   "found: %s" % in_schema)
ok("ui_panels" in arg.schema_for("Blocks")["properties"],
   "though a model may add per-panel direction for panels this theme changes")

for group in ("Build-step icons", "Material guide states", "Playable block art",
              "Board backgrounds", "Scene photos", "Header backgrounds"):
    ok(group in blocks_md, "section 5 lists %s" % group)

ok("38 icons, one per construction step" in blocks_md,
   "build-step icons are counted from the steps, not guessed",
   "Icon/Building1..5 ships 7/7/8/8/8 = 38")
ok("7 materials x 3 states = 21" in blocks_md,
   "and material guide states from the materials",
   "Icon/Book ships icon_mine_N_1..3")

# Section 8 used to say "to be agreed with the Art Lead" for every type.
ok("Build puff" in blocks_md and "Scene ambience" in blocks_md,
   "section 8 names the VFX a Blocks event ships instead of deferring it",
   "VX/Spine/Building_Cloud, Elf, Scene1..5")
merge_md = arg._render_markdown(dict(BLOCKS_DOC, zones=[]), "Merge", "Draft")
ok("Build puff" not in merge_md,
   "but only for Blocks - the other types still defer, honestly")

# Materials: the hint said 8-10, the builds say 6 and 7.
cm = arg.schema_for("Blocks")["properties"]["construction_materials"]
ok(cm.get("minItems") == 6 and cm.get("maxItems") == 7,
   "the material count is banded to what shipped, not to the old 8-10 claim")
ok("8-10" in arg.TYPE_GAMEPLAY_HINTS["Blocks"],
   "and the hint names the wrong number it replaced, so it is not reintroduced")
if fails:
    print("%d CHECK(S) FAILED" % len(fails)); sys.exit(1)
print("ALL ARTREQ SCHEMA TESTS PASSED")

