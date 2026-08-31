"""Does a generated Art Requirement match the production team's template?

The template arrived as a PDF from the production team. It has a header block
and ten numbered sections, and it is ONE document for every event type - its own
section 9 says so: "This allows the same Art Requirements template to support
Sim, Management, Blocks, Sort, Merge, or future event types".

So the thing to test is the pairing: the SHAPE is identical across types, and
section 5 carries each type's own elements - Merge's producers and chains,
Dessert's recipes and ingredients, Blocks' materials, Sort's goods.

Two things this also guards, both of which were wrong before:
  * the tool emitted six sections of its own invention, not these ten;
  * it was grounded in Event_Final_Art_Requirement.md, which is a
    multi-discipline Final Spec (Engineering, Product, Economy, QA) and does not
    describe an Art Requirement at all.

Run: python test_artreq_template.py
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import art_requirement_generator as arg                        # noqa: E402

fails = []


def ok(cond, msg, detail=""):
    if cond:
        print("OK   " + msg)
    else:
        print("FAIL " + msg + (("\n     " + str(detail)) if detail else ""))
        fails.append(msg)


def _zone(n, **kw):
    z = {"zone_number": n, "zone_theme": "Zone %d" % n, "palette_or_mood": "teal",
         "reference_notes": "a note"}
    z.update(kw)
    return z


def _building(i):
    return {"name": "Shed %d" % i, "area": "A", "building_number": i, "max_level": 2,
            "silhouette": "a low shed", "key_features": ["nets", "floats"], "signage": []}


MERGE = {
    "title": "[Artistic Requirement] Merge2612 Artisan Harbour",
    "gameplay_overview": {"theme": "Artisan Harbour", "duration_summary": "14 days",
                          "loop_summary": "Merge -> build -> progress",
                          "narrative_packaging": "A working harbour of makers."},
    "zones": [_zone(i, buildings=[_building(b) for b in range(8)]) for i in (1, 2, 3)],
    "producers": [{"producer_number": p, "name": "Kiln %d" % p, "palette": "teal",
                   "closed_description": "shut", "open_description": "open",
                   "chains": [{"chain_number": c, "family": "rope", "palette": "teal",
                               "items": ["a", "b", "c", "d", "e"]} for c in (1, 2)]}
                  for p in (1, 2, 3, 4)],
    "characters": [{"name": "Netmaker", "role": "NPC", "design_notes": "kind",
                    "scene_actions": ["mends a net"]}],
    "event_extras": {"achievement_badge": "a brass anchor", "hot_air_balloon": "",
                     "task_icons": ["net", "rope"]},
    "art_direction_guidance": {"lean_into": ["salt-worn timber"], "avoid": ["neon"]},
}

DESSERT = {
    "title": "[Artistic Requirement] Dessert2612 Orchard Patisserie",
    "gameplay_overview": {"theme": "Orchard Patisserie", "duration_summary": "17 days",
                          "loop_summary": "Collect -> cook -> serve"},
    "zones": [_zone(i) for i in (1, 2, 3, 4)],
    "recipes": ([{"name": "d%d" % i, "secret": False, "zone_number": 1,
                  "recipe_number": i, "description": "x", "palette": "p",
                  "ingredients": ["c0"]} for i in range(24)]
                + [{"name": "s%d" % i, "secret": True, "zone_number": 1,
                    "recipe_number": 100 + i, "description": "x", "palette": "p",
                    "ingredients": ["c0"]} for i in range(15)]),
    "ingredients": ([{"name": "c%d" % i, "secret": False, "description": "d", "palette": "p"}
                     for i in range(6)]
                    + [{"name": "x%d" % i, "secret": True, "description": "d", "palette": "p"}
                       for i in range(5)]),
    "characters": [],
}

BLOCKS = {
    "title": "[Artistic Requirement] Blocks2612 Harvest Fair",
    "gameplay_overview": {"theme": "Harvest Fair", "duration_summary": "14 days"},
    "zones": [_zone(i, construction_steps=["step %d" % s for s in range(6)],
                    before_state="bare ground") for i in (1, 2, 3, 4, 5)],
    "construction_materials": [{"name": "m%d" % i, "description": "d", "palette": "p"}
                               for i in range(6)],
    "characters": [],
}

SORT = {
    "title": "[Artistic Requirement] Sort2612 Hobby Loft",
    "gameplay_overview": {"theme": "Hobby Loft", "duration_summary": "14 days"},
    "zones": [_zone(i, goods=[{"name": "g%d" % g} for g in range(15)],
                    construction_steps=["s%d" % s for s in range(9)]) for i in range(1, 6)],
    "characters": [],
}

# ---- the shape is the template's, and it is the same for every type ---------
SECTIONS = [
    "## 1. Event / Gameplay Overview",
    "## 2. Theme & Overall Art Direction",
    "## 3. Art Content Overview",
    "## 4. Scene / Environment Requirements",
    "## 5. Gameplay / Event Asset Requirements",
    "## 7. UI / Icon Requirements",
    "## 8. Animation / VFX Requirements",
    "## 9. Additional Event-Specific Art Requirements",
    "## 10. SKB Feedback",
]
for doc, t in ((MERGE, "Merge"), (DESSERT, "Dessert"), (BLOCKS, "Blocks"), (SORT, "Sort")):
    md = arg._render_markdown(doc, t, "Draft")
    missing = [s for s in SECTIONS if s not in md]
    ok(not missing, "%s carries every non-optional template section, in order" % t, missing)
    ok(md.index("## 1.") < md.index("## 5.") < md.index("## 10."),
       "%s keeps the template's numbering order" % t)

# The template's own instruction: "Delete this section if the event does not
# require characters." So it is present when there are characters and absent when
# there are none - not an empty table either way.
ok("## 6. Character Requirements" in arg._render_markdown(MERGE, "Merge", "Draft"),
   "section 6 appears when the event has characters")
ok("## 6. Character Requirements" not in arg._render_markdown(DESSERT, "Dessert", "Draft"),
   "and is deleted when it has none, as the template instructs")

# ---- the header block ------------------------------------------------------
md = arg._render_markdown(MERGE, "Merge", "Design + Art Review", "Naveen")
for field in ("Event Name", "Event ID / Release", "Event Type / Gameplay", "Event Theme",
              "Event Duration", "Design Spec", "Design Owner", "Art Owner", "Status"):
    ok("| %s |" % field in md, "the Event Summary carries %s" % field)
ok(md.startswith("# [Art Requirements] Merge2612 - Artisan Harbour"),
   "and the title follows the naming convention", md.split("\n")[0])
ok("[Artistic Requirement]" not in md.split("\n")[0],
   "without double-prefixing the generated title into it")
ok(arg.STUDIO_DESIGN_OWNER in md and arg.STUDIO_ART_OWNER in md,
   "the owners default to the ones the template names")

# ---- section 5 is where the types differ -----------------------------------
def section5(doc, t):
    md = arg._render_markdown(doc, t, "Draft")
    return md[md.index("## 5."):md.index("### Detail")]

s5 = section5(MERGE, "Merge")
ok("Producers / generators" in s5 and "Merge chains" in s5,
   "Merge's section 5 is producers and chains")
ok("4 producers x 2 levels = 8 icons" in s5,
   "counted off the document, not asserted", s5)
ok("8 chains x 5 rungs = 40 item icons" in s5, "including the chain rungs")
ok("24 buildings, 48 upgrade points across 3 scenes" in s5,
   "and the buildings with their upgrade points")

s5 = section5(DESSERT, "Dessert")
ok("Common recipes" in s5 and "Secret recipes" in s5 and "Ingredients" in s5,
   "Dessert's section 5 is recipes and ingredients - not producers")
ok("Producers" not in s5, "with no trace of another type's elements")
ok("24 dishes across 4 zones" in s5 and "15 dishes" in s5,
   "split common from secret, as the shipped documents do", s5)
ok("6 common + 5 secret" in s5, "and the same for ingredients")

s5 = section5(BLOCKS, "Blocks")
ok("Construction material icons" in s5, "Blocks' section 5 is material icons")
ok("6 materials" in s5 and "30 steps across 5 scenes" in s5,
   "and its build steps, counted across the scenes", s5)

s5 = section5(SORT, "Sort")
ok("Goods / shelf icons" in s5, "Sort's section 5 is the goods catalogue")
ok("75 goods across 5 floors" in s5, "counted across the floors", s5)

# A type the map has never seen must still produce a document rather than an
# empty section - the template exists to absorb future event types.
s5 = section5({"title": "X", "zones": [_zone(1)], "characters": []}, "Hunt")
ok("Gameplay assets" in s5 and "section 9" in s5,
   "an unknown event type falls through to section 9 rather than emitting nothing", s5)

# ---- Art Content Overview counts the same elements -------------------------
md = arg._render_markdown(SORT, "Sort", "Draft")
overview = md[md.index("## 3."):md.index("## 4.")]
ok("Goods / shelf icons" in overview,
   "section 3's quantities are derived from the same per-type map, so the two "
   "sections cannot disagree")
ok("| Characters | 0 |" in overview, "and it reports a real zero rather than omitting the row")

# ---- grounding -------------------------------------------------------------
ok(arg.ART_REQ_TEMPLATE == "Event_Art_Requirement_Template.md",
   "the generator is grounded in the production team's template")
# Checking for the STRING would fail on the comments that explain why that file
# is the wrong one - which are worth keeping. What matters is that nothing READS
# it any more.
_src = open(os.path.join(os.path.dirname(os.path.abspath(__file__)),
                         "art_requirement_generator.py"), encoding="utf-8").read()
ok('_read_template(events_root, "Event_Final_Art_Requirement.md")' not in _src,
   "and nothing reads the multi-discipline Final Spec as if it were one")
ok(_src.count("_read_template(events_root, ART_REQ_TEMPLATE)") == 2,
   "both the generate and regenerate paths are grounded in it - regenerating a "
   "variation against a different template than it was drafted from would drift "
   "the document out of shape",
   _src.count("_read_template(events_root, ART_REQ_TEMPLATE)"))

print("")
if fails:
    print("%d TEMPLATE CHECK(S) FAILED" % len(fails))
    sys.exit(1)
print("ARTREQ TEMPLATE OK")
