"""Tests for the ComfyUI / Nano Banana reference-image path.

The HTTP layer is stubbed so the whole thing is verifiable while the studio's
ComfyUI worker is wedged. One live check at the end reports the server's real
state without asserting on it.

Run with the bundled interpreter:
    pybackend_dist\\python\\python.exe test_comfy_images.py
"""
import io, json, os, shutil, sys, tempfile

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
HERE = os.path.dirname(os.path.abspath(__file__))
os.chdir(HERE)
sys.path.insert(0, HERE)

import art_reference as ar
import comfy_client as cc

DATA = r"C:\SkipBo\tools\EventAutomationDesktop\ideationData\Events"
fails = []


def ok(cond, msg, detail=""):
    if cond:
        print("OK  " + msg)
    else:
        print("FAIL " + msg + ("\n     " + detail if detail else ""))
        fails.append(msg)


# 1. The approved workflow is reproduced exactly ------------------------------
UI = json.load(io.open(
    r"C:\SkipBo\tools\EventAutomationDesktop\ideationData\StyleReference\nano_banana_workflow.ui.json",
    encoding="utf-8"))
nb = next(n for n in UI["nodes"] if n["type"] == "NanoBananaAIO")
w = nb["widgets_values"]           # model, prompt, image_count, use_search, aspect, size, temp
wf = cc.build_workflow("P", ref_names=["ref.png"])
inp = wf["1"]["inputs"]
ok(wf["1"]["class_type"] == "NanoBananaAIO", "uses NanoBananaAIO, the node the approved workflow uses")
ok(inp["model_name"] == w[0], "model_name matches the approved workflow (%s)" % w[0])
ok(inp["image_count"] == w[2] and inp["use_search"] == w[3],
   "image_count/use_search match the approved workflow")
ok(inp["aspect_ratio"] == w[4] and inp["image_size"] == w[5] and inp["temperature"] == w[6],
   "aspect/size/temperature match the approved workflow (%s, %s, %s)" % (w[4], w[5], w[6]))
scale = next(n for n in UI["nodes"] if n["type"] == "ImageScale")["widgets_values"]
ok([wf["2"]["inputs"][k] for k in ("upscale_method", "width", "height", "crop")] == scale,
   "ImageScale matches the approved workflow (%s %dx%d)" % (scale[0], scale[1], scale[2]))
ok(wf["101"]["class_type"] == "LoadImage" and inp["image_1"] == ["101", 0],
   "reference images are attached via LoadImage -> image_1 (section 8: attach, don't describe)")

# 2. Option values are ones the live server actually accepts ------------------
OI = json.load(io.open(
    r"C:\Users\NAVEEN~1\AppData\Local\Temp\claude\C--SkipBo\b4d775b4-e255-4b0e-9923-1aa47c2b64cf\scratchpad\object_info.json",
    encoding="utf-8")) if os.path.exists(
    r"C:\Users\NAVEEN~1\AppData\Local\Temp\claude\C--SkipBo\b4d775b4-e255-4b0e-9923-1aa47c2b64cf\scratchpad\object_info.json") else None
if OI:
    spec = OI["NanoBananaAIO"]["input"]
    opts = lambda sec, k: (spec[sec][k][1] or {}).get("options") or spec[sec][k][0]
    ok(inp["model_name"] in opts("required", "model_name"), "model_name is a valid server option")
    ok(inp["aspect_ratio"] in opts("optional", "aspect_ratio"), "aspect_ratio is a valid server option")
    ok(inp["image_size"] in opts("optional", "image_size"), "image_size is a valid server option")
else:
    print("--  skipped server-option cross-check (no cached object_info.json)")

# 3. Prompt assembly order (ROOM_RENDER_STYLE.md section 8) -------------------
calls = []
def fake_distil(system, user, schema):
    calls.append(len(user))
    return {"feel": ["Warm nutty interior"], "palette": ["butter-cream + teal"],
            "silhouette": ["hazelnut-shaped cabinet"]}

zone = {"zone_number": 2, "zone_theme": "Roastery", "palette_or_mood": "teal + cream",
        "construction_steps": ["Drum", "Rack"], "reference_notes": "warm"}
p = ar.build_scene_prompt(DATA, "Sim", zone, doc_title="Sim2610",
                          art_direction={"lean_into": ["candy"], "avoid": ["olive"]},
                          distil_fn=fake_distil)
idx = [p.find(m) for m in ("APPROVED RENDER SPEC", "NAMED FAILURES",
                           "HOUSE VISUAL LANGUAGE", "THIS SCENE")]
ok(all(i >= 0 for i in idx) and idx == sorted(idx),
   "prompt order: spec -> named failures -> house language -> scene", str(idx))
ok(idx[0] < 10, "the render spec is at the very top (buried rules get ignored)")
for phrase in ("no soft-serve", "must depart from them", "no valve wheels" if "valve wheels" in p else "valve wheels"):
    ok(phrase.lower() in p.lower(), "names the specific failure: %r" % phrase)

# 4. Distillation reads the Chinese docs, and caches --------------------------
tmp = tempfile.mkdtemp()
try:
    fake_root = os.path.join(tmp, "Events")
    shutil.copytree(os.path.join(DATA, "Sim"), os.path.join(fake_root, "Sim"))
    calls.clear()
    v = ar.extract_visual_language(fake_root, "Sim", distil_fn=fake_distil)
    ok(len(calls) == 1 and calls[0] > 2000,
       "Sim's Chinese PDFs reach the distiller (%d chars)" % (calls[0] if calls else 0))
    ok(v["feel"] and v["sources"], "distilled language comes back with sources")
    v2 = ar.extract_visual_language(fake_root, "Sim", distil_fn=fake_distil)
    ok(len(calls) == 1 and v2.get("cached"), "second call is served from cache (no extra model call)")
    # touching the corpus must invalidate
    os.utime(os.path.join(fake_root, "Sim", "Art Requirements",
                          os.listdir(os.path.join(fake_root, "Sim", "Art Requirements"))[0]), None)
    ar.extract_visual_language(fake_root, "Sim", distil_fn=fake_distil)
    ok(len(calls) == 2, "editing a PDF invalidates the cache and re-distils")
    # no provider -> must degrade, never raise
    v3 = ar.extract_visual_language(fake_root, "Blocks", distil_fn=None)
    ok(isinstance(v3, dict) and v3["feel"] == [], "no provider degrades to empty, does not raise")
finally:
    shutil.rmtree(tmp, ignore_errors=True)

# 5. Blank-output guard (section 8: ~5 KB means a blank) ----------------------
runs = []
def fake_run(base, wf_, dest, timeout):
    runs.append(1)
    n = 4_800 if len(runs) < 3 else 250_000      # two blanks, then a real render
    return {"ok": True, "path": dest, "bytes": n, "prompt_id": "x"}

real_run, cc._run_once = cc._run_once, fake_run
try:
    r = cc.generate("http://x", "p", os.path.join(tempfile.gettempdir(), "z.png"))
    ok(r["ok"] and r["bytes"] == 250_000 and len(runs) == 3,
       "a ~5 KB blank is retried and the real render is kept (%d attempts)" % len(runs))
    runs.clear()
    cc._run_once = lambda *a, **k: (runs.append(1), {"ok": True, "path": "z", "bytes": 4800})[1]
    r = cc.generate("http://x", "p", "z.png")
    ok((not r["ok"]) and "blank" in r["error"].lower(),
       "all-blank runs fail loudly instead of returning a blank image")
finally:
    cc._run_once = real_run

# 6. Failures never raise -----------------------------------------------------
r = cc.generate("http://127.0.0.1:9/nope", "p", os.path.join(tempfile.gettempdir(), "q.png"))
ok(r["ok"] is False and r["error"], "an unreachable host returns an error, never raises")

# 7. Live server state (reported, not asserted) -------------------------------
live = cc.check("https://skipbo.comfyui.liquidnitro.games", timeout=12)
print("\n--  live server: state=%s  reachable=%s  node=%s  queue_ok=%s"
      % (live["state"], live["reachable"], live["node_present"], live["queue_ok"]))
print("--  %s" % live["detail"])

# ---- the aspect ratio actually reaches the output size -----------------------
# It did not. The workflow ends in an ImageScale whose size was hardcoded to
# 1115x994, so a 3:4 request was sent to the model, came back portrait, and was
# then resampled to near-square. Every render the tool has produced was
# landscape regardless of what was asked for.
import comfy_client as cc

def wh(aspect, **kw):
    n = cc.build_workflow("p", aspect_ratio=aspect, **kw)["2"]["inputs"]
    return n["width"], n["height"]

ok(cc.build_workflow("p", aspect_ratio="3:4")["1"]["inputs"]["aspect_ratio"] == "3:4",
   "the aspect ratio still reaches the render node")
w, h = wh("3:4")
ok(h > w, "a 3:4 request produces a PORTRAIT output size (the bug this fixes)",
   "got %dx%d" % (w, h))
ok(abs((w / float(h)) - 0.75) < 0.01, "and the ratio is 3:4 to within a percent",
   "got %.3f from %dx%d" % (w / float(h), w, h))
w2, h2 = wh("16:9")
ok(w2 > h2 and abs((w2 / float(h2)) - (16 / 9.0)) < 0.01,
   "16:9 produces a landscape output at the right ratio", "got %dx%d" % (w2, h2))
ok(abs(w * h - cc.BASE_PIXELS) / float(cc.BASE_PIXELS) < 0.02
   and abs(w2 * h2 - cc.BASE_PIXELS) / float(cc.BASE_PIXELS) < 0.02,
   "every aspect keeps the exported pixel budget, so cost and detail hold steady",
   "%d vs %d vs %d" % (w * h, w2 * h2, cc.BASE_PIXELS))
ok(wh("Auto") == cc.BASE_SCALE, "Auto keeps the size the approved workflow exported at",
   str(wh("Auto")))
ok(wh("nonsense") == cc.BASE_SCALE, "and an unparseable ratio falls back rather than crashing")
ok(wh("1:1", scale_w=800, scale_h=600) == (800, 600),
   "an explicit size still wins over the ratio")

# ---- the named-failure block is routed by event type -------------------------
# Same bug class as the render spec and the reference folder before it: one
# type's rules were injected into every type. For Dessert the default block is
# not merely irrelevant, it is wrong - it tells the model to depart from
# references that are now the shipped art, to stay matte when the shipped art is
# glossy, and not to draw producer fittings on a type that has no producers.
dess = ar.named_failures("Dessert")
dflt = ar.named_failures("Merge")
ok("SHIPPED GAME ART" in dess, "Dessert is told its references ARE the target")
ok("must depart from them" not in dess,
   "and is NOT told to depart from them (the bug this fixes)")
ok("Matte only" not in dess, "nor to stay matte, which the shipped art is not")
ok("producer" not in dess.lower(), "nor about producer fittings it does not have")
ok("CONTINUOUS PARK" in dess, "and is told it is one continuous park, not a floating plate")
ok("must depart from them" in dflt and "Matte only" in dflt,
   "while every other type keeps the block it was written for")
for shared in ("no soft-serve", "No text, letters or numbers", "Nothing mirrored"):
    ok(shared.lower() in dess.lower() and shared.lower() in dflt.lower(),
       "the genuinely shared ban survives in both: %r" % shared)
ok(ar.named_failures("") == dflt, "an unknown type gets the default block")
pd = ar.build_scene_prompt(DATA, "Dessert", zone, doc_title="Dessert2611",
                           style_spec="# Dessert zone art", distil_fn=fake_distil)
ok("SHIPPED GAME ART" in pd and "must depart from them" not in pd,
   "and build_scene_prompt injects the type's own block, not the default")

print("")
if fails:
    print("%d CHECK(S) FAILED" % len(fails))
    sys.exit(1)
print("ALL COMFY IMAGE TESTS PASSED")
