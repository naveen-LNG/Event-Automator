"""Tests for the within-frame duplicate-building check.

SCOPE, STATED UP FRONT: these tests cover the contract, the message the model is
sent, and the parsing of what comes back. They do NOT measure whether the model
actually spots a duplicated building, because neither provider was reachable
when this was written - the app's anthropicApiKey is empty and its configured
Ollama host (10.16.80.22) did not answer. So the detector's real accuracy is
UNVALIDATED. What would validate it: run check() over a scene with a known
repeated building and one without, with a vision model configured.

That gap is why the contract tested hardest here is that "did not check" can
never be mistaken for "checked and clean" - ran False leaves duplicate_found
None, never False.
"""
import os, sys, json, tempfile, shutil

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
HERE = os.path.dirname(os.path.abspath(__file__))
os.chdir(HERE); sys.path.insert(0, HERE)

from PIL import Image
import duplicate_check as dc
import artreq_images as ai

fails = []
def ok(c, m, d=""):
    print(("OK  " if c else "FAIL ") + m + (("\n     " + d) if (d and not c) else ""))
    if not c: fails.append(m)

# ---- the "did not check" contract -------------------------------------------
r = dc.check("whatever.png")
ok(r["ran"] is False, "no provider configured means the check does not run")
ok(r["duplicate_found"] is None,
   "and duplicate_found is None, never False",
   "False would read as 'checked and clean' and hide the defect")
r = dc.check("whatever.png", provider="ollama")
ok(r["ran"] is False and "Ollama" in r["detail"],
   "ollama with no URL is refused with a provider-specific message", r["detail"])
r = dc.check("missing.png", api_key="sk-test")
ok(r["ran"] is False and "not found" in r["detail"],
   "a missing image is reported, not silently passed", r["detail"])

# ---- parsing the model's answer ---------------------------------------------
dup = dc._parse({"duplicate_found": True, "building_count": 9,
                 "repeated": [{"what": "terracotta kiln dome",
                               "where": "upper left and bottom centre"}],
                 "confidence": "high", "notes": ""}, expected_count=8)
ok(dup["duplicate_found"] is True, "a reported duplicate is surfaced")
ok("terracotta kiln dome" in dup["detail"] and "upper left" in dup["detail"],
   "the finding names what repeated and where", dup["detail"])
ok("asked for 8" in dup["detail"], "and flags the count against the brief", dup["detail"])

clean = dc._parse({"duplicate_found": False, "building_count": 8, "repeated": [],
                   "confidence": "high", "notes": ""}, expected_count=8)
ok(clean["duplicate_found"] is False and clean["ran"] is True,
   "a clean answer is clean, and marked as actually checked")
ok("asked for" not in clean["detail"],
   "a matching count is not flagged", clean["detail"])

miscount = dc._parse({"duplicate_found": False, "building_count": 6, "repeated": [],
                      "confidence": "medium", "notes": "two structures are ambiguous"},
                     expected_count=8)
ok("Counted 6" in miscount["detail"] and "asked for 8" in miscount["detail"],
   "a wrong count is reported even when no duplicate was found", miscount["detail"])
ok("ambiguous" in miscount["detail"], "the model's own note is carried through")
ok(dc._parse({}, None)["duplicate_found"] is False,
   "an empty response degrades to 'no duplicate found' rather than crashing")

# ---- the message actually sent ----------------------------------------------
tmp = tempfile.mkdtemp()
try:
    p = os.path.join(tmp, "scene.png")
    Image.new("RGB", (1456, 1092), (120, 150, 120)).save(p)
    b64 = dc._encode(p)
    ok(len(b64) > 100, "the image encodes to base64")
    msg = dc.build_user_message(b64, expected_count=8)
    ok(msg[0]["type"] == "image" and msg[0]["source"]["media_type"] == "image/jpeg",
       "the image rides as base64 JPEG, not a raw PNG blob")
    ok("EXACTLY 8" in msg[1]["text"],
       "the expected building count is put to the model", msg[1]["text"])
    ok("EXACTLY" not in dc.build_user_message(b64)[1]["text"],
       "and omitted when the brief did not state one")
    im = Image.open(p)
    small = dc._encode(p)
    ok(len(small) < 2_000_000, "a full-size scene stays well inside the upload limit")
finally:
    shutil.rmtree(tmp, ignore_errors=True)

# the prompt must forbid reporting the things that are SUPPOSED to repeat
for word in ("trees", "benches", "lamp posts", "planters"):
    ok(word in dc.SYSTEM, "the brief tells the model not to report repeated %s" % word)
ok("EXACTLY ONCE" in dc.SYSTEM, "and states the rule being enforced")
ok(json.dumps(dc.SCHEMA) and dc.SCHEMA["properties"]["duplicate_found"]["type"] == "boolean",
   "the output schema is well-formed JSON Schema")

# ---- how the reviewer picks up the provider ---------------------------------
ok(ai._vision_cfg({}) is None, "no provider configured yields no vision config")
ok(ai._vision_cfg({"provider": "ollama"}) is None, "ollama without a URL yields none")
v = ai._vision_cfg({"provider": "ollama", "ollama_base_url": "http://x:11434",
                    "ollama_model": "qwen3:32b"})
ok(v and v["provider"] == "ollama" and v["ollama_base_url"] == "http://x:11434",
   "an Ollama host is passed through", str(v))
v2 = ai._vision_cfg({"provider": "ollama", "ollama_base_url": "http://x:11434",
                     "ollama_model": "qwen3:32b", "ollama_vision_model": "qwen2.5vl:32b"})
ok(v2["ollama_model"] == "qwen2.5vl:32b",
   "a dedicated vision model overrides the text model",
   "qwen3:32b is text-only and would ignore the image")
ok(ai._vision_cfg({"api_key": "sk-x"})["provider"] == "anthropic",
   "an Anthropic key yields the anthropic provider")

# ---- the building count is read out of the prompt ---------------------------
ok(ai.count_from_prompt("*** EXACTLY NINE BUILDINGS, WELL APART ***") == 9,
   "a spelled-out count is parsed")
ok(ai.count_from_prompt("contains EXACTLY EIGHT buildable objects") == 8,
   "'buildable objects' is recognised as the same rule")
ok(ai.count_from_prompt("- EXACTLY 12 BUILDINGS.") == 12, "a numeric count is parsed")
ok(ai.count_from_prompt("many buildings, no count given") is None,
   "a brief with no stated count yields None")

# ---- the reviewer without a vision provider ---------------------------------
tmp2 = tempfile.mkdtemp()
try:
    q = os.path.join(tmp2, "a.png")
    im = Image.new("RGB", (400, 300), (200, 80, 60))
    for y in range(150, 300):
        for x in range(400):
            im.putpixel((x, y), (40, 90, 180))
    im.save(q)
    r_off = ai.check_generated([q])
    ok(not any("duplicate check" in f["message"] for f in r_off["findings"]),
       "with no vision config the duplicate check is silent, not noisy",
       str(r_off["findings"]))
    r_on = ai.check_generated([q], vision={"provider": "anthropic", "api_key": ""})
    ok(any(f["level"] == "warn" and "did not run" in f["message"]
           for f in r_on["findings"]),
       "but a configured-yet-unusable provider warns that it did not check",
       str(r_on["findings"]))
finally:
    shutil.rmtree(tmp2, ignore_errors=True)

print("")
if fails:
    print("%d CHECK(S) FAILED" % len(fails)); sys.exit(1)
print("ALL DUPLICATE CHECK TESTS PASSED")
