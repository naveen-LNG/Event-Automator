"""Tests for Recraft custom-style setup.

Everything here is offline. The one thing that cannot be tested offline is the
comfy.org authorisation, which is exactly where the live run stops: the graph
validates and RecraftCreateStyleNode executes, then answers "Unauthorized:
Please login first to use this node". So these tests pin the parts that were
expensive to work out - the autogrow socket naming and the 5 MB budget - so they
survive until the host is signed in.
"""
import os, sys, json, tempfile, shutil

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
HERE = os.path.dirname(os.path.abspath(__file__))
os.chdir(HERE); sys.path.insert(0, HERE)

from PIL import Image
import recraft_style as rs

fails = []
def ok(c, m, d=""):
    print(("OK  " if c else "FAIL ") + m + (("\n     " + d) if (d and not c) else ""))
    if not c: fails.append(m)

# --- socket naming: recovered from the validation error, worth pinning --------
wf = rs._build(["a.png", "b.png", "c.png"], "images.image%d:0-based")
node = wf["1"]["inputs"]
ok(node.get("images.image0") == ["200", 0],
   "autogrow sockets are dotted and 0-based (images.image0)",
   "got keys: %s" % sorted(node))
ok("images.image2" in node and "images.image3" not in node,
   "one socket per reference, no more")
ok(node["style"] == "digital_illustration",
   "base style is digital_illustration, not realistic_image")
ok(rs.SOCKET_PATTERNS[0] == "images.image%d:0-based",
   "the known-good pattern is tried first so a run costs one queue, not five")
ok(wf["2"]["class_type"] == "PreviewAny" and wf["2"]["inputs"]["source"] == ["1", 0],
   "a PreviewAny sink is present - without it /history publishes no style_id")
ok(all(v["class_type"] == "LoadImage" for k, v in wf.items() if k.startswith("2") and k != "2"),
   "each reference gets its own LoadImage")

# 1-based / list fallbacks still build, in case ComfyUI changes the flattening
ok(rs._build(["a.png"], "image%d")["1"]["inputs"].get("image1") == ["200", 0],
   "the 1-based fallback pattern still builds")
ok(rs._build(["a.png", "b.png"], "images-as-list")["1"]["inputs"]["images"]
   == [["200", 0], ["201", 0]],
   "the as-a-list fallback pattern still builds")

# --- the generation graph cites the style and attaches NO images --------------
g = rs.build_workflow("a quiet pottery studio", "STYLE-UUID")
ok(g["1"]["inputs"]["style_id"] == "STYLE-UUID",
   "style_id feeds RecraftStyleV3InfiniteStyleLibrary")
ok(g["2"]["inputs"]["recraft_style"] == ["1", 0],
   "the resulting RECRAFT_V3_STYLE feeds the text-to-image node")
ok(not any(v["class_type"] == "LoadImage" for v in g.values()),
   "NO LoadImage in the generation graph",
   "this is the whole point: no attached image means nothing can bleed")
ok("negative_prompt" not in g["2"]["inputs"],
   "negative_prompt is omitted when empty rather than sent blank")
ok(rs.build_workflow("p", "S", negative_prompt="text")["2"]["inputs"]["negative_prompt"] == "text",
   "negative_prompt is passed through when given")
ok(g["2"]["inputs"]["size"] == "1024x1365",
   "default size is the 3:4 option, matching the Sort floor aspect")

# --- the 5 MB budget: the real floors are 8.27 MB and must be shrunk ---------
tmp = tempfile.mkdtemp()
try:
    src = []
    for i in range(5):
        p = os.path.join(tmp, "f%d.png" % i)
        # noisy image so PNG cannot trivially compress it away
        im = Image.new("RGB", (1190, 1590))
        im.putdata([((x * 7 + y * 13) % 256, (x * 3) % 256, (y * 5) % 256)
                    for y in range(1590) for x in range(1190)])
        im.save(p); src.append(p)
    small = rs._shrink(src, os.path.join(tmp, "out"))
    got = sum(os.path.getsize(q) for q in small)
    ok(len(small) == 5, "all five references are shrunk")
    ok(got <= rs.MAX_TOTAL_BYTES,
       "shrunk set fits Recraft's 5 MB cap (%.2f MB)" % (got / 1048576.0))
    dims = [Image.open(q).size for q in small]
    ok(all(max(d) <= rs.SHRINK_MAX_DIM for d in dims),
       "every shrunk reference has its longest edge capped at %d px" % rs.SHRINK_MAX_DIM,
       "got %s" % dims)
    ok(all(Image.open(q).mode == "RGB" for q in small),
       "alpha is flattened - transparency is not part of a style")
    ok(all(Image.open(q).size == Image.open(s).size for q, s in zip(small, small[1:])),
       "uniform source floors shrink to a uniform size")

    # load() round-trip
    ok(rs.load(tmp) is None, "load() returns None when no style has been created")
    with open(os.path.join(tmp, rs.STYLE_FILE), "w", encoding="utf-8") as f:
        json.dump({"style_id": "abc", "base_style": "digital_illustration"}, f)
    ok((rs.load(tmp) or {}).get("style_id") == "abc", "load() reads a stored style back")
    with open(os.path.join(tmp, rs.STYLE_FILE), "w", encoding="utf-8") as f:
        json.dump({"style_id": ""}, f)
    ok(rs.load(tmp) is None, "load() treats an empty style_id as absent")
finally:
    shutil.rmtree(tmp, ignore_errors=True)

ok(rs.MAX_REFS == 5, "reference count cap matches the node's documented limit of 5")

print("")
if fails:
    print("%d CHECK(S) FAILED" % len(fails)); sys.exit(1)
print("ALL RECRAFT STYLE TESTS PASSED")
