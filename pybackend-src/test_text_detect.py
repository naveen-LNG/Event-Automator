"""Tests for the rendered-text detector.

Text is the one defect that makes a generated asset unusable outright, and prose
bans do not reliably hold - a labelled list came back as a caption strip, a
numbered list as circled callout markers plus a key. Both got past visual review
during the Sort work, which is why this check exists.

Validated against real output from those runs, not just synthetic images.
"""
import os, sys, tempfile, shutil

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
HERE = os.path.dirname(os.path.abspath(__file__))
os.chdir(HERE); sys.path.insert(0, HERE)

from PIL import Image, ImageDraw
import text_detect as td

fails = []
def ok(c, m, d=""):
    print(("OK  " if c else "FAIL ") + m + (("\n     " + d) if (d and not c) else ""))
    if not c: fails.append(m)

tmp = tempfile.mkdtemp()
try:
    W, H = 600, 800

    def blank(bg=(235, 225, 200)):
        return Image.new("RGBA", (W, H), bg + (255,))

    # 1. a synthetic caption strip is caught -----------------------------------
    im = blank(); d = ImageDraw.Draw(im)
    d.polygon([(300, 300), (520, 420), (300, 540), (80, 420)], fill=(200, 150, 120, 255))
    d.text((60, 720), "Zone A Melting Bay Pouring Line Curing", fill=(30, 30, 30, 255))
    p = os.path.join(tmp, "caption.png"); im.save(p)
    r = td.detect_text(p)
    ok(r["has_text"], "a caption strip under the artwork is detected", r["detail"])
    if r["lines"]:
        ok(r["lines"][0]["y"] > 0.8, "and located near the bottom of the frame",
           str(r["lines"][0]))

    # 2. artwork with no text passes -------------------------------------------
    im = blank(); d = ImageDraw.Draw(im)
    d.polygon([(300, 300), (520, 420), (300, 540), (80, 420)], fill=(200, 150, 120, 255))
    for i, x in enumerate(range(120, 480, 60)):        # a row of similar objects
        d.ellipse([x, 380, x + 40, 430], fill=(120 + i * 12, 90, 160, 255))
    p2 = os.path.join(tmp, "clean.png"); im.save(p2)
    ok(not td.detect_text(p2)["has_text"],
       "a row of similar objects is NOT mistaken for text",
       "this is the false positive that forced the threshold up from 5 to 10")

    # 3. an empty plate passes --------------------------------------------------
    im = blank(); d = ImageDraw.Draw(im)
    d.polygon([(300, 300), (520, 420), (300, 540), (80, 420)], fill=(200, 150, 120, 255))
    p3 = os.path.join(tmp, "empty.png"); im.save(p3)
    ok(not td.detect_text(p3)["has_text"], "a bare plate produces no false positive")
finally:
    shutil.rmtree(tmp, ignore_errors=True)

# ---- real output from this session, with known ground truth ------------------
B = r"C:\Users\NAVEEN~1\AppData\Local\Temp\claude\C--SkipBo\b4d775b4-e255-4b0e-9923-1aa47c2b64cf\scratchpad"
S = r"C:\SkipBo\tools\EventAutomationDesktop\ideationData\StyleReference\images\Sort"
REAL = [(os.path.join(B, "sort_v4", "floor5_candle_terrace.png"), True, "generated, caption strip")]
for n in ("floor1_mixing_room", "floor2_oven_hall", "floor3_chocolate_room",
          "floor4_decorating_studio", "floor5_cafe_terrace"):
    REAL.append((os.path.join(B, "bakery", n + ".png"), False, "generated, clean: " + n))
if os.path.isdir(S):
    for n in sorted(os.listdir(S)):
        REAL.append((os.path.join(S, n), False, "shipped: " + n.replace("sort2609_zone_", "")))

present = [(p, e, l) for p, e, l in REAL if os.path.exists(p)]
if present:
    print("\n--- real images (ground truth from this session) ---")
    wrong = []
    for p, expect, label in present:
        got = td.detect_text(p)["has_text"]
        print("  %-42s expect %-5s got %-5s %s" % (label[:42], expect, got, "" if got == expect else "<-- MISS"))
        if got != expect:
            wrong.append(label)
    ok(not wrong, "detector agrees with ground truth on all %d real images" % len(present),
       "disagreed on: " + ", ".join(wrong))

# ---- signage checking: the Merge case --------------------------------------
# On Sort any text is a defect, so "is there text" is the whole question. On
# Merge signage is REQUIRED and only the wrong words are a defect - a design
# name written on a building reads structurally the same as an allowed sign.
# Telling them apart means reading them, so this half depends on OCR.
print("\n--- signage / OCR ---")
print("  ocr_available: %s" % td.ocr_available())

ok(td._tokens("DEPT. STORE") == ["DEPT", "STORE"], "punctuation is dropped when tokenising")
ok(td._tokens("MUSIC BOX PAVILION") == ["MUSIC", "BOX", "PAVILION"], "multi-word names tokenise")
ok(td._tokens("") == [] and td._tokens(None) == [], "empty input tokenises to nothing")

# near-miss handling: OCR is not perfect on stylised signage, and a
# correctly painted MOSAIC came back as 'MOSAIO' and hard-failed a batch.
# One character out is ambiguous between a misread and genuinely misspelled
# art, so it warns instead of failing. These are pure-function checks and do
# not depend on what the engine happens to read.
ok(td._edit_distance("MOSAIO", "MOSAIC") == 1, "one substitution is distance 1")
ok(td._edit_distance("TEA", "PAVILION") > td.NEAR_MISS_MAX,
   "plainly different words exceed the cap")
ok(td._near_match("MOSAIO", {"MOSAIC", "TILE", "WORKS"}) == "MOSAIC",
   "a one-character misread resolves to the word it meant")
ok(td._near_match("PAVILION", {"TEA", "LAMPS", "POST"}) is None,
   "a design name is NOT excused as a near miss")
ok(td._near_match("SEA", {"TEA"}) is None,
   "words under 4 characters are never near-matched",
   "SEA and TEA are one apart but both are real, distinct signs")

# A damaged FRAGMENT of a sign is the other half of the same problem. Lettering
# embossed in glass on glass, correctly reading "GLASS HOUSE", came back as
# 'DUSE' and 'OSE'. Neither is a substring of GLASSHOUSE and neither is within
# one edit of GLASS or HOUSE - but both are one edit from the substring 'OUSE'.
ok(td._near_fragment("DUSE", ["GLASSHOUSE"]) == "GLASSHOUSE",
   "a fragment with one wrong character resolves to the sign it came from")
ok(td._near_fragment("OSE", ["GLASSHOUSE"]) == "GLASSHOUSE",
   "a short damaged fragment resolves too")
ok(td._near_fragment("USE", ["GLASSHOUSE"]) == "GLASSHOUSE",
   "a clean fragment resolves")
ok(td._near_fragment("GROUT", ["TILEWORKS", "MOSAIC"]) is None,
   "a genuinely different word is not excused as a fragment",
   "this is the design-name case the whole check exists for")
ok(td._near_fragment("KITCHEN", ["TILEWORKS", "MOSAIC"]) is None,
   "nor is the second half of a design name")
ok(td._near_fragment("PAVILION", ["TEA", "LAMPS", "POST"]) is None,
   "nor a long design name against short signs")
ok(td._near_fragment("OS", ["GLASSHOUSE"]) is None,
   "fragments under 3 characters are never matched - too little to go on")

# The no-OCR branch is FORCED rather than skipped. The backend is vendored now,
# so this path would otherwise never run and would quietly rot - and it is the
# path that protects against a broken or missing install, which is exactly when
# nobody is watching. _OCR is the one-slot engine cache; [None] means "probed,
# absent".
_saved_ocr = list(td._OCR)
td._OCR[:] = [None]
try:
    ok(td.ocr_available() is False, "the no-OCR path can be forced for testing")
    # The contract that matters: "could not check" must be distinguishable from
    # "checked and clean", so a caller never reads a missing backend as a pass.
    r = td.check_signage("nonexistent.png", ["TEA"])
    ok(r["ok"] is None, "without OCR, check_signage returns ok=None, not True",
       "a missing backend must never read as a pass")
    ok(r["engine"] == "none", "the engine is reported as none")
    ok("VENDORING" in td.__doc__ or "vendor" in r["detail"].lower(),
       "the detail points at how to install the backend")
    tmp3 = tempfile.mkdtemp()
    try:
        im3 = Image.new("RGBA", (600, 800), (235, 225, 200, 255))
        ImageDraw.Draw(im3).polygon([(300, 300), (520, 420), (300, 540), (80, 420)],
                                    fill=(200, 150, 120, 255))
        q3 = os.path.join(tmp3, "plain.png"); im3.save(q3)
        ok(td.detect_text(q3)["has_text"] is False,
           "detect_text still falls back to the structural check")
    finally:
        shutil.rmtree(tmp3, ignore_errors=True)
finally:
    td._OCR[:] = _saved_ocr

if td.ocr_available():
    tmp2 = tempfile.mkdtemp()
    try:
        from PIL import ImageFont
        img = Image.new("RGB", (900, 500), (150, 170, 140))
        d = ImageDraw.Draw(img)
        try:
            f = ImageFont.truetype(r"C:\Windows\Fonts\arialbd.ttf", 44)
        except Exception:
            f = ImageFont.load_default()
        d.rectangle([60, 200, 520, 280], fill=(210, 190, 120))
        d.text((80, 212), "MUSIC BOX PAVILION", font=f, fill=(35, 25, 10))
        d.rectangle([600, 200, 760, 280], fill=(210, 190, 120))
        d.text((620, 212), "TEA", font=f, fill=(35, 25, 10))
        q = os.path.join(tmp2, "signs.png"); img.save(q)

        words = td.read_text(q)
        joined = " ".join(w["text"].upper() for w in words or [])
        ok(words is not None and len(words) >= 2,
           "read_text returns the rendered words", str(words))
        ok("TEA" in joined, "an allowed short sign is read", joined)
        ok("PAVILION" in joined, "a long design name is read", joined)

        r = td.check_signage(q, ["TEA"])
        ok(r["ok"] is False, "a design name is flagged when only TEA is authorised", str(r))
        ok(any("PAVILION" in u.upper() for u in r["unexpected"]),
           "the design name is named in the report", str(r["unexpected"]))

        r2 = td.check_signage(q, ["TEA", "MUSIC BOX PAVILION"])
        ok(r2["ok"] is True, "the same image passes once that wording is authorised", str(r2))

        # OCR is inconsistent about spacing on stylised signs - the same label
        # read back as "MUSIC BOX PAVILION" on one render and "MUSICBOXPAVILION"
        # on another. Both must resolve against the same authorisation.
        ok(td.check_signage(q, ["TEA", "MUSIC", "BOX", "PAVILION"])["ok"] is True,
           "a squashed read still segments into its authorised words")
        # and one word of a longer authorised sign is a fragment, not a defect -
        # "PAPER BAKERY" on two lines reads back as "PAPER" and "BAKERY"
        ok(td.check_signage(q, ["TEA", "MUSIC BOX PAVILION EXTRA"])["ok"] is True,
           "a partial read of a longer authorised sign passes")
        ok(td.check_signage(q, ["TEA", "LAMPS", "POST"])["ok"] is False,
           "an unauthorised design name is still caught after all that leniency")

    finally:
        shutil.rmtree(tmp2, ignore_errors=True)

    # thorough (tiled) reading: a whole-image pass misses small, low-contrast
    # signage. A generated scene plainly carrying "GROUT KITCHEN" returned
    # nothing for it even at zero confidence, and the gate reported it clean.
    tmp4 = tempfile.mkdtemp()
    try:
        from PIL import ImageFont
        big = Image.new("RGB", (1456, 1092), (150, 165, 140))
        db = ImageDraw.Draw(big)
        for gx in range(0, 1456, 40):          # busy background, as a real scene is
            db.line([(gx, 0), (gx, 1092)], fill=(140, 156, 132), width=2)
        try:
            fsmall = ImageFont.truetype(r"C:\Windows\Fonts\arialbd.ttf", 15)
        except Exception:
            fsmall = ImageFont.load_default()
        db.rectangle([1120, 940, 1300, 972], fill=(214, 196, 150))
        db.text((1130, 948), "GROUT KITCHEN", font=fsmall, fill=(40, 32, 16))
        q4 = os.path.join(tmp4, "small.png"); big.save(q4)

        plain = td.read_text(q4) or []
        deep = td.read_text(q4, thorough=True) or []
        ok(len(deep) >= len(plain),
           "thorough reading never returns less than the plain pass (%d vs %d)"
           % (len(deep), len(plain)))
        ok(any("GROUT" in w["text"].upper() or "KITCHEN" in w["text"].upper() for w in deep),
           "thorough reading finds small signage",
           "plain=%s deep=%s" % ([w["text"] for w in plain], [w["text"] for w in deep]))
        ok(td.check_signage(q4, ["TILE WORKS", "MOSAIC"])["ok"] is False,
           "and check_signage therefore catches a small design name",
           "check_signage must read thoroughly by default - it is a gate")
    finally:
        shutil.rmtree(tmp4, ignore_errors=True)

    SEEDED = (r"C:\SkipBo\tools\EventAutomationDesktop\ideationData"
              r"\StyleReference\images\Merge\_excluded\merge2606_scene2_brick_town.png")
    if os.path.exists(SEEDED):
        w = td.read_text(SEEDED) or []
        blob = " ".join(x["text"].upper() for x in w)
        ok(len(w) >= 3, "the seeded Merge reference reads as several signs (%d)" % len(w))
        ok("MARKET" in blob or "BAKERY" in blob,
           "a known sign from the shipped concept art is read back", blob[:80])

print("")
if fails:
    print("%d CHECK(S) FAILED" % len(fails)); sys.exit(1)
print("ALL TEXT DETECT TESTS PASSED")
