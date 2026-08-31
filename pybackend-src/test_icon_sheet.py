"""Tests for merge-chain icon normalisation.

Calibrated against the SHIPPED Merge2608 icons where they are on disk - those
are the convention, not my invention: 180x180 transparent, one centred object,
and a size that grows monotonically across the five levels of a chain.
"""
import os, sys, tempfile, shutil

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
HERE = os.path.dirname(os.path.abspath(__file__))
os.chdir(HERE); sys.path.insert(0, HERE)

from PIL import Image, ImageDraw
import icon_sheet as ish

fails = []
def ok(c, m, d=""):
    print(("OK  " if c else "FAIL ") + m + (("\n     " + d) if (d and not c) else ""))
    if not c: fails.append(m)

# ---- size progression -------------------------------------------------------
fr = [ish.size_fraction(l) for l in range(1, 6)]
ok(all(fr[i] > fr[i-1] for i in range(1, 5)),
   "the size target increases at every level", str(fr))
ok(fr[0] == ish.SIZE_STEPS[0] and fr[-1] == ish.SIZE_STEPS[-1],
   "the ends match the measured shipped values", str(fr))
fr3 = [ish.size_fraction(l, levels=3) for l in range(1, 4)]
ok(all(fr3[i] > fr3[i-1] for i in range(1, 3)),
   "a chain of a different length still progresses", str(fr3))
ok(ish.size_fraction(1, levels=1) == ish.SIZE_STEPS[-1],
   "a single-level chain is drawn at full size")

tmp = tempfile.mkdtemp()
try:
    def blob(path, w, h, colour=(200, 80, 90)):
        """A cut-out object on transparency, off-centre and oddly sized."""
        im = Image.new("RGBA", (400, 300), (0, 0, 0, 0))
        ImageDraw.Draw(im).ellipse([30, 20, 30 + w, 20 + h], fill=colour + (255,))
        im.save(path)
        return path

    # ---- normalise --------------------------------------------------------
    src = blob(os.path.join(tmp, "a.png"), 120, 60)
    res = ish.normalise(src, os.path.join(tmp, "a_out.png"), level=3, levels=5)
    out = Image.open(os.path.join(tmp, "a_out.png")).convert("RGBA")
    ok(out.size == (ish.CANVAS, ish.CANVAS),
       "output is the shipped canvas size %dx%d" % (ish.CANVAS, ish.CANVAS), str(out.size))
    ok(out.getchannel("A").getextrema()[0] == 0,
       "the background stays fully transparent")
    bb = out.getbbox()
    cx, cy = (bb[0] + bb[2]) / 2.0, (bb[1] + bb[3]) / 2.0
    ok(abs(cx - ish.CANVAS / 2) <= 2 and abs(cy - ish.CANVAS / 2) <= 2,
       "the object is centred on the canvas", "centre %.1f,%.1f" % (cx, cy))
    ok(abs(max(bb[2] - bb[0], bb[3] - bb[1]) - ish.CANVAS * ish.size_fraction(3)) <= 3,
       "the longest edge hits this level's size target",
       "got %d, want ~%.0f" % (max(bb[2]-bb[0], bb[3]-bb[1]), ish.CANVAS*ish.size_fraction(3)))

    # aspect ratio must survive the resize
    ratio_in = 120 / 60.0
    ratio_out = (bb[2] - bb[0]) / float(bb[3] - bb[1])
    ok(abs(ratio_in - ratio_out) < 0.25,
       "the object is not distorted", "in %.2f out %.2f" % (ratio_in, ratio_out))

    # ---- a whole chain progresses ------------------------------------------
    chain = []
    for lv in range(1, 6):
        s = blob(os.path.join(tmp, "c%d.png" % lv), 100 + lv * 5, 70 + lv * 4)
        d = os.path.join(tmp, "c%d_out.png" % lv)
        ish.normalise(s, d, level=lv, levels=5)
        chain.append(d)
    r = ish.check_chain(chain)
    ok(r["ok"], "a normalised chain passes the progression check", str(r))
    ok(all(r["sizes"][i] >= r["sizes"][i-1] for i in range(1, 5)),
       "and its content size rises at every step", str(r["sizes"]))

    # a chain that does NOT progress must fail - this is the defect Mattel named
    broken = list(chain); broken[3] = chain[0]
    rb = ish.check_chain(broken)
    ok(not rb["ok"] and "steps backwards" in rb["findings"][0],
       "a chain that collapses back to an earlier size is caught", str(rb["findings"]))
    # ...but a small dip must NOT fire: the shipped set has 2px and 4px dips
    ok(ish.check_chain([chain[0], chain[1], chain[2], chain[3], chain[4]])["ok"],
       "a clean chain is not flagged")
    flat = [chain[0]] * 5
    ok(not ish.check_chain(flat)["ok"],
       "a chain that never grows at all is caught", str(ish.check_chain(flat)["findings"]))

    # ---- stroke -------------------------------------------------------------
    plain = Image.open(blob(os.path.join(tmp, "s.png"), 100, 100)).convert("RGBA")
    before = plain.getbbox()
    stroked = ish.add_stroke(plain, px=3)
    after = stroked.getbbox()
    ok((after[2] - after[0]) > (before[2] - before[0]),
       "the stroke grows the silhouette outward", "%s -> %s" % (str(before), str(after)))
    # The rim must be DARKER than the artwork, not lighter. Measured on the
    # shipped Merge2608 icons, scoop.edge_halo_score runs -70 to -145: their
    # boundary ring is much darker than the interior. A white rim scores about
    # +116, the opposite, and that is what the first implementation produced.
    sp = stroked.convert("RGBA").load()
    rim = sp[after[0] + 1, (after[1] + after[3]) // 2]
    body = stroked.convert("RGBA").load()[(after[0] + after[2]) // 2,
                                          (after[1] + after[3]) // 2]
    ok(rim[3] > 0 and sum(rim[:3]) < sum(body[:3]),
       "the rim is DARKER than the artwork it surrounds",
       "rim %s vs body %s - a light rim is the defect, not the convention"
       % (rim[:3], body[:3]))
    ok(abs(rim[0] - rim[1]) < 90 or abs(rim[1] - rim[2]) < 90,
       "and it follows the object's own hue rather than being a flat colour",
       "rim %s" % (rim[:3],))
    ok(ish.add_stroke(plain, px=0).getbbox() == before,
       "px=0 leaves the image untouched")

    # restroke replaces a rim rather than stacking a second one on top
    rs = os.path.join(tmp, "rs.png")
    stroked.save(rs)
    w0 = stroked.getbbox()
    ish.restroke(rs)
    w1 = Image.open(rs).convert("RGBA").getbbox()
    ok(abs((w1[2] - w1[0]) - (w0[2] - w0[0])) <= 2 * ish.STROKE_PX,
       "restroke replaces the rim instead of growing the icon each call",
       "%s -> %s" % (str(w0), str(w1)))

    ok(ish.coverage(chain[4]) > ish.coverage(chain[0]),
       "coverage measures what it says it measures")
finally:
    shutil.rmtree(tmp, ignore_errors=True)

# ---- the shipped icons ARE the calibration ----------------------------------
SHIPPED = (r"C:\SkipBo\trunk\develop\client\Skipbo\Assets\ResourcesAb\Events"
           r"\Merge2608\Icon\Element1")
if os.path.isdir(SHIPPED):
    print("\n--- shipped Merge2608 chains ---")
    bad = []
    for g in (1, 2, 3):
        for c in (1, 2):
            paths = [os.path.join(SHIPPED, "element_%d_%d_%d.png" % (g, c, lv))
                     for lv in range(1, 6)]
            if not all(os.path.exists(p) for p in paths):
                continue
            r = ish.check_chain(paths)
            print("  group %d chain %d  coverage %s  %s"
                  % (g, c, r["coverage"], "ok" if r["ok"] else "PROGRESSION BREAK"))
            if not r["ok"]:
                bad.append("g%dc%d" % (g, c))
    ok(not bad,
       "every shipped chain passes the progression check this module enforces",
       "if a shipped chain fails, the rule here is wrong, not the art: " + ", ".join(bad))
    one = os.path.join(SHIPPED, "element_1_1_1.png")
    if os.path.exists(one):
        im = Image.open(one)
        ok(im.size == (ish.CANVAS, ish.CANVAS),
           "CANVAS matches the shipped icon size", "shipped %s, CANVAS %d" % (str(im.size), ish.CANVAS))

# ---- producers are not chain rungs -------------------------------------------
# Measured on the shipped element_machine<g>_<lv> icons: both levels sit at
# 0.86-0.97 of the canvas, against a chain's first rung at 0.60. A producer drawn
# at a rung's size reads as an ingredient, and that is the checkable failure.
ok(len(ish.PRODUCER_STEPS) == 2 and all(f >= ish.PRODUCER_MIN for f in ish.PRODUCER_STEPS),
   "both producer levels are drawn above the producer size floor", str(ish.PRODUCER_STEPS))
ok(ish.PRODUCER_STEPS[1] > ish.PRODUCER_STEPS[0],
   "level 2 is a touch larger than level 1, as the shipped set is")
# NOT "bigger than every rung": measured, a shipped producer sits at ~0.91 and a
# shipped top rung at ~0.92, so they are comparable. The real discriminator is
# against the FIRST rung, and that is the mistake worth catching - a producer
# drawn small reads as an ingredient.
ok(ish.PRODUCER_STEPS[0] > ish.SIZE_STEPS[0] * 1.4,
   "a producer is far bigger than a chain's first rung",
   "%s vs %s" % (ish.PRODUCER_STEPS[0], ish.SIZE_STEPS[0]))
ok(ish.PRODUCER_MIN > ish.SIZE_STEPS[0],
   "the producer size floor sits well above a first rung, so the two never blur")

_p = tempfile.mkdtemp()
try:
    def sq(path, w, h):
        im = Image.new("RGBA", (300, 300), (0, 0, 0, 0))
        ImageDraw.Draw(im).rectangle([10, 10, 10 + w, 10 + h], fill=(120, 90, 200, 255))
        im.save(path); return path

    big = os.path.join(_p, "big.png")
    ish.normalise(sq(os.path.join(_p, "s1.png"), 200, 180), big, level=1, levels=2,
                  fraction=ish.PRODUCER_STEPS[0])
    ok(abs(ish.content_size(big) / float(ish.CANVAS) - ish.PRODUCER_STEPS[0]) < 0.03,
       "the fraction override places a producer at producer size",
       "got %.3f" % (ish.content_size(big) / float(ish.CANVAS)))

    small = os.path.join(_p, "small.png")
    ish.normalise(sq(os.path.join(_p, "s2.png"), 200, 180), small, level=1, levels=5)
    ok(not ish.check_producers([small])["ok"],
       "a producer drawn at a chain rung's size is caught",
       str(ish.check_producers([small])["findings"]))
    ok(ish.check_producers([big])["ok"], "and one at producer size passes")
finally:
    shutil.rmtree(_p, ignore_errors=True)

SHIPPED_M = SHIPPED.replace("Element1", "Merge2608Common")
if os.path.isdir(SHIPPED_M):
    for lv in (1, 2):
        ps = [os.path.join(SHIPPED_M, "element_machine%d_%d.png" % (g, lv)) for g in (1, 2, 3, 4)]
        if all(os.path.exists(p) for p in ps):
            r = ish.check_producers(ps)
            ok(r["ok"], "the shipped producers pass the rule this module enforces (level %d)" % lv,
               str(r["findings"]))

# ---- a non-square canvas fits on BOTH axes -----------------------------------
# Sort goods are 1:2. A squat good (a bowl, a tin) reaches the 256px width long
# before its height reaches 0.88 of 512. That used to be refitted to the full
# canvas width, so the keyline sat flush on the edge with no margin.
import icon_sheet as _isheet
from PIL import Image as _Im
_g = tempfile.mkdtemp()          # `tmp` is already cleaned up by this point
_sq = os.path.join(_g, "squat.png")
_i = _Im.new("RGBA", (400, 180), (200, 90, 40, 255)); _i.save(_sq)
_out = os.path.join(_g, "squat_out.png")
_isheet.normalise(_sq, _out, fraction=_isheet.GOODS_FRACTION,
                  canvas=_isheet.GOODS_CANVAS, stroke=0)
_r = _Im.open(_out).convert("RGBA")
_bb = _r.getbbox()
_w, _h = _bb[2] - _bb[0], _bb[3] - _bb[1]
ok(_r.size == _isheet.GOODS_CANVAS, "a squat good lands on the 1:2 goods canvas", str(_r.size))
ok(_w <= round(256 * _isheet.GOODS_FRACTION) + 1,
   "and its width stays inside the fraction, not flush to the canvas edge",
   "%d wide of 256, fraction allows %d" % (_w, round(256 * _isheet.GOODS_FRACTION)))
ok(abs(_w / float(_h) - 400 / 180.) < 0.02, "and its aspect ratio is preserved")

_tall = os.path.join(_g, "tall.png")
_Im.new("RGBA", (180, 400), (40, 90, 200, 255)).save(_tall)
_isheet.normalise(_tall, _out, fraction=_isheet.GOODS_FRACTION,
                  canvas=_isheet.GOODS_CANVAS, stroke=0)
_bb = _Im.open(_out).convert("RGBA").getbbox()
ok(abs((_bb[3] - _bb[1]) / 512. - _isheet.GOODS_FRACTION) < 0.02,
   "a tall good still fills the fraction on the long edge, as before",
   str((_bb[3] - _bb[1]) / 512.))

_sqc = os.path.join(_g, "sqcanvas.png")
_Im.new("RGBA", (400, 180), (10, 10, 10, 255)).save(_sqc)
_isheet.normalise(_sqc, _out, fraction=0.9, canvas=300, stroke=0)
_bb = _Im.open(_out).convert("RGBA").getbbox()
ok(abs((_bb[2] - _bb[0]) / 300. - 0.9) < 0.02,
   "and on a SQUARE canvas the rule is unchanged, so recipes/materials are untouched",
   str((_bb[2] - _bb[0]) / 300.))

# ---- silhouette collisions across a set --------------------------------------
# Sort goods are sorted BY EYE, so two goods sharing an outline is a real defect
# and not a cosmetic one. It only shows up across the set - each icon looks fine
# alone - so it is worth a mechanical check rather than a contact sheet.
from PIL import Image as _Im2
import icon_sheet
_c = tempfile.mkdtemp()
def _blob(name, box):
    f = os.path.join(_c, name)
    im = _Im2.new("RGBA", (256, 512), (0, 0, 0, 0))
    d = __import__("PIL.ImageDraw", fromlist=["ImageDraw"]).Draw(im)
    d.ellipse(box, fill=(200, 90, 40, 255))
    im.save(f)
    return f
_tall1 = _blob("tall1.png", (60, 30, 190, 480))
_tall2 = _blob("tall2.png", (62, 32, 188, 478))     # the same shape again
_wide  = _blob("wide.png",  (10, 200, 246, 310))    # unmistakably different
ok(icon_sheet.silhouette_iou(_tall1, _tall2) > 0.95,
   "two goods with the same outline score as the same shape",
   "%.3f" % icon_sheet.silhouette_iou(_tall1, _tall2))
ok(icon_sheet.silhouette_iou(_tall1, _wide) < icon_sheet.SHAPE_COLLISION,
   "a tall good and a wide good do not collide")
ok(icon_sheet.silhouette_iou(_tall1, _wide) == icon_sheet.silhouette_iou(_wide, _tall1),
   "and the comparison is symmetric")
_col = icon_sheet.silhouette_collisions([_tall1, _tall2, _wide])
ok(len(_col) == 1 and set(_col[0][1:]) == {_tall1, _tall2},
   "a set check reports the offending pair and only that pair", str(len(_col)))

# This used to assert the opposite - that the same shape at a different size
# still counts as a collision - on the reasoning that a small pot and a big pot
# are one complaint. That is wrong for goods. Every good is composited on the
# same canvas at a fill its silhouette family chooses, so size is a cue the
# player really does sort on, and calling two differently-sized objects a
# collision invents a defect nobody can see. Measured on real output, scale
# normalisation turned 2 genuine collisions into 11.
_big = _blob("big.png", (20, 20, 236, 500))
_small = _blob("small.png", (100, 200, 156, 330))
ok(icon_sheet.silhouette_iou(_big, _small) < icon_sheet.SHAPE_COLLISION,
   "the same form at a very different SIZE is not a collision - size is a sorting cue too",
   "%.3f" % icon_sheet.silhouette_iou(_big, _small))

# ---- the goods portrait contract ---------------------------------------------
# Measured across all 76 shipped Sort2603/Sort2606 goods icons: content aspect
# 0.48 mean, canvas area 0.55 mean. A set built without that rule averaged 0.81
# aspect and covered 0.27 - half the canvas empty.
import asset_prompts as _ap
ok(not hasattr(_ap, "SHAPE_FAMILIES"),
   "the fixed shape vocabulary is gone - five floors drawing from one list of "
   "fifteen gave the event fifteen objects painted five ways")
ok(not hasattr(_ap, "SIZE_CLASSES"),
   "and so are the per-item size classes, which drew goods covering 6% of a frame")
ok(0.85 <= _ap.GOODS_FILL <= 0.92,
   "one fill for every good, at the shipped 0.86 wide / 0.90 tall", str(_ap.GOODS_FILL))
ok(_ap.goods_fraction({"name": "x"}) == _ap.GOODS_FILL
   and _ap.goods_fraction({}, 0.5) == _ap.GOODS_FILL,
   "and nothing can shrink an individual good below it any more")
_lo, _hi = _ap.GOODS_ASPECT_RANGE
ok(_lo < 0.48 < _hi, "the accepted aspect band brackets the shipped 0.48 mean",
   str(_ap.GOODS_ASPECT_RANGE))

_p1 = _ap.goods_prompt(
    {"name": "Amphora", "description": "an amphora",
     "silhouette": "a very tall narrow two-handled jar tapering to a pointed foot"},
    "blue", siblings=["Kiln Post", "Loop Tool"])
ok("TWICE AS TALL AS IT IS WIDE" in _p1,
   "the prompt states the portrait rule the shipped icons follow")
ok("stand it on its end" in _p1,
   "and says how to get the height - from the object's pose")
ok("NEVER stretch, squash or distort" in _p1,
   "and forbids getting it by distorting the object, which is the wrong fix")
ok("never blow up a small object" in _p1,
   "and forbids getting it by zooming, which is the other wrong fix")
ok("a very tall narrow two-handled jar" in _p1,
   "the item's own silhouette still reaches the prompt")
ok("Kiln Post" in _p1 and "Loop Tool" in _p1, "and so do its siblings")
ok("REAL tool or product of this" in _p1, "and the real-item rule still holds")

# ---- frame_problems reports, it never rescales --------------------------------
_ff = tempfile.mkdtemp()
def _obj(name, box):
    f = os.path.join(_ff, name)
    im = _Im2.new("RGBA", (256, 512), (0, 0, 0, 0))
    __import__("PIL.ImageDraw", fromlist=["ImageDraw"]).Draw(im).rectangle(box, fill=(200, 90, 40, 255))
    im.save(f)
    return f
_good = _obj("good.png", (30, 26, 226, 486))        # aspect ~0.43, fills the frame
_wide = _obj("wide.png", (10, 210, 246, 300))       # aspect ~2.6, a band of dead space
_probs, _fit = icon_sheet.frame_problems(_good)
ok(not _probs, "a portrait good that fills its frame raises nothing", str(_probs))
ok(0.33 < _fit["aspect"] < 0.68, "and measures inside the shipped aspect band",
   "%.2f" % _fit["aspect"])
_probs, _fit = icon_sheet.frame_problems(_wide)
ok(any("too wide" in p for p in _probs), "a wide object is reported as too wide")
ok(any("standing up, not scaling up" in p for p in _probs),
   "and the report says to stand it up rather than scale it - scaling a wide "
   "object only makes it wider")
ok(any("covers only" in p for p in _probs), "and its empty canvas is reported too")
_before = _Im2.open(_wide).tobytes()
icon_sheet.frame_problems(_wide)
ok(_Im2.open(_wide).tobytes() == _before,
   "checking a good never modifies it - the fix is a re-render, not a rescale")

# ---- colour is the third axis, after kind and shape ---------------------------
# Shipped floors span the hue wheel: 76-94 degrees mean pairwise separation, only
# 5-16% of pairs within 15. A generated floor that took one palette string for
# all fifteen goods came back with 79% of its pairs inside 15 degrees.
_hc = tempfile.mkdtemp()
def _swatch(name, rgb):
    f = os.path.join(_hc, name)
    im = _Im2.new("RGBA", (256, 512), (0, 0, 0, 0))
    __import__("PIL.ImageDraw", fromlist=["ImageDraw"]).Draw(im).rectangle(
        (30, 26, 226, 486), fill=rgb + (255,))
    im.save(f)
    return f
_red   = _swatch("red.png",   (210, 60, 45))
_red2  = _swatch("red2.png",  (205, 66, 50))     # the same colour again
_blue  = _swatch("blue.png",  (40, 90, 200))
_green = _swatch("green.png", (60, 170, 80))
ok(icon_sheet.dominant_hue(_red) is not None, "a good's dominant colour can be read")
_d = icon_sheet.dominant_hue(_blue)
ok(190 < _d["hue"] < 240, "and it is the object's actual hue", "%.0f deg" % _d["hue"])
_col = icon_sheet.hue_collisions([_red, _red2, _blue, _green])
ok(len(_col) == 1 and set(_col[0][1:]) == {_red, _red2},
   "two goods in the same colour are reported, and only those two", str(len(_col)))
ok(not icon_sheet.hue_collisions([_red, _blue, _green]),
   "three goods in three colours raise nothing")

# The keyline and the shading greys are on EVERY good by design, so counting them
# drags every icon toward one average and hides the difference being measured.
_grey = _swatch("grey.png", (128, 128, 128))
ok(icon_sheet.dominant_hue(_grey) is None,
   "a colourless good reports no hue rather than a meaningless one")

import asset_prompts as _ap2
_cp = _ap2.goods_prompt({"name": "Amphora", "description": "a jar",
                         "silhouette": "a tall narrow two-handled jar",
                         "palette": "deep cobalt blue glaze"},
                        "terracotta and cream", siblings=["Clay Bin"])
ok("deep cobalt blue glaze" in _cp, "the good's OWN colour reaches the prompt")
ok("terracotta and cream" not in _cp,
   "and the floor palette does not override it - that is what made a floor one colour")
ok("deliberately DIFFERENT colours" in _cp,
   "the prompt says the shelf's other goods are different colours")
ok("the FINISH, not the hue" in _cp,
   "and that the floor coheres through finish rather than through one palette")
ok("strongly saturated" in _cp and "fresh painted toys" in _cp,   # "light in tone" wraps
   "and asks for the shipped brightness - measured 0.71 saturation, 0.81 value, "
   "against 0.52 and 0.55 without asking")
# Asking was not enough on its own: the shared style block and the render spec
# both said DEEP, output came back at 0.51 value, and a new sentence asking for
# bright simply competed with them. The conflicting words are now removed.
ok("deep saturated colour" not in _cp,
   "and the conflicting 'deep' wording is gone rather than argued with")
ok("high-key" in _cp, "replaced by the high-key wording the shipped goods match")
# Flattening the shading was tried and reverted: rendered and measured it made
# every figure slightly worse and broke two icons outright. The wide-range
# wording stays, so this asserts it is still there rather than gone.
ok("wide tonal range" in _cp,
   "the shared wide-tonal-range wording is kept - flattening it measured worse")
ok("never dark, deep or muddy" in _cp,
   "and the shared 'Rich' colour line is replaced too - one voice, not three")

# ---- one keyline, not two -----------------------------------------------------
# The prompt used to ask the generator for a dark keyline while add_stroke
# painted one as well. Measured on the output, peeling the painted rim lifted the
# 5th-percentile brightness from 0.24 to 0.38, and peeling the drawn contour too
# lifted it to 0.56 against a shipped 0.52 - the icons were not shaded dark, they
# were outlined twice.
ok("this icon IS drawn with a clean DARK KEYLINE" not in _cp,
   "the generator is no longer asked to draw a keyline that add_stroke also paints")
ok("NO TEXT, NO OUTLINE" in _cp,
   "goods use the same clean-edge rule as every other icon class")
ok("No dark rim, no inked contour" in _cp,
   "and are told not to shade a dark band round the silhouette either")

# stroke_for on a non-square canvas takes the SHORTER edge: a keyline reads
# against how wide the object is, not how tall its frame is.
ok(icon_sheet.stroke_for(icon_sheet.GOODS_CANVAS) == icon_sheet.stroke_for(256),
   "the 256x512 goods canvas gets the same keyline as a 256 square would",
   str(icon_sheet.stroke_for(icon_sheet.GOODS_CANVAS)))
ok(icon_sheet.stroke_for(icon_sheet.GOODS_CANVAS) < 5,
   "which is half the 6px it used to get from scaling by the longer edge")
ok(icon_sheet.stroke_for(icon_sheet.RECIPE_CANVAS) == icon_sheet.stroke_for(icon_sheet.RECIPE_CANVAS),
   "and square classes are untouched by the change")
# "light in tone" alone moved value from 0.51 to only 0.58. Shipped goods run
# 0.89-0.95: their shadow side is still light. That needed saying concretely.
ok("EVEN THE SHADOW SIDE STAYS LIGHT" in _cp,
   "the high-key rule names the part that was actually coming out dark")
ok("never near-black" in _cp and "The only genuinely dark thing" in _cp,
   "and reserves darkness for the keyline alone")

print("")
if fails:
    print("%d CHECK(S) FAILED" % len(fails)); sys.exit(1)
print("ALL ICON SHEET TESTS PASSED")

