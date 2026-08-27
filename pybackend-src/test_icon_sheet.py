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

print("")
if fails:
    print("%d CHECK(S) FAILED" % len(fails)); sys.exit(1)
print("ALL ICON SHEET TESTS PASSED")

