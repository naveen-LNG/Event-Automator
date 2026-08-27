"""Tests for the asset cut-out post-step: knockout, trim, scoop.

Synthetic images with a known plate, plus the real shipped Sort floors as the
ground truth for the scoop curve.
"""
import os, shutil, sys, tempfile

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
HERE = os.path.dirname(os.path.abspath(__file__))
os.chdir(HERE); sys.path.insert(0, HERE)

from PIL import Image, ImageDraw
import scoop as sc

fails = []
def ok(c, m, d=""):
    print(("OK  " if c else "FAIL ") + m + (("\n     " + d) if (d and not c) else ""))
    if not c: fails.append(m)

W, H = 400, 520
BG = (250, 250, 250)

def make_plate(path, bg=BG, wall=(180, 210, 235), plate=(240, 225, 190)):
    """A backdrop, two walls, an isometric-ish plate, and an interior region
    that matches the backdrop colour exactly - the trap a naive global
    colour-replace would fall into."""
    im = Image.new("RGBA", (W, H), bg + (255,))
    d = ImageDraw.Draw(im)
    d.polygon([(60, 120), (200, 60), (340, 120), (340, 300), (60, 300)], fill=wall + (255,))
    d.polygon([(200, 250), (360, 340), (200, 440), (40, 340)], fill=plate + (255,))
    # interior patch the SAME colour as the backdrop, enclosed by the plate
    d.rectangle([170, 320, 230, 360], fill=bg + (255,))
    im.save(path)
    return path

tmp = tempfile.mkdtemp()
try:
    p = make_plate(os.path.join(tmp, "plate.png"))
    before = sc.transparency(p)
    ok(before == 0.0, "test fixture starts fully opaque")

    # 1. knockout ------------------------------------------------------------
    sc.knockout_background(p)
    after = sc.transparency(p)
    ok(after > 0.25, "knockout makes the backdrop transparent (%.1f%%)" % (after * 100))

    im = Image.open(p)
    a = im.getchannel("A").load()
    ok(a[5, 5] == 0 and a[W - 5, 5] == 0, "the outer corners are cut")
    ok(a[200, 340] > 200, "the plate itself survives")
    ok(a[200, 200] > 200, "the walls survive")
    # the enclosed same-colour patch must NOT be eaten: flood fill starts at the
    # border, so an interior region is unreachable however well it matches
    ok(a[200, 340] > 200 and a[195, 335] > 200,
       "an interior region matching the backdrop colour is NOT eaten")

    # 1b. a GRADIENT backdrop must clear too ---------------------------------
    # A fixed-seed comparison only clears a perfectly flat backdrop; the
    # generator often returns a gradient sky and half of it stayed opaque.
    pg = os.path.join(tmp, "grad.png")
    img = Image.new("RGBA", (W, H))
    dg = ImageDraw.Draw(img)
    for y in range(H):                      # smooth vertical sky gradient
        v = 130 + int(110 * y / H)
        dg.line([(0, y), (W, y)], fill=(v, v + 8, 255, 255))
    dg.polygon([(60, 120), (200, 60), (340, 120), (340, 300), (60, 300)],
               fill=(180, 210, 235, 255))
    dg.polygon([(200, 250), (360, 340), (200, 440), (40, 340)], fill=(240, 225, 190, 255))
    img.save(pg)
    sc.knockout_background(pg)
    gt = sc.transparency(pg)
    ok(gt > 0.30, "a gradient backdrop is cleared, not abandoned partway (%.1f%%)" % (gt * 100))
    ga = Image.open(pg).getchannel("A").load()
    ok(ga[5, 5] == 0 and ga[W - 5, H - 5] == 0,
       "both the light end and the dark end of the gradient are cut")
    ok(ga[200, 340] > 200, "the plate still survives a gradient backdrop")

    # 2. trim ----------------------------------------------------------------
    p2 = make_plate(os.path.join(tmp, "trim.png"))
    sc.knockout_background(p2)
    big = Image.open(p2).size
    sc.trim_to_content(p2)
    small = Image.open(p2).size
    ok(small[0] < big[0] and small[1] < big[1],
       "trim crops to content (%dx%d -> %dx%d)" % (big + small))
    ok(sc.transparency(p2) < after,
       "trimming reduces the transparent share (%.1f%% -> %.1f%%)"
       % (after * 100, sc.transparency(p2) * 100))

    # 3. scoop ---------------------------------------------------------------
    p3 = make_plate(os.path.join(tmp, "scoop.png"))
    ok(sc.measure_scoop(p3) < 0.02, "an unscooped image measures ~0")
    sc.apply_top_scoop(p3)
    got = sc.measure_scoop(p3)
    ok(abs(got - sc.DEFAULT_DEPTH) < 0.03,
       "scoop depth lands on target (%.3f vs %.3f)" % (got, sc.DEFAULT_DEPTH))
    sc.apply_top_scoop(p3)
    ok(abs(sc.measure_scoop(p3) - got) < 0.02, "applying the scoop twice is idempotent")

    # 4. the metric must not mistake a plate shape for a scoop ---------------
    p4 = make_plate(os.path.join(tmp, "shape.png"))
    sc.cut_asset(p4, scoop=False)
    ok(sc.measure_scoop(p4) < 0.05,
       "after knockout+trim an UNSCOOPED plate still measures ~0 (%.3f)" % sc.measure_scoop(p4),
       "raw depth reported 0.74 here before the metric was fixed")

    # 5. the whole pipeline in order -----------------------------------------
    p5 = make_plate(os.path.join(tmp, "full.png"))
    sc.cut_asset(p5, scoop=True)
    ok(sc.transparency(p5) > 0.05 and sc.measure_scoop(p5) > 0.10,
       "cut_asset knocks out, trims and scoops in one call")
finally:
    shutil.rmtree(tmp, ignore_errors=True)

# ---- against the real shipped assets ---------------------------------------
S = r"C:\SkipBo\tools\EventAutomationDesktop\ideationData\StyleReference\images\Sort"
if os.path.isdir(S):
    print("\n--- shipped Sort floors ---")
    measured = {}
    for n in sorted(os.listdir(S)):
        measured[n] = sc.measure_scoop(os.path.join(S, n))
        print("  %-36s %.3f" % (n, measured[n]))
    scooped = [v for v in measured.values() if v > 0.05]
    flat = [v for v in measured.values() if v <= 0.05]
    ok(len(scooped) == 2 and len(flat) == 3,
       "exactly 2 of the 5 shipped floors carry a scoop", str(measured))
    ok(all(0.13 < v < 0.20 for v in scooped),
       "the shipped scoops sit in the 0.13-0.20 band DEFAULT_DEPTH targets", str(scooped))

# ---- halo direction: the defect is a PALE fringe, not any fringe -------------
# This was a real bug. Both the reviewer and dehalo gated on abs(score), so a
# dark contact shadow at the plate rim read as the Mattel white-edge defect, and
# "correcting" it drove one floor from -20.0 to -43.4. Ground truth: the five
# shipped, Mattel-approved Sort floors score -19.4, -13.9, +1.2, 0.0, -19.4.
S = r"C:\SkipBo\tools\EventAutomationDesktop\ideationData\StyleReference\images\Sort"
if os.path.isdir(S):
    shipped = {n: sc.edge_halo_score(os.path.join(S, n)) for n in sorted(os.listdir(S))}
    print("\n--- shipped art halo scores ---")
    for n, v in shipped.items():
        print("  %-34s %+6.1f" % (n.replace("sort2609_zone_", ""), v))
    ok(all(v <= sc.HALO_TOLERANCE for v in shipped.values()),
       "no shipped floor is flagged as having a pale fringe",
       "flagged: %s" % {k: round(v, 1) for k, v in shipped.items() if v > sc.HALO_TOLERANCE})
    ok(any(v < -sc.HALO_TOLERANCE / 2 for v in shipped.values()),
       "shipped art does carry a dark rim, so the negative case is real and must be left alone")

    # dehalo must be a no-op on a dark-rimmed plate
    import shutil as _sh, tempfile as _tf
    d = _tf.mkdtemp()
    try:
        darkest = min(shipped, key=shipped.get)
        src = os.path.join(d, darkest)
        _sh.copyfile(os.path.join(S, darkest), src)
        before = sc.edge_halo_score(src)
        sc.dehalo(src)
        after = sc.edge_halo_score(src)
        ok(abs(after - before) < 0.01,
           "dehalo leaves a dark-rimmed plate untouched (%+.1f -> %+.1f)" % (before, after),
           "it used to erode it further and report a worse score")
    finally:
        _sh.rmtree(d, ignore_errors=True)

# ---- enclosed backdrop: the border flood cannot reach it ---------------------
# knockout_background floods inward from the border, so backdrop enclosed by the
# artwork survives - the hole in a rope grommet, the mesh of a net, the middle of
# a wreath. Eight of forty generated icons came back with a grey disc in the hole.
_t = tempfile.mkdtemp()
try:
    BG = (146, 146, 146)                      # measured off the real leftovers
    ring = Image.new("RGBA", (200, 200), (0, 0, 0, 0))
    dr = ImageDraw.Draw(ring)
    dr.ellipse([20, 20, 180, 180], fill=(210, 120, 60, 255))     # the object
    dr.ellipse([70, 70, 130, 130], fill=BG + (255,))             # enclosed backdrop
    q = os.path.join(_t, "ring.png"); ring.save(q)
    before = sc.transparency(q)
    res = sc.knockout_enclosed(q)
    after = sc.transparency(q)
    ok(after > before, "enclosed backdrop is removed (transparency %.2f -> %.2f)" % (before, after))
    ok(Image.open(q).convert("RGBA").load()[100, 100][3] == 0,
       "the hole in the middle is actually open now")
    ok(Image.open(q).convert("RGBA").load()[30, 100][3] == 255,
       "and the object itself is untouched")

    # a tinted near-grey must survive: the cream rope in those icons reads
    # saturation 0.32 and is artwork, not backdrop
    tint = Image.new("RGBA", (100, 100), (154, 134, 105, 255))
    q2 = os.path.join(_t, "tint.png"); tint.save(q2)
    sc.knockout_enclosed(q2)
    ok(sc.transparency(q2) < 0.01,
       "a tinted near-neutral colour is NOT mistaken for backdrop",
       "the cream rope would be eaten if this fired on it")

    # so must a bright highlight and a dark shadow - only the mid band goes
    for name, col in (("white", (250, 250, 250, 255)), ("dark", (30, 30, 30, 255))):
        q3 = os.path.join(_t, name + ".png")
        Image.new("RGBA", (60, 60), col).save(q3)
        sc.knockout_enclosed(q3)
        ok(sc.transparency(q3) < 0.01,
           "a neutral %s stays - only the mid value band is backdrop" % name)
finally:
    shutil.rmtree(_t, ignore_errors=True)

print("")
if fails:
    print("%d CHECK(S) FAILED" % len(fails)); sys.exit(1)
print("ALL SCOOP TESTS PASSED")


