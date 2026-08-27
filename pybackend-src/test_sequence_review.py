"""Tests for the scene-sequence reviewer.

Synthetic frames with known-good and known-bad build orders, so the checks are
verified without spending renders. Ends with a run against the REAL Harvest Fair
scenes. Scene 5 originally shipped with a stall standing in frames it should not
have been in; the reviewer caught it, it was repaired, and all four now pass.
"""
import os, sys, tempfile, shutil

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
HERE = os.path.dirname(os.path.abspath(__file__))
os.chdir(HERE); sys.path.insert(0, HERE)

from PIL import Image, ImageDraw
import sequence_review as sr

fails = []
def ok(c, m, d=""):
    print(("OK  " if c else "FAIL ") + m + (("\n     " + d) if (d and not c) else ""))
    if not c: fails.append(m)

W, H = 240, 420
def frame(objects, bg=(240, 232, 200), tint=0):
    # Static scenery, so an "empty" frame is not literally a flat colour - a real
    # empty frame still has sky, ground and buildings in it.
    im = Image.new("RGB", (W, H), bg)
    d = ImageDraw.Draw(im)
    d.rectangle([0, 0, W, 170], fill=(150, 200, 235))          # sky
    d.rectangle([0, 170, W, 215], fill=(170, 215, 165))        # distant fields
    d.polygon([(10, 170), (60, 110), (110, 170)], fill=(200, 140, 120))   # left building
    d.rectangle([150, 140, 235, 175], fill=(190, 170, 150))    # right shed
    d.rectangle([0, 355, W, H], fill=(205, 175, 120))          # foreground frame band
    boxes = {"hero":   (95, 250, 150, 340, (20, 120, 130)),
             "tall":   (20, 150, 65, 340,  (240, 110, 90)),
             "wide":   (20, 300, 150, 350, (120, 220, 170)),
             "round":  (165, 240, 225, 300,(190, 160, 235)),
             "compact":(170, 305, 225, 350,(230, 80, 140)),
             "crown":  (80, 60, 160, 110,  (235, 200, 110))}
    for o in objects:
        x0, y0, x1, y1, c = boxes[o]
        d.rectangle([x0, y0, x1, y1], fill=c)
    if tint:
        # uniform shift across every pixel, small enough to stay under the
        # "changed" threshold so it registers as drift rather than as content
        px = im.load()
        for y in range(H):
            for x in range(W):
                r, g, b = px[x, y]
                px[x, y] = (max(0, r - tint), max(0, g - tint), max(0, b - tint))
    return im

tmp = tempfile.mkdtemp()
try:
    BUILD = [[], ["hero"], ["hero","tall"], ["hero","tall","wide"],
             ["hero","tall","wide","round"], ["hero","tall","wide","round","compact"],
             ["hero","tall","wide","round","compact","crown"]]
    labels = ["F2_empty","F3_hero","F4_tall","F5_wide","F6_round","F7_compact","F8_base"]
    paths = []
    for objs, lab in zip(BUILD, labels):
        p = os.path.join(tmp, lab + ".png"); frame(objs).save(p); paths.append(p)
    base = paths[-1]

    # 1. a correct sequence passes
    r = sr.review(base, paths, labels=labels)
    ok(r["ok"], "a correct build sequence passes", sr.format_report(r))
    d = [m["diff_from_base"] for m in r["metrics"]["frames"]]
    ok(all(d[i] > d[i+1] for i in range(len(d)-1)),
       "diff-from-base decreases monotonically along the build order", str(d))

    # 2. THE SCENE 5 BUG: an object that should have been removed is still present.
    bad = list(paths)
    leak = os.path.join(tmp, "F2_leak.png")
    frame(["wide"]).save(leak)            # 'wide' must not exist yet in the empty frame
    bad[0] = leak
    r = sr.review(base, bad, labels=labels)
    ok(not r["ok"], "a leaked object in the empty frame FAILS the review")
    ok(any("did not take" in f["message"] or "went backwards" in f["message"]
           or "added nothing" in f["message"] for f in r["findings"]),
       "the finding names the real cause (a removal that did not take)",
       str([f["message"][:90] for f in r["findings"]]))

    # 3. a step that changed nothing is caught
    dup = list(paths)
    same = os.path.join(tmp, "F5_dup.png")
    frame(["hero","tall"]).save(same)     # identical content to F4
    dup[3] = same
    r = sr.review(base, dup, labels=labels)
    ok(not r["ok"], "a build step that adds nothing FAILS the review")

    # 4a. a real blank render is a VALID png that is one flat colour
    blanks = list(paths)
    bp = os.path.join(tmp, "blank.png")
    Image.new("RGB", (W, H), (255, 255, 255)).save(bp)
    blanks[2] = bp
    r = sr.review(base, blanks, labels=labels)
    ok(not r["ok"] and any("blank" in f["message"] for f in r["findings"]),
       "a flat blank render FAILS the review",
       str([f["message"][:70] for f in r["findings"]]))

    # 4b. a truncated write is a different fault and is caught separately
    trunc = list(paths)
    tp = os.path.join(tmp, "trunc.png")
    open(tp, "wb").write(b"\x89PNG" + b"\0" * 200)
    trunc[2] = tp
    r = sr.review(base, trunc, labels=labels)
    ok(not r["ok"] and any(("truncated" in f["message"]) or ("corrupt" in f["message"])
                           for f in r["findings"]),
       "a truncated/corrupt file FAILS the review",
       str([f["message"][:70] for f in r["findings"]]))

    # 5. palette drift in unchanged areas is flagged
    drifted = list(paths)
    dp = os.path.join(tmp, "F4_drift.png")
    frame(["hero","tall"], tint=20).save(dp)   # whole frame shifted subtly
    drifted[2] = dp
    r = sr.review(base, drifted, labels=labels)
    ok(any("palette drift" in f["message"] for f in r["findings"]),
       "palette drift in unchanged areas is reported")

    # 6. derelict must actually be drained, not just dim
    der = os.path.join(tmp, "derelict.png")
    Image.new("RGB", (W, H), (150, 150, 155)).save(der)
    r = sr.review(base, paths, derelict_path=der, labels=labels)
    ok(r["metrics"]["derelict_saturation"] < r["metrics"]["empty_saturation"],
       "derelict saturation measured below the restored frame")
finally:
    shutil.rmtree(tmp, ignore_errors=True)

# ---- against the real generated scenes -------------------------------------
REAL = r"C:\SkipBo\tools\EventAutomationDesktop\ideationData\Events\Blocks\Art Requirements\Drafts"
ORDER = ["01_F2_empty.png","02_F3.png","03_F4.png","04_F5.png","05_F6.png","06_F7.png","07_F8_complete.png"]
print("\n--- real Harvest Fair scenes ---")
seen = {}
for slug in ["scene1_fairgate","scene3_barndance","scene4_cideryard","scene5_nightfair"]:
    d = os.path.join(REAL, "Blocks2610_HarvestFair_" + slug)
    if not os.path.isdir(d):
        print("  %-18s (not generated)" % slug); continue
    fp = [os.path.join(d, f) for f in ORDER]
    if not all(os.path.exists(p) for p in fp):
        print("  %-18s (incomplete)" % slug); continue
    r = sr.review(os.path.join(d, "07_F8_complete.png"), fp,
                  derelict_path=os.path.join(d, "00_F1_derelict.png"), labels=ORDER)
    seen[slug] = r["ok"]
    print("  %-18s %s" % (slug, "PASS" if r["ok"] else "FAIL"))
    for f in r["findings"]:
        if f["level"] == "error":
            print("      - %s: %s" % (f["frame"], f["message"][:120]))

if seen:
    # Scene 5 shipped with a leaked stall and the reviewer caught it; it has
    # since been repaired, so the live assertion is that ALL scenes now pass.
    # The leak itself stays permanently guarded by the synthetic test above,
    # which does not depend on the state of the real artwork.
    ok(all(seen.values()),
       "every generated scene passes the reviewer",
       str({k: ("PASS" if v else "FAIL") for k, v in seen.items()}))

print("")
if fails:
    print("%d CHECK(S) FAILED" % len(fails)); sys.exit(1)
print("ALL SEQUENCE REVIEW TESTS PASSED")
