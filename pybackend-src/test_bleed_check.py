"""Tests for the reference-bleed / tone-conformance check.

Ground truth is the real Sort output from this session: one render that came back
with an entire reference room composited into it, and five that did not.

The first implementation of this module compared output to reference pixel-wise
on a coarse grid and scored the bled image LOWER than every clean one. The
ground-truth block below is what caught that, so it stays.
"""
import os, sys, tempfile, shutil

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
HERE = os.path.dirname(os.path.abspath(__file__))
os.chdir(HERE); sys.path.insert(0, HERE)

from PIL import Image, ImageDraw
import bleed_check as bc

fails = []
def ok(c, m, d=""):
    print(("OK  " if c else "FAIL ") + m + (("\n     " + d) if (d and not c) else ""))
    if not c: fails.append(m)

tmp = tempfile.mkdtemp()
try:
    W, H = 240, 320

    def solid(rgb):
        p = os.path.join(tmp, "%d_%d_%d.png" % rgb)
        Image.new("RGBA", (W, H), rgb + (255,)).save(p)
        return p

    ok(bc.check(solid((90, 150, 140)), "COOL")["bleed"] is False,
       "an all-sea-green plate passes a COOL brief")
    ok(bc.check(solid((230, 140, 170)), "COOL")["bleed"] is True,
       "an all-pink plate fails a COOL brief")
    ok(bc.check(solid((230, 140, 170)), "WARM")["bleed"] is False,
       "the same pink plate passes a WARM brief")
    ok(bc.check(solid((230, 140, 170)), None)["bleed"] is False,
       "no declared tone means no check")

    # neutrals must not vote: a grey room is not an off-tone room
    ok(bc.check(solid((150, 150, 150)), "COOL")["bleed"] is False,
       "a neutral grey plate passes either brief - near-neutrals do not vote")

    # half pink over a cool room clears the threshold; a small accent does not
    im = Image.new("RGBA", (W, H), (90, 150, 140, 255))
    ImageDraw.Draw(im).rectangle([0, 0, W, H // 2], fill=(230, 140, 170, 255))
    p = os.path.join(tmp, "half.png"); im.save(p)
    ok(bc.check(p, "COOL")["bleed"] is True, "half the frame off-tone is caught")

    im = Image.new("RGBA", (W, H), (90, 150, 140, 255))
    ImageDraw.Draw(im).rectangle([10, 10, 50, 50], fill=(220, 160, 80, 255))
    p = os.path.join(tmp, "accent.png"); im.save(p)
    ok(bc.check(p, "COOL")["bleed"] is False,
       "a small warm accent is allowed - the Mattel rules require one")
finally:
    shutil.rmtree(tmp, ignore_errors=True)

# ---- real output from this session, with known ground truth ------------------
B = r"C:\Users\NAVEEN~1\AppData\Local\Temp\claude\C--SkipBo\b4d775b4-e255-4b0e-9923-1aa47c2b64cf\scratchpad"
REAL = [
    (os.path.join(B, "bakery", "floor3_glaze_room.png"), "COOL", True,
     "cat room composited in"),
    (os.path.join(B, "seaside_fix", "floor3_glaze_room.png"), "COOL", False,
     "same brief, 2 refs, regenerated"),
    (os.path.join(B, "bakery", "floor4_guest_room.png"), "COOL", False, "guest room"),
    (os.path.join(B, "bakery", "floor5_sea_terrace.png"), "COOL", False,
     "sea terrace - golden hour over cool"),
    (os.path.join(B, "bakery", "floor1_welcome_hall.png"), "WARM", False, "welcome hall"),
    (os.path.join(B, "bakery", "floor2_pottery_studio.png"), "WARM", False, "pottery studio"),
]
present = [r for r in REAL if os.path.exists(r[0])]
if present:
    print("\n--- real images (ground truth from this session) ---")
    wrong = []
    for path, tone, expect, label in present:
        r = bc.check(path, tone)
        got = r["bleed"]
        print("  %-38s %-4s off-tone %5.1f%%  expect %-5s got %-5s %s"
              % (label[:38], tone, r["off_tone"] * 100, expect, got,
                 "" if got == expect else "<-- MISS"))
        if got != expect:
            wrong.append(label)
    ok(not wrong, "check agrees with ground truth on all %d real images" % len(present),
       "disagreed on: " + ", ".join(wrong))
else:
    print("\n(no session images on disk - synthetic checks only)")

print("")
if fails:
    print("%d CHECK(S) FAILED" % len(fails)); sys.exit(1)
print("ALL BLEED CHECK TESTS PASSED")
