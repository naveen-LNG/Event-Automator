"""Detect reference bleed: the generator pasting an attached reference into the output.

This happened during the Sort work. With five reference floors attached, one
glaze-room render came back with the entire Sort2609 cat room composited over it
- pink striped wallpaper, arched window, cat tower, food bowls - with the
requested sea-green room underneath. Text, halo, blank and duplicate detection
all passed it, because nothing was looking for "did the model copy its reference".

WHAT DOES NOT WORK, and why it is not in here: comparing the output to each
reference pixel-for-pixel on a coarse grid. Measured on the known-bad image it
scored 0.10 against the room it had copied, while clean floors scored 0.11-0.27
against references they had nothing to do with - the detector ranked the
defective image as the cleanest of the set. Two reasons. The bleed arrives as a
semi-transparent overlay blended with the room that was actually requested, so
no region is ever pixel-identical to the source; and the output and the
references differ in aspect, so a resize to a common grid misaligns every
feature anyway. Regional matching (per-quarter, best window) failed the same way.
Colour-similarity to the reference is the wrong signal. Do not re-try it.

WHAT DOES WORK: the bleed drags the reference's palette in with it, and that
shows up as a violation of the tone the prompt specified. The brief for the glaze
room said COOL and named sea-green; the bled render came back 53.9% warm-hued
because the cat room it copied is pink. Measuring off-tone area separates the
known-bad case from every clean one with real margin:

    BLED  floor3 glaze  (spec COOL)   53.9% warm   <-- caught
    clean floor3 glaze  (spec COOL)    1.1% warm
    clean floor4 guest  (spec COOL)    1.9% warm
    clean floor5 terrace(spec COOL)   19.5% warm   (deliberate golden-hour light)
    shipped aviary      (mixed)       27.5% warm
    clean floor1 hall   (spec WARM)   99.9% warm   (correct - warm room)

Hence OFF_TONE_FRACTION below. This is narrower than "any bleed" - a reference
copied in whose palette happens to agree with the brief slips through. It catches
the failure mode that actually shipped, which is the one worth a gate.
"""

import os

# Hue bands in degrees. Near-neutrals are excluded before any of this applies -
# a grey wall has no tone opinion and should not vote.
WARM_MAX = 45        # 0-45: red, orange, amber
WARM_MIN = 310       # 310-360: pink, magenta, rose
COOL_LO, COOL_HI = 150, 280   # green, cyan, blue, violet

MIN_SAT = 0.18       # below this the pixel reads as neutral, not as a hue
MIN_VAL = 0.15       # near-black carries no usable hue
SAMPLE_STEP = 4      # every 4th pixel both ways - 16x cheaper, same fractions

OFF_TONE_FRACTION = 0.35
# Sits above the highest legitimate reading measured (27.5%, a shipped floor that
# genuinely mixes warm into a cool scheme) and well below the bled 53.9%.


def tone_profile(path, step=SAMPLE_STEP):
    """Return (warm_fraction, cool_fraction) over hue-bearing opaque pixels."""
    import colorsys
    from PIL import Image

    im = Image.open(path).convert("RGBA")
    px = im.load()
    w, h = im.size
    warm = cool = total = 0
    for y in range(0, h, step):
        for x in range(0, w, step):
            r, g, b, a = px[x, y]
            if a < 128:
                continue
            hh, ss, vv = colorsys.rgb_to_hsv(r / 255.0, g / 255.0, b / 255.0)
            if ss < MIN_SAT or vv < MIN_VAL:
                continue
            total += 1
            deg = hh * 360.0
            if deg <= WARM_MAX or deg >= WARM_MIN:
                warm += 1
            elif COOL_LO <= deg <= COOL_HI:
                cool += 1
    if not total:
        return 0.0, 0.0
    return warm / float(total), cool / float(total)


def check(out_path, tone, threshold=OFF_TONE_FRACTION):
    """Does the output honour the tone its brief specified?

    tone: "COOL" or "WARM" - the MAIN TONE the prompt demanded. Anything else
    (None, "", "MIXED") means the brief did not commit, so there is nothing to
    check against and this returns clean.

    Returns {"bleed": bool, "off_tone": float, "warm": float, "cool": float,
             "tone": str, "detail": str}
    """
    t = (tone or "").strip().upper()
    if t not in ("COOL", "WARM"):
        return {"bleed": False, "off_tone": 0.0, "warm": 0.0, "cool": 0.0, "tone": t,
                "detail": "brief specifies no main tone - nothing to check"}

    warm, cool = tone_profile(out_path)
    off = warm if t == "COOL" else cool
    bled = off >= threshold
    if bled:
        wrong = "warm (red, orange, pink, magenta)" if t == "COOL" else "cool (green, cyan, blue)"
        detail = (
            "%.0f%% of the coloured area is %s in a room the brief specified as %s. That much "
            "off-tone area means content was carried in from an attached reference rather than "
            "drawn to the brief. Cut the number of attached references and drop any whose palette "
            "is close to what you see wrongly appearing." % (off * 100, wrong, t)
        )
    else:
        detail = "%s room, %.0f%% off-tone area - within the tone the brief asked for" % (t, off * 100)
    return {"bleed": bled, "off_tone": round(off, 3), "warm": round(warm, 3),
            "cool": round(cool, 3), "tone": t, "detail": detail}
