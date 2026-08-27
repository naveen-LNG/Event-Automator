"""Automated review of a generated scene-evolution sequence.

Every frame in a build sequence is produced as a subtractive edit of one base
image, which gives a structural invariant worth checking:

    the earlier the frame, the MORE it differs from the completed base.

So the diff-from-base must shrink monotonically along the build order, and each
step must change a real, localised amount. That single idea catches the two
failure modes that actually happen:

  * a removal silently did not take   -> that step barely differs from the next
  * a removal took too much, or the model re-drew the scene
                                      -> that step changes a huge share of the frame

Both got through visual inspection during the Harvest Fair run. Scene 5 shipped
with a stall standing in three frames it should not have been in, and nobody
caught it until the contact sheet was studied closely.

Also checks blank output, palette drift against the base, and that the derelict
frame is genuinely desaturated rather than merely dim.

CLI (payload JSON: {base, frames:[...in build order], derelict?}):
    pybackend_cli.py sequence_review payload.json
"""

import json
import os
import sys

# Fractions are of total pixel count.
TRUNCATED_BYTES = 512         # smaller than this is a truncated/failed write
BLANK_STDEV = 3.0             # decoded-content test for a blank render
DIFF_THRESHOLD = 28           # per-channel delta that counts as "changed"
MIN_STEP = 0.004              # a build step must change >=0.4% of the frame
MAX_STEP = 0.55               # >55% changed means the scene was re-drawn, not edited
PALETTE_TOLERANCE = 14.0      # mean per-channel drift allowed in unchanged areas
DERELICT_SAT_RATIO = 0.6      # derelict must be <=60% the saturation of the empty frame
SAMPLE = 4                    # sample every Nth pixel; plenty at 1080x1920


def _stdev(im):
    """Spread of pixel values. A blank render is flat, whatever its file size."""
    px = im.load(); w, h = im.size
    n = w * h
    tot = 0.0
    for y in range(h):
        for x in range(w):
            r, g, b = px[x, y]
            tot += (r + g + b) / 3.0
    mean = tot / n
    var = 0.0
    for y in range(h):
        for x in range(w):
            r, g, b = px[x, y]
            var += ((r + g + b) / 3.0 - mean) ** 2
    return (var / n) ** 0.5


def _load(path):
    from PIL import Image
    im = Image.open(path).convert("RGB")
    w, h = im.size
    if SAMPLE > 1:
        im = im.resize((max(1, w // SAMPLE), max(1, h // SAMPLE)), Image.BILINEAR)
    return im


def _diff_mask(a, b):
    """Fraction of pixels that differ, plus the mean drift over the rest."""
    pa, pb = a.load(), b.load()
    w, h = a.size
    changed = 0
    same_drift = 0
    same_n = 0
    for y in range(h):
        for x in range(w):
            r1, g1, b1 = pa[x, y]
            r2, g2, b2 = pb[x, y]
            d = max(abs(r1 - r2), abs(g1 - g2), abs(b1 - b2))
            if d >= DIFF_THRESHOLD:
                changed += 1
            else:
                same_drift += (abs(r1 - r2) + abs(g1 - g2) + abs(b1 - b2)) / 3.0
                same_n += 1
    total = w * h
    return changed / total, (same_drift / same_n if same_n else 0.0)


def _mean_saturation(im):
    px = im.load()
    w, h = im.size
    tot = 0.0
    for y in range(h):
        for x in range(w):
            r, g, b = px[x, y]
            mx, mn = max(r, g, b), min(r, g, b)
            tot += 0.0 if mx == 0 else (mx - mn) / mx
    return tot / (w * h)


def review(base_path, frame_paths, derelict_path=None, labels=None):
    """frame_paths must be in BUILD ORDER (emptiest first, completed base last)."""
    labels = labels or [os.path.basename(p) for p in frame_paths]
    findings = []
    metrics = {"frames": []}

    def add(level, frame, msg):
        findings.append({"level": level, "frame": frame, "message": msg})

    # Blankness is a property of the decoded image, not the file size: a byte
    # threshold misjudges anything that is not a full-resolution render.
    for p, lab in zip(frame_paths, labels):
        if not os.path.exists(p):
            add("error", lab, "frame file is missing"); continue
        if os.path.getsize(p) < TRUNCATED_BYTES:
            add("error", lab, "file is %d bytes - truncated or failed write"
                % os.path.getsize(p)); continue
        try:
            sd = _stdev(_load(p))
        except Exception as e:
            add("error", lab, "unreadable image (%s) - blank or corrupt render"
                % e.__class__.__name__); continue
        if sd < BLANK_STDEV:
            add("error", lab, "image is flat (stdev %.2f) - a blank render" % sd)
    if any(f["level"] == "error" for f in findings):
        return {"ok": False, "findings": findings, "metrics": metrics}

    base = _load(base_path)
    diffs = []
    for p, lab in zip(frame_paths, labels):
        im = _load(p)
        if im.size != base.size:
            add("error", lab, "frame size %s != base size %s" % (im.size, base.size))
            continue
        changed, drift = _diff_mask(im, base)
        diffs.append(changed)
        metrics["frames"].append({"frame": lab, "diff_from_base": round(changed, 5),
                                  "palette_drift": round(drift, 2)})
        if drift > PALETTE_TOLERANCE:
            add("warn", lab, "palette drift %.1f in unchanged areas (tolerance %.0f) - "
                             "the edit re-coloured the scene" % (drift, PALETTE_TOLERANCE))

    # The core invariant: diff-from-base must shrink as the build progresses.
    for i in range(len(diffs) - 1):
        a, b = diffs[i], diffs[i + 1]
        la, lb = labels[i], labels[i + 1]
        step = a - b
        if step <= 0:
            add("error", lb,
                "sequence broken: '%s' differs from the completed base by %.1f%%, but the "
                "EARLIER frame '%s' differs by only %.1f%%. Content went backwards - a "
                "removal did not take." % (lb, b * 100, la, a * 100))
        elif step < MIN_STEP:
            add("error", lb,
                "step from '%s' changed only %.2f%% of the frame (minimum %.2f%%) - that "
                "build step added nothing visible." % (la, step * 100, MIN_STEP * 100))
        elif step > MAX_STEP:
            add("warn", lb,
                "step from '%s' changed %.0f%% of the frame - too large for a single "
                "addition; the scene was probably re-drawn." % (la, step * 100))

    if derelict_path and os.path.exists(derelict_path) and frame_paths:
        d_sat = _mean_saturation(_load(derelict_path))
        e_sat = _mean_saturation(_load(frame_paths[0]))
        metrics["derelict_saturation"] = round(d_sat, 4)
        metrics["empty_saturation"] = round(e_sat, 4)
        if e_sat > 0 and d_sat > e_sat * DERELICT_SAT_RATIO:
            add("warn", os.path.basename(derelict_path),
                "derelict frame saturation %.3f is not clearly below the restored frame's "
                "%.3f - it reads dim rather than drained." % (d_sat, e_sat))

    return {"ok": not any(f["level"] == "error" for f in findings),
            "findings": findings, "metrics": metrics}


def format_report(res):
    lines = []
    for m in res["metrics"].get("frames", []):
        lines.append("    %-24s diff_from_base %6.2f%%  drift %5.2f"
                     % (m["frame"], m["diff_from_base"] * 100, m["palette_drift"]))
    if res["findings"]:
        lines.append("")
        for f in res["findings"]:
            lines.append("    [%s] %s: %s" % (f["level"].upper(), f["frame"], f["message"]))
    lines.append("")
    lines.append("    RESULT: " + ("PASS" if res["ok"] else "FAIL"))
    return "\n".join(lines)


def main():
    if len(sys.argv) < 2:
        print(json.dumps({"ok": False, "findings": [
            {"level": "error", "frame": "-", "message": "No payload file given"}]}))
        return
    with open(sys.argv[1], "r", encoding="utf-8-sig") as f:
        p = json.load(f)
    res = review(p["base"], p["frames"], p.get("derelict"), p.get("labels"))
    print(json.dumps(res))
