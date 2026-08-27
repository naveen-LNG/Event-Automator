"""Detect rendered text, captions and numbered legends in generated artwork.

Text is the one defect that makes an asset unusable outright, and it is the one
the generator reproduces most stubbornly: a labelled list in the prompt comes
back as a caption strip, and a numbered list comes back as circled callout
markers plus a key. Both shipped past visual review during the Sort work.

No OCR available in the bundled runtime, so this uses the structural signature
of text rather than reading it: a row of many small, similarly-sized, high-
contrast blobs sitting on a shared baseline. Illustration detail does not
produce that pattern - object shapes vary wildly in size and do not line up.

    detect_text(path) -> {"has_text": bool, "lines": [...], "detail": str}
OCR
---
The structural check above answers "is there text" and nothing more. That was
enough for Sort, where all text is a defect. It is not enough for Merge, where
shop signage is REQUIRED and only the wrong words are a defect - a design name
written on a building ("MUSIC BOX PAVILION" across the music box pavilion) is
structurally identical to an allowed sign ("TEA"). Telling them apart means
reading them.

Worse, the structural check does not even see Merge text. Measured: it returned
zero candidate lines on three Merge scenes carrying eleven correctly spelled
signs, and zero on a 30px bold label composited straight onto a scene. Two
rewrites were tried and both failed under measurement:

  * spatial adjacency chaining instead of global baseline banding - better, but
    no minimum run length separated Merge signage from clean art;
  * a stroke-transition test on merged word-blobs - clean Sort art reached 10.4
    transitions per row against real text at 5.1, so it does not discriminate.

The root cause is that adjacent letters bridge through their anti-aliased edges
when the image is downscaled for analysis, so a word becomes one or two fat
blobs rather than N glyphs. The assumed signature - a row of many small separate
glyphs - simply is not present.

So `read_text()` uses real OCR when it is available, and `detect_text()` prefers
it. On the full ground truth OCR scored 19/19 where the structural check scores
11/11 on Sort alone and 0/3 on Merge.

The backend is VENDORED into pybackend_dist and is imported lazily - the engine
loads once, on first use, and costs about 1.4s; each image then costs 1-2s,
against 35-45s to generate one. It is still an OPTIONAL import: if the packages
are ever missing, every function here keeps working and falls back to the
structural check, which remains correct for Sort and blind on Merge. See
`ocr_available()`.

VENDORING: rapidocr-onnxruntime pulls onnxruntime, opencv, numpy and friends -
89 MB of wheels, and it took pybackend_dist from 75 MB to 313 MB. The embedded
runtime has no pip of its own, so install it with a full CPython 3.14 of the
same version (the dist is 3.14.6) targeting the dist's site-packages:

    <sys-python> -m pip download rapidocr-onnxruntime --only-binary=:all: \\
        --python-version 314 --platform win_amd64 -d <wheeldir>
    <sys-python> -m pip install --no-index --find-links <wheeldir> \\
        --target <dist>/python/Lib/site-packages --upgrade rapidocr-onnxruntime

`--upgrade` matters: without it pip refuses to write into a populated --target.
The wheel set includes Pillow at the version already vendored, so it is
overwritten with itself rather than moved.

electron-builder needs no change - its extraResources copies
`pybackend_dist/python` wholesale, so the packaged app picks these up (and grows
by the same ~238 MB).
"""

import os
import re

MAX_DIM = 700              # analysis resolution; glyphs stay resolvable
CONTRAST = 42              # how far a pixel must sit from local background
MIN_GLYPH_FRAC = 0.008     # glyph height as a fraction of image height
MAX_GLYPH_FRAC = 0.075
MIN_GLYPHS_PER_LINE = 10   # a "line" needs at least this many aligned blobs.
                           # Measured on real output: the actual caption strip ran
                           # 17 blobs, while clean artwork topped out at 7 from
                           # incidentally aligned repeats (mugs on a shelf, tiles).
                           # 10 separates them cleanly. The cost is that a very
                           # short label (under ~10 characters) would slip past;
                           # the failures that actually occur are caption strips
                           # and numbered keys, which are long.
MAX_ASPECT = 2.6           # glyphs are not long thin bars
HEIGHT_SPREAD = 0.42       # glyphs on one line have similar heights


def _components(mask, w, h, max_boxes=4000):
    """Connected components of a bytearray mask, as bounding boxes."""
    from collections import deque
    seen = bytearray(w * h)
    boxes = []
    for sy in range(h):
        row = sy * w
        for sx in range(w):
            if not mask[row + sx] or seen[row + sx]:
                continue
            seen[row + sx] = 1
            q = deque([(sx, sy)])
            x0 = x1 = sx
            y0 = y1 = sy
            n = 0
            while q:
                x, y = q.popleft()
                n += 1
                if x < x0: x0 = x
                if x > x1: x1 = x
                if y < y0: y0 = y
                if y > y1: y1 = y
                for nx, ny in ((x-1, y), (x+1, y), (x, y-1), (x, y+1)):
                    if 0 <= nx < w and 0 <= ny < h:
                        i = ny * w + nx
                        if mask[i] and not seen[i]:
                            seen[i] = 1
                            q.append((nx, ny))
            if n >= 4:
                boxes.append((x0, y0, x1, y1))
            if len(boxes) > max_boxes:      # pathological image; bail out
                return boxes
    return boxes


MIN_OCR_CONF = 0.5         # below this the read is a guess, not a word
_OCR = []                  # one-slot cache: [engine] or [None] once probed


def _ocr_engine():
    """The OCR engine, or None. Loaded once; the model load costs about a second."""
    if _OCR:
        return _OCR[0]
    try:
        from rapidocr_onnxruntime import RapidOCR
        _OCR.append(RapidOCR())
    except Exception:
        _OCR.append(None)
    return _OCR[0]


def ocr_available():
    """Is a real OCR backend present? If not, everything degrades to structure."""
    return _ocr_engine() is not None


TILE_GRID = 2              # tiles per axis in thorough mode
TILE_ZOOM = 2.0            # each tile is magnified before reading
TILE_OVERLAP = 0.18        # so a sign on a tile seam is whole in one of them


def _read_one(eng, image, min_conf, y_offset=0.0, y_scale=1.0):
    try:
        res, _ = eng(image)
    except Exception:
        return []
    out = []
    for r in res or ():
        try:
            box, txt, conf = r[0], str(r[1]).strip(), float(r[2])
        except Exception:
            continue
        if not txt or conf < min_conf:
            continue
        try:
            ys = [pt[1] for pt in box]
            y = y_offset + (sum(ys) / float(len(ys))) / y_scale
        except Exception:
            y = y_offset
        out.append({"text": txt, "conf": round(conf, 2), "y": y})
    return out


def read_text(path, min_conf=MIN_OCR_CONF, thorough=False):
    """Words actually rendered in the image, or None if OCR is unavailable.

    Returns [{"text": str, "conf": float, "y": float}, ...] top to bottom.

    thorough=True also reads the image as magnified overlapping tiles. A single
    whole-image pass misses small, low-contrast signage: a scene that plainly
    carried "GROUT KITCHEN" across a building returned nothing for it even at
    zero confidence, because at full-frame scale the lettering is a few pixels
    tall. Tiling recovered it. Costs roughly 3x the time (about 7s against 2.5s
    on a busy scene), which is cheap next to the 30-45s it took to generate.
    """
    eng = _ocr_engine()
    if eng is None:
        return None

    merged = {}                       # squashed text -> best reading

    def collect(items):
        for it in items:
            k = _squash(it["text"])
            if not k:
                continue
            if k not in merged or it["conf"] > merged[k]["conf"]:
                merged[k] = it

    # The whole-image pass stays even in thorough mode: a wide sign spanning a
    # tile seam reads best when it is not cut in half.
    collect(_read_one(eng, path, min_conf))

    if thorough:
        try:
            import numpy as np
            from PIL import Image
            im = Image.open(path).convert("RGB")
            W, H = im.size
            tw, th = W // TILE_GRID, H // TILE_GRID
            ox, oy = int(tw * TILE_OVERLAP), int(th * TILE_OVERLAP)
            for gy in range(TILE_GRID):
                for gx in range(TILE_GRID):
                    x0, y0 = max(0, gx * tw - ox), max(0, gy * th - oy)
                    x1 = min(W, (gx + 1) * tw + ox)
                    y1 = min(H, (gy + 1) * th + oy)
                    crop = im.crop((x0, y0, x1, y1))
                    crop = crop.resize((int(crop.width * TILE_ZOOM),
                                        int(crop.height * TILE_ZOOM)), Image.LANCZOS)
                    collect(_read_one(eng, np.array(crop), min_conf,
                                      y_offset=y0, y_scale=TILE_ZOOM))
        except Exception:
            pass                      # tiling is an enhancement, never a hard failure

    out = _dedupe_variants(list(merged.values()))
    out.sort(key=lambda d: d["y"])
    return out


def _dedupe_variants(items):
    """Collapse several degraded readings of the same sign into the best one.

    Overlapping tiles read the same lettering more than once, and the copies cut
    by a tile edge come back truncated or one character wrong: one sign produced
    MUSICBOXPAVILION, MUSICBOXPAVILIOI, MUSICBOXPAVILIO and AVILION. Judged
    independently, the damaged copies all read as unauthorised words and the
    scene fails on a sign that is perfectly correct.

    Longest and most confident wins; anything that is a fragment of a kept
    reading, or within a character or two of one, is the same sign again.
    """
    ordered = sorted(items, key=lambda d: (-len(_squash(d["text"])), -d["conf"]))
    kept = []
    for it in ordered:
        s = _squash(it["text"])
        if not s:
            continue
        dup = False
        for k in kept:
            ks = _squash(k["text"])
            if s in ks:                                   # a fragment of it
                dup = True
                break
            # Only treat a near-identical pair as the same sign when the words
            # are long enough for one character to be noise rather than meaning:
            # TEA and SEA are one apart and are two different signs.
            if min(len(s), len(ks)) >= 5 and _edit_distance(s, ks) <= NEAR_MISS_MAX:
                dup = True
                break
        if not dup:
            kept.append(it)
    return kept


def _tokens(s):
    """Comparable word tokens: upper-case, punctuation dropped."""
    return [t for t in re.split(r"[^0-9A-Za-z]+", (s or "").upper()) if t]


def _squash(s):
    """Upper-case with every non-alphanumeric removed.

    OCR is inconsistent about word spacing on stylised signs - the same label
    came back as 'MUSIC BOX PAVILION' on one render and 'MUSICBOXPAVILION' on
    another, and 'DEPT. STORE' reads as 'DEPT.STORE'. Comparing squashed forms
    makes the match independent of how the engine happened to split it.
    """
    return "".join(_tokens(s))


NEAR_MISS_MAX = 2          # edit distance still treated as "probably the same word"


def _edit_distance(a, b, cap=NEAR_MISS_MAX):
    """Levenshtein distance, abandoned once it exceeds cap."""
    if abs(len(a) - len(b)) > cap:
        return cap + 1
    prev = list(range(len(b) + 1))
    for i, ca in enumerate(a, 1):
        cur = [i]
        for j, cb in enumerate(b, 1):
            cur.append(min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + (ca != cb)))
        if min(cur) > cap:
            return cap + 1
        prev = cur
    return prev[-1]


def _near_match(tok, allowed_tokens):
    """The authorised word this is probably a misread of, or None.

    OCR is not perfect on stylised signage: a correctly painted MOSAIC came back
    as 'MOSAIO' and hard-failed a batch. One wrong character is ambiguous between
    "the engine misread it" and "the artwork really is misspelled" - both want a
    human glance and neither should fail the run outright, so these are reported
    separately from words that are plainly not in the brief at all.
    """
    best = None
    for a in allowed_tokens:
        if len(a) < 4:                      # TEA vs SEA - too short to guess at
            continue
        tol = 1 if len(a) < 8 else NEAR_MISS_MAX
        d = _edit_distance(tok, a)
        if d <= tol and (best is None or d < best[0]):
            best = (d, a)
    return best[1] if best else None


def _near_fragment(tok, allowed_squashed):
    """The authorised sign this is probably a damaged fragment of, or None.

    Tiled and overlapping reads produce partial words, and low-contrast
    lettering adds a wrong character on top: a sign correctly reading
    "GLASS HOUSE" - embossed in glass on glass - came back as 'DUSE' and 'OSE'.
    Neither is a substring of GLASSHOUSE, and neither is within one edit of
    GLASS or HOUSE, so both read as unauthorised words and failed the batch.

    Both ARE within one edit of the substring 'OUSE'. So compare the token
    against every substring of the authorised wording at roughly its own
    length. Short tokens need an exact-ish match, which keeps genuinely
    different words - GROUT against TILEWORKSMOSAIC - well clear.
    """
    if len(tok) < 3:
        return None
    tol = 1 if len(tok) < 6 else NEAR_MISS_MAX
    for phrase in allowed_squashed:
        for span in (len(tok) - 1, len(tok), len(tok) + 1):
            if span < 3 or span > len(phrase):
                continue
            for i in range(len(phrase) - span + 1):
                if _edit_distance(tok, phrase[i:i + span]) <= tol:
                    return phrase
    return None


def _segmentable(s, tokens):
    """Can s be split end to end into allowed tokens? Exact, via DP.

    Catches the squashed case where OCR dropped the spaces: 'MUSICBOXPAVILION'
    segments into MUSIC + BOX + PAVILION, so if all three are authorised the
    region is authorised however the engine chose to break it.
    """
    if not s or not tokens:
        return False
    n = len(s)
    reach = [False] * (n + 1)
    reach[0] = True
    for i in range(n):
        if not reach[i]:
            continue
        for t in tokens:
            if t and s.startswith(t, i):
                reach[i + len(t)] = True
    return reach[n]


def check_signage(path, allowed=(), min_conf=MIN_OCR_CONF):
    """Flag rendered words that the brief did not authorise.

    This is the Merge gate. Signage is required there, so "is there text" says
    nothing; what matters is whether the text is text the brief asked for. Pass
    the exact sign wording the brief specifies as `allowed`.

    Matching is by TOKEN, not by whole string, because one sign often reads back
    as several regions - "PAPER BAKERY" on two lines comes back as "PAPER" and
    "BAKERY", and "DEPT. STORE" as "DEPT.STORE". A region passes when every one
    of its tokens appears somewhere in the allowed wording. The known gap: a
    forbidden phrase built entirely from allowed tokens would pass. Design names
    in practice introduce their own words, which is what makes this work.

    Returns {"ok", "unexpected", "found", "engine", "detail"}.
    ok is None - not True - when OCR is unavailable, so a caller can tell
    "checked and clean" from "could not check".

    Reads thoroughly (tiled) by default: this is a gate, and a design name it
    fails to see is worse than a few extra seconds. A whole-image-only pass
    missed a "GROUT KITCHEN" label entirely and reported the scene clean.
    """
    words = read_text(path, min_conf, thorough=True)
    if words is None:
        return {"ok": None, "unexpected": [], "found": [], "engine": "none",
                "detail": "no OCR backend installed, so rendered words cannot be "
                          "checked against the brief (see VENDORING in text_detect)"}

    ok_tokens = set()
    ok_squashed = []
    for phrase in allowed or ():
        ok_tokens.update(_tokens(phrase))
        sq = _squash(phrase)
        if sq:
            ok_squashed.append(sq)

    found, unexpected, near = [], [], []
    for wd in words:
        found.append(wd["text"])
        toks = _tokens(wd["text"])
        if not toks:
            continue
        sq = _squash(wd["text"])
        authorised = (
            all(t in ok_tokens for t in toks)          # every word is authorised
            or any(sq in a for a in ok_squashed)       # a fragment of one sign
            or _segmentable(sq, ok_tokens)             # OCR dropped the spaces
        )
        if authorised:
            continue
        # One character out from an authorised word is ambiguous - misread, or
        # genuinely misspelled art. Report it, but do not fail the run on it.
        guesses = [(t, _near_match(t, ok_tokens) or _near_fragment(t, ok_squashed))
                   for t in toks]
        if all(g for _, g in guesses):
            near.append("%s (did you mean %s?)"
                        % (wd["text"], " ".join(g for _, g in guesses)))
        else:
            unexpected.append(wd["text"])

    if unexpected:
        detail = ("rendered text the brief did not authorise: %s. Allowed wording was: %s. "
                  "A building's design name written on the building is the usual cause."
                  % ("; ".join(repr(u) for u in unexpected[:6]),
                     ", ".join(repr(a) for a in allowed) or "(none)"))
    elif near:
        detail = ("rendered text is one character off the authorised wording: %s. Either the "
                  "reader misread it or the sign really is misspelled - check this one by eye."
                  % "; ".join(near[:4]))
    elif found:
        detail = "all %d rendered text region(s) match the brief's wording" % len(found)
    else:
        detail = "no rendered text found"
    return {"ok": not unexpected, "unexpected": unexpected, "near_misses": near,
            "found": found, "engine": "rapidocr", "detail": detail}


def detect_text(path, min_glyphs=MIN_GLYPHS_PER_LINE):
    # Prefer OCR when it is present: it scored 19/19 on the ground truth against
    # the structural check's 11/11-on-Sort and 0/3-on-Merge, and it reports the
    # words themselves rather than just "something text-shaped is here".
    words = read_text(path)
    if words is not None:
        if words:
            detail = ("rendered text found: %s. Captions, legends and numbered keys make "
                      "the asset unusable." % "; ".join(repr(w["text"])[:40] for w in words[:5]))
        else:
            detail = "no rendered text found"
        return {"has_text": bool(words), "lines": [], "words": [w["text"] for w in words],
                "engine": "rapidocr", "detail": detail}
    return _detect_text_structural(path, min_glyphs)


def _detect_text_structural(path, min_glyphs=MIN_GLYPHS_PER_LINE):
    from PIL import Image, ImageFilter

    im = Image.open(path).convert("RGBA")
    if max(im.size) > MAX_DIM:
        s = MAX_DIM / max(im.size)
        im = im.resize((max(1, int(im.size[0] * s)), max(1, int(im.size[1] * s))), Image.LANCZOS)
    w, h = im.size

    grey = im.convert("L")
    # local background = heavy blur; text is what departs sharply from it
    bg = grey.filter(ImageFilter.GaussianBlur(6))
    g, b = grey.load(), bg.load()
    alpha = im.getchannel("A").load()

    mask = bytearray(w * h)
    for y in range(h):
        row = y * w
        for x in range(w):
            if alpha[x, y] < 128:
                continue
            if abs(g[x, y] - b[x, y]) >= CONTRAST:
                mask[row + x] = 1

    boxes = _components(mask, w, h)

    lo, hi = MIN_GLYPH_FRAC * h, MAX_GLYPH_FRAC * h
    glyphs = []
    for (x0, y0, x1, y1) in boxes:
        bw, bh = x1 - x0 + 1, y1 - y0 + 1
        if lo <= bh <= hi and bw <= bh * MAX_ASPECT and bw >= 2:
            glyphs.append((x0, y0, x1, y1, bh))

    # group glyphs into baselines: similar bottom edge, similar height
    glyphs.sort(key=lambda g_: g_[3])
    lines, used = [], [False] * len(glyphs)
    for i, gi in enumerate(glyphs):
        if used[i]:
            continue
        band = [gi]
        used[i] = True
        for j in range(i + 1, len(glyphs)):
            if used[j]:
                continue
            gj = glyphs[j]
            if abs(gj[3] - gi[3]) > max(2, gi[4] * 0.5):
                continue                                   # different baseline
            if abs(gj[4] - gi[4]) / max(1, gi[4]) > HEIGHT_SPREAD:
                continue                                   # different size
            band.append(gj)
            used[j] = True
        if len(band) >= min_glyphs:
            band.sort(key=lambda g_: g_[0])
            # a real text line is a run of blobs with small, regular gaps
            gaps = [band[k + 1][0] - band[k][2] for k in range(len(band) - 1)]
            tight = [gp for gp in gaps if 0 <= gp <= band[0][4] * 1.6]
            if len(tight) >= min_glyphs - 2:
                lines.append({"count": len(band),
                              "y": round(band[0][3] / h, 3),
                              "x0": round(band[0][0] / w, 3),
                              "x1": round(band[-1][2] / w, 3),
                              "glyph_h": round(band[0][4] / h, 4)})

    lines.sort(key=lambda l: -l["count"])
    if lines:
        worst = lines[0]
        where = ("below the artwork" if worst["y"] > 0.82 else
                 "above the artwork" if worst["y"] < 0.18 else "across the artwork")
        detail = ("looks like rendered text: %d text-like line(s), the longest %d characters "
                  "%s at y=%.2f. Captions, legends and numbered keys make the asset unusable."
                  % (len(lines), worst["count"], where, worst["y"]))
    else:
        detail = "no text-like structure found"
    return {"has_text": bool(lines), "lines": lines[:4], "detail": detail}
