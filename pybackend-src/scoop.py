"""Cut the UI scoop out of a generated Sort zone floor.

Two of the five shipped Sort2609 floors (aquarium, reptile desert) have a curved
bite taken out of the top of the frame for the game UI. It is ALPHA TRANSPARENCY
in the PNG, not painted background - measured at 49-53% transparent in the top
band, versus 0% for the cat room and aviary.

So it cannot be prompted for: an image model returns a filled rectangle. The
scoop is a post-step, applied here.

The curve is not invented. It was measured off the shipped assets by walking
each column down to the first opaque pixel:

    reptile  y/h across width: 0.00 .029 .055 .083 .110 .136 .156 .168 .173 ...
    aquarium y/h across width: 0.00 .033 .058 .083 .107 .129 .137 .154 .152 ...

Zero at both edges, peaking near 0.17 of the image height at centre - a
downward sine arc, which DEFAULT_DEPTH and the sine below reproduce.
"""

import math
import os

DEFAULT_DEPTH = 0.17       # peak depth as a fraction of image height
FEATHER_PX = 2             # soft edge so the cut does not alias


def scoop_profile(width, height, depth=DEFAULT_DEPTH):
    """y (in pixels) of the art's top boundary, per column."""
    return [depth * height * math.sin(math.pi * (x / max(1, width - 1)))
            for x in range(width)]


def apply_top_scoop(path, dest=None, depth=DEFAULT_DEPTH, feather=FEATHER_PX):
    """Make the scoop region transparent. Returns the written path.

    Idempotent in practice: re-running on an already-scooped image just clears
    pixels that are already clear.
    """
    from PIL import Image

    im = Image.open(path).convert("RGBA")
    w, h = im.size
    # Image.split() hands back COPIES of the bands - writing into one of those
    # never reaches the image, which silently produced a no-op. Take the band
    # out, edit it, and put it back.
    band = im.getchannel("A")
    alpha = band.load()
    prof = scoop_profile(w, h, depth)

    for x in range(w):
        cut = int(prof[x])
        for y in range(min(cut, h)):
            alpha[x, y] = 0
        # feather the boundary so the curve does not stair-step
        for i in range(feather):
            y = cut + i
            if 0 <= y < h:
                frac = (i + 1) / (feather + 1)
                alpha[x, y] = int(alpha[x, y] * frac)

    im.putalpha(band)
    out = dest or path
    os.makedirs(os.path.dirname(os.path.abspath(out)), exist_ok=True)
    im.save(out)
    return out


BG_LOCAL_TOL = 12          # step-to-step difference allowed while spreading
BG_GLOBAL_TOL = 78         # hard rail: how far the fill may drift from the seed
BG_TOLERANCE = BG_LOCAL_TOL   # back-compat alias
BG_FEATHER = 2
HALO_TOLERANCE = 20        # |edge - interior| lightness below this is already clean
HALO_HEAVY = 60            # above this, erode harder


def knockout_background(path, dest=None, local_tol=BG_LOCAL_TOL,
                        global_tol=BG_GLOBAL_TOL, feather=BG_FEATHER):
    """Make the backdrop outside the floor plate transparent.

    The shipped Sort floors are cut-out assets: 4-17% of each is transparent,
    everything outside the plate and its walls. Generated renders come back on
    a solid backdrop instead, so the scoop alone still leaves an opaque
    rectangle. This is what turns a picture into an asset.

    Two properties matter:

    * It floods inward FROM THE BORDER ONLY, so an interior region that happens
      to match the backdrop colour (a white wall, a pale tile) is never eaten -
      it simply is not reachable from the edge.

    * It compares each pixel to THE NEIGHBOUR IT SPREAD FROM, not to a fixed
      seed colour. A fixed seed only clears a perfectly flat backdrop; the
      generator often returns a gradient sky, which a flat comparison abandons
      partway up and leaves half the backdrop opaque. Local comparison walks a
      smooth gradient all the way. `global_tol` is the rail that stops that walk
      wandering out of the backdrop and into the art.
    """
    from collections import deque
    from PIL import Image

    im = Image.open(path).convert("RGBA")
    w, h = im.size
    px = im.load()
    band = im.getchannel("A")
    alpha = band.load()

    # Seed: median of the four corners, so one odd corner cannot define it.
    corners = [px[0, 0], px[w - 1, 0], px[0, h - 1], px[w - 1, h - 1]]
    seed = tuple(sorted(c[i] for c in corners)[1] for i in range(3))

    # The global rail is derived from the backdrop's OWN observed range along
    # the border rather than being a fixed number. A fixed rail has to guess how
    # steep a gradient might be: 78 per channel silently refused to finish a sky
    # spanning 110. The border is backdrop by definition, so its own spread is
    # the honest bound, widened a little for noise.
    edge = ([px[x, 0] for x in range(0, w, 4)] + [px[x, h - 1] for x in range(0, w, 4)] +
            [px[0, y] for y in range(0, h, 4)] + [px[w - 1, y] for y in range(0, h, 4)])
    edge = [c for c in edge if c[3] > 0] or [seed + (255,)]
    lo = [min(c[i] for c in edge) - global_tol // 3 for i in range(3)]
    hi = [max(c[i] for c in edge) + global_tol // 3 for i in range(3)]

    def near(a, b, tol):
        return (abs(a[0] - b[0]) <= tol and abs(a[1] - b[1]) <= tol
                and abs(a[2] - b[2]) <= tol)

    def in_backdrop_range(c):
        return all(lo[i] <= c[i] <= hi[i] for i in range(3))

    seen = bytearray(w * h)
    q = deque()

    def seed_px(x, y):
        if seen[y * w + x]:
            return
        c = px[x, y]
        if c[3] == 0 or in_backdrop_range(c):
            seen[y * w + x] = 1
            q.append((x, y))

    for x in range(w):
        seed_px(x, 0); seed_px(x, h - 1)
    for y in range(h):
        seed_px(0, y); seed_px(w - 1, y)

    while q:
        x, y = q.popleft()
        here = px[x, y]
        alpha[x, y] = 0
        for nx, ny in ((x - 1, y), (x + 1, y), (x, y - 1), (x, y + 1)):
            if not (0 <= nx < w and 0 <= ny < h) or seen[ny * w + nx]:
                continue
            c = px[nx, ny]
            # local step keeps a gradient connected; the global rail keeps the
            # fill from drifting shade-by-shade into the artwork.
            if c[3] == 0 or (near(c, here, local_tol) and in_backdrop_range(c)):
                seen[ny * w + nx] = 1
                q.append((nx, ny))

    im.putalpha(band)
    if feather:
        # one soft pass so the cut edge is not a hard jaggy staircase
        from PIL import ImageFilter
        soft = im.getchannel("A").filter(ImageFilter.GaussianBlur(feather * 0.5))
        im.putalpha(soft)

    out = dest or path
    os.makedirs(os.path.dirname(os.path.abspath(out)), exist_ok=True)
    im.save(out)
    return out


NEUTRAL_TOL = 8            # max-min channel spread still counted as "colourless"
NEUTRAL_BAND = (105, 180)  # value range of the flat backdrop these renders use


def knockout_enclosed(path, dest=None, tol=NEUTRAL_TOL, band=NEUTRAL_BAND):
    """Remove leftover backdrop that the border flood could not reach.

    knockout_background floods inward from the border, so background ENCLOSED by
    the artwork survives: the hole in a rope grommet, the mesh of a net, the
    middle of a necklace loop or a wreath. Six of forty generated icons came back
    with a grey disc sitting in the hole.

    Those leftovers are the flat backdrop, and the backdrop is PERFECTLY neutral -
    measured at (146,146,146), (134,134,134), (155,155,155), saturation 0.00 -
    while the artwork carries a tint even where it looks grey: the cream rope in
    the same icons reads saturation 0.32. So neutrality inside a mid value band
    identifies the backdrop and leaves the art alone.

    This is deliberately narrow. It is for renders made against a flat neutral
    backdrop, which is what the icon prompt asks for; it is not a general matte.
    """
    from PIL import Image
    dest = dest or path
    im = Image.open(path).convert("RGBA")
    px = im.load()
    w, h = im.size
    lo, hi = band
    removed = 0
    for y in range(h):
        for x in range(w):
            r, g, b, a = px[x, y]
            if a == 0:
                continue
            mx, mn = max(r, g, b), min(r, g, b)
            if mx - mn <= tol and lo <= mx <= hi:
                px[x, y] = (r, g, b, 0)
                removed += 1
    im.save(dest)
    return {"removed_px": removed, "fraction": round(removed / float(w * h), 4)}


def dehalo(path, dest=None, passes=2):
    """Remove the pale fringe left around cut-out edges.

    Mattel's art feedback, verbatim: "AI-generated resources will have obvious
    white edges. We need to use the mask tool in Photoshop to cut out these
    white lines to make the edges look clean overall."

    Two causes, both handled here:
      * the render itself paints a light halo where an object meets the
        backdrop, so the outermost opaque ring is lighter than the art behind it
      * an alpha feather leaves partly-transparent pixels still carrying that
        light colour, which composite as a white rim over any background

    Erodes the alpha by a pixel or two at the boundary and pulls the remaining
    edge colour inward from fully-opaque neighbours, so the cut edge takes the
    object's own colour instead of the backdrop's.
    """
    from PIL import Image, ImageFilter

    # Measure before acting. An edge that is already clean does not need
    # eroding, and correcting it anyway overshoots: one test image sat at +1.6
    # and came back at -68.7 after a blind two-pass erode. Only images with a
    # real fringe get touched, and badly-fringed ones get more passes.
    # The defect is DIRECTIONAL: Mattel flagged pale edges, so only a POSITIVE
    # score is a fault. This guard used abs() and that was wrong - a negative
    # score means the boundary ring is darker than the interior, which is the
    # contact shadow every one of these plates is supposed to have. Measured on
    # the shipped, Mattel-approved Sort floors: -19.4, -13.9, +1.2, 0.0, -19.4.
    # Eroding a dark rim also cannot converge - each pass peels the ring back
    # deeper into the shadow, so the next ring is darker still. One floor went
    # -20.0 -> -43.4 being "corrected".
    score = edge_halo_score(path)
    if score < HALO_TOLERANCE:
        return dest or path
    passes = 1 if score < HALO_HEAVY else max(1, passes)

    im = Image.open(path).convert("RGBA")
    w, h = im.size
    for _ in range(max(1, passes)):
        a = im.getchannel("A")
        # MinFilter erodes the opaque region by one pixel per pass
        eroded = a.filter(ImageFilter.MinFilter(3))
        im.putalpha(eroded)

    # Bleed interior colour outward with an ALPHA-WEIGHTED blur, so an edge
    # pixel takes the colour of the opaque art beside it.
    #
    # A plain MaxFilter was tried first and made things worse - it picks the
    # brightest neighbour, so it lightens exactly the fringe it is supposed to
    # remove (one test image went from +1.6 to +12.5). Premultiplying by alpha,
    # blurring, then dividing back weights each neighbour by how opaque it is,
    # which pulls in real colour and ignores the transparent side entirely.
    a = im.getchannel("A")
    ap = a.load()
    rgb = im.convert("RGB").load()

    pm = Image.new("RGB", (w, h))
    pmp = pm.load()
    for y in range(h):
        for x in range(w):
            av = ap[x, y]
            r, g, b = rgb[x, y]
            pmp[x, y] = (r * av // 255, g * av // 255, b * av // 255)

    R = 2
    pm_b = pm.filter(ImageFilter.GaussianBlur(R)).load()
    a_b = a.filter(ImageFilter.GaussianBlur(R)).load()

    out = Image.new("RGBA", (w, h))
    op = out.load()
    for y in range(h):
        for x in range(w):
            av = ap[x, y]
            if av == 0:
                op[x, y] = (0, 0, 0, 0)
            elif av < 250:
                wsum = a_b[x, y]
                if wsum > 4:
                    pr, pg, pb = pm_b[x, y]
                    op[x, y] = (min(255, pr * 255 // wsum),
                                min(255, pg * 255 // wsum),
                                min(255, pb * 255 // wsum), av)
                else:
                    op[x, y] = rgb[x, y] + (av,)
            else:
                op[x, y] = rgb[x, y] + (av,)

    target = dest or path
    out.save(target)
    return target


def edge_halo_score(path, sample=4):
    """Mean lightness of the semi-transparent edge ring minus that of the
    interior. A positive number means a pale fringe - the defect Mattel flagged."""
    from PIL import Image
    im = Image.open(path).convert("RGBA")
    w, h = im.size
    px = im.load()
    edge, inner = [], []
    for y in range(0, h, sample):
        for x in range(0, w, sample):
            r, g, b, a = px[x, y]
            if a == 0:
                continue
            lum = (r + g + b) / 3.0
            (edge if a < 250 else inner).append(lum)
    if not edge or not inner:
        return 0.0
    return sum(edge) / len(edge) - sum(inner) / len(inner)


def trim_to_content(path, dest=None, margin_frac=0.01):
    """Crop to the opaque bounding box, leaving a small even margin.

    After knockout the plate floats in a large transparent field - 41-45% of
    the frame, against 4-17% for the shipped assets, which are framed tightly
    to the plate. Trimming brings the framing (and the transparency ratio) into
    line, and stops the game scaling a mostly-empty texture.
    """
    from PIL import Image
    im = Image.open(path).convert("RGBA")
    bbox = im.getchannel("A").point(lambda v: 255 if v > 8 else 0).getbbox()
    if not bbox:
        return dest or path
    m = int(min(im.size) * margin_frac)
    box = (max(0, bbox[0] - m), max(0, bbox[1] - m),
           min(im.size[0], bbox[2] + m), min(im.size[1], bbox[3] + m))
    out = dest or path
    im.crop(box).save(out)
    return out


def cut_asset(path, dest=None, scoop=False, depth=DEFAULT_DEPTH, trim=True, halo=True):
    """Full post-process, in the order that matters.

    knockout -> dehalo -> trim -> scoop.

    dehalo runs straight after the knockout, while the fringe is still the
    outermost ring of the cut: Mattel's first note is that AI output shows
    obvious white edges that have to be masked away by hand. Trim comes after
    so it measures the cleaned silhouette, and the scoop is cut last so its
    curve is measured against the FINAL asset height.
    """
    out = knockout_background(path, dest)
    if halo:
        dehalo(out)
    if trim:
        trim_to_content(out)
    if scoop:
        apply_top_scoop(out, depth=depth)
    return out


def transparency(path):
    """Fraction of the image that is fully transparent."""
    from PIL import Image
    im = Image.open(path)
    if im.mode != "RGBA":
        return 0.0
    a = im.getchannel("A")
    return sum(1 for v in a.get_flattened_data() if v < 16) / (im.size[0] * im.size[1])


def measure_scoop(path):
    """Depth of a top scoop, as a fraction of height. 0.0 if there is none.

    Measured as the CENTRE column's transparent run minus the shallowest run
    across the width, because raw depth cannot tell a scoop from the shape of
    the asset. After trimming, an isometric plate's transparent corners give
    edge columns a deep run and naive depth reported 0.74 on an unscooped
    image. A scoop is deepest at centre; a diamond plate is shallowest there,
    so the difference separates them cleanly and reads 0 for a plain plate.
    """
    from PIL import Image

    im = Image.open(path)
    if im.mode != "RGBA":
        return 0.0
    w, h = im.size
    a = im.getchannel("A").load()

    def run(x):
        y = 0
        while y < h and a[x, y] < 128:
            y += 1
        return y if y < h else 0       # a fully transparent column tells us nothing

    runs = [run(x) for x in range(0, w, max(1, w // 64))]
    if not runs:
        return 0.0
    centre = run(w // 2)
    return max(0.0, (centre - min(runs)) / h)
