"""Normalise a generated merge-chain item into a shippable icon.

A merge chain is 5 levels of one product family, and the SHIPPED Merge2608 icons
(ResourcesAb/Events/Merge2608/Icon/Element1) set the convention precisely:

    canvas          180x180, transparent
    content         one object, centred
    size progression MONOTONIC across the chain - measured alpha coverage runs
                    ~10-20% at level 1 up to ~40-56% at level 5, and the content
                    bounding box grows with it

That progression is not decoration. Mattel's Merge2604 feedback calls it out as
missing, in the same breath as the other two things AI-generated icons arrive
without:

    目前全部图标还缺少统一的描边、投影、图标大小进阶的处理。
    （如果使用chatgpt生成了无描边的图标，需要额外添加1-2像素的描边效果）
    "All icons lack unified STROKE, DROP SHADOW and icon SIZE PROGRESSION. If you
     used ChatGPT to generate icons without strokes, add a 1-2 pixel stroke."

All three are handled here rather than asked of the generator, because all three
are geometry and geometry is exactly what a prompt is worst at. The generator
draws the object; this module places it at the right size on the right canvas and
gives it the stroke the art team asked for.

    normalise(src, dest, level, levels=5) -> dict of what it did
"""

import os

CANVAS = 180                   # px, matching the shipped icons exactly
STROKE_PX = 2                  # Mattel: "add 1-2 pixel stroke"
STROKE_DARKEN = 0.45           # the rim is this fraction of the object's own colour

# Longest content edge as a fraction of the canvas, per level. Read off the
# shipped Element1 icons: level 1 sits around 0.62 and level 5 around 0.93.
SIZE_STEPS = (0.60, 0.70, 0.78, 0.86, 0.93)

# PRODUCERS are not chain rungs and do not step up in size - they are the machine,
# drawn large at both levels. Measured off the shipped element_machine<g>_<lv>
# icons: level 1 longest edge averages 164px of 180 (0.91), level 2 averages
# 170px (0.95). A producer drawn at a chain rung's size reads as an ingredient.
PRODUCER_STEPS = (0.91, 0.95)
PRODUCER_MIN = 0.85            # below this it is not reading as a producer

# DESSERT recipes and ingredients are a flat set, not a ladder - every dish is
# drawn at the same scale as every other, and the two classes differ only in
# canvas and fill. All four numbers are measured off the shipped Dessert2608
# icons, not chosen:
#   Icon/Menu1..4      24 common recipes, 342x342, longest edge 0.68-0.86, mean 0.78
#   Icon/Menu_sp_1..3  15 secret recipes, 342x342, same treatment
#   Icon/Material      11 ingredients,    224x224, longest edge 0.84-0.94, mean 0.90
# An ingredient fills more of its canvas than a recipe does, which is what keeps
# a single ingredient reading as a raw component rather than as a small dish.
RECIPE_CANVAS = 342
RECIPE_FRACTION = 0.78
INGREDIENT_CANVAS = 224
INGREDIENT_FRACTION = 0.90


def stroke_for(canvas):
    """The rim, scaled so it stays the same weight relative to the icon.

    STROKE_PX is 2px on a 180px canvas. Left literal, a 342px recipe would get a
    hairline a third of the shipped weight, which is why this is a ratio.
    """
    return max(1, int(round(STROKE_PX * float(canvas) / CANVAS)))


def size_fraction(level, levels=5):
    """Target size for this rung of the chain. Always increasing."""
    if levels <= 1:
        return SIZE_STEPS[-1]
    lo, hi = SIZE_STEPS[0], SIZE_STEPS[-1]
    if levels == len(SIZE_STEPS):
        return SIZE_STEPS[max(0, min(len(SIZE_STEPS) - 1, level - 1))]
    t = (level - 1) / float(levels - 1)
    return lo + (hi - lo) * t


def add_stroke(im, px=STROKE_PX, darken=STROKE_DARKEN):
    """A uniform rim around the opaque shape, in a DARKENED version of the
    object's own colour at each point.

    The colour matters and a flat white rim is wrong. Measured on the shipped
    Merge2608 icons, scoop.edge_halo_score runs -70 to -145 (mean -101): their
    boundary ring is much DARKER than the interior. A white rim scores about
    +116 - the opposite. So the convention is a dark keyline that follows the
    local colour, which is also what a hand-drawn game icon does.

    Built by bleeding the object's colour outward with an alpha-weighted blur
    and multiplying it down. Alpha weighting matters for the same reason it does
    in scoop.dehalo: a plain blur drags transparent black into the edge and
    muddies it.
    """
    from PIL import Image, ImageFilter
    if px <= 0:
        return im
    px = int(px)
    a = im.getchannel("A")
    grown = a.filter(ImageFilter.MaxFilter(2 * px + 1))

    w, h = im.size
    ap = a.load()
    rgb_px = im.convert("RGB").load()
    pm = Image.new("RGB", (w, h))
    pmp = pm.load()
    for y in range(h):
        for x in range(w):
            av = ap[x, y]
            r, g, b = rgb_px[x, y]
            pmp[x, y] = (r * av // 255, g * av // 255, b * av // 255)

    radius = px + 1
    pm_b = pm.filter(ImageFilter.GaussianBlur(radius)).load()
    a_b = a.filter(ImageFilter.GaussianBlur(radius)).load()

    ring_rgb = Image.new("RGB", (w, h))
    rp = ring_rgb.load()
    for y in range(h):
        for x in range(w):
            wgt = a_b[x, y]
            if wgt <= 0:
                rp[x, y] = (0, 0, 0)
                continue
            r, g, b = pm_b[x, y]
            rp[x, y] = (min(255, int(r * 255 / wgt * darken)),
                        min(255, int(g * 255 / wgt * darken)),
                        min(255, int(b * 255 / wgt * darken)))

    ring = ring_rgb.convert("RGBA")
    ring.putalpha(grown)
    return Image.alpha_composite(ring, im)


def restroke(path, dest=None, px=STROKE_PX, darken=STROKE_DARKEN):
    """Replace an existing rim: erode it away, then draw the correct one.

    Needed because the first pass baked a white rim into 40 finished icons.
    Eroding by the stroke width removes the ring without touching the artwork
    inside it, which is far cheaper than re-rendering the set.
    """
    from PIL import Image, ImageFilter
    dest = dest or path
    im = Image.open(path).convert("RGBA")
    a = im.getchannel("A")
    im.putalpha(a.filter(ImageFilter.MinFilter(2 * int(px) + 1)))
    out = add_stroke(im, px, darken)
    out.save(dest)
    return dest


def normalise(src, dest=None, level=1, levels=5, canvas=CANVAS, stroke=STROKE_PX,
              fraction=None):
    """Trim, scale to this level's size, centre on a transparent square canvas.

    The source must already be cut out (transparent background) - see
    scoop.knockout_background. Returns what was done, so a caller can check.

    fraction overrides the chain-rung size ladder, for assets that are not chain
    rungs - pass PRODUCER_STEPS[lv-1] for a producer.
    """
    from PIL import Image
    dest = dest or src
    im = Image.open(src).convert("RGBA")
    bb = im.getbbox()
    if not bb:
        raise ValueError("image is fully transparent: %s" % os.path.basename(src))
    im = im.crop(bb)

    frac = fraction if fraction is not None else size_fraction(level, levels)
    target = max(1, int(round(canvas * frac)))
    scale = target / float(max(im.size))
    new = (max(1, int(round(im.width * scale))), max(1, int(round(im.height * scale))))
    im = im.resize(new, Image.LANCZOS)

    if stroke:
        # Stroke before placing, then re-fit: the rim adds a couple of pixels and
        # would otherwise push the content past its size target.
        im = add_stroke(im, stroke)
        bb2 = im.getbbox()
        if bb2:
            im = im.crop(bb2)
        if max(im.size) > target:
            s = target / float(max(im.size))
            im = im.resize((max(1, int(im.width * s)), max(1, int(im.height * s))), Image.LANCZOS)

    out = Image.new("RGBA", (canvas, canvas), (0, 0, 0, 0))
    out.alpha_composite(im, ((canvas - im.width) // 2, (canvas - im.height) // 2))
    os.makedirs(os.path.dirname(os.path.abspath(dest)), exist_ok=True)
    out.save(dest)
    return {"canvas": canvas, "level": level, "size_fraction": round(frac, 3),
            "content": im.size, "coverage": round(coverage(dest), 3)}


def coverage(path, alpha_min=200):
    """Fraction of the canvas that is solidly opaque - the size measure."""
    from PIL import Image
    im = Image.open(path).convert("RGBA")
    h = im.getchannel("A").histogram()
    return sum(h[alpha_min:]) / float(im.width * im.height)


STEP_DROP_TOL = 0.05           # of canvas: a step may dip this much and still count
CHAIN_GROWTH = 1.30            # level 5 must be at least this much bigger than level 1


def content_size(path):
    """Longest edge of the drawn content, in pixels."""
    from PIL import Image
    bb = Image.open(path).convert("RGBA").getbbox()
    return 0 if not bb else max(bb[2] - bb[0], bb[3] - bb[1])


def check_chain(paths, canvas=CANVAS):
    """Does this chain's size actually progress across the set?

    Mattel asked for "图标大小进阶" - size progression - as a property of the SET,
    so it can only be checked across the set, never on one icon.

    CALIBRATED AGAINST THE SHIPPED MERGE2608 CHAINS, and the first rule tried was
    wrong. Strict monotonic AREA fails three of the eight shipped chains: the
    men's formalwear chain ends bow tie -> shoes -> trousers -> jacket -> full
    suit, and the suit is taller but narrower than the spread jacket, so its area
    DROPS at the last step while it is plainly the bigger item. Longest-edge is
    the better measure but still dips by 2px and 4px in two shipped chains.

    So the real convention is "clearly bigger by the end, never a real step
    backwards" - a small dip is fine, a collapse is not. Both bounds below are
    what the shipped set actually satisfies.
    """
    sizes = [content_size(p) for p in paths]
    covs = [coverage(p) for p in paths]
    findings = []
    tol = canvas * STEP_DROP_TOL
    drops = [i for i in range(1, len(sizes)) if sizes[i] < sizes[i - 1] - tol]
    if drops:
        findings.append("size steps backwards at level(s) %s - longest edge %s"
                        % ([i + 1 for i in drops], sizes))
    if len(sizes) >= 2 and sizes[0] > 0 and sizes[-1] < sizes[0] * CHAIN_GROWTH:
        findings.append("the chain barely grows: level 1 is %dpx and the last is %dpx, "
                        "under the %.2fx the shipped chains show" % (sizes[0], sizes[-1], CHAIN_GROWTH))
    from PIL import Image
    square = [Image.open(p).size for p in paths if Image.open(p).size[0] != Image.open(p).size[1]]
    if square:
        findings.append("not every icon is a square canvas: %s" % square)
    return {"ok": not findings, "sizes": sizes,
            "coverage": [round(c, 3) for c in covs], "findings": findings}


def check_producers(paths, canvas=CANVAS):
    """A producer must read as the machine, not as one of its ingredients.

    The one rule worth enforcing is size: the shipped producer icons all sit at
    0.86-0.97 of the canvas, while a chain's first rung sits at 0.60. A producer
    drawn small is the failure mode, and it is the one that is measurable.
    """
    sizes = [content_size(p) for p in paths]
    fr = [s / float(canvas) for s in sizes]
    findings = []
    small = [i for i, f in enumerate(fr) if f < PRODUCER_MIN]
    if small:
        findings.append("producer(s) %s are drawn too small - %s of the canvas, under %.2f"
                        % ([i + 1 for i in small], [round(f, 2) for f in fr], PRODUCER_MIN))
    from PIL import Image
    bad = [Image.open(p).size for p in paths if Image.open(p).size != (canvas, canvas)]
    if bad:
        findings.append("not every producer is a %dx%d canvas: %s" % (canvas, canvas, bad))
    return {"ok": not findings, "sizes": sizes,
            "fractions": [round(f, 3) for f in fr], "findings": findings}
