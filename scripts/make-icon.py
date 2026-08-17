# Builds the Event Automator app icon from the in-game SKIP-BO wordmark.
#
# The source (img_num_skipbo.png) is a pure-white, alpha-masked wordmark. It is
# recoloured to match the app's own brand mark in the sidebar:
#   background: linear-gradient(140deg, var(--accent-2) #9d6bff, var(--accent-1) #4f7dff)
#   inset 0 1px 0 rgba(255,255,255,0.45)      -> top highlight
#   1px solid rgba(255,255,255,0.22) inset 3px -> hairline keyline
# so the taskbar icon and the in-app emblem read as the same object.
import math, os
from PIL import Image, ImageDraw, ImageFilter

SRC = r"C:\SkipBo\branches\Misc\develop\client\Skipbo\Assets\ResourcesAb\Icon\Card_Alice\img_num_skipbo.png"
OUT = r"C:\SkipBo\tools\EventAutomationDesktop\build"
os.makedirs(OUT, exist_ok=True)

ACCENT_2 = (157, 107, 255)   # --accent-2 violet  (gradient start)
ACCENT_1 = (79, 125, 255)    # --accent-1 blue    (gradient end)
ANGLE_DEG = 140              # matches the CSS brand-mark gradient

SS = 4                       # supersample factor for clean curves


def linear_gradient(size, c0, c1, angle_deg):
    """CSS-style linear-gradient: 0deg = to top, angles run clockwise."""
    w = h = size
    # CSS angle -> direction vector in screen space (y grows downward)
    rad = math.radians(angle_deg)
    dx, dy = math.sin(rad), -math.cos(rad)
    # project every pixel onto the gradient axis, normalised to 0..1
    half = (abs(dx) + abs(dy)) / 2.0
    img = Image.new("RGB", (w, h))
    px = img.load()
    for y in range(h):
        ny = (y + 0.5) / h - 0.5
        for x in range(w):
            nx = (x + 0.5) / w - 0.5
            t = (nx * dx + ny * dy) / (2 * half) + 0.5
            t = 0.0 if t < 0 else 1.0 if t > 1 else t
            px[x, y] = (
                round(c0[0] + (c1[0] - c0[0]) * t),
                round(c0[1] + (c1[1] - c0[1]) * t),
                round(c0[2] + (c1[2] - c0[2]) * t),
            )
    return img


def tinted(mask, rgb, alpha_scale=1.0, canvas=None, at=(0, 0)):
    """Straight-alpha RGBA: solid `rgb` everywhere, shape carried by `mask`.

    Image.paste(src, box, mask) multiplies the COLOUR channels by the mask as
    well as the alpha, leaving edge pixels premultiplied (a white glyph turns
    grey at the edges). alpha_composite then multiplies by alpha a second time,
    which shows up as a dark fringe. Building the channels explicitly keeps the
    colour flat and the shape purely in alpha.
    """
    a = mask if alpha_scale == 1.0 else mask.point(lambda v: int(v * alpha_scale))
    w, h = mask.size
    solid = Image.merge("RGBA", (
        Image.new("L", (w, h), rgb[0]),
        Image.new("L", (w, h), rgb[1]),
        Image.new("L", (w, h), rgb[2]),
        a))
    if canvas is None:
        return solid
    out = Image.new("RGBA", canvas, (0, 0, 0, 0))
    out.paste(solid, at)          # no mask -> channels copied verbatim
    return out


def rounded_mask(size, radius):
    m = Image.new("L", (size, size), 0)
    ImageDraw.Draw(m).rounded_rectangle([0, 0, size - 1, size - 1], radius=radius, fill=255)
    return m


def render(px_size, glyph_src):
    """Render one icon at px_size, supersampled then downscaled."""
    S = px_size * SS
    radius = int(S * 0.225)                       # squircle-ish tile

    tile = linear_gradient(S, ACCENT_2, ACCENT_1, ANGLE_DEG).convert("RGBA")

    # Depth: darken toward the bottom-right so the tile doesn't read flat.
    sd = Image.new("L", (S, S), 0)
    sdd = ImageDraw.Draw(sd)
    # i counts UP from the bottom row, so the strength must DECAY with i —
    # otherwise the shading peaks mid-tile and leaves a hard seam where the
    # band stops.
    band_d = int(S * 0.55)
    for i in range(band_d):
        v = int(58 * (1 - i / band_d) ** 1.5)
        sdd.line([(0, S - 1 - i), (S, S - 1 - i)], fill=v)
    tile = Image.composite(Image.new("RGBA", (S, S), (14, 18, 44, 255)), tile, sd)

    # Top sheen — the CSS is inset 0 1px 0 rgba(255,255,255,0.45): a thin
    # catch-light, not a wash over the whole upper half.
    hl = Image.new("L", (S, S), 0)
    hd = ImageDraw.Draw(hl)
    band = int(S * 0.13)
    for i in range(band):
        v = int(70 * (1 - i / band) ** 2.6)
        hd.line([(0, i), (S, i)], fill=v)
    tile = Image.composite(Image.new("RGBA", (S, S), (255, 255, 255, 255)), tile, hl)

    # hairline keyline, inset like .brand-mark::after (rgba(255,255,255,0.22)).
    # Below ~32px it costs more pixels than it earns, so the small entries drop
    # it and give that room back to the wordmark instead.
    if px_size >= 32:
        kd = ImageDraw.Draw(tile)
        inset = max(1, int(S * 0.072))
        kd.rounded_rectangle(
            [inset, inset, S - 1 - inset, S - 1 - inset],
            radius=int(radius * 0.72), outline=(255, 255, 255, 46), width=max(1, int(S / 150)))

    tile.putalpha(rounded_mask(S, radius))

    # --- the wordmark ---
    # Small icons get less padding so the strokes survive downscaling.
    pad = 0.165 if px_size >= 64 else 0.105 if px_size >= 32 else 0.055
    box = int(S * (1 - 2 * pad))
    g = glyph_src.resize((box, box), Image.LANCZOS)
    # The source wordmark is only 60px, so this is a ~20x upscale and LANCZOS
    # leaves a wide feathered edge. Re-steepen the alpha ramp to recover a
    # crisp outline while keeping a narrow antialiased band.
    g = g.point([0 if i < 104 else 255 if i > 152 else int((i - 104) * 255 / 48)
                 for i in range(256)])

    # tight contact shadow so white-on-violet keeps its edge
    sh = tinted(g, (8, 10, 26), 120 / 255, canvas=(S, S),
                at=(int(S * pad), int(S * pad) + max(1, int(S / 130))))
    sh = sh.filter(ImageFilter.GaussianBlur(S / 200))
    tile = Image.alpha_composite(tile, sh)

    layer = tinted(g, (255, 255, 255), canvas=(S, S), at=(int(S * pad), int(S * pad)))
    tile = Image.alpha_composite(tile, layer)

    return tile.resize((px_size, px_size), Image.LANCZOS)


glyph = Image.open(SRC).convert("RGBA")
# keep only the alpha shape; the source is already pure white
glyph = glyph.split()[3]
# trim to the ink bbox so centring is true
bbox = glyph.getbbox()
glyph = glyph.crop(bbox)
# pad back to square so the wordmark isn't distorted
gw, gh = glyph.size
side = max(gw, gh)
sq = Image.new("L", (side, side), 0)
sq.paste(glyph, ((side - gw) // 2, (side - gh) // 2))
glyph = sq
print(f"glyph ink bbox {bbox} -> square {glyph.size}")

# --- sidebar brand mark -----------------------------------------------------
# The renderer draws its own gradient tile (.brand-mark), so it only needs the
# white wordmark on transparency. Emitting it here keeps the sidebar emblem and
# the taskbar icon cut from the same glyph, with the same edge sharpening.
BRAND = os.path.join(os.path.dirname(OUT), "renderer", "brand-wordmark.png")
_bw = 128
_g = glyph.resize((_bw, _bw), Image.LANCZOS)
_g = _g.point([0 if i < 104 else 255 if i > 152 else int((i - 104) * 255 / 48)
               for i in range(256)])
tinted(_g, (255, 255, 255)).save(BRAND)
print("wrote", BRAND)

SIZES = [16, 24, 32, 48, 64, 128, 256]
imgs = {s: render(s, glyph) for s in SIZES}

# PNG for the BrowserWindow / non-Windows targets
imgs[256].save(os.path.join(OUT, "icon.png"))
render(512, glyph).save(os.path.join(OUT, "icon@512.png"))

# Multi-resolution ICO for electron-builder / the Windows shell
base = imgs[256]
try:
    base.save(os.path.join(OUT, "icon.ico"), format="ICO",
              sizes=[(s, s) for s in SIZES],
              append_images=[imgs[s] for s in SIZES if s != 256])
    mode = "per-size renders"
except TypeError:
    base.save(os.path.join(OUT, "icon.ico"), format="ICO", sizes=[(s, s) for s in SIZES])
    mode = "downscaled from 256"

print("wrote icon.ico (" + mode + "), icon.png, icon@512.png ->", OUT)
