# Sort zone art — COLOUR RENDER STYLE

Derived from the five shipped Sort2609 (Pet Sanctuary) zone floors in
`images/Sort/`: cat room, dog yard, aquarium, reptile desert, aviary.

> **This is NOT the Sim room spec.** `ROOM_RENDER_STYLE.md` is titled "Sim room
> art" and demands flat toy-like vector, matte, and absolutely no material
> texture or specular. Applying it to Sort produced flat, washed-out, vector
> renders that read nothing like the shipped floors. Sort is **richer**: soft
> painted, glossy where the material is glossy, and it does carry stylised
> material detail.

---

## 1. The render

**Soft-painted, richly shaded 2.5D game asset. Lineless. Saturated.**

| Do | Don't |
|---|---|
| Smooth airbrushed gradients across every form, with a wide range from lit face to deep shadow | Flat unshaded fills, two-step cel ramps, sticker flatness |
| Deep, confident, saturated colour | Washed-out pastel, chalky or milky tints |
| **Specular highlights where the material is glossy** — water, glass, glazed ceramic, polished wood, plastic, metal, pearls | A uniformly matte scene; killing all gloss makes it read as plastic vector |
| **Stylised material suggestion** — wood plank seams and light grain, rope twist, woven basket bands, grass tufts, water caustics, fabric folds | Photoreal texture, scanned material, noise, grunge, ambient-occlusion dirt |
| Soft ambient occlusion in crevices, plus a firm contact shadow under every object | Objects floating, or a hard black drop shadow |
| Chunky rounded toy forms with bold silhouettes | Thin, spindly or fiddly detail that dies at thumbnail size |

**No outlines, anywhere.** No dark contour, ink line, keyline or sticker edge,
and never a darker band ringing a shape. Forms separate by colour and by
shading. This is the one rule Sort shares with the Sim spec, and the one the
generator most often breaks — name it explicitly every time.

**Not a 3D render.** No CG/Blender product-render look, no photoreal materials,
no depth-of-field blur. Painted, not rendered.

---

## 1b. The style bible — from `Sort Art.docx`

The studio's own 22-section breakdown. Its one-line essence:

> *"A premium, colourful, highly detailed isometric toy-diorama environment where
> every room feels like a miniature playable world, built from soft rounded 3D
> objects with strong silhouettes, clean gradients, gentle lighting, and dense
> thematic storytelling."*

Most of it agrees with what is already here and sharpens it.

**§2 Camera.** Looking down at **30–45°**, showing front, side and top faces,
very little perspective distortion — orthographic/pseudo-isometric rather than
cinematic. *"Each floor behaves like a small diorama platform"*, which is exactly
the plate this file already describes.

**§3 Building composition — the part that is genuinely new here.** Sort is a
**vertical cutaway dollhouse**: a stack of floors, each with a floor platform,
back and side walls, and an open front. Our zone rules describe one floor; this
describes the tower they belong to. Structural hierarchy per floor:
architecture → main furniture → gameplay objects → small props → characters →
decorative details.

**§7 Materials.** Three named classes, and the split matters: **plastic-like**
(appliances, toys, tech — smooth, soft reflections, clean gradients),
**matte painted** (walls, furniture, floors — soft shading, minimal noise), and
**glossy accents used selectively** (glass, screens, water, bottles, metal).
Not heavy realistic PBR.

**§9 Shading — called "extremely important for reproducing the style".** Every
object: bright top surface → mid-tone front → darker side face → small highlight
→ soft ambient occlusion underneath.

**§10 Outlines.** Minimal or selective. Separation comes from colour contrast,
soft shadow, ambient occlusion and slight bevels. A subtle darker contour is
allowed; thick black outlines, ink lines and cel borders are not. That is
consistent with §5.8 below.

**§15 Floor layering.** Back layer walls and architecture · middle layer main
furniture and the activity · front layer small props and collectibles.

**§18 Level of detail.** Large objects moderate detail, small objects strong
colour and silhouette with minimal geometry, tiny decorations as simple forms and
colour blobs. *"Do not make every object equally detailed."*

**§22 Theme formula.** `[THEME] + miniature isometric dollhouse + 3–5 themed
furniture pieces + 10–20 small storytelling props + dominant theme palette +
cute character activity + soft rounded toy-like geometry + polished casual
mobile-game rendering`.

### Where this document and the shipped assets DISAGREE — density

§11 asks for *"high-density, controlled clutter"* and §21 lists **"empty rooms"**
as a thing to avoid. The measured shipped Sort2609 floors say the opposite:

| | measured on the five shipped floors |
|---|---|
| Detail fill | **5.5–7.0%**, mean 6.4% |
| Top band | 2.9–5.5% — sparsest in all five |
| Middle band | 7.5–9.4% — densest in all five |
| Centre column | quietest of the three in four of five |

Roughly **93% of a shipped plate is quiet surface**. A first generated attempt
came in at 9.9–18.7% fill and read "correct but cluttered" — it had followed
exactly the instinct §11 describes.

**The measurement wins.** It is taken off the assets that shipped, and Sort's
gameplay requires every object to read as its own separate silhouette with
nothing overlapping — density directly fights that. Read §11 as *"many small
themed props, each in its own clear pocket"*, not as filling the plate. The
per-band and per-column targets in §7 of `Sort_zone_floor_rules.md` are the
operative numbers.

---

## 2. Colour

Each zone is built almost entirely from tints and shades of **one dominant
hue** — walls, floor and most furniture. Players identify a zone by its colour
at thumbnail size, so the five hues stay far apart.

The **goods** are the only place that breaks: they keep their own natural
saturated colours and provide all the accent contrast. The dog yard's
primary-coloured ball pit, the aquarium's gold coins and pink coral, the reptile
zone's green cactus and fruit, the aviary's corn and carrot.

Unlike the Sim spec, Sort is **not restricted to a candy set**. The reptile zone
is built on terracotta and warm sandstone; the dog yard uses natural wood browns
and grass green. Earthy colour is allowed here where the habitat calls for it —
what is banned is *dull, greyed, muddy* colour, not warm natural colour.

---

## 3. Composition

See `Events/Sort/Art Requirements/Sort_zone_floor_rules.md` for the full
structure. In short: isometric high camera, rimmed floor plate, partial back
walls, 10–14 discrete non-overlapping objects around an open centre, no hero.

---

## 4. The top scoop is ALPHA, not art

Two of the five floors (aquarium, reptile) have a curved bite taken out of the
top for the game UI. **It is transparency in the PNG, not painted background** —
measured at 49–53% transparent in the top band, versus 0% for cat and aviary.

So it cannot be prompted for. Generate the full frame, then cut it in post with
`scoop.apply_top_scoop()`, which reproduces the measured curve: a downward
sine arc, zero at both edges, peaking at ~0.17 of the image height at centre.

Only apply it to zones whose layout calls for it — cat and aviary ship without.

---

## 5. Mattel art feedback — treat as Bible

From the Mattel review notes and the first Mattel meeting. These override
anything above that conflicts with them.

### 5.1 Colour — analogous, not complementary

> "When designing, we usually try to choose **analogous colors rather than
> complementary colors**. Even when using complementary colors, we use them in
> small areas to ensure the color progression isn't disrupted."

The rejected example was a kitchen with red, teal, green, yellow and pink all
fighting at full strength. **Neighbouring hues, one family, plus small
complementary accents only.**

> **Note the tension with the Sim spec.** `ROOM_RENDER_STYLE.md` tells the four
> display objects to take four *different* hue families (coral / mint /
> lavender / raspberry) and to separate "by HUE, not merely brightness". That is
> Sim guidance. For **Sort**, Mattel's analogous rule wins: keep the furniture
> within a family of the room tone and let value and saturation do the
> separating, with complementary colour reserved for small accents.

### 5.2 Define the main tone first

Mattel's stated order of operations:

1. **Decide the room's main tone — cool or warm — before anything else.**
2. **Floor and walls set that tone**, because they cover the largest area and
   are unlocked earliest.
3. Furniture only has to sit inside that tone. Each room's tone must
   **distinguish it from the other rooms** — that is what the tone is for.
4. A character's clothing only needs to *echo* the tone; it does not have to
   match walls, floor or furniture.

### 5.3 Walls and floor share one colour

> "Walls and floor should have the **same color with a slight variation in
> tone**. This permits elements standing out."

Not two colours that harmonise — the same colour, separated only by a small
step in value. The furniture is what the player should see.

### 5.4 Props must sit inside the tone

A picture frame was sent back for being too warm for a cool room: *"the color
could be adjusted to be cooler, as it looks somewhat out of place."* Every prop,
however small, is tinted toward the room tone.

### 5.5 Light and shadow — restrained, for layering

> "You can add some light and shadow to the room appropriately, **but not too
> much**. The main purpose is to **highlight the layering of the furniture**."

Shadow is a separation device, not drama. Enough to read depth between
overlapping pieces; no theatrical pools, no strong cast patterns.

### 5.6 Characters need edge separation

A character's hat blended into the background. The fix Mattel gave:

> "Add a **top light to the top of the hat (within the outline)** to create an
> edge line. This would both match the lighting settings and highlight the depth
> between the character and the background."

So a character over a same-tone room gets a rim/top light inside its silhouette.

### 5.7 Consistent detail standard

> "Some details must be depicted in a way that maintains a consistent production
> standard."

Recurring motifs — flowers, foliage, jars — are drawn to one spec everywhere
they appear, not reinvented per asset. Pick the shape once and repeat it.

### 5.8 Edges — the AI-specific defect

> "Pay attention to these edges. **AI-generated resources will have obvious
> white edges.** We need to use the mask tool in Photoshop to cut out these
> white lines to make the edges look clean overall."

This is handled in the pipeline rather than by hand: `scoop.dehalo()` measures
the fringe (edge lightness minus interior lightness), erodes the alpha and
bleeds interior colour outward with an alpha-weighted blur. Measured across five
renders, mean |fringe| fell from 53.7 to 10.2. It **measures before acting**, so
an already-clean edge is left alone rather than over-corrected. `cut_asset()`
runs it automatically, and the reviewer warns when any surviving fringe exceeds
tolerance.

### 5.9 Reference bleed — attach at most two

Attaching all five Sort floors as image inputs made NanoBananaAIO **composite**
them instead of imitating them. One glaze-room render came back with the entire
cat room laid over it: pink striped wallpaper, the arched window with its star
bow, the cat tower, the food bowls. The requested sea-green room was underneath.

Measured: that output's top band was `(226,172,193)` against the cat room
reference's `(243,179,203)` — effectively the reference pixels.

Regenerating with **two** references (and dropping the cat room, the closest
match compositionally to an interior) produced a clean room, top band
`(148,182,177)`, no pet furniture anywhere.

`MAX_REF_IMAGES` is now **2**. Style transfers from one or two examples; content
leaks from five. Prefer references whose *content* is far from the target — an
underwater or outdoor floor is a safer style anchor for an interior than another
interior.

Pair the cap with the named ban in the prompt: *"DO NOT COPY, TRACE, BLEND OR
OVERLAY ANY PART OF THEM… The attached rooms are pet rooms. NOTHING from them
belongs here."* The cap alone is the load-bearing fix; the ban is belt and braces.

**How the gate detects it.** Not by comparing the output to the reference.
That was tried and it inverted: on a coarse grid the bled render scored 0.10
against the room it had copied while clean floors scored 0.11–0.27 against
references they had nothing to do with. The bleed arrives as a semi-transparent
overlay, so nothing is ever pixel-identical, and the differing aspect ratios
misalign every feature anyway. Do not re-try reference similarity.

What works is that the bleed drags the reference's **palette** in, breaking the
MAIN TONE the brief set. `bleed_check.check(path, tone)` measures the fraction of
hue-bearing area in the wrong tone family and fails above **35%**:

| render | brief | off-tone |
|---|---|---|
| glaze room, cat room composited in | COOL | **53.9%** ← caught |
| glaze room, 2 refs, regenerated | COOL | 1.1% |
| guest room | COOL | 1.9% |
| sea terrace, deliberate golden-hour light | COOL | 19.5% |
| shipped aviary (genuinely mixed) | mixed | 27.5% |

The threshold sits above the highest legitimate reading and well below the
failure. Near-neutrals are excluded — a grey wall casts no tone vote — so the
one permitted warm accent never trips it. A reference copied in whose palette
happens to agree with the brief still slips through; this catches the failure
that actually shipped.
