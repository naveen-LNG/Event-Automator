# Blocks zone art — COLOUR RENDER STYLE

Derived from **`Blocks Art.docx`** (the studio's own "Art Direction Breakdown:
Generic Reusable Style Guide", 22 sections), checked against the shipped
Blocks2601 / 2604 / 2607 client art and the Blocks2604 game data.

> **This is NOT the Sim spec.** Before this file existed,
> `load_style_spec(events, "Blocks")` fell through to `ROOM_RENDER_STYLE.md` —
> titled "Sim room art" — with **no reference images at all**. That spec demands
> FLAT TOY-LIKE VECTOR, MATTE, no material texture and no specular, which is close
> to the opposite of what the Blocks guide asks for. Every Blocks render made
> before this file was briefed against the wrong spec.

> **The document is credible, and it was checked rather than trusted.** Its
> headline reading — soft painterly 3D, warm cinematic sunset key against cool
> ambient, atmospheric depth, vertical composition — is exactly what the shipped
> `Blocks2607/Texture/Scene*/img_*_101.png` backdrops do (960×1708 portrait,
> sunset over water, palm silhouette, soft gradient sky). Its mentions of shells,
> snorkels and a penguin-like character line up with the shipped Blocks2607
> beach set. Treat it as describing the house style, not proposing a new one.

---

## 1. The render — the one-line brief

The document's own summary, verbatim:

> *A premium, highly polished casual mobile-game illustration style featuring
> soft 3D-inspired painterly rendering, rounded toy-like shapes, cute
> anthropomorphic characters with oversized expressive faces and compact bodies,
> colorful chunky props, cinematic warm-and-cool lighting, soft volumetric depth,
> simplified backgrounds, and excellent readability at mobile-screen scale.*

It is **a hybrid between 3D animation and digital painting** — not flat 2D
illustration, and not realistic 3D.

| Do | Don't |
|---|---|
| Soft **gradient** shading, light to dark | Hard cel shading; harsh black shadows |
| **Broad, soft** highlights | Tiny metallic specular hits; sharp photoreal reflections |
| Gentle ambient occlusion where objects touch | Realistic skin pores, texture noise |
| Matte, **clay-like / soft painted toy** materials | Overly glossy plastic; harsh metallic reflection |
| Simplified stylised wood, metal, skin, fabric | Physically accurate materials |
| **Soft edge contrast** — colour and light separate forms | Thick black outlines, heavy comic-book linework |

**Edges (§16).** Not lineless, and not outlined — this is a distinction worth
holding. Forms separate by **colour, lighting and soft shadow**. An important
silhouette may carry *a slight darker edge, ambient occlusion, or a rim light*,
but never a uniform ink contour.

---

## 2. Shape language — "one of the most important aspects"

Everything is **soft, rounded, slightly inflated**. Rounded corners, inflated
forms, soft curves, slight asymmetry, chunky silhouettes, thick visual mass.

**The detail ratio is stated numerically, so use it:**

| | |
|---|---|
| Large and medium readable forms | **70%** |
| Secondary details | **20%** |
| Tiny accents | **10%** |

Large simple shapes → medium recognisable shapes → small decorative details.
Do not overcrowd with tiny realistic detail.

---

## 3. Colour — a relationship, not a palette

The governing rule (§5):

> **Warm light + cool shadows + saturated accent colours.**

Warm sunset family (golden yellow, orange, coral, peach, warm red, soft pink)
dominates the focal area; cool contrast (teal, cyan, blue, turquoise, purple)
separates objects from it. **The palette changes per theme; the relationship does
not.** The document's own worked examples:

| Theme | Reading |
|---|---|
| Winter | Warm focal lights + cool blue environment |
| Jungle | Warm sunlight + deep green shadows + turquoise accents |
| Space | Purple/blue environment + warm orange artificial lights |
| Candy | Pastel pink environment + turquoise shadows + yellow highlights |

---

## 4. Lighting — five named layers

1. **Key light** — warm yellow/orange, strong, *behind* the central scene
2. **Ambient** — soft cool blue/purple
3. **Rim light** — subtle bright edge around silhouettes
4. **Contact shadows** — soft and dark directly beneath objects
5. **Practical lights** — small decorative glows for warmth

Golden-hour sunset combined with stylised game-engine global illumination. The
backlight is what produces the bright background glow, the rim separation and
the atmospheric depth.

---

## 5. Depth, background and camera

**Background is deliberately simplified** and must not compete: large simple
shapes, soft gradients, atmospheric haze, low-frequency detail, reduced contrast.

**Detail density is graded (§17), and this is called out as critical:**

| Layer | Detail |
|---|---|
| Foreground | Highest |
| Characters | Medium-high |
| Midground | Medium |
| Background | Low |

Forward = more saturation, contrast, detail, edge definition. Back = less.

**Camera (§18):** stylised cinematic — slight perspective exaggeration, large
foreground elements, compressed background, character-centric framing, slightly
elevated viewpoint. Foreground props frame the shot; the hero sits deeper in it.
A layered storybook effect. **Vertical composition** — matching the shipped
960×1708 portrait backdrops.

---

## 6. Characters — a separate style system

Cute stylised anthropomorphic cartoon characters. **Head is 40–55% of the
character's total visual mass.** Body oval, pear, bean or rounded-rectangle.
Short simplified limbs; hands as mittens or rounded paws; legs very short or
partly hidden. Eyes large, rounded, set wide, with dark pupil, light iris and a
bright highlight. Mouth exaggerated — it may take a surprisingly large share of
the face.

> **The rule that decides it: the face must read emotion at thumbnail size.**

Expressions are broad — excitement, happiness, curiosity, surprise,
determination, playfulness, slight confusion. No subtle realistic acting.

Accessories (goggles, snorkels, hats, flowers, tools, food props) are
**oversized and readable from a distance** — no tiny buckles, no stitching, no
micro-mechanical detail. Avoid athletic anatomy, realistic muscles, long limbs,
narrow faces.

---

## 7. Props

Chunky, rounded, slightly exaggerated, simplified, colourful, easy to recognise —
the same universe as the characters. Food is slightly oversized, more colourful
than reality, reduced to an iconic shape: *immediately recognisable, appetising,
visually playful*, never scientifically accurate.

---

## 8. The document's own prompt formulas

Used verbatim by `asset_prompts` / `art_reference` for this type. Kept here so a
brief and a render are built from the same words.

**Scene / theme (§19):**

> [THEME / LOCATION], premium stylized casual mobile game artwork, cute
> anthropomorphic cartoon characters, rounded toy-like proportions, oversized
> expressive faces, compact soft bodies, smooth organic shapes, simplified
> stylized anatomy, colorful chunky props, soft painterly 3D rendering, matte
> clay-like materials, gentle volumetric shading, smooth gradient lighting, warm
> cinematic key light, cool ambient shadows, soft rim lighting, atmospheric
> depth, strong foreground midground and background layering, highly readable
> silhouettes, playful composition, polished mobile game illustration, vibrant
> but harmonious colors, whimsical environment, simplified background detail,
> soft contact shadows, high-end game promotional artwork, vertical composition

**Environment (§21):**

> Whimsical [THEME] environment, premium stylized casual mobile game art, rounded
> simplified architecture and props, chunky organic shapes, soft toy-like
> geometry, colorful stylized materials, cinematic warm lighting, cool ambient
> shadows, soft atmospheric perspective, simplified painterly background,
> volumetric glow, broad soft highlights, gentle contact shadows, strong depth
> layering, foreground props framing the composition, vibrant harmonious color
> palette, polished high-end mobile game illustration

**Character (§20):**

> Cute stylized anthropomorphic [ANIMAL/CREATURE/CHARACTER], premium casual game
> character design, large rounded head, compact pear-shaped body, short simplified
> limbs, oversized expressive eyes, soft rounded facial features, exaggerated
> friendly expression, simplified anatomy, toy-like proportions, smooth organic
> shapes, soft matte materials, subtle painterly 3D shading, broad soft
> highlights, gentle ambient occlusion, colorful oversized accessories, highly
> readable silhouette, charming playful personality, polished stylized mobile
> game artwork

**Negatives (§22)** — wired into this type's named-failure block:

> photorealistic, realistic anatomy, realistic skin texture, pores, hyper-detailed
> textures, gritty, dark horror, sharp cel shading, thick black outlines, flat
> vector art, anime proportions, realistic proportions, long limbs, thin
> characters, excessive micro-details, harsh metallic reflections, complex
> realistic backgrounds, noisy textures, monochromatic lighting, hard-edged
> shadows, overly glossy plastic, realistic photography

---

## 9. Structure, from the shipped data

`blocks2604_scene.json` — the mechanic-specific shape, distinct from Merge's
producers and Dessert's recipes:

| | |
|---|---|
| Zones | **5** |
| Buildings per zone | **7 / 7 / 8 / 8 / 8 = 38** (`buildingList`) |
| Zone backdrop | `Texture/Scene<N>/img_<N>_101.png`, **960×1708 portrait** |
| Zone art | a painted backdrop plus separate building and prop sprites layered on it |

As with Dessert, the backdrop ships **without** its buildings — they are cut-out
sprites the game places on top. So a zone brief specifies the backdrop and the
buildings separately.

The Blocks gameplay section is a **before/after restoration narrative** — each
zone starts damaged or incomplete and is revealed step by step — plus a themed
**construction material icon** list. See `TYPE_GAMEPLAY_HINTS["Blocks"]`.

---

## 10. Reference images

Seeded in `images/Blocks/`, both shipped Blocks2607 zone backdrops:

| File | Why |
|---|---|
| `blocks2607_scene1_sunset_beach.png` | The warm key light behind the scene, sunset over water, palm silhouette |
| `blocks2607_scene4_dusk.png` | The same relationship pushed further — deep warm sun against cool blue |

Chosen because the doc's single strongest rule is the **warm-key / cool-ambient
relationship**, and these two carry it most plainly. They are backdrops, so what
transfers is lighting, depth and finish rather than a layout.

> The three things §"Art Director Summary" says must survive into every new
> theme: **rounded shape language**, **character readability**, and **cinematic
> warm-and-cool lighting**.
