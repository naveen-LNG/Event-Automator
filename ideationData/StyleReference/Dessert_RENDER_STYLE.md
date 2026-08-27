# Dessert zone art — COLOUR RENDER STYLE

Derived from the Dessert2608 (SkipBo Picnic Park) and Dessert2605 (Garden
Restaurant) art requirements, the shipped Dessert2608 game data
(`WorkSpcae3/convert/RawJson/dessert2608_*.json`) and the shipped client art
(`ResourcesAb/Events/Dessert2608/`).

> **This is NOT the Sim spec, and NOT the Merge spec.** Before this file existed,
> `load_style_spec(events, "Dessert")` fell through to `ROOM_RENDER_STYLE.md` —
> titled "Sim room art" — with **no reference images at all**, the same gap Merge
> had. Dessert is further from Sim than Merge is.

> **Dessert is ISOMETRIC. Merge is three-quarter.** Do not carry Merge's rules
> over. Merge scenes are perspective vignettes full of micro-detail with readable
> shop signage; Dessert is a single continuous isometric park, softly airbrushed
> and lineless, with **no text anywhere**. Both are polychrome — see §3.

---

## 1. The render

> **This section was rewritten after a sample generation came back off-model in
> every structural way.** The first version of this file was written from the art
> requirement documents' **concept sketches**, which are outlined, isolated,
> single-plate drawings. The shipped game art is none of those things. Anything
> derived from the sketches — the outlines, the "flat-shaded high chroma"
> reading, the isolated-diorama framing — was wrong. What follows is read off
> `Texture/SceneBg/img_scene_bg_1..8` and `Texture/Photo/img_photo_1..4`, which
> are the real thing.

**Soft-shaded, airbrushed, lineless, in a continuous painted park.**

| Do | Don't |
|---|---|
| True **isometric**, high angle, one consistent camera across the whole park | Perspective with vanishing points, eye level, or a camera that changes per zone |
| **Smooth airbrushed gradients** — every form modelled with a soft light-to-shade falloff | Flat single-tone colour fills; broad areas with one hard shading step |
| **LINELESS.** Forms separate by colour and shading alone | Any dark keyline, contour, ink line or sticker edge — the single most visible way to be off-model |
| Bright but **gentle** colour: pastel greens, blush pinks, warm terracotta, soft violet | Heavy saturation, poster-vector punch, neon |
| Chunky rounded forms; a silhouette readable at icon size | Thin spindly parts that vanish when scaled |
| A soft contact shadow directly under each object | Hard black drop shadows, or objects floating |
| Each building delivered as a **cut-out sprite on transparency** | A building baked into a background |

**The surround is painted, not blank** — but *what* surrounds the pad is chosen
per event, not fixed. See §1c: this was checked across four shipped Dessert
events and the outdoor-park framing turned out to belong to 2608 alone. What is
constant is that the pad sits in *something* — a zone rendered on a white or
empty background is wrong, because that emptiness is what makes a render read as
a sticker instead of a place.

**No text anywhere on scene or building art.** Unlike Merge, Dessert buildings
carry no shop signage — motifs and shapes do the identifying work. Nothing in the
shipped Dessert2608 scene folders contains a word.

---

## 1b. What a "zone" asset actually is

This is the part the first draft got most wrong, and it is worth stating on its
own because it changes what you ask for.

**The park is ONE continuous map, delivered as tiles.** `img_scene_bg_1..8` are
1060×1448 each and stitch 4-across × 2-down into a single 4240×2896 image: one
park, with all four zone pads sitting in it at once, plus a wooden deck, small
link pads, water between them and landscape around the edge. There is no such
thing as a standalone zone picture in the shipped set.

**A zone pad ships EMPTY.** In the background art each zone is just its ground:
a coloured pad with a rim and its own surface treatment, and *nothing on it*.
Every building, cart, ride and prop is a separate cut-out sprite from
`UI/Scene1..4` that the game places on top. So a zone brief specifies **the pad**
— its shape, its surface, its rim — and the buildings are briefed separately.

**The pads are different shapes, and that is the point.** Not four copies of one
template:

| Zone | Pad |
|---|---|
| 1 | A soft **organic blob** of mown grass, irregular edge, low stone steps down to the path |
| 2 | An **organic blob** paved in pink-and-cream diamond checkerboard, flower clumps along the rim |
| 3 | A **rounded square** of terracotta, ringed by a white picket fence with festoon lights |
| 4 | A **rounded square** of violet, gold-trimmed rim, rope stands, on a raised tier |

**The composed view is a tall portrait crop.** `img_photo_1..4` are 344×488 —
roughly 3:4 portrait — and each is a crop of the continuous map centred on one
zone with its neighbour visibly bleeding in at an edge. If a render comes back as
one isolated square diorama floating on white, the framing is wrong even when the
contents are right.

---

## 1c. What is constant across Dessert events, and what is not

Checked against **four** shipped events — Dessert2512, 2602, 2605 and 2608 —
because a spec built on one event cannot tell a house rule from that event's
own choice. It could not, and one of the rules above was wrong.

**Constant in all four:**

| | |
|---|---|
| Composed zone view | `Texture/Photo/img_photo_*`, **344×488 portrait**, in every event |
| Projection | high-angle isometric, one camera per event |
| Finish | lineless, smooth airbrushed shading, bright candy colour |
| Objects | polychrome — never all one hue, in any event |
| Zone pads | each has its own shape, surface and rim |
| Text | none anywhere |
| Zones | 4 |

**NOT constant — chosen per event:**

- **The surround.** 2608 is an outdoor park: pads floating in pale blue water
  with green hills, a tree line and sky. **2512, 2602 and 2605 are none of
  those** — they are enclosed settings, densely dressed, with candy ground,
  interior walls, and no water or landscape at all (2602 runs a tavern, a
  dragon's lair, a castle and a space station; 2605 a piano room, an octopus
  pool, a pumpkin bakery and an ice palace). **2608 is the outlier, not the
  template.** Choose the surround from the event's own concept and state it in
  the brief; do not inherit 2608's water by default.
- **Density.** The earlier events pack their pads much fuller than 2608 does.
- **Pan bounds** (§2b): 2605 is `[-2000, 2000] × [-528, 528]`, 2608 is
  `[-1788, 1788] × [-690, 690]`. Per event.
- **Recipe and ingredient counts**: 2605 ships 50 recipes and 12 materials,
  2608 ships 39 and 11.

**The one structural invariant, now confirmed on two events independently of
the documents:** `scene.json` × `build_point.json` gives **11 / 11 / 12 / 12 =
46 buildings** for *both* Dessert2605 and Dessert2608, with `build_task_list`
matching the type-1 point count in every zone. Zone cost climbs the same way in
both (2605: 3,665 → 8,425 → 13,625 → 18,820; 2608: 3,225 → 7,265 → 11,705 →
16,150).

---

## 2. The fixed resource shape

Both shipped documents agree, and the game data confirms it exactly. This is a
template, not a per-event invention:

| | Dessert2605 | Dessert2608 | in the data |
|---|---|---|---|
| Zones | 4 | 4 | `scene.json` = 4 rows |
| Buildings | 11 / 11 / 12 / 12 = **46** | same | `build_point.json` type 1 = 46 |
| Secondary points | — | — | type 2 = 15 |
| Common recipes | 4 categories × 6 | 4 × 6 = **24** | `Icon/Menu1..4` = 6 each |
| Secret recipes | 1 category × 6 | **15** | `research_menu_info.json` = 15 |
| Recipes total | 30 | **39** | `menu.json` = 39 rows |
| Common ingredients | 6 | 6 | |
| Secret ingredients | — | 5 | `material.json` = 11 rows |
| Guests | 12 | 12 | `guest.json` = 17 (12 + 5 judges) |
| Judges | 5 | 5 | `judges.json` = 5 |
| Duration | 14 days | 14 days | |

The **4 zones, 46 buildings and 4×6 common recipes are the invariant**. The
secret-recipe count varies by event.

Beyond the four buildable zones there are fixed areas that are part of the base
scene and are not counted as buildings: an **Exhibition Area**, an **R&D Area**,
and a **centre celebration piece** (in 2608 a growing anniversary cake and
trophy pedestal that upgrades as competitions are won).

---

## 2b. Zone dimensions, from the shipped data

Retrieved from `WorkSpcae3/convert/RawJson/dessert2608_scene.json` and
`_build_point.json`, and measured off the stitched background art. These are
facts about the delivered event, not estimates.

### Construction elements per zone

`scene.json.point_list` joined to `build_point.json.point_type`:

| Zone | build_task_list | type 1 (buildings) | type 2 | playerPointList | complete_need_resource |
|---|---|---|---|---|---|
| 1 | 11 | **11** (`11101`–`11111`) | 3 (`11117`–`11119`) | 3 | 3,225 |
| 2 | 11 | **11** (`21101`–`21111`) | 4 (`21116`–`21119`) | 4 | 7,265 |
| 3 | 12 | **12** (`31101`–`31112`) | 4 (`31117`–`31120`) | 4 | 11,705 |
| 4 | 12 | **12** (`41101`–`41112`) | 4 (`41117`–`41120`) | 4 | 16,150 |

**11 / 11 / 12 / 12 = 46 buildings**, and `build_task_list` matches the type-1
count exactly in every zone — one build task per building. Node names are
`Z1101`…: zone digit, then a fixed `1` group, then the index. The type-2 points
are the secondary/dressing slots, and they are *not* part of the 46.

The resource cost roughly doubles then adds half again across the zones
(3.2k → 7.3k → 11.7k → 16.2k), so zone 4 is the long tail of the event.

### The pannable area

`horizontal_area` and `vertical_area` are **identical for all four zones**:

| field | value | travel |
|---|---|---|
| `horizontal_area` | `[-1788, 1788]` | 3576 |
| `vertical_area` | `[-690, 690]` | 1380 |

Verified against the client: `UIDessert2608BuildView:RefreshScreenMoveRange()`
clamps `sceneContentRcTrans.anchoredPosition` to these, so they are **the pan
limits of the scene content in canvas units** — how far the player can drag the
view — not the size of a zone pad.

> Two things worth knowing. The pan is **2.6× wider than it is tall**, so the
> park reads as a wide landscape the player scrolls sideways through, even though
> each composed view is a portrait crop. And that same function opens with
> `if self.curSceneId == nil or self.curSceneId > 3 then return end`, so **zone 4
> is never clamped** even though the data defines its bounds — worth a question
> to the client team rather than an assumption either way.

### The art, measured

The 12 `Texture/SceneBg` tiles assemble into one map. The arrangement is settled
by seam matching (mean edge difference ≤ 3.2 between true neighbours, ≥ 46
otherwise), and it is **not** the 4×2 an earlier draft of this file assumed —
tiles 9–12 form a short strip **above** the two main rows, and tiles 4, 8 and 12
are 1096 wide rather than 1060:

```
  9   10   11   12      252 px tall
  1    2    3    4     1448
  5    6    7    8     1448
1060 1060 1060 1096
```

**Total map: 4276 × 3148 px.**

Each zone pad, as a pixel bounding box in that map:

| Zone | Pad | bbox (x0,y0,x1,y1) | w × h px | w : h |
|---|---|---|---|---|
| 1 | green organic blob | 24, 712, 1572, 1532 | 1548 × 820 | 1.89 |
| 2 | pink checkerboard blob | 2, 1416, 1694, 2780 | 1692 × 1364 | 1.24 |
| 3 | terracotta rounded square | 2134, 741, 4232, 1729 | 2098 × 988 | 2.12 |
| 4 | violet rounded square | 2238, 1338, 4274, 2954 | 2036 × 1616 | 1.26 |

The pads occupy roughly the lower two-thirds of the map; the top strip and the
upper band are landscape and sky. Zones 3 and 4 run to the right-hand edge, and
zone 2 to the left, so the pads are cropped by the map boundary rather than
floating clear of it.

> The canvas-unit pan bounds and these pixel measurements cannot be reconciled
> into a single scale from what is on disk: content-minus-viewport would put the
> viewport at 700 × 1768, which is no real device. The scene content
> RectTransform is built at runtime, not stored in the prefabs, so its `sizeDelta`
> is not recoverable here. **Treat the two tables as separate facts** — the
> pixel dimensions for briefing art, the pan bounds for gameplay framing — and
> ask the client team for the scale if a brief ever needs both at once.

---

## 3. Zones are colour-coded, and the code is fixed

Dessert2608 states the progression outright: **Green → Pink → Amber → Purple**,
carried from the original event framework. A player identifies a zone by its hue
before they read anything in it.

| Zone | Theme | Hue |
|---|---|---|
| 1 | Camping Picnic Meadow | Forest Green / autumn warm |
| 2 | Sweet Bakery Open Market | Blush Pink |
| 3 | SkipBo Park Carnival | Amber / Terracotta |
| 4 | Anniversary DJ Club | Deep Violet |

**What the hue actually governs — corrected against the shipped sprites.** An
earlier draft of this file said every building holds strictly to its zone's
colour family. Measured, that is false: `UI/Scene1..4` hold 30 / 27 / 32 / 40
sprites and Scene 1 contains a teal camper van, a teal tent, a brown basket and a
white coffee cup. The objects are **polychrome candy colours**, the same as
Merge.

What carries the Green → Pink → Amber → Purple code is the **ground plate and the
zone's overall read** — its paving, its rim, its large fixed dressing — plus a
run of hero pieces in the zone hue. Zone 4 is the strongest case and the one that
misled the earlier draft: its plate, speakers and disco ball really are violet
throughout. Zones 1–3 are not.

So a brief should specify the zone hue on the **plate and the anchor pieces**,
and leave individual small buildings free to be their own colour. Forcing all 46
into one family would flatten the zone rather than identify it.

The RECIPES are the class where the zone family genuinely is strict — see §5.

**The zone plate.** Each zone is an isometric ground plate with a raised rim, its
own surface treatment (checkerboard paving, grass, dance floor) and a soft
irregular or square edge. The whole park is framed as an oval green island in a
blue moat with a roller coaster looping the perimeter and cobbled paths joining
the zones.

---

## 4. Buildings

Each of the 46 is specified with a **size class**, a material and colour note,
and the identifying details — for example:

> *Large anchor building. Classic A-frame canvas tent in deep forest green with
> warm cream canvas panels. Entrance faces the player with tent flap tied back…
> A recipe notebook pinned to the front tent pole — key food-connection detail.*

Two things worth carrying into any generated brief:

- **Size class** — "Large anchor", "Medium", small dressing. A zone needs one or
  two anchors and a spread beneath them, not eleven equal objects.
- **The food connection.** This is a cooking event, and buildings are asked to
  show it — the notebook pinned to the tent pole, the display case of pastries,
  the coffee machine on the counter. A building with no culinary hook is off-brief.

Some entries carry an **idle animation** note ("slight sway on idle"). That is an
animation instruction and never affects a still.

Not every sprite is a building: the shipped sets include ground and foreground
dressing — hedge runs, mats, fences, rope barriers, paving pads, some named
`front_*`. These are part of the zone, not part of the 46.

---

## 5. Recipes — the mechanic-specific asset class

Where Merge has producer chains, Dessert has **recipes**, and they are the
signature asset.

- **342×342, transparent, one dish centred.**
- **Four categories of six, one per zone, each in its zone's colour family.**
  Zone 1 golden-brown baked goods; Zone 2 pink frosted cakes and macarons;
  Zone 3 amber carnival treats — pretzels, candy apples, popcorn; Zone 4 violet
  drinks, parfaits and jellies.
- Glossy, saturated, soft-shaded, with a **dark keyline** and a drop shadow —
  note this differs from the scene art, which is lineless.
- **Every recipe carries the event's motif.** In 2608, the 5th-anniversary "5"
  appears on or in *every one of the 24* — a 5-shaped cookie, a 5 piped on a
  pie, a 5 candle, a 5-shaped pretzel. The general rule is that the event's motif
  is baked into each dish; the "5" is that event's instance of it.

**Ingredients are re-skinned into fixed slots, not invented.** The filenames are
a stable vocabulary shared across Dessert events — `icon_event_flour`, `_fruit`,
`_ice`, `_matcha`, `_meat`, `_milk`, `_mint`, `_spice`, `_sweet`, `_flower`,
`_pumpkin` — while the artwork behind them is redrawn per event (5 of 7 sampled
files differ between 2605 and 2608; the rest are carried over unchanged). So a
brief names its six common ingredients in the event's own language — "Party
Flour, Sweet Butter, Fresh Berry, Honey Drop, Carnival Sugar, Garden Herb" — and
those map onto the existing slots. Ingredient icons are **224×224**.

**The measured geometry.** Taken off the shipped Dessert2608 icons, and now the
constants the pipeline applies after each render (`icon_sheet.RECIPE_CANVAS`
and friends):

| | canvas | longest edge as a fraction of it |
|---|---|---|
| Recipe (`Icon/Menu1..4`, 24 common) | 342×342 | 0.68–0.86, mean **0.78** |
| Recipe (`Icon/Menu_sp_1..3`, 15 secret) | 342×342 | same treatment |
| Ingredient (`Icon/Material`, 11) | 224×224 | 0.84–0.94, mean **0.90** |

An ingredient fills more of its canvas than a recipe does — that difference is
what keeps a single raw component from reading as a small plated dish.
`Icon/Material/ImgMaterial5.png` is 408×408 and is *not* one of the 11; it is the
material-bundle image. `Icon/MenuHidden` is 128×128 ×20 and is the hidden-order
thumbnail set, not a recipe class.

---

## 6. Characters are RE-COSTUMED, not invented

The same house convention as Merge, and Dessert2605 says it in plain words:

> *12 garden-themed guest characters (**reused from existing roster**, updated
> with garden outfits — floral accessories, aprons, sun hats) + 5 judges*

The shipped set is the familiar SkipBo cast — lion, sheep, polar bear, turtle,
squirrel, parrot, chipmunk, pig — dressed for the theme. Each guest ships in
three forms:

| asset | size | where |
|---|---|---|
| `Icon/Character/img_character_NN` | 304×400 | codex portrait, on a rounded card with a warm orange gradient |
| `Icon/Animal/img_guest_NN` | 320×512 | full body, transparent, on the order page |
| Spine rig | — | walking in the built scene |

Plus **5 judges** (`Icon/Judges/img_reward_NN`, 288×288 head-and-shoulders,
unlocked progressively by Craftsman Level) and **1 chef** — in 2608 an Octopus
Anniversary Ringmaster, explicitly *"existing base, new anniversary skin"*.

So the art requirement for a character is the **costume**, not the animal.

---

## 7. Reference images

Seeded in `images/Dessert/`, both lifted from the **shipped client art**, not
from the documents:

| File | Source | Why |
|---|---|---|
| `dessert2608_shipped_zone1_picnic.png` | `Texture/Photo/img_photo_1.png` | Green organic pad, camper van, campfire |
| `dessert2608_shipped_zone4_club.png` | `Texture/Photo/img_photo_4.png` | Violet raised square, dance floor, multi-level |

Two zones deliberately far apart in hue **and in pad shape**, so what transfers
is the render, the camera and the framing rather than a palette or a layout.

> The document concept sketches that were seeded here first
> (`dessert2608_zone3_park_carnival.png`, `dessert2608_zone4_dj_club.png`) are in
> `images/_quarantine/`. They are outlined, isolated, single-plate drawings, and
> attaching them made the model reproduce all three of those properties — a
> reviewer's words: *"art style is completely different from shipped, camera
> angle is different, same composition repeated."* Concept sketches are how the
> art was **requested**; they are not how it **shipped**. Reference the shipped
> art.

`MAX_REF_IMAGES` is **2** and `_images_in()` selects by modification time — keep
the folder at two.
