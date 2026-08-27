# Merge scene art — COLOUR RENDER STYLE

Derived from the Merge2606 (Toy World) concept art in `images/Merge/`, the
Merge2604 (Blossom Gala) art requirement, and the Merge2604 art feedback round
of 08 April 2026.

> **This is NOT the Sim room spec, and NOT the Sort zone spec.** Before this
> file existed, `load_style_spec(events, "Merge")` fell through to
> `ROOM_RENDER_STYLE.md` — titled "Sim room art" — and
> `style_reference_images(events, "Merge")` returned **nothing at all**.
> Generating Merge art in that state would have repeated the exact failure
> already caught on Sort ("Art style is completely different from the attached
> references"), only worse: no reference to anchor it and the wrong spec
> injected.

> **Merge is not a cut-out plate.** Sort and Sim produce an object on a flat
> backdrop that gets knocked out to transparency. A Merge scene is a **complete
> landscape vignette** — sky, clouds, distant hills, ground — delivered whole.
> Do not knock out the background. Do not apply the alpha scoop. Those steps
> belong to Sort.

---

## 1. The render

**Soft-shaded dimensional toy diorama, lineless, saturated, high micro-detail.**

| Do | Don't |
|---|---|
| A full scene: sky, clouds, background trees or hills, an organic green base, winding paths joining the buildings | A single object floating on a flat colour field |
| **Three-quarter top-down** with real perspective — camera high and wide, taking in the whole settlement | True flat isometric; eye-level; a straight-down plan |
| Dense, generous micro-detail — every building rewards a close look | Sparse, empty, under-dressed ground |
| **A material identity carried across the whole scene** — see §2 | Generic "cute buildings" with no shared material logic |
| Soft ambient occlusion and a firm contact shadow grounding every structure | Buildings floating, or hard black drop shadows |
| Chunky rounded forms that hold their silhouette at thumbnail size | Thin fiddly detail that dies when scaled down |

**No outlines on the scene art.** Forms separate by colour and shading. (Icons
are the opposite — see §6.2. Do not confuse the two.)

**Not a photoreal 3D render.** No depth-of-field, no scanned materials.

---

## 2. Material identity per scene

This is the strongest organising idea in Merge2606 and it has no equivalent in
Sort. Each scene commits to **one physical material** and builds every structure
out of it:

| Scene | Material | Everything is made of |
|---|---|---|
| 1 — Plush Toy Kingdom | Felt and stitch | Fabric, seams, buttons, thread spools, cotton stuffing, patches, zips |
| 2 — Brick Town | Injection-moulded brick | Studded blocks, moulded slopes, transparent dome pieces, minifigure-scale fittings |
| 3 — Toy Car World | Painted toy plastic and tin | Rounded die-cast bodies, moulded ramps, glossy enamel finish |

The material is not decoration on top of ordinary architecture — it **is** the
architecture. A plush building is sewn, not built. A brick building is assembled
from discrete studded parts with visible seams between them.

Merge2604 does the same thing with a softer palette of materials per scene:
timber frames and honeycomb stonework (Village), gold-trimmed timber and
candy-striped canvas (Fairground), rough stone and terracotta tile (Banquet).

---

## 3. Colour

Each scene declares a **named palette up front** and stays inside it. Merge2604
states them literally, e.g.:

- Scene 1 — Cherry Blossom Pink, Honey Gold, Soft Violet, Fresh Meadow Green
- Scene 2 — Gold, Candy Pink, Royal Purple, Teal, Blossom Lavender

Name the palette in the brief and repeat it in the final check. Unlike Sort,
Merge is **not** built on one dominant hue per scene — a Merge scene is
polychrome by design. What holds it together is the material (§2) and the
palette list, not hue restriction.

**The sky is a scene property, not a default.** Merge2604's Fairground sky is
**purple** with a full rainbow. The feedback called this out explicitly:

> 从场景2的设定中提取元素时，不要随意更改元素识别，天空是紫色的不是蓝色的
> *When extracting elements from the Scene 2 setting, do not arbitrarily change
> the element's identity — the sky is purple, not blue.*

Never normalise a scene's sky, water or ground colour toward the expected one.

---

## 4. Composition — the radial plan

**Five of the six scenes across both documents use the same plan:** a circular
paved plaza at the centre holding a single hero object, buildings arranged
around the perimeter, paths radiating outward to connect them.

| Scene | Centre |
|---|---|
| Merge2604 S1 | Maypole Beacon on a circular stone platform |
| Merge2604 S2 | Three-tiered Fountain Plaza |
| Merge2606 S1 | Button Fountain on radial patchwork paving |
| Merge2606 S3 | Rotating car display platform |

The exception is Merge2606 Scene 2 (Brick Town), which uses a **crossroads
street grid** with pavements, crossings and street lamps instead — a commercial
high street, deliberately not a village green.

**Verticality is a per-scene decision, stated explicitly.** Merge2604:

- Scene 1 — "relatively low and wide, with the Maypole as the **sole** vertical accent"
- Scene 2 — "strong vertical variety … a dynamic, **jagged** fairground skyline"

Decide which of the two a scene is before placing anything, and say so in the
brief.

**Building count.** Merge2606 specifies **exactly 8 buildings per scene**.
Merge2604 gives 3 areas per scene at 2–4 buildings each. Both land at 3 scenes
and **45 total upgrades**, with **12 affinity NPCs at 3 quests each**.

---

## 5. Buildings — silhouette first

A Merge building must be identifiable **from its outline alone at thumbnail
size**. Both documents are built around this, in two different formats.

**Merge2606** — the more directly usable form, two fields per building:

- **轮廓特征** *(silhouette features)* — the shape
- **关键识别点** *(key identification points)* — a 3–5 item recognisability
  checklist, e.g. for the Plush Welcome House: bear ears, button windows, patch
  fabric, sagging soft roof, welcome sign at the door

**Merge2604** — six fields per building:

| Field | What it carries |
|---|---|
| Architectural Description | One-sentence identity hook |
| Structure | The build itself |
| **Vertical Focus** | Which element is tallest **and why it earns that** |
| **Bottom of the structure** | What sits at ground level — props, doors, steps, planting |
| Visual Features | Material, light, what reads from a distance |
| **Occasionally** | The idle-animation beat |

*Vertical Focus* and *Bottom of the structure* are the two fields worth carrying
into any generated brief — they are what stop a building reading as a
featureless box. *Occasionally* is an animation note, not a rendering
instruction; it describes a loop ("the dew droplet shimmers briefly", "the bee
finial glints"), so it never affects a still.

Write the checklist for every building and verify against it before delivery.

**The rule the generator breaks most often is "each appears EXACTLY ONCE."** It
has failed three times: a crane tower drawn on both sides of one scene, a
terracotta kiln dome drawn twice in another, and a repeated burner on a Sort
floor. Every one was caught by eye.

`duplicate_check.check(path, ...)` now asks a vision model, and
`check_generated` runs it when a provider is configured (`vision=...`), with the
expected count read out of the prompt — so write it as **`EXACTLY EIGHT
BUILDINGS`**, which `count_from_prompt` parses.

> This is **not** a pixel check, and it deliberately is not one. Block
> self-similarity was measured in four configurations against real scenes plus
> synthetic duplicates, and in every one the most similar pair of blocks in a
> *clean* scene was at least as similar as two genuine copies of a building —
> sky, grass and plaza correlate more strongly with each other than two
> independently drawn copies of a building do. Filtering by edge density removed
> the sky matches and the ranges still overlapped. See the module docstring for
> the numbers. Do not re-try block self-similarity.

---

## 5b. The deliverable is scenes **and** every building separately

A Merge art requirement carries **two** kinds of original art, and it is easy to
produce only the first:

| | count for a 3-scene event |
|---|---|
| Scene overview — the whole settlement | 3 |
| **Individual building sheet — one per building** | **24** (8 per scene) |

Merge2606 and Merge2604 both do this. The finishes differ completely — 2606 is
3D-rendered, 2604 is painted — but the building-sheet convention is identical:

- **ONE building, centred, filling most of the frame.** Not a scene, not a row
  of shops, no plaza or road network.
- **A small patch of its own ground** under it — grass, paving or path, just
  wide enough to carry it, edges softening away. Never floating, never cut out,
  never on a flat colour field.
- **A simple soft backdrop**: open sky and a hint of distant greenery in the
  scene's colours. No other structures in it.
- **The same three-quarter top-down camera** as the scene, roughly 30–40°.
- **The same material and palette as its scene** — this is what makes the sheet
  and the scene read as one event.
- **A little dressing at the base** — planters, lamps, barrels, benches — small
  and sparse, framing the building without competing with it.
- **Its own signage**, where the building has any, and nothing else written.
- **Roughly 4:5, portrait-ish.** The shipped references measure about 234×286.

**Generate a building sheet with its own scene attached as the style
reference.** The scene carries the material, palette, light and finish the sheet
has to match. Name the failure mode in the prompt — *do not redraw the whole
settlement, do not bring the other buildings into this frame* — because that is
what the model reaches for when it is shown a scene.

---

## 5c. Every building has UPGRADE LEVELS — that is the mechanic

The art requirement describes each building at its **final level only**.
Merge2603, the template, is explicit:

> 建筑需求描述的是**最终等级**的需求。每个建筑的升级步骤，由原画根据最终等级进行分级拆分。
> *The building requirement describes the final level. Each building's upgrade
> steps are split out by the artist from that final level.*

**The real numbers**, from the shipped Merge2608 design data
(`trunk/design/DesignData/WorkSpcae3/convert/RawJson/merge2608_building.json`):

| | |
|---|---|
| buildings | 25 |
| **total upgrade stages** | **45** — exactly 15 per scene |
| `maxLevel` distribution | 7 buildings at 1, 16 at 2, 2 at 3 |
| naming | `iconList[level]` → `MERGE2608_BUILDING_<scene>_<building>_<level>` |

So a "3 scenes × 8 buildings" event is **45 assets, not 24**. Both 2606 and 2608
state "45 total upgrades" in their stage plan; that is what the number means.

### What a level actually contains

Not what the phrase suggests. From the shipped client art
(`ResourcesAb/Events/Merge2608/UI/Scene1`), files
`img_<scene>_<building>_<level><part>.png`:

- **Building 7, Grand Restaurant** — L1: sushi counter + stools + fish display
  case · L2: condiments, cloche, bowls, **and the fish case re-authored with
  more stock**
- **Building 3, Aqua dome** — L1: heart pool + flamingo float + SKIP-BO sign ·
  L2: adds the slide · L3: adds parasol and loungers

A level is **a cluster of props added to the area, cumulatively** — not one
structure re-rendered mid-construction. Production ships them as **individual
cut-out sprites on transparent background**, positioned by the prefab, 2–4 parts
per level, each part its own canvas size.

### What to generate, and where the line is

Generate **one image per level, cumulative, showing the area as the player sees
it at that level** — the artist's derivation step that Merge2603 asks for. The
final level is the building sheet from §5b, so only the earlier stages are new.

Cutting each level into individual transparent prop sprites is a **production
step and is not generated** — that is a different asset class from anything in
this spec.

Two rules that make earlier stages checkable:

1. **No signage until the final level.** A building is signed when it opens, so
   every earlier stage declares `signage=[]` — an explicit "no text at all",
   which the reviewer enforces strictly (see §7).
2. **The chain must progress.** `sequence_review.review(final, frames)` takes the
   stages in build order and fails a chain that stalls (nothing changed) or jumps
   (the scene was redrawn rather than built on).

---

## 6. Mattel art feedback — treat as Bible

From the Merge2604 feedback round, 08 April 2026. These are corrections on
delivered work, so they override anything above.

### 6.1 Merge chains — upgrade by decorating, not by counting

> 建议升级的设计思路改为铃铛装饰在某个物品上，例如铃铛+蝴蝶结、铃铛+花环…… 而不是盲目的增加铃铛数量
> *Change the upgrade approach to a bell decorating some object — bell+bow,
> bell+wreath — rather than blindly increasing the number of bells.*

Each step up a merge chain must add **design**, not quantity. Increments of
"one more of the same thing" were rejected.

**Category purity across the chain.** 应该统一为面包类，所以不能出现第四个非面包类的
— *must be unified as the bread category, so no fourth non-bread item.* A chain
keeps one subject the whole way up. If the chain is costume, a bell appearing in
it is costume decoration, and the theme stays costume.

**Late items must not over-complicate.** 绿叶类目前第4个和第5个都太复杂了 — the
4th and 5th leaf items were both too complex. And generally: 这些细碎的点太多了
— *too many tiny fragmented bits.*

**The reference progression shape**, from the approved past design:
single item → pair → cluster → overflowing open box → treasure chest.
The chest is the top tier and should be the most refined piece — 把宝箱设计的再
精致一些，提升升级感.

**The producer belongs at the top of the chain** and should combine materials
(e.g. wood + ribbon) rather than being a plain machine.

**Fill the tile.** 让图标尽可能饱满，不要留出太多空白区域 — *make the icon as full
as possible, do not leave too much empty area.*

### 6.2 Icon craft — the AI-specific defect

> 目前全部图标还缺少统一的描边、投影、图标大小进阶的处理。
> （如果使用chatgpt生成了无描边的图标，需要额外添加1-2像素的描边效果）
> *All icons currently lack unified stroke, drop shadow, and icon-size
> progression. If you used ChatGPT to generate icons without strokes, you need
> to add a 1–2 pixel stroke.*

They have already seen AI-generated icons and named exactly how they arrive
wrong. Every icon needs:

1. A **stroke** — 1–2 px, consistent width across the whole set
2. A **drop shadow**, consistent in angle and softness
3. **Size progression** — later chain items read larger than earlier ones

This is the direct inverse of the scene rule in §1. Scene art is lineless; icons
are stroked. Never apply one rule to the other.

**Perspective consistency**: the coin icon must match the reference's tilt and
perspective angle, stroke weight and thickness — 保证描边、角度、厚度统一.

**No hollow outlines**: the entrance icon needs a stroke *and* a background —
"try not to have only a hollow outer contour shape". Explain what is behind the
arch, or occlude it with foliage.

### 6.3 The merge gameboard — four named layers

The gameboard is a **separate deliverable** from the scene art: a portrait
mobile UI with the merge grid over a built background. It has four layers, each
with its own rule.

| Layer | Rule |
|---|---|
| **背景 Background** | Only elements the scene setting actually contains. *"场景设定中没有河流，不要出现与场景不匹配的元素"* — Scene 1 has no river; do not introduce elements that do not match the scene. Sky colour per §3. |
| **前景 Foreground** | **Extract** from the concept art — grass, path surface, fencing. Scene 1 has no picnic or road signs, so they must not appear. **Never paste a scene building in as foreground decoration** (前景装饰不能直接用场景建筑); pull design elements from it instead. |
| **立面 Facade** | **Never a flat colour.** Treat it as a wall: take material from the concept art — wood, brick, stone — and decoration from it too — texture, climbing vines, windows. |
| **Bottom structure** | Must be **uniform across all three scenes**, e.g. every one has steps + ground with objects placed on the ground. All three delivered boards differed from each other. |

**Small sizes keep their bottom decoration.** At reduced sizes the bottom
decoration was being dropped; the reference (the clover event) keeps it present
and orderly at every size.

### 6.4 Characters and animation

- Event NPCs appear only in the NPC folder and the **DialoguePanel**. Walking
  and idle animations are **not** required unless the designer asks.
- Do **not** reuse an existing rig's walk when it does not suit the character
  design — the bear's walk was rejected for this.
- Cloaks and capes need their **own** animation driven by the walk dynamics.
- **Walk posture bug, systemic:** the left leg must be behind the right leg.
  Incorrect lower-limb hierarchy was found across **multiple** character Spine
  files — check all of them, not just the one reported.

---

## 7. Text — Merge is the opposite of Sort

**Readable signage text is REQUIRED in Merge scene art.** Observed in the
seeded references: `MARKET`, `BAKERY`, `DEPT. STORE`, `FLOWER SHOP`,
`HOME SWEET HOME`, `FRESH BREAD`, `FRESH LOCAL`, `WELCOME`,
`RAINBOW CAR WASH`, `GARDEN TRAIN STATION`, `PET SIGHTSEEING BUS STOP`,
`GIFT DELIVERY STATION`, `NEWS`.

Shop signage is part of how a Merge building announces its function, and §5's
recognisability requirement leans on it.

This inverts the Sort rule, where any rendered text is a defect.

**Signage is scenery. Annotation is not.** What stays banned: callout markers,
numbered keys, caption strips, legend blocks — and **a building's design name**.
Write `TEA`, `MARKET`, `SWEET OVEN` on the shop. Never write "Music Box
Pavilion" on the music box pavilion; it is identified by its wind-up key.
A generated scene did exactly that and had to be redone.

**How this is checked.** `text_detect.check_signage(path, allowed)` reads the
words actually rendered and flags any that the brief did not authorise. Pass the
exact sign wording from the brief:

```python
check_signage(scene, ["TEA", "LAMPS", "POST"])
# -> ok False, unexpected ['MUSIC BOX PAVILION']
```

**In the automated pass this is wired up and routed by event type.**
`artreq_images.check_generated(paths, tones, event_type, signage)` sends Merge
down the signage check and every other type down the blanket text ban — the two
rules are opposite and a single rule cannot serve both. Both call sites forward
the type automatically.

The authorised wording is read back out of the prompt, the same way MAIN TONE
is, so nothing has to be added to the document schema. **Write each sign in the
brief as `sign reads MARKET`** — one per building, on its own line or clause:

```
- Origami Flower Shop - a low shop with folded blooms, sign reads FLOWER FOLD
- Confetti Bakery - a rounded shop with card awning, sign reads PAPER BAKERY
```

If a Merge brief declares **no** wording, the reviewer emits a **warning**, not
a pass: it cannot judge text it was given no expectation for, and staying silent
there is what let the design name through in the first place.

Matching is deliberately lenient about how the words arrive, because OCR is not
consistent on stylised signs: it accepts a per-word match, a fragment of a longer
sign (`PAPER BAKERY` on two lines reads back as `PAPER` and `BAKERY`), and a
squashed read where the spaces were dropped (`MUSICBOXPAVILION`). What it will
not accept is a word the brief never mentioned. Known gap: a forbidden phrase
built entirely out of authorised words would pass — design names in practice
introduce their own vocabulary, which is what makes this work.

> **This runs on OCR, which is vendored into `pybackend_dist`.** It cost the
> dist 75 MB → 313 MB. The engine loads once per process (~1.4s) and reads an
> image in 1–2s.
>
> The import stays optional by design: if the packages are ever missing,
> `check_signage` returns `ok: None` — "could not check", deliberately
> distinguishable from "checked and clean" so a missing backend never reads as a
> pass. See VENDORING in `text_detect.py` to reinstall.
>
> The plain structural detector is **blind to Merge text** and cannot substitute.
> Measured: zero candidate lines across three Merge scenes carrying eleven
> correct signs, and zero on a 30px bold label composited straight onto a scene —
> while firing correctly on the Sort caption strip it was tuned for. Letters
> bridge through their anti-aliased edges when the image is downscaled, so a word
> becomes one fat blob instead of N glyphs and the "row of small glyphs"
> signature is simply absent.
>
> `check_generated(paths, tones)` still receives no event type, so the generic
> reviewer cannot yet route Sort and Merge to different text rules.

---

## 8. Reference images

Seeded in `images/Merge/`:

| File | Source |
|---|---|
| `merge2606_scene1_plush_kingdom.png` | Merge2606 p.2 — Plush Toy Kingdom |
| `merge2606_scene2_brick_town.png` | Merge2606 p.5 — Brick commercial town |

`MAX_REF_IMAGES` is **2** and `_images_in()` selects by modification time,
newest first — so any third image added to this folder can silently displace one
of these. Keep the folder at two.

**Quarantined** in `images/Merge/_excluded/` — Toy Car World (Merge2606 p.8).
Good art, unusable as an anchor: it carries circled callout numbers 1–8 stamped
over the buildings as a document annotation. Attaching a numbered image is how
an earlier Sort render learned to draw callout markers and a numbered key into
finished artwork. The subfolder is what excludes it — `_images_in()` only lists
files ending in an image extension directly inside `images/Merge/`. See the
README there.
