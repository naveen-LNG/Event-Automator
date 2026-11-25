# Merge scene art — COLOUR RENDER STYLE

Derived from the art that SHIPPED in the game: the scene share pictures of
Merge2511, Merge2512, Merge2603, Merge2606, Merge2608 and Merge2610
(`ResourcesAb/Events/<event>/Texture/Map/img_share_N.png`), plus the Merge2604
art feedback round of 08 April 2026 and the shipped Merge design data.
Rewritten 2026-11-09: earlier versions of this file were built from
pre-production concept paintings and a style document that described a
different finish from the one players actually see, and art generated from them
was rejected as "not the Merge event art style".

> **This is NOT the Sim room spec, and NOT the Sort zone spec.**
> `load_style_spec(events, "Merge")` returns this file and
> `style_reference_images(events, "Merge")` returns `images/Merge/`.

> **Merge is not a cut-out plate.** Sort and Sim produce an object on a flat
> backdrop that gets knocked out to transparency. A Merge scene is a **whole
> map** delivered in one piece. Do not knock out the background. Do not apply the
> alpha scoop. Those steps belong to Sort.

> **Note for anyone editing this file:** every word here is injected verbatim
> into the render prompt, so a paragraph *describing* a banned instruction reads
> as the instruction itself. Describe superseded rules; never quote them.

---

## 1. The render

**2D hand-painted casual mobile-game illustration: soft gradients, crisp
readable shapes, bright and cheerful.**

| Do | Don't |
|---|---|
| Clean digital painting: flat-to-soft colour fills, soft gradients, painted highlights, crisp shape edges with a slightly darker edge tone | A rendered 3D look: clay, plastic, glossy toy materials, ray-traced light |
| **High top-down map camera** (roughly 50–60° down), looking down on the scene like a game map | Eye-level views; a low three-quarter view with a big sky |
| **Ground fills the frame edge to edge.** At most a band of trees, hills or cloud along the top edge | A floating island, diorama base or plinth; a large sky; distant landscape vistas |
| Big, chunky buildings that read from their outline at thumbnail size | Thin, fiddly detail that dies when scaled down |
| A soft contact shadow under every building and character | Buildings floating; hard black drop shadows |
| Soft, even lighting across the whole map | Dramatic spot-lit focal points with dark surroundings |

**No black outlines on the scene art.** Forms separate by colour and a slightly
darker edge tone. (Icons are the opposite — see §6.2.)

**Not photoreal.** No depth-of-field, no scanned textures, no realistic
materials.

---

## 2. The visual DNA, read off the shipped scenes

| | Shipped Merge scenes |
|---|---|
| Finish | 2D painted, soft gradients, crisp shapes |
| Camera | high top-down map view |
| Framing | ground to the edges; at most a top band of trees or sky |
| Buildings | **novelty shapes**: buildings that ARE an object or an animal |
| Layout | each building on its own pad round an open central plaza, joined by wide paths |
| Colour | very bright, saturated candy colours; one dominant ground colour per scene |
| Light | soft and even; night scenes stay bright and saturated |
| Characters | small cute animal cast in themed outfits, scattered on the paths |
| Detail | moderate: clear props round each building, generous open ground between |

### Novelty buildings — the signature

The thing that makes a scene read as Merge is that the buildings are shaped
like **objects or animals**, not ordinary houses with themed props on them:

| Event | Shipped novelty buildings |
|---|---|
| Merge2512 Merry Merge | a giant lucky-cat house, a daruma house, an ice castle |
| Merge2606 Toy Towns | a teddy-bear house, a plush bunny cottage, a sewing-spool house, a bow-topped playhouse |
| Merge2610 Spooky Streets | pumpkin houses, a witch-hat house, a black-cat cottage |
| Merge2603 Green Hat Day | a mushroom house, a pretzel bakery, clover-topped cottages |

Give every building in a brief one of these hooks. A plain building with a
themed sign is the weakest option.

### Composition

A central open **plaza** (round or organic) with the hero building or hero
object at or near the middle; the other buildings sit on their own pads
around it, each with a little breathing room, joined by **wide winding paths**.
Nothing overlaps another building. Edges are framed by trees, snowdrifts,
water or terrain in the scene's colours.

```
   [trees / terrain along the top edge]
   [A]        [HERO]        [B]
   [C]  ---- open plaza ---- [D]
   [E]        [F]           [G]
   [terrain along the bottom edge]
```

---

## 3. Colour

**Bright, saturated and cheerful, with ONE dominant ground colour per scene.**
The ground colour is the scene's identity:

| Scene | Ground |
|---|---|
| Merge2512 S1 (Christmas) | lavender-blue snow |
| Merge2512 S2 (Japan, winter) | warm wood-plank plaza on white snow |
| Merge2511 S2 (Barbie DreamHouse) | candy pink |
| Merge2606 S1 (Toy Towns) | bright grass green with a cream plaza |
| Merge2610 S2 (Spooky Streets, night) | violet-blue cliffs and green grass |

Building colours are polychrome on top of that ground, but stay in the same
bright candy key. Name the palette in the brief and repeat it in the final
check.

**Night is bright.** A night scene is deep violet-blue with glowing lanterns,
string lights and lit windows — still saturated and easy to read, never dark,
murky or gloomy.

**The ground or sky colour is a scene property, not a default.** Merge2604's
Fairground sky is **purple**. The feedback called this out explicitly:

> 从场景2的设定中提取元素时，不要随意更改元素识别，天空是紫色的不是蓝色的
> *When extracting elements from the Scene 2 setting, do not arbitrarily change
> the element's identity — the sky is purple, not blue.*

Never normalise a scene's ground, water or sky colour toward the expected one.

---

## 4. Scene count, building count, and one material idea (optional)

From the shipped design data (`WorkSpcae3/backup_xlsx/<event>.xlsx`, sheet 建筑):

| Event | Days | Scenes | Buildings per scene | Stages per scene |
|---|---|---|---|---|
| Merge2512 | — | 4 | 8 / 8 / 8 / 8 | 18 / 18 / 17 / 18 |
| Merge2603 | — | 3 | 8 / 8 / 7 | 18 / 18 / 18 |
| Merge2606 | — | 3 | 8 / 9 / 8 | 15 / 15 / 15 |
| Merge2608 | 14 | 3 | 7 / 9 / 9 | 15 / 15 / 15 |

A 14-day Merge is **3 scenes, 25 buildings, 45 upgrade stages, 15 per scene**,
with **12 affinity NPCs** (101–104, 201–204, 301–304).

**A shared material per scene** (Merge2606's felt-and-stitch, brick, toy
plastic) is one strong way to tie a scene together, but the shipped scenes do
not require it: Merge2512's Christmas village mixes timber, ice and gift wrap.
Use it when the theme suggests one; the novelty shapes and the ground colour
carry the scene either way.

---

## 5. Buildings — silhouette first

A Merge building must be identifiable **from its outline alone at thumbnail
size**. Two fields per building, the form Merge2606 uses:

- **轮廓特征** *(silhouette features)* — the shape, in one sentence
- **关键识别点** *(key identification points)* — a 3–5 item recognisability
  checklist, e.g. for a teddy-bear house: bear ears, button windows, patch
  fabric, soft sagging roof, welcome sign at the door

Write the checklist for every building and verify against it before delivery.

**Every building must differ from every other building in the event.** Do not
give two buildings the same core shape (two decorated Christmas trees, two
snails, two stumps). The hero object appears once.

**The rule the generator breaks most often is "each appears EXACTLY ONCE."** It
has failed repeatedly: a crane tower drawn on both sides of one scene, a kiln
dome drawn twice, an unlisted extra tree beside the listed buildings, two
buildings drawn merged into one. Every one was caught by eye. Write the count
in the prompt as **`EXACTLY EIGHT BUILDINGS`** (parsed by `count_from_prompt`)
and add "there are NO other buildings anywhere — only these".

`duplicate_check.check(path, ...)` asks a vision model, and `check_generated`
runs it when a provider is configured. It is **not** a pixel check, deliberately:
block self-similarity was measured and could not separate genuine duplicates
from sky, grass and plaza. Do not re-try block self-similarity.

---

## 5b. The deliverable is scenes **and** every building separately

| | count for a 3-scene event |
|---|---|
| Scene overview — the whole map | 3 |
| **Individual building sheet — one per building** | **25** |

Building-sheet convention:

- **ONE building, centred, filling most of the frame.** Not a scene, no other
  buildings, no plaza or road network.
- **Its own pad of ground** in the scene's ground colour, just wide enough to
  carry it, edges softening away. Never floating, never on a flat colour field.
- **The same high top-down camera, painted finish and palette as its scene.**
- **A little dressing at the base** — lanterns, crates, benches — small and
  sparse.
- **Its own signage**, where the building has any, and nothing else written.
- **Roughly 4:5, portrait-ish.**

Generate a building sheet with its own scene attached as the style reference,
and say in the prompt *do not redraw the whole map, do not bring the other
buildings into this frame*.

---

## 5c. Every building has UPGRADE LEVELS — that is the mechanic

The art requirement describes each building at its **final level**; the
artist splits the earlier stages out of it (Merge2603: 建筑需求描述的是最终等级的需求。
每个建筑的升级步骤，由原画根据最终等级进行分级拆分。).

Each building has **1 to 3 levels**; **level 1 is the first build** and counts
as a stage. Shipped `maxLevel` mixes for a 14-day event:

| Event | Scene 1 | Scene 2 | Scene 3 |
|---|---|---|---|
| Merge2606 | 2,2,2,1,2,2,2,2 | 1,2,1,2,2,1,2,2,2 | 3,2,2,2,1,1,2,2 |
| Merge2608 | 1,2,3,2,3,2,2 | 1,2,2,2,2,2,1,1,2 | 1,1,2,2,2,2,1,2,2 |

Naming: `iconList[level]` → `MERGE<event>_BUILDING_<scene>_<building>_<level>`.

### What a level actually contains

From the shipped client art (`ResourcesAb/Events/Merge2608/UI/Scene1`, files
`img_<scene>_<building>_<level><part>.png`):

- **Building 7, Grand Restaurant** — L1: sushi counter + stools + fish display
  case · L2: condiments, cloche, bowls, and the fish case re-authored with more
  stock
- **Building 3, Aqua dome** — L1: heart pool + flamingo float + sign · L2: adds
  the slide · L3: adds parasol and loungers

A level is **a cluster of props added to the area, cumulatively** — not one
structure re-rendered mid-construction. Production ships them as individual
cut-out sprites on transparent background, positioned by the prefab, 2–4 parts
per level.

Generate **one image per level, cumulative, showing the area as the player sees
it at that level**. Cutting each level into transparent prop sprites is a
production step and is not generated.

1. **No signage until the final level** — every earlier stage declares
   `signage=[]`.
2. **The chain must progress.** `sequence_review.review(final, frames)` fails a
   chain that stalls or jumps.

### Scene background files

Each scene ships as two portrait halves, `img_scenebg1.png` and
`img_scenebg2.png`, 960×1708 each, with cloud overlays `img_cloud2/3`; the share
pictures `img_share_1..4` in `Texture/Map/` are about 568×396.

---

## 6. Mattel art feedback — treat as Bible

From the Merge2604 feedback round, 08 April 2026. These are corrections on
delivered work, so they override anything above.

### 6.1 Merge chains — upgrade by decorating, not by counting

> 建议升级的设计思路改为铃铛装饰在某个物品上，例如铃铛+蝴蝶结、铃铛+花环…… 而不是盲目的增加铃铛数量
> *Change the upgrade approach to a bell decorating some object — bell+bow,
> bell+wreath — rather than blindly increasing the number of bells.*

Each step up a merge chain must add **design**, not quantity.

**Category purity across the chain.** 应该统一为面包类，所以不能出现第四个非面包类的
— *must be unified as the bread category, so no fourth non-bread item.* A chain
keeps one subject the whole way up.

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

Every icon needs:

1. A **stroke** — 1–2 px, consistent width across the whole set
2. A **drop shadow**, consistent in angle and softness
3. **Size progression** — later chain items read larger than earlier ones

Scene art has no black outlines; icons are stroked. Never apply one rule to the
other.

**Perspective consistency**: the coin icon must match the reference's tilt and
perspective angle, stroke weight and thickness — 保证描边、角度、厚度统一.

**No hollow outlines**: the entrance icon needs a stroke *and* a background —
"try not to have only a hollow outer contour shape".

### 6.3 The merge gameboard — four named layers

The gameboard is a **separate deliverable** from the scene art: a portrait
mobile UI with the merge grid over a built background.

| Layer | Rule |
|---|---|
| **背景 Background** | Only elements the scene setting actually contains. *"场景设定中没有河流，不要出现与场景不匹配的元素"* — Scene 1 has no river; do not introduce elements that do not match the scene. Colour per §3. |
| **前景 Foreground** | **Extract** from the concept art — grass, path surface, fencing. **Never paste a scene building in as foreground decoration** (前景装饰不能直接用场景建筑); pull design elements from it instead. |
| **立面 Facade** | **Never a flat colour.** Treat it as a wall: take material and decoration from the concept art — wood, brick, stone, vines, windows. |
| **Bottom structure** | Must be **uniform across all three scenes**, e.g. every one has steps + ground with objects placed on the ground. |

**Small sizes keep their bottom decoration.**

### 6.4 Characters and animation

- Event NPCs appear only in the NPC folder and the **DialoguePanel**. Walking
  and idle animations are **not** required unless the designer asks.
- Events **re-costume the existing cast** (hippo, raccoon, poodle, turtle,
  squirrel, rabbit, owl, monkey, pig, bear) rather than inventing new animals.
- Do **not** reuse an existing rig's walk when it does not suit the character
  design — the bear's walk was rejected for this.
- Cloaks and capes need their **own** animation driven by the walk dynamics.
- **Walk posture bug, systemic:** the left leg must be behind the right leg.
  Check every character Spine file, not just the one reported.

---

## 7. Text — Merge is the opposite of Sort

**Readable signage text is allowed in Merge scene art**, one short word or two
per building that needs it: `MARKET`, `BAKERY`, `FLOWER SHOP`, `POST`,
`LANTERNS`. Shop signage is part of how a Merge building announces its
function.

This inverts the Sort rule, where any rendered text is a defect.

**Signage is scenery. Annotation is not.** What stays banned: callout markers,
numbered keys, caption strips, legend blocks — and **a building's design name**.
Write `TEA`, `MARKET`, `SWEET OVEN` on the shop. Never write the design name
("Music Box Pavilion") on the building.

**How this is checked.** `text_detect.check_signage(path, allowed)` reads the
words actually rendered and flags any that the brief did not authorise.
`artreq_images.check_generated(paths, tones, event_type, signage)` routes Merge
to the signage check and every other type to the blanket text ban. The
authorised wording is read back out of the prompt — **write each sign in the
brief as `sign reads MARKET`**, one per building.

If a Merge brief declares **no** wording, the reviewer emits a **warning**, not
a pass. OCR matching is lenient about split or squashed words but never accepts
a word the brief did not mention. OCR is vendored into `pybackend_dist`; if it
is missing, `check_signage` returns `ok: None` ("could not check"), never a pass.
The plain structural text detector is blind to Merge signs and cannot
substitute.

---

## 8. Reference images

Seeded in `images/Merge/` — shipped share pictures, NOT concept art:

| File | Source |
|---|---|
| `merge2512_share2_shipped.png` | Merge2512 Merry Merge, scene 2 (Christmas, novelty buildings round a plaza) |
| `merge2606_share1_shipped.png` | Merge2606 Toy Towns, scene 1 (object-shaped houses, open plaza, bright ground) |

`MAX_REF_IMAGES` is **2** and `_images_in()` selects by modification time,
newest first — so any third image added to this folder can silently displace one
of these. Keep the folder at two.

**Quarantined** in `images/Merge/_excluded/`:

- the two Merge2606 **pre-production concept paintings** (plush kingdom, brick
  town) that seeded earlier versions of this file. Their finish is not the
  shipped finish; attaching them pulls renders away from the game's look.
- Toy Car World (Merge2606 p.8): carries circled callout numbers 1–8 stamped
  over the buildings, which teaches the model to draw callout markers.

The subfolder is what excludes them — `_images_in()` only lists files directly
inside `images/Merge/`. More shipped scenes for comparison:
`ResourcesAb/Events/Merge2511|2512|2603|2606|2608|2610/Texture/Map/img_share_*.png`.
