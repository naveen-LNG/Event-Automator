# Style Reference

Feeds the Nano Banana reference-image generation in the Ideation / Art
Requirement tabs.

| File | What it is |
|------|------------|
| `PROMPT_TEMPLATE_recovered.txt` | The studio's real production prompt, recovered from the PNG metadata of `2 - vC.png`. **The master architecture** — top-priority task first, then absolute requirements, render style, reference-finish clause, colour budget, palette discipline, scene, asymmetry, hero, character stage, numbered final check. |
| `ROOM_RENDER_STYLE.md` | The approved colour-render spec (09-17 feedback pass). Injected at the TOP of every image prompt. |
| `nano_banana_workflow.api.json` | The exact ComfyUI graph from the PNG metadata. |
| `nano_banana_workflow.ui.json` | The same workflow as exported from the ComfyUI editor. |
| `images/<Type>/` | Render-quality references, **resolved per event type**. |

## Per-type references

`art_reference.style_reference_images(events_root, event_type)` looks in
`images/<Type>/` first and only falls back to `images/` when a type has none of
its own.

This matters: one flat folder meant a **Sort** render was anchored on the **Sim**
hazelnut-roastery room — a frontal interior with a producer machine, nothing
like Sort's isometric habitat plates — and pulled the result off-model.

| Type | References | Notes |
|---|---|---|
| `Sort/` | 5 — the Sort2609 Pet Sanctuary zone floors | cat, dog, aquarium, reptile, aviary. See `Events/Sort/Art Requirements/Sort_zone_floor_rules.md`. |
| `Sim/` | 1 — the hazelnut roastery | This is the **pre-revision** `2 - vC.png`. The spec names `2 - vC rev.png` (approved) and `image (27) (1).png` (toy shop) — neither was available. |
| `Blocks/`, `Merge/`, `Dessert/` | none | They fall back to the shared folder. Dropping a shipped scene into each closes the gap. |

Max 6 images per generation (NanoBananaAIO exposes `image_1`..`image_6`).
