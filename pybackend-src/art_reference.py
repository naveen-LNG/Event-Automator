"""Grounding material for reference-image prompts.

Two sources, both real studio artefacts rather than invented description:

1. StyleReference/ROOM_RENDER_STYLE.md - the approved colour-render spec from
   the 09-17 feedback pass. Injected verbatim and FIRST, because section 8 of
   that document says top-of-prompt wins and the object spec gets ignored when
   buried.

2. ideationData/Events/<Type>/Art Requirements/*.pdf - the studio's real art
   requirement documents, which carry the visual language new prompts have to
   match.

The rest of the generator only ever read *.md, so it never opened these PDFs
at all - the per-type hints in art_requirement_generator were hand-derived once
and frozen. This module reads them directly, so dropping a new PDF into a
type's folder changes the grounding with no code edit.
"""

import json
import os

STYLE_DIR_NAME = "StyleReference"
STYLE_SPEC_NAME = "ROOM_RENDER_STYLE.md"
IMAGE_EXTS = (".png", ".jpg", ".jpeg", ".webp")
# NanoBananaAIO exposes image_1..image_6, but attaching that many pushes it to
# BLEND the references rather than imitate their style. With five Sort floors
# attached, one generation composited an entire reference room into the output -
# its pink striped wallpaper, arched window, cat tower and food bowls all landed
# on top of the scene actually requested. Measured: the output's top band came
# back (226,172,193) against the cat room's (243,179,203).
#
# Two references reproduced the style with no bleed at all (top band went to
# (148,182,177), the requested sea-green). Keep the cap low; style transfers from
# one or two examples, content leaks from five.
MAX_REF_IMAGES = 2

# How much distilled language to carry into a prompt: enough to set the house
# voice, short enough that it cannot bury the render spec above it.
MAX_FEEL_LINES = 8
MAX_PALETTE_LINES = 8
MAX_SILHOUETTE_LINES = 10


def style_dir(events_root):
    """StyleReference lives beside Events/, not inside it."""
    if not events_root:
        return ""
    return os.path.join(os.path.dirname(os.path.abspath(events_root)), STYLE_DIR_NAME)


def load_style_spec(events_root, event_type=None):
    """The approved render spec for this event type, verbatim.

    Looks for StyleReference/<Type>_RENDER_STYLE.md first and only falls back to
    ROOM_RENDER_STYLE.md when a type has no spec of its own.

    ROOM_RENDER_STYLE.md is titled "Sim room art" and demands flat toy-like
    vector, matte, and absolutely no material texture or specular. It was being
    injected for EVERY type, which flattened Sort renders into washed-out vector
    that looked nothing like the shipped Sort floors - those are soft-painted,
    glossy where the material is glossy, and do carry stylised material detail.
    The finish is genuinely per-type; only the no-outlines rule is shared.
    """
    d = style_dir(events_root)
    names = ([f"{event_type}_RENDER_STYLE.md"] if event_type else []) + [STYLE_SPEC_NAME]
    for name in names:
        try:
            with open(os.path.join(d, name), "r", encoding="utf-8") as f:
                return f.read()
        except Exception:
            continue
    return ""


def _images_in(d, limit):
    if not os.path.isdir(d):
        return []
    files = [os.path.join(d, n) for n in os.listdir(d)
             if n.lower().endswith(IMAGE_EXTS) and os.path.isfile(os.path.join(d, n))]
    files.sort(key=os.path.getmtime, reverse=True)
    return files[:limit]


def style_reference_images(events_root, event_type=None, limit=MAX_REF_IMAGES):
    """Render-quality references to attach as image inputs, newest first.

    Looks in StyleReference/images/<Type>/ first, and only falls back to the
    shared StyleReference/images/ folder when a type has none of its own.

    This used to read one flat folder for every event type, which meant a Sort
    generation would attach the Sim hazelnut-roastery room as its style anchor
    and pull the render off-model. Each type's look is genuinely different -
    Sort2609's five zone floors are isometric habitat dioramas, nothing like a
    Sim production room - so the reference has to be type-specific.
    """
    root = os.path.join(style_dir(events_root), "images")
    if event_type:
        own = _images_in(os.path.join(root, event_type), limit)
        if own:
            return own
    return _images_in(root, limit)


ZONE_DIR_NAME = "zones"
ZONE_REF_STANDALONE = True   # see zone_reference_images


def zone_reference_images(events_root, event_type=None, zone_number=None,
                          limit=MAX_REF_IMAGES):
    """References for ONE zone: its own pad first, then the type's style anchors.

    Looks in StyleReference/images/<Type>/zones/ for files named zone<N>_*, and
    falls straight back to style_reference_images() when a zone has none - so a
    type that has not been seeded per zone behaves exactly as before.

    Why this exists: the pad shape and proportions could not be held by prose.
    Briefs stating "roughly 2.1 : 1" produced 1.31, and one zone moved the wrong
    way entirely. Attaching the shipped art is what fixed the render style
    earlier, for the same reason - a reference beats a description of one - and
    the pad is geometry, which prose controls worst of all.

    A seeded zone is rendered against its pad ALONE. Topping the slot up with the
    type-level style anchor was tried and measured: with both attached the model
    blends them rather than taking geometry from one and finish from the other.
    Zone 4 moved most (0.77 -> 1.16 against a 1.26 target) and its anchor happened
    to agree with its pad; zone 1, whose anchor was a violet club against green
    terraces, did not move at all. The finish rules are carried in prose by the
    render spec, and prose has held finish reliably - it is geometry it cannot
    hold. So the one reference slot goes to geometry.

    Set ZONE_REF_STANDALONE False to restore the topped-up behaviour.
    """
    fallback = style_reference_images(events_root, event_type, limit)
    if not event_type or zone_number is None:
        return fallback
    d = os.path.join(style_dir(events_root), "images", event_type, ZONE_DIR_NAME)
    prefix = "zone%s_" % zone_number
    # Listed unlimited THEN filtered: the folder holds one file per zone, so
    # taking the newest `limit` first would drop the zone actually being asked for.
    own = [p for p in _images_in(d, None)
           if os.path.basename(p).lower().startswith(prefix)]
    if not own:
        return fallback
    out = own[:limit]
    if ZONE_REF_STANDALONE:
        return out[:1]
    for p in fallback:                      # top up with the type's style anchors
        if len(out) >= limit:
            break
        if p not in out:
            out.append(p)
    return out


def _pdf_text(path):
    try:
        from pypdf import PdfReader
    except ImportError:
        return ""
    try:
        return "\n".join((pg.extract_text() or "") for pg in PdfReader(path).pages)
    except Exception:
        return ""


def art_requirement_pdfs(events_root, event_type):
    """Every real art requirement PDF for a type, newest first."""
    d = os.path.join(events_root or "", event_type, "Art Requirements")
    if not os.path.isdir(d):
        return []
    files = [os.path.join(d, n) for n in os.listdir(d) if n.lower().endswith(".pdf")]
    files.sort(key=os.path.getmtime, reverse=True)
    return files


# ------------------------------------------------------- visual language ---
VL_SCHEMA = {
    "type": "object",
    "properties": {
        "feel": {"type": "array", "items": {"type": "string"},
                 "description": "Scene mood/atmosphere phrases, in English."},
        "palette": {"type": "array", "items": {"type": "string"},
                    "description": "Colour palettes actually used, in English."},
        "silhouette": {"type": "array", "items": {"type": "string"},
                       "description": "Building/prop shape-language phrases, in English."},
    },
    "required": ["feel", "palette", "silhouette"],
    "additionalProperties": False,
}

VL_SYSTEM = "\n".join([
    "You distil the visual language out of a mobile-game art requirement document "
    "so later image prompts match the studio's established look.",
    "These documents are inconsistent: some are English, some are Simplified Chinese, "
    "and each event type lays its tables out differently. Read whatever is there and "
    "normalise it.",
    "Rules:",
    "- Always answer in English, even when the source is Chinese.",
    "- Quote the document's own concrete language (materials, colours, moods). Never "
    "invent art direction the text does not support.",
    "- Drop table scaffolding, page numbers, dates, IDs and broken column fragments.",
    "- Each entry is one self-contained phrase. Skip anything under four words.",
    "- Return empty arrays rather than padding with guesses.",
])


def _cache_path(events_root, event_type):
    return os.path.join(style_dir(events_root), "_visual_language", event_type + ".json")


def _corpus_fingerprint(paths):
    """Filenames + sizes + mtimes: a new or edited PDF invalidates the cache, an
    unchanged folder reuses it, so this costs one model call per type per change."""
    parts = []
    for p in sorted(paths):
        try:
            st = os.stat(p)
            parts.append("%s:%d:%d" % (os.path.basename(p), st.st_size, int(st.st_mtime)))
        except OSError:
            continue
    return "|".join(parts)


def extract_visual_language(events_root, event_type, distil_fn=None, max_pdfs=3,
                            max_chars_per_pdf=14000):
    """House visual language for a type, distilled from its real art PDFs.

    A regex over label lines ("Overall feel:", "Palette:") was tried first and
    only worked for Merge - the Sim and Blocks documents are written in Chinese
    and Dessert/Sort use different headings, so it returned nothing for four of
    the five types and fragments like "2 P" for the fifth. The text goes to the
    model instead, which reads all of them, and the result is cached per type.

    distil_fn(system_prompt, user_message, schema) -> dict. When it is None or
    raises, a cached result is used if present, otherwise empty lists come back:
    grounding is an enhancement and must never block a document.
    """
    paths = art_requirement_pdfs(events_root, event_type)[:max_pdfs]
    empty = {"sources": [], "feel": [], "palette": [], "silhouette": [], "cached": False}
    if not paths:
        return empty

    fp = _corpus_fingerprint(paths)
    cpath = _cache_path(events_root, event_type)
    cached = None
    try:
        with open(cpath, "r", encoding="utf-8") as f:
            cached = json.load(f)
        if cached.get("fingerprint") == fp:
            cached["cached"] = True
            return cached
    except Exception:
        cached = None

    def _fallback():
        if cached:
            cached["cached"] = True
            cached["stale"] = True
            return cached
        return empty

    if distil_fn is None:
        return _fallback()

    chunks, used = [], []
    for path in paths:
        text = _pdf_text(path)
        if not text.strip():
            continue
        used.append(os.path.basename(path))
        chunks.append("--- %s ---\n%s" % (os.path.basename(path), text[:max_chars_per_pdf]))
    if not chunks:
        return empty

    try:
        got = distil_fn(VL_SYSTEM,
                        "Event type: %s\n\n%s" % (event_type, "\n\n".join(chunks)),
                        VL_SCHEMA) or {}
    except Exception:
        return _fallback()

    out = {
        "fingerprint": fp,
        "sources": used,
        "feel": [s for s in (got.get("feel") or []) if s][:MAX_FEEL_LINES],
        "palette": [s for s in (got.get("palette") or []) if s][:MAX_PALETTE_LINES],
        "silhouette": [s for s in (got.get("silhouette") or []) if s][:MAX_SILHOUETTE_LINES],
        "cached": False,
    }
    try:
        os.makedirs(os.path.dirname(cpath), exist_ok=True)
        with open(cpath, "w", encoding="utf-8") as f:
            json.dump(out, f, indent=2, ensure_ascii=False)
    except Exception:
        pass
    return out


def visual_language_block(events_root, event_type, distil_fn=None):
    """The distilled language, formatted for a prompt. '' when nothing found."""
    v = extract_visual_language(events_root, event_type, distil_fn=distil_fn)
    if not (v["feel"] or v["palette"] or v["silhouette"]):
        return ""
    src = ", ".join(v["sources"]) or "this event type's art requirement documents"
    parts = ["HOUSE VISUAL LANGUAGE - distilled from this event type's own shipped art "
             "requirement documents (%s). Match this voice; do not drift to generic "
             "mobile-game art." % src]
    if v["feel"]:
        parts.append("Scene moods previously approved:\n" +
                     "\n".join("  - " + s for s in v["feel"]))
    if v["palette"]:
        parts.append("Palettes previously approved:\n" +
                     "\n".join("  - " + s for s in v["palette"]))
    if v["silhouette"]:
        parts.append("Building silhouette language:\n" +
                     "\n".join("  - " + s for s in v["silhouette"]))
    return "\n\n".join(parts)


# ------------------------------------------------------------------ prompt --
# Bans that hold for every type. Each is here because a render failed on it.
_SHARED_FAILURES = (
    "No brush strokes, no paper grain, no noise, no bump. No soft-serve or swirled peaks "
    "of chocolate; they read as a poop emoji, so use solid chunks with crisp edges. No "
    "text, letters or numbers. No characters, people, animals, mascots or hands. Nothing "
    "mirrored."
)

# The rest is NOT shared, and pretending it was is a bug this fixes. The block
# below was written for the Sim/Merge renders and then injected into every type,
# so a Dessert render was told its references had outlines (they no longer do,
# they are shipped art), told to stay matte (Dessert scene art is glossy), and
# told not to draw producer fittings (Dessert has no producers). Same bug class
# as the render spec and the reference folder before them: one type's rules
# quietly applied to all of them.
_DEFAULT_FAILURES = (
    "The attached reference images have outlines; you must depart from them - render "
    "lineless, with no dark contour, keyline, ink line or sticker edge anywhere, and "
    "never a darker band ringing a shape. No material texture of any kind: no wood or "
    "plank grain, no woven or wicker weave, no fabric fibre, no stone speckle, no metal "
    "brushing, no glass refraction. "
    "Matte only - no wet plastic sheen, specular hotspots or mirror highlights. No "
    "buttons, control panels, gauges, dials, levers, valve wheels, hoppers, chutes, "
    "funnels, spouts, pipes or cranks on anything except the producer. "
) + _SHARED_FAILURES

# Dessert's references ARE the shipped art, so they are the target rather than
# something to depart from - and the shipped art is glossy and softly modelled,
# which the default block forbids.
_DESSERT_FAILURES = (
    "The attached reference images are the SHIPPED GAME ART for this event type. Match "
    "them: their camera, their finish, their framing. Do not depart from them, and do "
    "not flatten or simplify them.\n"
    "Render LINELESS - no dark contour, keyline, ink line or sticker edge anywhere, and "
    "never a darker band ringing a shape. An outlined render is the single most visible "
    "way to be off-model here.\n"
    "Shade with SMOOTH AIRBRUSHED GRADIENTS, not flat single-tone fills. A soft sheen on "
    "glass, glaze and polished floors is correct and wanted; what is wrong is scanned "
    "material texture - no wood or plank grain, no fabric fibre, no stone speckle.\n"
    "Colour is bright but GENTLE - pastel and candy, never neon or poster-saturated.\n"
    "This is ONE CONTINUOUS PARK, not an isolated diorama: never render the zone as a "
    "square plate floating on a white or empty background. "
) + _SHARED_FAILURES

# Blocks' negatives are the studio's own, from section 22 of Blocks Art.docx.
# They matter more than usual here because the default block contradicts the
# Blocks guide on three points: it demands lineless (the guide asks for soft edge
# contrast WITH an allowed slight darker edge and rim light), matte-only with no
# specular (the guide asks for broad soft highlights and volumetric glow), and no
# material texture (the guide asks for simplified stylised wood, metal and
# fabric). Injecting it into a Blocks render would brief against the house style.
_BLOCKS_FAILURES = (
    "The attached reference images are the SHIPPED GAME ART for this event type. Match their "
    "lighting, their depth and their finish.\n"
    "Avoid all of the following: photorealistic, realistic anatomy, realistic skin texture, pores, "
    "hyper-detailed textures, gritty, dark horror, sharp cel shading, thick black outlines, flat "
    "vector art, anime proportions, realistic proportions, long limbs, thin characters, excessive "
    "micro-details, harsh metallic reflections, complex realistic backgrounds, noisy textures, "
    "monochromatic lighting, hard-edged shadows, overly glossy plastic, realistic photography.\n"
    "EDGES: no thick black outline and no uniform ink contour - forms separate by colour, lighting "
    "and soft shadow. A slight darker edge, ambient occlusion or a rim light on an important "
    "silhouette is correct and wanted; a drawn outline is not.\n"
    "DETAIL IS GRADED, not uniform: highest in the foreground, medium-high on characters, medium in "
    "the midground, LOW in the background. Roughly 70% large and medium readable forms, 20% "
    "secondary detail, 10% tiny accents. Do not detail every part equally.\n"
    "No text, letters or numbers anywhere."
)

# Merge's negatives are the studio's own, from section 27 of Merge Art.docx,
# plus the lineless rule the Merge spec's section 6.2 depends on. The default
# block contradicts the document twice - it demands matte-only with no specular
# and bans material texture, where the document asks for glossy clean materials,
# PBR-inspired shading and controlled emissive glow. Same bug class as Blocks.
_MERGE_FAILURES = (
    "Avoid all of the following: photorealistic, realistic architecture, realistic proportions, "
    "gritty textures, dark horror aesthetic, realistic human anatomy, ultra-wide-angle "
    "perspective, fisheye lens, excessive perspective distortion, low-poly assets, flat 2D "
    "illustration, anime, manga, watercolour, painterly brush strokes, realistic environmental "
    "wear, dirty surfaces, excessive texture noise, desaturated palette, monochromatic palette, "
    "harsh shadows, dramatic photographic lighting, excessive bloom, overly thin geometry, "
    "microscopic details, cluttered composition, physically realistic industrial design.\n"
    "EDGES: the SCENE art is LINELESS - no dark contour, keyline or ink line around anything, and "
    "no darker band ringing a shape. Forms separate by colour, shading and ambient occlusion. "
    "(Item and producer ICONS are the opposite and do carry a rim - do not confuse the two.)\n"
    "LIGHT IS NOT EVEN: dark atmospheric surroundings with bright, readable, illuminated hero "
    "objects. Do NOT light the whole scene equally, and keep the background darker, softer and "
    "less saturated than the play area.\n"
    "GLOW IS CONTROLLED: soft and magical, never neon cyberpunk, never blown-out bloom.\n"
    "No text, letters or numbers anywhere except signage the brief explicitly authorises."
)

# Sort's negatives are the studio's own, from section 21 of Sort Art.docx, minus
# "empty rooms" - the measured shipped floors are 93% quiet surface and the
# density instinct is exactly what made the first generated attempt read
# cluttered. The measured composition targets are added as positive instruction
# because a negative list cannot express them.
_SORT_FAILURES = (
    "Avoid all of the following: photorealistic, realistic architecture, gritty, dark horror, "
    "harsh shadows, cinematic realism, ultra-realistic materials, rough textures, dirty surfaces, "
    "thin objects, sharp edges, heavy black outlines, comic book style, flat 2D illustration, "
    "watercolour, sketch, realistic human proportions, excessive perspective distortion, fisheye "
    "lens, dramatic contrast, muted colours, monochrome palette.\n"
    "EDGES: separation comes from colour contrast, soft shadow, ambient occlusion and slight "
    "bevels. A subtle darker contour on an important silhouette is fine; a thick black outline, "
    "an ink line or a cel-shaded border is not.\n"
    "SHADING, per object: bright top surface, mid-tone front face, darker side face, one small "
    "highlight, soft ambient occlusion underneath. Not detailed realistic lighting.\n"
    "MATERIALS: smooth plastic-like for props and appliances, MATTE painted for walls, floors and "
    "furniture, and glossy only as a selective accent on glass, screens, water and metal. Never "
    "heavy realistic PBR.\n"
    "*** THE PLATE IS MOSTLY EMPTY - THIS IS THE MOST COMMONLY MISSED RULE ***\n"
    "Roughly 90% of the floor plate is quiet, undecorated surface. Every object sits in its own "
    "clear pocket with floor visible all around it; NOTHING overlaps or occludes anything else, "
    "because the player has to read each object as a separate silhouette. Ten to fourteen objects "
    "in total, no more.\n"
    "The UPPER wall band carries at most two or three small fittings - never crowd the top. The "
    "MIDDLE band carries the main furniture mass and is the densest. The CENTRE LANE stays "
    "clearest; weight the left and right flanks.\n"
    "No text, letters or numbers anywhere."
)

FAILURES_BY_TYPE = {"Dessert": _DESSERT_FAILURES,
                    "Blocks": _BLOCKS_FAILURES,
                    "Merge": _MERGE_FAILURES,
                    "Sort": _SORT_FAILURES}
NAMED_FAILURES = _DEFAULT_FAILURES          # kept for callers that import it


def named_failures(event_type=""):
    """The failure block for this type - see FAILURES_BY_TYPE for why it varies."""
    return ("=== NAMED FAILURES TO AVOID ===\n"
            + FAILURES_BY_TYPE.get(event_type or "", _DEFAULT_FAILURES))


def build_scene_prompt(events_root, event_type, zone, doc_title="",
                       art_direction=None, style_spec=None, distil_fn=None):
    """Assemble the image prompt for one zone/scene.

    Order is deliberate, per ROOM_RENDER_STYLE.md section 8:
      1. the render spec        top-of-prompt wins; buried rules get ignored
      2. the named failures     generic bans lose to the reference image
      3. house visual language  from the type's real PDFs
      4. this specific scene    what actually changes run to run
    """
    spec = style_spec if style_spec is not None else load_style_spec(events_root, event_type)
    blocks = []

    if spec:
        blocks.append("=== APPROVED RENDER SPEC - FOLLOW EXACTLY ===\n" + spec.strip())

    blocks.append(named_failures(event_type))

    vl = visual_language_block(events_root, event_type, distil_fn=distil_fn)
    if vl:
        blocks.append("=== " + vl)

    ad = art_direction or {}
    scene = ["=== THIS SCENE ===",
             "Event: %s (%s)." % (doc_title or event_type, event_type)]
    if zone.get("zone_theme"):
        scene.append("Zone %s - %s." % (zone.get("zone_number", "?"), zone["zone_theme"]))
    if zone.get("palette_or_mood"):
        scene.append("Palette / mood: " + str(zone["palette_or_mood"]))
    if zone.get("reference_notes"):
        scene.append("Reference notes from the art requirement: " + str(zone["reference_notes"]))
    steps = zone.get("construction_steps") or []
    if steps:
        scene.append("Buildings / upgrades to show: " +
                     "; ".join(str(s) for s in steps[:10]))

    # The zone's BUILDINGS reach the scene prompt. They did not before: this
    # builder only ever read construction_steps, and a Merge zone carries its
    # eight buildings in `buildings` instead - so a Merge scene was briefed with
    # no idea which structures it was supposed to contain, and came back with
    # four or five vague ones. The count is stated as a number because "several
    # buildings" is exactly the instruction that produces five.
    blds = [b for b in (zone.get("buildings") or []) if b.get("name")]
    if blds:
        lines = ["EVERY ONE OF THESE %d BUILDINGS IS IN THIS SCENE, all visible and all "
                 "distinguishable from each other by silhouette alone. Do not draw fewer, do not "
                 "merge two into one, and do not substitute a generic structure for a named one:"
                 % len(blds)]
        for b in blds:
            bits = [str(b.get("name"))]
            if b.get("silhouette"):
                bits.append(str(b["silhouette"]))
            feats = [str(f) for f in (b.get("key_features") or []) if str(f).strip()]
            if feats:
                bits.append("recognisable by " + ", ".join(feats[:3]))
            lines.append("  %s. %s" % (b.get("building_number", "-"), " - ".join(bits)))
        levels = sum(int(b.get("max_level") or 1) for b in blds)
        lines.append("Upgrade points in this scene: %d across %d buildings - the taller, more "
                     "elaborate buildings are the ones that upgrade furthest, so give those a "
                     "visibly grander silhouette than the single-level ones." % (levels, len(blds)))
        scene.append("\n".join(lines))
    if ad.get("lean_into"):
        scene.append("Lean into: " + "; ".join(ad["lean_into"]))
    if ad.get("avoid"):
        scene.append("Avoid: " + "; ".join(ad["avoid"]))

    blocks.append("\n".join(scene))
    return "\n\n".join(blocks)
