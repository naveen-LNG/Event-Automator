"""
Art Requirement tab backend: once an Ideation theme is finalized, generates
N variations of the event's Art Requirement document, and regenerates a
single variation to address reviewer feedback. Mirrors ideation_generator.py's
architecture and reuses its provider/feedback-store primitives directly
rather than duplicating them.

Grounding, verified by actually opening real examples (not assumed from the
generic Event_Final_Art_Requirement.md template, whose content turned out to
be a broader multi-discipline "Final Spec" - not what the real per-event
"<Type>/Art Requirements/*.pdf" files in this repo actually look like):

- Merge2608 Art Requirement.pdf: Gameplay Overview table -> per-scene
  Producer/Product-Chain tables (2 upgrade chains per building) -> 3-scene
  visual progression -> Character Requirements.
- Dessert2605 Art Requirement.pdf: same Gameplay Overview + Scene backbone,
  but the gameplay-specific section is a Recipe/Ingredient requirement
  (common + secret recipe matrices), not producer chains.
- Blocks2604 Art Requirement.pdf: same backbone again, but the
  gameplay-specific section is a numbered before/after zone-restoration
  narrative plus a themed Construction Material Icon list.
- Sim2605 Art Requirement.pdf (a Hotel Sim / resort-management event): same
  backbone, but the gameplay-specific section is a Room/Furniture Set
  Requirement - themed decor/furniture icon sets tied to each room, plus a
  guest-service loop (guest NPCs arrive with requests; host/staff NPCs
  fulfill them).

- Sort2606 Art Requirement.pdf (a Sort/shelf-matching + building event,
  compressed under 20MB and read after Sim): same Gameplay Overview + Scene
  backbone (5 rooms, 9 construction steps each), but the gameplay-specific
  section is a Goods/Shelf Icon Requirement - a themed catalog of ~15
  small collectible item icons per zone (e.g. Living Room: Candle Jar,
  Photo Frame, Ceramic Vase...), since Sort's core loop is literally
  sorting/matching those icons on shelves.

So the common backbone (Gameplay Overview, per-zone Scene Requirements,
Character Requirements) is genuinely shared across types, but the
mechanic-specific middle section is NOT - forcing one rigid schema for it
would misrepresent Merge's chains, Dessert's recipes, Blocks' material
icons, Sim's decor sets, or Sort's shelf-icon catalog.
ART_REQUIREMENT_SCHEMA therefore keeps that one section as grounded
freeform markdown (see TYPE_GAMEPLAY_HINTS) rather than a fixed shape.

Sort2606 also surfaced two things that are NOT mechanic-specific -
genuinely generic to every event type - and are now their own schema
fields rather than folded into a freeform section: a global
art_direction_guidance (a "lean into X / avoid Y" moodboard statement
distinct from any single zone's palette - Sort2606 called this out
explicitly; the other four docs likely fold the same guidance into their
per-zone "Overall feel" text instead, so this field just makes it
first-class), and ui_icon_refresh_notes (which existing UI icons - badges,
entry button, currency token, event pack art - need re-skinning to this
event's palette, reusing existing shapes/positions).
"""
import sys
import os
import re
import json

import anthropic

import ideation_store as store
from ideation_generator import (
    KNOWN_TYPES, log, _read_template, _client, _call_claude, _call_ollama,
    write_feedback_digest, find_latest_doc,
)

MODEL_MAX_TOKENS_BASE = 5000
MODEL_MAX_TOKENS_PER_VARIATION = 4200  # richer document than a theme card - more sections, more zones

# Verified against the three real docs above; used only when this tool has
# not yet produced a prior Art Requirement for the type (see find_latest_doc
# in generate_variations(), which prefers this tool's own real past output
# once any exists).
TYPE_GAMEPLAY_HINTS = {
    "Merge": (
        "Merge has a FIXED resource shape, set by the Merge2603 template and "
        "confirmed against the shipped Merge2606 and Merge2608 documents and "
        "their game data. Fill the STRUCTURED fields, not just prose:\n"
        "  - zones: 3 scenes for a 14-16 day event, 4 for 17-21 days. Give each "
        "a scene_material - the one physical material every structure in it is "
        "built from - and leave main_tone empty, because Merge is polychrome.\n"
        "  - buildings: 8 per scene, spread across areas A, B and C. Each needs "
        "a silhouette, 3-5 key_features, and signage (exact wording, or an "
        "empty array if it carries no text).\n"
        "  - max_level: the upgrade stages per building. These must sum to "
        "about 15 PER SCENE (45 across three) - Merge2608 ships 7 buildings at "
        "level 1, 16 at level 2 and 2 at level 3. Vary them; do not give every "
        "building the same number.\n"
        "  - producers: exactly 4, each with exactly 2 chains of exactly 5 "
        "items. Level 1 of a producer is the vessel closed, level 2 the same "
        "vessel open with goods from both chains inside.\n"
        "Every chain names its family and holds to it, and each rung is a "
        "distinct more significant item - never more copies of the last one."
    ),
    "Dessert": (
        "Dessert has a FIXED resource shape, confirmed against the shipped "
        "Dessert2605 and Dessert2608 documents and the Dessert2608 game data. "
        "Fill the STRUCTURED fields, not just prose:\n"
        "  - zones: exactly 4, and their colour progression is fixed by the "
        "framework - zone 1 green, zone 2 pink, zone 3 amber, zone 4 purple. "
        "Give each a main_tone of WARM or COOL to match, and remember Dessert "
        "is ISOMETRIC and FLAT, never painted like Merge.\n"
        "  - buildings: 11, 11, 12, 12 across the four zones - 46 in all. Each "
        "needs a silhouette and 3-5 key_features, and EVERY building shows a "
        "culinary hook (a pastry case, a coffee machine, a recipe notebook) "
        "because this is a cooking event. Signage is ALWAYS an empty array: "
        "unlike Merge, Dessert buildings carry no text at all.\n"
        "  - recipes: 24 common - 4 categories of 6, one category per zone, "
        "each dish in its zone's colour family - plus the secret recipes "
        "(Dessert2608 shipped 15, Dessert2605 6).\n"
        "  - ingredients: 6 common and 5 secret. These are RAW components, not "
        "dishes, and they map onto a fixed slot vocabulary that is re-skinned "
        "each event rather than re-invented: flour, fruit, ice, matcha, meat, "
        "milk, mint, spice, sweet, flower, pumpkin. Name them in the event's "
        "own language over those slots.\n"
        "  - event_motif: the one motif stamped into EVERY recipe. "
        "Dessert2608's was the 5th-anniversary numeral 5, worked into all 24 "
        "common dishes - a 5-shaped cookie, a 5 piped on a pie, a 5 candle.\n"
        "Dessert has NO producers and NO merge chains - leave producers empty."
    ),
    "Blocks": (
        "A numbered before/after zone-restoration narrative (each zone "
        "starts in a 'damaged' or incomplete state and is revealed step by "
        "step, e.g. 'clear debris -> plant trees -> flowers bloom -> animals "
        "appear') plus a themed Construction Material Icon list (2-3 "
        "material icons per building step, thematically coherent)."
    ),
    "Sort": (
        "A Goods/Shelf Icon Requirement: for each zone, a themed catalog of "
        "roughly 15 small collectible item icons that populate the sort-shelf "
        "gameplay (e.g. a Living Room zone might need Candle Jar, Photo "
        "Frame, Ceramic Vase, Mini Plant Pot, Book Stack...) - each zone's "
        "catalog thematically distinct from the others."
    ),
    "Sim": (
        "A Room/Furniture Set Requirement: for each room or zone, a themed "
        "set of decor/furniture icons (e.g. 3-5 items per room, thematically "
        "coordinated) that fills in incrementally as the player upgrades "
        "it, plus guest-service NPCs who arrive with requests for "
        "host/staff NPCs to fulfill."
    ),
}

# A "zone" is this document's top-level VISUAL unit: one Merge scene, one Sort
# floor, one Sim room. It is deliberately not renamed to "scene" - the whole
# image pipeline, the per-zone regenerate path and every stored draft key off
# `zones`, and the rename would buy nothing.
#
# Note that Merge's own documents use "zone" for a SUBDIVISION of a scene (each
# scene has 3, labelled A/B/C, and buildings belong to one). That subdivision is
# carried on the building as `area`, so both meanings can coexist without
# renaming anything.

BUILDING_SCHEMA = {
    "type": "object",
    "properties": {
        "building_number": {"type": "integer"},
        "name": {"type": "string", "description": "e.g. 'Weaver's Round House'. A design name for the brief - it must NEVER be drawn on the building."},
        "area": {"type": "string", "description": "Which subdivision of the scene this sits in, e.g. 'A', 'B', 'C'. Empty for event types that do not subdivide."},
        "silhouette": {"type": "string", "description": "The shape, in one sentence. A building must be identifiable from its outline alone at thumbnail size."},
        "key_features": {
            "type": "array",
            "items": {"type": "string"},
            "description": "3-5 recognisability anchors, the form Merge2606 uses (关键识别点) - e.g. 'bear ears', 'button windows', 'patch fabric'. This is the checklist the finished art is verified against."
        },
        "max_level": {
            "type": "integer", "minimum": 1, "maximum": 3,
            "description": "How many upgrade stages this building has. The brief describes the FINAL level; earlier stages are derived from it. Merge2608 ships 45 stages across 25 buildings - 7 at 1, 16 at 2, 2 at 3, summing to 15 per scene. Use 1 where the type has no upgrade mechanic."
        },
        "signage": {
            "type": "array",
            "items": {"type": "string"},
            "description": "EXACT wording of any sign on this building, e.g. ['SWEET OVEN']. An EMPTY array is meaningful and is not the same as omitting it: it declares the building carries NO text, which the reviewer then enforces strictly."
        }
    },
    "required": ["building_number", "name", "area", "silhouette", "key_features", "max_level", "signage"],
    "additionalProperties": False
}

CHAIN_SCHEMA = {
    "type": "object",
    "properties": {
        "chain_number": {"type": "integer", "minimum": 1, "maximum": 2},
        "family": {
            "type": "string",
            "description": "The single category ALL five items belong to, e.g. 'willow basketry', 'seafood'. Mattel rejected a chain with a non-matching item at position 4: 应该统一为面包类. Naming the family up front is what keeps the chain pure."
        },
        "palette": {"type": "string", "description": "The one colour family this chain is built from, so it reads as a set."},
        "items": {
            "type": "array", "minItems": 5, "maxItems": 5,
            "items": {"type": "string"},
            "description": "The five rungs, level 1 to 5. Each is a DISTINCT, more significant item in the family, ending in the assembled or grand version - shipped example: bow tie, shoes, trousers, jacket, full suit. NEVER one-then-two-then-three of the same object: 而不是盲目的增加铃铛数量."
        }
    },
    "required": ["chain_number", "family", "palette", "items"],
    "additionalProperties": False
}

PRODUCER_SCHEMA = {
    "type": "object",
    "properties": {
        "producer_number": {"type": "integer"},
        "name": {"type": "string"},
        "palette": {"type": "string"},
        "closed_description": {"type": "string", "description": "Level 1: the vessel CLOSED."},
        "open_description": {"type": "string", "description": "Level 2: the SAME vessel OPEN, with goods from both its chains visible inside. Same shape, same materials, same fittings - only the lid and contents change."},
        "chains": {"type": "array", "minItems": 2, "maxItems": 2, "items": CHAIN_SCHEMA}
    },
    "required": ["producer_number", "name", "palette", "closed_description", "open_description", "chains"],
    "additionalProperties": False
}

INGREDIENT_SCHEMA = {
    "type": "object",
    "properties": {
        "name": {"type": "string", "description": "The event's own name for it, e.g. 'Carnival Sugar'."},
        "description": {"type": "string", "description": "The RAW component in one line - what it physically is before cooking. Not a dish."},
        "palette": {"type": "string"},
        "secret": {"type": "boolean", "description": "True for a secret-recipe ingredient, false for a common one. Dessert2608 shipped 6 common and 5 secret."}
    },
    "required": ["name", "description", "palette", "secret"],
    "additionalProperties": False
}

RECIPE_SCHEMA = {
    "type": "object",
    "properties": {
        "recipe_number": {"type": "integer"},
        "name": {"type": "string"},
        "zone_number": {"type": "integer", "description": "The zone whose colour family this dish belongs to. Dessert2608 ships 6 common recipes per zone."},
        "description": {"type": "string", "description": "The finished dish, plated, in one or two lines."},
        "palette": {"type": "string", "description": "This zone's colour family - zone 1 golden-brown baked goods, zone 2 pink frosted, zone 3 amber carnival, zone 4 violet drinks and jellies."},
        "ingredients": {
            "type": "array", "items": {"type": "string"},
            "description": "The ingredient names this dish needs, matching entries in the top-level ingredients list."
        },
        "secret": {"type": "boolean", "description": "True for a secret recipe (research-unlocked), false for a common one."}
    },
    "required": ["recipe_number", "name", "zone_number", "description", "palette", "ingredients", "secret"],
    "additionalProperties": False
}

ZONE_SCHEMA = {
    "type": "object",
    "properties": {
        "zone_number": {"type": "integer"},
        "zone_theme": {"type": "string", "description": "e.g. 'Tropical Rainforest', 'Rock Concert Stage'."},
        "palette_or_mood": {"type": "string"},
        "main_tone": {
            "type": "string", "enum": ["COOL", "WARM", ""],
            "description": "The dominant tone, where the type commits to one (Sort and Sim do; Merge is polychrome and leaves this empty). The reviewer checks the render against it - a large off-tone area means content was carried in from a reference image."
        },
        "scene_material": {
            "type": "string",
            "description": "The ONE physical material every structure in this zone is built from, e.g. 'folded paper and cut card', 'blown glass'. The strongest organising idea in the Merge documents; empty for types that do not work this way."
        },
        "construction_steps": {
            "type": "array",
            "items": {"type": "string"},
            "description": "Numbered visual build-up steps for types that reveal a zone in stages (Blocks). Where the type has discrete buildings, use `buildings` instead and leave this empty."
        },
        "buildings": {
            "type": "array",
            "items": BUILDING_SCHEMA,
            "description": "The buildable structures in this zone, each described at its FINAL level. Empty for types with no discrete buildings."
        },
        "reference_notes": {"type": "string", "description": "What reference art/mood imagery Art needs for this zone."}
    },
    "required": ["zone_number", "zone_theme", "palette_or_mood", "main_tone", "scene_material",
                 "construction_steps", "buildings", "reference_notes"],
    "additionalProperties": False
}

CHARACTER_SCHEMA = {
    "type": "object",
    "properties": {
        "name": {"type": "string"},
        "role": {"type": "string", "enum": ["Gift-giving/mechanic NPC", "Judge/special guest", "Ordinary guest", "Chef/host", "Other"]},
        "npc_id": {
            "type": "string",
            "description": "The favourability slot this fills: 101-104 for scene 1, 201-204 for scene 2, 301-304 for scene 3. Empty for characters outside that scheme. Merge2608 ships exactly these twelve."
        },
        "base_character": {
            "type": "string",
            "description": "WHICH EXISTING SkipBo cast member this is - hippo, raccoon, poodle, turtle, squirrel, rabbit, owl, monkey, pig, bear. These events RE-COSTUME the existing cast; they do not invent new animals. Merge2608 says so outright: 'reuse existing guests as boarding travelers', 'reuse existing guests in gala formalwear'. Leave empty only if this genuinely is a new character."
        },
        "costume": {
            "type": "string",
            "description": "The re-skin: what this cast member WEARS for this event, themed to its scene. This is the actual art requirement - the character underneath is already drawn."
        },
        "design_notes": {"type": "string"},
        "scene_actions": {"type": "array", "items": {"type": "string"}}
    },
    "required": ["name", "role", "npc_id", "base_character", "costume", "design_notes", "scene_actions"],
    "additionalProperties": False
}

EVENT_EXTRAS_SCHEMA = {
    "type": "object",
    "properties": {
        "hot_air_balloon": {
            "type": "string",
            "description": "Every Merge event needs at least one themed hot air balloon - 每次活动需要至少1种热气球. Describe its look in one or two sentences, themed to the event. Empty for types that have none."
        },
        "achievement_badge": {
            "type": "string",
            "description": "The event achievement badge: which motif from the event it is built from, e.g. Merge2604 used 'the Maypole and blossoms'. Empty if the type has none."
        },
        "task_icons": {
            "type": "array",
            "items": {"type": "string"},
            "description": "The 3-4 find-the-object task icons - 任务事件图标一共需要3-4种. One short object description each. Empty for types with none."
        }
    },
    "required": ["hot_air_balloon", "achievement_badge", "task_icons"],
    "additionalProperties": False
}

ART_REQUIREMENT_SCHEMA = {
    "type": "object",
    "properties": {
        "title": {"type": "string", "description": "e.g. '[Artistic Requirement] Merge2609 Sunday Garden Club'."},
        "gameplay_overview": {
            "type": "object",
            "properties": {
                "theme": {"type": "string"},
                "duration_summary": {"type": "string"},
                "narrative_packaging": {"type": "string"},
                "loop_summary": {"type": "string", "description": "This event type's fixed core loop, described briefly - do not invent a new loop."},
                "resource_planning": {"type": "string", "description": "Zone count, steps/buildings per zone, character count - a short plain-text summary."},
                "major_changes": {"type": "string", "description": "What differs from the standard template this run, or 'None - standard template.'"}
            },
            "required": ["theme", "duration_summary", "narrative_packaging", "loop_summary", "resource_planning", "major_changes"],
            "additionalProperties": False
        },
        "zones": {"type": "array", "minItems": 3, "maxItems": 6, "items": ZONE_SCHEMA},
        "producers": {
            "type": "array",
            "items": PRODUCER_SCHEMA,
            "description": "Merge producers. The template fixes the shape: 4 producers, each with 2 chains, each chain 5 levels - 每次活动固定需要4组生产器，每组生产器有2条生产链，每个生产链有5级产品. That is 8 producer icons and 40 item icons. Empty for event types with no producers."
        },
        "recipes": {
            "type": "array",
            "items": RECIPE_SCHEMA,
            "description": "Dessert recipes. The template fixes the common half: 4 categories of 6, one category per zone, in that zone's colour family - 24 in all. The secret half varies (Dessert2608 shipped 15, Dessert2605 6). Empty for event types with no recipes."
        },
        "ingredients": {
            "type": "array",
            "items": INGREDIENT_SCHEMA,
            "description": "Dessert ingredients: the raw components the recipes are built from. Dessert2608 shipped 11 - 6 common and 5 secret. Empty for event types with no recipes."
        },
        "event_motif": {
            "type": "string",
            "description": "The one recurring motif this event stamps on its assets - Dessert2608's was the 5th-anniversary numeral 5, which appears worked into every one of its 24 common recipes. One short phrase. Empty if the event has no such motif."
        },
        "gameplay_specific_content": {
            "type": "string",
            "description": "The mechanic-specific art requirement section for this event type, as markdown - see the type-specific guidance given. If no verified historical structure exists for this type, this MUST start with '[INFERRED - no historical example for this type; confirm structure with Art Lead before use.]'."
        },
        "characters": {"type": "array", "items": CHARACTER_SCHEMA},
        "art_direction_guidance": {
            "type": "object",
            "properties": {
                "lean_into": {"type": "array", "items": {"type": "string"}, "description": "Mood/style references to lean into - e.g. 'Scandinavian modern interiors', 'soft sunlight', 'oak wood + beige fabrics'. Global, distinct from any single zone's palette."},
                "avoid": {"type": "array", "items": {"type": "string"}, "description": "What NOT to do - e.g. 'dark realism', 'heavy construction machinery', 'mansion extravagance'."}
            },
            "required": ["lean_into", "avoid"],
            "additionalProperties": False
        },
        "event_extras": EVENT_EXTRAS_SCHEMA,
        "ui_icon_refresh_notes": {
            "type": "string",
            "description": "Which existing UI icons/panels (achievement badge, entry icon, restock/currency icon, event pack art) need re-skinning to this event's palette - reuse existing icon shapes and positions, only the motif/color changes."
        },
        "notes": {"type": "string"}
    },
    "required": ["title", "gameplay_overview", "zones", "producers", "recipes", "ingredients", "event_motif", "gameplay_specific_content", "characters", "event_extras", "art_direction_guidance", "ui_icon_refresh_notes", "notes"],
    "additionalProperties": False
}


def _variations_schema(count):
    return {
        "type": "object",
        "properties": {
            "variations": {"type": "array", "minItems": count, "maxItems": count, "items": ART_REQUIREMENT_SCHEMA}
        },
        "required": ["variations"],
        "additionalProperties": False
    }


REGENERATE_SCHEMA = {
    "type": "object",
    "properties": {
        "variation": ART_REQUIREMENT_SCHEMA,
        "extracted_lesson": {
            "type": "string",
            "description": "One generalizable sentence capturing the rule behind this feedback. Empty string if it carries no generalizable rule."
        }
    },
    "required": ["variation", "extracted_lesson"],
    "additionalProperties": False
}


def _build_system_prompt(event_type, known_lessons_block, final_spec_template_text, past_doc_text):
    hint = TYPE_GAMEPLAY_HINTS.get(event_type, "[NO VERIFIED EXAMPLE - infer cautiously and flag as such.]")
    past_block = (
        f"\n\nHere is the most recent Art Requirement document this studio actually "
        f"produced for a {event_type} event - match its level of concreteness and "
        f"structure closely (do not copy its specific content):\n---\n{past_doc_text[:8000]}\n---"
        if past_doc_text else
        f"\n\nNo prior Art Requirement document exists yet for {event_type} in this "
        f"tool's history - this will be the first. Be extra explicit about structure "
        f"since there is no concrete precedent to match."
    )
    template_block = (
        f"\n\nThe studio's general Final Spec template's Art/Visual Spec acceptance "
        f"criteria (for the closing checklist section only, not the main document "
        f"structure):\n---\n{final_spec_template_text[:3000]}\n---"
        if final_spec_template_text else ""
    )
    return (
        "You are an experienced art production lead at a mobile card-game studio "
        "(Skip-Bo) writing an Art Requirement handoff document - the real internal "
        "spec the Art team builds concept art from, once an event's Ideation "
        "1-Pager theme has already been approved. This is NOT a pitch document; "
        "the theme is already decided. Your job is translating that theme into "
        "concrete, buildable art direction.\n\n"
        "Structure, verified against this studio's real Art Requirement documents:\n"
        "1. Gameplay Overview: theme, duration, narrative packaging, this event "
        "type's fixed gameplay loop (describe it, never invent a new one), "
        "resource planning (zone/building counts), and any changes from the "
        "standard template.\n"
        "2. Scene Requirements: 3-6 zones, each with a theme, palette/mood, an "
        "ordered list of construction/reveal steps or buildings, and reference "
        "notes for what art needs to see.\n"
        "3. A gameplay-specific art requirement section unique to this event "
        f"type. For {event_type}: {hint}\n"
        "4. Character Requirements: any mechanic-tied NPCs (e.g. a gift-giving "
        "NPC, or a themed staff role per zone like an Interior Hostess/Chef/"
        "Gardener), judges/special guests, and ordinary reused guests - each "
        "with a design note and their in-scene actions.\n"
        "5. Art Direction Guidance: a global 'lean into X / avoid Y' moodboard "
        "statement (references, materials, lighting to lean into; aesthetics "
        "to avoid) - distinct from any single zone's palette, applies to the "
        "whole document.\n"
        "6. UI Icon Refresh Notes: which existing UI icons (achievement badge, "
        "entry icon, restock/currency icon, event pack art) need re-skinning "
        "to this event's palette - reuse existing shapes/positions, only the "
        "motif changes.\n\n"
        "Never invent asset delivery dates, file-naming conventions, or budget "
        "numbers - those are Production's job, not Art direction's.\n"
        f"{known_lessons_block}"
        f"{past_block}"
        f"{template_block}"
    )


def _generate_variations(provider, payload, system_prompt, user_message, count):
    schema = _variations_schema(count)
    if provider == "ollama":
        result = _call_ollama(
            payload.get("ollama_base_url", ""),
            payload.get("ollama_model") or "qwen3:32b",
            system_prompt, user_message, schema
        )
    else:
        max_tokens = min(MODEL_MAX_TOKENS_BASE + MODEL_MAX_TOKENS_PER_VARIATION * count, 100000)
        result = _call_claude(payload.get("api_key", ""), system_prompt, user_message, schema, max_tokens)
    return result["variations"]


def _regenerate_variation(provider, payload, system_prompt, user_message):
    if provider == "ollama":
        return _call_ollama(
            payload.get("ollama_base_url", ""),
            payload.get("ollama_model") or "qwen3:32b",
            system_prompt, user_message, REGENERATE_SCHEMA
        )
    return _call_claude(payload.get("api_key", ""), system_prompt, user_message, REGENERATE_SCHEMA, 8500)


def _zone_image_md(doc, zone, md_dir):
    """Markdown for a zone's generated reference image, if one exists.

    Written as a path RELATIVE to the .md file so the document keeps working
    when the ideationData folder is synced to another machine or a different
    OneDrive root - an absolute path would break for every other designer.
    """
    path = (zone or {}).get("image_path") or ""
    if not path or not os.path.exists(path):
        return ""
    try:
        rel = os.path.relpath(path, md_dir).replace("\\", "/")
    except ValueError:                      # different drive - no relative path exists
        rel = path.replace("\\", "/")
    alt = "Zone %s - %s" % (zone.get("zone_number", "?"), zone.get("zone_theme", ""))
    return "\n![%s](%s)\n" % (alt.replace("]", ")"), rel)


def _buildings_md(zone):
    """The buildings table for one zone, or nothing if the type has none.

    Level and Signage are columns rather than prose because both are consumed:
    Level drives how many upgrade stages get rendered, and Signage is the exact
    wording the reviewer checks the finished art against - with an empty cell
    meaning "this building carries no text", which is enforced strictly.
    """
    blds = zone.get("buildings") or []
    if not blds:
        return ""
    total = sum(int(b.get("max_level") or 1) for b in blds)
    rows = "\n".join(
        "| {n} | {name} | {area} | {lv} | {sil} | {feat} | {sign} |".format(
            n=b.get("building_number", ""), name=b.get("name", ""),
            area=b.get("area", "") or "—", lv=b.get("max_level", 1),
            sil=(b.get("silhouette", "") or "").replace("|", "/"),
            feat=", ".join(b.get("key_features") or []).replace("|", "/"),
            sign=", ".join(b.get("signage") or []) or "— none —")
        for b in blds)
    return (
        f"**Buildings ({len(blds)}) — {total} upgrade stages in this zone**\n\n"
        "| # | Building | Area | Levels | Silhouette | Key Identification Points | Signage |\n"
        "|---|---|---|---|---|---|---|\n"
        f"{rows}\n\n")


def _producers_md(doc):
    """The producer and chain tables, or nothing if the type has no producers."""
    prods = doc.get("producers") or []
    if not prods:
        return ""
    out = [f"**Producers ({len(prods)})** — "
           f"{len(prods) * 2} producer icons, "
           f"{sum(len(p.get('chains') or []) for p in prods) * 5} item icons\n"]
    for p in prods:
        out.append(f"\n#### Producer {p.get('producer_number')} — {p.get('name')}\n")
        out.append(f"**Palette:** {p.get('palette')}\n")
        out.append(f"- **Level 1 (closed):** {p.get('closed_description')}")
        out.append(f"- **Level 2 (open):** {p.get('open_description')}\n")
        out.append("| Chain | Family | 1 | 2 | 3 | 4 | 5 |")
        out.append("|---|---|---|---|---|---|---|")
        for ch in (p.get("chains") or []):
            items = list(ch.get("items") or []) + [""] * 5
            out.append("| {c} | {f} | {i} |".format(
                c=ch.get("chain_number", ""), f=ch.get("family", ""),
                i=" | ".join(x.replace("|", "/") for x in items[:5])))
        out.append("")
    return "\n".join(out) + "\n"


def _recipes_md(doc):
    """The recipe and ingredient tables, or nothing if the type has no recipes."""
    recipes = doc.get("recipes") or []
    ings = doc.get("ingredients") or []
    if not recipes and not ings:
        return ""
    out = []
    if ings:
        common = [i for i in ings if not i.get("secret")]
        secret = [i for i in ings if i.get("secret")]
        out.append(f"**Ingredients ({len(ings)})** — {len(common)} common, {len(secret)} secret\n")
        out.append("| Ingredient | Secret | Description |")
        out.append("|---|---|---|")
        for i in ings:
            out.append("| {n} | {s} | {d} |".format(
                n=i.get("name", ""), s="yes" if i.get("secret") else "",
                d=str(i.get("description", "")).replace("|", "/")))
        out.append("")
    if recipes:
        common = [r for r in recipes if not r.get("secret")]
        secret = [r for r in recipes if r.get("secret")]
        motif = (doc.get("event_motif") or "").strip()
        out.append(f"\n**Recipes ({len(recipes)})** — {len(common)} common, {len(secret)} secret")
        if motif:
            out.append(f"\nEvery recipe carries the event motif: **{motif}**.\n")
        else:
            out.append("")
        for label, group in (("Common", common), ("Secret", secret)):
            if not group:
                continue
            out.append(f"\n#### {label} recipes ({len(group)})\n")
            out.append("| # | Zone | Dish | Ingredients |")
            out.append("|---|---|---|---|")
            for r in group:
                out.append("| {n} | {z} | {d} | {i} |".format(
                    n=r.get("recipe_number", ""), z=r.get("zone_number", ""),
                    d=("**%s** — %s" % (r.get("name", ""), r.get("description", ""))).replace("|", "/"),
                    i=", ".join(str(x).replace("|", "/") for x in (r.get("ingredients") or []))))
            out.append("")
    return "\n".join(out) + "\n"


def _render_markdown(doc, event_type, status_line, author_name="", md_dir=None):
    go = doc.get("gameplay_overview") or {}
    md_dir = md_dir or os.getcwd()
    zones_md = ""
    for z in (doc.get("zones") or []):
        steps = "\n".join(f"  {i + 1}. {s}" for i, s in enumerate(z.get("construction_steps") or []))
        zones_md += (
            f"\n### Zone {z.get('zone_number')} — {z.get('zone_theme')}\n\n"
            f"**Palette/Mood:** {z.get('palette_or_mood')}\n\n"
        )
        if z.get("scene_material"):
            zones_md += f"**Scene Material:** {z['scene_material']}\n\n"
        if z.get("main_tone"):
            zones_md += f"**Main Tone:** {z['main_tone']}\n\n"
        if steps:
            zones_md += f"**Construction/Reveal Steps:**\n{steps}\n\n"
        zones_md += _buildings_md(z)
        zones_md += (
            f"**Reference Notes:** {z.get('reference_notes')}\n"
            f"{_zone_image_md(doc, z, md_dir)}"
        )
    chars_md = "\n".join(
        f"| {c.get('name')} | {c.get('role')} | {c.get('design_notes')} | {', '.join(c.get('scene_actions') or [])} |"
        for c in (doc.get("characters") or [])
    )
    adg = doc.get("art_direction_guidance") or {}
    lean_into_md = "\n".join(f"- {x}" for x in (adg.get("lean_into") or []))
    avoid_md = "\n".join(f"- {x}" for x in (adg.get("avoid") or []))

    by_line = f"\n**By:** {author_name}" if author_name else ""

    return f"""# {doc.get('title')}

**Status:** {status_line}{by_line}

## 1. Gameplay Overview

| Field | Value |
|---|---|
| Theme | {go.get('theme')} |
| Duration | {go.get('duration_summary')} |
| Narrative Packaging | {go.get('narrative_packaging')} |
| Core Loop | {go.get('loop_summary')} |
| Resource Planning | {go.get('resource_planning')} |
| Changes from Standard Template | {go.get('major_changes')} |

## 2. Scene Requirements
{zones_md}

## 3. Gameplay-Specific Art Requirement ({event_type})

{_producers_md(doc)}{_recipes_md(doc)}
{doc.get('gameplay_specific_content')}

## 4. Character Requirements

| Name | Role | Design Notes | Scene Actions |
|---|---|---|---|
{chars_md}

## 5. Art Direction Guidance

**Lean into:**
{lean_into_md}

**Avoid:**
{avoid_md}

## 6. UI Icon Refresh Notes

{doc.get('ui_icon_refresh_notes')}

## Notes

{doc.get('notes')}

---
AI-drafted via the Event Automator Art Requirement tab. Treat as a first draft for Art Lead review, not an approved spec.
"""


def _write_variation_draft(events_root, event_type, doc, index, author_name=""):
    if not events_root:
        return None
    title = doc.get("title") or "Untitled"
    ref = re.sub(r'[\\/:*?"<>|]', '_', title)[:60]
    draft_dir = os.path.join(events_root, event_type, "Art Requirements", "Drafts")
    os.makedirs(draft_dir, exist_ok=True)
    path = os.path.join(draft_dir, f"{ref} Variation {index + 1}.md")
    with open(path, "w", encoding="utf-8") as f:
        f.write(_render_markdown(doc, event_type,
                                 f"Draft variation {index + 1} — under review, not yet selected",
                                 author_name, md_dir=draft_dir))
    return path


def generate_main():
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    if len(sys.argv) < 2:
        print(json.dumps({"error": "Usage: artreq_generate <payload.json>"}))
        sys.exit(1)

    with open(sys.argv[1], "r", encoding="utf-8-sig") as f:
        payload = json.load(f)

    event_type = payload.get("event_type", "")
    ideation_theme = payload.get("ideation_theme") or {}
    events_root = payload.get("events_root", "")
    provider = payload.get("provider") or "anthropic"
    author_id = payload.get("author_id", "")
    author_name = payload.get("author_name", "")
    try:
        variation_count = max(1, min(10, int(payload.get("variation_count", 3))))
    except (TypeError, ValueError):
        variation_count = 3

    if event_type not in KNOWN_TYPES:
        print(json.dumps({"error": f"Unknown event type: {event_type}"}))
        sys.exit(1)
    if not ideation_theme.get("theme_name"):
        print(json.dumps({"error": "No finalized Ideation theme supplied - select one as Final first."}))
        sys.exit(1)

    try:
        lessons_block = store.lessons_prompt_block(events_root, event_type)
        final_spec_template_text = _read_template(events_root, "Event_Final_Art_Requirement.md")
        art_req_folder = os.path.join(events_root, event_type, "Art Requirements") if events_root else ""
        _, past_doc_text = find_latest_doc(art_req_folder)

        user_message = (
            f"Finalized Ideation theme to translate into art direction:\n"
            f"{json.dumps(ideation_theme, ensure_ascii=False)}\n\n"
            f"Generate exactly {variation_count} distinct Art Requirement variations "
            f"for this theme - different concrete visual executions of the SAME "
            f"approved theme (not different themes). Vary zone composition, "
            f"character designs, and specific reveal/construction steps across "
            f"variations; keep them all faithful to the theme above."
        )

        system_prompt = _build_system_prompt(event_type, lessons_block, final_spec_template_text, past_doc_text)
        variations = _generate_variations(provider, payload, system_prompt, user_message, variation_count)

        draft_paths = []
        if events_root:
            for i, v in enumerate(variations):
                p = _write_variation_draft(events_root, event_type, v, i, author_name)
                if p:
                    draft_paths.append(p)

        print(json.dumps({"variations": variations, "provider": provider, "draft_paths": draft_paths}, ensure_ascii=False))
    except anthropic.AuthenticationError:
        print(json.dumps({"error": "Invalid Anthropic API key. Check Settings."}))
        sys.exit(1)
    except anthropic.RateLimitError:
        print(json.dumps({"error": "Rate limited by the Anthropic API. Try again shortly."}))
        sys.exit(1)
    except anthropic.APIConnectionError:
        print(json.dumps({"error": "Could not reach the Anthropic API. Check network access."}))
        sys.exit(1)
    except Exception as e:
        print(json.dumps({"error": str(e)}))
        sys.exit(1)


def regenerate_main():
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    if len(sys.argv) < 2:
        print(json.dumps({"error": "Usage: artreq_regenerate <payload.json>"}))
        sys.exit(1)

    with open(sys.argv[1], "r", encoding="utf-8-sig") as f:
        payload = json.load(f)

    event_type = payload.get("event_type", "")
    events_root = payload.get("events_root", "")
    provider = payload.get("provider") or "anthropic"
    author_id = payload.get("author_id", "")
    author_name = payload.get("author_name", "")
    variation = payload.get("variation")
    variation_index = payload.get("variation_index", 0)
    feedback_text = (payload.get("feedback") or "").strip()

    if not variation:
        print(json.dumps({"error": "No variation supplied to regenerate."}))
        sys.exit(1)
    if not feedback_text:
        print(json.dumps({"error": "No feedback supplied."}))
        sys.exit(1)

    try:
        lessons_block = store.lessons_prompt_block(events_root, event_type)
        final_spec_template_text = _read_template(events_root, "Event_Final_Art_Requirement.md")
        art_req_folder = os.path.join(events_root, event_type, "Art Requirements") if events_root else ""
        _, past_doc_text = find_latest_doc(art_req_folder)

        user_message = (
            f"Here is the current Art Requirement variation draft:\n"
            f"{json.dumps(variation, ensure_ascii=False)}\n\n"
            f"Reviewer feedback to address:\n\"{feedback_text}\"\n\n"
            "Rewrite it so it fully addresses this feedback - every point, not just "
            "the easiest one. Keep everything the feedback didn't ask to change. "
            "Also extract one short, generalizable lesson from this feedback."
        )

        system_prompt = _build_system_prompt(event_type, lessons_block, final_spec_template_text, past_doc_text)
        result = _regenerate_variation(provider, payload, system_prompt, user_message)
        new_variation = result["variation"]

        lesson = (result.get("extracted_lesson") or "").strip()
        draft_path = None
        ref_name = payload.get("ref_name") or variation.get("title", "")
        if events_root:
            store.write_feedback_entry(
                events_root, event_type, "artreq", ref_name,
                variation.get("title", ""), feedback_text, lesson,
                author_id, author_name,
                revised_title=new_variation.get("title", ""),
            )
            write_feedback_digest(events_root, event_type, ref_name, "artreq", "Art Requirement")
            draft_path = _write_variation_draft(events_root, event_type, new_variation, variation_index, author_name)

        print(json.dumps({"variation": new_variation, "provider": provider, "draft_path": draft_path}, ensure_ascii=False))
    except anthropic.AuthenticationError:
        print(json.dumps({"error": "Invalid Anthropic API key. Check Settings."}))
        sys.exit(1)
    except anthropic.RateLimitError:
        print(json.dumps({"error": "Rate limited by the Anthropic API. Try again shortly."}))
        sys.exit(1)
    except anthropic.APIConnectionError:
        print(json.dumps({"error": "Could not reach the Anthropic API. Check network access."}))
        sys.exit(1)
    except Exception as e:
        print(json.dumps({"error": str(e)}))
        sys.exit(1)


def save_final_main():
    """No LLM call - just re-renders the already-generated variation with a
    Final status and writes it to the real (non-Drafts) location, matching
    the studio's actual naming convention (e.g. 'Merge2608 Art
    Requirement.pdf' -> 'Merge2609 Art Requirement.md')."""
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    if len(sys.argv) < 2:
        print(json.dumps({"error": "Usage: artreq_save_final <payload.json>"}))
        sys.exit(1)

    with open(sys.argv[1], "r", encoding="utf-8-sig") as f:
        payload = json.load(f)

    event_type = payload.get("event_type", "")
    events_root = payload.get("events_root", "")
    ref_name = payload.get("ref_name", "")
    variation = payload.get("variation")

    if not variation or not events_root or not ref_name:
        print(json.dumps({"error": "Missing variation, events_root, or ref_name."}))
        sys.exit(1)

    try:
        ref = re.sub(r'[\\/:*?"<>|]', '_', ref_name)
        final_dir = os.path.join(events_root, event_type, "Art Requirements")
        os.makedirs(final_dir, exist_ok=True)
        path = os.path.join(final_dir, f"{ref} Art Requirement.md")
        selector = payload.get("author_name", "")
        status = "Selected as Final (draft — pending Art Lead review)"
        if selector:
            status += f" · selected by {selector}"
        with open(path, "w", encoding="utf-8") as f:
            f.write(_render_markdown(variation, event_type, status, selector, md_dir=final_dir))
        print(json.dumps({"path": path}))
    except Exception as e:
        print(json.dumps({"error": str(e)}))
        sys.exit(1)
