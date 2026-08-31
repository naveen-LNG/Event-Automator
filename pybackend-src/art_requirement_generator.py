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
import copy
import json

import anthropic

import ideation_store as store
# The portrait band is shared with the renderer rather than restated, so the two
# cannot drift. Shape itself is deliberately NOT an enum - see the goods block in
# asset_prompts.py: a fixed list of shapes caps the catalogue at that many
# distinct objects, and five floors drawing from one list of fifteen produced a
# set that repeated down every column.
from asset_prompts import GOODS_ASPECT_RANGE as _GOODS_ASPECT
from ideation_generator import (
    KNOWN_TYPES, log, _read_template, _client, _call_claude, _call_ollama,
    _ollama_supports_vision, write_feedback_digest, find_latest_doc,
    ollama_timeout_for,
)

# The production team's own Art Requirement template, which is a DIFFERENT
# document from Event_Final_Art_Requirement.md. That file is a multi-discipline
# "Final Spec" - it carries Engineering, Product/Analytics, Economy and QA
# sections - and grounding art requirements in it was wrong: it does not contain
# the shape an Art Requirement actually has. The real template arrived as a PDF
# and is installed beside it, converted, so the model is grounded in the document
# the art team is actually given.
ART_REQ_TEMPLATE = "Event_Art_Requirement_Template.md"

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
        "Dessert has NO producers and NO merge chains - leave producers empty. "
        "Its buildings have NO upgrade levels either: unlike Merge, a Dessert "
        "building is drawn once and never upgraded. Two structures ARE an "
        "exception and Dessert2605 spends a page on them - the Exhibition Area, "
        "which has 3 upgrade stages as the chef's level rises with only the "
        "surrounding architecture and the chef-level sign changing, and the "
        "bronze/silver/gold trophy. Describe both in "
        "gameplay_specific_content; they are named structures, not a level on "
        "every building."
    ),
    "Blocks": (
        "Blocks has a FIXED resource shape, confirmed against the shipped "
        "Blocks2604 Art Requirement document and the blocks2604 game data. "
        "Fill the STRUCTURED fields, not just prose:\n"
        "  - zones: 5 scene areas for a 14-day event, scaled by duration. Each "
        "is one place, restored from a before-state.\n"
        "  - construction_steps: 6-8 per zone - 每个区域有6-8个建造步骤. These "
        "ARE the buildable content: numbered, each one a visible change to the "
        "scene, e.g. 'clear the dead wood, water appears -> plant the left-hand "
        "trees -> flowers open in the foreground -> the birds arrive'. "
        "blocks2604 ships 38 across five scenes: 7, 7, 8, 8, 8.\n"
        "  - before_state: every zone needs one, and the five must differ in "
        "KIND - shut, bare, dormant, unlit, stripped - not merely in wording. A "
        "before-state must not carry the scene's finished light.\n"
        "  - construction_materials: a themed icon set. Each step consumes 2-3 "
        "of them - 每个建造步骤会用2-3种材料 - so 8-10 icons covers an event, and "
        "they must share one palette and read as one family.\n"
        "  - characters: the gift NPCs. Blocks levels have a 赠礼机制 - every N "
        "clears an NPC appears with a gift - and the same NPC also appears in "
        "the restored scene, which is what ties the two halves together.\n"
        "Blocks has NO buildings, NO upgrade levels and NO A/B/C areas. Those "
        "belong to Merge. A Blocks scene is built once, step by step, and never "
        "upgraded: the game data carries a build animation and a material cost "
        "per step and no level field at all. Leave buildings empty."
    ),
    "Sort": (
        "Sort has a FIXED resource shape, confirmed against the shipped Sort2606 "
        "and Sort2609 documents. Fill the STRUCTURED fields, not just prose:\n"
        "  - zones: 5 floors, each a room with its own use - Sort2606 ran 5 "
        "phases of home decoration, Sort2609 five wings of a pet cafe.\n"
        "  - construction_steps: 9 per floor, 45 across the event. Sort2609 "
        "calls these '9 Construction Buildings Per Floor'; Sort2606 calls the "
        "same thing '9 construction steps' per room. They are steps, built once "
        "in order - do NOT read the word 'buildings' there as Merge's "
        "upgradeable buildings, which is exactly the mistake this wording "
        "invites.\n"
        "  - goods: roughly 15 small collectible icons per floor, and every one "
        "distinct from every other in the event in BOTH silhouette and colour - "
        "no two pots, no four flat rectangles, no floor painted in one hue. "
        "This is the rule that has actually failed in production.\n"
        "Sort has NO buildings and NO upgrade levels of its own: the only "
        "'level' in a Sort document is the sort-shelf PVE level the player "
        "enters, which is gameplay, not art. Leave buildings empty."
    ),
    "Sim": (
        "Confirmed against the sim2604 game data: 3 scenes holding 12 rooms "
        "between them, each room carrying a pendantList of decor and chaining "
        "to the next through nextRoom.\n"
        "  - zones: rooms. Each has a use, a guest capacity and a decor set.\n"
        "  - construction_steps: how the room fills in, numbered - 3-5 decor or "
        "furniture pieces per room, thematically coordinated.\n"
        "  - characters: guest-service NPCs who arrive with requests, plus the "
        "host/staff who fulfil them.\n"
        "Sim rooms unlock and fill; individual pieces are NOT upgraded through "
        "levels. Sim has NO buildings and NO upgrade levels - leave buildings "
        "empty."
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

MATERIAL_SCHEMA = {
    "type": "object",
    "properties": {
        "name": {"type": "string", "description": "e.g. 'Driftwood Plank', 'Coral Brick'."},
        "description": {"type": "string", "description": "The raw material as a resource icon - a stack, a bundle, a worked block. Not a finished structure."},
        "palette": {"type": "string"}
    },
    "required": ["name", "description", "palette"],
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
        "reference_notes": {"type": "string", "description": "What reference art/mood imagery Art needs for this zone."},
        "goods": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {
                    "name": {"type": "string", "description": "e.g. 'Glazed Bowl', 'Ink Tin'."},
                    "description": {"type": "string", "description": "The single product as it stands on a shelf, in one line. It must MATCH the shape family, and it must not repeat another good's use on this floor - one thing to store liquid, one to cut, one to measure, and so on."},
                    "silhouette": {"type": "string", "description": "This good's OWN outline, in one line, as a shape rather than as a name - 'a deep round cup on a long handle bent at an angle', 'four flat shelves stacked with gaps between them'. It must be the shape the real object actually has. Two rules on it: no two goods in the WHOLE event may share an outline, not just no two on this floor; and the outline must be ROUGHLY TWICE AS TALL AS IT IS WIDE (%s-%s), because the icon frame is 1:2 and a square or wide object leaves half the frame empty. Get that height from the object - a tall tool, or a wide one stood on its end - never by stretching it." % (_GOODS_ASPECT[0], _GOODS_ASPECT[1])},
                    "palette": {"type": "string", "description": "THIS GOOD's own dominant colour, not the floor's. Every good on a floor must be a clearly different colour from every other - the player sorts by colour as well as by shape, and a shelf of one palette is a shelf of one item repeated. Spread them right round the hue wheel: the shipped goods sets average 76-94 degrees of separation between any two items on a floor. What keeps a floor coherent is the shared finish - same gloss, same keyline, bright and clean - not a shared hue."}
                },
                "required": ["name", "description", "silhouette", "palette"],
                "additionalProperties": False
            },
            "description": "Sort goods: the shelf-icon catalogue for THIS zone, roughly 15 small collectible items. Sort's loop is sorting these BY EYE, so three strict rules hold. (1) Every good is a REAL tool or product of THIS floor's craft, the sort that shop would actually stock - not a generic shape with the theme painted on. (2) No two goods anywhere in the event are the same KIND of object: if one floor has a bottle, no other floor gets a bottle. (3) No two goods anywhere share an OUTLINE. (4) No two goods on a FLOOR share a COLOUR - spread them round the hue wheel and let the shared finish, not a shared palette, hold the floor together. Build the list craft-first - pick real items that already differ in shape - because picking shapes first and theming them afterwards produces invented filler, and giving every floor the same list of shapes produces a set that repeats down every column. Empty for event types with no sorting."
        },
        "before_state": {
            "type": "string",
            "description": "How THIS scene looks before anything is built, in its own terms - shuttered and dusty, bare dry ground under a flat sky, dim and cavernous, a winter-bare orchard, a dark unlit field. It must be EMPTY of everything humanised, and it must differ from the other scenes' before-states: the point of this field is variety, because a previous set gave every scene the same treatment and they all looked alike. Frozen is one option among many, not the rule - pick what suits the theme."
        },
        "interior": {
            "type": "boolean",
            "description": "True when this scene is INSIDE - a shop, a barn, a hall. Blocks scenes should mix interiors and exteriors rather than being all outdoors. It changes the 'before' frame: an exterior freezes the land, an interior freezes the room and keeps its shell."
        },
        "blunt_steps": {
            "type": "array",
            "items": {"type": "integer", "minimum": 1},
            "description": "Build-step numbers (1-based) that must be removed by their OWN dedicated edit rather than in a bulk list. A bulk removal silently leaves three kinds of object standing: the hero, a large structure, and a large flat ground element such as a laid floor. Steps 1 and 2 are handled automatically; list any OTHER step that is one of those. Empty for event types with no build sequence."
        }
    },
    # interior and blunt_steps are deliberately NOT required: both are optional
    # production hints, usually absent, and forcing a constrained decoder to
    # reason about them on every zone would cost more than it buys.
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
        "construction_materials": {
            "type": "array",
            "items": MATERIAL_SCHEMA,
            "description": "Blocks construction material icons. Each construction step consumes 2-3 of these, so they are a themed set sharing one palette. Empty for event types with no build materials."
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
    "required": ["title", "gameplay_overview", "zones", "producers", "recipes", "ingredients", "construction_materials", "event_motif", "gameplay_specific_content", "characters", "event_extras", "art_direction_guidance", "ui_icon_refresh_notes", "notes"],
    "additionalProperties": False
}




# ---------------------------------------------------------------------------
# What shape is a zone, for THIS event type?
#
# This table exists because a Blocks document shipped with Merge's mechanics in
# it. Every zone carried a table of "buildings" with 1-3 upgrade levels each and
# an A/B/C area column, none of which Blocks has. The cause was not the prompt:
# `buildings` and its `max_level` were REQUIRED fields on every zone for every
# event type, the only concrete instructions for filling them were Merge's, and
# section 4 of the renderer printed the buildings table without ever consulting
# the event type. A constrained decoder handed a required array does not leave
# it empty, so it did the only thing it could and answered in Merge.
#
# Being a table rather than a set of `if event_type ==` branches is the point.
# It drives three things that were previously free to disagree - what the schema
# will ACCEPT, what the prompt ASKS for, and what the renderer PRINTS - so a new
# event type is one row here, and cannot silently inherit Merge's shape.
#
# The conservative default matters as much as the rows. An unknown type gets no
# buildings and no levels, so the failure mode for a type nobody has calibrated
# is a thin document that says so, not a confident Merge one.
#
# Grounding, per row:
#   Merge   - the shipped Merge2606 / Merge2608 documents. 8 buildings a scene
#             across areas A/B/C, upgrade stages summing to ~15 a scene, signage
#             carrying exact wording. The one type that genuinely has all three.
#   Dessert - the shipped Dessert2605 / Dessert2608 documents: 46 buildings, but
#             signage is ALWAYS empty and nothing in the data carries a level.
#             dessert2605_scene.json has no level-ish field at all.
#   Blocks  - the shipped Blocks2604 Art Requirement (场景区域数量：5，每个区域有
#             6-8个建造步骤，每个建造步骤会用2-3种材料) and blocks2604_building.json:
#             38 rows keyed id/group/expValue/buildAnim/needMatMap, and NO level
#             field of any kind. A Blocks "building" is a one-shot construction
#             step with a build animation and a material cost - which is why the
#             38 count is real but calling them upgradeable buildings was not.
#   Sim     - sim2604_room.json / sim2604_scene.json: rooms hold a pendantList of
#             decor and chain through `nextRoom`. No per-building levels.
#   Sort    - the shipped Sort2606 (738 KB) and Sort2609 (558 KB) documents,
#             read directly. Both are small and always were: the note claiming
#             they exceed a 20MB limit was wrong. They confirm the row through
#             the same vocabulary collision this table exists to stop - Sort2609
#             writes "5 Floors (9 Construction Buildings Per Floor i.e. 9*5 = 45
#             Buildings)" where Sort2606 writes "each room has 9 construction
#             steps". Steps, not upgradeable buildings.

ZONE_SHAPES = {
    #            what one zone is   buildings levels  areas signage  steps  goods
    "Merge":   dict(unit="Scene",   buildings=True,  levels=True,  areas=True,
                    signage=True,  steps=False, goods=False, verified=True),
    "Dessert": dict(unit="Zone",    buildings=True,  levels=False, areas=False,
                    signage=False, steps=False, goods=False, verified=True),
    "Blocks":  dict(unit="Scene",   buildings=False, levels=False, areas=False,
                    signage=False, steps=True,  goods=False, verified=True),
    "Sim":     dict(unit="Room",    buildings=False, levels=False, areas=False,
                    signage=False, steps=True,  goods=False, verified=True),
    "Sort":    dict(unit="Floor",   buildings=False, levels=False, areas=False,
                    signage=False, steps=True,  goods=True,  verified=True),
}

# An unknown type inherits nothing. Prose only, and section 9 carries the
# INFERRED warning the schema already demands.
DEFAULT_ZONE_SHAPE = dict(unit="Zone", buildings=False, levels=False, areas=False,
                          signage=False, steps=True, goods=False, verified=False)

# Which top-level asset collections each type may fill. Same reasoning one level
# up: `producers` is required at the top level too, so without this a Blocks
# document could grow Merge's four producers and forty chain icons.
TYPE_COLLECTIONS = {
    "Merge":   {"producers"},
    "Dessert": {"recipes", "ingredients"},
    "Blocks":  {"construction_materials"},
    "Sort":    {"construction_materials"},
    "Sim":     set(),
}
ALL_COLLECTIONS = ("producers", "recipes", "ingredients", "construction_materials")


def zone_shape(event_type):
    return ZONE_SHAPES.get(event_type, DEFAULT_ZONE_SHAPE)


def _forbid(node, why):
    """Make an array field structurally unfillable rather than merely discouraged.

    maxItems 0 is the whole mechanism. A description saying "empty for types that
    do not have these" is advice, and advice lost to a required field is exactly
    how the Merge buildings got into the Blocks document. A constrained decoder
    cannot emit an item into an array capped at zero, so the leak stops being a
    matter of how persuasive the prompt was.
    """
    node = dict(node)
    node["maxItems"] = 0
    node["description"] = why
    node.pop("minItems", None)
    # Drop the item schema too. An array capped at zero needs none, and leaving it
    # in would keep shipping Merge's building vocabulary - "Merge2608 ships 45
    # stages across 25 buildings" - inside every Blocks prompt. Removing the
    # temptation beats out-arguing it.
    node.pop("items", None)
    return node


def schema_for(event_type):
    """ART_REQUIREMENT_SCHEMA narrowed to what THIS event type actually has.

    The module-level schema stays the full vocabulary - every field any type
    might use - and this is the only thing that should ever be handed to a model.
    Passing the base schema directly is what produced a Blocks document full of
    Merge buildings.
    """
    shape = zone_shape(event_type)
    s = copy.deepcopy(ART_REQUIREMENT_SCHEMA)
    z = s["properties"]["zones"]["items"]
    zp, zreq = z["properties"], list(z["required"])

    if shape["buildings"]:
        b = zp["buildings"]["items"]
        # Levels, areas and signage are Merge's, and Dessert borrows only the
        # buildings themselves. Dropping the properties outright beats setting
        # them to a constant: a field that does not exist cannot be filled in,
        # and cannot be rendered as a column that means nothing.
        for field, keep in (("max_level", shape["levels"]),
                            ("area", shape["areas"]),
                            ("signage", shape["signage"])):
            if not keep:
                b["properties"].pop(field, None)
                if field in b["required"]:
                    b["required"].remove(field)
    else:
        zp["buildings"] = _forbid(
            zp["buildings"],
            "This event type has no discrete buildings. Its zones are built up "
            "through construction_steps instead - leave this empty.")
        if "buildings" in zreq:
            zreq.remove("buildings")

    if not shape["steps"]:
        zp["construction_steps"] = _forbid(
            zp["construction_steps"],
            "This event type does not reveal a zone in numbered steps - it has "
            "discrete buildings instead. Leave this empty.")
        if "construction_steps" in zreq:
            zreq.remove("construction_steps")

    if not shape["goods"] and "goods" in zp:
        zp["goods"] = _forbid(zp["goods"],
                              "Only Sort zones carry a goods catalogue. Leave this empty.")

    z["required"] = zreq

    # The title example was a Merge one for every type, and an example is the
    # strongest instruction in a schema - it gets copied.
    s["properties"]["title"]["description"] = (
        "e.g. '[Artistic Requirement] %s2610 <Theme Name>'." % (event_type or "Event"))

    allowed = TYPE_COLLECTIONS.get(event_type, set())
    for name in ALL_COLLECTIONS:
        if name in s["properties"] and name not in allowed:
            s["properties"][name] = _forbid(
                s["properties"][name],
                "Not used by %s events. Leave this empty." % (event_type or "this type"))
    return s
def _variations_schema(count, event_type):
    return {
        "type": "object",
        "properties": {
            "variations": {"type": "array", "minItems": count, "maxItems": count,
                           "items": schema_for(event_type)}
        },
        "required": ["variations"],
        "additionalProperties": False
    }


def _regenerate_schema(event_type):
    """A rewrite is bound by the same per-type shape the first draft was.

    This used to be a module constant built on the full ART_REQUIREMENT_SCHEMA,
    which meant a regeneration could reintroduce exactly the fields the initial
    generation was narrowed to exclude - feedback on a Blocks document could
    hand back Merge buildings.
    """
    return {
        "type": "object",
        "properties": {
            "variation": schema_for(event_type),
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


def _generate_variations(provider, payload, system_prompt, user_message, count,
                         event_type, images=None):
    schema = _variations_schema(count, event_type)
    # One budget, both providers: Claude caps the response with it, Ollama sizes
    # its timeout from it. An Art Requirement is the largest thing this tool
    # generates, and a fixed 10-minute timeout could not fit one at any speed a
    # 32B model reaches - see ollama_timeout_for.
    max_tokens = min(MODEL_MAX_TOKENS_BASE + MODEL_MAX_TOKENS_PER_VARIATION * count, 100000)
    if provider == "ollama":
        result = _call_ollama(
            payload.get("ollama_base_url", ""),
            payload.get("ollama_model") or "qwen3:32b",
            system_prompt, user_message, schema, images=images,
            timeout=ollama_timeout_for(max_tokens)
        )
    else:
        result = _call_claude(payload.get("api_key", ""), system_prompt, user_message,
                              schema, max_tokens, images=images)
    return result["variations"]


def _regenerate_variation(provider, payload, system_prompt, user_message, event_type):
    if provider == "ollama":
        return _call_ollama(
            payload.get("ollama_base_url", ""),
            payload.get("ollama_model") or "qwen3:32b",
            system_prompt, user_message, _regenerate_schema(event_type),
            timeout=ollama_timeout_for(8500)
        )
    return _call_claude(payload.get("api_key", ""), system_prompt, user_message,
                        _regenerate_schema(event_type), 8500)


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


def _zone_content_summary(zone, shape):
    """The "Required Content / Elements" cell in section 4's summary table.

    This used to read "%d buildings, %d build steps" for every event type, which
    on a Blocks document counted construction steps and announced them as
    buildings - the same miscount that put a Levels column on scenery that is
    built once and never upgraded.
    """
    bits = []
    if shape["buildings"] and _n(zone.get("buildings")):
        bits.append("%d buildings" % _n(zone.get("buildings")))
    if shape["steps"] and _n(zone.get("construction_steps")):
        bits.append("%d construction steps" % _n(zone.get("construction_steps")))
    if shape["goods"] and _n(zone.get("goods")):
        bits.append("%d goods" % _n(zone.get("goods")))
    return ", ".join(bits) or "—"


def _buildings_md(zone, shape=None):
    """The buildings table for one zone, or nothing if the type has none.

    The columns are chosen by the event type's shape rather than fixed. Levels,
    Area and Signage are Merge's: Level drives how many upgrade stages get
    rendered, Area is Merge's A/B/C subdivision of a scene, and Signage is the
    exact wording the reviewer checks the art against - with an empty cell
    meaning "this building carries no text", which is enforced strictly.

    Printing all three for every type is how a Blocks document ended up claiming
    "Buildings (7) - 14 upgrade stages in this zone" about scenery that has no
    upgrade mechanic at all. A column nobody can fill is not neutral: it invites
    someone to fill it.
    """
    blds = zone.get("buildings") or []
    shape = shape or zone_shape(None)
    if not blds or not shape["buildings"]:
        return ""

    cols = ["#", "Building"]
    if shape["areas"]:
        cols.append("Area")
    if shape["levels"]:
        cols.append("Levels")
    cols += ["Silhouette", "Key Identification Points"]
    if shape["signage"]:
        cols.append("Signage")

    rows = []
    for b in blds:
        cells = [str(b.get("building_number", "")), b.get("name", "") or ""]
        if shape["areas"]:
            cells.append(b.get("area", "") or "—")
        if shape["levels"]:
            cells.append(str(b.get("max_level", 1)))
        cells.append((b.get("silhouette", "") or "").replace("|", "/"))
        cells.append(", ".join(b.get("key_features") or []).replace("|", "/"))
        if shape["signage"]:
            cells.append(", ".join(b.get("signage") or []) or "— none —")
        rows.append("| " + " | ".join(cells) + " |")

    if shape["levels"]:
        total = sum(int(b.get("max_level") or 1) for b in blds)
        caption = "**Buildings (%d) — %d upgrade stages in this zone**" % (len(blds), total)
    else:
        caption = "**Buildings (%d)**" % len(blds)
    return ("%s\n\n| %s |\n|%s|\n%s\n\n"
            % (caption, " | ".join(cols), "---|" * len(cols), "\n".join(rows)))


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


def _materials_md(doc):
    """The construction material icon table - Blocks' mechanic-specific class.

    Merge's producers and Dessert's recipes each had a detail table; Blocks'
    materials and Sort's goods were COUNTED in section 5 and never listed, so a
    Blocks document told the art team "6 materials" and never said which six.
    """
    mats = doc.get("construction_materials") or []
    if not mats:
        return ""
    out = ["**Construction material icons (%d)**\n" % len(mats),
           "| # | Material | Description | Palette |", "|---|---|---|---|"]
    for i, m in enumerate(mats, 1):
        out.append("| {i} | {n} | {d} | {p} |".format(
            i=i, n=m.get("name", ""),
            d=str(m.get("description", "")).replace("|", "/"),
            p=str(m.get("palette", "")).replace("|", "/")))
    return "\n".join(out) + "\n"


def _goods_md(doc):
    """The per-zone goods catalogue - Sort's mechanic-specific class."""
    zones = [z for z in (doc.get("zones") or []) if z.get("goods")]
    if not zones:
        return ""
    out = []
    for z in zones:
        goods = z.get("goods") or []
        out.append("**Zone %s - %s: %d goods**\n"
                   % (z.get("zone_number"), z.get("zone_theme"), len(goods)))
        out.append("| # | Good | Silhouette | Colour |")
        out.append("|---|---|---|---|")
        for i, g in enumerate(goods, 1):
            out.append("| {i} | {n} | {s} | {p} |".format(
                i=i, n=g.get("name", ""),
                s=str(g.get("silhouette", "")).replace("|", "/"),
                p=str(g.get("palette", "")).replace("|", "/")))
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


# ---------------------------------------------------------------------------
# The production team's Art Requirement template is ONE document for every
# event type - section 9 says so in its own words: "This allows the same Art
# Requirements template to support Sim, Management, Blocks, Sort, Merge, or
# future event types". So the difference between a Merge document and a Dessert
# one is not a different shape; it is what fills section 5.
#
# That mapping is data, kept here rather than spread through the renderer, so
# adding an event type is one entry and not a new branch in eight places. Each
# row is (group name, gameplay function, how to count it) and the count is
# derived from the document the model actually produced - never asserted - so a
# document that came back short reports what it really has.
STUDIO_DESIGN_OWNER = "Saravanan Sankar"
STUDIO_ART_OWNER = "leonardo"


def _n(seq):
    return len(seq or [])


def _asset_groups(doc, event_type):
    """Section 5 rows: (group, gameplay function, quantity) for THIS type."""
    zones = doc.get("zones") or []
    rows = []

    if event_type == "Merge":
        prods = doc.get("producers") or []
        chains = [c for p in prods for c in (p.get("chains") or [])]
        items = [i for c in chains for i in (c.get("items") or [])]
        rows.append(("Producers / generators",
                     "Tapped to yield chain items; two visual levels (closed, open)",
                     "%d producers x 2 levels = %d icons" % (_n(prods), _n(prods) * 2)))
        rows.append(("Merge chains",
                     "Merged upward; each rung a distinct, more significant item",
                     "%d chains x 5 rungs = %d item icons" % (_n(chains), _n(items) or _n(chains) * 5)))
        levels = sum(int(b.get("max_level") or 1)
                     for z in zones for b in (z.get("buildings") or []))
        rows.append(("Buildings & upgrade stages",
                     "Built and upgraded with merged goods; the scene's progression",
                     "%d buildings, %d upgrade points across %d scenes"
                     % (sum(_n(z.get("buildings")) for z in zones), levels, _n(zones))))

    elif event_type == "Dessert":
        recipes = doc.get("recipes") or []
        common = [r for r in recipes if not r.get("secret")]
        secret = [r for r in recipes if r.get("secret")]
        ing = doc.get("ingredients") or []
        rows.append(("Common recipes",
                     "Completed from ingredients to progress the event",
                     "%d dishes across %d zones" % (_n(common), _n(zones))))
        rows.append(("Secret recipes",
                     "Discovered rather than listed; the collection hook",
                     "%d dishes" % _n(secret)))
        rows.append(("Ingredients",
                     "Raw components collected in play and spent on recipes",
                     "%d common + %d secret"
                     % (_n([i for i in ing if not i.get("secret")]),
                        _n([i for i in ing if i.get("secret")]))))

    elif event_type == "Blocks":
        mats = doc.get("construction_materials") or []
        steps = sum(_n(z.get("construction_steps")) for z in zones)
        rows.append(("Construction material icons",
                     "Spent to complete a build step",
                     "%d materials" % _n(mats)))
        rows.append(("Scene build / reveal steps",
                     "Each step visibly changes the scene as the event progresses",
                     "%d steps across %d scenes" % (steps, _n(zones))))

    elif event_type == "Sort":
        goods = sum(_n(z.get("goods")) for z in zones)
        rows.append(("Goods / shelf icons",
                     "Sorted onto shelves by eye - the core loop",
                     "%d goods across %d floors" % (goods, _n(zones))))
        steps = sum(_n(z.get("construction_steps")) for z in zones)
        if steps:
            rows.append(("Room build steps",
                         "Unlocked with sorting progress",
                         "%d steps across %d rooms" % (steps, _n(zones))))

    elif event_type == "Sim":
        rows.append(("Room / furniture sets",
                     "Placed to furnish and upgrade each room",
                     "%d rooms" % _n(zones)))
        rows.append(("Guest & service assets",
                     "Guest requests and the staff actions that fulfil them",
                     "see section 6"))

    # Anything the type map does not know about, but the document carries.
    if not rows:
        rows.append(("Gameplay assets", "See section 9",
                     "%d zones" % _n(zones)))
    return rows


def _pipe_table(headers, rows, empty="_None specified._"):
    if not rows:
        return empty + "\n"
    out = ["| " + " | ".join(headers) + " |",
           "|" + "|".join(["---"] * len(headers)) + "|"]
    for r in rows:
        cells = [str(c if c is not None else "").replace("|", r"\|").replace("\n", " ")
                 for c in r]
        cells += [""] * (len(headers) - len(cells))
        out.append("| " + " | ".join(cells[:len(headers)]) + " |")
    return "\n".join(out) + "\n"


def _art_content_overview(doc, event_type):
    """Section 3 - the quantities, counted off the document rather than claimed."""
    zones = doc.get("zones") or []
    buildings = sum(_n(z.get("buildings")) for z in zones)
    steps = sum(_n(z.get("construction_steps")) for z in zones)
    ex = doc.get("event_extras") or {}
    ui_bits = [k.replace("_", " ") for k in ("hot_air_balloon", "achievement_badge") if ex.get(k)]
    ui_bits += ["%d task icons" % _n(ex.get("task_icons"))] if ex.get("task_icons") else []
    shape = zone_shape(event_type)
    unit = shape["unit"]
    rows = [
        ("Main %s / Environment" % unit,
         "%d %s%s" % (_n(zones), unit.lower(), "" if _n(zones) == 1 else "s"),
         doc.get("event_motif") or ""),
    ]
    # Only name a row the type actually has. "Areas / Rooms / Buildings: 38"
    # on a Blocks document was counting build steps and calling them buildings.
    if shape["buildings"]:
        rows.append(("Buildings", str(buildings) if buildings else "-",
                     "spread across areas A/B/C" if shape["areas"] else ""))
    if shape["steps"]:
        rows.append(("Construction / Reveal Steps",
                     "%d steps across %d %ss" % (steps, _n(zones), unit.lower())
                     if steps else "-",
                     "each one a visible change to the scene"))
    for group, _fn, qty in _asset_groups(doc, event_type):
        rows.append(("Gameplay Assets - " + group, qty, ""))
    rows.append(("Characters", "%d" % _n(doc.get("characters")), ""))
    rows.append(("UI / Icons", ", ".join(ui_bits) or "-",
                 doc.get("ui_icon_refresh_notes") or ""))
    rows.append(("Animation / VFX", "-", "to be specified with Art Lead"))
    return _pipe_table(["Art Content", "Scope / Quantity", "Notes"], rows)


def _render_markdown(doc, event_type, status_line, author_name="", md_dir=None):
    """The document in the production team's own template shape.

    This used to emit six sections of its own invention. The template the art
    team is actually given has a header block and ten numbered sections, and the
    tool was also grounded in the wrong file - Event_Final_Art_Requirement.md is
    a multi-discipline Final Spec, not this. Both are fixed now: the numbering,
    the section names and the table columns below follow the template, so a
    generated draft can be reviewed against it line for line.

    Sections the template marks optional are emitted only when the document has
    something to put in them, since it says to delete them otherwise.
    """
    go = doc.get("gameplay_overview") or {}
    md_dir = md_dir or os.getcwd()
    zones = doc.get("zones") or []
    ex = doc.get("event_extras") or {}

    # ---- header: Event Summary --------------------------------------------
    title = doc.get("title") or event_type
    # The template's naming convention: [Art Requirements] [Event ID - Event
    # Name - Event Type]. Recover the ID from the title when it carries one.
    m = re.search(r"\b([A-Z][a-z]+\d{4})\b", title)
    event_id = m.group(1) if m else ""
    # The generated title already carries a bracketed prefix and the event id, so
    # reusing it raw as the event NAME produced
    # "[Art Requirements] Blocks2610 - [Artistic Requirement] Blocks2610 Harvest
    # Fair". Strip both, and fall back to it only when no theme was stated.
    plain_title = re.sub(r"^\s*\[[^\]]*\]\s*", "", title)
    if event_id:
        plain_title = re.sub(r"\b%s\b\s*" % re.escape(event_id), "", plain_title).strip()
    theme = go.get("theme") or ""
    summary = _pipe_table(["Field", "Value"], [
        ("Event Name", theme or plain_title),
        ("Event ID / Release", event_id or "_enter Event ID / Release_"),
        ("Event Type / Gameplay", event_type),
        ("Event Theme", theme),
        ("Event Duration", go.get("duration_summary") or ""),
        ("Design Spec", "_add link_"),
        ("Design Owner", doc.get("design_owner") or STUDIO_DESIGN_OWNER),
        ("Art Owner", doc.get("art_owner") or STUDIO_ART_OWNER),
        ("Status", status_line),
    ])

    # ---- 1. Event / Gameplay Overview -------------------------------------
    structure = _pipe_table(["Content Scope", "Detail"], [
        ("Main Scene(s)", "%d" % len(zones)),
        ("Areas / Phases", ", ".join(str(z.get("zone_theme") or "") for z in zones)),
        ("Build / Upgrade Steps",
         "%d" % sum(_n(z.get("construction_steps")) for z in zones) or "-"),
        ("Gameplay Asset Groups",
         ", ".join(g for g, _f, _q in _asset_groups(doc, event_type))),
        ("Characters", "%d" % _n(doc.get("characters"))),
        ("Other Event Content",
         ", ".join(k.replace("_", " ") for k in ("hot_air_balloon", "achievement_badge")
                   if ex.get(k)) or "-"),
    ])
    screens = _pipe_table(
        ["Screen / Area", "Purpose / Main Content", "Reference"],
        [("Zone %s - %s" % (z.get("zone_number"), z.get("zone_theme")),
          (z.get("reference_notes") or "")[:160],
          "see Scene Detail below" if z.get("image_path") else "_add reference_")
         for z in zones])

    # ---- 2. Theme & Overall Art Direction ---------------------------------
    adg = doc.get("art_direction_guidance") or {}
    lean = "\n".join("- %s" % x for x in (adg.get("lean_into") or [])) or "_none stated._"
    avoid = "\n".join("- %s" % x for x in (adg.get("avoid") or [])) or "_none stated._"

    # ---- 4. Scene / Environment -------------------------------------------
    # Everything below asks the shape what this type has rather than assuming
    # Merge's. Section 4 never used to receive event_type at all.
    shape = zone_shape(event_type)
    area_rows = []
    for z in zones:
        area_rows.append((
            "Zone %s - %s" % (z.get("zone_number"), z.get("zone_theme")),
            (z.get("reference_notes") or "")[:200],
            " / ".join(x for x in (z.get("palette_or_mood"), z.get("scene_material"),
                                   z.get("main_tone")) if x),
            _zone_content_summary(z, shape),
            "see Scene Detail below" if z.get("image_path") else "_add reference_",
        ))
    areas = _pipe_table(["Scene / Area", "Requirement / Function",
                         "Detailed Art Direction", "Required Content / Elements",
                         "Visual Reference + Reference Intent"], area_rows)

    # The template's tables above are a summary. The per-zone detail this tool
    # already produced - the buildings table with its Levels and Signage columns,
    # the reveal steps, the generated reference image - is exactly what section 4
    # means by "Detailed Art Direction", so it is kept here rather than flattened
    # into a cell. The image especially: markdown inside a table cell does not
    # render as an image, so it needs a line of its own.
    detail_lines = []
    for z in zones:
        detail_lines.append("#### Zone %s - %s\n" % (z.get("zone_number"),
                                                     z.get("zone_theme")))
        detail_lines.append("**Palette / Mood:** %s\n" % (z.get("palette_or_mood") or ""))
        if z.get("scene_material"):
            detail_lines.append("**Scene Material:** %s\n" % z["scene_material"])
        if z.get("main_tone"):
            detail_lines.append("**Main Tone:** %s\n" % z["main_tone"])
        if z.get("before_state"):
            detail_lines.append("**Before State:** %s\n" % z["before_state"])
        steps = z.get("construction_steps") or []
        if steps:
            detail_lines.append("**Construction / Reveal Steps:**\n%s\n" % "\n".join(
                "  %d. %s" % (i + 1, st) for i, st in enumerate(steps)))
        bt = _buildings_md(z, shape)
        if bt:
            detail_lines.append(bt)
        if z.get("reference_notes"):
            detail_lines.append("**Reference Notes:** %s\n" % z["reference_notes"])
        img = _zone_image_md(doc, z, md_dir)
        if img:
            detail_lines.append(img.strip() + "\n")
    detail_md = "\n".join(detail_lines)

    element_rows = []
    for z in zones:
        for b in (z.get("buildings") or []):
            element_rows.append((
                "%s (zone %s%s)" % (b.get("name"), z.get("zone_number"),
                                    ", area " + b["area"] if b.get("area") else ""),
                b.get("silhouette") or "",
                z.get("palette_or_mood") or "",
                "; ".join(b.get("key_features") or []),
                ("signage: " + ", ".join(b.get("signage"))) if b.get("signage")
                else "no signage",
            ))
    elements = _pipe_table(["Element", "Function / Content", "Detailed Art Direction",
                            "Key Props / Visual Elements",
                            "Visual Reference + Reference Intent"], element_rows,
                           empty="_This event type has no per-building breakdown - "
                                 "its zones are built through the steps below._")

    prog_rows = []
    for z in zones:
        steps = z.get("construction_steps") or []
        if not steps:
            continue
        prog_rows.append((
            "Zone %s - %s" % (z.get("zone_number"), z.get("zone_theme")),
            "Built up step by step" + (" from a before-state" if z.get("before_state") else ""),
            "%d steps" % len(steps),
            "; ".join("%d. %s" % (i + 1, s) for i, s in enumerate(steps)),
            "see Scene Detail below" if z.get("image_path") else "_add reference_",
        ))
    progression = _pipe_table(["Element / Area", "Progression Requirement",
                              "# Steps / States", "Detailed Art Direction",
                              "Reference"], prog_rows,
                             empty="_Nothing in this event changes visually as it progresses._")

    # The template's own section 4 headings, named for what THIS type has. A
    # Blocks document headed "Room / Building / Workshop Breakdown" over an empty
    # table, followed by "Build / Upgrade / Visual Progression" over content that
    # is never upgraded, was the most visible half of the Merge leak.
    element_heading = ("Room / Building / Workshop Breakdown" if shape["buildings"]
                       else "%s Element Breakdown" % shape["unit"])
    progression_heading = ("Build / Upgrade / Visual Progression" if shape["levels"]
                           else "Construction / Reveal Progression")

    # ---- 5. Gameplay / Event Asset Requirements ---------------------------
    # The one section that differs by event type, which is exactly what the
    # template intends: Merge's producers and chains, Dessert's recipes and
    # ingredients, Blocks' materials, Sort's goods.
    asset_rows = [(g, fn, "see the detail below", qty, "_add reference_")
                  for g, fn, qty in _asset_groups(doc, event_type)]
    assets = _pipe_table(["Asset / Asset Group", "Gameplay Function",
                          "Detailed Art Requirement", "Quantity / Variants / States",
                          "Visual Reference + Reference Intent"], asset_rows)

    # ---- 6. Characters -----------------------------------------------------
    chars = doc.get("characters") or []
    char_rows = [(c.get("name"), c.get("role"), c.get("design_notes"),
                  ", ".join(c.get("scene_actions") or []), "_add reference_")
                 for c in chars]
    char_table = _pipe_table(["Character / Role", "Function / Personality",
                              "Detailed Visual Direction",
                              "Costume / Props / Key Features",
                              "Visual Reference + Reference Intent"], char_rows)

    # ---- 7. UI / Icon ------------------------------------------------------
    icon_rows = []
    for group, _fn, qty in _asset_groups(doc, event_type):
        icon_rows.append((group, "Gameplay", doc.get("event_motif") or "", qty, "_add reference_"))
    if ex.get("achievement_badge"):
        icon_rows.append(("Achievement badge", "Progress / Collection",
                          ex["achievement_badge"], "1", "_add reference_"))
    if ex.get("hot_air_balloon"):
        icon_rows.append(("Event entry art", "Entry / promo",
                          ex["hot_air_balloon"], "1", "_add reference_"))
    if ex.get("task_icons"):
        icon_rows.append(("Task icons", "Progress / Collection",
                          "; ".join(ex["task_icons"]), str(_n(ex["task_icons"])),
                          "_add reference_"))
    icons = _pipe_table(["Icon / Group", "Function", "Visual / Theme Direction",
                         "Quantity / Variants", "Visual Reference"], icon_rows)

    by_line = "\n**By:** %s" % author_name if author_name else ""
    optional_chars = ""
    if chars:
        optional_chars = """
## 6. Character Requirements

**Total characters:** %d

%s
""" % (len(chars), char_table)

    return """# [Art Requirements] %s - %s

**Status:** %s%s

%s

## 1. Event / Gameplay Overview

**Theme:** %s

**Event Overview**

%s

**Core Gameplay Loop**

%s

### Event Structure

%s
### Main Screens / Flow

%s
## 2. Theme & Overall Art Direction

### Overall Art Direction

%s

**Key Visual Requirements**

%s

**Avoid**

%s

**Reference Intent:** the references in this document communicate atmosphere, palette,
shape language and content. They are not final production assets.

## 3. Art Content Overview

%s
## 4. Scene / Environment Requirements

### Scene / Area Breakdown

%s
### %s

%s
### %s

%s
### Scene Detail

%s
## 5. Gameplay / Event Asset Requirements

%s
### Detail

%s%s
%s
%s
## 7. UI / Icon Requirements

%s
## 8. Animation / VFX Requirements

%s

## 9. Additional Event-Specific Art Requirements

%s

## 10. SKB Feedback

%s
---
AI-drafted via the Event Automator Art Requirement tab, in the production team's
Art Requirement template. Treat as a first draft for Art Lead and SKB review,
not an approved spec.
""" % (
        event_id or event_type, theme or plain_title,
        status_line, by_line,
        summary,
        theme or plain_title or "_not stated_",
        go.get("narrative_packaging") or "_not stated_",
        go.get("loop_summary") or "_not stated_",
        structure,
        screens,
        adg.get("overall") if isinstance(adg.get("overall"), str) and adg.get("overall")
        else (go.get("narrative_packaging") or "_see the per-zone direction in section 4._"),
        lean,
        avoid,
        _art_content_overview(doc, event_type),
        areas,
        element_heading,
        elements,
        progression_heading,
        progression,
        detail_md,
        assets,
        _producers_md(doc), _recipes_md(doc) + _materials_md(doc) + _goods_md(doc),
        doc.get("gameplay_specific_content") or "",
        optional_chars,
        icons,
        _pipe_table(["Asset / Character / Moment", "Trigger / State",
                     "Animation / VFX Requirement", "Visual Reference"],
                    [], empty="_Not specified in this draft - to be agreed with the Art Lead._"),
        doc.get("notes") or "_None._",
        _pipe_table(["#", "Area / Requirement", "SKB Feedback",
                     "Design + Art Response / Action", "Owner", "Status"],
                    [], empty="_To be filled during SKB review._"),
    )


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


def preview_main():
    """The exact Markdown the art team receives, for on-screen review.

    Deliberately calls the same _render_markdown() the draft/final writers
    use rather than reformatting the doc dict a second time - a preview that
    is separately rendered is a preview that can silently disagree with the
    file on disk, which is precisely the drift the renderer used to have.
    md_dir matters: zone image links are written relative to it, so preview
    against the directory the doc actually lives in.
    """
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    if len(sys.argv) < 2:
        print(json.dumps({"error": "Usage: artreq_preview <payload.json>"}))
        sys.exit(1)

    with open(sys.argv[1], "r", encoding="utf-8-sig") as f:
        payload = json.load(f)

    variation = payload.get("variation")
    if not variation:
        print(json.dumps({"error": "No variation supplied to preview."}))
        sys.exit(1)

    event_type = payload.get("event_type", "")
    events_root = payload.get("events_root", "")
    status = payload.get("status_line") or "Draft — under review"
    author = payload.get("author_name", "")

    # Match the directory the corresponding writer would use, so relative
    # image paths in the preview resolve the same way they will in the file.
    if payload.get("is_final") and events_root:
        md_dir = os.path.join(events_root, event_type, "Art Requirements")
    elif events_root:
        md_dir = os.path.join(events_root, event_type, "Art Requirements", "Drafts")
    else:
        md_dir = os.getcwd()

    try:
        print(json.dumps({
            "ok": True,
            "markdown": _render_markdown(variation, event_type, status, author, md_dir=md_dir),
        }, ensure_ascii=False))
    except Exception as e:
        print(json.dumps({"error": str(e)}))
        sys.exit(1)


def generate_main():
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    if len(sys.argv) < 2:
        print(json.dumps({"error": "Usage: artreq_generate <payload.json>"}))
        sys.exit(1)

    with open(sys.argv[1], "r", encoding="utf-8-sig") as f:
        payload = json.load(f)

    event_type = payload.get("event_type", "")
    ideation_theme = payload.get("ideation_theme") or {}
    # An Art Requirement can start from a theme this tool just generated, OR from
    # a One-Pager that already exists on disk - one finalized in an earlier
    # session, written by hand, or saved by a colleague into the shared folder.
    # Only the markdown is kept when a theme is finalized, so the document text
    # is the honest thing to pass back rather than a theme dict reconstructed
    # from it - and it is also what carries any edits made after saving.
    ideation_doc_text = (payload.get("ideation_doc_text") or "").strip()
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
    if not ideation_theme.get("theme_name") and not ideation_doc_text:
        print(json.dumps({"error": "No Ideation theme or One-Pager supplied - either "
                                   "select a generated theme as Final, or pick an "
                                   "existing One-Pager to build from."}))
        sys.exit(1)

    try:
        lessons_block = store.lessons_prompt_block(events_root, event_type)
        final_spec_template_text = _read_template(events_root, ART_REQ_TEMPLATE)
        art_req_folder = os.path.join(events_root, event_type, "Art Requirements") if events_root else ""
        _, past_doc_text = find_latest_doc(art_req_folder)

        if ideation_doc_text:
            source_block = (
                "Approved One-Pager Ideation document to translate into art "
                "direction. This is the finished, signed-off brief - treat every "
                "detail in it as decided, and do not invent a different theme:\n\n"
                "----- BEGIN ONE-PAGER -----\n"
                f"{ideation_doc_text}\n"
                "----- END ONE-PAGER -----"
            )
        else:
            source_block = (
                "Finalized Ideation theme to translate into art direction:\n"
                f"{json.dumps(ideation_theme, ensure_ascii=False)}"
            )

        user_message = (
            f"{source_block}\n\n"
            f"Generate exactly {variation_count} distinct Art Requirement variations "
            f"for this theme - different concrete visual executions of the SAME "
            f"approved theme (not different themes). Vary zone composition, "
            f"character designs, and specific reveal/construction steps across "
            f"variations; keep them all faithful to the theme above."
        )

        # The reference art from a converted Word/PDF brief, attached to the call
        # rather than only described. OCR recovers a mood board's WORDS; it does
        # not recover what the board looks like, and that is the half of a
        # One-Pager the art direction is actually derived from.
        ref_images = []
        doc_path = (payload.get("ideation_doc_path") or "").strip()
        if doc_path:
            try:
                import doc_convert
                ref_images = doc_convert.reference_images(doc_path)
            except Exception:
                ref_images = []         # never lose the generation over an image

        # Ollama hard-rejects the whole request (HTTP 400) when a model without
        # vision support gets `images` attached - it does not ignore them.
        # Check before referencing them in the prompt text below, so the two
        # stay consistent: never tell the model to "read the N images attached
        # above" when we're about to send zero.
        if ref_images and provider == "ollama":
            model = payload.get("ollama_model") or "qwen3:32b"
            if not _ollama_supports_vision(payload.get("ollama_base_url", ""), model):
                log(f"'{model}' has no vision support - generating from the brief's text only "
                    f"({len(ref_images)} reference image(s) from the source One-Pager dropped).")
                ref_images = []

        if ref_images:
            user_message += (
                "\n\nThe %d image(s) attached above are the reference art from that "
                "One-Pager - the look the theme was approved on. Read the art "
                "direction from them: palette, materials, lighting, the level of "
                "detail and how built-up a scene is. They are REFERENCE, not a "
                "layout to copy, and the brief's words win wherever the two "
                "disagree." % len(ref_images))

        system_prompt = _build_system_prompt(event_type, lessons_block, final_spec_template_text, past_doc_text)
        variations = _generate_variations(provider, payload, system_prompt, user_message,
                                          variation_count, event_type, images=ref_images)

        draft_paths = []
        if events_root:
            for i, v in enumerate(variations):
                p = _write_variation_draft(events_root, event_type, v, i, author_name)
                if p:
                    draft_paths.append(p)

        print(json.dumps({"variations": variations, "provider": provider,
                          "draft_paths": draft_paths,
                          "reference_images": len(ref_images)}, ensure_ascii=False))
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
        final_spec_template_text = _read_template(events_root, ART_REQ_TEMPLATE)
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
        result = _regenerate_variation(provider, payload, system_prompt, user_message,
                                       event_type)
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


def schemas_main():
    """CLI: pybackend_cli.py artreq_schemas

    Prints {event_type: schema} for every known type, so a caller outside Python
    - the MCP server - can declare one tool per event type whose INPUT schema is
    that type's own. That is the whole point: an MCP host constrains a tool call
    to its input schema the same way the Anthropic API constrains a response, so
    a Blocks tool that has no max_level in its schema cannot be called with one.
    The guarantee stops depending on how persuasive the prompt was.
    """
    import json
    import sys
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    try:
        out = {t: schema_for(t) for t in KNOWN_TYPES}
        out["_shapes"] = {t: zone_shape(t) for t in KNOWN_TYPES}
        print(json.dumps(out, ensure_ascii=False))
    except Exception as e:
        print(json.dumps({"error": "%s: %s" % (e.__class__.__name__, e)}))
        sys.exit(1)


def mcp_save_main():
    """CLI: pybackend_cli.py artreq_mcp_save payload.json

    payload: {event_type, variation, ref_name, events_root, author_name}

    The MCP host already constrains a tool call to the tool's input schema, so a
    Blocks save cannot arrive carrying max_level. This validates anyway, for two
    reasons: a host that does not enforce input schemas would otherwise write a
    malformed document silently, and the check costs nothing next to the render.
    Belt and braces on the one property this whole design exists to guarantee.
    """
    import json
    import sys
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    if len(sys.argv) < 2:
        print(json.dumps({"error": "Usage: artreq_mcp_save <payload.json>"}))
        sys.exit(1)
    try:
        with open(sys.argv[1], "r", encoding="utf-8-sig") as f:
            payload = json.load(f)

        event_type = payload.get("event_type", "")
        if event_type not in KNOWN_TYPES:
            raise ValueError("Unknown event type: %s" % event_type)
        doc = payload.get("variation") or {}
        if not doc.get("zones"):
            raise ValueError("The variation has no zones.")

        shape = zone_shape(event_type)
        problems = []
        for z in doc["zones"]:
            n = z.get("zone_number")
            if not shape["buildings"] and (z.get("buildings") or []):
                problems.append(
                    "zone %s carries buildings, and %s has none - its zones are "
                    "built through construction_steps" % (n, event_type))
            if not shape["steps"] and (z.get("construction_steps") or []):
                problems.append(
                    "zone %s carries construction_steps, and %s builds through "
                    "discrete buildings instead" % (n, event_type))
            if not shape["goods"] and (z.get("goods") or []):
                problems.append("zone %s carries goods, which only Sort has" % n)
            for b in (z.get("buildings") or []):
                if not shape["levels"] and b.get("max_level") not in (None, 1):
                    problems.append(
                        "zone %s building %s has max_level %s, and %s has no "
                        "upgrade mechanic" % (n, b.get("name"), b.get("max_level"),
                                              event_type))
                if not shape["areas"] and (b.get("area") or "").strip():
                    problems.append(
                        "zone %s building %s has an area, and only Merge "
                        "subdivides a scene into A/B/C" % (n, b.get("name")))
        allowed = TYPE_COLLECTIONS.get(event_type, set())
        for coll in ALL_COLLECTIONS:
            if coll not in allowed and (doc.get(coll) or []):
                problems.append("%s is not used by %s events" % (coll, event_type))
        if problems:
            raise ValueError("This document carries mechanics %s does not have. "
                             "Fix and call again:\n  - %s"
                             % (event_type, "\n  - ".join(problems)))

        events_root = payload.get("events_root", "")
        if not events_root:
            raise ValueError("No events_root configured.")
        ref = re.sub(r'[\\/:*?"<>|]', "_", payload.get("ref_name") or "Untitled")[:80]
        author = payload.get("author_name", "")
        draft_dir = os.path.join(events_root, event_type, "Art Requirements", "Drafts")
        os.makedirs(draft_dir, exist_ok=True)
        path = os.path.join(draft_dir, "%s.md" % ref)

        status = "Draft - written in Claude Desktop, pending Art Lead review"
        if author:
            status += " - by %s" % author
        md = _render_markdown(doc, event_type, status, author, md_dir=draft_dir)
        with open(path, "w", encoding="utf-8", newline="") as f:
            f.write(md)

        zones = doc["zones"]
        summary = ("%s: %d zones, %d construction steps, %d buildings, %d characters, "
                   "%d chars rendered."
                   % (event_type, len(zones),
                      sum(len(z.get("construction_steps") or []) for z in zones),
                      sum(len(z.get("buildings") or []) for z in zones),
                      len(doc.get("characters") or []), len(md)))
        print(json.dumps({"path": path, "summary": summary}, ensure_ascii=False))
    except ValueError as e:
        print(json.dumps({"error": str(e)}, ensure_ascii=False))
        sys.exit(1)
    except Exception as e:
        print(json.dumps({"error": "%s: %s" % (e.__class__.__name__, e)}))
        sys.exit(1)
