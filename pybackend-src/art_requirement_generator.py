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
        "A Production Equipment & Product Chains table: for each producer "
        "building, two upgrade chains (Chain A / Chain B) of themed items, "
        "shown as an icon progression per level."
    ),
    "Dessert": (
        "A Recipe / Ingredient Requirement: a common-recipes table and a "
        "secret-recipes table, each mapping a dish to its required "
        "ingredient icons."
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

ZONE_SCHEMA = {
    "type": "object",
    "properties": {
        "zone_number": {"type": "integer"},
        "zone_theme": {"type": "string", "description": "e.g. 'Tropical Rainforest', 'Rock Concert Stage'."},
        "palette_or_mood": {"type": "string"},
        "construction_steps": {
            "type": "array",
            "items": {"type": "string"},
            "description": "Numbered visual build-up steps or a building/level list for this zone - whichever this event type actually uses."
        },
        "reference_notes": {"type": "string", "description": "What reference art/mood imagery Art needs for this zone."}
    },
    "required": ["zone_number", "zone_theme", "palette_or_mood", "construction_steps", "reference_notes"],
    "additionalProperties": False
}

CHARACTER_SCHEMA = {
    "type": "object",
    "properties": {
        "name": {"type": "string"},
        "role": {"type": "string", "enum": ["Gift-giving/mechanic NPC", "Judge/special guest", "Ordinary guest", "Chef/host", "Other"]},
        "design_notes": {"type": "string"},
        "scene_actions": {"type": "array", "items": {"type": "string"}}
    },
    "required": ["name", "role", "design_notes", "scene_actions"],
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
        "ui_icon_refresh_notes": {
            "type": "string",
            "description": "Which existing UI icons/panels (achievement badge, entry icon, restock/currency icon, event pack art) need re-skinning to this event's palette - reuse existing icon shapes and positions, only the motif/color changes."
        },
        "notes": {"type": "string"}
    },
    "required": ["title", "gameplay_overview", "zones", "gameplay_specific_content", "characters", "art_direction_guidance", "ui_icon_refresh_notes", "notes"],
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


def _render_markdown(doc, event_type, status_line, author_name=""):
    go = doc.get("gameplay_overview") or {}
    zones_md = ""
    for z in (doc.get("zones") or []):
        steps = "\n".join(f"  {i + 1}. {s}" for i, s in enumerate(z.get("construction_steps") or []))
        zones_md += (
            f"\n### Zone {z.get('zone_number')} — {z.get('zone_theme')}\n\n"
            f"**Palette/Mood:** {z.get('palette_or_mood')}\n\n"
            f"**Construction/Reveal Steps:**\n{steps}\n\n"
            f"**Reference Notes:** {z.get('reference_notes')}\n"
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
        f.write(_render_markdown(doc, event_type, f"Draft variation {index + 1} — under review, not yet selected", author_name))
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
            f.write(_render_markdown(variation, event_type, status, selector))
        print(json.dumps({"path": path}))
    except Exception as e:
        print(json.dumps({"error": str(e)}))
        sys.exit(1)
