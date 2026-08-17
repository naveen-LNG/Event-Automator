"""
Event Ideation tab backend: generates 5 themed 1-Pager Ideation drafts for a
given event type + duration, and regenerates a single theme to address
reviewer feedback.

Two providers, chosen by Electron (main.js) and passed in as payload["provider"]:
- "anthropic" (default whenever a key is configured): the real Anthropic
  Messages API via the `anthropic` package, bundled into the embeddable
  interpreter alongside openpyxl - see scripts/build-pybackend.ps1. Requires
  a metered Anthropic Console API key (Settings tab), separate from any
  claude.ai subscription.
- "ollama": local fallback used only when no Anthropic key is configured -
  a self-hosted Ollama server (e.g. qwen3:32b on a shared office-LAN
  machine), called over plain HTTP, no key required. See
  _call_ollama()/_call_claude() - both providers share the same JSON
  Schemas (GENERATE_SCHEMA/REGENERATE_SCHEMA) since Ollama's `format` field
  accepts a JSON Schema the same way Claude's output_config.format does.

Feedback loop: every piece of feedback a designer submits on a theme is
(1) applied immediately via a regeneration call, and (2) distilled into a
short, reusable "lesson". Both are persisted by ideation_store as one
immutable file per entry, so several designers sharing the ideationData
folder (OneDrive/network share, path configurable in Settings) never write
the same file and never lose each other's feedback to a sync conflict.
Every future generate/regenerate call folds the accumulated, non-retracted
lessons into its system prompt, so the same note doesn't need giving twice.
A human-readable "<Type>/Feedback/<ref> Ideation Feedback.md" digest is
regenerated from those entries so the trail is readable outside this app.
"""
import sys
import os
import re
import json

import openpyxl
import anthropic

import ideation_store as store

MODEL = "claude-opus-5"

KNOWN_TYPES = ["Merge", "Sim", "Blocks", "Sort", "Dessert"]
ALIAS_RE = re.compile(r'^([A-Za-z]+)(\d{4})')

THEME_SCHEMA = {
    "type": "object",
    "properties": {
        "theme_name": {"type": "string", "description": "Short, evocative theme name a designer would put on a 1-Pager, e.g. 'Sunday Garden Club'."},
        "internal_reference_name": {"type": "string", "description": "Internal code in the studio's own convention: <Type><YY><MM>, e.g. 'Merge2609'."},
        "reasoning": {"type": "string", "description": "1-3 sentences: why this theme, why now, why it fits the 35+ US women audience."},
        "audience_fit": {
            "type": "object",
            "properties": {
                "progression_completionist": {"type": "string", "enum": ["Green", "Yellow", "Red"]},
                "casual_escape": {"type": "string", "enum": ["Green", "Yellow", "Red"]},
                "social_collector": {"type": "string", "enum": ["Green", "Yellow", "Red"]},
                "notes": {"type": "string", "description": "Why any Yellow/Red rating is acceptable, if present."}
            },
            "required": ["progression_completionist", "casual_escape", "social_collector", "notes"],
            "additionalProperties": False
        },
        "visual_direction": {
            "type": "object",
            "properties": {
                "art_style": {"type": "string"},
                "color_palette": {"type": "string"},
                "key_features": {"type": "array", "items": {"type": "string"}, "description": "Setting/scene elements, e.g. named zones or landmarks."},
                "ui_direction": {"type": "array", "items": {"type": "string"}}
            },
            "required": ["art_style", "color_palette", "key_features", "ui_direction"],
            "additionalProperties": False
        },
        "scenes": {
            "type": "array",
            "minItems": 3,
            "maxItems": 3,
            "items": {
                "type": "object",
                "properties": {
                    "name": {"type": "string"},
                    "description": {"type": "string"}
                },
                "required": ["name", "description"],
                "additionalProperties": False
            },
            "description": "The board's 3-scene visual progression, matching the studio's real 1-Pager format."
        },
        "new_features": {"type": "string", "description": "New mechanics/additions, or the literal string 'No additions or modifications.' if this is a config-only theme refresh."},
        "expected_impact": {
            "type": "object",
            "properties": {
                "engagement_note": {"type": "string", "description": "Qualitative directional note only - never a fabricated precise ARPU/DAU number. Label as a draft estimate for PO/econ review."},
                "monetization_note": {"type": "string"}
            },
            "required": ["engagement_note", "monetization_note"],
            "additionalProperties": False
        },
        "effort_estimate": {
            "type": "object",
            "properties": {
                "art": {"type": "string", "enum": ["Low", "Medium", "High"]},
                "engineering": {"type": "string", "enum": ["Low", "Medium", "High"]},
                "design": {"type": "string", "enum": ["Low", "Medium", "High"]},
                "overall": {"type": "string", "enum": ["Low", "Medium", "High"]}
            },
            "required": ["art", "engineering", "design", "overall"],
            "additionalProperties": False
        }
    },
    "required": [
        "theme_name", "internal_reference_name", "reasoning", "audience_fit",
        "visual_direction", "scenes", "new_features", "expected_impact", "effort_estimate"
    ],
    "additionalProperties": False
}

def _generate_schema(count):
    return {
        "type": "object",
        "properties": {
            "themes": {
                "type": "array",
                "minItems": count,
                "maxItems": count,
                "items": THEME_SCHEMA
            }
        },
        "required": ["themes"],
        "additionalProperties": False
    }

REGENERATE_SCHEMA = {
    "type": "object",
    "properties": {
        "theme": THEME_SCHEMA,
        "extracted_lesson": {
            "type": "string",
            "description": "One generalizable sentence capturing the rule behind this feedback, phrased so it prevents the same mistake on a future, unrelated theme. Empty string if the feedback carries no generalizable rule (e.g. a one-off factual correction)."
        }
    },
    "required": ["theme", "extracted_lesson"],
    "additionalProperties": False
}

AUDIENCE_GUIDANCE = """\
Target player: 35+ US women. Lean into themes this audience already gravitates \
toward - wellness and self-care, home and garden, cozy seasonal moments, travel \
and culture, baking and crafts, book clubs, family milestones and reunions, \
nostalgia and heritage, spa and retreat settings, community and social events. \
Avoid themes that skew toward young/teen pop culture, gamer-culture references, \
or generic fantasy/sci-fi with no connection to this audience's actual interests."""



def log(msg):
    print(f"[Ideation] {msg}", file=sys.stderr, flush=True)


def write_feedback_digest(events_root, event_type, ref_name, doc_kind, kind_label):
    """Human-readable feedback trail, alongside the studio's real
    <Type>/Feedback/<code> ... .pdf files - so a designer browsing the event
    folder sees the discussion without opening this app.

    Regenerated wholesale from the entry files every time rather than
    appended to. Append-to-one-file is exactly the pattern that makes a
    synced folder produce "conflicted copy" duplicates; a fully derived file
    converges to the same content on every machine once sync settles, so a
    conflicted copy is harmless rather than lossy."""
    if not events_root:
        return None
    entries = store.read_feedback_entries(events_root, event_type, doc_kind=doc_kind, ref_name=ref_name)
    if not entries:
        return None
    retracted = store.read_retractions(events_root, event_type)

    safe_ref = re.sub(r'[\\/:*?"<>|]', '_', ref_name or "UnnamedEvent")
    log_dir = os.path.join(events_root, event_type, "Feedback")
    os.makedirs(log_dir, exist_ok=True)
    log_path = os.path.join(log_dir, f"{safe_ref} {kind_label} Feedback.md")

    lines = [
        f"# {kind_label} Feedback Log — {safe_ref}",
        "",
        "_Generated from the feedback entries in `Feedback/entries/`. "
        "Edits here are overwritten - add feedback through the tool._",
        "",
    ]
    for e in entries:
        is_retracted = e.get("entry_id") in retracted
        lines.append(f"## {e.get('author_name')} — {e.get('created_at')}")
        lines.append("")
        lines.append(f"**On:** {e.get('subject_title')}")
        lines.append("")
        lines.append(f"**Feedback:** {e.get('feedback')}")
        lines.append("")
        lesson = (e.get("extracted_lesson") or "").strip()
        if lesson:
            if is_retracted:
                r = retracted[e["entry_id"]]
                lines.append(
                    f"**Lesson (RETRACTED by {r.get('author_name')}"
                    f"{' — ' + r['reason'] if r.get('reason') else ''}):** ~~{lesson}~~"
                )
            else:
                lines.append(f"**Lesson (applied to future {event_type} {kind_label.lower()} runs):** {lesson}")
            lines.append("")
        if e.get("revised_title"):
            lines.append(f"**Revised to:** {e['revised_title']}")
            lines.append("")
        lines.append("---")
        lines.append("")

    with open(log_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    return log_path


def find_latest_doc(folder, exclude_dirname="Drafts"):
    """Most-recently-written finalized document in a folder (skips the
    Drafts/ subfolder) - used to ground a new generation in whatever this
    tool has actually produced for that event type so far. Returns
    (path, text) or (None, None). Best-effort/non-fatal by design: on a
    type's first-ever run there is nothing here yet, and that's expected,
    not an error."""
    if not folder or not os.path.isdir(folder):
        return None, None
    candidates = []
    for name in os.listdir(folder):
        full = os.path.join(folder, name)
        if name == exclude_dirname or not os.path.isfile(full) or not name.lower().endswith(".md"):
            continue
        candidates.append(full)
    if not candidates:
        return None, None
    latest = max(candidates, key=os.path.getmtime)
    try:
        with open(latest, "r", encoding="utf-8") as f:
            return latest, f.read()
    except Exception as e:
        log(f"Warning: could not read {latest}: {e}")
        return None, None


def _recent_events_for_type(design_path, event_type, limit=5):
    """Best-effort: recent event names/dates of this type, for grounding and
    to steer the model away from repeating a very recent theme. Never fatal -
    ideation should still work even if events.xlsx is missing or unreadable."""
    events_path = os.path.join(design_path or "", "events.xlsx")
    if not design_path or not os.path.exists(events_path):
        return []
    try:
        wb = openpyxl.load_workbook(events_path, data_only=True, read_only=True)
        ws = wb["events"]
    except Exception as e:
        log(f"Warning: could not read events.xlsx for history context: {e}")
        return []

    matches = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        alias = row[1] if len(row) > 1 else None
        if not alias:
            continue
        alias = str(alias).strip()
        m = ALIAS_RE.match(alias)
        if not m or m.group(1) != event_type:
            continue
        start_time = row[7] if len(row) > 7 else None
        matches.append({"alias": alias, "start_time": str(start_time) if start_time else ""})
    return matches[-limit:]


def _read_template(events_root, filename):
    """Reads the studio's real ideation template verbatim, so the model is
    grounded in the actual document, not a paraphrase baked into this
    script. Best-effort: ideation should still work if TEMPLATES/ is ever
    moved or renamed, just with weaker grounding."""
    if not events_root:
        return ""
    path = os.path.join(events_root, "TEMPLATES", filename)
    try:
        with open(path, "r", encoding="utf-8") as f:
            return f.read()
    except Exception as e:
        log(f"Warning: could not read template {filename}: {e}")
        return ""


def _render_theme_markdown(theme, event_type, duration_days, status_line, author_name=""):
    """Mirrors the real Event_Ideation_1Pager.md section structure (Event
    Overview / Theme & Reasoning / New Features / Product Alignment /
    Event Item Visual References / Effort Estimation / Submission
    Checklist) so the file this writes is an actual reviewable document in
    the studio's own format, not an ad hoc layout."""
    af = theme.get("audience_fit") or {}
    vd = theme.get("visual_direction") or {}
    imp = theme.get("expected_impact") or {}
    eff = theme.get("effort_estimate") or {}
    scenes = "\n".join(f"{i + 1}. **{s.get('name')}** — {s.get('description')}" for i, s in enumerate(theme.get("scenes") or []))
    key_features = "\n".join(f"- {f}" for f in (vd.get("key_features") or []))
    ui_direction = "\n".join(f"- {f}" for f in (vd.get("ui_direction") or []))

    by_line = f"\n**By:** {author_name}" if author_name else ""

    return f"""# Event Ideation 1-Pager — {theme.get('theme_name')}

**Status:** {status_line}{by_line}

## 1. Event Overview

| Field | Value |
|---|---|
| Event Theme | {theme.get('theme_name')} |
| Event Type | {event_type} |
| Internal Reference Name | {theme.get('internal_reference_name')} |
| Sprint Duration | {duration_days} days |

## 2. Theme & Reasoning

{theme.get('reasoning')}

**Theme Visual Direction**

- Art Style: {vd.get('art_style')}
- Color Palette: {vd.get('color_palette')}

**Audience Fit** (Green / Yellow / Red per the studio's Ideation Checklist)

- Progression Completionist: {af.get('progression_completionist')}
- Casual Escape: {af.get('casual_escape')}
- Social/Collector: {af.get('social_collector')}

{af.get('notes') or ''}

**Key Features:**
{key_features}

**UI Direction:**
{ui_direction}

**Scene Progression:**

{scenes}

## 3. New Features & Additions

{theme.get('new_features')}

## 4. Product Alignment

**Expected Impact** (draft estimate — needs PO/Economy Designer validation, per the checklist's "Honest Assessment" step)

- Engagement: {imp.get('engagement_note')}
- Monetization: {imp.get('monetization_note')}

## 5. Event Item Visual References

*Not yet produced at ideation stage — per the studio's checklist, attach before submission:*

- [ ] Event icon (bottom-center button)
- [ ] Event Packs mockups
- [ ] Event rewards cosmetics (card frames/backs)
- [ ] Scene building elements / references

## 6. Effort Estimation

| Area | Effort |
|---|---|
| Art | {eff.get('art')} |
| Engineering | {eff.get('engineering')} |
| Design | {eff.get('design')} |
| Overall Complexity | {eff.get('overall')} |

## 7. Checklist: Ready for Submission?

- [x] Theme has reasoning (AI-drafted — verify against real audience/product knowledge)
- [x] Audience impact assessed (Green/Yellow/Red for all archetypes)
- [x] Effort sizing provided
- [ ] Visual references attached (mood board, sketches) — not yet done
- [ ] Jira ticket linked — not yet done

---
AI-drafted via the Event Automator Ideation tab. Treat as a first draft for PO review, not an approved spec — per the studio's Event Ideation Checklist, "if you can't explain why, it's not ready for submission."
"""


def _write_theme_draft(events_root, event_type, theme, duration_days, author_name=""):
    """Writes every generated/regenerated theme to disk immediately, so
    feedback happens against an actual document instead of only the
    in-app card. Lives in a Drafts/ subfolder of the same "One-Pager
    Ideation" folder the real PDFs use, distinct from the final file
    "Select as Final" writes."""
    if not events_root:
        return None
    ref_name = re.sub(r'[\\/:*?"<>|]', '_', theme.get("internal_reference_name") or "UnnamedEvent")
    draft_dir = os.path.join(events_root, event_type, "One-Pager Ideation", "Drafts")
    os.makedirs(draft_dir, exist_ok=True)
    draft_path = os.path.join(draft_dir, f"{ref_name} Draft.md")
    md = _render_theme_markdown(theme, event_type, duration_days, "Draft — under review, not yet selected", author_name)
    with open(draft_path, "w", encoding="utf-8") as f:
        f.write(md)
    return draft_path


def _build_system_prompt(known_lessons_block, template_text):
    template_block = (
        f"\n\nFollow this exact section structure - it is the studio's real "
        f"1-Pager template, verbatim:\n---\n{template_text}\n---"
        if template_text else ""
    )
    return (
        "You are an experienced game event designer at a mobile card-game studio "
        "(Skip-Bo, a casual solitaire/rummy-style card game) writing an Event "
        "Ideation 1-Pager - the studio's real internal document for pitching a "
        "new limited-time event theme to a Product Owner before any art or "
        "engineering work begins. Follow the studio's actual 1-Pager structure: "
        "a theme name and internal reference code, a short reasoning paragraph, "
        "an audience-fit rating (Green/Yellow/Red) for three player archetypes - "
        "Progression Completionist, Casual Escape, and Social/Collector - visual "
        "direction (art style, palette, key scene features, UI direction), a "
        "3-scene visual progression (the board escalates across three scenes as "
        "players advance, matching how this studio's real events are built), a "
        "new-features note, a qualitative expected-impact note (never invent a "
        "precise ARPU/DAU number - flag it as a draft estimate for the Product "
        "Owner and Economy Designer to validate), and an Art/Engineering/Design/"
        "Overall effort estimate (Low/Medium/High).\n\n"
        f"{AUDIENCE_GUIDANCE}"
        f"{known_lessons_block}"
        f"{template_block}"
    )


def _extract_json(response):
    if response.stop_reason == "refusal":
        raise RuntimeError(
            "The generation request was declined by Claude's safety classifiers. "
            "This is usually a false positive on an unrelated theme - try again "
            "or adjust the event type."
        )
    for block in response.content:
        if block.type == "text":
            return json.loads(block.text)
    raise RuntimeError("No text content in the model response.")


def _client(api_key):
    if not api_key:
        raise RuntimeError("No Anthropic API key configured. Add one in Settings.")
    return anthropic.Anthropic(api_key=api_key)


def _call_claude(api_key, system_prompt, user_message, schema, max_tokens):
    client = _client(api_key)
    # Streamed internally (still one buffered JSON line on our own stdout -
    # no change to the Electron side): theme_count is now user-controlled,
    # so max_tokens can run well past the ~16K non-streaming SDK guard.
    with client.messages.stream(
        model=MODEL,
        max_tokens=max_tokens,
        system=system_prompt,
        messages=[{"role": "user", "content": user_message}],
        output_config={"format": {"type": "json_schema", "schema": schema}},
    ) as stream:
        response = stream.get_final_message()
    return _extract_json(response)


def _call_ollama(base_url, model, system_prompt, user_message, schema):
    """Local fallback for when no Anthropic key is configured - a shared
    Ollama server (e.g. on the office LAN) running a model like qwen3:32b.
    Raw HTTP via stdlib: Ollama's API is a single unauthenticated JSON POST,
    and pulling in the `ollama` pip package (another compiled-wheel surface
    to keep working on the embeddable interpreter) buys nothing over that.
    Ollama's `format` field accepts the same JSON Schema Claude's
    output_config.format does, so the schemas are shared as-is between the
    two providers - no fallback-specific schema to maintain."""
    import urllib.request
    import urllib.error

    if not base_url:
        raise RuntimeError("No Ollama Base URL configured. Add one in Settings.")

    url = base_url.rstrip("/") + "/api/chat"
    body = json.dumps({
        "model": model,
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_message},
        ],
        "format": schema,
        "stream": False,
    }).encode("utf-8")
    req = urllib.request.Request(url, data=body, headers={"Content-Type": "application/json"}, method="POST")

    try:
        # Local network call, not billed per token - a generous timeout costs
        # nothing and a 32B model on modest hardware can genuinely take minutes.
        with urllib.request.urlopen(req, timeout=600) as resp:
            data = json.loads(resp.read().decode("utf-8"))
    except urllib.error.URLError as e:
        raise RuntimeError(f"Could not reach Ollama at {base_url}: {e.reason if hasattr(e, 'reason') else e}")
    except TimeoutError:
        raise RuntimeError(f"Ollama at {base_url} timed out (10 min). The model may be overloaded or too slow on this hardware.")

    content = (data.get("message") or {}).get("content", "")
    if not content:
        raise RuntimeError("Ollama returned an empty response - check the server logs.")
    try:
        return json.loads(content)
    except json.JSONDecodeError:
        raise RuntimeError("Ollama's response wasn't valid JSON even with a schema - this model may not support structured outputs well.")


def _generate_themes(provider, payload, system_prompt, user_message, theme_count):
    schema = _generate_schema(theme_count)
    if provider == "ollama":
        return _call_ollama(
            payload.get("ollama_base_url", ""),
            payload.get("ollama_model") or "qwen3:32b",
            system_prompt, user_message, schema
        )
    # ~2.6K tokens/theme plus a fixed overhead, covers the fuller 7-section
    # document per theme now that drafts get written verbatim to disk too.
    max_tokens = min(4000 + 2600 * theme_count, 100000)
    return _call_claude(payload.get("api_key", ""), system_prompt, user_message, schema, max_tokens)


def _regenerate_theme(provider, payload, system_prompt, user_message):
    if provider == "ollama":
        return _call_ollama(
            payload.get("ollama_base_url", ""),
            payload.get("ollama_model") or "qwen3:32b",
            system_prompt, user_message, REGENERATE_SCHEMA
        )
    return _call_claude(payload.get("api_key", ""), system_prompt, user_message, REGENERATE_SCHEMA, 8000)


def generate_main():
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    if len(sys.argv) < 2:
        print(json.dumps({"error": "Usage: ideation_generate <payload.json>"}))
        sys.exit(1)

    with open(sys.argv[1], "r", encoding="utf-8-sig") as f:
        payload = json.load(f)

    event_type = payload.get("event_type", "")
    duration_days = payload.get("duration_days", 14)
    design_path = payload.get("design_path", "")
    events_root = payload.get("events_root", "")
    provider = payload.get("provider") or "anthropic"
    author_id = payload.get("author_id", "")
    author_name = payload.get("author_name", "")
    try:
        theme_count = max(1, min(10, int(payload.get("theme_count", 5))))
    except (TypeError, ValueError):
        theme_count = 5

    if event_type not in KNOWN_TYPES:
        print(json.dumps({"error": f"Unknown event type: {event_type}"}))
        sys.exit(1)

    try:
        lessons_block = store.lessons_prompt_block(events_root, event_type)
        template_text = _read_template(events_root, "Event_Ideation_1Pager.md")
        recent_events = _recent_events_for_type(design_path, event_type)
        past_themes_txt = store.past_themes_block(events_root, event_type)

        recent_txt = (
            "\n".join(f"- {e['alias']} (started {e['start_time'] or 'unknown date'})" for e in recent_events)
            or "(no recent event history found)"
        )

        # Grounds scenes/visual_direction in what this event type's Art
        # Requirements actually look like once any exist (this tool's own
        # prior output - see art_requirement_generator.py). Nothing exists
        # on a type's first-ever run, which is expected, not an error.
        art_req_folder = os.path.join(events_root, event_type, "Art Requirements") if events_root else ""
        _, art_req_text = find_latest_doc(art_req_folder)
        art_req_block = (
            f"\n\nFor visual/scene grounding only (do not copy content, just match the "
            f"level of concreteness), here is the most recent Art Requirement doc this "
            f"studio produced for a {event_type} event:\n---\n{art_req_text[:6000]}\n---"
            if art_req_text else ""
        )

        user_message = (
            f"Event type: {event_type}\n"
            f"Sprint duration: {duration_days} days\n\n"
            f"Recent {event_type} events shipped (codes/dates only, for calendar context):\n{recent_txt}\n\n"
            f"Themes already used for {event_type} - do NOT repeat any of these, or propose a "
            f"close variant of one (same setting/hook with a different label counts as a repeat):\n{past_themes_txt}\n\n"
            f"Generate exactly {theme_count} distinct theme options for this event, each as a "
            f"complete draft 1-Pager, ready for a Product Owner to review. All {theme_count} must be "
            f"genuinely distinct from each other too - vary the setting, mood, and visual palette, not just the label."
            f"{art_req_block}"
        )

        result = _generate_themes(provider, payload, _build_system_prompt(lessons_block, template_text), user_message, theme_count)
        themes = result["themes"]

        draft_paths = []
        if events_root:
            for t in themes:
                p = _write_theme_draft(events_root, event_type, t, duration_days, author_name)
                if p:
                    draft_paths.append(p)
            store.record_generated_themes(events_root, event_type, themes, author_id, author_name)

        print(json.dumps({"themes": themes, "provider": provider, "draft_paths": draft_paths}, ensure_ascii=False))
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
        print(json.dumps({"error": "Usage: ideation_regenerate <payload.json>"}))
        sys.exit(1)

    with open(sys.argv[1], "r", encoding="utf-8-sig") as f:
        payload = json.load(f)

    event_type = payload.get("event_type", "")
    duration_days = payload.get("duration_days", 14)
    events_root = payload.get("events_root", "")
    provider = payload.get("provider") or "anthropic"
    author_id = payload.get("author_id", "")
    author_name = payload.get("author_name", "")
    theme = payload.get("theme")
    feedback_text = (payload.get("feedback") or "").strip()

    if not theme:
        print(json.dumps({"error": "No theme supplied to regenerate."}))
        sys.exit(1)
    if not feedback_text:
        print(json.dumps({"error": "No feedback supplied."}))
        sys.exit(1)

    try:
        lessons_block = store.lessons_prompt_block(events_root, event_type)
        template_text = _read_template(events_root, "Event_Ideation_1Pager.md")

        user_message = (
            f"Event type: {event_type}\n"
            f"Sprint duration: {duration_days} days\n\n"
            f"Here is the current theme draft:\n{json.dumps(theme, ensure_ascii=False)}\n\n"
            f"Reviewer feedback to address:\n\"{feedback_text}\"\n\n"
            "Rewrite the theme so it fully addresses this feedback - every point, "
            "not just the easiest one. Keep everything the feedback didn't ask to "
            "change. Also extract one short, generalizable lesson from this "
            "feedback (see the extracted_lesson field) so the same mistake isn't "
            "repeated on a future, unrelated theme."
        )

        result = _regenerate_theme(provider, payload, _build_system_prompt(lessons_block, template_text), user_message)
        new_theme = result["theme"]

        lesson = (result.get("extracted_lesson") or "").strip()
        draft_path = None
        ref_name = theme.get("internal_reference_name", "")
        if events_root:
            store.write_feedback_entry(
                events_root, event_type, "ideation", ref_name,
                theme.get("theme_name", ""), feedback_text, lesson,
                author_id, author_name,
                revised_title=new_theme.get("theme_name", ""),
            )
            write_feedback_digest(events_root, event_type, ref_name, "ideation", "Ideation")
            draft_path = _write_theme_draft(events_root, event_type, new_theme, duration_days, author_name)
            store.record_generated_themes(events_root, event_type, [new_theme], author_id, author_name)

        print(json.dumps({"theme": new_theme, "provider": provider, "draft_path": draft_path}, ensure_ascii=False))
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


def entries_main():
    """Read-only: the feedback thread + active lessons for a type, so the UI
    can show who said what and offer to retract a bad lesson. No LLM call."""
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    if len(sys.argv) < 2:
        print(json.dumps({"error": "Usage: ideation_entries <payload.json>"}))
        sys.exit(1)

    with open(sys.argv[1], "r", encoding="utf-8-sig") as f:
        payload = json.load(f)

    events_root = payload.get("events_root", "")
    event_type = payload.get("event_type", "")
    doc_kind = payload.get("doc_kind") or None
    ref_name = payload.get("ref_name") or None

    try:
        retracted = store.read_retractions(events_root, event_type)
        entries = []
        for e in store.read_feedback_entries(events_root, event_type, doc_kind=doc_kind, ref_name=ref_name):
            r = retracted.get(e.get("entry_id"))
            entries.append({**e, "retracted": bool(r), "retracted_by": (r or {}).get("author_name")})
        # Profiles ride along with the thread so the UI can draw each
        # author's avatar without a second round trip per render.
        print(json.dumps({
            "entries": entries,
            "lessons": store.active_lessons(events_root, event_type),
            "profiles": store.read_profiles(events_root),
        }, ensure_ascii=False))
    except Exception as e:
        print(json.dumps({"error": str(e)}))
        sys.exit(1)


def retract_main():
    """Retract a lesson so it stops being folded into future prompts. Writes
    a separate retraction file rather than editing the original entry - two
    people retracting different lessons must never touch the same file."""
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    if len(sys.argv) < 2:
        print(json.dumps({"error": "Usage: ideation_retract <payload.json>"}))
        sys.exit(1)

    with open(sys.argv[1], "r", encoding="utf-8-sig") as f:
        payload = json.load(f)

    events_root = payload.get("events_root", "")
    event_type = payload.get("event_type", "")
    entry_id = payload.get("entry_id", "")

    if not events_root or not event_type or not entry_id:
        print(json.dumps({"error": "Missing events_root, event_type, or entry_id."}))
        sys.exit(1)

    try:
        store.write_retraction(events_root, event_type, entry_id,
                               payload.get("author_id", ""), payload.get("author_name", ""),
                               payload.get("reason", ""))
        # Keep the human-readable digests in step with the retraction.
        for e in store.read_feedback_entries(events_root, event_type):
            if e.get("entry_id") == entry_id:
                kind_label = "Art Requirement" if e.get("doc_kind") == "artreq" else "Ideation"
                write_feedback_digest(events_root, event_type, e.get("ref_name"), e.get("doc_kind"), kind_label)
                break
        print(json.dumps({"ok": True, "lessons": store.active_lessons(events_root, event_type)}, ensure_ascii=False))
    except Exception as e:
        print(json.dumps({"error": str(e)}))
        sys.exit(1)


def profile_save_main():
    """Write your own profile (name + avatar) into the shared folder so
    colleagues see it beside your feedback. No LLM call."""
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    if len(sys.argv) < 2:
        print(json.dumps({"error": "Usage: profile_save <payload.json>"}))
        sys.exit(1)

    with open(sys.argv[1], "r", encoding="utf-8-sig") as f:
        payload = json.load(f)

    events_root = payload.get("events_root", "")
    author_id = payload.get("author_id", "")
    if not events_root or not author_id:
        print(json.dumps({"error": "Missing events_root or author_id."}))
        sys.exit(1)

    try:
        rec = store.write_profile(
            events_root, author_id,
            payload.get("display_name", ""),
            payload.get("avatar_color", ""),
            payload.get("avatar_image", ""),
        )
        print(json.dumps({"profile": rec}, ensure_ascii=False))
    except Exception as e:
        print(json.dumps({"error": str(e)}))
        sys.exit(1)


def profiles_read_main():
    """Every known profile in the shared folder, for the sidebar chip and
    the avatars in feedback threads."""
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    if len(sys.argv) < 2:
        print(json.dumps({"error": "Usage: profiles_read <payload.json>"}))
        sys.exit(1)

    with open(sys.argv[1], "r", encoding="utf-8-sig") as f:
        payload = json.load(f)

    try:
        print(json.dumps({"profiles": store.read_profiles(payload.get("events_root", ""))}, ensure_ascii=False))
    except Exception as e:
        print(json.dumps({"error": str(e)}))
        sys.exit(1)
