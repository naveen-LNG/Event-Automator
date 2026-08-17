"""
Shared, conflict-free storage for ideation/art-requirement feedback,
extracted lessons, and generated-theme history.

Why file-per-entry instead of one aggregate JSON:

This data is meant to be shared between designers via a synced folder
(OneDrive for Business, or a network share - the root path is configurable
in Settings). File-sync services resolve simultaneous edits to the SAME file
by silently creating "conflicted copy" duplicates, so a single mutable
_ideation_feedback_store.json that everyone appends to would quietly lose
one designer's feedback whenever two people reviewed on the same afternoon.

So every write here creates its own immutable file, named with the author
and a UTC timestamp. Two people never touch the same file, so there is
nothing to conflict. Aggregates (the lesson list folded into prompts, the
theme history used for de-duplication) are derived at read time by globbing
the folder. Same property makes this safe under SVN/git if the folder is
ever version-controlled instead.

Retraction follows the same rule: a bad lesson is not edited out of its
entry (another author might be doing the same thing concurrently) - a
separate retraction file is written referencing the entry id, and read-time
aggregation filters retracted lessons out.

Layout under <root>/Events/:
  <Type>/Feedback/entries/<ts>_<author>_<ref>_<kind>.json
  <Type>/Feedback/retractions/<ts>_<author>_retract_<entry_id>.json
  _theme_history/<Type>/<ts>_<author>.json
"""
import os
import re
import json
import glob
from datetime import datetime, timezone

SCHEMA_VERSION = 1

# Keep the prompt from being swamped once a type accumulates a lot of
# history; both are "most recent wins" trims applied at read time.
MAX_LESSONS_IN_PROMPT = 40
MAX_THEME_HISTORY_PER_TYPE = 60


def utc_stamp():
    return datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")


def iso_now():
    return datetime.now(timezone.utc).isoformat(timespec="seconds").replace("+00:00", "Z")


def safe_token(value, fallback="unknown", maxlen=40):
    """Filename-safe slug. Author ids look like 'AzureAD\\NaveenKumar', and
    titles can contain anything a designer typed."""
    token = re.sub(r"[^A-Za-z0-9._-]+", "-", str(value or "")).strip("-")
    return (token[:maxlen] or fallback)


def _read_json(path):
    try:
        with open(path, "r", encoding="utf-8-sig") as f:
            return json.load(f)
    except Exception:
        # A half-synced or hand-edited file must never take down a
        # generation run - skip it and keep reading the rest.
        return None


def _feedback_dir(events_root, event_type, sub):
    return os.path.join(events_root, event_type, "Feedback", sub)


# ---------------- Feedback entries ----------------

def write_feedback_entry(events_root, event_type, doc_kind, ref_name, subject_title,
                         feedback, extracted_lesson, author_id, author_name,
                         revised_title=None):
    """One immutable file per piece of feedback. Returns the entry dict."""
    if not events_root:
        return None
    ts = utc_stamp()
    entry_id = f"{ts}_{safe_token(author_id)}_{safe_token(ref_name)}_{safe_token(doc_kind)}"
    entry = {
        "schema": SCHEMA_VERSION,
        "entry_id": entry_id,
        "author_id": author_id or "unknown",
        "author_name": author_name or author_id or "Unknown",
        "created_at": iso_now(),
        "event_type": event_type,
        "doc_kind": doc_kind,          # "ideation" | "artreq"
        "ref_name": ref_name,
        "subject_title": subject_title,
        "revised_title": revised_title,
        "feedback": feedback,
        "extracted_lesson": extracted_lesson or "",
    }
    d = _feedback_dir(events_root, event_type, "entries")
    os.makedirs(d, exist_ok=True)
    with open(os.path.join(d, f"{entry_id}.json"), "w", encoding="utf-8") as f:
        json.dump(entry, f, indent=2, ensure_ascii=False)
    return entry


def read_feedback_entries(events_root, event_type, doc_kind=None, ref_name=None):
    """All entries for a type, newest last. Optional filters for showing a
    single document's thread in the UI."""
    if not events_root:
        return []
    d = _feedback_dir(events_root, event_type, "entries")
    entries = []
    for path in sorted(glob.glob(os.path.join(d, "*.json"))):
        e = _read_json(path)
        if not e:
            continue
        if doc_kind and e.get("doc_kind") != doc_kind:
            continue
        if ref_name and e.get("ref_name") != ref_name:
            continue
        entries.append(e)
    entries.sort(key=lambda e: e.get("created_at", ""))
    return entries


# ---------------- Retractions ----------------

def write_retraction(events_root, event_type, entry_id, author_id, author_name, reason=""):
    if not events_root or not entry_id:
        return None
    ts = utc_stamp()
    rec = {
        "schema": SCHEMA_VERSION,
        "retracts_entry_id": entry_id,
        "author_id": author_id or "unknown",
        "author_name": author_name or author_id or "Unknown",
        "created_at": iso_now(),
        "reason": reason or "",
    }
    d = _feedback_dir(events_root, event_type, "retractions")
    os.makedirs(d, exist_ok=True)
    fname = f"{ts}_{safe_token(author_id)}_retract_{safe_token(entry_id, maxlen=70)}.json"
    with open(os.path.join(d, fname), "w", encoding="utf-8") as f:
        json.dump(rec, f, indent=2, ensure_ascii=False)
    return rec


def read_retractions(events_root, event_type):
    """Map of retracted entry_id -> retraction record."""
    if not events_root:
        return {}
    d = _feedback_dir(events_root, event_type, "retractions")
    out = {}
    for path in sorted(glob.glob(os.path.join(d, "*.json"))):
        r = _read_json(path)
        if r and r.get("retracts_entry_id"):
            out[r["retracts_entry_id"]] = r
    return out


# ---------------- Derived: lessons ----------------

def active_lessons(events_root, event_type):
    """Lessons that still apply, newest last, retracted ones removed.
    Each item carries its author so the UI can attribute and retract it."""
    retracted = read_retractions(events_root, event_type)
    lessons = []
    for e in read_feedback_entries(events_root, event_type):
        lesson = (e.get("extracted_lesson") or "").strip()
        if not lesson or e.get("entry_id") in retracted:
            continue
        lessons.append({
            "entry_id": e.get("entry_id"),
            "lesson": lesson,
            # author_id (not just the name) so the UI can match this to a
            # profile and draw the right avatar.
            "author_id": e.get("author_id"),
            "author_name": e.get("author_name"),
            "created_at": e.get("created_at"),
            "doc_kind": e.get("doc_kind"),
        })
    return lessons[-MAX_LESSONS_IN_PROMPT:]


def lessons_prompt_block(events_root, event_type, doc_kind=None):
    """The '- lesson' bullet list folded into the system prompt. doc_kind
    filters to lessons from that pipeline; None uses all of them."""
    items = active_lessons(events_root, event_type)
    if doc_kind:
        items = [l for l in items if l.get("doc_kind") in (doc_kind, None)]
    if not items:
        return ""
    bullets = "\n".join(f"- {l['lesson']}" for l in items)
    return (
        "\n\nLessons from past feedback on this studio's events - do not repeat "
        f"these mistakes:\n{bullets}"
    )


# ---------------- Derived: generated-theme history ----------------

def _theme_history_dir(events_root, event_type):
    return os.path.join(events_root, "_theme_history", event_type)


def record_generated_themes(events_root, event_type, themes, author_id, author_name):
    """One file per generation run, not one mutable list for everyone."""
    if not events_root or not themes:
        return None
    ts = utc_stamp()
    rec = {
        "schema": SCHEMA_VERSION,
        "created_at": iso_now(),
        "author_id": author_id or "unknown",
        "author_name": author_name or author_id or "Unknown",
        "event_type": event_type,
        "themes": [
            {
                "theme_name": t.get("theme_name", ""),
                "internal_reference_name": t.get("internal_reference_name", ""),
            }
            for t in themes
        ],
    }
    d = _theme_history_dir(events_root, event_type)
    os.makedirs(d, exist_ok=True)
    with open(os.path.join(d, f"{ts}_{safe_token(author_id)}.json"), "w", encoding="utf-8") as f:
        json.dump(rec, f, indent=2, ensure_ascii=False)
    return rec


def past_themes(events_root, event_type):
    """Flat list of every theme generated for this type, newest last."""
    if not events_root:
        return []
    out = []
    for path in sorted(glob.glob(os.path.join(_theme_history_dir(events_root, event_type), "*.json"))):
        rec = _read_json(path)
        if not rec:
            continue
        for t in rec.get("themes", []):
            out.append({
                "theme_name": t.get("theme_name", ""),
                "internal_reference_name": t.get("internal_reference_name", ""),
                "author_name": rec.get("author_name"),
            })
    return out[-MAX_THEME_HISTORY_PER_TYPE:]


# ---------------- Profiles ----------------
# One file per person, in the shared folder rather than local config, so a
# colleague's name and avatar show up next to their feedback on everyone
# else's machine. Same file-per-entry rule: you only ever write your own
# file, so profiles can't conflict either.

def _profiles_dir(events_root):
    return os.path.join(events_root, "_profiles")


def write_profile(events_root, author_id, display_name, avatar_color, avatar_image=""):
    if not events_root or not author_id:
        return None
    rec = {
        "schema": SCHEMA_VERSION,
        "author_id": author_id,
        "display_name": display_name or author_id,
        "avatar_color": avatar_color or "#7c5cfc",
        # Downscaled PNG data URI written by the renderer, or "" for the
        # initials-on-colour fallback. Kept inline rather than as a sibling
        # image file so a profile is a single self-contained sync unit.
        "avatar_image": avatar_image or "",
        "updated_at": iso_now(),
    }
    d = _profiles_dir(events_root)
    os.makedirs(d, exist_ok=True)
    with open(os.path.join(d, f"{safe_token(author_id, maxlen=80)}.json"), "w", encoding="utf-8") as f:
        json.dump(rec, f, indent=2, ensure_ascii=False)
    return rec


def read_profiles(events_root):
    """All known profiles, keyed by author_id."""
    if not events_root:
        return {}
    out = {}
    for path in sorted(glob.glob(os.path.join(_profiles_dir(events_root), "*.json"))):
        rec = _read_json(path)
        if rec and rec.get("author_id"):
            out[rec["author_id"]] = rec
    return out


def past_themes_block(events_root, event_type):
    items = past_themes(events_root, event_type)
    if not items:
        return "(no themes generated for this event type yet - nothing to avoid repeating)"
    return "\n".join(
        f"- {t['theme_name']} ({t['internal_reference_name']})" for t in items
    )
