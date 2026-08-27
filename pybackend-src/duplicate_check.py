"""Detect the same building drawn twice inside ONE generated scene.

This is the standing failure of the "EXACTLY N buildings, none twice" rule. It
has now happened three times: the crane nesting tower drawn on both sides of a
Merge origami scene, the terracotta kiln dome drawn twice in a mosaic scene, and
a Sort floor that repeated a burner. Every one was caught by eye.

`artreq_images.check_generated` already has a duplicate test, but it compares
whole FRAMES against each other - it answers "did two zones come back as the
same picture", never "does this one picture repeat itself".

WHY THIS IS NOT A PIXEL CHECK
-----------------------------
Block self-similarity was tried properly and does not work. Four configurations
were measured against real scenes plus synthetic duplicates:

  descriptor          stride   clean max   known duplicate pair
  grey 16x16          B/3      0.76-0.93   0.31-0.59
  grey 16x16          B/10     0.80-0.94   0.66-0.82
  gradient 16x16      B/3      0.65-0.75   0.07-0.47
  gradient 16x16      B/10     0.72-0.77   0.64-0.72
  grey + edge filter  B/10     0.53-0.86   0.56-0.87  (overlapping)

In every one the most similar pair of blocks in a CLEAN scene was at least as
similar as two genuine copies of one building. The reason is structural: these
scenes contain large smooth regions - sky, grass, paved plaza - and repeated
dressing like trees and benches, and those correlate more strongly with each
other than two independently drawn copies of a building do, because the copies
differ in lighting, angle and detail. Filtering candidate blocks by edge density
removed the sky matches and still left the ranges overlapping. There is no
threshold that separates them. Do not re-try block self-similarity.

WHAT THIS DOES
--------------
Asks a vision model, which reads the scene semantically the way the reviewer
does. It also counts the buildings, which is the actual rule being enforced.

Costs one vision API call per image, so it is OPT-IN: no api_key, no check, and
the caller is told it did not run rather than being told the image is clean.

PRIVACY: this uploads the artwork to Anthropic. That is the same exposure the
pipeline already has via ComfyUI's NanoBanana node, which forwards prompts and
reference images to Google. Unannounced event art leaves the network either way -
worth knowing, not new.
"""

import base64
import io
import json
import os

MAX_EDGE = 1400            # Claude downsamples above ~1568 anyway; this keeps the
                           # upload small without throwing away building detail
JPEG_QUALITY = 88

SCHEMA = {
    "type": "object",
    "additionalProperties": False,
    "required": ["duplicate_found", "building_count", "repeated", "confidence", "notes"],
    "properties": {
        "duplicate_found": {
            "type": "boolean",
            "description": "True only if two or more separate BUILDINGS in the scene are "
                           "the same design drawn more than once.",
        },
        "building_count": {
            "type": "integer",
            "description": "How many distinct buildable structures are in the scene. "
                           "Count buildings only, never trees, benches, lamps, planters, "
                           "fences, paths or loose dressing.",
        },
        "repeated": {
            "type": "array",
            "description": "One entry per repeated design. Empty when there is none.",
            "items": {
                "type": "object",
                "additionalProperties": False,
                "required": ["what", "where"],
                "properties": {
                    "what": {"type": "string",
                             "description": "The repeated building, in a few words."},
                    "where": {"type": "string",
                              "description": "Where each copy sits, e.g. 'upper left and "
                                             "lower left'."},
                },
            },
        },
        "confidence": {"type": "string", "enum": ["high", "medium", "low"]},
        "notes": {"type": "string", "description": "One sentence. Empty if nothing to add."},
    },
}

SYSTEM = """You are reviewing generated isometric game art for a single specific defect: the same
building design drawn more than once in one scene.

The brief for these scenes always says every building appears EXACTLY ONCE. Two copies of one
design is a defect that has to be caught before the art is handed over.

What counts as a repeat:
- Two structures with the same silhouette, roof shape and fittings, even if they differ in
  colour, size or which way they face.
- A design that clearly reads as "the same building again" to a person glancing at the scene.

What does NOT count, and must never be reported:
- Repeated small dressing: trees, bushes, planters, benches, lamp posts, fences, barrels, crates,
  paths, paving, market goods, bunting. These are SUPPOSED to repeat.
- Two different buildings that merely share a material or palette. Everything in these scenes is
  built from one material by design, so shared material is not a repeat.
- Symmetry within a single building, such as matching towers on one structure.

Count buildings conservatively: a buildable structure a player would construct. An arch, a
gateway, a stall, a tower and a kiln all count. A tree does not.

Be decisive. If you are unsure whether two structures are the same design, say duplicate_found
false and set confidence to low, and explain the doubt in notes."""


def _encode(path, max_edge=MAX_EDGE, quality=JPEG_QUALITY):
    from PIL import Image
    im = Image.open(path).convert("RGB")
    if max(im.size) > max_edge:
        s = max_edge / float(max(im.size))
        im = im.resize((max(1, int(im.width * s)), max(1, int(im.height * s))), Image.LANCZOS)
    buf = io.BytesIO()
    im.save(buf, "JPEG", quality=quality)
    return base64.standard_b64encode(buf.getvalue()).decode("ascii")


def build_user_message(b64, expected_count=None):
    """The vision message. Split out so it is testable without an API call."""
    ask = "Review this scene for a building drawn more than once, and count the buildings."
    if expected_count:
        ask += (" The brief for this scene specifies EXACTLY %d buildings, each appearing once. "
                "If you count a different number, that itself is worth reporting in notes."
                % int(expected_count))
    return [
        {"type": "image",
         "source": {"type": "base64", "media_type": "image/jpeg", "data": b64}},
        {"type": "text", "text": ask},
    ]


def _skipped(why):
    return {"ran": False, "duplicate_found": None, "building_count": None,
            "repeated": [], "confidence": None, "detail": why}


def _ask_claude(api_key, b64, expected_count, max_tokens):
    from ideation_generator import _client, _extract_json, MODEL
    client = _client(api_key)
    with client.messages.stream(
        model=MODEL,
        max_tokens=max_tokens,
        system=SYSTEM,
        messages=[{"role": "user", "content": build_user_message(b64, expected_count)}],
        output_config={"format": {"type": "json_schema", "schema": SCHEMA}},
    ) as stream:
        response = stream.get_final_message()
    return _extract_json(response)


def _ask_ollama(base_url, model, b64, expected_count, timeout=180):
    """Ollama's vision path: the image rides on the message as base64.

    Supported because this is the provider the app is actually configured for,
    and because a LAN model keeps the artwork on the network - see PRIVACY above.
    The model must be vision-capable; a text-only one (qwen3:32b) will simply
    ignore the image and its answer must not be trusted.
    """
    import urllib.request

    if not base_url:
        raise RuntimeError("No Ollama Base URL configured. Add one in Settings.")
    ask = build_user_message(b64, expected_count)[1]["text"]
    body = json.dumps({
        "model": model,
        "stream": False,
        "format": SCHEMA,
        "messages": [{"role": "system", "content": SYSTEM},
                     {"role": "user", "content": ask, "images": [b64]}],
    }).encode("utf-8")
    req = urllib.request.Request(base_url.rstrip("/") + "/api/chat", data=body,
                                 headers={"Content-Type": "application/json"}, method="POST")
    with urllib.request.urlopen(req, timeout=timeout) as r:
        res = json.loads(r.read().decode("utf-8"))
    return json.loads((res.get("message") or {}).get("content") or "{}")


def check(path, provider="anthropic", api_key="", ollama_base_url="", ollama_model="",
          expected_count=None, max_tokens=800):
    """Is any building drawn twice in this one image?

    Returns {"ran", "duplicate_found", "building_count", "repeated", "confidence",
             "detail"}.

    ran is False when no provider was configured or the call failed - and
    duplicate_found is then None, never False, so "did not check" can never be
    mistaken for "checked and clean".
    """
    provider = (provider or "anthropic").lower()
    if provider == "ollama":
        if not ollama_base_url:
            return _skipped("no Ollama URL configured, so the duplicate check did not run")
    elif not api_key:
        return _skipped("no API key supplied, so the within-frame duplicate check did not run")
    if not os.path.isfile(path):
        return _skipped("image not found: %s" % os.path.basename(path))

    try:
        b64 = _encode(path)
        if provider == "ollama":
            data = _ask_ollama(ollama_base_url, ollama_model or "qwen2.5vl:32b",
                               b64, expected_count)
        else:
            data = _ask_claude(api_key, b64, expected_count, max_tokens)
    except Exception as e:
        return _skipped("duplicate check failed (%s: %s)"
                        % (e.__class__.__name__, str(e)[:160]))
    return _parse(data, expected_count)


def _parse(data, expected_count=None):
    """Turn a provider response into the result dict. Split out so the whole
    result contract is testable without reaching a model."""
    data = data or {}
    dup = bool(data.get("duplicate_found"))
    reps = data.get("repeated") or []
    count = data.get("building_count")
    conf = data.get("confidence") or "low"

    if dup:
        detail = ("the same building appears more than once: %s. Every building in the brief "
                  "appears exactly once."
                  % "; ".join("%s (%s)" % (r.get("what", "?"), r.get("where", "?"))
                              for r in reps[:3]))
    else:
        detail = "no repeated building found"
    if count is not None:
        detail += " Counted %d building(s)." % count
        if expected_count and int(count) != int(expected_count):
            detail += " The brief asked for %d." % int(expected_count)
    if data.get("notes"):
        detail += " " + str(data["notes"])

    return {"ran": True, "duplicate_found": dup, "building_count": count,
            "repeated": reps, "confidence": conf, "detail": detail}
