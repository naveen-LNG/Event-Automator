"""Reference-image generation for Art Requirement documents.

Renders one image per zone of a finalized/draft art requirement variation via
the studio's ComfyUI Nano Banana workflow, writes them beside the document, and
returns the paths so the renderer can embed them.

Prompts are grounded in three real artefacts, never invented:
  - StyleReference/ROOM_RENDER_STYLE.md  (the approved render spec)
  - StyleReference/images/*              (attached as reference image inputs)
  - <Type>/Art Requirements/*.pdf        (the type's own shipped documents)

CLI payload (JSON file, same convention as the other backend scripts):
    { comfy_url, events_root, event_type, doc (the variation object),
      index, api_key | provider/ollama_*, aspect_ratio?, image_size?,
      max_images? }
"""

import json
import os
import re
import sys

import art_reference
import comfy_client
import ideation_store as store
from ideation_generator import log


def _slug(s, limit=48):
    s = re.sub(r"[^A-Za-z0-9]+", "_", str(s or "")).strip("_")
    return (s[:limit] or "untitled").lower()


def _vision_cfg(payload):
    """Provider config for the within-frame duplicate check, or None.

    None when nothing is configured, so the reviewer reports "did not check"
    instead of running and failing. Mirrors _distiller: the app is dual-provider
    and an Ollama vision model keeps the artwork on the LAN.
    """
    provider = (payload.get("provider") or "anthropic").lower()
    if provider == "ollama":
        if not payload.get("ollama_base_url"):
            return None
        return {"provider": "ollama",
                "ollama_base_url": payload.get("ollama_base_url", ""),
                "ollama_model": payload.get("ollama_vision_model")
                                or payload.get("ollama_model") or ""}
    if not payload.get("api_key"):
        return None
    return {"provider": "anthropic", "api_key": payload.get("api_key", "")}


def _distiller(payload):
    """Adapter so art_reference can distil PDFs with whichever provider the
    app is configured for, without importing the provider plumbing itself."""
    provider = payload.get("provider") or "anthropic"

    def distil(system_prompt, user_message, schema):
        from ideation_generator import _call_claude, _call_ollama
        if provider == "ollama":
            return _call_ollama(payload.get("ollama_base_url", ""),
                                payload.get("ollama_model") or "qwen3:32b",
                                system_prompt, user_message, schema)
        return _call_claude(payload.get("api_key", ""), system_prompt,
                            user_message, schema, 3000)
    return distil


def generate_for_doc(payload):
    comfy_url = (payload.get("comfy_url") or "").strip()
    events_root = payload.get("events_root", "")
    event_type = payload.get("event_type", "")
    doc = payload.get("doc") or {}
    index = int(payload.get("index", 0))
    max_images = int(payload.get("max_images", 6))

    if not comfy_url:
        return {"ok": False, "error": "No ComfyUI URL configured. Set it in Settings."}

    health = comfy_client.check(comfy_url)
    if health["state"] != "ok":
        # Surfaced verbatim: 'wedged' and 'unreachable' need different fixes and
        # a generic "generation failed" would send someone hunting the wrong one.
        return {"ok": False, "error": health["detail"], "state": health["state"]}

    refs = art_reference.style_reference_images(events_root, event_type)
    if not refs:
        log("Warning: no images in StyleReference/images - generating from prose "
            "only, which the render spec (section 8) says is materially weaker.")

    spec = art_reference.load_style_spec(events_root, event_type)
    if not spec:
        log("Warning: no render spec found in StyleReference/ for %s - the "
            "approved render spec will not be in the prompt." % event_type)

    distil = _distiller(payload)
    title = doc.get("title") or event_type
    zones = (doc.get("zones") or [])[:max_images]
    if not zones:
        return {"ok": False, "error": "This variation has no zones to illustrate."}

    # The scene pass is opt-out, not mandatory. It used to run unconditionally,
    # so unticking "Scene overviews" and asking for only recipes still spent four
    # scene renders - and the cost line, which comes from plan_passes, quietly
    # under-reported by exactly those four. Re-rendering scenes to reach the
    # icons also re-rolls art that was already approved.
    requested = list(payload.get("passes") or DEFAULT_PASSES)
    do_scenes = "scene" in requested

    out_dir = os.path.join(events_root, event_type, "Art Requirements",
                           "Drafts", "images", "v%d_%s" % (index + 1, _slug(title)))
    results = []
    for zone in (zones if do_scenes else []):
        zn = zone.get("zone_number", len(results) + 1)
        prompt = art_reference.build_scene_prompt(
            events_root, event_type, zone, doc_title=title,
            art_direction=doc.get("art_direction_guidance"),
            style_spec=spec, distil_fn=distil)
        dest = os.path.join(out_dir, "zone%s_%s.png" % (zn, _slug(zone.get("zone_theme"), 32)))
        # Per-zone references when the type has been seeded with them, so a pad's
        # shape and proportions come from the shipped art rather than from prose.
        zrefs = art_reference.zone_reference_images(events_root, event_type, zn) or refs
        log("rendering zone %s - %s (refs: %s)"
            % (zn, zone.get("zone_theme", ""),
               ", ".join(os.path.basename(r) for r in zrefs) or "none"))
        # A zone may override the render aspect - a 2.12:1 pad has no room to be
        # itself inside a 3:4 frame - and crop_aspect brings the result back to
        # the shipped portrait framing afterwards.
        z_aspect = (zone.get("render_aspect") or payload.get("aspect_ratio") or "Auto")
        res = comfy_client.generate(
            comfy_url, prompt, dest, ref_paths=zrefs, log_fn=log,
            aspect_ratio=z_aspect,
            image_size=payload.get("image_size", "2K"))
        crop = zone.get("crop_aspect") or payload.get("crop_aspect")
        if res.get("ok") and crop:
            try:
                crop_to_aspect(res["path"], crop)
            except Exception as e:
                log("crop failed for zone %s (%s), keeping the uncropped render"
                    % (zn, e.__class__.__name__))
        if res.get("ok"):
            # Stamped onto the zone itself so _render_markdown can embed it and
            # so a later single-image regenerate knows which file to edit.
            zone["image_path"] = res["path"]
            zone["image_prompt"] = prompt
        results.append({
            "zone_number": zn,
            "zone_theme": zone.get("zone_theme", ""),
            "ok": bool(res.get("ok")),
            "path": res.get("path", ""),
            "bytes": res.get("bytes", 0),
            "error": res.get("error", ""),
            "prompt_chars": len(prompt),
            "refs": [os.path.basename(r) for r in zrefs],
        })

    ok_n = sum(1 for r in results if r["ok"])

    # Always review what was just produced. Generation reporting "ok" only means
    # ComfyUI returned bytes - it says nothing about whether the images are
    # blank or whether two zones came back as the same picture. Eyeballing a
    # contact sheet missed exactly that class of fault once already.
    # With the scene pass off there is nothing here to review, and the per-zone
    # maps would key off image_path values left over from an earlier run.
    scened = [z for z in zones if z.get("image_path")] if do_scenes else []
    review = check_generated(
        [r["path"] for r in results if r.get("ok")],
        {os.path.basename(z["image_path"]): zone_tone(z) for z in scened},
        event_type=event_type,
        signage={os.path.basename(z["image_path"]): zone_signage(z) for z in scened},
        vision=_vision_cfg(payload),
        counts={os.path.basename(z["image_path"]): zone_building_count(z)
                for z in scened})
    for f in review["findings"]:
        log("REVIEW %s: %s - %s" % (f["level"].upper(), f["frame"], f["message"]))
    if review["ok"] and ok_n:
        log("REVIEW PASS: %d image(s) checked, none blank, all distinct" % ok_n)

    extra = _generate_extra_passes(payload, doc, zones, out_dir, comfy_url, event_type)
    results.extend(extra["images"])
    ok_n += extra["generated"]

    return {"ok": ok_n > 0, "images": results, "out_dir": out_dir,
            "generated": ok_n,
            "requested": (len(zones) if do_scenes else 0) + extra["requested"],
            "review": review, "extra_reviews": extra["reviews"], "doc": doc,
            "plan": plan_passes(doc, payload.get("passes") or DEFAULT_PASSES),
            "style_refs": [os.path.basename(p) for p in refs]}


def crop_to_aspect(path, aspect, dest=None):
    """Centre-crop a render to a target w:h, in place by default.

    Renders wide and crops to portrait so the model draws a pad at its true
    proportions - in a 3:4 frame it fills instead, which is why a 2.12:1 pad
    came back at 1.39. The cost is real and worth stating: cropping keeps the
    full height and discards width, so whatever the model drew at the sides is
    gone. Pair it with a render aspect only a little wider than the target, or
    the pad is clipped rather than framed.
    """
    from PIL import Image
    try:
        aw, ah = [float(x) for x in str(aspect).split(":")]
        target = aw / ah
    except (ValueError, TypeError):
        return path
    dest = dest or path
    im = Image.open(path)
    w, h = im.size
    if abs((w / float(h)) - target) < 0.01:
        return path
    if w / float(h) > target:            # too wide: trim the sides
        nw = int(round(h * target)); nx = (w - nw) // 2
        box = (nx, 0, nx + nw, h)
    else:                                # too tall: trim top and bottom
        nh = int(round(w / target)); ny = (h - nh) // 2
        box = (0, ny, w, ny + nh)
    im.crop(box).save(dest)
    return dest


def _render_one(comfy_url, prompt, dest, refs, kind, aspect, size, post=None):
    """One render plus whatever post-processing its asset class needs."""
    res = comfy_client.generate(comfy_url, prompt, dest, ref_paths=refs, log_fn=log,
                                aspect_ratio=aspect, image_size=size)
    if res.get("ok") and post:
        try:
            post(dest)
        except Exception as e:
            res = dict(res, ok=False,
                       error="rendered but post-processing failed (%s: %s)"
                             % (e.__class__.__name__, str(e)[:120]))
    return {"kind": kind, "ok": bool(res.get("ok")), "path": res.get("path", dest),
            "bytes": res.get("bytes", 0), "error": res.get("error", "")}


def _icon_post(level, levels, fraction=None, canvas=None):
    """Cut out, clear enclosed backdrop, then apply the geometry icon_sheet owns.

    canvas defaults to the 180px merge-icon canvas; Dessert recipes and
    ingredients ship at 342 and 224, so they pass their own.
    """
    def run(dest):
        import icon_sheet
        import scoop
        scoop.knockout_background(dest)
        scoop.knockout_enclosed(dest)       # holes the border flood cannot reach
        scoop.dehalo(dest)                  # guarded: no-op unless a pale fringe exists
        cv = canvas or icon_sheet.CANVAS
        icon_sheet.normalise(dest, dest, level=level, levels=levels, fraction=fraction,
                             canvas=cv, stroke=icon_sheet.stroke_for(cv))
    return run


def _generate_extra_passes(payload, doc, zones, out_dir, comfy_url, event_type):
    """Building sheets, upgrade stages, chain and producer icons, recipes and
    ingredients.

    Each class has its own aspect, its own post-processing and its own review,
    so they cannot share the scene loop. Buildings and stages keep their
    background - they are illustrations. Icons are cut out - they are assets.
    """
    import asset_prompts
    requested = [p for p in (payload.get("passes") or DEFAULT_PASSES) if p in PASSES and p != "scene"]
    out = {"images": [], "generated": 0, "requested": 0, "reviews": {}}
    if not requested:
        return out

    vision = _vision_cfg(payload)
    producers = doc.get("producers") or []

    def review(paths, signage, label):
        if not paths:
            return
        r = check_generated(paths, tones=None, event_type=event_type, signage=signage)
        out["reviews"][label] = r
        for f in r["findings"]:
            log("REVIEW %s [%s]: %s - %s" % (f["level"].upper(), label, f["frame"], f["message"]))

    # ---- building sheets ----------------------------------------------------
    if "building" in requested:
        made, sign = [], {}
        for z in zones:
            scene_img = z.get("image_path")
            for b in (z.get("buildings") or []):
                dest = os.path.join(out_dir, "buildings", "z%s_b%s_%s.png"
                                    % (z.get("zone_number"), b.get("building_number"),
                                       _slug(b.get("name"), 28)))
                log("rendering building %s - %s" % (b.get("building_number"), b.get("name", "")))
                r = _render_one(comfy_url, asset_prompts.building_prompt(z, b), dest,
                                [scene_img] if scene_img else [], "building", "4:5", "2K")
                out["images"].append(r); out["requested"] += 1
                if r["ok"]:
                    b["image_path"] = dest
                    made.append(dest); sign[os.path.basename(dest)] = b.get("signage") or []
                    out["generated"] += 1
        review(made, sign, "building")

    # ---- upgrade stages -----------------------------------------------------
    if "stage" in requested:
        made, sign, chains = [], {}, []
        for z in zones:
            for b in (z.get("buildings") or []):
                n = int(b.get("max_level") or 1)
                final = b.get("image_path")
                if n < 2 or not final:
                    continue
                frames = []
                for lv in range(1, n):
                    if not asset_prompts.stage_rule(n, lv):
                        continue
                    dest = os.path.join(out_dir, "stages", "z%s_b%s_%s_L%d.png"
                                        % (z.get("zone_number"), b.get("building_number"),
                                           _slug(b.get("name"), 24), lv))
                    log("rendering %s level %d of %d" % (b.get("name", ""), lv, n))
                    r = _render_one(comfy_url, asset_prompts.stage_prompt(z, b, n, lv), dest,
                                    [final], "stage", "4:5", "2K")
                    out["images"].append(r); out["requested"] += 1
                    if r["ok"]:
                        made.append(dest); frames.append(dest)
                        # An unfinished building is unsigned, so [] - an explicit
                        # "no text", which the reviewer enforces strictly.
                        sign[os.path.basename(dest)] = []
                        out["generated"] += 1
                if frames:
                    chains.append((b.get("name", ""), final, frames + [final]))
        review(made, sign, "stage")
        _review_progression(chains, out)

    # ---- chain item icons ---------------------------------------------------
    if "icon" in requested:
        made, sign = [], {}
        for p in producers:
            for ch in (p.get("chains") or []):
                items = ch.get("items") or []
                for lv in range(1, len(items) + 1):
                    dest = os.path.join(out_dir, "icons", "element_%s_%s_%d.png"
                                        % (p.get("producer_number"), ch.get("chain_number"), lv))
                    log("rendering icon %s/%s level %d" % (p.get("producer_number"),
                                                           ch.get("chain_number"), lv))
                    r = _render_one(comfy_url, asset_prompts.chain_item_prompt(ch, lv), dest,
                                    [], "icon", "1:1", "2K",
                                    post=_icon_post(lv, len(items)))
                    out["images"].append(r); out["requested"] += 1
                    if r["ok"]:
                        made.append(dest); sign[os.path.basename(dest)] = []
                        out["generated"] += 1
        review(made, sign, "icon")
        _review_chains(producers, out_dir, out)

    # ---- producer icons -----------------------------------------------------
    if "producer" in requested:
        import icon_sheet
        made, sign = [], {}
        for p in producers:
            for lv in (1, 2):
                dest = os.path.join(out_dir, "icons", "element_machine%s_%d.png"
                                    % (p.get("producer_number"), lv))
                log("rendering producer %s level %d" % (p.get("producer_number"), lv))
                r = _render_one(comfy_url, asset_prompts.producer_prompt(p, lv), dest,
                                [], "producer", "1:1", "2K",
                                post=_icon_post(lv, 2, fraction=icon_sheet.PRODUCER_STEPS[lv - 1]))
                out["images"].append(r); out["requested"] += 1
                if r["ok"]:
                    made.append(dest); sign[os.path.basename(dest)] = []
                    out["generated"] += 1
        review(made, sign, "producer")

    # ---- recipes and ingredients (Dessert) ----------------------------------
    # A flat set, not a ladder: every dish is drawn at the same scale as every
    # other, so there is no size progression to review - only the canvas and
    # fill measured off the shipped Dessert2608 icons.
    if "recipe" in requested or "ingredient" in requested:
        import icon_sheet
        motif = (doc.get("event_motif") or "").strip()
        zone_pal = {z.get("zone_number"): z.get("palette_or_mood", "") for z in zones}

        if "recipe" in requested:
            made, sign = [], {}
            for i, rc in enumerate(doc.get("recipes") or [], 1):
                n = rc.get("recipe_number") or i
                dest = os.path.join(out_dir, "recipes", "recipe_%02d_%s.png"
                                    % (int(n), _slug(rc.get("name"), 24)))
                log("rendering recipe %s - %s" % (n, rc.get("name", "")))
                r = _render_one(comfy_url,
                                asset_prompts.recipe_prompt(rc, motif,
                                                            zone_pal.get(rc.get("zone_number"), "")),
                                dest, [], "recipe", "1:1", "2K",
                                post=_icon_post(1, 1, fraction=icon_sheet.RECIPE_FRACTION,
                                                canvas=icon_sheet.RECIPE_CANVAS))
                out["images"].append(r); out["requested"] += 1
                if r["ok"]:
                    rc["image_path"] = dest
                    made.append(dest); sign[os.path.basename(dest)] = []
                    out["generated"] += 1
            review(made, sign, "recipe")

        if "ingredient" in requested:
            made, sign = [], {}
            for i, ing in enumerate(doc.get("ingredients") or [], 1):
                dest = os.path.join(out_dir, "ingredients", "ingredient_%02d_%s.png"
                                    % (i, _slug(ing.get("name"), 24)))
                log("rendering ingredient %s - %s" % (i, ing.get("name", "")))
                r = _render_one(comfy_url, asset_prompts.ingredient_prompt(ing), dest,
                                [], "ingredient", "1:1", "2K",
                                post=_icon_post(1, 1, fraction=icon_sheet.INGREDIENT_FRACTION,
                                                canvas=icon_sheet.INGREDIENT_CANVAS))
                out["images"].append(r); out["requested"] += 1
                if r["ok"]:
                    ing["image_path"] = dest
                    made.append(dest); sign[os.path.basename(dest)] = []
                    out["generated"] += 1
            review(made, sign, "ingredient")

    # ---- balloon, badge and task icons --------------------------------------
    # Small standalone objects. Same cut-out path as the icons, but no chain and
    # so no size ladder - each is normalised at the top of the scale.
    if "extras" in requested:
        import icon_sheet
        ex = doc.get("event_extras") or {}
        pal = ((doc.get("zones") or [{}])[0]).get("palette_or_mood", "")
        jobs = []
        if (ex.get("hot_air_balloon") or "").strip():
            jobs.append(("balloon", asset_prompts.balloon_prompt(ex["hot_air_balloon"], pal)))
        if (ex.get("achievement_badge") or "").strip():
            jobs.append(("badge", asset_prompts.badge_prompt(ex["achievement_badge"], pal)))
        for i, t in enumerate([t for t in (ex.get("task_icons") or []) if str(t).strip()], 1):
            jobs.append(("task_%02d" % i, asset_prompts.task_icon_prompt(t, pal)))
        made, sign = [], {}
        for name, prompt in jobs:
            dest = os.path.join(out_dir, "icons", "extra_%s.png" % name)
            log("rendering extra: %s" % name)
            r = _render_one(comfy_url, prompt, dest, [], "extras", "1:1", "2K",
                            post=_icon_post(1, 1, fraction=icon_sheet.SIZE_STEPS[-1]))
            out["images"].append(r); out["requested"] += 1
            if r["ok"]:
                made.append(dest); sign[os.path.basename(dest)] = []
                out["generated"] += 1
        review(made, sign, "extras")
    return out


def _review_progression(chains, out):
    """Each building's stages must actually progress - not stall, not restart."""
    try:
        import sequence_review
    except Exception:
        return
    findings = []
    for name, final, frames in chains:
        try:
            r = sequence_review.review(final, frames,
                                       labels=["L%d" % (i + 1) for i in range(len(frames))])
        except Exception:
            continue
        for f in r["findings"]:
            if f["level"] == "error":
                findings.append({"level": "error", "frame": "%s %s" % (name, f["frame"]),
                                 "message": f["message"]})
                log("REVIEW ERROR [stage-progression]: %s %s - %s" % (name, f["frame"], f["message"]))
    out["reviews"]["stage_progression"] = {"ok": not findings, "findings": findings}


def _review_chains(producers, out_dir, out):
    """A chain's icons must grow across the set - Mattel's 图标大小进阶."""
    try:
        import icon_sheet
    except Exception:
        return
    findings = []
    for p in producers:
        for ch in (p.get("chains") or []):
            n = len(ch.get("items") or [])
            paths = [os.path.join(out_dir, "icons", "element_%s_%s_%d.png"
                                  % (p.get("producer_number"), ch.get("chain_number"), lv))
                     for lv in range(1, n + 1)]
            if not paths or not all(os.path.exists(x) for x in paths):
                continue
            r = icon_sheet.check_chain(paths)
            for msg in r["findings"]:
                findings.append({"level": "error",
                                 "frame": "producer %s chain %s" % (p.get("producer_number"),
                                                                    ch.get("chain_number")),
                                 "message": msg})
                log("REVIEW ERROR [icon-progression]: %s" % msg)
    out["reviews"]["icon_progression"] = {"ok": not findings, "findings": findings}


PASSES = ("scene", "building", "stage", "icon", "producer", "recipe", "ingredient", "extras")
DEFAULT_PASSES = ("scene",)     # unchanged behaviour unless the caller asks for more


def plan_passes(doc, passes):
    """What each requested pass would produce, without producing any of it.

    Separate from generation so the app can show the cost before spending it -
    a full Merge event is about 120 renders at half a minute each, and nobody
    should discover that by starting it.
    """
    zones = doc.get("zones") or []
    producers = doc.get("producers") or []
    plan = []
    for name in passes:
        if name == "scene":
            n = len(zones)
        elif name == "building":
            n = sum(len(z.get("buildings") or []) for z in zones)
        elif name == "stage":
            n = sum(max(0, int(b.get("max_level") or 1) - 1)
                    for z in zones for b in (z.get("buildings") or []))
        elif name == "icon":
            n = sum(len(ch.get("items") or []) for p in producers for ch in (p.get("chains") or []))
        elif name == "producer":
            n = len(producers) * 2
        elif name == "recipe":
            n = len(doc.get("recipes") or [])
        elif name == "ingredient":
            n = len(doc.get("ingredients") or [])
        elif name == "extras":
            ex = doc.get("event_extras") or {}
            n = (1 if (ex.get("hot_air_balloon") or "").strip() else 0)                 + (1 if (ex.get("achievement_badge") or "").strip() else 0)                 + len([t for t in (ex.get("task_icons") or []) if str(t).strip()])
        else:
            continue
        plan.append({"pass": name, "count": n})
    return {"passes": plan, "total": sum(p["count"] for p in plan)}


IMAGE_DOC_KIND = "image"        # keeps image lessons separate from document lessons


def _feedback_prompt(zone, feedback, lessons_block, original_prompt=""):
    """A single-hop edit prompt: hold everything, change only what was asked.

    Deliberately NOT a fresh render. A regenerate-from-scratch throws away the
    composition the reviewer already accepted and only objected to one part of;
    the edit path preserves it. This is the same hold-then-change shape that
    the studio's own workflow uses and that survived every test this session.
    """
    lessons = ("\nLESSONS FROM EARLIER IMAGE FEEDBACK - apply all of these too:\n" + lessons_block
               if lessons_block else "")
    # Built by concatenation, not %-formatting: the style block contains a
    # literal "15% deeper", which %-formatting reads as a format specifier.
    context = "Scene context (for reference only - do not re-render it from scratch): zone {}, {}.".format(
        zone.get("zone_number", "?"), zone.get("zone_theme", ""))
    return (
        "Revise this image. Everything in it stays EXACTLY as it is - the same camera, the same "
        "perspective and horizon, the same composition, the same objects in the same positions, the "
        "same lighting and the same palette - except for the change requested below.\n\n"
        "THE REQUESTED CHANGE:\n" + feedback.strip() + "\n" + lessons + "\n\n"
        "Hold the art style exactly: flat toy-like vector 2.5D, softly modelled with smooth gradients, "
        "MATTE, absolutely NO OUTLINES or darker edge around anything, absolutely NO material texture "
        "(no wood grain, no woven or wicker weave, no fibres, no speckle), shading 15% deeper in each "
        "object's own hue, firm contact shadows. Clean candy palette - no olive, khaki, mustard, taupe, "
        "greige or rust on anything built. One single full-bleed illustration, not a grid or a sheet. "
        "No text, letters or numbers. No people unless they are already present.\n\n" + context
    )


def regenerate_images(payload):
    """Apply reviewer feedback to one generated image, or to all of them.

    payload: { comfy_url, events_root, event_type, doc, index, feedback,
               zone_number, apply_to_all?, author_id, author_name }

    Single-image is the default because most notes are local ("the press should
    be teal"). apply_to_all exists for global notes ("everything is too warm"),
    where fixing one image would just make it inconsistent with the rest.
    """
    comfy_url = (payload.get("comfy_url") or "").strip()
    events_root = payload.get("events_root", "")
    event_type = payload.get("event_type", "")
    doc = payload.get("doc") or {}
    feedback = (payload.get("feedback") or "").strip()
    apply_all = bool(payload.get("apply_to_all"))
    target_zone = payload.get("zone_number")

    if not feedback:
        return {"ok": False, "error": "No feedback given."}
    if not comfy_url:
        return {"ok": False, "error": "No ComfyUI URL configured. Set it in Settings."}

    health = comfy_client.check(comfy_url)
    if health["state"] != "ok":
        return {"ok": False, "error": health["detail"], "state": health["state"]}

    zones = doc.get("zones") or []
    targets = [z for z in zones if z.get("image_path")]
    if not apply_all:
        targets = [z for z in targets if str(z.get("zone_number")) == str(target_zone)]
    if not targets:
        return {"ok": False, "error": "No generated image found for that zone. Generate images first."}

    # Image lessons accumulate under their own doc_kind, so a note about art
    # never leaks into the prompt that writes the document text.
    lessons_block = ""
    try:
        lessons_block = store.lessons_prompt_block(events_root, event_type, doc_kind=IMAGE_DOC_KIND)
    except Exception:
        pass

    results = []
    for zone in targets:
        src = zone["image_path"]
        prompt = _feedback_prompt(zone, feedback, lessons_block, zone.get("image_prompt", ""))
        log("revising zone %s - %s" % (zone.get("zone_number"), zone.get("zone_theme", "")))
        res = comfy_client.generate(comfy_url, prompt, src, ref_paths=[src], log_fn=log,
                                    aspect_ratio=payload.get("aspect_ratio", "Auto"),
                                    image_size=payload.get("image_size", "2K"))
        if res.get("ok"):
            zone["image_prompt"] = prompt
        results.append({"zone_number": zone.get("zone_number"),
                        "zone_theme": zone.get("zone_theme", ""),
                        "ok": bool(res.get("ok")), "path": res.get("path", src),
                        "bytes": res.get("bytes", 0), "error": res.get("error", "")})

    review = check_generated(
        [r["path"] for r in results if r["ok"]],
        {os.path.basename(z["image_path"]): zone_tone(z)
         for z in targets if z.get("image_path")},
        event_type=event_type,
        signage={os.path.basename(z["image_path"]): zone_signage(z)
                 for z in targets if z.get("image_path")},
        vision=_vision_cfg(payload),
        counts={os.path.basename(z["image_path"]): zone_building_count(z)
                for z in targets if z.get("image_path")})
    for f in review["findings"]:
        log("REVIEW %s: %s - %s" % (f["level"].upper(), f["frame"], f["message"]))

    # Record the note so it conditions every future image for this event type.
    entry = None
    try:
        entry = store.write_feedback_entry(
            events_root, event_type, IMAGE_DOC_KIND,
            doc.get("title") or event_type,
            "Zone %s image%s" % (target_zone, " (all zones)" if apply_all else ""),
            feedback, feedback, payload.get("author_id", ""), payload.get("author_name", ""))
    except Exception as e:
        log("Warning: could not record image feedback: %s" % e)

    ok_n = sum(1 for r in results if r["ok"])
    return {"ok": ok_n > 0, "images": results, "doc": doc, "review": review,
            "regenerated": ok_n, "applied_to_all": apply_all,
            "entry_id": (entry or {}).get("entry_id", "")}


TONE_RE = re.compile(r"MAIN\s*TONE\s*[:\-]?\s*(COOL|WARM)", re.I)

# Event types whose finished art legitimately carries rendered words. On Sort and
# Sim any text at all is a defect - a caption strip and a numbered key both
# shipped past visual review, which is why the check exists. On Merge signage IS
# the art: the shipped Merge2606 concept art carries MARKET, BAKERY, DEPT. STORE,
# FLOWER SHOP. There the defect is not "text" but "text the brief did not ask
# for", usually a building's design name written across the building.
TEXT_IS_SCENERY = {"Merge"}

SIGN_RE = re.compile(r"sign reads\s+([^\n,;.]+)", re.I)

_NUMBER_WORDS = {"one": 1, "two": 2, "three": 3, "four": 4, "five": 5, "six": 6,
                 "seven": 7, "eight": 8, "nine": 9, "ten": 10, "eleven": 11,
                 "twelve": 12, "thirteen": 13, "fourteen": 14}
COUNT_RE = re.compile(
    r"EXACTLY\s+(\d+|%s)\s+(?:BUILDING|BUILDABLE|OBJECT)" % "|".join(_NUMBER_WORDS), re.I)


def zone_tone(zone):
    """The zone's MAIN TONE: the declared field first, the prompt as fallback.

    The three zone_* helpers below all follow the same rule. The schema now
    carries these declarations, so the reviewer reads them directly instead of
    scraping the prompt it just wrote - but drafts made before the schema change
    have no such fields, and they still have to be reviewable.
    """
    t = (zone.get("main_tone") or "").strip().upper()
    if t in ("COOL", "WARM"):
        return t
    return tone_from_prompt(zone.get("image_prompt", ""))


def zone_signage(zone):
    """Every word authorised anywhere in this zone, or None if nothing says.

    Returns [] - explicitly "no text anywhere" - when the zone HAS buildings and
    not one of them declares a sign. That is a real expectation and the reviewer
    enforces it strictly; None means nobody said, and it only warns. Conflating
    the two silently skipped the check on 14 assets once already.
    """
    blds = zone.get("buildings")
    if blds:
        out = []
        for b in blds:
            out.extend(b.get("signage") or [])
        return out
    from_prompt = signage_from_prompt(zone.get("image_prompt", ""))
    return from_prompt or None


def zone_building_count(zone):
    """How many buildings the brief demands here, or None."""
    blds = zone.get("buildings")
    if blds:
        return len(blds)
    return count_from_prompt(zone.get("image_prompt", ""))


def count_from_prompt(prompt):
    """How many buildings the brief demanded, or None.

    The rule the generator breaks most often is "EXACTLY N buildings, none
    twice", so the reviewer needs N to check against. Read back out of the
    prompt like the tone and the signage - it is stated there and nowhere else.
    """
    m = COUNT_RE.search(prompt or "")
    if not m:
        return None
    tok = m.group(1).lower()
    return int(tok) if tok.isdigit() else _NUMBER_WORDS.get(tok)


def signage_from_prompt(prompt):
    """The exact sign wording the brief authorised, read back out of the prompt.

    Same approach as tone_from_prompt: the prompt is the only place the wording
    is actually stated and it is already in hand at both call sites, so nothing
    new has to be plumbed through the document schema. The convention is
    "sign reads MARKET" - documented in Merge_RENDER_STYLE.md.
    """
    return [m.group(1).strip() for m in SIGN_RE.finditer(prompt or "") if m.group(1).strip()]


def tone_from_prompt(prompt):
    """The MAIN TONE the brief committed to, or "" if it did not commit to one.

    Read back out of the prompt rather than plumbed through the zone data: the
    prompt is the only place the tone is actually stated, and it is already in
    hand at both call sites.
    """
    m = TONE_RE.search(prompt or "")
    return m.group(1).upper() if m else ""


def check_generated(paths, tones=None, event_type=None, signage=None,
                    vision=None, counts=None):
    """Blank + near-duplicate check over a freshly generated set of images.

    event_type routes the text check, which is genuinely opposite between types -
    see TEXT_IS_SCENERY. signage maps a file's basename to the sign wording its
    brief authorised, and is only consulted for types in that set.

    vision, when given, is the provider config for the within-frame duplicate
    check (see duplicate_check) - e.g. {"provider": "ollama", "ollama_base_url":
    ...}. Omitted means the check does not run, and each image says so. counts
    maps a basename to the number of buildings its brief demanded.

    Zone images are independent renders rather than a build sequence, so the
    sequence invariant does not apply; what does apply is that none may be
    blank and no two may be the same picture (a repeated render means one
    zone's prompt was ignored).
    """
    findings = []
    try:
        import sequence_review as sr
    except Exception as e:
        return {"ok": True, "findings": [], "skipped": "Pillow unavailable (%s)" % e}

    loaded = []
    for p in paths:
        name = os.path.basename(p)
        try:
            im = sr._load(p)
        except Exception as e:
            findings.append({"level": "error", "frame": name,
                             "message": "unreadable image (%s)" % e.__class__.__name__})
            continue
        if sr._stdev(im) < sr.BLANK_STDEV:
            findings.append({"level": "error", "frame": name,
                             "message": "image is flat - a blank render"})
            continue

        # Rendered text is the one defect that makes an asset unusable outright,
        # and prose bans do not reliably hold: a labelled list came back as a
        # caption strip, a numbered list as circled callouts plus a key. Both
        # got past visual review. What counts as the defect is per type, though -
        # see TEXT_IS_SCENERY.
        try:
            import text_detect
            if (event_type or "") in TEXT_IS_SCENERY:
                # None and [] mean different things and must not be conflated.
                # None  - the brief said nothing about signage, so there is no
                #         expectation to judge against: warn, do not fail.
                # []    - the brief declared this asset carries NO signage, which
                #         IS an expectation: any rendered word is a defect.
                # A building sheet for a structure with no sign passes [], and
                # treating that as "undeclared" turned every one of them into a
                # noise warning while silently not checking them.
                allowed = (signage or {}).get(name)
                if allowed is not None:
                    s = text_detect.check_signage(p, allowed)
                    if s["ok"] is False:
                        findings.append({"level": "error", "frame": name,
                                         "message": s["detail"]})
                    elif s.get("near_misses"):
                        # One character out: OCR misread, or genuinely misspelled
                        # art. Ambiguous, so it wants eyes, not a hard failure.
                        findings.append({"level": "warn", "frame": name,
                                         "message": s["detail"]})
                    elif s["ok"] is None:
                        # OCR missing: could not check. Never silently a pass.
                        findings.append({"level": "warn", "frame": name,
                                         "message": s["detail"]})
                else:
                    findings.append({"level": "warn", "frame": name,
                                     "message": "%s art may carry signage, so a blanket text ban "
                                                "does not apply - but the brief declared no "
                                                "wording, so rendered text was NOT checked. Name "
                                                "each sign in the prompt as 'sign reads MARKET' "
                                                "to enable the check."
                                                % (event_type or "This")})
            else:
                t = text_detect.detect_text(p)
                if t["has_text"]:
                    findings.append({"level": "error", "frame": name,
                                     "message": t["detail"]})
        except Exception as e:
            findings.append({"level": "warn", "frame": name,
                             "message": "text check skipped (%s)" % e.__class__.__name__})

        # Mattel: "AI-generated resources will have obvious white edges."
        try:
            import scoop as sc
            # Positive only. A negative score is the contact shadow at the plate
            # rim, which the shipped art has too (-19.4, -13.9, +1.2, 0.0, -19.4
            # across the five approved Sort floors). Warning on abs() flagged
            # correct artwork and invited a "fix" that made it worse.
            halo = sc.edge_halo_score(p)
            if halo > sc.HALO_TOLERANCE:
                findings.append({"level": "warn", "frame": name,
                                 "message": "pale edge fringe of %+.0f lightness against the interior - "
                                            "run scoop.dehalo() before handing this over" % halo})
        except Exception:
            pass

        # Reference bleed: one render came back with an entire attached reference
        # room composited into it. Comparing against the reference pixel-wise
        # does not find this - see bleed_check's docstring. What does find it is
        # the reference dragging its palette in, breaking the tone the brief set.
        tone = (tones or {}).get(name, "")
        if tone:
            try:
                import bleed_check
                b = bleed_check.check(p, tone)
                if b["bleed"]:
                    findings.append({"level": "error", "frame": name,
                                     "message": b["detail"]})
            except Exception as e:
                findings.append({"level": "warn", "frame": name,
                                 "message": "bleed check skipped (%s)" % e.__class__.__name__})

        # The same building drawn twice INSIDE this one frame. The pairwise test
        # below cannot see it - that compares whole frames against each other.
        # Opt-in: costs one vision call per image, and without a provider the
        # finding says so rather than staying silent.
        if vision:
            try:
                import duplicate_check
                d = duplicate_check.check(p, expected_count=(counts or {}).get(name),
                                          **vision)
                if d["duplicate_found"]:
                    findings.append({"level": "error", "frame": name, "message": d["detail"]})
                elif not d["ran"]:
                    findings.append({"level": "warn", "frame": name, "message": d["detail"]})
                elif d.get("building_count") is not None and (counts or {}).get(name) \
                        and int(d["building_count"]) != int(counts[name]):
                    findings.append({"level": "warn", "frame": name, "message": d["detail"]})
            except Exception as e:
                findings.append({"level": "warn", "frame": name,
                                 "message": "duplicate check skipped (%s)" % e.__class__.__name__})

        loaded.append((name, im))

    for i in range(len(loaded)):
        for j in range(i + 1, len(loaded)):
            (na, a), (nb, b) = loaded[i], loaded[j]
            if a.size != b.size:
                continue
            changed, _ = sr._diff_mask(a, b)
            if changed < 0.02:
                findings.append({"level": "error", "frame": nb,
                                 "message": "only %.1f%% different from %s - the same "
                                            "picture was rendered twice" % (changed * 100, na)})
    return {"ok": not any(f["level"] == "error" for f in findings), "findings": findings}


def _read_payload():
    # pybackend_cli rewrites sys.argv to [script_name, payload_path], so the
    # payload is argv[1] here - same contract as every other backend script.
    # utf-8-sig because the Electron side can write a BOM.
    if len(sys.argv) < 2:
        raise RuntimeError("No payload file given")
    with open(sys.argv[1], "r", encoding="utf-8-sig") as f:
        return json.load(f)


def regenerate_main():
    try:
        print(json.dumps(regenerate_images(_read_payload())))
    except Exception as e:
        print(json.dumps({"ok": False, "error": "%s: %s" % (e.__class__.__name__, e)}))


def images_main():
    try:
        print(json.dumps(generate_for_doc(_read_payload())))
    except Exception as e:
        print(json.dumps({"ok": False, "error": "%s: %s" % (e.__class__.__name__, e)}))


def plan_main():
    """What the requested passes would render, without rendering anything.

    Deliberately touches neither ComfyUI nor the network: a full Merge event is
    about a hundred renders at half a minute each, and that is not a number
    anyone should discover by starting the job.
    """
    try:
        payload = _read_payload()
        doc = payload.get("doc") or {}
        plan = plan_passes(doc, payload.get("passes") or PASSES)
        plan["ok"] = True
        plan["estimate_seconds"] = plan["total"] * 32      # measured average per render
        print(json.dumps(plan))
    except Exception as e:
        print(json.dumps({"ok": False, "error": "%s: %s" % (e.__class__.__name__, e)}))


DEFAULT_COMFY_URL = "https://skipbo.comfyui.liquidnitro.games"


def check_main():
    """Accepts a payload file, a bare URL, or nothing at all.

    The Electron side passes a payload file like every other command, but this
    is also the command a human runs by hand to ask "is the box up yet?", and
    making them author JSON first is a good way to get a FileNotFoundError
    instead of an answer. All three of these work:

        pybackend_cli.py comfy_check
        pybackend_cli.py comfy_check https://skipbo.comfyui.liquidnitro.games
        pybackend_cli.py comfy_check payload.json
    """
    try:
        arg = sys.argv[1].strip() if len(sys.argv) > 1 else ""
        if not arg:
            url = DEFAULT_COMFY_URL
        elif arg.startswith(("http://", "https://")):
            url = arg
        else:
            with open(arg, "r", encoding="utf-8-sig") as f:
                url = (json.load(f).get("comfy_url") or DEFAULT_COMFY_URL).strip()
        print(json.dumps(comfy_client.check(url)))
    except Exception as e:
        print(json.dumps({"state": "unreachable", "reachable": False,
                          "detail": "%s: %s" % (e.__class__.__name__, e)}))
