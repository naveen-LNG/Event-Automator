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
import shutil
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


def _aspect_ratio_value(spec):
    """"9:16" -> 0.5625. Returns None for "Auto" or anything unparseable."""
    try:
        w, h = str(spec).split(":")
        return float(w) / float(h)
    except Exception:
        return None


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
        # References are resolved BEFORE the prompt, because the prompt now has to
        # say what is attached. This used to run after, which is how the scene
        # pass ended up attaching two shipped scenes and never mentioning them -
        # the model took the finish from them and invented the geometry.
        # Per-zone references when the type has been seeded with them, so a pad's
        # shape and proportions come from the shipped art rather than from prose.
        zrefs = art_reference.zone_reference_images(events_root, event_type, zn) or refs
        prompt = art_reference.build_scene_prompt(
            events_root, event_type, zone, doc_title=title,
            art_direction=doc.get("art_direction_guidance"),
            style_spec=spec, distil_fn=distil, reference_images=zrefs)
        dest = os.path.join(out_dir, "zone%s_%s.png" % (zn, _slug(zone.get("zone_theme"), 32)))
        log("rendering zone %s - %s (refs: %s)"
            % (zn, zone.get("zone_theme", ""),
               ", ".join(os.path.basename(r) for r in zrefs) or "none"))
        # A zone may override the render aspect - a 2.12:1 pad has no room to be
        # itself inside a 3:4 frame - and crop_aspect brings the result back to
        # the shipped portrait framing afterwards.
        # A zone may override, then the caller, then the type's MEASURED aspect,
        # and only then Auto. Auto used to be third and returned a near-square
        # frame for a 9:16 scene, which is why this order matters.
        z_aspect = (zone.get("render_aspect") or payload.get("aspect_ratio")
                    or art_reference.scene_aspect(event_type) or "Auto")
        res = comfy_client.generate(
            comfy_url, prompt, dest, ref_paths=zrefs, log_fn=log,
            aspect_ratio=z_aspect,
            image_size=payload.get("image_size", "2K"))
        # Verify the shape that came back. The renderer is asked for an aspect,
        # not guaranteed one, and a wrong-shaped frame is the failure that looks
        # most like success - it is a complete, pretty, unusable image.
        want = _aspect_ratio_value(z_aspect)
        if want and res.get("ok") and os.path.isfile(dest):
            try:
                from PIL import Image
                with Image.open(dest) as _im:
                    got = _im.width / float(_im.height)
                if abs(got - want) / want > 0.05:
                    msg = ("zone %s rendered at ratio %.3f but %s is %.3f - the frame "
                           "shape is wrong, not just the content"
                           % (zn, got, z_aspect, want))
                    log("WARNING " + msg)
                    out.setdefault("warnings", []).append(msg)
            except Exception:
                pass

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
            "plan": plan_passes(doc, payload.get("passes") or DEFAULT_PASSES, event_type),
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

    # ---- scene evolution sequences (Blocks) ----------------------------------
    # Structure and method are Blocks_scene_progression_rules.md, which was
    # written after a chained-forward run drifted its palette into the banned
    # taupe band. Everything below encodes a rule that document paid for:
    #   * the built base is rendered ONCE and every earlier frame is a single
    #     CUMULATIVE hop from it, so nothing accumulates
    #   * a bulk removal silently leaves large objects standing, so the hero and
    #     the first structure each get their own blunt single-change hop
    #   * the derelict frame is a different WEATHER AND LIGHT state, not a
    #     desaturated copy
    #   * characters are one ADDITIVE hop from the base, final frame only
    #   * every set is checked by sequence_review, never by eye
    if "progression" in requested and (event_type or "") in PROGRESSION_TYPES:
        # Derived here rather than passed in: this helper only receives the
        # payload, and the sequence needs the same spec and references the scene
        # pass uses. distil_fn is None - the visual-language distillation needs a
        # provider and a sequence must not depend on one being reachable.
        events_root = payload.get("events_root", "")
        title = doc.get("title") or event_type
        spec = art_reference.load_style_spec(events_root, event_type)
        refs = art_reference.style_reference_images(events_root, event_type)
        # The build-up frames must be the SAME SHAPE as the finished scene they
        # build up to, and the scene's shape is the type's measured one. The old
        # default here was "4:5", and since nothing passes aspect_ratio every
        # frame rendered near-square while its own scene rendered 9:16.
        prog_aspect = (payload.get("aspect_ratio")
                       or art_reference.scene_aspect(event_type) or "4:5")
        distil = None
        made_any = False
        for z in zones:
            steps = [s for s in (z.get("construction_steps") or []) if str(s).strip()]
            if len(steps) < 2:
                continue
            made_any = True
            zn = z.get("zone_number")
            zdir = os.path.join(out_dir, "scene%s_%s" % (zn, _slug(z.get("zone_theme"), 24)))
            frames, labels = [], []

            # reference_images so the prompt names what is attached and assigns
            # it the geometry - without it the attachment is an unlabelled hint
            # and the horizon gets invented, which is what happened on the scene
            # pass for three rounds.
            base_prompt = art_reference.build_scene_prompt(
                events_root, event_type,
                dict(z, reference_notes="%s\n\n%s" % (
                    z.get("reference_notes", ""), asset_prompts.built_scene(steps))),
                doc_title=title, art_direction=doc.get("art_direction_guidance"),
                style_spec=spec, distil_fn=distil,
                reference_images=art_reference.zone_reference_images(
                    events_root, event_type, z.get("zone_number")) or refs)
            base = os.path.join(zdir, "%02d_F%d_complete.png" % (len(steps), len(steps) + 1))
            # If this zone already has an ACCEPTED scene image, that is the
            # completed frame - copy it rather than rendering a sibling of it.
            # Rendering a second "finished scene" produced a strip whose last
            # frame did not match the hero image above it in the document.
            accepted = z.get("image_path") or ""
            if accepted and os.path.isfile(accepted):
                if not os.path.isdir(zdir):
                    os.makedirs(zdir)
                shutil.copyfile(accepted, base)
                log("scene %s: using the accepted scene as the completed base" % zn)
            else:
                log("scene %s: rendering the completed base" % zn)
                zrefs = art_reference.zone_reference_images(
                    events_root, event_type, zn) or refs
                r = _render_one(comfy_url, base_prompt, base, zrefs, "progression",
                                prog_aspect, "2K")
                out["images"].append(r); out["requested"] += 1
                if not r["ok"]:
                    log("scene %s: base failed, skipping the rest of this sequence" % zn)
                    continue
                out["generated"] += 1

            # steps 1..N-1, each the base minus everything not yet built.
            # Step 1 (the hero alone) and step 2 are cut from their OWN
            # successor rather than in bulk, because a long removal list leaves
            # the biggest object standing.
            prev = base
            per_step = {}
            # Steps 1 and 2 always get their own blunt hop - the hero is step 1
            # and the largest structure is usually step 2, and a bulk list leaves
            # both standing. blunt_steps names any OTHER step that is a large
            # structure or a flat ground element: the barn's laid dance-floor sat
            # at step 5, survived the bulk removal, and cost five renders to
            # repair because the fix cascades down the chain.
            blunt = set()
            for x in (z.get("blunt_steps") or []):
                try:
                    blunt.add(int(x))
                except (TypeError, ValueError):
                    log("scene %s: ignoring non-numeric blunt_steps entry %r" % (zn, x))
            # Everything at or below the highest blunt step must CHAIN from its
            # successor, not be cut from the base. Cutting from the base means
            # one bulk list containing the difficult object, and it survives
            # again - which is exactly what happened when blunt_steps=[5] fixed
            # frame 5 and frame 4 then re-introduced the dance floor. Chaining
            # works because the source frame already lacks it.
            chain_below = max([2] + [s - 1 for s in blunt])
            for k in range(len(steps) - 1, 0, -1):
                absent = steps[k:]
                dest = os.path.join(zdir, "%02d_F%d.png" % (k, k + 1))
                if k <= chain_below:             # dedicated single-change hop
                    src, instr = prev, asset_prompts.remove_one(steps[k])
                else:
                    src, instr = base, asset_prompts.remove_many(absent)
                log("scene %s: frame %d of %d (%s)"
                    % (zn, k + 1, len(steps) + 2,
                       "one change" if k <= chain_below else "cumulative"))
                r = _render_one(comfy_url, instr, dest, [src], "progression",
                                prog_aspect, "2K")
                out["images"].append(r); out["requested"] += 1
                if r["ok"]:
                    out["generated"] += 1; per_step[k] = dest; prev = dest

            empty = os.path.join(zdir, "00_F1_empty.png")
            hero_src = per_step.get(1, base)
            log("scene %s: the restored-but-empty frame" % zn)
            r = _render_one(comfy_url, asset_prompts.remove_one(steps[0]), empty,
                            [hero_src], "progression", prog_aspect, "2K")
            out["images"].append(r); out["requested"] += 1
            if r["ok"]:
                out["generated"] += 1

            derelict = os.path.join(zdir, "00_F0_derelict.png")
            log("scene %s: the derelict frame" % zn)
            r = _render_one(comfy_url, asset_prompts.derelict_prompt(z), derelict,
                            [empty if os.path.exists(empty) else base], "progression",
                            prog_aspect, "2K")
            out["images"].append(r); out["requested"] += 1
            if r["ok"]:
                out["generated"] += 1

            # Characters arrive ONE BY ONE across the last two frames, so the
            # population reads as arriving rather than appearing at once.
            chars = [c for c in (doc.get("characters") or []) if c.get("name")][:2]
            prev_frame = base
            for i, c in enumerate(chars):
                cf = os.path.join(zdir, "%02d_F%d_characters%d.png"
                                  % (len(steps) + 1 + i, len(steps) + 2 + i, i + 1))
                log("scene %s: character %d of %d arrives" % (zn, i + 1, len(chars)))
                r = _render_one(comfy_url,
                                asset_prompts.add_characters(
                                    chars[:i + 1], arriving=[c],
                                    scale=(z.get("character_scale") or "beside")),
                                cf, [prev_frame], "progression",
                                prog_aspect, "2K")
                out["images"].append(r); out["requested"] += 1
                if r["ok"]:
                    out["generated"] += 1; prev_frame = cf

            # build order for the review: empty, then each step, then the base
            ordered = [empty] + [per_step[k] for k in sorted(per_step)] + [base]
            labels = ["empty"] + ["+ %s" % steps[k][:24] for k in sorted(per_step)] + ["complete"]
            frames = [p for p in ordered if os.path.exists(p)]
            _review_progression_set(frames, base, derelict, labels, zn, out)

            # Record the sequence ON THE ZONE so the document can show it. The
            # frames were being written to disk and then forgotten: the generator
            # printed the step LIST and none of the step ART.
            #
            # The derelict frame leads here although the review excludes it - the
            # review needs diff-from-base to shrink along the build order and the
            # derelict state breaks that, but an artist needs to see the start.
            shown = ([(derelict, "F0 derelict")] if os.path.exists(derelict) else [])
            shown += [(p, l) for p, l in zip(ordered, labels) if os.path.exists(p)]
            # ...then the characters arriving, which the review set leaves out on
            # purpose (they are added AFTER the scene is complete, so they break
            # the shrinking-diff invariant it checks) but which the artist still
            # has to draw. The document is not the review.
            for i, c in enumerate(chars):
                cf = os.path.join(zdir, "%02d_F%d_characters%d.png"
                                  % (len(steps) + 1 + i, len(steps) + 2 + i, i + 1))
                if os.path.exists(cf):
                    shown.append((cf, "+ %s" % (c.get("name") or "character")))
            z["progression_frames"] = [{"label": l, "path": p} for p, l in shown]
            sheet = progression_sheet([p for p, _ in shown], [l for _, l in shown],
                                      os.path.join(zdir, "progression_sheet.png"))
            if sheet:
                z["progression_sheet"] = sheet
                log("scene %s: progression strip written (%d frames)" % (zn, len(shown)))
        if not made_any:
            log("no zone carries construction_steps, so there is no sequence to build")

    # ---- goods icons (Sort) --------------------------------------------------
    # Per ZONE, not per document: each floor carries its own themed catalogue,
    # and Sort's loop is sorting these onto shelves. Portrait canvas, because the
    # shipped sets are 1:2 and squaring a can or a baguette wastes half the frame.
    if "goods" in requested:
        import icon_sheet
        made, sign = [], {}
        for z in zones:
            zn = z.get("zone_number")
            goods = z.get("goods") or []
            names = [g.get("name", "") for g in goods]
            zone_made = []
            for i, g in enumerate(goods, 1):
                dest = os.path.join(out_dir, "goods", "z%s_good_%02d_%s.png"
                                    % (zn, i, _slug(g.get("name"), 22)))
                log("rendering goods %s/%s - %s" % (zn, i, g.get("name", "")))
                # The siblings are the rest of THIS floor's catalogue. Without
                # them each icon is drawn blind to the set it belongs to and the
                # generator settles on one safe product blob for all of them.
                sibs = [n for j, n in enumerate(names, 1) if j != i and n]
                r = _render_one(comfy_url,
                                asset_prompts.goods_prompt(g, z.get("palette_or_mood", ""),
                                                           siblings=sibs),
                                # 9:16 (0.563) is the closest portrait the node
                                # offers to the shipped 1:2; 1:2 itself is not in
                                # its list and is refused outright.
                                dest, [], "goods", "9:16", "2K",
                                # One fill for every good, matching the shipped
                                # 0.86 wide / 0.90 tall. Sizing goods down
                                # individually is what left half a frame empty.
                                post=_icon_post(1, 1,
                                                fraction=asset_prompts.goods_fraction(
                                                    g, icon_sheet.GOODS_FRACTION),
                                                canvas=icon_sheet.GOODS_CANVAS))
                out["images"].append(r); out["requested"] += 1
                if r["ok"]:
                    g["image_path"] = dest
                    made.append(dest); sign[os.path.basename(dest)] = []
                    zone_made.append(dest)
                    out["generated"] += 1

            # Colour is the third axis the player sorts on, after kind and shape,
            # and it is checked PER FLOOR because that is where a shelf is seen.
            # A floor whose fifteen goods share one palette is a floor of one
            # item repeated, however different their outlines are.
            for deg, a, b in icon_sheet.hue_collisions(zone_made):
                msg = ("zone %s: '%s' and '%s' are the same colour (%d degrees apart) "
                       "- goods on a floor must differ in colour as well as shape"
                       % (zn, os.path.basename(a)[8:-4], os.path.basename(b)[8:-4], deg))
                log("WARNING " + msg)
                out.setdefault("warnings", []).append(msg)

        # Does each good actually sit in its frame the way the shipped ones do?
        # Reported, never corrected: the only way to fix an object that is too
        # wide for a 1:2 frame is to draw a different object or stand it on its
        # end. Scaling it up just makes it wider and softer.
        for f in made:
            probs, fit = icon_sheet.frame_problems(
                f, asset_prompts.GOODS_ASPECT_RANGE, asset_prompts.GOODS_MIN_AREA)
            for why in probs:
                msg = "%s %s" % (os.path.basename(f)[:-4], why)
                log("WARNING " + msg)
                out.setdefault("warnings", []).append(msg)

        # Checked across ALL the goods, not floor by floor. Checking each floor
        # in isolation is what let a catalogue ship where every floor was
        # internally fine and all five were the same fifteen shapes - the repeat
        # ran down the columns, where a per-floor check cannot look.
        for iou, a, b in icon_sheet.silhouette_collisions(made):
            msg = ("'%s' and '%s' share a silhouette (%.2f) - the player cannot sort "
                   "apart two goods that look the same"
                   % (os.path.basename(a)[:-4], os.path.basename(b)[:-4], iou))
            log("WARNING " + msg)
            out.setdefault("warnings", []).append(msg)
        review(made, sign, "goods")

    # ---- construction materials (Blocks) ------------------------------------
    # Blocks' mechanic-specific class, the way recipes are Dessert's: each
    # construction step consumes 2-3 themed material icons, so they are a set.
    if "material" in requested:
        import icon_sheet
        pal = ((doc.get("zones") or [{}])[0]).get("palette_or_mood", "")
        made, sign = [], {}
        for i, m in enumerate(doc.get("construction_materials") or [], 1):
            dest = os.path.join(out_dir, "materials", "material_%02d_%s.png"
                                % (i, _slug(m.get("name"), 24)))
            log("rendering material %s - %s" % (i, m.get("name", "")))
            r = _render_one(comfy_url, asset_prompts.material_prompt(m, pal), dest,
                            [], "material", "1:1", "2K",
                            post=_icon_post(1, 1, fraction=icon_sheet.SIZE_STEPS[-1]))
            out["images"].append(r); out["requested"] += 1
            if r["ok"]:
                m["image_path"] = dest
                made.append(dest); sign[os.path.basename(dest)] = []
                out["generated"] += 1
        review(made, sign, "material")

    # ---- UI panels -----------------------------------------------------------
    # A re-skin, not an invention: the previous panel's own art goes in as the
    # reference so the layout, the proportions and the part list survive, and
    # the prompt changes only the theme over the top of them.
    if "panel" in requested:
        made, sign = [], {}
        theme = ((doc.get("gameplay_overview") or {}).get("theme") or "").strip()
        adg = doc.get("art_direction_guidance") or {}
        lean = "; ".join(adg.get("lean_into") or [])
        pal = ((doc.get("zones") or [{}])[0]).get("palette_or_mood", "")
        panel_notes = {(e.get("panel") or "").strip().lower(): (e.get("note") or "")
                       for e in (doc.get("ui_panels") or [])}
        panels = _panels_with_art(event_type, events_root)
        doc.setdefault("panel_images", {})
        for name, asset, sheet in panels:
            note = panel_notes.get(name.lower()) or panel_notes.get(asset.lower()) or ""
            prompt = (
                "The attached image is a contact sheet of every texture in the "
                "shipped %s UI panel '%s'. RE-SKIN this exact set of parts for the "
                "event theme '%s'. Keep every part, its shape, its proportion and "
                "its role - a button stays that button, a frame stays that frame, "
                "a badge stays that badge. Change ONLY the surface: material, "
                "colour and motif. Palette: %s. Art direction: %s.%s "
                "Lay the re-skinned parts out on the same grid as the reference, "
                "on a plain mid-grey background, with no text, no lettering and no "
                "numbers anywhere."
                % (event_type, name, theme or "the event theme", pal or "as the scenes",
                   lean or "match the scenes", (" " + note) if note else ""))
            dest = os.path.join(out_dir, "panels", "panel_%s.png" % _slug(asset, 40))
            log("rendering panel %s (ref: %s)" % (name, os.path.basename(sheet)))
            r = _render_one(comfy_url, prompt, dest, [sheet], "panel", "Auto", "2K")
            out["images"].append(r); out["requested"] += 1
            if r["ok"]:
                doc["panel_images"][asset] = dest
                made.append(dest); sign[os.path.basename(dest)] = []
                out["generated"] += 1
        review(made, sign, "panel")

    # ---- characters ----------------------------------------------------------
    # Cut out but NOT normalised: icon_sheet forces a square canvas, and a
    # character ships as a tall portrait sprite (Dessert guests are 320x512).
    # Squaring one would letterbox it and throw away the proportion that the
    # style guide spends six sections defining.
    if "character" in requested:
        pal = ((doc.get("zones") or [{}])[0]).get("palette_or_mood", "")

        def cutout(dest):
            import scoop
            scoop.knockout_background(dest)
            scoop.knockout_enclosed(dest)
            scoop.dehalo(dest)

        made, sign = [], {}
        for i, c in enumerate(doc.get("characters") or [], 1):
            dest = os.path.join(out_dir, "characters", "character_%02d_%s.png"
                                % (i, _slug(c.get("name"), 24)))
            log("rendering character %s - %s (%s)"
                % (i, c.get("name", ""), c.get("base_character", "new")))
            r = _render_one(comfy_url, asset_prompts.character_prompt(c, pal), dest,
                            [], "character", "2:3", "2K", post=cutout)
            out["images"].append(r); out["requested"] += 1
            if r["ok"]:
                c["image_path"] = dest
                made.append(dest); sign[os.path.basename(dest)] = []
                out["generated"] += 1
        review(made, sign, "character")

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
        # Recorded on the doc, keyed the way the icon table asks for them. The
        # materials, characters and panels all wrote their paths back; the extras
        # were the one pass that rendered six icons and told nobody where they
        # went, so the document's icon rows stayed on "_add reference_".
        doc.setdefault("extra_images", {})
        for name, prompt in jobs:
            dest = os.path.join(out_dir, "icons", "extra_%s.png" % name)
            log("rendering extra: %s" % name)
            r = _render_one(comfy_url, prompt, dest, [], "extras", "1:1", "2K",
                            post=_icon_post(1, 1, fraction=icon_sheet.SIZE_STEPS[-1]))
            out["images"].append(r); out["requested"] += 1
            if r["ok"]:
                doc["extra_images"][name] = dest
                made.append(dest); sign[os.path.basename(dest)] = []
                out["generated"] += 1

        # The task icons also go in as ONE strip. Four separate images in a single
        # table cell came out at 260px each in the .docx - 1040px inside a column
        # a third that wide - and they are judged as a set anyway: four icons that
        # each read well while sharing no visual family is the failure this is for.
        tasks = [(k, doc["extra_images"][k]) for k in sorted(doc["extra_images"])
                 if k.startswith("task_")]
        if len(tasks) > 1:
            names = [t for t in (ex.get("task_icons") or []) if str(t).strip()]
            strip = progression_sheet(
                [p for _k, p in tasks],
                [str(names[i])[:22] if i < len(names) else k
                 for i, (k, _p) in enumerate(tasks)],
                os.path.join(out_dir, "icons", "task_icons_set.png"), tile_h=170)
            if strip:
                doc["extra_images"]["task_sheet"] = strip
                log("task icons composed into one strip (%d icons)" % len(tasks))
        review(made, sign, "extras")
    return out


PROGRESSION_TILE_H = 300        # px per frame in the strip
PROGRESSION_PER_ROW = 6         # wraps to a second row beyond this
PROGRESSION_LABEL_H = 34


def _progression_font(size=15):
    """A real font if Windows has one, otherwise PIL's bitmap default.

    Wrapped because a missing font must degrade the labels, never lose the sheet:
    the frames are the deliverable and the captions are an aid to reading them.
    """
    from PIL import ImageFont
    for name in ("arial.ttf", "segoeui.ttf", "calibri.ttf"):
        try:
            return ImageFont.truetype("C:/Windows/Fonts/" + name, size)
        except Exception:
            continue
    try:
        return ImageFont.load_default()
    except Exception:
        return None


def _fit_label(draw, text, font, max_w):
    """Trim text to fit max_w PIXELS, not to a character count.

    Cutting at 30 characters is what made the first strip's captions overlap:
    "+ a tall TIERED FRUIT SHELF" is 27 characters and far wider than the tile it
    sat under. Falls back to a crude character estimate on any PIL without
    textlength, because a slightly short caption beats a crash.
    """
    try:
        if draw.textlength(text, font=font) <= max_w:
            return text
        while text and draw.textlength(text + "\u2026", font=font) > max_w:
            text = text[:-1]
        return text + "\u2026" if text else ""
    except Exception:
        return text[:max(4, int(max_w // 7))]


def progression_sheet(frames, labels, dest, tile_h=PROGRESSION_TILE_H):
    """One labelled strip showing a scene's build-up in order. Path, or "".

    Returns "" rather than raising: a document with a scene image and no strip is
    worth shipping, and a pass that dies here would lose the 58 frames it just
    spent an hour rendering.
    """
    try:
        import os
        from PIL import Image, ImageDraw
        live = [(l, p) for l, p in zip(labels, frames) if p and os.path.isfile(p)]
        if len(live) < 2:
            return ""
        tiles = []
        for label, p in live:
            with Image.open(p) as im:
                im = im.convert("RGB")
                w = max(1, int(round(im.width * (tile_h / float(im.height)))))
                tiles.append((label, im.resize((w, tile_h), Image.LANCZOS)))

        per_row = min(PROGRESSION_PER_ROW, len(tiles))
        rows = [tiles[i:i + per_row] for i in range(0, len(tiles), per_row)]
        gap, pad = 10, 14
        row_w = [sum(t[1].width for t in r) + gap * (len(r) - 1) for r in rows]
        cell_h = tile_h + PROGRESSION_LABEL_H
        sheet_w = max(row_w) + pad * 2
        sheet_h = len(rows) * cell_h + gap * (len(rows) - 1) + pad * 2

        sheet = Image.new("RGB", (sheet_w, sheet_h), (244, 243, 240))
        draw = ImageDraw.Draw(sheet)
        font = _progression_font()
        y = pad
        for r in rows:
            x = pad
            for label, im in r:
                sheet.paste(im, (x, y))
                draw.rectangle([x, y, x + im.width - 1, y + tile_h - 1],
                               outline=(205, 202, 196))
                if font is not None:
                    text = _fit_label(draw, str(label), font, im.width - 6)
                    draw.text((x + 3, y + tile_h + 8), text, fill=(60, 58, 55), font=font)
                x += im.width + gap
            y += cell_h + gap

        d = os.path.dirname(os.path.abspath(dest))
        if d and not os.path.isdir(d):
            os.makedirs(d)
        sheet.save(dest)
        return dest
    except Exception:
        return ""


def _review_progression_set(frames, base, derelict, labels, zn, out):
    """Check one scene sequence with sequence_review - never by eye.

    The rules document is blunt about why: visual inspection missed a leaked
    stall across three frames of a shipped set, and it would have reached the
    art team. The subtractive method gives the invariant this exploits - each
    frame is the base minus what is not yet built, so diff-from-base must shrink
    monotonically along the build order.
    """
    key = "progression_scene%s" % zn
    if len(frames) < 2:
        out["reviews"][key] = {"ok": False, "findings": [
            {"level": "error", "frame": "-", "message": "too few frames rendered to review"}]}
        return
    try:
        import sequence_review
        r = sequence_review.review(base, frames,
                                   derelict_path=derelict if os.path.exists(derelict) else None,
                                   labels=labels[:len(frames)])
    except Exception as e:
        out["reviews"][key] = {"ok": False, "findings": [
            {"level": "error", "frame": "-",
             "message": "sequence_review failed to run (%s: %s)" % (e.__class__.__name__, e)}]}
        return
    out["reviews"][key] = r
    for f in r["findings"]:
        log("REVIEW %s [scene %s]: %s - %s"
            % (f["level"].upper(), zn, f.get("frame", ""), f["message"]))
    if r["ok"]:
        log("REVIEW PASS [scene %s]: %d frames, build order holds" % (zn, len(frames)))


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


PASSES = ("scene", "progression", "building", "stage", "icon", "producer", "recipe",
          "ingredient", "goods", "material", "character", "extras", "panel")
DEFAULT_PASSES = ("scene",)     # unchanged behaviour unless the caller asks for more


def plan_passes(doc, passes, event_type=""):
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
        elif name == "progression":
            # per scene: derelict + empty + one frame per build step + characters
            n = 0
            for z in (zones if (event_type or "") in PROGRESSION_TYPES else []):
                st = [x for x in (z.get("construction_steps") or []) if str(x).strip()]
                if len(st) >= 2:
                    # derelict + empty + one per build step + one per arriving
                    # character (they come in one by one over the last frames)
                    n += len(st) + 2 + len((doc.get("characters") or [])[:2])
                    # ...minus the completed base, when the zone already has an
                    # accepted scene image: the pass copies that instead of
                    # rendering a second finished scene. Spelled the same way as
                    # the check in the pass itself, and test_artreq_images pins
                    # the two together - the cost line comes from this number.
                    accepted = z.get("image_path") or ""
                    if accepted and os.path.isfile(accepted):
                        n -= 1
        elif name == "goods":
            n = sum(len(z.get("goods") or []) for z in zones)
        elif name == "material":
            n = len(doc.get("construction_materials") or [])
        elif name == "character":
            n = len(doc.get("characters") or [])
        elif name == "panel":
            # From the table, not the disk: plan_passes is the cost line and must
            # not depend on whether a reference folder happens to be synced yet.
            try:
                import art_requirement_generator as _arg
                n = len([p for p in (_arg.UI_PANELS_BY_TYPE.get(event_type) or []) if p[2]])
            except Exception:
                n = 0
        elif name == "extras":
            ex = doc.get("event_extras") or {}
            n = (1 if (ex.get("hot_air_balloon") or "").strip() else 0)                 + (1 if (ex.get("achievement_badge") or "").strip() else 0)                 + len([t for t in (ex.get("task_icons") or []) if str(t).strip()])
        else:
            continue
        plan.append({"pass": name, "count": n})
    return {"passes": plan, "total": sum(p["count"] for p in plan)}


def _panels_with_art(event_type, events_root):
    """(panel name, shipped asset, reference sheet path) for re-skinnable panels.

    Only panels that ship art of their own. The rest are built from the shared
    chrome, so they change when it does and have nothing of their own to hold.
    """
    try:
        import art_requirement_generator as arg
    except Exception:
        return []
    panels = arg.UI_PANELS_BY_TYPE.get(event_type) or []
    root = _panel_reference_dir(event_type, events_root)
    out = []
    for name, asset, tex in panels:
        if not tex:
            continue
        sheet = os.path.join(root, "%s.jpg" % asset) if root else ""
        if sheet and os.path.isfile(sheet):
            out.append((name, asset, sheet))
    return out


def _panel_reference_dir(event_type, events_root):
    """Where the shipped-panel contact sheets live, beside the documents."""
    if not events_root:
        return ""
    import art_requirement_generator as arg
    return os.path.join(events_root, event_type, "Art Requirements",
                        arg.PANEL_REFERENCE_DIR)


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

# Only Blocks delivers a scene as a BUILD SEQUENCE. Every other type uses
# construction_steps as free-text reveal notes - a pre-schema Sort draft carries
# two of them - so without this gate every legacy draft would suddenly be offered
# a nine-frame sequence it has no use for, and charged for it in the cost line.
PROGRESSION_TYPES = {"Blocks"}

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
        plan = plan_passes(doc, payload.get("passes") or PASSES,
                           payload.get("event_type", ""))
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
