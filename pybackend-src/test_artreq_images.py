"""Tests for image-in-document embedding and image feedback regeneration.

ComfyUI is stubbed, so the whole path is verifiable without spending renders or
depending on the server being up.
"""
import io, json, os, shutil, sys, tempfile

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
HERE = os.path.dirname(os.path.abspath(__file__))
os.chdir(HERE); sys.path.insert(0, HERE)

from PIL import Image
import comfy_client
import artreq_images
import art_requirement_generator as arg

fails = []
def ok(c, m, d=""):
    print(("OK  " if c else "FAIL ") + m + (("\n     " + d) if (d and not c) else ""))
    if not c: fails.append(m)

tmp = tempfile.mkdtemp()
try:
    events = os.path.join(tmp, "ideationData", "Events")
    draft_dir = os.path.join(events, "Merge", "Art Requirements", "Drafts")
    img_dir = os.path.join(draft_dir, "images", "v1_test")
    os.makedirs(img_dir, exist_ok=True)

    def make_png(path, colour):
        # Two strongly different halves, so the image is still clearly non-flat
        # after the reviewer's 4x downsample. Sparse single-pixel dots blur away
        # and get (correctly) reported as blank.
        im = Image.new("RGB", (120, 200), colour)
        inv = tuple(255 - c for c in colour)
        for y in range(100, 200):
            for x in range(120):
                im.putpixel((x, y), inv)
        im.save(path)
        return path

    z1 = make_png(os.path.join(img_dir, "zone1.png"), (200, 90, 70))
    z2 = make_png(os.path.join(img_dir, "zone2.png"), (70, 150, 200))

    DOC = {"title": "Merge2610 Harvest", "gameplay_overview": {},
           "zones": [
               {"zone_number": 1, "zone_theme": "Press Yard", "palette_or_mood": "teal",
                "construction_steps": ["a"], "reference_notes": "n", "image_path": z1},
               {"zone_number": 2, "zone_theme": "Barn", "palette_or_mood": "mint",
                "construction_steps": ["b"], "reference_notes": "n", "image_path": z2}],
           "characters": [], "art_direction_guidance": {}, "gameplay_specific_content": "",
           "ui_icon_refresh_notes": "", "notes": ""}

    # 1. images are embedded in the rendered markdown -------------------------
    md = arg._render_markdown(DOC, "Merge", "Draft", "Naveen", md_dir=draft_dir)
    ok("![Zone 1 - Press Yard](images/v1_test/zone1.png)" in md,
       "zone 1 image is embedded with a path relative to the .md", md[:200])
    ok("![Zone 2 - Barn](images/v1_test/zone2.png)" in md,
       "zone 2 image is embedded too")
    ok("\\" not in [l for l in md.splitlines() if l.startswith("![")][0],
       "the embedded path uses forward slashes, so it survives a synced folder")

    # 2. a zone with no image contributes no broken link ----------------------
    doc_noimg = json.loads(json.dumps(DOC))
    doc_noimg["zones"][1].pop("image_path")
    md2 = arg._render_markdown(doc_noimg, "Merge", "Draft", "", md_dir=draft_dir)
    ok(md2.count("![") == 1, "a zone without an image produces no image tag at all")

    # 3. a stale path (image deleted) is skipped rather than embedded broken ---
    doc_stale = json.loads(json.dumps(DOC))
    doc_stale["zones"][0]["image_path"] = os.path.join(img_dir, "gone.png")
    md3 = arg._render_markdown(doc_stale, "Merge", "Draft", "", md_dir=draft_dir)
    ok("gone.png" not in md3, "a deleted image is skipped, not embedded as a broken link")

    # 3b. style references are picked PER EVENT TYPE ---------------------------
    # Before this, one flat folder served every type, so a Sort render would be
    # anchored on the Sim hazelnut room and pulled off-model.
    import art_reference as ar
    sr_root = os.path.join(tmp, "ideationData", "StyleReference", "images")
    os.makedirs(os.path.join(sr_root, "Sort"), exist_ok=True)
    os.makedirs(os.path.join(sr_root, "Sim"), exist_ok=True)
    make_png(os.path.join(sr_root, "Sort", "sort_zone.png"), (200, 120, 160))
    make_png(os.path.join(sr_root, "Sim", "sim_room.png"), (120, 160, 200))
    make_png(os.path.join(sr_root, "shared_fallback.png"), (160, 160, 160))

    sort_refs = [os.path.basename(x) for x in ar.style_reference_images(events, "Sort")]
    sim_refs = [os.path.basename(x) for x in ar.style_reference_images(events, "Sim")]
    ok(sort_refs == ["sort_zone.png"], "Sort gets only its own reference art", str(sort_refs))
    ok(sim_refs == ["sim_room.png"], "Sim gets only its own reference art", str(sim_refs))
    ok("sim_room.png" not in sort_refs,
       "a Sort render is never anchored on the Sim room (the bug this fixes)")
    fallback = [os.path.basename(x) for x in ar.style_reference_images(events, "Merge")]
    ok(fallback == ["shared_fallback.png"],
       "a type with no folder of its own falls back to the shared references", str(fallback))

    # 3b-i. PER-ZONE references ------------------------------------------------
    # Pad shape and proportions could not be held by prose: briefs stating
    # "roughly 2.1 : 1" rendered at 1.31, and one zone moved the wrong way. A
    # reference beats a description of one, so a zone can carry its own.
    zdir = os.path.join(sr_root, "Sort", ar.ZONE_DIR_NAME)
    os.makedirs(zdir, exist_ok=True)
    make_png(os.path.join(zdir, "zone2_pad.png"), (90, 190, 140))
    make_png(os.path.join(zdir, "zone3_pad.png"), (190, 90, 140))
    zr2 = [os.path.basename(x) for x in ar.zone_reference_images(events, "Sort", 2)]
    zr3 = [os.path.basename(x) for x in ar.zone_reference_images(events, "Sort", 3)]
    ok(zr2 and zr2[0] == "zone2_pad.png", "a zone's own reference is attached FIRST", str(zr2))
    ok(zr3 and zr3[0] == "zone3_pad.png", "and each zone gets its own, not the first one found",
       str(zr3))
    ok("zone3_pad.png" not in zr2 and "zone2_pad.png" not in zr3,
       "one zone's pad never leaks into another's render", "%s / %s" % (zr2, zr3))
    # A seeded zone gets its pad ALONE. Measured: with the style anchor also
    # attached the model blends the two, and only the zone whose anchor happened
    # to match its pad moved towards the target. Finish is carried by the spec in
    # prose, which holds it; geometry is what prose cannot hold, so it gets the slot.
    ok(zr2 == ["zone2_pad.png"],
       "a seeded zone is rendered against its own pad alone, no style anchor", str(zr2))
    ok(len(zr2) <= ar.MAX_REF_IMAGES, "and the 2-reference cap still applies", str(zr2))
    ar.ZONE_REF_STANDALONE = False
    try:
        topped = [os.path.basename(x) for x in ar.zone_reference_images(events, "Sort", 2)]
        ok(topped == ["zone2_pad.png", "sort_zone.png"],
           "and the flag restores the topped-up behaviour, so the change is reversible",
           str(topped))
    finally:
        ar.ZONE_REF_STANDALONE = True
    zr9 = [os.path.basename(x) for x in ar.zone_reference_images(events, "Sort", 9)]
    ok(zr9 == [os.path.basename(x) for x in ar.style_reference_images(events, "Sort")],
       "a zone with no reference of its own falls back to the type's, unchanged", str(zr9))
    ok([os.path.basename(x) for x in ar.zone_reference_images(events, "Sim", 1)]
       == ["sim_room.png"],
       "and a type never seeded per zone behaves exactly as before")
    ok("zone2_pad.png" not in [os.path.basename(x)
                               for x in ar.style_reference_images(events, "Sort")],
       "the zones/ subfolder does not pollute the type-level reference list")

    # 3b-i-b. wide render, portrait crop ---------------------------------------
    # A 2.12:1 pad has no room to be itself in a 3:4 frame, so the model fills
    # instead. Rendering wider gives it room; the crop restores the shipped
    # portrait framing. The trade is that cropping discards width, which is why
    # the render aspect wants to be only a little wider than the target.
    wide = make_png(os.path.join(img_dir, "wide.png"), (140, 170, 120))
    Image.new("RGB", (1600, 900), (140, 170, 120)).save(wide)
    artreq_images.crop_to_aspect(wide, "3:4")
    cw, ch = Image.open(wide).size
    ok(abs((cw / float(ch)) - 0.75) < 0.01,
       "a wide render is centre-cropped to the portrait target", "got %dx%d" % (cw, ch))
    ok(ch == 900, "the crop keeps full height and takes the width off the sides",
       "height %d" % ch)
    tall = os.path.join(img_dir, "tall.png")
    Image.new("RGB", (600, 1600), (120, 140, 170)).save(tall)
    artreq_images.crop_to_aspect(tall, "3:4")
    tw, th = Image.open(tall).size
    ok(abs((tw / float(th)) - 0.75) < 0.01 and tw == 600,
       "and a too-tall render is trimmed top and bottom instead", "got %dx%d" % (tw, th))
    same = os.path.join(img_dir, "same.png")
    Image.new("RGB", (900, 1200), (170, 120, 140)).save(same)
    artreq_images.crop_to_aspect(same, "3:4")
    ok(Image.open(same).size == (900, 1200),
       "an image already at the target aspect is left untouched")
    ok(artreq_images.crop_to_aspect(same, "nonsense") == same,
       "and an unparseable aspect is a no-op rather than a crash")

    # 3b-ii. the reference cap ------------------------------------------------
    # Five attached references made the model composite one of them into the
    # output instead of imitating its style. Two reproduced the style cleanly.
    for i in range(5):
        make_png(os.path.join(sr_root, "Sort", "extra%d.png" % i), (100 + i * 20, 130, 170))
    many = ar.style_reference_images(events, "Sort")
    ok(len(many) <= 2,
       "at most 2 references are attached even when more exist (%d available, %d used)"
       % (6, len(many)),
       "five references caused an entire reference room to be composited in")
    ok(ar.MAX_REF_IMAGES == 2, "MAX_REF_IMAGES is pinned at 2")

    # 3c. the RENDER SPEC is also per event type ------------------------------
    # Same bug class as the image folder: ROOM_RENDER_STYLE.md is titled "Sim
    # room art" and demands flat vector / matte / no texture / no specular, and
    # it was injected for every type. That is what flattened the Sort renders.
    sr_dir = os.path.join(tmp, "ideationData", "StyleReference")
    os.makedirs(sr_dir, exist_ok=True)
    io.open(os.path.join(sr_dir, "ROOM_RENDER_STYLE.md"), "w", encoding="utf-8").write(
        "# Sim room art\nFLAT TOY-LIKE VECTOR. MATTE. NO MATERIAL TEXTURE.")
    io.open(os.path.join(sr_dir, "Sort_RENDER_STYLE.md"), "w", encoding="utf-8").write(
        "# Sort zone art\nSOFT-PAINTED. Specular where glossy. Material detail wanted.")

    sort_spec = ar.load_style_spec(events, "Sort")
    sim_spec = ar.load_style_spec(events, "Sim")
    ok(sort_spec.startswith("# Sort zone art"), "Sort resolves its own render spec")
    ok("FLAT TOY-LIKE VECTOR" not in sort_spec,
       "the Sim flat-vector spec is NOT injected into a Sort prompt (the bug this fixes)")
    ok(sim_spec.startswith("# Sim room art"), "Sim falls back to ROOM_RENDER_STYLE.md")
    ok(ar.load_style_spec(events, "Blocks").startswith("# Sim room art"),
       "a type with no spec of its own falls back to the shared one")
    # and the prompt builder must actually use the per-type spec
    zprompt = ar.build_scene_prompt(events, "Sort", {"zone_number": 1, "zone_theme": "T"},
                                    distil_fn=None)
    ok("SOFT-PAINTED" in zprompt and "FLAT TOY-LIKE VECTOR" not in zprompt,
       "build_scene_prompt injects the type's own spec, not the Sim one")

    # 3d. the TEXT check is routed by event type ------------------------------
    # Sort and Merge are genuinely opposite here. On Sort any rendered text is a
    # defect. On Merge shop signage IS the art - the shipped Merge2606 concept
    # art carries MARKET, BAKERY, DEPT. STORE - and the defect is text the brief
    # did not authorise, typically a design name written on the building.
    import text_detect as _td
    ok(artreq_images.signage_from_prompt(
        "- Flower Shop - a low shop, sign reads FLOWER FOLD\n- sign reads MARKET\n")
        == ["FLOWER FOLD", "MARKET"], "sign wording is parsed out of the prompt")
    ok(artreq_images.signage_from_prompt("no signs here") == [],
       "a prompt declaring no signs yields no wording")
    ok("Merge" in artreq_images.TEXT_IS_SCENERY and "Sort" not in artreq_images.TEXT_IS_SCENERY,
       "Merge is the type where text is scenery; Sort is not")

    from PIL import ImageDraw, ImageFont
    def text_png(path, *labels):
        im = Image.new("RGB", (1000, 420), (140, 165, 135))
        dr = ImageDraw.Draw(im)
        try:
            fnt = ImageFont.truetype(r"C:\Windows\Fonts\arialbd.ttf", 46)
        except Exception:
            fnt = ImageFont.load_default()
        y = 60
        for lb in labels:
            dr.rectangle([60, y - 12, 60 + 26 * len(lb), y + 58], fill=(215, 195, 130))
            dr.text((76, y), lb, font=fnt, fill=(30, 22, 10))
            y += 120
        im.save(path)
        return path

    good = text_png(os.path.join(img_dir, "sig_good.png"), "MARKET")
    bad = text_png(os.path.join(img_dir, "sig_bad.png"), "MARKET", "MUSIC BOX PAVILION")
    allow = {os.path.basename(good): ["MARKET"], os.path.basename(bad): ["MARKET"]}

    r_sort = artreq_images.check_generated([good], event_type="Sort")
    ok(any(f["level"] == "error" for f in r_sort["findings"]),
       "on Sort, rendered signage is an error", str(r_sort["findings"]))

    r_legacy = artreq_images.check_generated([good])
    ok(any(f["level"] == "error" for f in r_legacy["findings"]),
       "with no event type the old Sort-style behaviour is kept",
       "callers that predate the routing must not silently stop checking")

    r_nosig = artreq_images.check_generated([good], event_type="Merge")
    ok(not any(f["level"] == "error" for f in r_nosig["findings"])
       and any(f["level"] == "warn" for f in r_nosig["findings"]),
       "on Merge with no declared wording it warns rather than erroring",
       str(r_nosig["findings"]))

    # None and [] are DIFFERENT declarations and conflating them is a real bug.
    # A per-building sheet for a structure with no sign passes [] - "this asset
    # carries no text" - which is an expectation to enforce, not an absence of
    # one. Treating [] as "undeclared" turned 14 building sheets into noise
    # warnings while silently not checking any of them.
    r_empty = artreq_images.check_generated(
        [good], event_type="Merge", signage={os.path.basename(good): []})
    if _td.ocr_available():
        ok(any(f["level"] == "error" for f in r_empty["findings"]),
           "an empty allowed-list means NO text is permitted, and text is an error",
           str(r_empty["findings"]))
    ok(not any("declared no wording" in f["message"] for f in r_empty["findings"]),
       "an explicit empty list is never reported as 'the brief said nothing'",
       str(r_empty["findings"]))

    if _td.ocr_available():
        r_ok = artreq_images.check_generated([good], event_type="Merge", signage=allow)
        ok(not any(f["level"] == "error" for f in r_ok["findings"]),
           "on Merge, authorised signage passes clean", str(r_ok["findings"]))
        r_bad = artreq_images.check_generated([bad], event_type="Merge", signage=allow)
        ok(any(f["level"] == "error" for f in r_bad["findings"]),
           "on Merge, an unauthorised design name is an error", str(r_bad["findings"]))
        ok(any("PAVILION" in f["message"].upper() for f in r_bad["findings"]),
           "and the offending wording is named in the finding", str(r_bad["findings"]))
    else:
        r_noocr = artreq_images.check_generated([good], event_type="Merge", signage=allow)
        ok(any(f["level"] == "warn" for f in r_noocr["findings"]),
           "without OCR the Merge text check warns - never a silent pass",
           str(r_noocr["findings"]))

    # ---- stub ComfyUI -------------------------------------------------------
    calls = []
    def fake_generate(base_url, prompt, dest, ref_paths=(), **kw):
        calls.append({"dest": dest, "prompt": prompt, "refs": list(ref_paths)})
        make_png(dest, (120, 200, 120))
        return {"ok": True, "path": dest, "bytes": os.path.getsize(dest)}
    real_gen, real_check = comfy_client.generate, comfy_client.check
    comfy_client.generate = fake_generate
    comfy_client.check = lambda url, timeout=20: {"state": "ok", "detail": "stub",
                                                  "reachable": True, "queue_ok": True,
                                                  "node_present": True, "version": "stub"}
    try:
        base = dict(comfy_url="http://stub", events_root=events, event_type="Merge",
                    author_id="AzureAD\\Tester", author_name="Tester")

        # 4. single-image regeneration touches ONLY that zone -----------------
        calls.clear()
        doc = json.loads(json.dumps(DOC))
        r = artreq_images.regenerate_images(dict(base, doc=doc, zone_number=1,
                                                 feedback="make the press teal"))
        ok(r["ok"] and r["regenerated"] == 1, "single-image regenerate reports one image")
        ok(len(calls) == 1 and calls[0]["dest"] == z1,
           "only zone 1 was re-rendered", str([c["dest"] for c in calls]))
        ok(calls[0]["refs"] == [z1],
           "the existing image is attached as the edit source, not re-rendered from scratch")
        ok("make the press teal" in calls[0]["prompt"], "the note reaches the prompt")
        ok("stays EXACTLY as it is" in calls[0]["prompt"],
           "the prompt holds the rest of the scene (edit, not fresh render)")

        # 5. apply-to-all touches every zone that has an image ----------------
        calls.clear()
        doc = json.loads(json.dumps(DOC))
        r = artreq_images.regenerate_images(dict(base, doc=doc, zone_number=1,
                                                 feedback="everything is too warm",
                                                 apply_to_all=True))
        ok(r["ok"] and r["regenerated"] == 2 and r["applied_to_all"],
           "apply-to-all regenerates every zone image")
        ok(sorted(c["dest"] for c in calls) == sorted([z1, z2]),
           "both zone images were re-rendered", str([c["dest"] for c in calls]))

        # 6. the note is stored under its own doc_kind ------------------------
        import ideation_store as store
        entries = store.read_feedback_entries(events, "Merge", doc_kind="image")
        ok(len(entries) == 2, "image feedback is recorded (%d entries)" % len(entries))
        ok(all(e["doc_kind"] == "image" for e in entries),
           "recorded under doc_kind 'image', not mixed into document feedback")
        doc_entries = store.read_feedback_entries(events, "Merge", doc_kind="artreq")
        ok(len(doc_entries) == 0, "document feedback is left untouched by image notes")

        # 7. an earlier image lesson conditions the next regeneration ---------
        calls.clear()
        doc = json.loads(json.dumps(DOC))
        artreq_images.regenerate_images(dict(base, doc=doc, zone_number=2,
                                             feedback="add more pumpkins"))
        ok("too warm" in calls[0]["prompt"] or "teal" in calls[0]["prompt"],
           "past image lessons are folded into the next image prompt")

        # 9b. the call site actually forwards the type -------------------------
        # Routing is worthless if the caller never passes it, and that wiring is
        # exactly the kind of thing a refactor drops silently.
        seen = {}
        _real_cg = artreq_images.check_generated
        def spy(paths, tones=None, event_type=None, signage=None, vision=None, counts=None):
            seen.update({"event_type": event_type, "signage": signage, "tones": tones,
                         "vision": vision, "counts": counts})
            return {"ok": True, "findings": []}
        artreq_images.check_generated = spy
        try:
            artreq_images.regenerate_images(dict(base, doc=json.loads(json.dumps(DOC)),
                                                 zone_number=1,
                                                 feedback="sign reads MARKET, EXACTLY EIGHT "
                                                          "buildings"))
            ok(seen.get("event_type") == "Merge",
               "regenerate_images forwards the event type to the reviewer",
               "got %r" % seen.get("event_type"))
            ok(isinstance(seen.get("signage"), dict),
               "and forwards a signage map keyed by file name", repr(seen.get("signage"))[:120])
            vals = [v for v in (seen.get("signage") or {}).values()]
            ok(vals and "MARKET" in vals[0],
               "wording named in the feedback reaches the reviewer as authorised",
               str(vals))
            ok(isinstance(seen.get("counts"), dict),
               "and forwards the expected building count per image",
               repr(seen.get("counts"))[:120])
            cvals = [v for v in (seen.get("counts") or {}).values()]
            ok(cvals and cvals[0] == 8,
               "the count stated in the brief reaches the reviewer", str(cvals))
            ok(seen.get("vision") is None,
               "with no provider in the payload, no vision config is forwarded",
               "the duplicate check must not silently try to run")

            # and with a provider configured, it IS forwarded
            seen.clear()
            artreq_images.regenerate_images(dict(base, doc=json.loads(json.dumps(DOC)),
                                                 zone_number=1, feedback="x",
                                                 provider="ollama",
                                                 ollama_base_url="http://stub:11434"))
            ok((seen.get("vision") or {}).get("provider") == "ollama",
               "a configured provider is forwarded for the duplicate check",
               repr(seen.get("vision")))
        finally:
            artreq_images.check_generated = _real_cg

        # 7b. the Dessert recipe and ingredient passes -------------------------
        # End to end through the real post-processing, because the geometry is
        # the part a prompt cannot control: the cut-out, the canvas and the fill
        # are applied afterwards, and a recipe that came back at the merge-icon
        # canvas would be silently wrong everywhere it is used.
        import icon_sheet
        def blob_generate(base_url, prompt, dest, ref_paths=(), **kw):
            # records refs too, so a sequence's derivation chain is checkable
            # What the prompt actually asks for: one object on a flat mid-grey
            # backdrop, so knockout_background has something to cut.
            calls.append({"dest": dest, "prompt": prompt, "refs": list(ref_paths)})
            os.makedirs(os.path.dirname(dest), exist_ok=True)
            im = Image.new("RGB", (512, 512), (128, 128, 128))
            for y in range(120, 400):
                for x in range(120, 400):
                    im.putpixel((x, y), (210, 120, 60))
            im.save(dest)
            return {"ok": True, "path": dest, "bytes": os.path.getsize(dest)}
        comfy_client.generate = blob_generate
        try:
            calls.clear()
            ddoc = json.loads(json.dumps(DOC))
            ddoc["event_motif"] = "the 5th-anniversary numeral 5"
            ddoc["producers"] = []
            ddoc["recipes"] = [{"recipe_number": 1, "name": "Honey Tart", "zone_number": 1,
                                "description": "a glazed tart", "palette": "amber",
                                "ingredients": ["Party Flour"], "secret": False}]
            ddoc["ingredients"] = [{"name": "Party Flour", "description": "a paper sack of flour",
                                    "palette": "cream", "secret": False}]
            dout = os.path.join(tmp, "dessert_out")
            res = artreq_images._generate_extra_passes(
                {"passes": ["recipe", "ingredient"]}, ddoc, ddoc["zones"], dout,
                "http://stub", "Dessert")
            ok(res["generated"] == 2 and res["requested"] == 2,
               "the recipe and ingredient passes each render their one asset", str(res["generated"]))
            rp = os.path.join(dout, "recipes", "recipe_01_honey_tart.png")
            ip = [c["dest"] for c in calls if "ingredient" in c["dest"]]
            ok(os.path.exists(rp), "the recipe lands under recipes/ with a readable name",
               str([c["dest"] for c in calls]))
            ok(bool(ip) and os.path.exists(ip[0]), "and the ingredient under ingredients/")

            def geom(p):
                im = Image.open(p).convert("RGBA")
                bb = im.getchannel("A").point(lambda a: 255 if a > 8 else 0).getbbox()
                return im.size, max(bb[2] - bb[0], bb[3] - bb[1]) / float(max(im.size))
            (rw, rh), rfrac = geom(rp)
            (iw, ih), ifrac = geom(ip[0])
            ok((rw, rh) == (icon_sheet.RECIPE_CANVAS, icon_sheet.RECIPE_CANVAS),
               "a recipe is normalised to the shipped 342px canvas", "%dx%d" % (rw, rh))
            ok((iw, ih) == (icon_sheet.INGREDIENT_CANVAS, icon_sheet.INGREDIENT_CANVAS),
               "an ingredient to the shipped 224px canvas", "%dx%d" % (iw, ih))
            ok(abs(rfrac - icon_sheet.RECIPE_FRACTION) < 0.04,
               "and each fills its canvas by the measured fraction, not the merge ladder",
               "recipe %.2f, want %.2f" % (rfrac, icon_sheet.RECIPE_FRACTION))
            ok(abs(ifrac - icon_sheet.INGREDIENT_FRACTION) < 0.04,
               "ingredient fill matches too",
               "ingredient %.2f, want %.2f" % (ifrac, icon_sheet.INGREDIENT_FRACTION))
            ok(ifrac > rfrac, "an ingredient reads bigger in frame than a recipe does")
            rprompt = [c["prompt"] for c in calls if c["dest"] == rp][0]
            ok("5th-anniversary numeral 5" in rprompt,
               "the event motif reaches the recipe prompt")

            # 7c. the scene pass is opt-out ------------------------------------
            # It used to run unconditionally, so asking for recipes alone still
            # spent one render per zone - re-rolling art that was already
            # approved - and the cost line, which comes from plan_passes,
            # under-reported by exactly that many.
            calls.clear()
            dd = json.loads(json.dumps(ddoc))
            r_noscene = artreq_images.generate_for_doc(
                dict(base, event_type="Dessert", doc=dd,
                     passes=["recipe"], index=500))
            zone_calls = [c for c in calls if os.sep + "zone" in c["dest"]
                          or os.path.basename(c["dest"]).startswith("zone")]
            ok(not zone_calls, "asking for recipes only renders no scenes at all",
               str([os.path.basename(c["dest"]) for c in calls]))
            ok(r_noscene["requested"] == 1,
               "and the requested count matches the plan rather than counting zones",
               "requested %s" % r_noscene["requested"])
            ok(r_noscene["requested"] == artreq_images.plan_passes(dd, ["recipe"])["total"],
               "generation and the cost estimate now agree")
            calls.clear()
            r_scene = artreq_images.generate_for_doc(
                dict(base, event_type="Dessert", doc=json.loads(json.dumps(ddoc)),
                     passes=["scene", "recipe"], index=501))
            ok(len([c for c in calls if os.path.basename(c["dest"]).startswith("zone")]) == 2,
               "while ticking scene still renders every zone", str(len(calls)))
            ok(r_scene["requested"] == 3, "and counts them", str(r_scene["requested"]))

            # 7d. the Blocks progression pass, end to end ----------------------
            # This exists because the pass shipped with a NameError in it - the
            # helper never had events_root in scope - and no unit test could see
            # it, only running the pass could. Stubbed ComfyUI, so it costs
            # nothing to keep.
            calls.clear()
            pdoc = json.loads(json.dumps(DOC))
            pdoc["zones"] = [dict(pdoc["zones"][0], zone_number=1, zone_theme="Cove",
                                  construction_steps=["a hero lantern", "a boat shed",
                                                      "a farm cart", "a hamper"])]
            pdoc["characters"] = [{"name": "Keeper", "base_character": "turtle",
                                   "costume": "a teal coat"}]
            pout = os.path.join(tmp, "prog_out")
            pres = artreq_images._generate_extra_passes(
                {"passes": ["progression"], "events_root": events, "aspect_ratio": "4:5"},
                pdoc, pdoc["zones"], pout, "http://stub", "Blocks")
            ok(pres["requested"] == 4 + 2 + 1,
               "a 4-step scene renders base + 3 earlier frames + empty + derelict + characters",
               "requested %s" % pres["requested"])
            ok(pres["requested"] == artreq_images.plan_passes(
                   pdoc, ["progression"], "Blocks")["total"],
               "and the plan agrees with what was actually rendered",
               "%s vs %s" % (pres["requested"],
                             artreq_images.plan_passes(pdoc, ["progression"], "Blocks")["total"]))
            names = sorted(os.path.basename(c["dest"]) for c in calls)
            ok(any(n.startswith("00_F0_derelict") for n in names)
               and any(n.startswith("00_F1_empty") for n in names),
               "the derelict and empty frames are both produced", str(names))
            ok(any("characters" in n for n in names), "and the character frame", str(names))
            byname = {os.path.basename(c["dest"]): c for c in calls}
            base_name = [n for n in names if n.endswith("_complete.png")][0]
            ok(not byname[base_name]["refs"] or "zone" not in byname[base_name]["refs"][0],
               "the base is rendered from the style references, not from a frame")
            hero = byname["01_F2.png"]
            ok("ONLY change" in hero["prompt"],
               "the hero-adjacent frame is cut with a single blunt change, not a bulk list")
            ok("ONLY change" not in byname["03_F4.png"]["prompt"],
               "while a later step is cut from the base with a cumulative list")

            # blunt_steps forces a dedicated hop at a named step. Without it a
            # large structure or flat ground element further up the ladder is cut
            # in bulk and silently survives - the barn's dance-floor sat at step
            # 5, and repairing it cost five renders because the fix cascades.
            calls.clear()
            bdoc = json.loads(json.dumps(pdoc))
            bdoc["zones"][0]["blunt_steps"] = [4]
            artreq_images._generate_extra_passes(
                {"passes": ["progression"], "events_root": events}, bdoc, bdoc["zones"],
                os.path.join(tmp, "blunt_out"), "http://stub", "Blocks")
            bn = {os.path.basename(c["dest"]): c for c in calls}
            ok("ONLY change" in bn["03_F4.png"]["prompt"],
               "naming step 4 in blunt_steps gives it its own single-change hop",
               bn["03_F4.png"]["prompt"][:90])
            ok(bn["03_F4.png"]["refs"]
               and "04_F5" in os.path.basename(bn["03_F4.png"]["refs"][0]),
               "cut from the frame directly after it, not from the base",
               str(bn["03_F4.png"]["refs"]))
            ok("ONLY change" in bn["01_F2.png"]["prompt"],
               "and steps 1 and 2 keep their automatic blunt hops")
            # Every frame BELOW a blunt step must chain too. Cutting one from the
            # base means a bulk list that still contains the difficult object, and
            # it survives again: blunt_steps=[5] fixed frame 5 and frame 4 then
            # put the dance floor straight back.
            ok("ONLY change" in bn["02_F3.png"]["prompt"],
               "frames below a blunt step chain too, rather than reverting to a bulk list",
               bn["02_F3.png"]["prompt"][:90])
            ok(bn["02_F3.png"]["refs"]
               and "03_F4" in os.path.basename(bn["02_F3.png"]["refs"][0]),
               "each cut from its own successor, so the difficult object is already gone",
               str(bn["02_F3.png"]["refs"]))
            calls.clear()
            gdoc = json.loads(json.dumps(pdoc))
            gdoc["zones"][0]["blunt_steps"] = ["nonsense", 4]
            artreq_images._generate_extra_passes(
                {"passes": ["progression"], "events_root": events}, gdoc, gdoc["zones"],
                os.path.join(tmp, "blunt_bad"), "http://stub", "Blocks")
            gn = {os.path.basename(c["dest"]): c for c in calls}
            ok("ONLY change" in gn["03_F4.png"]["prompt"],
               "a junk entry is skipped rather than crashing the whole sequence")
            der = byname["00_F0_derelict.png"]
            ok("FROZEN" in der["prompt"] and "REMOVE EVERYTHING HUMANISED" in der["prompt"],
               "the derelict frame is frozen and stripped of humanised things")
            ok(der["refs"] and "00_F1_empty" in der["refs"][0],
               "and is derived from the EMPTY frame, so nothing built survives into it",
               str(der["refs"]))
            ch = byname[[n for n in names if "characters" in n][0]]
            ok(ch["refs"] and ch["refs"][0].endswith("_complete.png"),
               "characters are one additive hop from the completed base", str(ch["refs"]))
            ok("progression_scene1" in pres["reviews"],
               "and the set is handed to sequence_review automatically",
               str(list(pres["reviews"])))
            calls.clear()
            skipped = artreq_images._generate_extra_passes(
                {"passes": ["progression"], "events_root": events}, pdoc, pdoc["zones"],
                os.path.join(tmp, "prog_sort"), "http://stub", "Sort")
            ok(skipped["requested"] == 0 and not calls,
               "a non-Blocks type renders no sequence even when asked", str(skipped))

            # and the other type's passes produce nothing from this document
            calls.clear()
            res2 = artreq_images._generate_extra_passes(
                {"passes": ["icon", "producer"]}, ddoc, ddoc["zones"],
                os.path.join(tmp, "dessert_out2"), "http://stub", "Dessert")
            ok(res2["requested"] == 0 and not calls,
               "a Dessert document renders no chain or producer icons", str(res2))
        finally:
            comfy_client.generate = fake_generate

        # 8. empty feedback is refused before any render ----------------------
        calls.clear()
        r = artreq_images.regenerate_images(dict(base, doc=json.loads(json.dumps(DOC)),
                                                 zone_number=1, feedback="   "))
        ok((not r["ok"]) and not calls, "empty feedback is refused without calling ComfyUI")

        # 9. a zone with no generated image is refused clearly ----------------
        d = json.loads(json.dumps(DOC))
        for z in d["zones"]:
            z.pop("image_path")
        r = artreq_images.regenerate_images(dict(base, doc=d, zone_number=1, feedback="x"))
        ok((not r["ok"]) and "Generate images first" in r["error"],
           "regenerating with no image yet gives an actionable error", str(r.get("error")))
    finally:
        comfy_client.generate, comfy_client.check = real_gen, real_check

    # 10. a wedged server is surfaced verbatim, not as a generic failure ------
    comfy_client.check = lambda url, timeout=20: {"state": "wedged", "detail": "worker stuck",
                                                  "reachable": True, "queue_ok": False,
                                                  "node_present": True, "version": "0.18.2"}
    try:
        r = artreq_images.regenerate_images(dict(comfy_url="http://stub", events_root=events,
                                                 event_type="Merge", doc=json.loads(json.dumps(DOC)),
                                                 zone_number=1, feedback="x"))
        ok((not r["ok"]) and r.get("state") == "wedged" and "worker stuck" in r["error"],
           "a wedged ComfyUI is reported as wedged, with its own detail text")
    finally:
        comfy_client.check = real_check
finally:
    shutil.rmtree(tmp, ignore_errors=True)

# ---- the REAL StyleReference tree: Merge seeding -----------------------------
# Everything above runs against a temp fixture. This block checks the shipped
# tree, because the Merge gap was a data problem, not a code one: Merge had no
# spec and no reference images, so it silently resolved to the Sim room spec
# with nothing attached.
REAL = r"C:\SkipBo\tools\EventAutomationDesktop\ideationData\Events"
if os.path.isdir(REAL):
    import art_reference as _ar
    print("\n--- real StyleReference tree ---")
    for t in ("Sort", "Sim", "Merge"):
        spec = (_ar.load_style_spec(REAL, t) or "").splitlines()
        refs = [os.path.basename(x) for x in _ar.style_reference_images(REAL, t)]
        print("  %-6s %-40s %s" % (t, (spec or ["(none)"])[0][:40], refs))

    dspec = _ar.load_style_spec(REAL, "Dessert") or ""
    drefs = _ar.style_reference_images(REAL, "Dessert")
    ok(dspec.startswith("# Dessert zone art"),
       "Dessert resolves its own render spec, not the Sim one",
       "got: " + (dspec.splitlines() or ["(empty)"])[0][:60])
    ok(len(drefs) == 2, "Dessert has exactly 2 seeded references", "got %d" % len(drefs))
    ok(all(os.path.getsize(x) > 20000 for x in drefs),
       "both seeded Dessert references are real images")

    mspec = _ar.load_style_spec(REAL, "Merge") or ""
    mrefs = _ar.style_reference_images(REAL, "Merge")
    ok(mspec.startswith("# Merge scene art"),
       "Merge resolves its own render spec, not the Sim one",
       "got: " + (mspec.splitlines() or ["(empty)"])[0][:60])
    ok("# Sim room art" not in mspec.splitlines()[0] if mspec else False,
       "the Sim spec is not what Merge gets")
    ok(len(mrefs) == 2, "Merge has exactly 2 seeded references", "got %d" % len(mrefs))
    ok(not any("CALLOUT" in os.path.basename(x).upper() for x in mrefs),
       "the callout-numbered Toy Car scene is never selected as a reference",
       "a numbered reference is what taught an earlier render to draw callout markers")
    ok(all(os.path.getsize(x) > 20000 for x in mrefs),
       "both seeded Merge references are real images, not stubs")
    # the quarantine mechanism itself: a subfolder must not be walked
    excl = os.path.join(os.path.dirname(mrefs[0]), "_excluded")
    if os.path.isdir(excl):
        ok(any(n.lower().endswith(".png") for n in os.listdir(excl)),
           "the quarantine folder still holds the excluded art")
    # Sort and Sim must be untouched by the Merge work
    ok((_ar.load_style_spec(REAL, "Sort") or "").startswith("# Sort zone art"),
       "Sort still resolves its own spec")
    ok(len(_ar.style_reference_images(REAL, "Sort")) == 2,
       "Sort still has its 2 references")

print("")
if fails:
    print("%d CHECK(S) FAILED" % len(fails)); sys.exit(1)
print("ALL ARTREQ IMAGE TESTS PASSED")
