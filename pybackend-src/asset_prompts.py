"""Prompt builders for the asset classes beyond the scene overview.

An art requirement carries four kinds of original art beyond the scene:

    building sheet    one per building, at its final level
    upgrade stage     one per level below the final one
    chain item icon   5 per chain, 40 for a Merge event
    producer icon     2 per producer, 8 for a Merge event
    recipe            1 per dish, 39 for a Dessert event (24 common + 15 secret)
    ingredient        1 per component, 11 for a Dessert event (6 common + 5 secret)

Merge and Dessert do not share their mechanic-specific classes: a Merge event
has producers and chains and no recipes, a Dessert event has recipes and
ingredients and no producers. The passes are gated on which the document
actually carries, so neither type is ever asked for the other's assets.

Every rule below was arrived at by generating, reviewing and correcting real
output, and each one is here because its absence produced a defect:

  * the ground under a building sheet grew into a paved courtyard with boundary
    walls, benches and a pool until it was capped by a SIZE RELATION - prose
    like "keep it small" did not hold, "no wider than one and a half times the
    footprint" did;
  * a building's design name was written across the building twice, so the ban
    is stated separately from the signage instruction and names the failure;
  * an upgrade stage came back as a different building until "same footprint,
    same camera, same ground patch" was made explicit;
  * icons must be generated on a FLAT NEUTRAL backdrop or they cannot be cut
    out cleanly - and geometry (canvas, size ladder, stroke) is applied
    afterwards by icon_sheet, never asked for here.

The builders take schema objects (a zone, a building, a producer, a chain) and
nothing else, so the document is the single source of truth.
"""

NO_ANNOTATION = (
    "*** NO ANNOTATION ***\n"
    "This is finished artwork, not a labelled diagram. NO callout numbers, NO circled markers, NO legend, "
    "NO caption, NO arrows. NEVER write the building's design name anywhere - a previous render wrote a "
    "building's name across its front and that is completely wrong. A building is identified by its shape "
    "and its fittings, never by a label."
)

RENDER_STYLE = (
    "*** RENDER STYLE ***\n"
    "Soft-shaded dimensional toy diorama. Smooth gradients, deep saturated colour, generous micro-detail. "
    "Soft ambient occlusion and a firm contact shadow grounding the structure. Chunky rounded forms that "
    "hold their silhouette at small size.\n"
    "NO OUTLINES - no dark contour, ink line or keyline. Forms separate by colour and shading alone.\n"
    "NOT a photoreal 3D render. No depth-of-field blur, no scanned materials."
)

GROUND_RULE = (
    "*** WHAT IT STANDS ON - KEEP THIS SMALL ***\n"
    "A SMALL, IRREGULAR, SOFT-EDGED PATCH OF GROUND, no wider than about ONE AND A HALF TIMES the "
    "building's own footprint - a scrap of ground it happens to stand on, nothing more. Its edges are "
    "uneven and fade into the backdrop; it never ends in a straight cut or a clean border.\n"
    "A previous attempt turned the ground into a designed courtyard: a paved precinct with a geometric "
    "outline, raised planter walls round the edge, benches down both sides, matching trees flanking left "
    "and right, corner lamp posts and an ornamental pool - filling the bottom half while the building "
    "shrank into the middle. That is completely wrong.\n"
    "So NONE of: plaza, courtyard, terrace, precinct, boundary or planter walls, run of benches, "
    "symmetrical flanking trees, corner lamp posts, pools, wide steps wrapping the building, or a raised "
    "platform with visible straight sides. The ground is not architecture. It is just ground."
)

BACKDROP_RULE = (
    "*** BACKGROUND ***\n"
    "A simple, soft, uncluttered backdrop: open sky and a hint of distant greenery in the scene's colours, "
    "quietly out of focus. No other structures anywhere in it. Do not leave it blank and washed-out - it "
    "reads as the same countryside the scene sits in, only softened."
)


def _signage_block(signage):
    """What may be written, stated so that an empty list is unambiguous."""
    if signage:
        joined = ", ".join(signage)
        return ("It carries a clear, correctly spelled sign reading exactly: %s. Those are the ONLY words "
                "in the picture." % joined,
                "The only words anywhere are the sign: %s." % joined)
    return ("This building carries NO sign and NO text of any kind.",
            "There is NO text anywhere in this picture at all.")


def _material_block(zone):
    mat = (zone.get("scene_material") or "").strip()
    pal = (zone.get("palette_or_mood") or "").strip()
    out = []
    if mat:
        out.append("SCENE MATERIAL: %s\nEvery wall, roof, sign and fitting is built out of it and shows how "
                   "it is made - its seams, its joins, its edges, its surface. The material is not "
                   "decoration applied to ordinary architecture; it IS the architecture." % mat)
    if pal:
        out.append("PALETTE: %s" % pal)
    tone = (zone.get("main_tone") or "").strip().upper()
    if tone in ("COOL", "WARM"):
        out.append("MAIN TONE: %s" % tone)
    return "\n\n".join(out)


def building_prompt(zone, building):
    """One building on its own - the per-building sheet the documents carry."""
    sign_line, text_rule = _signage_block(building.get("signage"))
    feats = building.get("key_features") or []
    feat_md = "\n".join("- %s" % f for f in feats)
    return """*** ONE SINGLE BUILDING, ON ITS OWN ***
Draw EXACTLY ONE building, centred in the frame and filling most of it. This is a single asset sheet for one
structure, NOT a scene and NOT a settlement. DO NOT draw the other buildings, a plaza, a fountain, a road
network or a row of shops. There is ONE building here and nothing else of any size.

*** HOW BIG IT IS IN FRAME ***
THE BUILDING IS THE PICTURE. It stands tall in the middle and takes up most of the frame's height.
Everything else is a thin margin around it. If the building looks small in its setting, the picture is wrong.

{ground}

{backdrop}

*** CAMERA ***
THREE-QUARTER TOP-DOWN with real perspective, roughly 30 to 40 degrees, seeing its front and one side.
NOT flat elevation, NOT eye level, NOT straight down. Portrait format, slightly taller than wide.

{material}

{style}

*** THE BUILDING ***
{name}: {silhouette}

It must be identifiable from its outline alone at thumbnail size. These are its recognition points and every
one of them must be visible:
{features}

{sign_line}

{no_annotation}
{text_rule}

*** FINAL CHECK ***
1. Exactly ONE building, standing tall and taking most of the frame's height.
2. A SMALL irregular soft-edged ground patch - no courtyard, no walls, no benches, no flanking trees.
3. Three-quarter top-down camera. 4. Built from the scene's material, showing how it is made.
5. Every recognition point above is visible. 6. {text_rule}
7. No outlines, no characters, no annotation of any kind.""".format(
        ground=GROUND_RULE, backdrop=BACKDROP_RULE, material=_material_block(zone),
        style=RENDER_STYLE, name=building.get("name", "The building"),
        silhouette=building.get("silhouette", ""), features=feat_md or "- (none given)",
        sign_line=sign_line, no_annotation=NO_ANNOTATION, text_rule=text_rule)


# What is present at an earlier level. Cumulative: each keeps everything before it.
STAGE_RULES = {
    (2, 1): ("THE EARLY STAGE OF TWO. The main structure stands but is plainly UNFINISHED: its core mass is "
             "built and its shape is already recognisable, but it is bare. No sign. No goods, stock or "
             "produce on display. No awning, banners, bunting or hanging decoration. The surface is plain "
             "and undressed. A little raw material of the scene's own kind is stacked beside it, as if work "
             "is still going on."),
    (3, 1): ("THE FIRST STAGE OF THREE - BARELY STARTED. Only the FOUNDATION and a partial FRAME: the "
             "footprint marked out on the ground and the skeleton of the structure going up, open and "
             "see-through, with no walls closed in and no roof. Clearly a building site, but already in the "
             "right footprint and the right shape."),
    (3, 2): ("THE SECOND STAGE OF THREE. The structure is now COMPLETE and closed in - walls and roof done, "
             "the silhouette finished - but entirely UNDECORATED. No sign. No goods or stock. No awning, "
             "banners or hanging decoration. Plain and functional, waiting to be dressed."),
}


def stage_rule(max_level, level):
    """The description of one earlier level, or None if there is no rule for it."""
    return STAGE_RULES.get((int(max_level), int(level)))


def stage_prompt(zone, building, max_level, level):
    """The same building at an earlier construction level.

    A building is signed when it opens, so an earlier stage carries NO text -
    which the reviewer then enforces strictly rather than merely warning.
    """
    rule = stage_rule(max_level, level)
    if not rule:
        raise ValueError("no stage rule for level %s of %s" % (level, max_level))
    return """*** THIS IS THE SAME BUILDING AT AN EARLIER CONSTRUCTION STAGE ***
The attached image is this building FINISHED. Draw the SAME building, in the SAME place, at the same size, at
the same camera angle, on the same ground patch - but at an earlier stage of its construction.

{rule}

*** WHAT MUST NOT CHANGE ***
Same footprint and position in frame. Same camera. Same ground patch, same shape and size. Same material,
palette and light as the attached image. Same background. This must read as ONE building photographed twice
during its build, never as two different buildings.

*** WHAT IT IS ***
{name}: {silhouette}

{material}

{style}

{ground}

*** NO TEXT AT ALL ***
There is NO text anywhere in this picture. The building has NOT been signed yet - the sign goes up only when
it is finished. No lettering, no numbers, no callout markers, no legend, no caption, and never the building's
design name.

*** FINAL CHECK ***
1. Recognisably the SAME building as the attached image, same place, same size, same ground.
2. {short}
3. NO text of any kind anywhere. 4. No outlines, no characters, no annotation.""".format(
        rule=rule, short=rule.split(".")[0], name=building.get("name", "The building"),
        silhouette=building.get("silhouette", ""), material=_material_block(zone),
        style=RENDER_STYLE, ground=GROUND_RULE)


ICON_BACKDROP = (
    "*** BACKGROUND - THIS MUST BE EASY TO CUT OUT ***\n"
    "A PLAIN FLAT SOLID MID-GREY BACKGROUND, one single uniform colour edge to edge. No gradient, no "
    "vignette, no scenery, no texture, no pattern. The object casts NO shadow onto the background - the "
    "background is completely empty and even, so the object can be cut out cleanly."
)

ICON_STYLE = (
    "*** RENDER STYLE - MATCH A MOBILE GAME ITEM ICON ***\n"
    "Soft-shaded and glossy, deep saturated colour, smooth airbrushed gradients with a wide tonal range. "
    "Bright specular highlights where the material is glossy. Chunky, rounded, slightly exaggerated toy-like "
    "forms with a bold clear silhouette that still reads when small. Soft self-shading and ambient occlusion "
    "on the object itself. Not flat vector, not matte, not a photoreal 3D render, no depth-of-field blur."
)

ICON_NO_TEXT = (
    "*** NO TEXT, NO OUTLINE ***\n"
    "There is NO text, NO lettering, NO numbers and NO labels anywhere in this picture.\n"
    "Do NOT draw an outline, keyline or border around the object - its edge is left clean."
)


def chain_item_prompt(chain, level):
    """One rung of a merge chain.

    Canvas, size ladder and stroke are deliberately absent: icon_sheet applies
    all three afterwards, because they are geometry and a prompt controls
    geometry worst. This asks only for the object.
    """
    items = list(chain.get("items") or [])
    item = items[level - 1] if 0 < level <= len(items) else ""
    ladder = "; ".join("%d) %s" % (i + 1, s.split(",")[0]) for i, s in enumerate(items))
    return """*** A SINGLE GAME ITEM ICON ***
Draw ONE object only, centred, filling the frame: {item}.
Nothing else is in the picture. No second copy of it, no scene, no table, no hands, no characters.

{backdrop}

{style}

*** COLOUR ***
Built from {palette}. Rich and appetising, never washed out or greyed.

*** SILHOUETTE ***
Compact and readable. Keep it roughly as wide as it is tall unless the object is naturally long.
No thin spindly parts that would disappear at icon size.

{no_text}

*** THIS IS RUNG {level} OF {n} IN A SET ***
Every rung belongs to ONE family: {family}. The set runs: {ladder}.
This one is rung {level}. It shares the family's materials and colours, but it is its own distinct object -
NEVER a copy of a neighbouring rung, and NEVER several copies of a smaller item stacked up to look bigger.""".format(
        item=item, backdrop=ICON_BACKDROP, style=ICON_STYLE,
        palette=chain.get("palette", "the event palette"), no_text=ICON_NO_TEXT,
        level=level, n=len(items) or 5, family=chain.get("family", "one family"), ladder=ladder)


PRODUCER_OPEN_RULE = """
*** THIS IS THE UPGRADED, OPEN VERSION ***
It is the SAME vessel as before - same shape, same materials, same colours, same fittings - now OPEN, showing
what it makes. Keep the vessel unmistakably recognisable; only the lid position and the visible contents
change. A previous attempt turned a square brass-bound case into a rounded pot, which is wrong. The contents
sit INSIDE it and stay small enough that the vessel still reads as the subject. Make this version feel a
clear step richer than the closed one: a little more trim, a little more polish."""


def producer_prompt(producer, level):
    """A producer at level 1 (the vessel closed) or level 2 (the same vessel open).

    The convention is read off the shipped Merge2608 producers: a closed
    suitcase becomes the same suitcase open with a sun hat and a pass inside.
    """
    subject = (producer.get("closed_description") if level == 1
               else producer.get("open_description")) or producer.get("name", "")
    return """*** A SINGLE GAME PRODUCER ICON ***
Draw ONE object only, centred, FILLING THE FRAME: {subject}.
It is the machine the player taps, so it is drawn BIG and bold - it fills almost the whole picture.
Nothing else is present. No second copy, no scene, no table, no hands, no characters.

{backdrop}

{style}

*** COLOUR ***
Built from {palette}. Rich and appetising, never washed out or greyed.

{no_text}{open_rule}""".format(
        subject=subject, backdrop=ICON_BACKDROP, style=ICON_STYLE,
        palette=producer.get("palette", "the event palette"), no_text=ICON_NO_TEXT,
        open_rule=PRODUCER_OPEN_RULE if level == 2 else "")


DISH_STYLE = (
    "*** RENDER STYLE - MATCH A MOBILE GAME FOOD ICON ***\n"
    "Glossy, appetising and deeply saturated, with smooth airbrushed shading and bright specular highlights "
    "on anything wet, glazed or sugared. Chunky, rounded, slightly exaggerated forms with a bold silhouette "
    "that still reads when small.\n"
    "UNLIKE the scene art, this asset IS drawn with a clean DARK KEYLINE around its outer edge and around "
    "its major internal shapes - a food icon in this game is outlined, the scenery is not. Keep the line "
    "even and confident, a dark version of the object's own colour rather than pure black."
)


def recipe_prompt(recipe, motif="", palette=""):
    """One finished dish - the signature Dessert asset.

    Two rules here are read off the shipped Dessert2608 set rather than invented:
    every one of the 24 common recipes carries the event's motif somewhere in the
    dish (that event's was the anniversary "5" - a 5-shaped cookie, a 5 piped on
    a pie, a 5 candle), and each dish stays inside its zone's colour family so a
    player can tell at a glance which zone a recipe belongs to.

    Canvas and fill are applied afterwards by icon_sheet - see RECIPE_CANVAS.
    """
    motif = (motif or "").strip()
    motif_block = ("""

*** THE EVENT MOTIF IS BAKED INTO THIS DISH ***
The event's motif is: {motif}.
It appears IN or ON this dish, worked into the food itself - piped, moulded, cut, stamped, stuck in as a
decoration or formed by the shape of the dish. Every recipe in the event carries it, so a player recognises
the set instantly. Work it in the way THIS dish would naturally carry it; do not repeat another dish's
solution, and do not paste it on as a flat sticker or a logo floating beside the food.""".format(motif=motif)
                    if motif else "")
    ings = [str(i).strip() for i in (recipe.get("ingredients") or []) if str(i).strip()]
    ing_block = ("""

*** WHAT IT IS MADE FROM ***
It is made from: {ings}. Those ingredients must be visible or clearly implied in the finished dish - a player
reads the recipe off the picture.""".format(ings=", ".join(ings)) if ings else "")
    return """*** A SINGLE FINISHED DISH ***
Draw ONE dish only, centred, filling the frame: {name} - {desc}.
It is plated or served as it would be handed to a customer. Nothing else is in the picture: no second
portion, no table, no cutlery laid out beside it, no hands, no characters, no kitchen behind it.

{backdrop}

{style}

*** COLOUR ***
Built from {palette}. This dish belongs to that zone's colour family and stays inside it - rich, appetising
and confident, never washed out, greyed or muddy.

*** SILHOUETTE ***
Compact and readable at small size. Roughly as wide as it is tall unless the dish is naturally tall.
No thin spindly garnish that would disappear when the icon is scaled down.

{no_text}{motif_block}{ing_block}""".format(
        name=recipe.get("name", "a dish"), desc=recipe.get("description", ""),
        backdrop=ICON_BACKDROP, style=DISH_STYLE, no_text=ICON_NO_TEXT,
        palette=recipe.get("palette") or palette or "the zone palette",
        motif_block=motif_block, ing_block=ing_block)


def ingredient_prompt(ingredient, palette=""):
    """One raw ingredient.

    The shipped ingredient set is a fixed vocabulary of slots - flour, fruit,
    ice, matcha, meat, milk, mint, spice, sweet, flower, pumpkin - re-drawn per
    event rather than re-invented, so this asks for a RAW COMPONENT and says so
    plainly. The failure it guards against is an ingredient coming back plated
    and garnished, at which point it is indistinguishable from a recipe.
    """
    return """*** A SINGLE RAW INGREDIENT ***
Draw ONE ingredient only, centred, filling the frame: {name} - {desc}.
This is a RAW COMPONENT, not a finished dish. It is the thing before it is cooked: the sack of flour, the
piece of fruit, the pat of butter. It is NOT plated, NOT garnished, NOT arranged on a dish and NOT served.
Nothing else is in the picture - no bowl of other ingredients, no scene, no hands, no characters.

A small neat portion of it is enough to read the ingredient: one item, or a tidy little group of the same
item if that is how it is naturally handled. Never a scattered pile.

{backdrop}

{style}

*** COLOUR ***
Built from {palette}. Fresh and appetising, never washed out or greyed.

*** SILHOUETTE ***
Compact and bold - this icon is shown small, next to a recipe, and must be told apart from ten others at a
glance. Roughly as wide as it is tall.

{no_text}""".format(
        name=ingredient.get("name", "an ingredient"), desc=ingredient.get("description", ""),
        backdrop=ICON_BACKDROP, style=DISH_STYLE, no_text=ICON_NO_TEXT,
        palette=ingredient.get("palette") or palette or "the event palette")


def _object_prompt(subject, palette, extra=""):
    """The shared shape for a standalone themed object rendered as a cut-out."""
    return """*** A SINGLE OBJECT, ON ITS OWN ***
Draw ONE object only, centred, filling the frame: {subject}.
Nothing else is in the picture. No second copy, no scene, no table, no hands, no characters.

{backdrop}

{style}

*** COLOUR ***
Built from {palette}. Rich and appetising, never washed out or greyed.

{no_text}{extra}""".format(subject=subject, backdrop=ICON_BACKDROP, style=ICON_STYLE,
                           palette=palette or "the event palette", no_text=ICON_NO_TEXT, extra=extra)


BALLOON_EXTRA = """

*** IT IS A HOT AIR BALLOON ***
A complete hot air balloon seen from the side and slightly above: a full rounded envelope on top, cords or
netting running down, and a basket beneath. It floats free with nothing holding it. The envelope is the
largest part of the picture and carries the event's motifs across its panels."""

BADGE_EXTRA = """

*** IT IS AN ACHIEVEMENT BADGE ***
A single emblem: a symmetrical medal or crest built from the motif described, with a clear outer frame and a
bold readable centre. It reads at small size, so keep the interior simple and the silhouette distinctive.
It is a badge, not a scene inside a circle."""


def balloon_prompt(description, palette=""):
    """The themed hot air balloon every Merge event needs at least one of."""
    return _object_prompt(description or "a themed hot air balloon", palette, BALLOON_EXTRA)


def badge_prompt(description, palette=""):
    """The event achievement badge, built from one of the event's own motifs."""
    return _object_prompt(description or "an event achievement badge", palette, BADGE_EXTRA)


def task_icon_prompt(description, palette=""):
    """One of the 3-4 find-the-object task icons."""
    return _object_prompt(description or "a small task object", palette)
