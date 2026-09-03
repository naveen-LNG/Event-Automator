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
3. Camera per the event type's composition rules - Merge and Dessert use a three-quarter top-down view, Blocks is upfront and straight on. 4. Built from the scene's material, showing how it is made.
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


# --------------------------------------------------------- scene evolution ---
# One hold clause, used by every hop in a sequence. It is long on purpose: the
# whole method depends on each edit changing exactly one thing, and an edit that
# quietly re-draws the scene shows up as palette drift three frames later.
HOLD = (
    "Everything else stays EXACTLY as it is - the same camera and framing, the same foreground "
    "framing device, the same ground, the same background and sky, the same light direction, the "
    "same warm key and cool ambient shadows, the same palette and the same finish. Do NOT re-draw "
    "the scene, do NOT move anything, do NOT add anything, do NOT change the time of day. "
    "No text anywhere."
)


# Where each build step lands, in order. The art director's note: populate from
# the bottom of the frame upwards, and JUMP between areas rather than clustering
# around the hero - the eye travelling makes each step land harder. The
# subtractive method can only reveal an arrangement the base already has, so the
# distribution has to be specified when the base is drawn, not later.
# Where each build step lands, in build order. Two rules hold at once, and the
# art lead asked for both: the sequence climbs the frame BOTTOM TO TOP, and it
# ALTERNATES LEFT AND RIGHT as it climbs. A step that appears next to the one
# before it is a small visible change; a step that jumps across the frame and up
# is a large one, and the build sequence is the thing the player watches.
#
# The previous ladder alternated sides but zig-zagged in depth - slot 6 was
# "furthest back" and slot 7 came back to "close to the camera" - so it was
# rising and falling through the frame while claiming, in its own docstring, to
# rise. Depth rank is recorded per slot now so the claim is checkable:
# PLACEMENT_DEPTH must never decrease.
_PLACEMENT_SLOTS = (
    (0, "C", "at the bottom centre"),
    (1, "L", "at the bottom LEFT, close to the camera"),
    (1, "R", "at the bottom RIGHT, close to the camera"),
    (2, "L", "on the LEFT, just above the near ground"),
    (2, "R", "on the RIGHT, just above the near ground"),
    (3, "L", "in the MIDDLE DISTANCE on the left"),
    (3, "R", "in the MIDDLE DISTANCE on the right"),
    (4, "C", "across the UPPER part of the scene, furthest back"),
)
PLACEMENT = [text for _d, _s, text in _PLACEMENT_SLOTS]
PLACEMENT_DEPTH = [d for d, _s, _t in _PLACEMENT_SLOTS]
PLACEMENT_SIDE = [s for _d, s, _t in _PLACEMENT_SLOTS]


def placement_slots(n):
    """Which ladder slots an n-step build uses, bottom to top.

    The first n-1 slots then the topmost one, so the sequence always finishes at
    the top of the frame however many steps a scene has - i % len(PLACEMENT) left
    a seven-step scene ending in the middle distance. Both invariants survive at
    every n: depth non-decreasing, sides alternating.
    """
    last = len(PLACEMENT) - 1
    if n <= 1:
        return [0]
    if n > len(PLACEMENT):
        return [min(i, last) for i in range(n)]
    return list(range(n - 1)) + [last]


def built_scene(steps):
    """The completed scene: everything present, hero first, distribution fixed.

    Each item gets its own region, assigned in build order, so the sequence
    climbs the frame bottom to top while alternating left and right. Both halves
    of that are checked - see PLACEMENT_DEPTH and PLACEMENT_SIDE.
    """
    slots = placement_slots(len(steps))
    body = "\n".join("%d. %s - placed %s"
                     % (i + 1, s, PLACEMENT[slots[i]]) for i, s in enumerate(steps))
    return ("*** THE COMPLETED SCENE ***\nEverything below is built and standing. The first item is "
            "the HERO. Each item sits in its own part of the frame, spread apart - do NOT cluster "
            "them together around the hero:\n%s\n"
            "SIZE BY DEPTH: the LARGEST structures belong in the middle ground or further back. "
            "The near ground takes small and medium props only. A big object sitting right at the "
            "camera hides everything behind it and swallows the changes the build steps make.\n"
            "SEPARATION: leave air between objects. They may sit in front of one another but must "
            "NOT overlap drastically - each silhouette has to read on its own, because a step that "
            "adds something largely hidden behind what is already there produces no visible change "
            "and the sequence stops reading as construction.\n"
            "Each is chunky, rounded and slightly inflated, lit by the warm key with a cool "
            "ambient shadow and a soft rim on its silhouette. The scene should read as HUMANISED "
            "ACTIVITY spread across the whole frame, near to far." % body)


def remove_one(item):
    """A single blunt removal - the only reliable way to take out a big object.

    A long removal list leaves the largest object standing: an eight-category
    instruction left the press and the cart in place, and a six-item one left a
    boat shed. Both were fixed by giving the big object its own hop.
    """
    return ("Revise this image. Delete %s completely, leaving clean empty ground where it stood. "
            "That is the ONLY change.\n\n%s" % (item, HOLD))


def remove_many(items):
    """Cumulative removal of the secondary props - safe in bulk, unlike a hero.

    Cumulative, not one-per-frame: removing a single item from the base left
    every middle frame still holding everything else, so the sequence never
    built up.
    """
    joined = ", ".join(items[:-1]) + (" and " + items[-1] if len(items) > 1 else items[0])
    return ("Revise this image. Remove %s - remove ALL of those, completely, leaving clean empty "
            "ground where they stood.\n\n%s" % (joined, HOLD))


def derelict_prompt(zone):
    """The 'before' state: EMPTY of everything humanised, and DIFFERENT per scene.

    Two corrections from the art director, in order.

    First: the instinct to show neglect - broken boards, slack rope, scattered
    debris - is wrong, because every one of those is itself a human trace. The
    frame wants the place WITHOUT people: its landscape or its shell present,
    everything humanised absent. That part is mandatory.

    Second, and this one corrected an over-reach of mine: FROZEN IS NOT THE RULE.
    It was offered as one way to make a before-state distinctive, because the
    previous generation's scenes all looked alike. Freezing every scene just
    reproduces that sameness in a colder palette. So the before-state is chosen
    per scene to suit its own theme - `before_state` on the zone says how - and
    the requirement is that the five differ from each other.
    """
    given = (zone.get("before_state") or "").strip()
    if given:
        return ("Revise this image into its 'before' state: %s\n"
                "*** REMOVE EVERYTHING HUMANISED. *** No buildings, stalls, carts, furniture, "
                "goods, crates, tools, fences, paths, hanging decoration, lights, people or "
                "animals. Nothing built, nothing placed, nothing arranged, and NO DEBRIS OR RUINS "
                "either - broken boards and scattered litter are human traces too and must not "
                "appear. The place is bare and unused, not abandoned.\n"
                "KEEP the same camera and the same %s, so it is recognisably this place. "
                "Almost EMPTY is the target. Keep the same soft-painterly finish. "
                "No text anywhere." % (given, "room shell - walls, floor, ceiling, window"
                                       if zone.get("interior") else
                                       "landscape - fields, hills, mountains, horizon"))
    if zone.get("interior"):
        # An interior has no fields to freeze. The equivalent is the SHELL kept -
        # walls, floor, window, beams, the thing that makes it this location - and
        # every trace of use stripped out, under cold light. Same idea as the
        # landscape version: the place exists, nobody has worked in it.
        return ("Revise this image into its FROZEN 'before' state - this room in deep winter, "
                "unused and shut up.\n"
                "KEEP the shell of the room exactly as it is: the same walls, floor, ceiling and "
                "beams, the same window and door in the same places, the same camera. The room is "
                "simply COLD - frost on the glass, pale blue-white light, breath-cold air, the "
                "colour drained out of the timber.\n"
                "*** REMOVE EVERYTHING THAT SHOWS IT IS USED. *** No goods, no produce, no crates "
                "or barrels, no furniture, no counters or shelves of stock, no tools, no hanging "
                "decoration, no lights lit, no people and no animals. Nothing placed, nothing "
                "stored, nothing arranged, and NO DEBRIS, dust sheets or ruin either - clutter is "
                "a human trace too and must not appear.\n"
                "Almost EMPTY is the target: a bare cold room and the light through the frosted "
                "window, nothing else. Keep the same soft-painterly finish. No text anywhere.")
    return ("Revise this image into its FROZEN 'before' state - the land in deep winter, before "
            "anyone has worked it.\n"
            "KEEP the natural landscape exactly where it is: the same fields, the same hills, the "
            "same mountains or horizon, the same camera. They are simply frozen - held under snow "
            "and frost, still and silent, the colour drained to pale blue-white and cold grey.\n"
            "*** REMOVE EVERYTHING HUMANISED. *** No buildings, no stalls, no carts, no fences or "
            "posts, no paths or paving, no tools, no crates, no bunting, no lights, no people and "
            "no animals. Nothing built, nothing placed, nothing repaired, and NO DEBRIS OR RUINS "
            "either - broken boards and scattered litter are human traces too and must not appear. "
            "The land is bare and untouched, not abandoned.\n"
            "Almost EMPTY is the target: quiet frozen ground, quiet frozen distance, nothing else. "
            "Keep the same soft-painterly finish. No text anywhere.")


def _who(c):
    """Describe a character fully enough to draw it.

    base_character + costume is the house convention - the animal is already
    designed and only the costume is new - and wins when both are present. When
    they are absent, design_notes carries the species and the costume in prose and
    is used instead. Without that fallback this returned "a cute stylized
    anthropomorphic character in this event's costume", and the arrival frames
    rendered pale doughy figures with no species while the document plainly said
    "a round marmot in a rust-red dungaree and a straw hat".
    """
    base = (c.get("base_character") or "").strip()
    costume = (c.get("costume") or "").strip()
    if base:
        return ("a cute stylized anthropomorphic %s in %s"
                % (base, costume or "this event's costume"))
    notes = (c.get("design_notes") or "").strip()
    if notes:
        return "a cute stylized anthropomorphic character - %s" % notes
    return "a cute stylized anthropomorphic character in this event's costume"


def _who_short(c):
    """Name a character just well enough to say "leave that one alone".

    Held characters are already drawn in the source frame. Re-describing them in
    full is an invitation to redraw them, which is the opposite of holding them.
    """
    for key in ("name", "base_character"):
        v = (c.get(key) or "").strip()
        if v:
            return v
    return "the character already there"


CHARACTER_SCALES = ("beside", "foreground")


def add_characters(characters, arriving=None, scale="beside"):
    """Characters arrive ONE BY ONE across the last two frames.

    The art director's note: not a single character frame. One character in the
    penultimate frame, the second joining in the last, so the population reads as
    arriving rather than appearing all at once.

    `arriving` names who is new in THIS frame; anyone else in `characters` is
    already standing there and must be held unchanged.
    """
    arriving = arriving if arriving is not None else characters
    already = [c for c in characters if c not in arriving]
    new = " and ".join(_who(c) for c in arriving)
    held = ("The character%s already in the picture - %s - stay%s exactly where they are, "
            "unchanged. " % ("s" if len(already) > 1 else "",
                             " and ".join(_who_short(c) for c in already),
                             "" if len(already) > 1 else "s") if already else "")
    # Two scales, because they trade against each other and both are wanted.
    #
    # "beside" is the default and composites cleanly: the characters stand
    # lateral to the hero at the scene's own scale, and the frame stays
    # faithful to the completed base.
    #
    # "foreground" follows the style guide's camera section - character-centric
    # framing, large elements close to camera - and the faces read far better.
    # It costs fidelity: measured across five scenes, 67-95% of image columns
    # change, because the model re-renders the frame around the character
    # instead of compositing one into it. Composition survives; pixels do not.
    if scale == "foreground":
        place = ("*** THEY ARE BIG AND CLOSE TO CAMERA ***\n"
                 "They stand in the FOREGROUND, on the near side of the scene, well in front of "
                 "the hero object - not back beside it. Each one is ABOUT A THIRD OF THE FRAME'S "
                 "HEIGHT, the largest thing in the picture after the hero, close enough that the "
                 "face reads instantly. They are the subject of this frame, not dressing in it. A "
                 "character that reads as a small background figure is wrong.\n"
                 "Place them to ONE SIDE of the centre line so the hero and the built scene behind "
                 "them stay visible - overlapping the lower corner of the frame is fine, hiding "
                 "the hero is not.\n")
    else:
        place = ("They stand on the ground BESIDE the central hero object, lateral to it so "
                 "nothing behind them is hidden, at the scene's own scale - part of the picture "
                 "rather than in front of it.\n")
    return ("Revise this image by adding %d character%s: %s.\n"
            "%s"
            "%sLarge rounded heads, compact pear-shaped bodies, short simplified limbs, mitten "
            "hands, oversized expressive eyes and broad friendly expressions readable at thumbnail "
            "size.\n"
            "Adding %s is the ONLY change. %s"
            % (len(arriving), "" if len(arriving) == 1 else "s", new, place, held,
               "that character" if len(arriving) == 1 else "those characters",
               HOLD.replace("do NOT add anything, ", "")))


# Measured off the shipped Sort2603 and Sort2606 goods sets - 76 icons across six
# folders - rather than chosen:
#
#     canvas                 96x192, a 1:2 PORTRAIT
#     bbox width  / canvas   0.86 mean   (0.60 - 1.00)
#     bbox height / canvas   0.90 mean   (0.74 - 0.99)
#     CONTENT ASPECT w/h     0.48 mean   (0.31 - 0.62)
#     canvas area covered    0.55 mean
#
# The third line is the one that matters and the one three earlier attempts here
# missed. Every shipped good is ITSELF a portrait object at roughly 1:2, which is
# why it fills both axes and why the frame never looks empty. An earlier set here
# measured 0.81 mean aspect and covered 0.27 of the canvas - half the shipped
# figure - because its objects were square or wide. A harmonica at aspect 1.85
# cannot fill a 1:2 frame at any scale factor, so the fix is never a bigger
# number: scaling it up only makes it wider and leaves the same band of dead
# space above and below.
#
# So the requirement moved from the post-processing into the brief. Each good is
# chosen and presented so that the OBJECT is about twice as tall as it is wide -
# stood on its end, leaning, hung, or simply a tool that is already long. That is
# how the shipped set does it, and it is why there is one fill number here now
# instead of a ladder of size classes: the variation in how much ink lands on the
# canvas should come from the object (a whisk is mostly air, a sack is solid),
# never from deliberately shrinking things.
GOODS_FILL = 0.90

# What a rendered good has to come back as. Outside these it is re-rendered, not
# rescaled - stretching an object to fit is what makes an icon look wrong.
GOODS_ASPECT_RANGE = (0.33, 0.68)
GOODS_MIN_AREA = 0.30

# Measured across the shipped goods: WITHIN one floor the hues run right round the
# wheel (Sort2603/Goods3 spans 23 to 332 degrees), mean pairwise separation 76-94
# degrees, only 5-16% of pairs within 15 degrees. A shipped floor is emphatically
# not one palette. Handing every good on a floor the same palette string put one
# generated floor's entire catalogue inside 29 degrees of hue.
#
# The same measurement also caught the finish being off: shipped goods average
# 0.71 saturation and 0.81 value, against 0.52 and 0.55 for a set generated
# without asking. They are brighter and cleaner than a plain "rich colour" brief
# produces, and that brightness is what makes fifteen different hues still read
# as one shelf.
COLOUR_RULE = """

*** ITS COLOUR IS ITS OWN ***
This good's colour is: %s
The other goods on this shelf are deliberately DIFFERENT colours from it and from each other - the
player sorts by colour as well as by shape, so a shelf of one colour is a shelf of one item repeated.
Do not drift this object toward the colours of the room it sits in; give it the colour named above,
strongly and unmistakably, as the colour a person would use to describe it.
What makes these fifteen a set is the FINISH, not the hue: the same clean gloss, the same soft
rounded shading, the same dark keyline. Colour them BRIGHT and CLEAN - strongly saturated, light in
tone, like fresh painted toys under good light. Never dull, never greyed, never dark and muddy.

*** HOW LIGHT IT IS - THE WHOLE OBJECT IS HIGH-KEY ***
Light the object brightly and evenly, as if photographed under a big soft light in a white room.
EVEN THE SHADOW SIDE STAYS LIGHT: the darkest part of the object is a soft mid-tone, never a deep
shadow, never near-black. Shading works by dropping a little brightness and adding a little colour,
not by going dark. Keep the highlights broad and clean. The only genuinely dark thing in the picture
is the thin keyline around the edge.
Aim for the object to read as a LIGHT, luminous colour overall - a pale-to-mid version of its hue -
rather than a deep or dark one. "Deep", "rich", "dark" and "moody" are all wrong here."""

PORTRAIT_RULE = """

*** THE FRAME IS TALL - THE OBJECT MUST BE TALL TOO ***
This icon is drawn on a TALL PORTRAIT frame, twice as high as it is wide, and the object has to be
about TWICE AS TALL AS IT IS WIDE so that it fills that frame properly. A short or wide object
leaves a dead band of empty space above and below it and reads as a tiny icon lost in its box.
Get the height from the OBJECT, never from the drawing:
  - stand it on its end, upright, the way it would sit on a shelf;
  - or lean it, or hang it, or tilt it so its length runs UP the frame;
  - or show the whole of it including the handle, the stem, the stand it sits on.
NEVER stretch, squash or distort it to reach the top and bottom, and never blow up a small object
until it goes soft. It must look like the real thing, in its right proportions, stood so that its
long axis runs up the picture."""


def goods_prompt(item, palette="", siblings=()):
    """One Sort goods icon.

    `siblings` is the rest of THIS floor's catalogue. Passing it matters: the
    generator sees one item at a time and cannot avoid a collision it does not
    know about - the same reason a Merge scene had to be told its own building
    list before it would draw all of them.
    """
    sil = (item.get("silhouette") or "").strip()

    extra = """

*** IT IS A SHELF GOOD ***
A single real object from this workshop, as the player sorts it onto a shelf. One whole thing on its
own - not a scene, not a shelf, not a group, no second copy. It is a REAL tool or product of this
craft, the sort a working shop would actually stock, never a generic blob with the theme's colours
painted on."""

    extra += PORTRAIT_RULE

    colour = (item.get("palette") or palette or "").strip()
    if colour:
        extra += COLOUR_RULE % colour

    if sil:
        extra += """

*** ITS SILHOUETTE - THIS IS THE POINT OF THE IMAGE ***
Its outline is: %s
Commit to that outline completely. The player picks this item off a shelf of similar items by its
SHAPE ALONE at thumbnail size, so the shape has to be unmistakable before any detail is read.
It should reach close to the top and the bottom of the frame, and most of the way across.""" % sil

    sibs = [x for x in siblings if x]
    if sibs:
        extra += ("""

*** IT MUST NOT LOOK LIKE ANY OF THESE ***
These are the other goods on this same shelf, and this one is sorted against them:
    %s
Its outline must be tellable from every one of them at a glance. If what you are drawing would
share an outline with any item on that list, change the shape - not the colour, not the label,
the SHAPE.""" % ", ".join(sibs))

    # This used to ask the generator for the dark keyline, on the reasoning that
    # a shelf good is outlined where a Sort floor is lineless. But icon_sheet
    # then paints a keyline too, so every good came back wearing two - a drawn
    # contour of a few pixels under a 6px painted rim. Measured on the output,
    # peeling the painted rim lifted the 5th-percentile brightness from 0.24 to
    # 0.38 and peeling the drawn contour as well lifted it to 0.56, against a
    # shipped 0.52. The frames were not shaded dark; they were outlined twice.
    #
    # So goods now do what every other icon class does: the generator leaves the
    # edge clean and add_stroke supplies one keyline at a controlled width. That
    # is also what Mattel's own note asks for - "add a 1-2 pixel stroke" - rather
    # than whatever weight a generator happens to draw.
    extra += """

*** NO CONTOUR SHADING AT THE EDGE ***
Do not darken the object around its own outline. No dark rim, no inked contour, no band of heavy
ambient occlusion following the silhouette, no vignette. The edge of the object is the same
brightness as the surface just inside it. The keyline is added afterwards and must not be drawn
here - drawn as well as added, it doubles up and turns the whole outer edge of the icon dark."""

    out = _object_prompt(item.get("description") or item.get("name", "a shelf good"),
                         item.get("palette") or palette, extra)

    # The shared style block and the injected render spec both ask for DEEP
    # colour. That is right for a room and wrong for a shelf icon: measured, the
    # shipped goods sit at 0.81 value where goods generated under the deep
    # wording came back at 0.51. Adding a sentence asking for brightness changed
    # nothing, because three other sentences were still asking for the opposite
    # and the model is free to pick whichever it likes. So the conflicting words
    # are REPLACED rather than argued with - the same lesson as the render spec
    # that quoted a banned instruction and put it back into the prompt.
    for a, b in (("deep saturated colour",
                  "bright, clean, high-key saturated colour"),
                 # TRIED AND REVERTED. Side by side the shipped goods look
                 # flatter than the shared icon style asks for - measured tonal
                 # spread 0.43 against 0.70 here - so this replaced "smooth
                 # airbrushed gradients with a wide tonal range" with bold nearly
                 # flat colour and a narrow range. Rendered and measured, it made
                 # every figure slightly WORSE (p25 0.69 -> 0.60, saturation 0.60
                 # -> 0.57) and two of fifteen icons came back as broken
                 # fragments. The wide-range wording stays. The tonal gap that
                 # remains is not caused by it.
                 ("Rich and appetising, never washed out or greyed.",
                  "Bright, clean and light in tone - the colour of a freshly painted toy "
                  "under good light. Never washed out or greyed, and never dark, deep "
                  "or muddy.")):
        out = out.replace(a, b)
    return out


def goods_fraction(item, default=GOODS_FILL):
    """One fill for every good, matching the shipped 0.86 wide / 0.90 tall.

    There is deliberately no per-item size any more. Sizing individual goods down
    is what produced icons covering 6% of their canvas, and the honest way to
    make an icon fill its frame is to draw an object that fits it.
    """
    return GOODS_FILL




MATERIAL_EXTRA = """

*** IT IS A CONSTRUCTION MATERIAL ***
A raw building material as a game resource icon: a neat, chunky, self-contained bundle or piece - a
stack, a bundle, a single worked block - not a finished structure and not a tool in use. It reads at a
glance as "the stuff you build with"."""


def material_prompt(material, palette=""):
    """One Blocks construction-material icon.

    Blocks' mechanic-specific asset class, the way recipes are Dessert's and
    chains are Merge's: each construction step consumes 2-3 themed material
    icons, so they are a set and share a palette.
    """
    return _object_prompt(material.get("description") or material.get("name", "a material"),
                          material.get("palette") or palette, MATERIAL_EXTRA)


# Section 20 of Blocks Art.docx, the studio's own character formula. Used
# verbatim because it is the studio's language for its own cast, and because
# sections 9-14 of that document are entirely about character construction -
# head 40-55% of visual mass, oval/pear body, mitten hands, thumbnail-readable
# face - which no generic prompt of mine would have got right.
CHARACTER_FORMULA = (
    "premium casual game character design, large rounded head, compact pear-shaped body, short "
    "simplified limbs, oversized expressive eyes, soft rounded facial features, exaggerated friendly "
    "expression, simplified anatomy, toy-like proportions, smooth organic shapes, soft matte "
    "materials, subtle painterly 3D shading, broad soft highlights, gentle ambient occlusion, "
    "colorful oversized accessories, highly readable silhouette, charming playful personality, "
    "polished stylized mobile game artwork"
)


def character_prompt(character, palette=""):
    """One event character: the EXISTING cast member in this event's costume.

    The house convention across Merge, Dessert and Blocks is that a character
    art requirement is the COSTUME - the animal underneath is already drawn and
    is reused. So the base animal is stated as a constraint, not invented here.
    """
    base = (character.get("base_character") or "").strip()
    # design_notes is the fallback because it is the field these documents
    # actually carry. Without it the subject line came out as "Cute stylized
    # anthropomorphic character: in this event's costume" - no species, no
    # costume, nothing of the marmot in a rust-red waistcoat that the document
    # describes - and the renders were faithful to exactly that.
    costume = ((character.get("costume") or "").strip()
               or (character.get("design_notes") or "").strip())
    actions = [a for a in (character.get("scene_actions") or []) if str(a).strip()]
    who = ("Cute stylized anthropomorphic %s" % base) if base else "Cute stylized anthropomorphic character"
    reuse = ("\n\n*** THIS IS AN EXISTING CHARACTER IN A NEW COSTUME ***\n"
             "The %s itself is already designed and must stay recognisably the same animal - same "
             "species, same friendly proportions. What is new is only the costume: %s. Draw the "
             "costume as the event's contribution, not a redesign of the character."
             % (base, costume)) if base and costume else ""
    pose = ("\n\nPose: standing, facing the viewer, full body visible head to feet, weight on both "
            "feet, arms relaxed and clear of the body so the silhouette reads."
            + (" In character: %s." % actions[0] if actions else ""))
    return """*** A SINGLE GAME CHARACTER ***
{who}: {costume}
{formula}

Draw ONE character only, centred, filling the frame top to bottom. Nothing else is in the picture - no
scene, no furniture, no props on the ground, no second character.

{backdrop}

*** FACE ***
The face must communicate its emotion at THUMBNAIL SIZE. Eyes large, rounded, set well apart, each with
a dark pupil, a lighter iris and a bright highlight. The mouth is exaggerated and may take a large share
of the face. The expression is broad and clear - no subtle realistic acting.

*** PROPORTIONS ***
The head is roughly 40-55% of the character's total visual mass. Body oval, pear or bean shaped. Limbs
short and simplified; hands as mittens or rounded paws; legs very short or partly hidden.
NOT athletic anatomy, NOT realistic muscles, NOT long limbs, NOT a narrow face, NOT anime proportions.

*** ACCESSORIES ***
Large and readable from a distance, in the same rounded shape language. No tiny buckles, no stitching,
no micro-mechanical detail.

*** COLOUR ***
Built from {palette}. Warm key light, cool ambient shadow, a subtle rim light around the silhouette.

{no_text}{reuse}{pose}""".format(
        who=who, costume=costume or "in this event's costume", formula=CHARACTER_FORMULA,
        backdrop=ICON_BACKDROP, palette=palette or "the event palette",
        no_text=ICON_NO_TEXT, reuse=reuse, pose=pose)


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
