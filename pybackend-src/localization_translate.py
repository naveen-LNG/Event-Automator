"""
Translate English UI strings into the six shipped languages, using what the
third-party translator has already taught us.

Nothing here is a model fine-tune. The "learning" is four concrete layers,
cheapest and most certain first, so the model is only ever asked the questions
nothing else can answer:

  1. APPROVED    the vendor's own corrected text for this exact English string
                 (loc_corrections_import.py writes these). Reused verbatim.
  2. MEMORY      our own shipped translation for this exact English string, out
                 of localization.xlsx + localization_quiz.xlsx. Reused verbatim.
  3. GENERATED   whatever is left goes to Claude, prompted with the nearest
                 memory neighbours and the vendor's real before/after pairs for
                 that language.
  4. CHECKED     every generated string is then put through the deterministic
                 rules below - the ones the vendor kept applying by hand.

Layer 4 exists because of a measurement, not a hunch: of 543 real corrections
in the first nine deliveries, 140 (26%) changed nothing but typography, casing
or punctuation. All 108 of the typographic ones did the same thing - replaced a
straight ' quote pair with the language's own quotation marks. That is a rule,
not a judgement call, so QUOTE_STYLE below applies it and the model is never
asked to get it right.

Every constant in this file was measured from the shipped corpus rather than
recalled from language references - see the comments on QUOTE_STYLE and
MAX_LENGTH_RATIO for what was counted and over how many rows.

Payload in (argv[1] is a JSON file), one JSON line out. See translate_main.
"""
import collections
import difflib
import hashlib
import json
import os
import re
import sys
import unicodedata

import openpyxl

from ideation_generator import _call_claude, _call_ollama, ollama_timeout_for

LANGS = ["DE", "ES", "NL", "SV", "FR", "DA"]
LANG_NAMES = {"DE": "German", "ES": "Spanish", "NL": "Dutch",
              "SV": "Swedish", "FR": "French", "DA": "Danish"}

# (opening, closing, padding inside the marks).
#
# Measured, not looked up: every one of the 108 typographic corrections in the
# first nine deliveries replaced a straight ' pair, and each language was
# unanimous about what it became - DE 27/27, ES 28/28, NL 12/12, SV 11/11,
# DA 10/11, FR 17/19. French was the only one that pads, and it pads with an
# ordinary U+0020 space (28 of 28 guillemets in the corrected French text), not
# the narrow no-break space typography guides ask for. We follow the vendor,
# because the vendor is who signs off.
QUOTE_STYLE = {
    "DE": ("„", "“", ""),
    "ES": ("«", "»", ""),
    "NL": ("‘", "’", ""),
    "SV": ("”", "”", ""),
    "FR": ("«", "»", " "),
    "DA": ("”", "”", ""),
}

# A translation longer than 99% of our shipped translations for that language
# is worth a designer's eye - it is the shape UI overflow takes. Derived from
# the 31,396 English->target pairs per language already in the workbooks
# (character ratio against the English source, strings of 4+ chars).
# These are calibration knobs: re-derive them when the corpus grows a lot.
MAX_LENGTH_RATIO = {"DE": 2.00, "ES": 2.08, "NL": 1.89,
                    "SV": 1.75, "FR": 2.12, "DA": 1.75}

# What must survive translation byte-for-byte. The corpus has 1,198 English
# strings carrying at least one of these, and a dropped %s is a runtime crash
# rather than a cosmetic problem.
PLACEHOLDER_RE = re.compile(r"%\d+\$[sd]|%[sd]|\{\d+\}|\{[A-Za-z_][A-Za-z0-9_]*\}|\\n|<[^<>]{1,24}>")

TM_FILES = ["localization.xlsx", "localization_quiz.xlsx"]
SKIP_SHEETS = {"languages", "country"}

# Parsing both workbooks costs ~5.1s; reading the snapshot back costs ~0.05s.
# The panel is interactive, so the snapshot is not optional.
SNAPSHOT_VERSION = 2   # bumped when tidy() stopped flattening newlines

MAX_BATCH = 20          # English strings per model request
MAX_NEIGHBOURS = 3      # memory examples shown per string
MAX_STYLE_EXAMPLES = 6  # vendor before/after pairs shown per language

# A button label and a line of dialogue are different writing jobs, and the
# style examples have to cover both. Measured on the 09-17 review: 46% of that
# delivery was 30 characters or under, and the old "longest English first"
# selection put ZERO short strings in front of the model - 32 of the 36 chosen
# examples were trivia questions, because questions are the longest strings in
# the corrections corpus. See load_style_examples.
SHORT_STRING = 30

# A trivia question, as opposed to game copy. The quiz workbook supplies most
# of the corrections corpus, so without this the style examples are all
# general-knowledge questions.
TRIVIA_RE = re.compile(r"(which|what|who|where|when|how)\b", re.I)

# Per-language house rules.
#
# Every line below was derived from the third-party review of the Candy
# Crafters delivery (178 corrections on 822 cells), not from general style
# advice, and each was checked for direction before being written down:
#   - 21 of the 24 casing corrections were the vendor LOWERCASING our text,
#     18 of them French. We were calquing English title case.
#   - German and Spanish dialogue was rewritten formal -> informal (5 and 1).
#   - French moved person in NEITHER direction across 7 corrections and its
#     final text is 23 formal markers to 0, so French keeps vous. Writing an
#     informal rule for French would have made the tool worse.
#   - Dutch and Danish noun phrases were closed up into single compounds.
LANG_RULES = {
    "DE": [
        "Address the player informally: du, dir, dein. Never Sie/Ihnen/Ihr.",
        "Capitalise only what German capitalises anyway (nouns). Do not copy "
        "English title case onto adjectives.",
    ],
    "ES": [
        "Address the player informally: tú. Never usted.",
        "Do not stack nouns the English way - link them with de/del/de la. "
        "\"Cocoa Workshop Pack\" is \"Paquete del Taller de Cacao\", not "
        "\"Paquete Taller de Cacao\".",
    ],
    "NL": [
        "Write compounds as ONE word: \"Cacaowerkplaatspakket\", never "
        "\"Cacao Werkplaats Pakket\". Hyphenate only after an English loan "
        "word, as in \"Chocolate Smash-game\".",
        "Address the player informally: je, jouw.",
    ],
    "SV": [
        "Write compounds as ONE word: \"Kakaoverkstadspaket\".",
        "Capitalise only the first word and proper nouns: "
        "\"Godismakarens kortspel\", not \"Godismakarens Kortspel\".",
    ],
    "FR": [
        "Address the player formally: vous. This game uses vous throughout - "
        "do not switch to tu.",
        "Capitalise only the first word and proper nouns. English title case "
        "is wrong: \"Étiquettes de voyage\", not \"Étiquettes de Voyage\". "
        "Common game nouns stay lowercase in running text: chocolaterie, "
        "jetons de cacao, pièces sucrées, pochettes de chocolat.",
        "Button and menu labels take the infinitive, not the imperative: "
        "\"Améliorer\", not \"Améliorez\".",
    ],
    "DA": [
        "Write compounds as ONE word, with the linking -s- Danish requires: "
        "\"Kakaoværkstedspakke\", not \"Kakaoværksted-pakke\".",
        "Common nouns stay lowercase in running text: \"chokoladeposer\", "
        "not \"Chokoladeposer\".",
    ],
}

# Above this many gaps, show one neighbour per language instead of three.
#
# Neighbours are per string per language, so the brief grows as strings x 6 x
# examples. A real 124-string batch of varied event copy produced a 188,347
# character brief - fine as a file, too big to hand a chat tool. Tapering keeps
# small batches (the common case) at full quality and stops a large one from
# being unusable. Measured effect is in the comment on brief().
TAPER_ABOVE_GAPS = 40
TAPER_NEIGHBOURS = 1

UNIFY = [("’", "'"), ("‘", "'"), ("“", '"'), ("”", '"'),
         (" ", " "), ("​", ""), ("–", "-"), ("—", "-")]


# ---------------- normalisation (must match loc_corrections_import) ----------------

def tidy(value):
    r"""Clean up a value for storage and display WITHOUT destroying line breaks.

    This replaced a collapse() that ran re.sub(r"\s+", " ", ...) over every
    value. That flattened real newlines into spaces, so a multi-line string -
    162 cells in localization.xlsx, 29 of them English - came back out of the
    panel as a DIFFERENT string than the one that went in. Pasting that into
    the workbook silently rewrote the string with nothing to catch it.

    Runs of spaces and tabs still collapse, and each line is trimmed, so the
    tidying that was wanted survives. Blank lines survive too: a paragraph
    break is content, not stray whitespace.

    Matching is unaffected - norm_key() does its own aggressive normalisation,
    which is what lets a lookup still succeed across a newline difference."""
    text = str(value).replace("\r\n", "\n").replace("\r", "\n")
    text = re.sub(r"[^\S\n]+", " ", text)     # spaces and tabs, never newlines
    return "\n".join(line.strip() for line in text.split("\n")).strip()


def norm_key(value):
    text = unicodedata.normalize("NFKC", str(value))
    for a, b in UNIFY:
        text = text.replace(a, b)
    return re.sub(r"\s+", " ", text).strip().lower().rstrip(".!?:;,")


# ---------------- layer 1: the vendor's approved text ----------------

def load_approved(corrections_root):
    """lang -> norm_key(english) -> entry, as written by the importer.

    Missing folder is not an error: the panel has to work before anyone has
    imported a delivery, it just has one fewer memory to draw on."""
    approved = {lang: {} for lang in LANGS}
    if not corrections_root:
        return approved
    for lang in LANGS:
        path = os.path.join(corrections_root, "approved", lang + ".json")
        if not os.path.exists(path):
            continue
        try:
            with open(path, "r", encoding="utf-8-sig") as handle:
                approved[lang] = json.load(handle)
        except (OSError, ValueError):
            # A half-synced OneDrive file must not take down a translation run.
            pass
    return approved


def load_style_examples(corrections_root):
    """lang -> the vendor's wording corrections, newest-looking first.

    Only 'wording' pairs are shown to the model. Feeding it the typographic
    ones would be teaching it a rule we already apply deterministically, and
    spending prompt on it twice."""
    examples = {lang: [] for lang in LANGS}
    if not corrections_root:
        return examples
    for lang in LANGS:
        path = os.path.join(corrections_root, "pairs", lang + ".json")
        if not os.path.exists(path):
            continue
        try:
            with open(path, "r", encoding="utf-8-sig") as handle:
                rows = json.load(handle)
        except (OSError, ValueError):
            continue
        wording = [r for r in rows
                   if r.get("category") == "wording" and not r.get("ambiguous")]
        # Longest English first WITHIN each length bucket: a long pair carries
        # the most transferable context, so it is still the right tie-break.
        #
        # But ranking on length alone was a trap. Questions are the longest
        # strings in the corrections corpus, so "longest first" handed the
        # model six trivia questions per language (32 of 36, and 6 of 6 for
        # ES/NL/DA) and not one button label - while 46% of the delivery it
        # was about to write consisted of strings 30 characters or shorter.
        # Splitting the budget between short and long guarantees the model
        # sees the register of both jobs.
        # Two axes, because length alone was not enough. Most of the
        # corrections corpus is quiz content (the trivia workbook draws the
        # most corrections), and after the length split alone 78% of the
        # chosen examples were still trivia questions. Deprioritising them
        # leaves the UI and dialogue copy the model is actually being asked
        # to write. Note the test is the interrogative opener, not a trailing
        # "?" - "Revert your previous move?" is a confirm dialog, exactly the
        # register we want to keep.
        wording.sort(key=lambda r: (bool(TRIVIA_RE.match((r.get("en") or "").strip())),
                                    -len(r.get("en") or "")))
        short = [r for r in wording if len(r.get("en") or "") <= SHORT_STRING]
        long_ = [r for r in wording if len(r.get("en") or "") > SHORT_STRING]
        half = MAX_STYLE_EXAMPLES // 2
        picked = short[:half] + long_[:MAX_STYLE_EXAMPLES - half]
        # If one bucket is short of stock, spend the remainder on the other.
        if len(picked) < MAX_STYLE_EXAMPLES:
            spare = [r for r in short[half:] + long_[MAX_STYLE_EXAMPLES - half:]]
            picked += spare[:MAX_STYLE_EXAMPLES - len(picked)]
        examples[lang] = picked
    return examples


# ---------------- the glossary: terms that must not drift ----------------
# Why this exists: on the 09-17 review the tool rendered "Manager" as both
# "Gérant" (x2) and "Manager" (x4) in French inside a single delivery, and
# "Chef" (x1) / "Manager" (x5) in Danish. No model improvement fixes that -
# the decision has to be made once and then held, which is what a glossary is.

GLOSSARY_FILE = "glossary.json"

# Unicode-aware: the terms this has to match are Süßmünze, Sötmynt, Slikmønt,
# Fábrica and pièce sucrée. The ASCII WORD_RE used by retrieval would shred
# every one of them.
WORD_ANY_RE = re.compile(r"[^\W\d_]+", re.UNICODE)


def load_glossary(corrections_root):
    """english term -> {lang: house rendering}. Keys starting with _ are notes."""
    if not corrections_root:
        return {}
    path = os.path.join(corrections_root, GLOSSARY_FILE)
    try:
        with open(path, "r", encoding="utf-8-sig") as handle:
            raw = json.load(handle)
    except (OSError, ValueError):
        return {}
    return {k: v for k, v in raw.items()
            if not k.startswith("_") and isinstance(v, dict)}


def glossary_hits(english, glossary):
    """The glossary entries whose English term appears in this string.

    Longest term first, so "Chocolate Factory" is reported rather than a bare
    "chocolatier" that happens to sit inside it."""
    text = str(english).lower()
    hits = [(term, langs) for term, langs in glossary.items()
            if term.lower() in text]
    hits.sort(key=lambda pair: -len(pair[0]))
    return hits


def _stem(word):
    """Enough of a word to survive inflection without matching its neighbours.

    Dropping two characters absorbs plurals and case endings (Süßmünze ->
    Süßmünzen, Slikmønt -> Slikmønter, pièce sucrée -> pièces sucrées) while
    keeping "chokolademager" from being satisfied by plain "chokolade"."""
    return word[:max(4, len(word) - 2)].lower()


def glossary_warning(english, translated, lang, glossary):
    """Which fixed term is missing from this translation, or None.

    Case is deliberately ignored: whether "Fábrica" or "fábrica" is right in
    a given sentence is a LANG_RULES question, and flagging it here would put
    two rules on one string and teach people to stop reading warnings."""
    got = [w.lower() for w in WORD_ANY_RE.findall(str(translated))]
    for term, langs in glossary_hits(english, glossary):
        want = langs.get(lang)
        if not want:
            continue
        # Every word of the wanted term that carries meaning must show up.
        needed = [w for w in WORD_ANY_RE.findall(want) if len(w) > 3]
        if needed and not any(g.startswith(_stem(w)) for w in needed for g in got):
            return ("glossary: %r should be %r in %s"
                    % (term, want, LANG_NAMES.get(lang, lang)))
    return None


# ---------------- layer 2: our own shipped translations ----------------

def _fingerprint(design_path):
    parts = [str(SNAPSHOT_VERSION)]
    for name in TM_FILES:
        path = os.path.join(design_path, name)
        if os.path.exists(path):
            stat = os.stat(path)
            parts.append("%s:%d:%d" % (name, stat.st_mtime_ns, stat.st_size))
        else:
            parts.append(name + ":absent")
    digest = hashlib.sha1("|".join(parts).encode("utf-8")).hexdigest()[:16]
    return digest


def _build_snapshot(design_path):
    memory = {}
    used = []
    for name in TM_FILES:
        path = os.path.join(design_path, name)
        if not os.path.exists(path):
            continue
        used.append(name)
        workbook = openpyxl.load_workbook(path, read_only=True, data_only=True)
        for sheet_name in workbook.sheetnames:
            if sheet_name in SKIP_SHEETS:
                continue
            for row in workbook[sheet_name].iter_rows(min_row=2, values_only=True):
                if not row or not row[0] or str(row[0]).startswith("//"):
                    continue
                if len(row) < 2 or row[1] in (None, ""):
                    continue
                key = norm_key(row[1])
                entry = memory.setdefault(key, {"en": tidy(row[1])})
                for offset, lang in enumerate(LANGS, start=2):
                    value = row[offset] if len(row) > offset else None
                    # First writer wins, so the live sheets (read before the
                    # archive) beat an old event's pre-correction translation.
                    if value not in (None, "") and lang not in entry:
                        entry[lang] = tidy(value)
        workbook.close()
    return {"files": used, "entries": memory}


def tm_snapshot(design_path, cache_dir):
    """Our shipped translations, cached against both workbooks' mtime+size."""
    if not design_path or not os.path.isdir(design_path):
        return {"files": [], "entries": {}}, "no-design-path"
    name = "tm_%s.json" % _fingerprint(design_path)
    path = os.path.join(cache_dir, name) if cache_dir else None
    if path and os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as handle:
                return json.load(handle), "cache"
        except (OSError, ValueError):
            pass
    snapshot = _build_snapshot(design_path)
    if path:
        try:
            os.makedirs(cache_dir, exist_ok=True)
            # Written via a temp name so a concurrent reader never sees a
            # half-written snapshot on a synced folder.
            temp = path + ".tmp"
            with open(temp, "w", encoding="utf-8") as handle:
                json.dump(snapshot, handle, ensure_ascii=False)
            os.replace(temp, path)
        except OSError:
            pass
    return snapshot, "built"


# ---------------- layer 4: the deterministic rules ----------------

def apply_quote_style(text, lang):
    """Straight quote PAIRS become the language's own quotation marks.

    Pairs only. A lone apostrophe is left alone on purpose: it is usually
    word-internal (Dutch "Dorothy's", French "l'origine"), and the vendor is
    not consistent about converting those - 2 of 19 French corrections did,
    the rest did not. Converting them anyway would invent a rule the sign-off
    does not support."""
    style = QUOTE_STYLE.get(lang)
    if not style:
        return text
    open_mark, close_mark, pad = style

    def replace(match):
        inner = match.group(2)
        return "%s%s%s%s%s" % (match.group(1) or "", open_mark + pad,
                               inner, pad + close_mark, "")

    # A quoted run: an opening straight quote at a word boundary, some text
    # with no quote in it, then a closing straight quote.
    for quote in ("'", '"'):
        pattern = re.compile(r"(^|[\s(\[–—-])%s([^%s\n]{1,120})%s" % (quote, quote, quote))
        text = pattern.sub(replace, text)
    return text


def placeholders(text):
    return sorted(PLACEHOLDER_RE.findall(str(text)))


# ---- what the DesignData convert cannot carry -------------------------------
# c_data copies a localization cell into localize.json VERBATIM - it escapes
# nothing. So a straight " ends the JSON string early and breaks coder_convert
# for the whole team, and a backslash that does not open a valid JSON escape
# does the same. localization.xlsx's own convention for a quote you really do
# want is a typed \" (see TUTORIAL_10301); the quotation marks in QUOTE_STYLE
# are non-ASCII, so they are always safe.
#
# This is a build break, not a typography opinion, so unlike the advisory
# checks in row_warnings it runs on approved and shipped text too - vendor
# signed-off text with a stray " stops the convert exactly the same way.
# Real case: SIM2609_LETTER reached localize_event wrapped in quotes and failed
# coder_convert with "Expected comma or object end but found invalid token at
# character 1401996".
JSON_ESCAPE_OPENERS = set('"\\/bfnrtu')
HEX = set("0123456789abcdefABCDEF")


def pipeline_unsafe(text):
    """Why this string would break the DesignData JSON convert, or None."""
    text = str(text)
    i = 0
    while i < len(text):
        if text[i] == "\\":
            nxt = text[i + 1:i + 2]
            if nxt not in JSON_ESCAPE_OPENERS:
                return ("backslash before %s is not a valid JSON escape - the "
                        "convert copies cells into JSON as-is"
                        % (repr(nxt) if nxt else "the end of the text"))
            if nxt == "u" and not all(c in HEX for c in text[i + 2:i + 6]):
                return r"\u must be followed by 4 hex digits or the convert fails"
            i += 2
            continue
        if text[i] == '"':
            return ('straight " would end the JSON string in coder_convert - '
                    'delete it, or type it as \\"')
        i += 1
    return None


# An opening straight quote at a word boundary, then a matching closing one.
# The leading boundary is what keeps elision apostrophes out: the ' in "J'ai"
# is preceded by a letter, so it never opens a pair.
QUOTE_PAIR_RE = re.compile(u"(?:^|[\\s(\\[\u2013\u2014-])([\"'])(?=\\S)[^\"'\\n]{0,400}\\1")


def row_warnings(english, translated, lang, source="generated", glossary=None):
    """Everything mechanically checkable about one translated string.

    Two tiers, and the split matters. A placeholder mismatch is a correctness
    invariant - a dropped %s crashes the client no matter who wrote the string -
    so it is checked on everything, and so is a character the DesignData
    convert cannot carry (see pipeline_unsafe). The rest are advisory style
    checks and only apply to text WE just generated.

    Running the advisory checks on reused text was actively harmful: the
    translator's own signed-off Danish for "Redwood" (Kæmpe-sekvoja, 1.9x the
    English) got flagged for length. Second-guessing the authority we are
    deliberately deferring to is noise, and noise is what teaches people to
    stop reading warnings. Shipped text has the same immunity - it is already
    in the game, so it demonstrably fits.

    The glossary check sits in the advisory tier for that same reason: if the
    translator signed off wording that departs from the glossary, the glossary
    is what is out of date, not their text."""
    warnings = []
    if not str(translated).strip():
        return ["empty translation"]

    want, got = placeholders(english), placeholders(translated)
    if want != got:
        missing = [p for p in want if p not in got]
        extra = [p for p in got if p not in want]
        if missing:
            warnings.append("placeholder missing: " + " ".join(missing))
        if extra:
            warnings.append("placeholder added: " + " ".join(extra))

    unsafe = pipeline_unsafe(translated)
    if unsafe:
        warnings.append(unsafe)

    if source != "generated":
        return warnings

    source_len = len(tidy(english))
    if source_len >= 4:
        ratio = len(tidy(translated)) / source_len
        limit = MAX_LENGTH_RATIO.get(lang)
        if limit and ratio > limit:
            warnings.append("%.1fx longer than the English (over the %.2fx seen in "
                            "99%% of shipped %s) - check for UI overflow"
                            % (ratio, limit, LANG_NAMES.get(lang, lang)))

    # A quote PAIR that our own typography rule did not convert. It must be a
    # pair: `[\"']\S` used to fire on every French elision - l'Usine, d'Or,
    # J'ai, qu'ils - which was 15 of the 21 warnings on the 09-17 delivery and
    # the only check that fired at all. It also contradicted
    # apply_quote_style, which leaves word-internal apostrophes alone on
    # purpose. A stray straight double quote is a separate and worse problem;
    # pipeline_unsafe owns that one.
    open_mark, close_mark, _ = QUOTE_STYLE.get(lang, ("", "", ""))
    if open_mark and QUOTE_PAIR_RE.search(translated) and open_mark not in translated:
        warnings.append("straight quote left in - %s uses %s%s"
                        % (LANG_NAMES.get(lang, lang), open_mark, close_mark))

    if glossary:
        missing = glossary_warning(english, translated, lang, glossary)
        if missing:
            warnings.append(missing)

    if tidy(english) and tidy(translated) == tidy(english) and source_len > 12:
        warnings.append("identical to the English - check it was actually translated")
    return warnings


# ---------------- retrieval ----------------

WORD_RE = re.compile(r"[a-z0-9]+")

# A shortlist big enough that the ranking below still has real choices, small
# enough that difflib is cheap. 400 of 20,315 keys is a 50x cut.
CANDIDATE_POOL = 800


STEM = 5    # characters of a word that must agree for it to be a candidate


def _tokens(text):
    """Words, plus a truncated form of each.

    The stem is what makes the shortlist survive morphology. Indexing whole
    words only, "Chocolates" never met "Chocolate" and the shortlist returned
    nothing where the exhaustive scan returned five good neighbours - a silent
    quality loss traded for speed, which is the worst kind. Five characters is
    enough to collide plural/singular and most inflections without collapsing
    unrelated words together."""
    words = set(WORD_RE.findall(text))
    return words | {"~" + w[:STEM] for w in words if len(w) > STEM}


def build_index(keys):
    """word -> keys containing it, for shortlisting before difflib.

    Why this exists, having once been left out on purpose: difflib compares the
    probe against EVERY key, so its cost scales with corpus size times probe
    length. Measured on the real 20,315-key memory, one call was 0.035s for an
    11-character probe but 1.25s for a 91-character one - and real event copy is
    the long kind. At six languages that was 7.5s per string, which is what made
    a 124-string batch time out. The earlier "no index needed" call was made
    against a smaller corpus and short probes, and it did not survive contact
    with a real batch."""
    index = collections.defaultdict(list)
    for key in keys:
        for token in _tokens(key):
            index[token].append(key)
    return index


def similar_keys(english, keys, index, limit):
    """Nearest keys by English text. Deliberately language-INDEPENDENT.

    This used to be called once per language, which repeated the whole
    expensive scan six times to compute the same answer six times over. The
    result depends only on the English, so it is computed once per string and
    the per-language filtering happens after."""
    probe = norm_key(english)
    tokens = _tokens(probe)
    if not tokens:
        return []

    # Rarest tokens first: a shortlist built from "the" is no shortlist at all.
    # Every token still gets to contribute before the cap, so a probe whose
    # rare word is a red herring is not stuck with that word's postings.
    pool, seen = [], set()
    for token in sorted(tokens, key=lambda t: len(index.get(t, ()))):
        for key in index.get(token, ()):
            if key not in seen:
                seen.add(key)
                pool.append(key)
        if len(pool) >= CANDIDATE_POOL:
            break
    if not pool:
        pool = keys                       # nothing shared a token; fall back
    return difflib.get_close_matches(probe, pool[:CANDIDATE_POOL], n=limit, cutoff=0.55)


def neighbours_by_lang(english, entries, keys, index, langs, limit=MAX_NEIGHBOURS):
    """lang -> the nearest strings we already ship in that language."""
    found = {lang: [] for lang in langs}
    for candidate in similar_keys(english, keys, index, limit * 4):
        entry = entries.get(candidate) or {}
        for lang in langs:
            if len(found[lang]) < limit and entry.get(lang):
                found[lang].append({"en": entry.get("en", candidate), "text": entry[lang]})
    return found


# ---------------- the model call ----------------

SYSTEM_PROMPT = (
    "You translate short UI strings for Skip-Bo, a casual mobile card game, from "
    "English into the language you are asked for.\n\n"
    "Rules:\n"
    "- Match the register of the APPROVED EXAMPLES you are shown. They were signed "
    "off by the game's professional translator and they are the target you are "
    "aiming at, not suggestions.\n"
    "- Keep every placeholder (%s, %d, {0}, \\n, <b>) exactly as it appears in the "
    "English, in an order that reads naturally in the target language.\n"
    "- Keep emoji exactly as they appear.\n"
    "- These strings sit on buttons and in dialogs. Stay as short as the English, "
    "or shorter. Never explain, never add words the English does not have.\n"
    "- Follow the house rules listed under each language heading. They beat "
    "your own sense of what reads well, and they beat the English's own "
    "shape: English title case, English noun stacking and English compound "
    "spacing are wrong in most of these languages.\n"
    "- Where a string lists a fixed term, use that exact term. Do not vary it "
    "between strings, and do not translate it afresh because a different "
    "wording would fit the sentence better.\n"
    "- Product names stay in English: Skip-Bo, Wild Card, Stock Pile. Nothing "
    "else stays in English - a job title like Manager gets translated.\n"
    "- Do not add quotation marks the English does not have. Typography is applied "
    "afterwards by rule, so write plain straight quotes if you need any at all.\n"
    "- ALL CAPS English stays ALL CAPS in the translation.\n"
)


def _schema(langs):
    fields = {"en": {"type": "string"}}
    for lang in langs:
        fields[lang] = {"type": "string"}
    return {
        "type": "object",
        "properties": {
            "rows": {
                "type": "array",
                "items": {"type": "object", "properties": fields,
                          "required": ["en"] + list(langs), "additionalProperties": False},
            }
        },
        "required": ["rows"],
        "additionalProperties": False,
    }


def _user_message(batch, langs, snapshot, style_examples, glossary=None):
    entries = snapshot["entries"]
    keys = list(entries.keys())
    index = build_index(keys)
    glossary = glossary or {}
    parts = []

    for lang in langs:
        open_mark, close_mark, _ = QUOTE_STYLE.get(lang, ("", "", ""))
        parts.append("## %s (%s)" % (LANG_NAMES[lang], lang))
        # Rules before examples. These are the classes the translator kept
        # correcting by hand, and unlike the examples they apply to every
        # string in the batch rather than the ones that happen to look similar.
        for rule in LANG_RULES.get(lang, []):
            parts.append("- %s" % rule)
        examples = style_examples.get(lang) or []
        if examples:
            parts.append("Corrections the translator made to our earlier attempts - "
                         "these are the mistakes to avoid:")
            for pair in examples:
                parts.append("  EN: %s" % pair["en"])
                parts.append("    we wrote : %s" % pair["before"])
                parts.append("    they want: %s" % pair["after"])
        parts.append("")

    parts.append("## Strings to translate")
    for item in batch:
        parts.append("")
        parts.append("EN: %s" % item["en"])
        marks = placeholders(item["en"])
        if marks:
            parts.append("  placeholders that must survive: %s" % " ".join(marks))
        # Fixed terms, per string rather than in one list up top: a glossary
        # the model has to hold in mind across 20 strings is a glossary it
        # drifts away from, which is the failure this is here to stop.
        for term, renderings in glossary_hits(item["en"], glossary):
            shown = ", ".join("%s=%s" % (l, renderings[l])
                              for l in langs if renderings.get(l))
            if shown:
                parts.append("  fixed term %r - use exactly: %s" % (term, shown))
        # One scan per string, then filtered per language - the scan result
        # depends only on the English.
        near_by_lang = neighbours_by_lang(item["en"], entries, keys, index, langs,
                                          limit=(TAPER_NEIGHBOURS
                                                 if len(batch) > TAPER_ABOVE_GAPS
                                                 else MAX_NEIGHBOURS))
        for lang in langs:
            near = near_by_lang.get(lang) or []
            if near:
                shown = "; ".join('"%s" -> "%s"' % (n["en"], n["text"]) for n in near)
                parts.append("  similar %s we already ship: %s" % (lang, shown))

    parts.append("")
    parts.append("Return one row per English string above, echoing `en` exactly as "
                 "given, with a translation for: %s." % ", ".join(langs))
    return "\n".join(parts)


OLLAMA_PREFLIGHT_SECONDS = 2.0


def ollama_unreachable(base_url):
    """Why a pre-flight: a configured-but-offline Ollama box does not refuse the
    connection, it black-holes it, so the socket sits for ~20 seconds before
    WinError 10060. The office LAN server being down is the normal case on a
    laptop off the network, and 20 seconds of nothing before an error message
    reads as the panel being broken. Returns a reason, or None if reachable."""
    import socket
    from urllib.parse import urlparse

    parsed = urlparse(base_url if "//" in base_url else "//" + base_url)
    host = parsed.hostname
    port = parsed.port or (443 if parsed.scheme == "https" else 80)
    if not host:
        return "no host in the configured Ollama URL (%s)" % base_url
    try:
        socket.create_connection((host, port), timeout=OLLAMA_PREFLIGHT_SECONDS).close()
        return None
    except OSError as exc:
        return "%s is not reachable (%s)" % (base_url, exc.strerror or exc)


# ollama_timeout_for() floors at 600s, which is right for an Art Requirement
# and wrong here. This panel is interactive: a designer clicking Translate and
# getting nothing for ten minutes concludes it is broken, and on the measured
# studio box (~10 tok/s) a full batch genuinely takes that long. Cap it, and let
# the failure fall through to the Claude Desktop handoff, which is both faster
# and better. Raise it only if a local model ever becomes the preferred path.
OLLAMA_TRANSLATE_CEILING = 60


def _call_model(payload, system_prompt, user_message, schema, max_tokens):
    provider = (payload.get("provider") or "anthropic").lower()
    if provider == "ollama":
        return _call_ollama(
            payload.get("ollama_base_url", ""),
            payload.get("ollama_model") or "qwen3:32b",
            system_prompt, user_message, schema,
            timeout=min(ollama_timeout_for(max_tokens), OLLAMA_TRANSLATE_CEILING),
        )
    return _call_claude(payload.get("api_key", ""), system_prompt, user_message,
                        schema, max_tokens)


def progress(stage, done, total, detail=""):
    """One progress line on STDERR, because stdout is the JSON result channel.

    A large batch is several sequential model calls and can legitimately run for
    minutes. Without this the panel shows "Translating..." and nothing else, and
    a slow run is indistinguishable from a hung one - which is exactly what a
    designer reports as broken."""
    sys.stderr.write("PROGRESS %s %d/%d %s\n" % (stage, done, total, detail))
    sys.stderr.flush()


def generate(payload, gaps, langs, snapshot, style_examples, glossary=None):
    """norm_key(english) -> lang -> generated text, for the strings no memory
    could answer. Batched, because one request per string wastes the shared
    context that makes the batch consistent with itself."""
    produced = collections.defaultdict(dict)
    chunks = (len(gaps) + MAX_BATCH - 1) // MAX_BATCH
    for index, start in enumerate(range(0, len(gaps), MAX_BATCH), start=1):
        batch = gaps[start:start + MAX_BATCH]
        progress("generating", index - 1, chunks,
                 "%d of %d strings done" % (start, len(gaps)))
        schema = _schema(langs)
        message = _user_message(batch, langs, snapshot, style_examples, glossary)
        max_tokens = min(4000 + 700 * len(batch) * max(1, len(langs)) // 6, 32000)
        result = _call_model(payload, SYSTEM_PROMPT, message, schema, max_tokens)
        progress("generating", index, chunks,
                 "%d of %d strings done" % (min(start + MAX_BATCH, len(gaps)), len(gaps)))
        by_key = {norm_key(r.get("en", "")): r for r in result.get("rows", [])}
        for index, item in enumerate(batch):
            row = by_key.get(norm_key(item["en"]))
            if row is None:
                # The model dropped or reworded the echo; fall back to position
                # rather than losing the string silently.
                rows = result.get("rows", [])
                row = rows[index] if index < len(rows) else {}
            for lang in langs:
                text = row.get(lang)
                if text:
                    produced[item["key"]][lang] = tidy(text)
    return produced


# ---------------- assembly ----------------

def translate(payload):
    raw = payload.get("strings") or []
    langs = [l for l in (payload.get("languages") or LANGS) if l in LANGS] or LANGS

    seen = set()
    items = []
    for line in raw:
        english = tidy(line)
        if not english:
            continue
        key = norm_key(english)
        if key in seen:
            continue
        seen.add(key)
        items.append({"en": english, "key": key})

    if not items:
        return {"error": "No English text supplied."}

    progress("memory", 0, 1, "loading the translation memory")
    approved = load_approved(payload.get("corrections_root", ""))
    style_examples = load_style_examples(payload.get("corrections_root", ""))
    glossary = load_glossary(payload.get("corrections_root", ""))
    snapshot, snapshot_origin = tm_snapshot(payload.get("design_path", ""),
                                            payload.get("cache_dir", ""))
    progress("memory", 1, 1, "%d strings in memory" % len(snapshot["entries"]))
    entries = snapshot["entries"]

    rows = []
    gaps = []
    for item in items:
        targets = {}
        for lang in langs:
            hit = approved.get(lang, {}).get(item["key"])
            if hit and hit.get("text"):
                targets[lang] = {"text": hit["text"], "source": "approved",
                                 "note": "translator-approved (%s, seen %sx)"
                                         % (hit.get("source", "import"),
                                            hit.get("occurrences", 1))}
                continue
            ours = (entries.get(item["key"]) or {}).get(lang)
            if ours:
                targets[lang] = {"text": ours, "source": "memory",
                                 "note": "already shipped in localization.xlsx"}
        rows.append({"en": item["en"], "key": item["key"], "targets": targets})
        if any(lang not in targets for lang in langs):
            gaps.append(item)

    counts = collections.Counter()
    error = None
    # Filled by generate() when it runs; stays empty on every path that skips
    # it, so the fill loop below reads the same either way.
    produced = {}
    provider = (payload.get("provider") or "").lower()
    offline = None
    if gaps and provider == "ollama":
        offline = ollama_unreachable(payload.get("ollama_base_url", ""))

    if gaps and (payload.get("skip_generation") or offline):
        # Nothing to wait for, so do not wait. Answering instantly with what
        # memory knows beats a dead socket that ends at the same message.
        error = offline or ("No model is configured for this app — nothing was "
                            "generated.")
        error += " The memory tiers below still answered what they could."
    elif gaps:
        needed = sorted({lang for row in rows for lang in langs
                         if lang not in row["targets"]},
                        key=LANGS.index)
        try:
            produced = generate(payload, gaps, needed, snapshot, style_examples,
                                glossary)
        except Exception as exc:                      # noqa: BLE001 - reported, not raised
            # One batch-level reason, not one per cell. Repeating "not
            # translated" across 54 cells hides the single fact the user needs:
            # there was no model to ask.
            error = str(exc) or type(exc).__name__

    # Outside the branches on purpose. This ran inside the generate branch once,
    # so on the skip path the unanswered cells were never created at all - the
    # renderer's own fallback drew them as FAILED while the counts, built from
    # what the backend actually produced, reported none. The table and its own
    # summary contradicted each other.
    if gaps:
        for row in rows:
            for lang in langs:
                if lang in row["targets"]:
                    continue
                text = produced.get(row["key"], {}).get(lang)
                if not text:
                    row["targets"][lang] = {
                        "text": "", "source": "failed",
                        "note": error or "no translation produced",
                        # Deliberately empty when the whole batch failed for one
                        # reason: that reason is reported once, above the table.
                        "warnings": [] if error else ["not translated"]}
                    continue
                # Typography is ours to get right, so it is applied here rather
                # than asked of the model - see QUOTE_STYLE.
                row["targets"][lang] = {"text": apply_quote_style(text, lang),
                                        "source": "generated", "note": ""}

    for row in rows:
        # The English goes into column B of localization.xlsx like every other
        # column, so it breaks the convert the same way - and row_warnings only
        # ever sees the translations. This was the actual path SIM2609_LETTER
        # took: the quotes arrived in the English and were mirrored outward.
        unsafe = pipeline_unsafe(row["en"])
        row["warnings"] = ["English: " + unsafe] if unsafe else []
        for lang, cell in row["targets"].items():
            cell.setdefault("warnings",
                            row_warnings(row["en"], cell["text"], lang, cell["source"],
                                         glossary))
            counts[cell["source"]] += 1
        row.pop("key", None)

    result = {
        "rows": rows,
        "languages": langs,
        "provider": (payload.get("provider") or "anthropic").lower(),
        "memory": {
            "approved_entries": sum(len(v) for v in approved.values()),
            "shipped_entries": len(entries),
            "workbooks": snapshot.get("files", []),
            "snapshot": snapshot_origin,
        },
        "counts": {k: counts[k] for k in ("approved", "memory", "generated", "failed")},
    }
    if error:
        result["provider_error"] = error
        # The two routes that actually work, so the panel can say what to do
        # instead of leaving the user with a wall of failed cells. The MCP
        # route needs no API key at all - there the host is the model.
        result["provider_hint"] = (
            "This panel translates through an Anthropic API key, which is not "
            "configured. Either add one in Settings, or do this batch in Claude "
            "Desktop — ask it to open the localization panel and it runs on your "
            "claude.ai subscription with no key. 'Copy brief for Claude' puts the "
            "whole prompt on your clipboard either way."
        )
    return result


# ---------------- MCP entry points ----------------
# The desktop panel needs an Anthropic API key. The MCP server does not: there
# the HOST is the model, so a designer can do this in Claude Desktop on their
# claude.ai subscription, exactly as they already write Art Requirements.
#
# That splits the work in two. brief_main() hands the host everything the
# memory knows - what is already settled, what still needs translating, the
# nearest shipped strings, and the vendor's own corrections per language. The
# host translates only the gaps. mcp_save_main() then takes the result back and
# applies the deterministic half (quote rules, the checks) server-side, because
# those are ours to get right and must not depend on the model having complied.

def brief(payload):
    """Everything the host model needs, as text. No model call happens here."""
    raw = payload.get("strings") or []
    langs = [l for l in (payload.get("languages") or LANGS) if l in LANGS] or LANGS

    seen, items = set(), []
    for line in raw:
        english = tidy(line)
        if not english:
            continue
        key = norm_key(english)
        if key not in seen:
            seen.add(key)
            items.append({"en": english, "key": key})
    if not items:
        return {"error": "No English text supplied."}

    approved = load_approved(payload.get("corrections_root", ""))
    style_examples = load_style_examples(payload.get("corrections_root", ""))
    glossary = load_glossary(payload.get("corrections_root", ""))
    snapshot, origin = tm_snapshot(payload.get("design_path", ""),
                                   payload.get("cache_dir", ""))
    entries = snapshot["entries"]
    keys = list(entries.keys())
    index = build_index(keys)

    settled, gaps = [], []
    for item in items:
        known = {}
        for lang in langs:
            hit = approved.get(lang, {}).get(item["key"])
            if hit and hit.get("text"):
                known[lang] = (hit["text"], "translator-approved")
                continue
            ours = (entries.get(item["key"]) or {}).get(lang)
            if ours:
                known[lang] = (ours, "already shipped")
        missing = [l for l in langs if l not in known]
        (gaps if missing else settled).append((item, known, missing))

    out = []
    if settled or any(k for _, k, _ in gaps):
        out.append("## Already settled - copy these through unchanged")
        out.append("")
        out.append("These are either the translator's own approved text or what the "
                   "game already ships. Do not retranslate them; pass them back "
                   "exactly as given.")
        out.append("")
        for item, known, _ in settled + gaps:
            if not known:
                continue
            out.append("EN: %s" % item["en"])
            for lang in langs:
                if lang in known:
                    text, why = known[lang]
                    out.append("  %s: %s   (%s)" % (lang, text, why))
            out.append("")

    if gaps:
        needed = sorted({l for _, _, m in gaps for l in m}, key=LANGS.index)
        # Fewer examples each when there are many strings - see TAPER_ABOVE_GAPS.
        per_string = (TAPER_NEIGHBOURS if len(gaps) > TAPER_ABOVE_GAPS
                      else MAX_NEIGHBOURS)
        out.append("## Needs translating")
        out.append("")
        for item, _, missing in gaps:
            out.append("EN: %s" % item["en"])
            out.append("  missing: %s" % ", ".join(missing))
            marks = placeholders(item["en"])
            if marks:
                out.append("  placeholders that MUST survive: %s" % " ".join(marks))
            for term, renderings in glossary_hits(item["en"], glossary):
                shown = ", ".join("%s=%s" % (l, renderings[l])
                                  for l in missing if renderings.get(l))
                if shown:
                    out.append("  fixed term %r - use exactly: %s" % (term, shown))
            near_by_lang = neighbours_by_lang(item["en"], entries, keys, index,
                                              missing, limit=per_string)
            for lang in missing:
                near = near_by_lang.get(lang) or []
                if near:
                    out.append("  similar %s we already ship: %s" % (
                        lang, "; ".join('"%s" -> "%s"' % (n["en"], n["text"])
                                        for n in near)))
            out.append("")

        out.append("## Per-language rules")
        out.append("")
        for lang in needed:
            open_mark, close_mark, pad = QUOTE_STYLE[lang]
            out.append("### %s (%s)" % (LANG_NAMES[lang], lang))
            out.append("- Quotation marks: %s%stext%s%s. Write plain straight quotes if "
                       "you need any - the save tool converts them for you."
                       % (open_mark, pad, pad, close_mark))
            out.append("- Stay at or under the English length; these sit on buttons.")
            for rule in LANG_RULES.get(lang, []):
                out.append("- %s" % rule)
            examples = style_examples.get(lang) or []
            if examples:
                out.append("- Corrections the translator made to our earlier attempts. "
                           "These are the mistakes to avoid, not suggestions:")
                for pair in examples:
                    out.append("    EN: %s" % pair["en"])
                    out.append("      we wrote : %s" % pair["before"])
                    out.append("      they want: %s" % pair["after"])
            out.append("")

    out.append("Product names stay in English: Skip-Bo, Wild Card, Stock Pile. "
               "ALL CAPS English stays ALL CAPS. Keep emoji as they are.")
    out.append("")
    out.append("When every row is complete, call save_translations with one row per "
               "English string above - including the settled ones, unchanged.")

    # The same row shape translate() produces, so the MCP App can render a
    # table straight from the brief - that is the "Check memory" view, which
    # answers with no model involved at all.
    rows, tally = [], collections.Counter()
    for item, known, missing in settled + gaps:
        targets = {}
        for lang in langs:
            if lang in known:
                text, why = known[lang]
                source = "approved" if why == "translator-approved" else "memory"
                targets[lang] = {"text": text, "source": source, "note": why,
                                 "warnings": row_warnings(item["en"], text, lang,
                                                          source, glossary)}
            else:
                targets[lang] = {"text": "", "source": "failed", "note": "needs translating",
                                 "warnings": []}
            tally[targets[lang]["source"]] += 1
        rows.append({"en": item["en"], "targets": targets})

    return {
        "text": "\n".join(out),
        "rows": rows,
        "languages": langs,
        "counts": {"strings": len(items), "settled": len(settled), "gaps": len(gaps),
                   "approved": tally["approved"], "memory": tally["memory"],
                   "generated": 0, "failed": tally["failed"]},
        "memory": {"approved_entries": sum(len(v) for v in approved.values()),
                   "shipped_entries": len(entries),
                   "workbooks": snapshot.get("files", []), "snapshot": origin},
    }


def mcp_save(payload):
    """Take the host's translations, apply the rules we own, and file them.

    Everything the model returns is treated as a draft: quote style is applied
    here rather than trusted, and the checks run whether or not the model
    claims it followed them."""
    rows_in = payload.get("rows") or []
    if not rows_in:
        return {"error": "No rows supplied."}
    out_root = payload.get("corrections_root", "")
    if not out_root:
        return {"error": "No Localization folder configured."}

    approved = load_approved(out_root)
    glossary = load_glossary(out_root)
    # The shipped memory too, not just the approved one. The brief tells the
    # model to pass settled rows back unchanged, so without this every one of
    # them comes back labelled "generated" - wrong provenance in the table, and
    # it would also put the advisory checks back onto text that is already in
    # the game. Absent design_path just means everything reads as generated.
    shipped = {}
    if payload.get("design_path"):
        snapshot, _ = tm_snapshot(payload["design_path"], payload.get("cache_dir", ""))
        shipped = snapshot["entries"]

    rows, warned = [], 0
    for row in rows_in:
        english = tidy(row.get("en", ""))
        if not english:
            continue
        key = norm_key(english)
        targets = {}
        for lang in LANGS:
            text = tidy(row.get(lang) or "")
            if not text:
                targets[lang] = {"text": "", "source": "failed",
                                 "warnings": ["not translated"]}
                warned += 1
                continue
            hit = approved.get(lang, {}).get(key)
            if hit and hit.get("text") == text:
                # Unchanged approved text: it is signed off, so the advisory
                # checks stay off it (see row_warnings).
                source = "approved"
            elif (shipped.get(key) or {}).get(lang) == text:
                # Unchanged from what the game already ships - reused, not new.
                source = "memory"
            else:
                source = "generated"
                text = apply_quote_style(text, lang)
            marks = row_warnings(english, text, lang, source, glossary)
            warned += 1 if marks else 0
            targets[lang] = {"text": text, "source": source, "warnings": marks}
        unsafe = pipeline_unsafe(english)
        rows.append({"en": english, "targets": targets,
                     "warnings": ["English: " + unsafe] if unsafe else []})

    stamp = _utc_stamp()
    ref = _safe_token(payload.get("ref_name") or "translations")
    author = _safe_token(payload.get("author_name") or "unknown")
    folder = os.path.join(out_root, "translations")
    os.makedirs(folder, exist_ok=True)
    stem = os.path.join(folder, "%s_%s_%s" % (stamp, author, ref))

    header = ["English"] + LANGS
    tsv = ["\t".join(header)]
    for row in rows:
        tsv.append("\t".join([row["en"]] + [row["targets"][l]["text"] for l in LANGS]))
    with open(stem + ".tsv", "w", encoding="utf-8", newline="\r\n") as handle:
        handle.write("\n".join(tsv))

    md = ["# %s" % (payload.get("ref_name") or "Translations"), "",
          "Saved %s by %s. Columns are localization.xlsx's own B..H order."
          % (stamp, payload.get("author_name") or "unknown"), "",
          "| " + " | ".join(header) + " |",
          "|" + "---|" * len(header)]
    for row in rows:
        md.append("| " + " | ".join(
            [row["en"].replace("|", "\\|")]
            + [row["targets"][l]["text"].replace("|", "\\|") for l in LANGS]) + " |")
    flagged = ([(r["en"], "EN", w) for r in rows for w in r["warnings"]]
               + [(r["en"], l, w) for r in rows for l in LANGS
                  for w in r["targets"][l]["warnings"]])
    if flagged:
        md += ["", "## Checks to look at", ""]
        md += ["- **%s** / %s: %s" % (en, l, w) for en, l, w in flagged]
    with open(stem + ".md", "w", encoding="utf-8") as handle:
        handle.write("\n".join(md) + "\n")

    tally = collections.Counter(cell["source"] for row in rows
                                for cell in row["targets"].values())
    job_id = payload.get("job_id") or ""
    result = {
        "rows": rows,
        "languages": LANGS,
        "job_id": job_id,
        "counts": {k: tally[k] for k in ("approved", "memory", "generated", "failed")},
        "tsv_path": stem + ".tsv",
        "md_path": stem + ".md",
        "flagged": [{"en": en, "lang": l, "warning": w} for en, l, w in flagged],
        "summary": "%d string(s) x %d languages filed; %d cell(s) flagged."
                   % (len(rows), len(LANGS), len(flagged)),
    }

    # A job_id means the Electron panel is sitting watching for this. Write the
    # answer beside the job and remove the job, so the panel stops waiting and
    # the same batch cannot be picked up twice.
    if job_id:
        job_path, result_path = pending_paths(out_root, job_id)
        try:
            os.makedirs(os.path.dirname(result_path), exist_ok=True)
            with open(result_path, "w", encoding="utf-8") as handle:
                json.dump(result, handle, ensure_ascii=False)
            if os.path.exists(job_path):
                os.remove(job_path)
        except OSError as exc:
            result["job_warning"] = "Could not file the result for %s: %s" % (job_id, exc)

    return result


# ---------------- the Claude Desktop handoff ----------------
# What this is for: the Electron panel cannot make Claude Desktop translate.
# An MCP server only ever answers the host, and no claude:// deep link accepts
# a prompt (checked against the app's own route table: claude://claude,
# resume, code/new, code/continue, code/needs-input, cowork/shared-artifact -
# none of them carry text). So exactly one human action is unavoidable: send
# one message in Claude Desktop.
#
# Everything either side of it is automated. The panel drops a job here, the
# user says one thing in Claude Desktop, the model picks the job up with
# localization_pending, and save_translations writes the answer back beside it
# where the panel is already watching for it.

PENDING_DIR = "pending"


def pending_paths(corrections_root, job_id):
    folder = os.path.join(corrections_root, PENDING_DIR)
    return (os.path.join(folder, job_id + ".json"),
            os.path.join(folder, job_id + ".result.json"))


def pending_take(payload):
    """The oldest unanswered job, with its brief already built.

    Oldest first, not newest: if two designers queued work, the one who has
    been waiting longer should not be starved by the other's later click."""
    root = payload.get("corrections_root", "")
    folder = os.path.join(root, PENDING_DIR)
    if not os.path.isdir(folder):
        return {"job": None, "reason": "No batches are waiting."}

    jobs = sorted(f for f in os.listdir(folder)
                  if f.endswith(".json") and not f.endswith(".result.json"))
    wanted = payload.get("job_id")
    if wanted:
        jobs = [f for f in jobs if f[:-5] == wanted]
    for name in jobs:
        job_id = name[:-5]
        job_path, result_path = pending_paths(root, job_id)
        if os.path.exists(result_path):
            continue                       # already answered
        try:
            with open(job_path, "r", encoding="utf-8-sig") as handle:
                job = json.load(handle)
        except (OSError, ValueError):
            continue
        detail = brief({"strings": job.get("strings") or [],
                        "languages": job.get("languages"),
                        "design_path": job.get("design_path", ""),
                        "corrections_root": root,
                        "cache_dir": payload.get("cache_dir", "")})
        if detail.get("error"):
            return {"job": None, "reason": detail["error"]}
        detail["job_id"] = job_id
        detail["requested_by"] = job.get("author_name", "")
        return {"job": detail}
    return {"job": None, "reason": "No batches are waiting."}


def pending_main():
    _payload_main(pending_take, "Usage: loc_pending <payload.json>")


# ---------------- Excel export ----------------

# The vendor's own header row. Deliberately English rather than localization
# .xlsx's Chinese headers (文本枚举/英语/德语/...): this file gets sent to the
# translator as often as it gets pasted into the workbook, and their delivered
# sheets use exactly these labels.
EXPORT_HEADERS = ["English", "German", "Spanish", "Dutch", "Swedish", "French", "Danish"]

# Verified against WorkSpcae3/localization.xlsx: A=文本枚举, B=英语, C=德语,
# D=西班牙语, E=荷兰语, F=瑞典语, G=法语, H=丹麦语. So English + LANGS in order
# lines up with B..H, and a paste starting at column B lands correctly. Get this
# order wrong and the wrong language ships in the game with nothing to catch it.
EXPORT_COLUMN_NOTE = "Columns match localization.xlsx B..H - paste starting at column B."


def export_xlsx(payload):
    """Write the finished table to .xlsx.

    Two sheets on purpose. The first is paste-clean: nothing but the seven
    columns, so it can go straight into localization.xlsx or to the translator.
    Provenance and warnings would corrupt that paste, so they go on the second
    sheet rather than being dropped."""
    rows = payload.get("rows") or []
    out_path = (payload.get("out_path") or "").strip()
    if not rows:
        return {"error": "Nothing to export."}
    if not out_path:
        return {"error": "No output path given."}

    langs = [l for l in (payload.get("languages") or LANGS) if l in LANGS] or LANGS

    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Translations"

    header = ["English"] + [LANG_NAMES[l] for l in langs]
    sheet.append(header)
    for row in rows:
        targets = row.get("targets") or {}
        sheet.append([row.get("en", "")]
                     + [(targets.get(l) or {}).get("text", "") for l in langs])

    head_font = Font(bold=True)
    head_fill = PatternFill(fill_type="solid", start_color="FFDDEBF7")
    for cell in sheet[1]:
        cell.font = head_font
        cell.fill = head_fill
    sheet.freeze_panes = "A2"
    for index in range(1, len(header) + 1):
        letter = get_column_letter(index)
        longest = max([len(str(header[index - 1]))]
                      + [len(str(sheet.cell(row=r, column=index).value or ""))
                         for r in range(2, sheet.max_row + 1)] or [10])
        sheet.column_dimensions[letter].width = min(max(longest + 2, 12), 60)
    for r in range(2, sheet.max_row + 1):
        for c in range(1, len(header) + 1):
            sheet.cell(row=r, column=c).alignment = Alignment(vertical="top",
                                                              wrap_text=True)

    detail = workbook.create_sheet("Provenance")
    detail.append(["English", "Language", "Source", "Checks"])
    for cell in detail[1]:
        cell.font = head_font
    flagged = 0
    for row in rows:
        targets = row.get("targets") or {}
        en_marks = "; ".join(row.get("warnings") or [])
        if en_marks:
            flagged += 1
            detail.append([row.get("en", ""), "English", "", en_marks])
        for lang in langs:
            cell = targets.get(lang) or {}
            warnings = "; ".join(cell.get("warnings") or [])
            if warnings:
                flagged += 1
            detail.append([row.get("en", ""), LANG_NAMES[lang],
                           cell.get("source", ""), warnings])
    detail.freeze_panes = "A2"
    for index, width in enumerate((46, 12, 12, 60), start=1):
        detail.column_dimensions[get_column_letter(index)].width = width

    folder = os.path.dirname(os.path.abspath(out_path))
    if folder:
        os.makedirs(folder, exist_ok=True)
    workbook.save(out_path)
    workbook.close()
    return {"path": out_path, "rows": len(rows), "languages": langs,
            "flagged": flagged, "note": EXPORT_COLUMN_NOTE}


def export_main():
    _payload_main(export_xlsx, "Usage: loc_export_xlsx <payload.json>")


def _utc_stamp():
    from datetime import datetime, timezone
    return datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")


def _safe_token(value, maxlen=48):
    token = re.sub(r"[^A-Za-z0-9._-]+", "-", str(value or "")).strip("-")
    return token[:maxlen] or "unknown"


def _payload_main(handler, usage):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    if len(sys.argv) < 2:
        print(json.dumps({"error": usage}))
        sys.exit(1)
    with open(sys.argv[1], "r", encoding="utf-8-sig") as handle:
        payload = json.load(handle)
    try:
        print(json.dumps(handler(payload), ensure_ascii=False))
    except Exception as exc:                          # noqa: BLE001
        print(json.dumps({"error": "%s: %s" % (type(exc).__name__, exc)}))
        sys.exit(1)


def brief_main():
    _payload_main(brief, "Usage: loc_mcp_brief <payload.json>")


def mcp_save_main():
    _payload_main(mcp_save, "Usage: loc_mcp_save <payload.json>")


def translate_main():
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    if len(sys.argv) < 2:
        print(json.dumps({"error": "Usage: loc_translate <payload.json>"}))
        sys.exit(1)
    with open(sys.argv[1], "r", encoding="utf-8-sig") as handle:
        payload = json.load(handle)
    try:
        print(json.dumps(translate(payload), ensure_ascii=False))
    except Exception as exc:                          # noqa: BLE001
        print(json.dumps({"error": "%s: %s" % (type(exc).__name__, exc)}))
        sys.exit(1)


# ---------------- self-check (offline, no model call) ----------------

def selfcheck():
    import tempfile

    # --- the quote rule, per language, from a straight-quote pair ---
    cases = {
        "DE": ("He said 'yes' today", "He said „yes“ today"),
        "ES": ("He said 'yes' today", "He said «yes» today"),
        "NL": ("He said 'yes' today", "He said ‘yes’ today"),
        "SV": ("He said 'yes' today", "He said ”yes” today"),
        "FR": ("He said 'yes' today", "He said « yes » today"),
        "DA": ("He said 'yes' today", "He said ”yes” today"),
    }
    for lang, (before, after) in cases.items():
        got = apply_quote_style(before, lang)
        assert got == after, "%s: %r -> %r, wanted %r" % (lang, before, got, after)

    # A word-internal apostrophe is NOT a quote pair and must survive untouched.
    assert apply_quote_style("Dorothy's shoes", "NL") == "Dorothy's shoes"
    assert apply_quote_style("l'origine de l'oeuvre", "FR") == "l'origine de l'oeuvre"

    # --- the mechanical checks ---
    assert row_warnings("Complete %s levels", "Voltooi %s niveaus", "NL") == []
    assert any("placeholder missing" in w
               for w in row_warnings("Complete %s levels", "Voltooi niveaus", "NL"))
    assert any("placeholder added" in w
               for w in row_warnings("Collect now", "Sammle %s jetzt", "DE"))
    assert any("longer than the English" in w
               for w in row_warnings("Play", "Spiele jetzt sofort ein ganzes Spiel", "DE"))
    assert any("straight quote" in w
               for w in row_warnings("Say 'go'", "Sag 'los'", "DE"))
    assert any("identical to the English" in w
               for w in row_warnings("Collect your daily reward", "Collect your daily reward", "DE"))
    assert row_warnings("OK", "", "DE") == ["empty translation"]

    # --- what would break coder_convert ---
    # The exact string that stopped the convert: a letter wrapped in quotes.
    assert pipeline_unsafe('"Dear Manager,\n\nRegards"')
    assert pipeline_unsafe('He said "no"')
    assert pipeline_unsafe("ends with a backslash \\")
    assert pipeline_unsafe(r"a \q escape")
    assert pipeline_unsafe(r"short \u12 escape")
    # Legal: the workbook's own convention, and every non-ASCII quote mark.
    assert pipeline_unsafe(r'Skip-Bo cards are \"wild\"') is None
    assert pipeline_unsafe(r"a \\ and a \n and a é") is None
    assert pipeline_unsafe("Dear Manager,\n\nRegards") is None
    for _open, _close, _pad in QUOTE_STYLE.values():
        assert pipeline_unsafe("%s%s%s" % (_open, "hallo", _close)) is None
    # It is a build break, so no source is exempt - not even signed-off text.
    for _src in ("generated", "memory", "approved"):
        assert any("end the JSON string" in w
                   for w in row_warnings("Hi", '"Hallo"', "DE", _src)), _src

    # --- the quote check must not fire on elision -------------------------
    # 15 of the 21 warnings on the 09-17 delivery were this: l'Usine, d'Or,
    # J'ai, qu'ils flagged as "straight quote left in".
    for _fr in ("L'usine est en désordre", "J'ai construit une chocolaterie",
                "Trésor de l'Usine à bonbons", "Équipement Étoile d'or",
                "des jetons qu'ils gagnent"):
        assert not any("straight quote" in w
                       for w in row_warnings("x", _fr, "FR")), _fr
    # A real pair still has to be caught, in both quote characters.
    assert any("straight quote" in w for w in row_warnings("Say 'go'", "Sag 'los'", "DE"))
    assert QUOTE_PAIR_RE.search("il a dit 'oui' hier")
    assert not QUOTE_PAIR_RE.search("Dorothy's shoes and l'origine")

    # --- the glossary ------------------------------------------------------
    assert _stem("chokolademager") == "chokolademag", _stem("chokolademager")
    assert not "chokolade".startswith(_stem("chokolademager")), \
        "the stem must not be satisfied by a shorter compound"
    assert _stem("Chef") == "chef"
    _gl = {"Manager": {"FR": "Gérant", "DA": "Chef"},
           "Sweet Coin": {"DA": "Slikmønt"}}
    assert glossary_warning("Hello Manager", "Bonjour Manager", "FR", _gl)
    assert glossary_warning("Hello Manager", "Bonjour Gérant", "FR", _gl) is None
    # Inflection must not trip it: plurals and case endings still count.
    assert glossary_warning("Earn Sweet Coins", "Optjen Slikmønter", "DA", _gl) is None
    # A term that is not in the English is not checked.
    assert glossary_warning("Collect now", "Saml nu", "DA", _gl) is None
    # Case is LANG_RULES' business, not the glossary's.
    assert glossary_warning("Hello Manager", "bonjour gérant", "FR", _gl) is None
    # Longest term first, so a nested term does not shadow the specific one.
    assert [t for t, _ in glossary_hits("The Chocolate Factory chocolatier",
                                        {"Chocolate Factory": {}, "chocolatier": {}})] \
        == ["Chocolate Factory", "chocolatier"]
    # And it stays in the advisory tier: signed-off text is never second-guessed.
    assert row_warnings("Hello Manager", "Bonjour Manager", "FR", "approved", _gl) == []

    # --- house rules reach the prompt --------------------------------------
    for _lang in LANGS:
        assert LANG_RULES.get(_lang), _lang
    # Direction matters more than wording here: French keeps vous (its final
    # text was 23 formal markers to 0), German moves to du. Getting either
    # backwards would make the tool worse than having no rule.
    # The direction is carried by which pronoun the rule PRESCRIBES, so assert
    # on the prescription, not on mere mention - "do not switch to tu" names tu
    # legitimately.
    assert "formally: vous" in " ".join(LANG_RULES["FR"]), "French must prescribe vous"
    assert "informally: du" in " ".join(LANG_RULES["DE"]), "German must prescribe du"
    assert "informally: tú" in " ".join(LANG_RULES["ES"]), "Spanish must prescribe tú"
    assert re.search(r"\bSie\b", " ".join(LANG_RULES["DE"])), "German must name Sie as wrong"

    # The pre-flight must call an unroutable address unreachable quickly, and
    # must not claim a bad URL is fine.
    import time as _time
    started = _time.time()
    reason = ollama_unreachable("http://10.255.255.1:11434")
    assert reason, "an unroutable host was reported as reachable"
    assert _time.time() - started < OLLAMA_PREFLIGHT_SECONDS + 1.5, \
        "the pre-flight took longer than its own timeout"
    assert ollama_unreachable(""), "an empty URL must not be treated as reachable"

    # A skipped generation must still ACCOUNT for the cells it did not fill.
    # This is the bug the table's own summary once contradicted: the renderer
    # drew FAILED badges while counts.failed said zero.
    scratch0 = tempfile.mkdtemp(prefix="loc_skip_check_")
    design0 = os.path.join(scratch0, "design")
    os.makedirs(design0)
    wb0 = openpyxl.Workbook()
    sh0 = wb0.active
    sh0.title = "localize"
    sh0.append(["key", "English", "German", "Spanish", "Dutch",
                "Swedish", "French", "Danish"])
    sh0.append(["BTN_COLLECT", "Collect", "Sammeln", "Recoger", "Verzamel",
                "Ta emot", "Collecter", "Indsaml"])
    wb0.save(os.path.join(design0, "localization.xlsx"))
    wb0.close()

    skipped = translate({"strings": ["Collect", "A brand new string"],
                         "design_path": design0, "corrections_root": "",
                         "cache_dir": os.path.join(scratch0, "cache"),
                         "skip_generation": True})
    assert skipped["counts"]["memory"] == 6, skipped["counts"]
    assert skipped["counts"]["failed"] == 6, \
        "unfilled cells were not accounted for: %s" % skipped["counts"]
    unanswered = skipped["rows"][1]["targets"]
    assert len(unanswered) == 6 and all(c["source"] == "failed"
                                       for c in unanswered.values()), unanswered
    assert skipped.get("provider_error"), "the skip must be reported, not silent"

    # --- Excel export: the column order is the part that can silently corrupt
    # a paste into localization.xlsx, so it is checked against the real order
    # (A=key, B=English, C=German, D=Spanish, E=Dutch, F=Swedish, G=French,
    # H=Danish) rather than against itself.
    export_rows = [{
        "en": "Collect",
        "targets": {
            "DE": {"text": "Sammeln", "source": "memory", "warnings": []},
            "ES": {"text": "Recoger", "source": "memory", "warnings": []},
            "NL": {"text": "Verzamel", "source": "memory", "warnings": []},
            "SV": {"text": "Ta emot", "source": "memory", "warnings": []},
            "FR": {"text": "Collecter", "source": "memory", "warnings": []},
            "DA": {"text": "Indsaml", "source": "approved", "warnings": ["check me"]},
        },
    }]
    xlsx_path = os.path.join(scratch0, "out", "export.xlsx")
    written = export_xlsx({"rows": export_rows, "out_path": xlsx_path})
    assert written.get("path") == xlsx_path, written
    assert written["flagged"] == 1, written

    check = openpyxl.load_workbook(xlsx_path)
    grid = check["Translations"]
    assert [c.value for c in grid[1]] == EXPORT_HEADERS, [c.value for c in grid[1]]
    # Row 2 must read across in localization.xlsx's B..H order.
    assert [c.value for c in grid[2]] == ["Collect", "Sammeln", "Recoger", "Verzamel",
                                          "Ta emot", "Collecter", "Indsaml"], \
        [c.value for c in grid[2]]
    # German is column B of the target languages, i.e. column C of the workbook.
    assert grid.cell(row=1, column=2).value == "German"
    assert grid.cell(row=1, column=7).value == "Danish"
    detail = check["Provenance"]
    assert [c.value for c in detail[1]] == ["English", "Language", "Source", "Checks"]
    flagged_rows = [r for r in detail.iter_rows(min_row=2, values_only=True) if r[3]]
    assert flagged_rows and flagged_rows[0][1] == "Danish", flagged_rows
    assert flagged_rows[0][2] == "approved", flagged_rows
    check.close()

    # --- retrieval shortlist: fast is worthless if it stops finding things ---
    # The shortlist replaced an exhaustive difflib scan that cost 4-8s per
    # string at six languages. It must still agree with that scan, and in
    # particular must survive morphology: a word-only index made "Chocolates"
    # miss "Chocolate" entirely and return nothing.
    probe_keys = ["chocolate", "chocolate bars", "collect", "collect all",
                  "play summer beach levels to get more", "wild card"]
    probe_index = build_index(probe_keys)
    plural = similar_keys("Chocolates", probe_keys, probe_index, 5)
    assert "chocolate" in plural, "the shortlist lost a plural/singular match: %s" % plural
    exhaustive = difflib.get_close_matches(norm_key("Chocolates"), probe_keys, n=5, cutoff=0.55)
    assert set(plural) >= set(exhaustive), \
        "shortlist %s missed something the full scan found %s" % (plural, exhaustive)
    # A probe sharing no token with anything must fall back, not return junk.
    assert similar_keys("zzzz qqqq", probe_keys, probe_index, 5) == []

    # And the per-language split must not re-run the scan or drop languages.
    entries_probe = {"collect": {"en": "Collect", "DE": "Sammeln", "FR": "Collecter"}}
    by_lang = neighbours_by_lang("Collect all", entries_probe, ["collect"],
                                 build_index(["collect"]), ["DE", "FR", "ES"])
    assert by_lang["DE"] and by_lang["DE"][0]["text"] == "Sammeln", by_lang
    assert by_lang["FR"] and by_lang["FR"][0]["text"] == "Collecter", by_lang
    assert by_lang["ES"] == [], "a language with no translation must come back empty"

    # A subset of languages must still export in canonical order, not the order
    # the caller happened to pass.
    subset_path = os.path.join(scratch0, "out", "subset.xlsx")
    export_xlsx({"rows": export_rows, "languages": ["DA", "DE"],
                 "out_path": subset_path})
    sub = openpyxl.load_workbook(subset_path)
    assert [c.value for c in sub["Translations"][1]] == ["English", "Danish", "German"], \
        [c.value for c in sub["Translations"][1]]
    sub.close()

    assert export_xlsx({"rows": [], "out_path": xlsx_path}).get("error")
    assert export_xlsx({"rows": export_rows, "out_path": ""}).get("error")

    # --- multi-line strings survive tidy(), matching, and the export ---
    # 162 cells in localization.xlsx hold a real line break. Flattening one
    # changes the string without saying so, which is the failure this guards.
    newline = chr(10)
    multi = "First line." + newline + newline + "Third line."
    assert tidy(multi) == multi, repr(tidy(multi))
    assert tidy("  spaced   out  ") == "spaced out"
    assert tidy("trailing " + newline + "  indented") == "trailing" + newline + "indented"
    assert tidy("crlf\r\nhere") == "crlf" + newline + "here"
    # Matching must still ignore the line break, so a lookup succeeds either way.
    assert norm_key(multi) == norm_key("First line. Third line.")

    multi_path = os.path.join(scratch0, "out", "multiline.xlsx")
    export_xlsx({"rows": [{"en": multi, "targets": {
        lang: {"text": multi, "source": "generated", "warnings": []}
        for lang in LANGS}}], "out_path": multi_path})
    check2 = openpyxl.load_workbook(multi_path)
    assert check2["Translations"].cell(row=2, column=1).value == multi, "export flattened it"
    assert check2["Translations"].cell(row=2, column=2).value == multi
    check2.close()

    # Reused text gets the correctness check and nothing else: the vendor's own
    # Danish for "Redwood" is 1.9x the English and must NOT be flagged for it.
    assert row_warnings("Redwood", "Kæmpe-sekvoja", "DA", "approved") == []
    assert row_warnings("Redwood", "Kæmpe-sekvoja", "DA", "memory") == []
    assert any("longer than the English"
               in w for w in row_warnings("Redwood", "Kæmpe-sekvoja", "DA", "generated"))
    # ...but a dropped placeholder is a crash whoever wrote it.
    assert any("placeholder missing" in w
               for w in row_warnings("Win %s rounds", "Vind runder", "DA", "approved"))

    # --- memory precedence and the snapshot cache, end to end, no model ---
    scratch = tempfile.mkdtemp(prefix="loc_translate_check_")
    design = os.path.join(scratch, "design")
    corrections = os.path.join(scratch, "corrections")
    os.makedirs(design)
    os.makedirs(os.path.join(corrections, "approved"))
    os.makedirs(os.path.join(corrections, "pairs"))

    # --- style examples must cover both writing jobs ----------------------
    # Ranking on length alone gave the model six trivia questions per language
    # and not one button label, while 46% of the delivery it was about to write
    # was 30 characters or shorter. Stock this with plenty of long trivia and a
    # few short UI strings: the short ones must still get through.
    _pairs = ([{"en": "Which %s did the explorer chart in the year %s of the voyage?" % (i, i),
                "before": "b%d" % i, "after": "a%d" % i, "category": "wording"}
               for i in range(12)]
              + [{"en": "Tap to collect", "before": "b", "after": "a", "category": "wording"},
                 {"en": "Revert your move?", "before": "b", "after": "a", "category": "wording"},
                 {"en": "Upgrade Now", "before": "b", "after": "a", "category": "wording"}])
    with open(os.path.join(corrections, "pairs", "DE.json"), "w", encoding="utf-8") as handle:
        json.dump(_pairs, handle)
    _chosen = load_style_examples(corrections)["DE"]
    assert len(_chosen) == MAX_STYLE_EXAMPLES, _chosen
    _short = [p for p in _chosen if len(p["en"]) <= SHORT_STRING]
    assert len(_short) == 3, "short UI strings were crowded out again: %s" % (
        [p["en"] for p in _chosen])
    assert not any(TRIVIA_RE.match(p["en"]) for p in _short), _short
    # With only 3 short pairs in stock the remaining budget must still fill up
    # rather than returning a short list.
    os.remove(os.path.join(corrections, "pairs", "DE.json"))

    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "localize"
    sheet.append(["key", "English", "German", "Spanish", "Dutch", "Swedish", "French", "Danish"])
    sheet.append(["BTN_COLLECT", "Collect", "Sammeln", "Recoger", "Verzamel",
                  "Ta emot", "Collecter", "Indsaml"])
    sheet.append(["BTN_NEXT", "Next", "Weiter", "Siguiente", "Volgende",
                  "Nästa", "Suivant", "Næste"])
    workbook.save(os.path.join(design, "localization.xlsx"))
    workbook.close()

    with open(os.path.join(corrections, "approved", "DE.json"), "w", encoding="utf-8") as handle:
        json.dump({"collect": {"en": "Collect", "text": "Einsammeln",
                               "source": "Rel-PR.xlsx", "occurrences": 3}}, handle)

    payload = {"strings": ["Collect", "Next", "Collect"], "design_path": design,
               "corrections_root": corrections, "cache_dir": os.path.join(scratch, "cache")}
    result = translate(payload)

    assert len(result["rows"]) == 2, "duplicate English was not collapsed"
    collect = result["rows"][0]["targets"]
    # Approved beats our shipped text; every other language falls back to it.
    assert collect["DE"] == {"text": "Einsammeln", "source": "approved",
                             "note": "translator-approved (Rel-PR.xlsx, seen 3x)",
                             "warnings": []}, collect["DE"]
    assert collect["ES"]["text"] == "Recoger", collect["ES"]
    assert collect["ES"]["source"] == "memory", collect["ES"]
    assert result["counts"]["generated"] == 0, "nothing should have been generated"
    assert result["memory"]["snapshot"] == "built", result["memory"]

    # Second run must hit the cache rather than reparse the workbook.
    again = translate(payload)
    assert again["memory"]["snapshot"] == "cache", again["memory"]

    # A quote in the ENGLISH has to reach the filed output. row_warnings only
    # ever sees translations, so nothing else would catch it - and the English
    # is column B of localization.xlsx, which is the exact shape that stopped
    # coder_convert on SIM2609_LETTER.
    quoted = '"Dear Manager,\n\nRegards"'
    assert translate({"strings": [quoted], "design_path": design,
                      "corrections_root": corrections,
                      "cache_dir": os.path.join(scratch, "cache"),
                      # No model: the English check must not depend on one.
                      "skip_generation": True,
                      })["rows"][0]["warnings"], "panel did not flag the English"

    saved = mcp_save({"rows": [{"en": quoted, "DE": "Lieber Manager"}],
                      "corrections_root": corrections, "ref_name": "ZZ_SELFTEST"})
    assert not saved.get("error"), saved
    assert saved["rows"][0]["warnings"], "filed row did not flag the English"
    assert any(f["lang"] == "EN" and "end the JSON string" in f["warning"]
               for f in saved["flagged"]), saved["flagged"]
    # Leave nothing behind; a silent cleanup failure is how selftest files end
    # up in the shared folder.
    for _key in ("tsv_path", "md_path"):
        os.remove(saved[_key])

    print("selfcheck OK - quote rules for %d languages, mechanical checks, "
          "memory precedence, snapshot cache" % len(cases))


if __name__ == "__main__":
    if "--selfcheck" in sys.argv:
        selfcheck()
    else:
        translate_main()
