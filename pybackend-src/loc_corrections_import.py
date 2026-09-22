"""
Import the third-party translator's correction workbooks into a reusable
translation memory.

The vendor delivers one .xlsx per release with the columns
English | German | Spanish | Dutch | Swedish | French | Danish, and marks every
cell they changed with a yellow fill. That fill is the only signal for which
cells are corrections - there is no key column, no change log, and no separate
"before" file. So this script reads the fills, not just the values.

Four things about the real deliveries drive the shape of this file. Each one
was found by a file that parsed "successfully" and produced wrong output:

1. Some workbooks are ISO 29500 STRICT OOXML. openpyxl does not support that
   dialect and returns a workbook with ZERO SHEETS rather than raising - so a
   naive importer reports success having read nothing. Two of the first nine
   deliveries were Strict, holding ~42% of all corrections. The only real
   difference is the namespace URIs, so we rewrite them into a temp copy
   (see _to_transitional) before handing the file to openpyxl.

2. The same yellow reads back as "FFFFFF00" or "00FFFF00" depending on how
   Excel happened to store the alpha byte. Comparing the full 8 characters
   silently matches nothing. Compare the last six.

3. Yellow is not the only fill in these files. Light blue (DDEBF7), light
   green (E2F0D9) and a few hundred THEME-based fills are used for section
   banding - one file bands 447 English cells that way. Theme fills carry no
   RGB at all, so we gate on fill_type == "solid" AND a real RGB string.

4. The header row is not always row 1 (one file has a title row above it, one
   has a blank row 2), and the labels vary - "English ", "ENGLISH", " Spanish".
   Find the header by looking for the language names; never hardcode a row.

Matching corrections back to our own translations is by ENGLISH TEXT, because
the vendor sheets carry no key column. That needs both localization.xlsx and
localization_quiz.xlsx - the trivia corrections live only in the quiz file, and
indexing just the main one drops ~70% of the matches on the floor.

WorkSpcae3 is the workspace to point --design at. It is the master that pipeline
step 5 clones into the others, and it is where corrections land first: run
against WorkSpcae_Design instead and 464 of the same corrections look unmerged
purely because that clone is behind.

That being the master is also why --history exists. Because corrections are
already applied there, comparing a delivery against the current workbook yields
almost no before/after signal - 79 pairs, with Dutch and Swedish at 5 each.
--history fetches older copies out of SVN and recovers what we had BEFORE each
correction was merged: two revisions turn those 79 pairs into 536.

Run:
  python loc_corrections_import.py --src "<LOC Corrections folder>"
      --design "<...\\DesignData\\WorkSpcae3\\design>"
      --history 2026-07-01,2026-06-01 [--out <dir>]
  python loc_corrections_import.py --selfcheck
"""
import argparse
import collections
import glob
import json
import os
import re
import subprocess
import tempfile
import unicodedata
import zipfile

import openpyxl

# Column order the vendor uses, which is also localization.xlsx's B..H order.
LANG_ORDER = ["EN", "DE", "ES", "NL", "SV", "FR", "DA"]
TARGETS = LANG_ORDER[1:]

HEADER_ALIASES = {
    "english": "EN", "german": "DE", "spanish": "ES", "dutch": "NL",
    "swedish": "SV", "french": "FR", "danish": "DA",
}

YELLOW = "FFFF00"

# Strict OOXML uses purl.oclc.org namespaces; Transitional uses
# schemas.openxmlformats.org. Everything else about the parts is identical,
# which is why a straight string swap is enough to make openpyxl read them.
STRICT_MARKER = b"http://purl.oclc.org/ooxml/"
NS_SWAPS = [
    ("http://purl.oclc.org/ooxml/spreadsheetml/main",
     "http://schemas.openxmlformats.org/spreadsheetml/2006/main"),
    ("http://purl.oclc.org/ooxml/officeDocument/relationships",
     "http://schemas.openxmlformats.org/officeDocument/2006/relationships"),
    ("http://purl.oclc.org/ooxml/officeDocument/extendedProperties",
     "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"),
    ("http://purl.oclc.org/ooxml/officeDocument/docPropsVTypes",
     "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"),
    ("http://purl.oclc.org/ooxml/officeDocument/customProperties",
     "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"),
]
STRICT_ATTR = 'conformance="strict"'

QUOTE_CHARS = "\"'‘’‚‛“”„‟«»‹›"

# Characters the vendor retypes differently from us but which mean the same
# thing. Used by norm_key (matching) and strip_typography (classification).
UNIFY = [
    ("’", "'"), ("‘", "'"), ("“", '"'), ("”", '"'),
    (" ", " "), ("​", ""), ("–", "-"), ("—", "-"),
]


# ---------------- text normalisation ----------------

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
    """Lookup form. Aggressive on purpose: the vendor retypes English rather
    than copying it, so curly quotes, non-breaking spaces and a trailing full
    stop routinely differ from ours for what is plainly the same string.
    Normalising these took the match rate against our workbooks from 27% to
    94% on the first nine deliveries."""
    text = unicodedata.normalize("NFKC", str(value))
    for a, b in UNIFY:
        text = text.replace(a, b)
    return re.sub(r"\s+", " ", text).strip().lower().rstrip(".!?:;,")


def strip_typography(text):
    """Everything a typographic-only correction is allowed to change."""
    text = unicodedata.normalize("NFKC", text)
    text = "".join(" " if ch in QUOTE_CHARS else ch for ch in text)
    for a, b in (("–", "-"), ("—", "-"), ("…", "..."), (" ", " ")):
        text = text.replace(a, b)
    return re.sub(r"\s+", " ", text).strip()


def _depunct(text):
    return re.sub(r"[^\w\s]", "", strip_typography(text)).lower()


def classify(before, after):
    """Which kind of correction this is. The split matters: on the first nine
    deliveries 26% of corrections were typographic/casing/punctuation only -
    those are deterministic per-language rules, not judgement calls, and the
    translator should never have to make them by hand twice."""
    if strip_typography(before) == strip_typography(after):
        return "typographic"
    if strip_typography(before).lower() == strip_typography(after).lower():
        return "casing"
    if _depunct(before) == _depunct(after):
        return "punctuation"
    return "wording"


# ---------------- workbook access ----------------

def is_strict(path):
    with zipfile.ZipFile(path) as archive:
        return STRICT_MARKER in archive.read("xl/workbook.xml")


def _rewrite_namespaces(src, dst, to_strict=False):
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.endswith((".xml", ".rels")):
                text = data.decode("utf-8")
                for strict_ns, transitional_ns in NS_SWAPS:
                    if to_strict:
                        text = text.replace(transitional_ns, strict_ns)
                    else:
                        text = text.replace(strict_ns, transitional_ns)
                if not to_strict:
                    text = text.replace(STRICT_ATTR, "")
                data = text.encode("utf-8")
            zout.writestr(item, data)


def open_workbook(path, scratch_dir):
    """Returns (workbook, was_strict). Always read_only - fills come through
    there identically to normal mode, verified, and these files are big."""
    strict = is_strict(path)
    if strict:
        safe = re.sub(r"[^A-Za-z0-9._-]", "_", os.path.basename(path))
        target = os.path.join(scratch_dir, "conv_" + safe)
        _rewrite_namespaces(path, target)
        path = target
    return openpyxl.load_workbook(path, read_only=True, data_only=True), strict


def fill_token(cell):
    """Normalised fill colour, or None when the cell has no solid fill.

    Theme fills return an openpyxl error-descriptor rather than a string for
    .rgb, so the isinstance checks are load-bearing, not defensive noise."""
    fill = cell.fill
    if fill is None or fill.fill_type != "solid":
        return None
    colour = fill.start_color
    rgb = getattr(colour, "rgb", None)
    if isinstance(rgb, str) and len(rgb) >= 6:
        return "RGB:" + rgb[-6:].upper()
    theme = getattr(colour, "theme", None)
    if isinstance(theme, int):
        tint = getattr(colour, "tint", 0) or 0
        return "THEME:%s/%s" % (theme, round(tint, 2))
    indexed = getattr(colour, "indexed", None)
    if isinstance(indexed, int):
        return "INDEXED:%s" % indexed
    return "UNKNOWN"


def is_correction(cell):
    return fill_token(cell) == "RGB:" + YELLOW


def find_header(rows):
    """Row index whose cells name the languages. None when absent."""
    for index, row in enumerate(rows[:12]):
        labels = {str(c.value or "").strip().lower() for c in row}
        if "english" in labels and "german" in labels:
            return index
    return None


def column_map(header_row):
    """Header cell index -> language code, so a file that reorders or omits a
    column still lands in the right place."""
    mapping = {}
    for index, cell in enumerate(header_row):
        code = HEADER_ALIASES.get(str(cell.value or "").strip().lower())
        if code:
            mapping[index] = code
    return mapping


# ---------------- extraction ----------------

def read_corrections(path, scratch_dir):
    """Every yellow target cell in one workbook, plus a census of what was seen."""
    workbook, strict = open_workbook(path, scratch_dir)
    name = os.path.basename(path)
    found = []
    colours = collections.Counter()
    sheets = []
    for sheet_name in workbook.sheetnames:
        sheet = workbook[sheet_name]
        max_col = min(sheet.max_column or 1, 12)
        rows = list(sheet.iter_rows(min_row=1, max_row=sheet.max_row, max_col=max_col))
        header = find_header(rows)
        if header is None:
            sheets.append({"sheet": sheet_name, "rows": len(rows),
                           "header": None, "corrections": 0})
            continue
        columns = column_map(rows[header])
        english_col = next((i for i, c in columns.items() if c == "EN"), None)
        before_count = len(found)
        for row in rows[header + 1:]:
            english = None
            if english_col is not None and english_col < len(row):
                english = row[english_col].value
            for index, code in columns.items():
                if code == "EN" or index >= len(row):
                    continue
                cell = row[index]
                token = fill_token(cell)
                if token:
                    colours[token] += 1
                if cell.value in (None, "") or not is_correction(cell):
                    continue
                # A correction with nothing to key it on is unusable, but it
                # must be reported rather than dropped in silence.
                found.append({
                    "lang": code,
                    "english": tidy(english) if english not in (None, "") else None,
                    "after": tidy(cell.value),
                    "source": name, "sheet": sheet_name, "row": cell.row,
                })
        sheets.append({"sheet": sheet_name, "rows": len(rows), "header": header + 1,
                       "corrections": len(found) - before_count})
    workbook.close()
    return {"file": name, "strict": strict, "sheets": sheets,
            "colours": dict(colours), "corrections": found}


def _generation_rank(source_name):
    """The folder holds several generations of one delivery (…-PR,
    …-PR-Updated, …-PR-Updated-07.07). More "updated" wins; the longer name
    breaks ties, which is the order the vendor actually names them in."""
    name = source_name.lower()
    return (name.count("updated"), len(name))


def dedupe(records):
    """Collapse to one correction per (English, language).

    Returns (best, conflicts, ambiguous). Later generations win a disagreement
    between files, and the clash is recorded either way.

    Ambiguity is different from a conflict and must not be resolved by picking
    one. When a single delivery corrects the SAME English text two different
    ways, that English is not a key - it is one string reused for two different
    pieces of content. Trivia answers are the obvious case: English "8" is the
    answer to more than one question, so "8" -> "8" and "8" -> "9" are both
    right. Those are held out of the approved memory entirely rather than
    silently collapsed into whichever one sorted first."""
    grouped = collections.defaultdict(list)
    for record in records:
        if record["english"] is None:
            continue
        grouped[(norm_key(record["english"]), record["lang"])].append(record)

    best = {}
    conflicts = []
    ambiguous = {}
    for key, group in grouped.items():
        within_file = collections.defaultdict(set)
        for record in group:
            within_file[record["source"]].add(record["after"])
        if any(len(values) > 1 for values in within_file.values()):
            ambiguous["%s|%s" % key] = {
                "english": group[0]["english"], "lang": group[0]["lang"],
                "variants": sorted({r["after"] for r in group}),
                "sources": sorted({r["source"] for r in group}),
            }
            continue
        winner = max(group, key=lambda r: _generation_rank(r["source"]))
        for losing in sorted({r["after"] for r in group} - {winner["after"]}):
            source = next(r["source"] for r in group if r["after"] == losing)
            conflicts.append({"english": winner["english"], "lang": winner["lang"],
                              "kept": winner["after"], "kept_from": winner["source"],
                              "dropped": losing, "dropped_from": source})
        winner = dict(winner, occurrences=len(group))
        best[key] = winner
    return best, conflicts, ambiguous


# ---------------- our own translations, for the "before" side ----------------

TM_FILES = ["localization.xlsx", "localization_quiz.xlsx"]
SKIP_SHEETS = {"languages", "country"}


def _read_translation_sets(paths):
    """norm_key(english) -> lang -> {every translation these files hold}."""
    memory = collections.defaultdict(lambda: collections.defaultdict(set))
    for path in paths:
        if not os.path.exists(path):
            continue
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
                for offset, code in enumerate(TARGETS, start=2):
                    value = row[offset] if len(row) > offset else None
                    if value not in (None, ""):
                        memory[key][code].add(tidy(value))
        workbook.close()
    return memory


def load_our_translations(design_dir):
    """norm_key(english) -> lang -> set of our current translations.

    Both workbooks, because the trivia corrections only ever match the quiz
    file. Column A is the key, B..H are EN + the six targets in LANG_ORDER."""
    paths = [os.path.join(design_dir, name) for name in TM_FILES]
    used = [name for name, path in zip(TM_FILES, paths) if os.path.exists(path)]
    return _read_translation_sets(paths), used


# ---------------- our own filed batches (the cheap "before" side) -----------
# Once a correction is merged into the workbook, the workbook can no longer
# say what we wrote before it - and the pair corpus is what teaches the model
# style. Importing the 09-17 reviews dropped it from 245 pairs to 81 for
# exactly that reason, with Dutch and Swedish down to 3 wording pairs each.
#
# But for anything the tool itself translated we already hold the answer: the
# filed job result is precisely what we produced, before any review touched it.
# That is better evidence than SVN history (exact, local, instant) and it needs
# no network - svn cat of a 3MB xlsx was still hanging after four minutes.

def load_filed_results(out_dir):
    """norm_key(english) -> lang -> set of texts this tool has filed."""
    produced = collections.defaultdict(lambda: collections.defaultdict(set))
    folder = os.path.join(out_dir, "pending")
    if not os.path.isdir(folder):
        return produced
    for name in sorted(os.listdir(folder)):
        if not name.endswith(".result.json"):
            continue
        try:
            with open(os.path.join(folder, name), "r", encoding="utf-8") as handle:
                payload = json.load(handle)
        except (OSError, ValueError):
            continue
        for row in payload.get("rows") or []:
            key = norm_key(row.get("en", ""))
            if not key:
                continue
            for code, cell in (row.get("targets") or {}).items():
                text = tidy((cell or {}).get("text"))
                if text:
                    produced[key][code].add(text)
    return produced


def recover_from_filed(produced, merged):
    """code -> pairs, for merged corrections whose 'before' we filed ourselves."""
    recovered = collections.defaultdict(list)
    for key, code, english, after in merged:
        candidates = [t for t in produced.get(key, {}).get(code, ()) if t != after]
        if not candidates:
            continue
        # Longest first is a stable choice when one English was filed twice;
        # they differ by wording, not by which is more recent.
        before = sorted(candidates, key=lambda t: (-len(t), t))[0]
        recovered[code].append({
            "en": english, "before": before, "after": after,
            "category": classify(before, after), "source": "filed batch",
            "ambiguous": len(candidates) > 1,
        })
    return recovered


# ---------------- pre-merge history (the other 97% of the "before" side) ----

def _svn_cat(design_dir, name, spec, out_path):
    """One historical copy of a workbook, out of SVN. `spec` is anything svn
    -r takes: a revision number, or a date as {YYYY-MM-DD}."""
    target = os.path.join(design_dir, name)
    peg = spec if spec.startswith(("{", "r")) or spec.isdigit() else "{%s}" % spec
    peg = peg[1:] if peg.startswith("r") else peg
    with open(out_path, "wb") as handle:
        proc = subprocess.run(["svn", "cat", "-r", peg, target],
                              stdout=handle, stderr=subprocess.PIPE, timeout=300)
    if proc.returncode != 0 or os.path.getsize(out_path) == 0:
        raise RuntimeError((proc.stderr or b"").decode("utf-8", "replace").strip()
                           or "svn cat returned nothing")


def recover_premerge(design_dir, specs, merged, scratch_dir):
    """Find what WE had before each already-applied correction was merged.

    Why this is needed at all: the corrections have mostly been merged into
    the master workspace already, so comparing a delivery against the CURRENT
    workbook shows them as identical and yields almost no before/after signal -
    79 pairs across six languages on the first nine deliveries, with Dutch and
    Swedish down to 5 each. That is too thin to distil a style from.

    The pre-merge text is still in SVN, so we fetch a couple of older copies
    and look for a translation that differs from what the vendor approved.
    Measured on this repo: two revisions recover 457 further pairs (79 -> 536),
    and 53 of them are strings that had shipped completely untranslated.

    `merged` is [(key, lang, english, after)] for corrections whose text the
    live workbook already matches. Returns (pairs_by_lang, per_spec_counts)."""
    return _match_premerge(_fetch_histories(design_dir, specs, scratch_dir), merged)


def _fetch_histories(design_dir, specs, scratch_dir):
    histories = []
    for spec in specs:
        paths = []
        for name in TM_FILES:
            out = os.path.join(scratch_dir, "%s_%s" % (
                re.sub(r"[^A-Za-z0-9]", "", spec), name))
            try:
                _svn_cat(design_dir, name, spec, out)
                paths.append(out)
            except Exception as exc:                  # noqa: BLE001 - reported
                raise RuntimeError("svn cat -r %s %s: %s" % (spec, name, exc))
        histories.append((spec, _read_translation_sets(paths)))
    return histories


def _match_premerge(histories, merged):
    """Split from the fetch so the matching rule is testable without a network
    or a working copy - see selfcheck."""
    recovered = collections.defaultdict(list)
    per_spec = collections.Counter()
    for key, lang, english, after in merged:
        # Newest history first, so the "before" is the most recent thing we
        # actually shipped rather than an ancient draft of it.
        for spec, memory in histories:
            older = sorted((memory.get(key, {}).get(lang) or set()) - {after})
            if not older:
                continue
            before = older[0]
            recovered[lang].append({
                "en": english, "before": before, "after": after,
                "category": classify(before, after),
                "source": "svn @ %s" % spec,
                "was_untranslated": norm_key(before) == norm_key(english),
                "ambiguous": len(older) > 1,
            })
            per_spec[spec] += 1
            break
    return recovered, per_spec


# ---------------- outputs ----------------

def _dump(path, payload):
    with open(path, "w", encoding="utf-8") as handle:
        json.dump(payload, handle, ensure_ascii=False, indent=1, sort_keys=True)


def write_outputs(out_dir, approved, pairs, report):
    for sub in ("approved", "pairs"):
        os.makedirs(os.path.join(out_dir, sub), exist_ok=True)
    for code in TARGETS:
        _dump(os.path.join(out_dir, "approved", code + ".json"), approved.get(code, {}))
        _dump(os.path.join(out_dir, "pairs", code + ".json"), pairs.get(code, []))
    _dump(os.path.join(out_dir, "ambiguous.json"), report.get("ambiguous", {}))
    _dump(os.path.join(out_dir, "report.json"), report)
    with open(os.path.join(out_dir, "report.md"), "w", encoding="utf-8") as handle:
        handle.write(render_report(report))


def render_report(report):
    lines = ["# Localization corrections import", ""]
    lines.append("Source folder: `%s`" % report["source_dir"])
    lines.append("Translation memory read from: %s"
                 % (", ".join(report["tm_files"]) or "(none)"))
    lines.append("")
    lines.append("## Files")
    lines.append("")
    lines.append("| File | Strict | Yellow cells | Header |")
    lines.append("|---|---|---|---|")
    for entry in report["files"]:
        header = ", ".join(
            ("%s:r%s" % (s["sheet"], s["header"])) if s["header"] else
            ("%s:NO HEADER" % s["sheet"]) for s in entry["sheets"])
        lines.append("| %s | %s | %d | %s |"
                     % (entry["file"], "yes" if entry["strict"] else "",
                        entry["yellow"], header))
    lines.append("")
    lines.append("## Totals")
    lines.append("")
    for label, value in report["totals"].items():
        lines.append("- %s: **%s**" % (label.replace("_", " "), value))
    lines.append("")
    history = report.get("history") or {}
    if history.get("revisions"):
        lines.append("## Pre-merge history mined")
        lines.append("")
        lines.append("Revisions asked for: %s" % ", ".join(history["revisions"]))
        for spec, count in history.get("recovered_per_revision", {}).items():
            lines.append("- `%s` supplied the \"before\" for **%d** correction(s)"
                         % (spec, count))
        lines.append("")
    lines.append("## Per language")
    lines.append("")
    lines.append("| Lang | Approved | Before/after | Typographic | Casing | Punctuation | Wording |")
    lines.append("|---|---|---|---|---|---|---|")
    for code in TARGETS:
        row = report["per_language"][code]
        lines.append("| %s | %d | %d | %d | %d | %d | %d |"
                     % (code, row["approved"], row["pairs"], row["typographic"],
                        row["casing"], row["punctuation"], row["wording"]))
    lines.append("")
    if report["fill_colours"]:
        lines.append("## Every fill colour seen")
        lines.append("")
        lines.append("Only yellow counts as a correction; the rest is section banding.")
        lines.append("")
        for token, count in sorted(report["fill_colours"].items(), key=lambda kv: -kv[1]):
            marker = "  <- corrections" if token == "RGB:" + YELLOW else ""
            lines.append("- `%s` x%d%s" % (token, count, marker))
        lines.append("")
    if report["warnings"]:
        lines.append("## Warnings")
        lines.append("")
        for warning in report["warnings"]:
            lines.append("- %s" % warning)
        lines.append("")
    return "\n".join(lines)


# ---------------- the run ----------------

def run(src_dir, design_dir, out_dir, history=None, svn_design_dir=None):
    files = sorted(p for p in glob.glob(os.path.join(src_dir, "*.xlsx"))
                   if not os.path.basename(p).startswith("~$"))
    if not files:
        raise SystemExit("No .xlsx files in %s" % src_dir)

    scratch = tempfile.mkdtemp(prefix="loc_import_")
    raw = []
    file_reports = []
    colours = collections.Counter()
    warnings = []

    for path in files:
        result = read_corrections(path, scratch)
        raw.extend(result["corrections"])
        colours.update(result["colours"])
        yellow = len(result["corrections"])
        file_reports.append({"file": result["file"], "strict": result["strict"],
                             "sheets": result["sheets"], "yellow": yellow})
        if result["strict"]:
            warnings.append("%s is ISO-Strict OOXML; converted before reading. Without "
                            "that conversion it reads as an empty workbook."
                            % result["file"])
        if yellow == 0:
            warnings.append("%s contained NO yellow cells - check the vendor still "
                            "highlights with %s." % (result["file"], YELLOW))
        for sheet in result["sheets"]:
            if sheet["header"] is None:
                warnings.append("%s / %s: no English+German header row found; sheet skipped."
                                % (result["file"], sheet["sheet"]))

    orphans = [r for r in raw if r["english"] is None]
    if orphans:
        warnings.append("%d yellow cell(s) had no English in their row and were skipped."
                        % len(orphans))

    best, conflicts, ambiguous = dedupe(raw)
    if conflicts:
        warnings.append("%d correction(s) differed between file generations; kept the "
                        "newer one. See conflicts in report.json." % len(conflicts))
    if ambiguous:
        warnings.append("%d English string(s) were corrected two different ways inside "
                        "one delivery, so the English is not a usable key. Held out of "
                        "the approved memory; see ambiguous.json." % len(ambiguous))

    memory, tm_files = load_our_translations(design_dir)
    if not tm_files:
        warnings.append("No localization workbooks found in %s - every correction will "
                        "be recorded as new content." % design_dir)

    approved = collections.defaultdict(dict)
    pairs = collections.defaultdict(list)
    counts = collections.Counter()
    per_lang = {code: collections.Counter() for code in TARGETS}
    merged = []

    for (key, code), record in best.items():
        # occurrences lets the translate step tell a string corrected once from
        # one the vendor has confirmed a dozen times over several deliveries.
        approved[code][key] = {"en": record["english"], "text": record["after"],
                               "source": record["source"],
                               "occurrences": record.get("occurrences", 1)}
        per_lang[code]["approved"] += 1
        ours = memory.get(key, {}).get(code)
        if not ours:
            counts["new_content"] += 1
            continue
        if record["after"] in ours:
            counts["already_merged"] += 1
            # Our current text already IS the corrected text, so there is no
            # before/after to learn from here - unless we go and find what we
            # had before it was merged. See recover_premerge.
            merged.append((key, code, record["english"], record["after"]))
            continue
        # Where our workbook holds several translations for one English string
        # the "before" is ambiguous - flag it rather than pretend otherwise.
        before = sorted(ours)[0]
        category = classify(before, record["after"])
        pairs[code].append({"en": record["english"], "before": before,
                            "after": record["after"], "category": category,
                            "source": record["source"], "ambiguous": len(ours) > 1})
        counts["before_after_pairs"] += 1
        per_lang[code]["pairs"] += 1
        per_lang[code][category] += 1

    # Our own filed batches first: free, exact, and no network.
    # Captured before the filed-batch recovery below, which adds pairs for
    # corrections ALREADY counted in already_merged - without this the derived
    # "matched to our text" double-counts them and exceeds the correction total.
    workbook_pairs = counts["before_after_pairs"]
    filed_pairs = 0
    if merged:
        from_filed = recover_from_filed(load_filed_results(out_dir), merged)
        answered = set()
        for code, rows in from_filed.items():
            pairs[code].extend(rows)
            per_lang[code]["pairs"] += len(rows)
            for row in rows:
                per_lang[code][row["category"]] += 1
                answered.add((row["en"], code))
            filed_pairs += len(rows)
        counts["recovered_from_filed"] = filed_pairs
        counts["before_after_pairs"] += filed_pairs
        # Whatever a filed batch answered does not need SVN as well.
        merged = [m for m in merged if (m[2], m[1]) not in answered]

    history_specs = [s.strip() for s in (history or "").split(",") if s.strip()]
    history_counts = {}
    if history_specs and merged:
        try:
            recovered, history_counts = recover_premerge(
                svn_design_dir or design_dir, history_specs, merged, scratch)
        except Exception as exc:                      # noqa: BLE001 - reported
            recovered = {}
            warnings.append("Pre-merge history unavailable, so only the %d pair(s) "
                            "visible against the current workbooks were kept - %s"
                            % (counts["before_after_pairs"], exc))
        for code, rows in recovered.items():
            pairs[code].extend(rows)
            per_lang[code]["pairs"] += len(rows)
            for row in rows:
                per_lang[code][row["category"]] += 1
            counts["recovered_from_history"] += len(rows)
            counts["shipped_untranslated"] += sum(
                1 for row in rows if row["was_untranslated"])

    report = {
        "source_dir": src_dir,
        "design_dir": design_dir,
        "tm_files": tm_files,
        "files": file_reports,
        "fill_colours": dict(colours),
        "totals": {
            "workbooks": len(files),
            "yellow_cells": len(raw),
            "unique_corrections": len(best),
            "matched_to_our_text": counts["already_merged"] + workbook_pairs,
            "already_merged": counts["already_merged"],
            "before_after_pairs": counts["before_after_pairs"],
            "recovered_from_filed": counts.get("recovered_from_filed", 0),
            "recovered_from_history": counts["recovered_from_history"],
            "total_pairs": counts["before_after_pairs"] + counts["recovered_from_history"],
            "shipped_untranslated": counts["shipped_untranslated"],
            "new_content": counts["new_content"],
            "ambiguous_keys_held_out": len(ambiguous),
        },
        "history": {"revisions": history_specs, "recovered_per_revision": dict(history_counts)},
        "ambiguous": ambiguous,
        "per_language": {code: {k: per_lang[code][k] for k in
                                ("approved", "pairs", "typographic", "casing",
                                 "punctuation", "wording")} for code in TARGETS},
        "conflicts": conflicts,
        "warnings": warnings,
    }
    write_outputs(out_dir, approved, pairs, report)
    return report


def main():
    parser = argparse.ArgumentParser(
        description="Import yellow-highlighted translator corrections.")
    parser.add_argument("--src", help="folder of vendor correction .xlsx files")
    parser.add_argument("--design", help="a DesignData workspace's design/ folder")
    parser.add_argument("--out", help="output folder (default: <src>/../Localization)")
    parser.add_argument("--history", help="comma-separated SVN revisions or dates to mine "
                                          "for pre-merge text, newest first, e.g. "
                                          "2026-07-01,2026-06-01 or 131550,129996")
    parser.add_argument("--svn-design", help="working-copy design/ folder to run svn cat "
                                             "against (default: --design)")
    parser.add_argument("--json", action="store_true", help="print the report as JSON")
    parser.add_argument("--selfcheck", action="store_true", help="run the built-in checks")
    args = parser.parse_args()

    if args.selfcheck:
        selfcheck()
        return
    if not args.src or not args.design:
        parser.error("--src and --design are required (or use --selfcheck)")

    out_dir = args.out or os.path.join(os.path.dirname(os.path.abspath(args.src)),
                                       "Localization")
    os.makedirs(out_dir, exist_ok=True)
    report = run(args.src, args.design, out_dir,
                 history=args.history, svn_design_dir=args.svn_design)

    if args.json:
        print(json.dumps(report, ensure_ascii=False))
        return
    print(render_report(report))
    print("Written to %s" % out_dir)


# ---------------- self-check ----------------

def _make_probe(path, header_row, short_alpha, strict):
    """A miniature vendor delivery exercising every trap at once."""
    from openpyxl.styles import PatternFill
    from openpyxl.styles.colors import Color

    workbook = openpyxl.Workbook()
    sheet = workbook.active
    if header_row == 2:
        sheet.cell(row=1, column=1, value="Trivia Challenge")   # title row above the header
    for col, label in enumerate(["English ", "German", " Spanish", "Dutch",
                                 "Swedish", "French", "DANISH"], start=1):
        sheet.cell(row=header_row, column=col, value=label)

    yellow = PatternFill(fill_type="solid",
                         start_color=("FFFF00" if short_alpha else "FFFFFF00"))
    banding = PatternFill(fill_type="solid", start_color=Color(theme=5, tint=0.4))
    blue = PatternFill(fill_type="solid", start_color="FFDDEBF7")
    base = header_row

    sheet.cell(row=base + 1, column=1, value="Collect")
    sheet.cell(row=base + 1, column=2, value="Einsammeln").fill = yellow   # real correction
    sheet.cell(row=base + 1, column=3, value="Recoger").fill = blue        # banding, not a fix
    sheet.cell(row=base + 1, column=4, value="Verzamel")                   # untouched

    sheet.cell(row=base + 2, column=1, value="'Starry Night'").fill = banding
    sheet.cell(row=base + 2, column=2, value="„Sternennacht“").fill = yellow

    sheet.cell(row=base + 3, column=1, value="Brand New String")
    sheet.cell(row=base + 3, column=7, value="Helt ny tekst").fill = yellow  # new content

    # The same English corrected two ways inside one file: a trivia answer,
    # not a key. Must be held out rather than resolved by whichever sorts first.
    sheet.cell(row=base + 4, column=1, value="8")
    sheet.cell(row=base + 4, column=2, value="8").fill = yellow
    sheet.cell(row=base + 5, column=1, value="8")
    sheet.cell(row=base + 5, column=2, value="9").fill = yellow

    workbook.save(path)
    workbook.close()

    if strict:
        # Turn the saved file INTO Strict, so the check exercises the real
        # failure mode rather than a mock of it.
        temp = path + ".strict"
        _rewrite_namespaces(path, temp, to_strict=True)
        os.replace(temp, path)


def selfcheck():
    scratch = tempfile.mkdtemp(prefix="loc_selfcheck_")
    src = os.path.join(scratch, "src")
    design = os.path.join(scratch, "design")
    os.makedirs(src)
    os.makedirs(design)

    # Our side: both strings already translated, both differently from the
    # vendor's correction - one a wording change, one purely typographic.
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "localize"
    sheet.append(["key", "English", "German", "Spanish", "Dutch",
                  "Swedish", "French", "Danish"])
    sheet.append(["BTN_COLLECT", "Collect", "Sammeln", "Recoger", "Verzamel",
                  "Ta emot", "Collecter", "Indsaml"])
    sheet.append(["QUIZ_PAINTING", "'Starry Night'", "\"Sternennacht\"",
                  "x", "x", "x", "x", "x"])
    workbook.save(os.path.join(design, "localization.xlsx"))
    workbook.close()

    _make_probe(os.path.join(src, "Rel_1-PR.xlsx"), header_row=1,
                short_alpha=False, strict=False)
    _make_probe(os.path.join(src, "Rel_2-PR-Updated.xlsx"), header_row=2,
                short_alpha=True, strict=True)

    out = os.path.join(scratch, "out")
    report = run(src, design, out)
    totals = report["totals"]

    strict_files = [f for f in report["files"] if f["strict"]]
    assert len(strict_files) == 1, "Strict workbook not detected: %s" % report["files"]

    # Both probes carry the same five yellow cells. Ten raw cells proves the
    # Strict file AND the short-alpha yellow were both actually read; three
    # unique proves dedupe collapsed the generations and held the ambiguous
    # "8" out instead of picking one of its two corrections.
    assert [f["yellow"] for f in report["files"]] == [5, 5], report["files"]
    assert totals["yellow_cells"] == 10, totals
    assert totals["unique_corrections"] == 3, totals
    assert totals["before_after_pairs"] == 2, totals
    assert totals["new_content"] == 1, totals
    assert totals["ambiguous_keys_held_out"] == 1, totals

    with open(os.path.join(out, "pairs", "DE.json"), encoding="utf-8") as handle:
        de_pairs = json.load(handle)
    categories = {p["en"]: p["category"] for p in de_pairs}
    assert categories.get("Collect") == "wording", categories
    assert categories.get("'Starry Night'") == "typographic", categories

    with open(os.path.join(out, "approved", "DE.json"), encoding="utf-8") as handle:
        approved_de = json.load(handle)
    assert "8" not in approved_de, "an ambiguous key leaked into the approved memory"
    assert approved_de["collect"]["occurrences"] == 2, approved_de["collect"]

    with open(os.path.join(out, "ambiguous.json"), encoding="utf-8") as handle:
        held_out = json.load(handle)
    assert held_out["8|DE"]["variants"] == ["8", "9"], held_out

    with open(os.path.join(out, "approved", "DA.json"), encoding="utf-8") as handle:
        approved_da = json.load(handle)
    assert approved_da["brand new string"]["text"] == "Helt ny tekst", approved_da

    # The blue-banded Spanish cell must NOT have been taken as a correction.
    with open(os.path.join(out, "approved", "ES.json"), encoding="utf-8") as handle:
        assert json.load(handle) == {}, "a non-yellow fill was imported as a correction"

    assert os.path.exists(os.path.join(out, "report.md"))

    # --- pre-merge recovery, without a network or a working copy ---
    # Newest history wins; a revision that already matches the approved text
    # supplies no "before" and must be skipped rather than counted.
    histories = [
        ("newest", {"collect": {"DE": {"Einsammeln"}},
                    "next": {"DE": {"Weiter alt"}}}),
        ("oldest", {"collect": {"DE": {"Sammeln"}},
                    "revert": {"DE": {"Revert your previous move?"}}}),
    ]
    merged = [("collect", "DE", "Collect", "Einsammeln"),
              ("next", "DE", "Next", "Weiter"),
              ("revert", "DE", "Revert your previous move?", "Letzten Zug rückgängig machen?"),
              ("absent", "DE", "Absent", "Fehlt")]
    rows, per_spec = _match_premerge(histories, merged)
    got = {r["en"]: r for r in rows["DE"]}
    assert set(got) == {"Collect", "Next", "Revert your previous move?"}, sorted(got)
    # 'newest' matched Collect but held the SAME text, so the older revision
    # had to supply that before.
    assert got["Collect"]["before"] == "Sammeln", got["Collect"]
    assert got["Collect"]["source"] == "svn @ oldest", got["Collect"]
    assert got["Next"]["source"] == "svn @ newest", got["Next"]
    # A string that shipped with the English still in it is flagged: that is a
    # different class of miss from a wording preference.
    assert got["Revert your previous move?"]["was_untranslated"] is True
    assert got["Next"]["was_untranslated"] is False
    assert per_spec["oldest"] == 2 and per_spec["newest"] == 1, dict(per_spec)

    print("selfcheck OK - %d files, %d yellow cells, %d unique, %d before/after pairs"
          % (totals["workbooks"], totals["yellow_cells"],
             totals["unique_corrections"], totals["before_after_pairs"]))


if __name__ == "__main__":
    main()
