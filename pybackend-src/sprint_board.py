"""
Sprint landing page backend: reads the studio's Feishu roadmap export and
joins it against what is actually configured in the design workbooks, so a
designer can pick a sprint and immediately see what still needs doing.

Two halves, deliberately kept separate:

1. read_roadmap()  - the PLAN. Parsed from ideationData/Roadmap/Roadmap.xlsx,
   exported out of Feishu. That file is *Strict* OOXML (conformance="strict",
   purl.oclc.org namespaces), which openpyxl cannot read at all - it silently
   reports zero sheets. So the sheet XML is parsed directly here, which also
   happens to handle the ordinary transitional namespace, meaning either
   export flavour works and nobody has to remember a save-as setting.

   The sheet is TRANSPOSED: one column per sprint, with
     row 2 = sprint name ("Aug Event 2")
     row 3 = launch date as M.DD
     row 4 = update method ("Patch" | "Store Submission")
     row 5 = free-text LiveOps block: an "Event:" list then a
             "Main Regular Updates:" list.

2. sprint_status() - the ACTUAL. Scans the same design workbooks the rest of
   this tool writes to, matching rows whose date range overlaps the sprint
   window, and reports planned-vs-configured per tool tab.

Dates are the join key between the two, never names: the roadmap uses
marketing names ("Anniversary Cruise") while events.xlsx uses internal codes
("Merge2608活动"), but their date ranges match exactly.

Deliberately NOT parsed: the cycle numbers in roadmap lines like "Battle Pass
2608" / "Holiday BP 2615". Their meaning is ambiguous (2621 rules out YYMM)
and this tool already derives the real next cycle ID from the workbooks in
get_bp_info.py / get_hbp_info.py. A roadmap line is therefore treated purely
as "this update is due in this sprint" - a presence signal, not an ID.
"""
import sys
import os
import re
import json
import zipfile
import datetime as dt
from xml.etree import ElementTree as ET

import openpyxl

# Roadmap line -> the tool tab that does the work. Matched case-insensitively
# as substrings, so "- - Streak Challenge refresh" (a real double-dash typo in
# the source sheet) and minor wording drift still land correctly.
UPDATE_KEYWORDS = [
    ("minigames",        ["minigame"]),
    ("trivia",           ["trivia challenge questions", "trivia questions"]),
    ("battlepass",       ["battle pass"]),
    ("holiday_bp",       ["holiday bp"]),
    ("decoration_shop",  ["decoration shop"]),
    ("streak_challenge", ["streak challenge"]),
    ("weekly_gift",      ["weekly offer"]),
    ("full_screen_gift", ["full screen offer", "full-screen offer"]),
    ("endless_treasure", ["endless discount", "endless treasure"]),
    # A compiled pattern rather than a substring: "Chapter 84 + Star Chest" is
    # actionable, a bare mention of "chapter" is not, and claiming the latter
    # would strand it in a tab that can never satisfy it instead of leaving it
    # visible in the read-only list.
    ("chapter",          [re.compile(r"chapter\s*\d")]),
]

TAB_LABELS = {
    "minigames": "Minigame Configs",
    "trivia": "Trivia Challenge",
    "battlepass": "Pet Battlepass",
    "holiday_bp": "Holiday Battlepass",
    "decoration_shop": "Decoration Shop",
    "streak_challenge": "Streak Challenge",
    "weekly_gift": "Weekly Gift",
    "full_screen_gift": "Full Screen Gift",
    "endless_treasure": "Endless Treasure",
    "chapter": "Chapter Update",
}

# Where "is it configured?" is read from. Columns are 0-indexed and were
# verified against real workbooks in branches/Misc. decoration_shop is absent
# on purpose: it writes to skin_shop/icon/item/card, none of which carry a
# date column, so it cannot be detected by window overlap. Reporting it as
# "manual check" is honest; reporting it as missing would be a false alarm.
CONFIG_SOURCES = {
    "minigames":        dict(file="mini_mgr.xlsx",         sheet="活动数据",   start=2, end=3,  switch=1,    name=5),
    "trivia":           dict(file="mini_mgr.xlsx",         sheet="活动数据",   start=2, end=3,  switch=1,    name=5, only_type="答题挑战"),
    "battlepass":       dict(file="battle_pass.xlsx",      sheet="battle_pass", start=2, end=3, switch=1,    name=4),
    "holiday_bp":       dict(file="festivals_bp.xlsx",     sheet="festivals_bp", start=2, end=3, switch=1,   name=4),
    "endless_treasure": dict(file="endless_treasure.xlsx", sheet="折扣链推荐", start=3, end=4,  switch=8,    name=9),
    "weekly_gift":      dict(file="store.xlsx",            sheet="weekly_gift", start=3, end=4, switch=2,    name=1),
    # recommend_gift holds every gift type; the full-screen one is 栏位归属=16
    # (col 7). Without that filter the whole table matches and the count is
    # meaningless.
    "full_screen_gift": dict(file="store.xlsx",            sheet="recommend_gift", start=11, end=12, switch=4, name=1,
                             where={7: "16"}),
    "streak_challenge": dict(file="resident_task.xlsx",    sheet="任务链",     start=2, end=3,  switch=None, name=0),
}

DATE_RANGE_RE = re.compile(r"\((\d{1,2})[./](\d{1,2})\s*[-–—]\s*(\d{1,2})[./](\d{1,2})\)")


def log(msg):
    print(f"[SprintBoard] {msg}", file=sys.stderr, flush=True)


# ---------------- Roadmap (the plan) ----------------

def _local(tag):
    return tag.split("}", 1)[1] if "}" in tag else tag


def _read_grid(path):
    """Cell grid from an .xlsx, tolerating both Strict and transitional OOXML."""
    z = zipfile.ZipFile(path)
    shared = []
    if "xl/sharedStrings.xml" in z.namelist():
        for si in ET.fromstring(z.read("xl/sharedStrings.xml")):
            if _local(si.tag) == "si":
                shared.append("".join(t.text or "" for t in si.iter() if _local(t.tag) == "t"))

    sheet_names = [n for n in z.namelist() if n.startswith("xl/worksheets/sheet")]
    if not sheet_names:
        raise RuntimeError("No worksheet found inside the roadmap workbook.")
    root = ET.fromstring(z.read(sorted(sheet_names)[0]))

    grid = {}
    for c in root.iter():
        if _local(c.tag) != "c" or not c.get("r"):
            continue
        holder = next((ch for ch in c if _local(ch.tag) in ("v", "is")), None)
        if holder is None:
            continue
        if _local(holder.tag) == "is":
            val = "".join(x.text or "" for x in holder.iter() if _local(x.tag) == "t")
        else:
            raw = holder.text or ""
            val = shared[int(raw)] if c.get("t") == "s" and raw.isdigit() and int(raw) < len(shared) else raw
        m = re.match(r"([A-Z]+)(\d+)", c.get("r"))
        if not m:
            continue
        col = 0
        for ch in m.group(1):
            col = col * 26 + (ord(ch) - 64)
        grid[(int(m.group(2)), col)] = val
    return grid


def _parse_liveops(block, year_hint):
    """Split the free-text LiveOps cell into event entries and update entries."""
    events, updates, section = [], [], None
    for raw in (block or "").split("\n"):
        line = raw.strip()
        if not line:
            continue
        low = line.lower()
        if low.startswith("event:"):
            section = "event"; continue
        if low.startswith("main regular updates") or low.startswith("update:"):
            section = "update"; continue

        text = line.lstrip("-").strip()
        # "- - Streak Challenge refresh" appears verbatim in the source sheet.
        text = text.lstrip("-").strip()
        if not text:
            continue

        if section == "event":
            m = DATE_RANGE_RE.search(text)
            entry = {"text": text, "start": None, "end": None}
            if m:
                sm, sd, em, ed = (int(x) for x in m.groups())
                sy = year_hint
                # A range that wraps to a lower month has crossed new year.
                ey = year_hint + 1 if em < sm else year_hint
                try:
                    entry["start"] = dt.date(sy, sm, sd).isoformat()
                    entry["end"] = dt.date(ey, em, ed).isoformat()
                except ValueError:
                    pass
                entry["title"] = DATE_RANGE_RE.sub("", text).strip()
            else:
                entry["title"] = text
            events.append(entry)
        elif section == "update":
            key = None
            for k, words in UPDATE_KEYWORDS:
                if any(w.search(low) if hasattr(w, "search") else w in low
                       for w in words):
                    key = k
                    break
            updates.append({"text": text, "tab": key})
    return events, updates


def read_roadmap(path, year_hint):
    """Parsed sprint columns, left to right."""
    grid = _read_grid(path)
    if not grid:
        return []
    maxcol = max(c for _, c in grid)
    sprints = []
    for col in range(2, maxcol + 1):
        name = (grid.get((2, col)) or "").strip()
        block = grid.get((5, col)) or ""
        if not name and not block.strip():
            continue
        method = (grid.get((4, col)) or "").strip()
        events, updates = _parse_liveops(block, year_hint)

        # The event's own date range is authoritative. Row 3's launch date is
        # a float with precision damage (8.1999999999999993) whose fraction is
        # a 2-digit day, so ".2" means the 20th, not the 2nd - too easy to
        # misread. It is kept only as a display hint.
        #
        # The window is the FIRST dated event, not the union of all of them.
        # A sprint can also carry a longer overlay campaign - "Aug Event 1"
        # lists both Five-Star Picnic (08.06-08.19) and Alice Themed Event
        # (08.06-09.03) - and letting the overlay stretch the window makes
        # every configured-row count bleed in from neighbouring sprints.
        dated = [e for e in events if e.get("start")]
        start = dated[0]["start"] if dated else None
        end = dated[0].get("end") if dated else None
        overlays = [e for e in dated[1:] if e.get("end") and end and e["end"] > end]

        sprints.append({
            "column": col,
            "name": name,
            "launch_hint": (grid.get((3, col)) or "").strip(),
            "method": method,
            "is_milestone": method.lower().startswith("store submission"),
            "start": start,
            "end": end,
            "events": events,
            "overlays": overlays,
            "updates": updates,
        })
    return sprints


# ---------------- Config workbooks (the actual) ----------------

def _to_date(v):
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    if isinstance(v, str):
        for f in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%Y/%m/%d %H:%M:%S", "%Y/%m/%d"):
            try:
                return dt.datetime.strptime(v.strip(), f).date()
            except ValueError:
                pass
    return None


_WB_CACHE = {}


def _sheet_rows(design_path, filename, sheet):
    key = (design_path, filename, sheet)
    if key in _WB_CACHE:
        return _WB_CACHE[key]
    p = os.path.join(design_path, filename)
    rows = []
    if os.path.exists(p):
        try:
            wb = openpyxl.load_workbook(p, data_only=True, read_only=True)
            if sheet in wb.sheetnames:
                rows = list(wb[sheet].iter_rows(min_row=2, values_only=True))
        except Exception as e:
            log(f"Warning: could not read {filename}/{sheet}: {e}")
    _WB_CACHE[key] = rows
    return rows


def _configured_rows(design_path, tab, win_start, win_end):
    src = CONFIG_SOURCES.get(tab)
    if not src:
        return None          # not auto-detectable
    hits = []
    for r in _sheet_rows(design_path, src["file"], src["sheet"]):
        if not r:
            continue
        si, ei = src["start"], src["end"]
        if len(r) <= si:
            continue
        s = _to_date(r[si])
        if not s:
            continue
        e = _to_date(r[ei]) if ei is not None and len(r) > ei else None
        e = e or s
        if s > win_end or e < win_start:
            continue
        if src.get("only_type"):
            ni = src["name"]
            if len(r) <= ni or str(r[ni] or "").strip() != src["only_type"]:
                continue
        if src.get("where"):
            ok = True
            for ci, want in src["where"].items():
                got = str(r[ci]).strip() if len(r) > ci and r[ci] is not None else ""
                # Sheet values arrive as "16" or 16 or 16.0 depending on cell type.
                if got.endswith(".0"):
                    got = got[:-2]
                if got != str(want):
                    ok = False
                    break
            if not ok:
                continue
        ni = src["name"]
        swi = src["switch"]
        hits.append({
            "name": str(r[ni] or "")[:40] if ni is not None and len(r) > ni else "",
            "start": s.isoformat(),
            "end": e.isoformat(),
            "switch": str(r[swi]) if swi is not None and len(r) > swi else "",
        })
    return hits


CHAPTER_RE = re.compile(r"chapter\s*(\d+)", re.I)


def _planned_chapters(sprint):
    """Display chapter numbers this sprint's roadmap lines ask for.

    The roadmap writes them as "Chapter 84 + Star Chest", and that number is
    the DISPLAY chapter (章节数字, col C of pve_chapter.xlsx) -- not the
    internal map id, which runs 10 ahead (94000).
    """
    out = []
    for u in sprint.get("updates", []):
        if u.get("tab") != "chapter":
            continue
        m = CHAPTER_RE.search(u.get("text") or "")
        if m:
            out.append(int(m.group(1)))
    return out


def _chapter_rows(design_path, sprint):
    """Chapter Update is the one tab checked by identity, not by date overlap.

    Every other config carries a start/end column, so "is this sprint's work
    done?" is a window question. A chapter has no dates at all -- but the
    roadmap names the exact chapter, so the far stronger check is simply
    whether that chapter number exists in pve_chapter.xlsx yet.
    """
    wanted = _planned_chapters(sprint)
    if not wanted:
        return []

    have = {}
    for r in _sheet_rows(design_path, "pve_chapter.xlsx", "chapter"):
        if not r or len(r) < 3 or r[0] is None:
            continue
        try:
            have[int(r[2])] = (int(r[0]), str(r[1] or ""))
        except (TypeError, ValueError):
            continue      # "//END" markers and other non-data rows

    hits = []
    for ch in wanted:
        if ch in have:
            map_id, alias = have[ch]
            hits.append({
                "name": f"章节{ch} · {alias}".strip(" ·"),
                "start": None,          # a chapter is not a dated window
                "end": None,
                "switch": f"map {map_id}",
            })
    return hits


def sprint_status(design_path, sprint):
    """Planned vs configured, per tool tab, for one sprint window."""
    if not sprint.get("start") or not sprint.get("end"):
        return {"tabs": [], "note": "This sprint has no parsable date range in the roadmap."}

    win_start = dt.date.fromisoformat(sprint["start"])
    win_end = dt.date.fromisoformat(sprint["end"])
    planned_tabs = {u["tab"] for u in sprint["updates"] if u.get("tab")}

    tabs = []
    for tab, label in TAB_LABELS.items():
        rows = (_chapter_rows(design_path, sprint) if tab == "chapter"
                else _configured_rows(design_path, tab, win_start, win_end))
        planned = tab in planned_tabs
        if rows is None:
            status = "manual"
        elif rows and planned:
            status = "done"
        elif rows and not planned:
            status = "extra"        # configured but the roadmap didn't ask for it
        elif planned:
            status = "todo"
        else:
            status = "n/a"
        tabs.append({
            "tab": tab, "label": label, "planned": planned,
            "status": status, "count": len(rows) if rows else 0,
            "rows": (rows or [])[:6],
        })

    # Roadmap lines with no owning tab (Star Chest on its own, Event Pack theme,
    # Webstore offers, Card Collection, surveys) - shown read-only so the
    # sprint's full scope is visible even though this tool can't action them.
    other = [u["text"] for u in sprint["updates"] if not u.get("tab")]
    return {"tabs": tabs, "other_updates": other}


def _match_event_config(design_path, sprint):
    """The events.xlsx row whose window matches this sprint, joined on dates
    (the roadmap's marketing name never appears in the config)."""
    if not sprint.get("start"):
        return None
    ws = dt.date.fromisoformat(sprint["start"])
    we = dt.date.fromisoformat(sprint["end"]) if sprint.get("end") else ws
    best = None
    for r in _sheet_rows(design_path, "events.xlsx", "events"):
        if not r or len(r) < 9 or not r[1]:
            continue
        s, e = _to_date(r[7]), _to_date(r[8])
        if not s:
            continue
        e = e or s
        if s > we or e < ws:
            continue
        overlap = (min(e, we) - max(s, ws)).days
        row = {"alias": str(r[1]), "start": s.isoformat(), "end": e.isoformat(),
               "switch": str(r[5]) if len(r) > 5 else "",
               "exact": s == ws and e == we}
        if best is None or overlap > best[0]:
            best = (overlap, row)
    return best[1] if best else None


def main():
    """CLI: emit the whole board (every sprint + the selected one's status)."""
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    if len(sys.argv) < 2:
        print(json.dumps({"error": "Usage: sprint_board <payload.json>"}))
        sys.exit(1)

    with open(sys.argv[1], "r", encoding="utf-8-sig") as f:
        payload = json.load(f)

    roadmap_path = payload.get("roadmap_path", "")
    design_path = payload.get("design_path", "")
    year_hint = int(payload.get("year_hint") or dt.date.today().year)
    selected = payload.get("selected_sprint")

    if not roadmap_path or not os.path.exists(roadmap_path):
        print(json.dumps({"error": f"Roadmap not found at {roadmap_path}. "
                                   f"Export the Feishu sheet to ideationData/Roadmap/Roadmap.xlsx."}))
        sys.exit(1)

    try:
        sprints = read_roadmap(roadmap_path, year_hint)
        if not sprints:
            print(json.dumps({"error": "No sprint columns found in the roadmap sheet."}))
            sys.exit(1)

        today = dt.date.today()
        current = None
        for sp in sprints:
            if sp.get("start") and sp.get("end"):
                if dt.date.fromisoformat(sp["start"]) <= today <= dt.date.fromisoformat(sp["end"]):
                    current = sp["name"]
                    break

        chosen = next((s for s in sprints if s["name"] == selected), None)
        if chosen is None:
            chosen = next((s for s in sprints if s["name"] == current), None) or \
                     next((s for s in sprints if not s["is_milestone"]), sprints[0])

        detail = {"sprint": chosen}
        if not chosen["is_milestone"] and design_path:
            detail.update(sprint_status(design_path, chosen))
            detail["event_config"] = _match_event_config(design_path, chosen)

        print(json.dumps({
            "sprints": [{k: sp[k] for k in ("name", "start", "end", "method", "is_milestone", "launch_hint")}
                        for sp in sprints],
            "current_sprint": current,
            "selected": chosen["name"],
            "detail": detail,
        }, ensure_ascii=False))
    except Exception as e:
        print(json.dumps({"error": str(e)}))
        sys.exit(1)
