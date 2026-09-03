"""Parses the real Feishu roadmap export and joins it against real config
workbooks. Skips cleanly if either isn't present on this machine."""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
import sprint_board as sb

ROADMAP = r"C:\SkipBo\tools\EventAutomationDesktop\ideationData\Roadmap\Roadmap.xlsx"
DESIGN  = r"C:\SkipBo\branches\Misc\design\DesignData\WorkSpcae_Design\design"

def check(c, m):
    if not c:
        print("FAIL:", m); sys.exit(1)

if not os.path.exists(ROADMAP):
    print("SKIP: roadmap export not present"); sys.exit(0)

sprints = sb.read_roadmap(ROADMAP, 2026)
check(len(sprints) >= 20, f"expected ~22 sprint columns, got {len(sprints)}")
print(f"OK  parsed {len(sprints)} columns from Strict-OOXML roadmap (openpyxl cannot read this file at all)")

by = {s["name"]: s for s in sprints}

# Store Submission columns are milestones, not sprints.
ms = [s for s in sprints if s["is_milestone"]]
check(len(ms) == 3, f"expected 3 Store Submission milestones, got {len(ms)}")
print(f"OK  {len(ms)} Store Submission milestones flagged, not offered as sprints")

# Window comes from the primary event, not the longer overlay campaign.
aug1 = by["Aug Event 1"]
check(aug1["start"] == "2026-08-06" and aug1["end"] == "2026-08-19",
      f"Aug Event 1 window wrong: {aug1['start']}..{aug1['end']}")
check([o["title"] for o in aug1["overlays"]] == ["Alice Themed Event"],
      f"overlay not separated: {aug1['overlays']}")
print("OK  window = primary event (08-06..08-19); 'Alice Themed Event' kept as an overlay")

# The double-dash typo in the source sheet must still map to a tab.
aug2 = by["Aug Event 2"]
streak = [u for u in aug2["updates"] if u["tab"] == "streak_challenge"]
check(streak, "'- - Streak Challenge refresh' (source typo) failed to map")
print("OK  source typo '- - Streak Challenge refresh' still maps to the Streak tab")

if not os.path.isdir(DESIGN):
    print("SKIP: Misc design data not present; roadmap-only checks passed"); sys.exit(0)

# Date join: roadmap marketing name -> internal config code.
st = sb.sprint_status(DESIGN, aug1)
ev = sb._match_event_config(DESIGN, aug1)
check(ev and ev["alias"].startswith("Dessert2608"), f"date join failed: {ev}")
check(ev["exact"], "Aug Event 1 should match its config window exactly")
print(f"OK  date join: 'Five-Star Picnic' -> {ev['alias']} (exact)")

counts = {t["tab"]: t["count"] for t in st["tabs"]}
check(counts["minigames"] == 5, f"expected 5 minigames in the window, got {counts['minigames']}")
check(counts["full_screen_gift"] <= 2,
      f"full_screen_gift not filtered to 栏位归属=16 (got {counts['full_screen_gift']})")
print(f"OK  per-tab counts sane (minigames=5, full-screen gift={counts['full_screen_gift']} after slot-16 filter)")

# Drift the board must surface.
sep1 = by["Sep Event 1"]
check(sb._match_event_config(DESIGN, sep1) is None,
      "Sep Event 1 unexpectedly has a configured event")
todo = [t["label"] for t in sb.sprint_status(DESIGN, sep1)["tabs"] if t["status"] == "todo"]
check(len(todo) >= 6, f"expected many TODOs for the unconfigured sprint, got {todo}")
print(f"OK  drift found: 'Sort Event' planned but not configured ({len(todo)} tabs outstanding)")

sep2 = by["Sep Event 2"]
ev2 = sb._match_event_config(DESIGN, sep2)
check(ev2 is not None, f"Sep Event 2 lost its matched config entirely: {ev2}")
print(f"OK  Sep Event 2 joins to its config ({ev2['alias']})")

# A configured event whose dates DISAGREE with the roadmap must come back
# inexact, so the board can say so.
#
# The sprint is searched for rather than named. This was pinned to Sep Event 2
# and its 09-28 vs 09-30 disagreement, so the assertion was really claiming that
# production design data still contains a specific mistake - and it failed the
# day someone corrected those dates, which is the opposite of what a test should
# reward. Searching keeps the real subject (an inexact match is reported as
# inexact) and survives the data being fixed or new drift turning up elsewhere.
inexact = [(s["name"], m) for s in sprints
           if (m := sb._match_event_config(DESIGN, s)) and not m["exact"]]
if inexact:
    nm, m = inexact[0]
    check(m.get("start") and m.get("end"),
          f"{nm} is flagged inexact but carries no dates to compare: {m}")
    print(f"OK  drift found: {nm} planned ->{by[nm]['end']} but "
          f"{m['alias']} configured ->{m['end']}")
else:
    # Every configured event agreeing with the roadmap is the desired state, so
    # this is not a failure - but it is announced, because a check that quietly
    # stops covering anything is worse than one that fails.
    print("SKIP  no configured event currently disagrees with the roadmap, so "
          "the inexact-match path was not exercised this run")

# Chapter Update is matched by identity, not by date window: the roadmap names
# the chapter ("Chapter 82 + Star Chest") and the check is simply whether that
# number exists in pve_chapter.xlsx. Both directions are pinned here because a
# regex that claimed too much (any line saying "chapter") or too little (only
# one exact phrasing) silently breaks planned-vs-done for every sprint.
check(sb._planned_chapters(aug1) == [82],
      f"Aug Event 1 should plan chapter 82, got {sb._planned_chapters(aug1)}")
ch_aug = [t for t in st["tabs"] if t["tab"] == "chapter"][0]
check(ch_aug["status"] == "done",
      f"chapter 82 exists in pve_chapter.xlsx, expected done, got {ch_aug['status']}")
check(all("hapter" not in o for o in st["other_updates"]),
      f"chapter line still listed as unactionable: {st['other_updates']}")
ch_sep1 = [t for t in sb.sprint_status(DESIGN, sep1)["tabs"] if t["tab"] == "chapter"][0]
check(ch_sep1["planned"], "Sep Event 1 should plan a chapter")
print(f"OK  chapter tab: Aug=ch82 {ch_aug['status']}, "
      f"Sep1=ch{sb._planned_chapters(sep1)} {ch_sep1['status']}, no longer unactionable")

# A line mentioning chapters without naming one must NOT be claimed, or it
# becomes a permanent TODO the tab can never satisfy.
_e, _u = sb._parse_liveops("update:" + chr(10) + "- Chapter art theme refresh", 2026)
check(_u[0]["tab"] is None, f"a numberless chapter line was wrongly claimed: {_u[0]}")
print("OK  numberless 'Chapter art theme refresh' stays unclaimed")

print("\nALL SPRINT BOARD TESTS PASSED")
