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
check(ev2 and not ev2["exact"], f"expected a date mismatch on Sep Event 2, got {ev2}")
print(f"OK  drift found: Hotel Sim planned ->{sep2['end']} but {ev2['alias']} configured ->{ev2['end']}")

print("\nALL SPRINT BOARD TESTS PASSED")
