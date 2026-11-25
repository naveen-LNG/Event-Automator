"""Step 10 (update_events_excel), 活动buff常量: col 1 keys must take the GUI's event id.

swap_name's digit fallback rewrites the month (202610 -> 202611), and the old code
tested the source-id prefix only AFTER that swap. Every cross-month clone then kept a
month-shifted id: Blocks2610 -> Tiles2611 got 20261101_* (Merge2611's id) instead of
20261102_*, and Tiles2611EventCtrl.lua:407 crashed on every PvE game end.

Runs the real update_events_excel on a synthetic events.xlsx in a temp dir.
"""
import json
import os
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
import openpyxl
import event_pipeline as ep

FAILS = []
BUFF, BACKUP = "活动buff常量", "buff老活动备份"


def check(ok, msg, detail=""):
    print(("OK   " if ok else "FAIL ") + msg)
    if not ok:
        if detail:
            print("     " + str(detail)[:300])
        FAILS.append(msg)


def run_clone(src_evt, tgt_evt, event_id):
    """Clone src_evt's buff rows (from the backup sheet) into tgt_evt; return main-sheet col 1/2."""
    with tempfile.TemporaryDirectory() as td:
        design = os.path.join(td, "design", "DesignData", ep.WORKSPACE_NAME, "design")
        os.makedirs(design)
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = BUFF
        for row in (["//名字", "值"],
                    ["//正常双周活动(Merge2610)", None],
                    ["20261015_MERGE2610_BUFF", "1"]):
            ws.append(row)
        bk = wb.create_sheet(BACKUP)
        for row in (["//正常双周活动(Blocks2610)", None],
                    ["20261001_BLOCKS2610_GAME_END", "1"],
                    ["20261001_BLOCKS2610_SCORE", "Blocks2610Score"]):
            bk.append(row)
        wb.save(os.path.join(design, "events.xlsx"))
        with open(os.path.join(td, "extra_args.json"), "w") as f:
            json.dump({"event_id": event_id}, f)

        ok = ep.update_events_excel(td, src_evt, tgt_evt, td, 10)
        check(ok, f"{src_evt} -> {tgt_evt}: update_events_excel returns True")
        out = openpyxl.load_workbook(os.path.join(design, "events.xlsx"))[BUFF]
        return [(out.cell(r, 1).value, out.cell(r, 2).value) for r in range(1, out.max_row + 1)
                if out.cell(r, 1).value]


def main():
    # Cross-month: the GUI's id wins, not the month-shifted source id (20261101).
    rows = run_clone("Blocks2610", "Tiles2611", "20261102")
    keys = [k for k, _ in rows]
    check("20261102_TILES2611_GAME_END" in keys and "20261102_TILES2611_SCORE" in keys,
          "cross-month: keys are 20261102_<key>", rows)
    check(not any(str(k).startswith("20261101") for k in keys),
          "cross-month: no month-shifted 20261101_* key", rows)
    check(("20261102_TILES2611_SCORE", "Tiles2611Score") in rows,
          "cross-month: other columns still get the name swap", rows)
    check("//正常双周活动(Tiles2611)" in keys, "cross-month: group header renamed", rows)

    # Same month: digit fallback is a no-op; the GUI's id must still be applied.
    rows = run_clone("Blocks2610", "Tiles2610", "20261015")
    keys = [k for k, _ in rows]
    check("20261015_TILES2610_GAME_END" in keys and "20261015_TILES2610_SCORE" in keys,
          "same-month: keys are 20261015_<key>", rows)

    print(f"\n{len(FAILS)} failure(s)" if FAILS else "\nall passed")
    sys.exit(1 if FAILS else 0)


if __name__ == "__main__":
    main()
