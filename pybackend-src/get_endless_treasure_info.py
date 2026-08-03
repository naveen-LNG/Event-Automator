import sys
import os
import json
import re
from datetime import datetime, timedelta
import openpyxl

# 折扣链数据 (chain data) columns — verified against real production data
CHAIN_COL_ID = 1
CHAIN_COL_DESC = 2
CHAIN_COL_TYPE = 3
CHAIN_COL_STAGE_LIST = 5
EVENT_CHAIN_TYPE = 2  # 活动折扣链

# 折扣链推荐 (recommendation/scheduler) columns
REC_COL_ID = 1
REC_COL_TYPE = 3
REC_COL_START = 4
REC_COL_END = 5
REC_COL_END_ACTIVATE = 6
REC_COL_FORMULA = 7
REC_COL_SWITCH = 9
REC_COL_THEME = 10

TIER_SUFFIXES = ["1.99折扣链", "4.99折扣链", "9.99折扣链"]


def parse_dt(v):
    if isinstance(v, datetime):
        return v
    if isinstance(v, str):
        try:
            return datetime.strptime(v.strip(), "%Y-%m-%d %H:%M:%S")
        except ValueError:
            return None
    return None


def find_last_two_event_triplets(ws_chain):
    """Return (latest_triplet_rows, previous_triplet_rows), each a list of 3 row
    numbers in ascending ID order. The sheet has no trailing section after the
    last chain row (verified: row 201 is EOF), so the most recent 3 type-2 rows
    by row position are always the most recent event triplet.
    """
    event_rows = []
    for r in range(2, ws_chain.max_row + 1):
        if ws_chain.cell(r, CHAIN_COL_TYPE).value == EVENT_CHAIN_TYPE:
            idv = ws_chain.cell(r, CHAIN_COL_ID).value
            if isinstance(idv, (int, float)):
                event_rows.append(r)
    if len(event_rows) < 6:
        return None, None
    return event_rows[-3:], event_rows[-6:-3]


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    try:
        design_path = sys.argv[1]
    except IndexError:
        print(json.dumps({"error": "Missing design folder path"}))
        sys.exit(1)

    file_path = os.path.join(design_path, "endless_treasure.xlsx")
    if not os.path.exists(file_path):
        print(json.dumps({"error": f"File not found: {file_path}"}))
        sys.exit(1)

    try:
        wb = openpyxl.load_workbook(file_path, data_only=True)
        ws_chain = wb["折扣链数据"]
        ws_rec = wb["折扣链推荐"]

        latest_triplet, prev_triplet = find_last_two_event_triplets(ws_chain)
        if not latest_triplet:
            print(json.dumps({"error": "Fewer than 2 event-chain (类型=2) triplets found in 折扣链数据 — cannot determine slot alternation."}))
            sys.exit(1)

        latest_ids = [ws_chain.cell(r, CHAIN_COL_ID).value for r in latest_triplet]
        latest_stage_lists = [ws_chain.cell(r, CHAIN_COL_STAGE_LIST).value for r in latest_triplet]
        prev_stage_lists = [ws_chain.cell(r, CHAIN_COL_STAGE_LIST).value for r in prev_triplet]

        next_chain_ids = [int(latest_ids[2]) + 1, int(latest_ids[2]) + 2, int(latest_ids[2]) + 3]

        # Find the active (switch=TRUE) recommendation row, or fall back to the
        # most recent by start date if none is currently flagged active.
        active_row = None
        active_start = None
        for r in range(2, ws_rec.max_row + 1):
            s = parse_dt(ws_rec.cell(r, REC_COL_START).value)
            if s and (active_start is None or s > active_start):
                if str(ws_rec.cell(r, REC_COL_SWITCH).value or "").strip().upper() == "TRUE":
                    active_row = r
                    active_start = s
        if active_row is None:
            for r in range(2, ws_rec.max_row + 1):
                s = parse_dt(ws_rec.cell(r, REC_COL_START).value)
                if s and (active_start is None or s > active_start):
                    active_start = s
                    active_row = r

        # Only within the type-2 (event chain) recommendation sequence — 20001+/
        # 30001+ are separate resident-chain (type 3/4) ID ranges, not shared.
        max_rec_id = 0
        for r in range(2, ws_rec.max_row + 1):
            if ws_rec.cell(r, REC_COL_TYPE).value != EVENT_CHAIN_TYPE:
                continue
            idv = ws_rec.cell(r, REC_COL_ID).value
            if isinstance(idv, (int, float)) and idv > max_rec_id:
                max_rec_id = idv

        last_start = parse_dt(ws_rec.cell(active_row, REC_COL_START).value) if active_row else None
        last_end = parse_dt(ws_rec.cell(active_row, REC_COL_END).value) if active_row else None
        window_days = (last_end - last_start).days + 1 if (last_start and last_end) else 7
        last_theme = str(ws_rec.cell(active_row, REC_COL_THEME).value or "") if active_row else ""
        last_formula = str(ws_rec.cell(active_row, REC_COL_FORMULA).value or "") if active_row else ""

        result = {
            "next_rec_id": max_rec_id + 1,
            "next_chain_ids": next_chain_ids,
            "tier_suffixes": TIER_SUFFIXES,
            "active_rec_row_theme": last_theme,
            "active_rec_row_start": last_start.strftime("%Y-%m-%d %H:%M:%S") if last_start else None,
            "active_rec_row_end": last_end.strftime("%Y-%m-%d %H:%M:%S") if last_end else None,
            "window_days": window_days,
            "suggested_next_start": (last_end + timedelta(seconds=1)).strftime("%Y-%m-%d") if last_end else None,
            "latest_triplet_ids": [int(x) for x in latest_ids],
            "latest_stage_lists": latest_stage_lists,
            "next_stage_lists": prev_stage_lists,  # the "other" slot — what the new triplet will clone
            "last_formula_preview": last_formula,
        }
        print(json.dumps(result))

    except Exception as e:
        print(json.dumps({"error": str(e)}))
        sys.exit(1)


if __name__ == "__main__":
    main()
