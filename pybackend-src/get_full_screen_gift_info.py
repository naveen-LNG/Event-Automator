import sys
import os
import json
from datetime import datetime, timedelta
import openpyxl

# recommend_gift column indices (1-based) — verified against real production data
COL_ID = 1
COL_ALIAS = 2
COL_SWITCH = 5
COL_SLOT = 8  # 栏位归属; '16' = 全屏礼包 (full screen gift)
COL_START = 12
COL_END = 13

SLOT_FULL_SCREEN_GIFT = "16"


def parse_dt(v):
    if isinstance(v, datetime):
        return v
    if isinstance(v, str):
        try:
            return datetime.strptime(v.strip(), "%Y-%m-%d %H:%M:%S")
        except ValueError:
            return None
    return None


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    try:
        design_path = sys.argv[1]
    except IndexError:
        print(json.dumps({"error": "Missing design folder path"}))
        sys.exit(1)

    file_path = os.path.join(design_path, "store.xlsx")
    if not os.path.exists(file_path):
        print(json.dumps({"error": f"File not found: {file_path}"}))
        sys.exit(1)

    try:
        wb = openpyxl.load_workbook(file_path, data_only=True)
        ws = wb["recommend_gift"]

        # IDs are banded per slot (184xx = 15, 185xx = 16, 186xx = 17): continue the
        # slot-16 band, skipping any ID already taken anywhere in the sheet.
        all_ids = set()
        max_slot_id = 0
        # Last full-screen-gift row by start date (the clone template + cadence anchor).
        last_row = None
        last_start = None
        for r in range(2, ws.max_row + 1):
            idv = ws.cell(r, COL_ID).value
            try:
                idv_int = int(str(idv).lstrip("/"))
            except (TypeError, ValueError):
                continue
            all_ids.add(idv_int)
            if str(ws.cell(r, COL_SLOT).value or "").strip() == SLOT_FULL_SCREEN_GIFT:
                max_slot_id = max(max_slot_id, idv_int)
                s = parse_dt(ws.cell(r, COL_START).value)
                if s and (last_start is None or s > last_start):
                    last_start = s
                    last_row = r

        if last_row is None:
            print(json.dumps({"error": "No existing slot-16 (Full Screen Gift) rows found in recommend_gift to use as a template."}))
            sys.exit(1)

        next_id = max_slot_id + 1
        while next_id in all_ids:
            next_id += 1

        last_end = parse_dt(ws.cell(last_row, COL_END).value)
        window_days = (last_end - last_start).days + 1 if last_end else 6

        result = {
            "next_id": next_id,
            "last_row_alias": str(ws.cell(last_row, COL_ALIAS).value or ""),
            "last_start": last_start.strftime("%Y-%m-%d %H:%M:%S") if last_start else None,
            "last_end": last_end.strftime("%Y-%m-%d %H:%M:%S") if last_end else None,
            "window_days": window_days,
            # 14-day start-to-start cadence, verified against 11 consecutive real
            # instances (18502-18514) — pure UI suggestion, user can override.
            "suggested_next_start": (last_start + timedelta(days=14)).strftime("%Y-%m-%d") if last_start else None,
        }
        print(json.dumps(result))

    except Exception as e:
        print(json.dumps({"error": str(e)}))
        sys.exit(1)


if __name__ == "__main__":
    main()
