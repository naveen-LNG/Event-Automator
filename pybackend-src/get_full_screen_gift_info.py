import sys
import os
import json
from datetime import datetime, timedelta
import openpyxl

SLOT_FULL_SCREEN_GIFT = "16"  # 栏位归属 value for 全屏礼包

# Header (first line) per column. Positions differ per workspace
# (WorkSpcae3/_NE carry an extra 开关$LNHC2 column), so look them up by name.
HEADERS = {
    "id": "礼包ID", "alias": "Alias", "switch": "开关", "slot": "栏位归属",
    "start": "开始时间", "end": "结束时间", "sync": "是否同步三方支付",
}


def find_cols(ws):
    names = {}
    for c in range(1, ws.max_column + 1):
        h = str(ws.cell(1, c).value or "").split("\n")[0].strip()
        names.setdefault(h, c)
    missing = [h for h in HEADERS.values() if h not in names]
    if missing:
        raise ValueError(f"recommend_gift header(s) not found: {missing}")
    return {k: names[h] for k, h in HEADERS.items()}


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
        col = find_cols(ws)
        COL_ID, COL_ALIAS, COL_SLOT, COL_START, COL_END = (
            col["id"], col["alias"], col["slot"], col["start"], col["end"])

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
