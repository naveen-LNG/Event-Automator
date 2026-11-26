import sys
import os
import json
import copy as _copy
from xlsx_safe import load_workbook_frozen
from datetime import datetime, timedelta
import openpyxl

COL_ID = 1
COL_ALIAS = 2
COL_SWITCH = 5
COL_SLOT = 8
COL_START = 12
COL_END = 13
COL_SYNC_3P_PAYMENT = 33

SLOT_FULL_SCREEN_GIFT = "16"


def copy_cell_format(src_cell, dst_cell):
    dst_cell.number_format = src_cell.number_format
    if src_cell.has_style:
        dst_cell.font = _copy.copy(src_cell.font)
        dst_cell.fill = _copy.copy(src_cell.fill)
        dst_cell.border = _copy.copy(src_cell.border)
        dst_cell.alignment = _copy.copy(src_cell.alignment)
        dst_cell.protection = _copy.copy(src_cell.protection)


def parse_dt(v):
    if isinstance(v, datetime):
        return v
    if isinstance(v, str):
        try:
            return datetime.strptime(v.strip(), "%Y-%m-%d %H:%M:%S")
        except ValueError:
            return None
    return None


def update_full_screen_gift(design_path, start_date_str, duration_days=None):
    """Append a new recommend_gift 'Full Screen Gift' (栏位归属=16) instance.

    Clones the most recent slot-16 row's format/config verbatim (recharge list,
    formula, purchase count, open condition, etc. — verified identical across
    every existing instance), regenerates ID/Alias/Start/End for the requested
    start date. 开关 (switch, col E) and 是否同步三方支付 (col AG) are a
    "current instance" pair of flags: exactly one slot-16 row is TRUE for both
    at a time (verified: 18502-18513 FALSE/FALSE, 18514 TRUE/TRUE — the newest
    row, not a stray mistake). The new row gets both set TRUE, and every other
    slot-16 row that was TRUE gets flipped back to FALSE.
    """
    file_path = os.path.join(design_path, "store.xlsx")
    if not os.path.exists(file_path):
        print(f"Error: store.xlsx not found at {file_path}")
        return False

    try:
        start_date = datetime.strptime(start_date_str.strip(), "%Y-%m-%d")
    except (ValueError, AttributeError):
        print(f"Error: invalid Start Date '{start_date_str}' (expected YYYY-MM-DD).")
        return False

    os.chmod(file_path, 0o666)

    try:
        wb = load_workbook_frozen(file_path)
        if "recommend_gift" not in wb.sheetnames:
            print("Error: recommend_gift sheet not found in store.xlsx")
            return False
        ws = wb["recommend_gift"]

        all_ids = set()
        max_slot_id = 0
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
            print("Error: No existing slot-16 (Full Screen Gift) rows found in recommend_gift to clone from.")
            return False

        last_end = parse_dt(ws.cell(last_row, COL_END).value)
        window_days = (last_end - last_start).days + 1 if last_end else 6
        if duration_days:  # UI override; default = previous instance's length
            window_days = int(duration_days)
            if window_days < 1:
                print(f"Error: Duration must be at least 1 day (got {window_days}).")
                return False

        new_alias = f"全屏礼包{start_date.strftime('%y%m%d')}"

        # Idempotency: a row already covering this exact start date means nothing to do.
        for r in range(2, ws.max_row + 1):
            if str(ws.cell(r, COL_SLOT).value or "").strip() != SLOT_FULL_SCREEN_GIFT:
                continue
            s = parse_dt(ws.cell(r, COL_START).value)
            if s and s.date() == start_date.date():
                print(f"recommend_gift: A Full Screen Gift row for {start_date_str} already exists "
                      f"(ID {ws.cell(r, COL_ID).value}, row {r}). Skipping.")
                return True

        # IDs are banded per slot (185xx = slot 16): continue this band, skip any taken ID.
        new_id = max_slot_id + 1
        while new_id in all_ids:
            new_id += 1
        new_start = start_date
        new_end = new_start + timedelta(days=window_days - 1, hours=23, minutes=59, seconds=59)

        insert_at = last_row + 1
        ws.insert_rows(insert_at)

        for c in range(1, ws.max_column + 1):
            src_c = ws.cell(last_row, c)
            dst_c = ws.cell(insert_at, c)
            if c == COL_ID:
                dst_c.value = new_id
            elif c == COL_ALIAS:
                dst_c.value = new_alias
            elif c == COL_SWITCH:
                dst_c.value = "TRUE"
            elif c == COL_SYNC_3P_PAYMENT:
                dst_c.value = "TRUE"
            elif c == COL_START:
                dst_c.value = new_start.strftime("%Y-%m-%d %H:%M:%S")
            elif c == COL_END:
                dst_c.value = new_end.strftime("%Y-%m-%d %H:%M:%S")
            else:
                dst_c.value = src_c.value
            copy_cell_format(src_c, dst_c)

        # Flip every other TRUE slot-16 row back to FALSE (only the new row stays TRUE).
        demoted = []
        for r in range(2, ws.max_row + 1):
            if r == insert_at:
                continue
            if str(ws.cell(r, COL_SLOT).value or "").strip() != SLOT_FULL_SCREEN_GIFT:
                continue
            changed = False
            if str(ws.cell(r, COL_SWITCH).value or "").strip().upper() == "TRUE":
                ws.cell(r, COL_SWITCH).value = "FALSE"
                changed = True
            if str(ws.cell(r, COL_SYNC_3P_PAYMENT).value or "").strip().upper() == "TRUE":
                ws.cell(r, COL_SYNC_3P_PAYMENT).value = "FALSE"
                changed = True
            if changed:
                demoted.append(r)

        wb.save(file_path)
        print(f"recommend_gift: Added Full Screen Gift ID {new_id} ('{new_alias}') at row {insert_at}, "
              f"{new_start.strftime('%Y-%m-%d')} -> {new_end.strftime('%Y-%m-%d')}, switch/sync=TRUE. "
              f"Demoted previous row(s) to FALSE: {demoted}.")
        print("store.xlsx (recommend_gift) successfully updated.")
        return True

    except Exception as e:
        print(f"Error updating store.xlsx (recommend_gift): {e}")
        import traceback
        traceback.print_exc()
        return False


def main():
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

    try:
        design_path = sys.argv[1]
        payload_path = sys.argv[2]
    except IndexError:
        print("Error: Missing args")
        sys.exit(1)

    with open(payload_path, "r", encoding="utf-8-sig") as f:
        payload = json.load(f)

    success = update_full_screen_gift(
        design_path,
        payload.get("start_date", ""),
        payload.get("duration_days"),
    )
    sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()
