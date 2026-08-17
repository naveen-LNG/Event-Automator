import sys
import os
import json
import openpyxl
from xlsx_safe import load_workbook_frozen
import copy as _copy
from datetime import datetime, timedelta


def copy_cell_format(src_cell, dst_cell):
    """Copy number format + style from src_cell to dst_cell (call after writing
    dst_cell.value). Plain value-only writes leave the destination at
    openpyxl's default "General" format, diverging from source cells
    formatted as e.g. Text."""
    dst_cell.number_format = src_cell.number_format
    if src_cell.has_style:
        dst_cell.font = _copy.copy(src_cell.font)
        dst_cell.fill = _copy.copy(src_cell.fill)
        dst_cell.border = _copy.copy(src_cell.border)
        dst_cell.alignment = _copy.copy(src_cell.alignment)
        dst_cell.protection = _copy.copy(src_cell.protection)


def parse_dt(s):
    return datetime.strptime(str(s).strip(), "%Y-%m-%d %H:%M:%S")


def update_weekly_gift(design_path, event_id, start_time, end_time):
    """Enable store.xlsx's weekly_gift rows for a main event's date range.

    Unlike Main Events, weekly_gift is not cloned per event: it's a single,
    continuous, gapless weekly grid (IDs 10001, 10002, ... one row per
    Mon-Sun week) that's pre-populated far in advance and toggled on/off per
    row via the switch column. A main event doesn't get its own rows — it
    simply "claims" however many of the fixed grid's weekly slots its date
    range overlaps (verified against production data: a 14-day event claims
    2 rows, a 21-day event claims 3, etc. — not a fixed 2-per-event count).

    This function: (1) extends the grid with new gapless weekly rows,
    switch=FALSE, if the event's end date runs past the last pre-populated
    row, then (2) flips switch to TRUE (and un-comments the ID if it was
    "//"-prefixed) on every row whose week overlaps [start_time, end_time].
    Idempotent: rows already TRUE are left alone and reported separately.
    """
    file_path = os.path.join(design_path, "store.xlsx")

    if not os.path.exists(file_path):
        print(f"Error: store.xlsx not found at {file_path}")
        return False

    event_id = str(event_id).strip()
    if not event_id:
        print("Error: Event ID is required.")
        return False
    if not start_time or not end_time:
        print("Error: Start Date and Duration (days) are required.")
        return False

    evt_start = parse_dt(start_time)
    evt_end = parse_dt(end_time)
    if evt_end < evt_start:
        print(f"Error: end time ({evt_end}) is before start time ({evt_start}).")
        return False

    os.chmod(file_path, 0o666)

    try:
        wb = load_workbook_frozen(file_path)
        if "weekly_gift" not in wb.sheetnames:
            print("Error: weekly_gift sheet not found in store.xlsx")
            return False
        ws = wb["weekly_gift"]

        # Anchor on the actual last populated row, not ws.max_row (can be inflated
        # by formatting/dimension metadata past the real data).
        last_row = None
        for r in range(ws.max_row, 1, -1):
            if ws.cell(r, 1).value not in (None, ""):
                last_row = r
                break
        if last_row is None:
            print("Error: weekly_gift sheet has no existing rows to anchor onto.")
            return False

        def parse_id(raw):
            s = str(raw).strip()
            if s.startswith("//"):
                s = s[2:]
            return int(s)

        def row_dates(r):
            s = ws.cell(r, 4).value
            e = ws.cell(r, 5).value
            s = parse_dt(s) if isinstance(s, str) else s
            e = parse_dt(e) if isinstance(e, str) else e
            return s, e

        last_id = parse_id(ws.cell(last_row, 1).value)
        last_start, last_end = row_dates(last_row)
        last_limit_tiers = ws.cell(last_row, 6).value
        src_row_for_format = last_row

        # --- Extend the gapless weekly grid if the event runs past it ---
        new_rows_added = 0
        cur_id = last_id
        cur_row = last_row
        cursor_start = last_end + timedelta(seconds=1)
        while last_end < evt_end:
            row_start = cursor_start
            row_end = row_start + timedelta(days=7, seconds=-1)
            cur_id += 1
            cur_row += 1

            for c in range(1, ws.max_column + 1):
                src_c = ws.cell(src_row_for_format, c)
                if c == 1:
                    val = str(cur_id)  # existing IDs are stored as text (col format is "@")
                elif c == 2:
                    val = None  # Alias: empty on every existing row, never populated
                elif c == 3:
                    val = "FALSE"  # default disabled until a future event claims it
                elif c == 4:
                    val = row_start.strftime("%Y-%m-%d %H:%M:%S")
                elif c == 5:
                    val = row_end.strftime("%Y-%m-%d %H:%M:%S")
                elif c == 6:
                    val = last_limit_tiers
                else:
                    val = None
                dst_c = ws.cell(cur_row, c)
                dst_c.value = val
                copy_cell_format(src_c, dst_c)

            print(f"weekly_gift: Appended row {cur_row} (ID {cur_id}, "
                  f"{row_start.strftime('%Y-%m-%d')} -> {row_end.strftime('%Y-%m-%d')}).")
            new_rows_added += 1
            last_end = row_end
            cursor_start = row_end + timedelta(seconds=1)

        # --- Enable every row whose week overlaps the event's date range, and disable
        # every other row that's currently TRUE (only one event's rows should be live
        # at a time — verified against production data: at most the current + next
        # couple of weeks are TRUE, everything else is FALSE). ---
        enabled_rows = []
        already_enabled = []
        uncommented_rows = []
        disabled_rows = []
        for r in range(2, ws.max_row + 1):
            idv = ws.cell(r, 1).value
            if idv in (None, ""):
                continue
            s, e = row_dates(r)
            if s is None or e is None:
                continue
            in_range = s <= evt_end and e >= evt_start
            switch_c = ws.cell(r, 3)
            cur_switch = str(switch_c.value or "").strip().upper()

            if in_range:
                id_c = ws.cell(r, 1)
                id_str = str(id_c.value).strip()
                if id_str.startswith("//"):
                    id_c.value = id_str[2:]
                    uncommented_rows.append(r)

                if cur_switch == "TRUE":
                    already_enabled.append(r)
                else:
                    switch_c.value = "TRUE"
                    enabled_rows.append(r)
            elif cur_switch == "TRUE":
                switch_c.value = "FALSE"
                disabled_rows.append(r)

        if not enabled_rows and not new_rows_added and not uncommented_rows and not disabled_rows:
            print(f"weekly_gift: No changes needed for event {event_id} "
                  f"({len(already_enabled)} row(s) already enabled).")
        else:
            print(f"weekly_gift: Enabled {len(enabled_rows)} row(s) {enabled_rows}, "
                  f"disabled {len(disabled_rows)} row(s) {disabled_rows} (previous cycle), "
                  f"uncommented {len(uncommented_rows)} row(s) {uncommented_rows}, "
                  f"added {new_rows_added} new row(s) for event {event_id}.")

        wb.save(file_path)
        print(f"store.xlsx (weekly_gift) successfully updated for event {event_id}.")
        return True

    except Exception as e:
        print(f"Error updating store.xlsx (weekly_gift): {e}")
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

    success = update_weekly_gift(
        design_path,
        payload.get("event_id", ""),
        payload.get("start_time", ""),
        payload.get("end_time", ""),
    )
    sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()
