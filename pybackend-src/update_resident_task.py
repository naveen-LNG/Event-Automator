import sys
import os
import re
import json
import openpyxl
import copy as _copy

from xlsx_safe import load_workbook_frozen
from datetime import datetime, timedelta, time as dt_time


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


def get_active_card_collection_info(card_collection_path, evt_start_date):
    """Find the card_collection.xlsx cycle covering evt_start_date (falling back to
    whichever cycle has 活动开关=TRUE) and return (cycle_id, pack_prefix) — e.g. (13, '2607').
    This is card_collection's own independent cycle numbering (~56-day cycles) — it is
    NOT derived from the target event's name/digits, so it must be looked up separately.
    cycle_id (周期ID) is what 任务链's 集卡活动ID column actually stores — verified against
    every historical row in production data, not the pack-name prefix."""
    if not os.path.exists(card_collection_path):
        return None, None
    wb = openpyxl.load_workbook(card_collection_path, data_only=True)
    try:
        if "card_collection" not in wb.sheetnames:
            return None, None
        ws = wb["card_collection"]
        date_match_row = None
        active_flag_row = None
        for r in range(2, ws.max_row + 1):
            cycle_id = ws.cell(r, 1).value
            if cycle_id in (None, ""):
                continue
            start_val = ws.cell(r, 3).value
            end_val = ws.cell(r, 4).value
            start_dt = start_val if isinstance(start_val, datetime) else (
                parse_dt(start_val) if start_val else None)
            end_dt = end_val if isinstance(end_val, datetime) else (
                parse_dt(end_val) if end_val else None)
            if start_dt and end_dt and start_dt.date() <= evt_start_date <= end_dt.date():
                date_match_row = r
            switch = ws.cell(r, 2).value
            if switch is True or str(switch).strip().upper() == "TRUE":
                active_flag_row = r
        target_row = date_match_row or active_flag_row
        if not target_row:
            return None, None
        cycle_id = ws.cell(target_row, 1).value
        seq_val = ws.cell(target_row, 12).value or ws.cell(target_row, 11).value
        pack_prefix = None
        if seq_val:
            m = re.search(r'(\d{4})[一二三四五]星', str(seq_val))
            pack_prefix = m.group(1) if m else None
        return cycle_id, pack_prefix
    finally:
        wb.close()


def update_resident_task(design_path, event_id, event_name, start_time, end_time,
                         preview=False, overrides=None):
    """Append a new streak-challenge cycle to resident_task.xlsx.

    Unlike the main event pipeline, this doesn't clone a source event's block:
    任务链 rows are a pure date-driven weekly chain, so this appends new rows
    continuing that chain (row count derived from Start Date + Duration), tagged
    with the given numeric Event ID (活动ID) and the matching card_collection.xlsx
    cycle's 周期ID (集卡活动ID).

    任务列表's task IDs are evergreen — the same 16xx/17xx rows get reused forever;
    only their reward content changes. So this does NOT mint new task rows: it
    updates col I/J (extra-reward token) on every existing 任务列表 row that
    currently carries an event token, swapping it to event_name, and col K
    (extra card-pack reward) on every row that currently carries a card-pack
    reference, swapping it to whichever card_collection.xlsx cycle covers this
    event's dates (a separate, independently-versioned lookup — not derived
    from the event name). New 任务链 rows then reference the same task-ID list
    the chain already uses, unchanged.

    preview=True: compute and print the plan as one "PREVIEW_JSON:" line, write nothing.
    overrides: user edits to that plan {rows, token_name, pack_prefix} applied on commit."""
    overrides = overrides or {}
    file_path = os.path.join(design_path, "resident_task.xlsx")
    card_collection_path = os.path.join(design_path, "card_collection.xlsx")

    if not os.path.exists(file_path):
        print(f"Error: resident_task.xlsx not found at {file_path}")
        return False

    event_id = str(event_id).strip()
    if not event_id:
        print("Error: Event ID is required.")
        return False
    if not start_time or not end_time:
        print("Error: Start Date and Duration (days) are required.")
        return False
    if not event_name:
        print("Error: Event Name is required.")
        return False

    evt_start = parse_dt(start_time)
    evt_end = parse_dt(end_time)
    duration_days = (evt_end.date() - evt_start.date()).days + 1
    if duration_days < 1:
        print(f"Error: computed duration ({duration_days} days) is invalid — check Start Date/Duration.")
        return False
    num_new_rows = (duration_days + 6) // 7  # ceil(duration_days / 7)

    tgt_upper = event_name.upper()
    tgt_pascal = event_name[0].upper() + event_name[1:]

    # The extra-reward currency's Chinese suffix word is tied to event TYPE, not carried
    # over from whatever event the template last belonged to: verified against every
    # {Type}{digits}(能量|代币) alias in item.xlsx — every "能量" ("energy") entry is a
    # Merge-type event, every other type (Sort, Sim, Blocks, Hunt, Drink, Dessert, ...)
    # uses "代币" ("token"), with zero exceptions. A prefix-only swap would otherwise
    # silently reproduce the wrong suffix (e.g. cloning a Merge template onto a Sort
    # event would keep "能量" instead of switching to "代币").
    tgt_family_match = re.match(r'^([A-Za-z]+)', event_name)
    tgt_family = tgt_family_match.group(1) if tgt_family_match else event_name
    tgt_suffix = "能量" if tgt_family.lower() == "merge" else "代币"
    tok = str(overrides.get("token_name") or "").strip()
    if tok:
        m_tok = re.match(r'^(.*?)(能量|代币)?$', tok)
        tgt_pascal, tgt_suffix = m_tok.group(1), (m_tok.group(2) or tgt_suffix)
        tgt_upper = tgt_pascal.upper()

    os.chmod(file_path, 0o666)

    try:
        # 任务列表's 任务ID column is formula-driven (e.g. "=A541+300") for ~30% of its
        # rows, and a downstream build step reads those cached values rather than
        # recalculating — so the formulas must be baked to literals before saving.
        # See xlsx_safe.load_workbook_frozen for why.
        wb = load_workbook_frozen(file_path)

        ws_chain = wb["任务链"]
        ws_list = wb["任务列表"]

        # Idempotency: this event's activity ID already present in 任务链?
        existing_rows = [r for r in range(2, ws_chain.max_row + 1)
                          if str(ws_chain.cell(r, 9).value or "").strip() == event_id]
        if existing_rows:
            print(f"resident_task [任务链]: activity ID {event_id} already present at rows {existing_rows}. Skipping.")
            return True

        # Anchor on the actual last populated row, not ws.max_row — openpyxl's max_row can
        # be inflated by formatting/dimension metadata on rows past the real data, and
        # inserting at max_row+1 there leaves a huge blank gap instead of appending cleanly.
        last_row = None
        for r in range(ws_chain.max_row, 1, -1):
            if ws_chain.cell(r, 1).value not in (None, ""):
                last_row = r
                break
        if last_row is None:
            print("Error: 任务链 sheet has no existing rows to anchor the new cycle onto.")
            return False

        last_id = ws_chain.cell(last_row, 1).value
        last_type = ws_chain.cell(last_row, 2).value
        last_task_list_raw = ws_chain.cell(last_row, 5).value
        last_condition = ws_chain.cell(last_row, 7).value
        last_pack_id = ws_chain.cell(last_row, 8).value
        last_end = ws_chain.cell(last_row, 4).value
        if isinstance(last_end, str):
            last_end = parse_dt(last_end)

        try:
            last_id_num = int(last_id)
        except (TypeError, ValueError):
            last_id_num = None
        if last_id_num is None:
            print(f"Error: could not parse last 任务链 row's ID ({last_id!r}) as a number.")
            return False

        def norm_type(t):
            try:
                return int(t)
            except (TypeError, ValueError):
                return 1
        next_type = 2 if norm_type(last_type) == 1 else 1

        if not str(last_task_list_raw or "").strip():
            print(f"Error: last 任务链 row's 固定任务列表 is empty — nothing to reuse.")
            return False

        cycle_id, card_pack_prefix = get_active_card_collection_info(card_collection_path, evt_start.date())
        if str(overrides.get("pack_prefix") or "").strip():
            card_pack_prefix = str(overrides["pack_prefix"]).strip()
        if overrides.get("cycle_id") not in (None, ""):
            cycle_id = overrides["cycle_id"]
        if not cycle_id:
            print("Warning: could not determine the active card_collection.xlsx cycle for this event's "
                  "dates — 任务链 col J (集卡活动ID) will be left blank; set it manually.")
        if not card_pack_prefix:
            print("Warning: could not determine the active card_collection.xlsx card-pack prefix — "
                  "任务列表 col K (额外卡包奖励内容) will be left as-is; check it manually.")

        # --- Update col I/J/K in place on every existing 任务列表 row that carries them ---
        # (evergreen rows, e.g. the 16xx/17xx blocks — never mint new task IDs)
        def swap_token_cell(cell, new_prefix):
            """Detect this cell's OWN prefix+suffix and swap independently of any other
            cell — verified in real data that col I and col J on the same row can carry
            different suffixes (e.g. I='Merge2608能量', J='MERGE2608代币' for the same
            reward), so J must never be assumed to mirror I's suffix."""
            val = cell.value
            if not isinstance(val, str):
                return False
            m = re.match(r'^([A-Za-z]+\d{3,4})(能量|代币)', val)
            if m:
                src_token, src_suffix = m.group(1), m.group(2)
                cell.value = val.replace(src_token + src_suffix, new_prefix + tgt_suffix, 1)
                return True
            m2 = re.match(r'^([A-Za-z]+\d{3,4})', val)
            if m2:
                cell.value = val.replace(m2.group(1), new_prefix)
                print(f"Warning: 任务列表 row {cell.row} col {cell.column_letter} ({val!r}) didn't match "
                      f"the expected 能量/代币 suffix pattern — swapped prefix only; verify the suffix manually.")
                return True
            return False

        token_updates = 0
        pack_updates = 0
        for r in range(2, ws_list.max_row + 1):
            i_changed = swap_token_cell(ws_list.cell(r, 9), tgt_pascal)
            j_changed = swap_token_cell(ws_list.cell(r, 10), tgt_upper)
            if i_changed or j_changed:
                token_updates += 1

            if card_pack_prefix:
                k_c = ws_list.cell(r, 11)
                k_val = k_c.value
                if isinstance(k_val, str) and k_val.strip():
                    m = re.search(r'(\d{4})([一二三四五]星)', k_val)
                    if m and m.group(1) != card_pack_prefix:
                        k_c.value = k_val.replace(m.group(1), card_pack_prefix)
                        pack_updates += 1

        print(f"resident_task [任务列表]: Updated extra-reward token on {token_updates} row(s) -> "
              f"{tgt_pascal + tgt_suffix!r}, card-pack prefix on {pack_updates} row(s) -> {card_pack_prefix!r}.")

        # --- Append new 任务链 rows spanning duration_days, alternating type, reusing the
        # chain's existing task list unchanged (task IDs are evergreen, see above) ---
        # 任务链 is a gapless back-to-back weekly chain (verified against production data), so new
        # rows always continue from the last existing row's end date, not from the given Start
        # Date — just warn if those two disagree instead of silently creating a gap or overlap.
        cursor_date = last_end.date() + timedelta(days=1)
        if cursor_date != evt_start.date():
            print(f"Warning: Start Date ({evt_start.date()}) does not match where the 任务链 chain "
                  f"actually continues from ({cursor_date}) — using {cursor_date} to keep the chain gapless.")
        remaining_days = duration_days
        cur_type = next_type
        planned = []
        for i in range(num_new_rows):
            days_this_row = min(7, remaining_days)
            row_end_date = cursor_date + timedelta(days=days_this_row - 1)
            planned.append({
                "id": last_id_num + 1 + i, "type": cur_type,
                "start": datetime.combine(cursor_date, dt_time(6, 0, 0)).strftime("%Y-%m-%d %H:%M:%S"),
                "end": datetime.combine(row_end_date, dt_time(23, 59, 59)).strftime("%Y-%m-%d %H:%M:%S"),
                "tasks": last_task_list_raw,
                "activity_id": int(event_id) if event_id.isdigit() else event_id,
                "card_id": cycle_id,
            })
            cursor_date = row_end_date + timedelta(days=1)
            remaining_days -= days_this_row
            cur_type = 2 if cur_type == 1 else 1
        if overrides.get("rows"):
            planned = overrides["rows"]

        if preview:
            print("PREVIEW_JSON:" + json.dumps({
                "rows": planned, "token_name": tgt_pascal + tgt_suffix,
                "pack_prefix": card_pack_prefix, "cycle_id": cycle_id,
                "token_rows": token_updates, "pack_rows": pack_updates,
                "last_row": last_row, "last_end": last_end.strftime("%Y-%m-%d"),
            }, ensure_ascii=False))
            return True

        chain_insert_at = last_row + 1
        for i, pr in enumerate(planned):
            dst_r = chain_insert_at + i
            src_r = last_row  # copy formatting/constants from the last real row
            for c in range(1, ws_chain.max_column + 1):
                src_c = ws_chain.cell(src_r, c)
                val = src_c.value
                if c == 1: val = pr["id"]
                elif c == 2: val = pr["type"]
                elif c == 3: val = pr["start"]
                elif c == 4: val = pr["end"]
                elif c == 5: val = pr["tasks"]
                elif c == 6: val = None
                elif c == 7: val = last_condition
                elif c == 8: val = last_pack_id
                elif c == 9: val = pr["activity_id"]
                elif c == 10: val = pr["card_id"]
                dst_c = ws_chain.cell(dst_r, c)
                dst_c.value = val
                copy_cell_format(src_c, dst_c)
            print(f"resident_task [任务链]: Added row {dst_r} (ID {pr['id']}, type {pr['type']}, "
                  f"{pr['start']} -> {pr['end']}, activity ID {pr['activity_id']}, 集卡活动ID {pr['card_id']}).")

        wb.save(file_path)
        print(f"resident_task.xlsx successfully updated for event ID {event_id} ({len(planned)} row(s), "
              f"{duration_days} days).")
        return True

    except Exception as e:
        print(f"Error updating resident_task.xlsx: {e}")
        import traceback
        traceback.print_exc()
        return False


def main():
    # Log messages include Chinese sheet/column names; when stdout isn't attached to a
    # real UTF-8 console (e.g. piped from Electron's spawn on a non-UTF-8 Windows locale)
    # the default encoding can be cp1252 and crash on write. Force UTF-8 defensively.
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

    success = update_resident_task(
        design_path,
        payload.get("event_id", ""),
        payload.get("event_name", ""),
        payload.get("start_time", ""),
        payload.get("end_time", ""),
        preview=bool(payload.get("preview")),
        overrides=payload.get("overrides"),
    )
    sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()
