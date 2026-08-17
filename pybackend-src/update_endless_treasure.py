import sys
import os
import re
import json
import copy as _copy
from xlsx_safe import load_workbook_frozen
from datetime import datetime, timedelta
import openpyxl

# 折扣链数据 (chain data) columns
CHAIN_COL_ID = 1
CHAIN_COL_DESC = 2
CHAIN_COL_TYPE = 3
CHAIN_COL_STAGE_LIST = 5
EVENT_CHAIN_TYPE = 2

# 折扣链推荐 (recommendation/scheduler) columns
REC_COL_ID = 1
REC_COL_TYPE = 3
REC_COL_START = 4
REC_COL_END = 5
REC_COL_END_ACTIVATE = 6
REC_COL_FORMULA = 7
REC_COL_SWITCH = 9
REC_COL_THEME = 10

# 折扣链常量 (constants) — ENDLESS_TREASURE_VALID_LIST row
CONST_COL_KEY = 1
CONST_COL_VALUE = 2
VALID_LIST_KEY = "ENDLESS_TREASURE_VALID_LIST"

TIER_SUFFIXES = ["1.99折扣链", "4.99折扣链", "9.99折扣链"]


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


def get_active_card_collection_info(card_collection_path, evt_start_date):
    """Find the card_collection.xlsx cycle covering evt_start_date (falling back to
    whichever cycle has 活动开关=TRUE) and return its card-pack prefix (e.g. '2607').
    Mirrors update_resident_task.py's get_active_card_collection_info exactly —
    card_collection.xlsx runs its own independent ~56-day cycle numbering, unrelated
    to the endless-treasure event window being created."""
    if not os.path.exists(card_collection_path):
        return None
    wb = openpyxl.load_workbook(card_collection_path, data_only=True)
    try:
        if "card_collection" not in wb.sheetnames:
            return None
        ws = wb["card_collection"]
        date_match_row = None
        active_flag_row = None
        for r in range(2, ws.max_row + 1):
            cycle_id = ws.cell(r, 1).value
            if cycle_id in (None, ""):
                continue
            start_val = ws.cell(r, 3).value
            end_val = ws.cell(r, 4).value
            start_dt = start_val if isinstance(start_val, datetime) else (parse_dt(start_val) if start_val else None)
            end_dt = end_val if isinstance(end_val, datetime) else (parse_dt(end_val) if end_val else None)
            if start_dt and end_dt and start_dt.date() <= evt_start_date <= end_dt.date():
                date_match_row = r
            switch = ws.cell(r, 2).value
            if switch is True or str(switch).strip().upper() == "TRUE":
                active_flag_row = r
        target_row = date_match_row or active_flag_row
        if not target_row:
            return None
        seq_val = ws.cell(target_row, 12).value or ws.cell(target_row, 11).value
        if not seq_val:
            return None
        m = re.search(r'(\d{4})[一二三四五]星', str(seq_val))
        return m.group(1) if m else None
    finally:
        wb.close()


def refresh_card_pack_prefix(ws_stage, stage_lists, card_pack_prefix):
    """Update every reward-content cell (across the full row, not just one known
    column — 折扣链阶段数据 repeats the same reward in multiple A/B-test-variant
    columns, e.g. 奖励内容/奖励内容$E1-E3/奖励内容$E4-E6 and their 显示内容1/2/3/4
    counterparts) on the stage rows the newly-activated triplet references, so a
    stale card-pack code (e.g. last cycle's '2408') doesn't go live again."""
    if not card_pack_prefix:
        return 0
    id_to_row = {}
    for r in range(2, ws_stage.max_row + 1):
        v = ws_stage.cell(r, 1).value
        if v is not None:
            id_to_row[v] = r
    stage_ids = []
    for stage_list in stage_lists:
        for tok in str(stage_list or "").split(","):
            tok = tok.strip()
            if tok:
                try:
                    stage_ids.append(int(tok))
                except ValueError:
                    pass
    updated = 0
    for sid in stage_ids:
        r = id_to_row.get(sid)
        if not r:
            continue
        for c in range(2, ws_stage.max_column + 1):
            cell = ws_stage.cell(r, c)
            val = cell.value
            if not isinstance(val, str):
                continue
            m = re.search(r'(\d{4})([一二三四五]星)', val)
            if m and m.group(1) != card_pack_prefix:
                cell.value = val.replace(m.group(1) + m.group(2), card_pack_prefix + m.group(2))
                updated += 1
    return updated


def find_last_two_event_triplets(ws_chain):
    event_rows = []
    for r in range(2, ws_chain.max_row + 1):
        if ws_chain.cell(r, CHAIN_COL_TYPE).value == EVENT_CHAIN_TYPE:
            idv = ws_chain.cell(r, CHAIN_COL_ID).value
            if isinstance(idv, (int, float)):
                event_rows.append(r)
    if len(event_rows) < 6:
        return None, None
    return event_rows[-3:], event_rows[-6:-3]


def update_endless_treasure(design_path, start_date_str, event_name):
    """Append a new endless-treasure event chain triplet + scheduling window.

    折扣链数据 event-chain (类型=2) triplets alternate between two fixed stage-ID
    templates every window (verified against 5 SVN snapshots spanning Aug 2025 -
    Jul 2026: the two templates' reward content has never diverged), purely so a
    player mid-purchase on the outgoing chain doesn't have their stage progression
    IDs reassigned when the next chain activates. So the new triplet clones the
    OTHER slot's most recent triplet (2 triplets back) column-for-column,
    tier-for-tier (1.99/4.99/9.99) — only ID and 说明 (description) change.

    折扣链推荐 gets one new scheduling row cloned from the currently-active row
    (formula regenerated by substituting the 3 new chain IDs for the 3 old ones,
    preserving whatever exact threshold structure the live formula already has),
    with 开关=TRUE; the previously-active row's 开关 is flipped to FALSE.

    折扣链常量's ENDLESS_TREASURE_VALID_LIST has the outgoing triplet's IDs
    removed and the new triplet's IDs appended.
    """
    file_path = os.path.join(design_path, "endless_treasure.xlsx")
    if not os.path.exists(file_path):
        print(f"Error: endless_treasure.xlsx not found at {file_path}")
        return False

    try:
        start_date = datetime.strptime(start_date_str.strip(), "%Y-%m-%d")
    except (ValueError, AttributeError):
        print(f"Error: invalid Start Date '{start_date_str}' (expected YYYY-MM-DD).")
        return False

    event_name = (event_name or "").strip()
    if not event_name:
        print("Error: Event Name is required.")
        return False

    os.chmod(file_path, 0o666)

    try:
        wb = load_workbook_frozen(file_path)
        ws_chain = wb["折扣链数据"]
        ws_rec = wb["折扣链推荐"]
        ws_const = wb["折扣链常量"]

        # Idempotency: a recommendation row already covering this exact start date.
        for r in range(2, ws_rec.max_row + 1):
            s = parse_dt(ws_rec.cell(r, REC_COL_START).value)
            if s and s.date() == start_date.date():
                print(f"折扣链推荐: A row for {start_date_str} already exists "
                      f"(ID {ws_rec.cell(r, REC_COL_ID).value}, row {r}). Skipping.")
                return True

        latest_triplet, prev_triplet = find_last_two_event_triplets(ws_chain)
        if not latest_triplet:
            print("Error: Fewer than 2 event-chain (类型=2) triplets found in 折扣链数据 — cannot determine slot alternation.")
            return False

        for rows, label in [(latest_triplet, "latest"), (prev_triplet, "previous")]:
            descs = [str(ws_chain.cell(r, CHAIN_COL_DESC).value or "") for r in rows]
            if not all(descs[i].endswith(TIER_SUFFIXES[i]) for i in range(3)):
                print(f"Error: {label} triplet at rows {rows} does not match the expected "
                      f"1.99/4.99/9.99折扣链 tier order ({descs}). Refusing to guess — check the sheet manually.")
                return False

        latest_ids = [int(ws_chain.cell(r, CHAIN_COL_ID).value) for r in latest_triplet]
        next_chain_ids = [latest_ids[2] + 1, latest_ids[2] + 2, latest_ids[2] + 3]

        # Clone the OTHER slot's most recent triplet (2 back) — that's the
        # template for the new one, since slots alternate every window.
        new_rows = []
        insert_at = ws_chain.max_row + 1
        for i, src_r in enumerate(prev_triplet):
            dst_r = insert_at + i
            for c in range(1, ws_chain.max_column + 1):
                src_c = ws_chain.cell(src_r, c)
                dst_c = ws_chain.cell(dst_r, c)
                if c == CHAIN_COL_ID:
                    dst_c.value = next_chain_ids[i]
                elif c == CHAIN_COL_DESC:
                    dst_c.value = f"{event_name}{TIER_SUFFIXES[i]}"
                else:
                    dst_c.value = src_c.value
                copy_cell_format(src_c, dst_c)
            new_rows.append(dst_r)

        # --- 折扣链阶段数据: refresh the newly-activated slot's card-pack reward
        # to the currently-active card_collection.xlsx cycle (mirrors the same
        # lookup update_resident_task.py uses for Streak Challenge) ---
        new_stage_lists = [ws_chain.cell(r, CHAIN_COL_STAGE_LIST).value for r in new_rows]
        card_collection_path = os.path.join(design_path, "card_collection.xlsx")
        card_pack_prefix = get_active_card_collection_info(card_collection_path, start_date.date())
        if not card_pack_prefix:
            print("Warning: could not determine the active card_collection.xlsx card-pack cycle — "
                  "折扣链阶段数据's card-pack reward references will be left as-is; check them manually.")
        else:
            ws_stage = wb["折扣链阶段数据"]
            pack_updates = refresh_card_pack_prefix(ws_stage, new_stage_lists, card_pack_prefix)
            print(f"折扣链阶段数据: Refreshed card-pack prefix on {pack_updates} cell(s) -> {card_pack_prefix!r}.")

        # --- 折扣链推荐: clone the active row, regenerate formula + dates ---
        active_row = None
        active_start = None
        for r in range(2, ws_rec.max_row + 1):
            s = parse_dt(ws_rec.cell(r, REC_COL_START).value)
            if s and str(ws_rec.cell(r, REC_COL_SWITCH).value or "").strip().upper() == "TRUE":
                if active_start is None or s > active_start:
                    active_start = s
                    active_row = r
        if active_row is None:
            print("Error: No 折扣链推荐 row currently has 开关=TRUE — cannot determine the active window to clone from.")
            return False

        old_start = parse_dt(ws_rec.cell(active_row, REC_COL_START).value)
        old_end = parse_dt(ws_rec.cell(active_row, REC_COL_END).value)
        window_days = (old_end - old_start).days + 1 if (old_start and old_end) else 7
        new_start = start_date
        new_end = new_start + timedelta(days=window_days - 1, hours=23, minutes=59, seconds=59)

        old_formula = str(ws_rec.cell(active_row, REC_COL_FORMULA).value or "")
        new_formula = old_formula
        for old_id, new_id in zip(latest_ids, next_chain_ids):
            new_formula = re.sub(rf'\b{old_id}\b', str(new_id), new_formula)

        # Only within the type-2 (event chain) recommendation sequence — 20001+/
        # 30001+ are separate resident-chain (type 3/4) ID ranges, not shared.
        max_rec_id = 0
        for r in range(2, ws_rec.max_row + 1):
            if ws_rec.cell(r, REC_COL_TYPE).value != EVENT_CHAIN_TYPE:
                continue
            idv = ws_rec.cell(r, REC_COL_ID).value
            if isinstance(idv, (int, float)) and idv > max_rec_id:
                max_rec_id = idv
        new_rec_id = max_rec_id + 1

        theme_tag = f"{start_date.strftime('%Y%m')}{re.sub(r'[^a-zA-Z0-9]', '', event_name).lower()}"

        # Insert right after the last type-2 row, not at the sheet's absolute end —
        # rows 63+ are unrelated type-3/4 resident chains (20001+/30001+), and a
        # blind append would land the new row past them, breaking the type-2
        # sequence's contiguity.
        last_type2_row = active_row
        for r in range(2, ws_rec.max_row + 1):
            if ws_rec.cell(r, REC_COL_TYPE).value == EVENT_CHAIN_TYPE and r > last_type2_row:
                last_type2_row = r
        rec_insert_at = last_type2_row + 1
        ws_rec.insert_rows(rec_insert_at)
        for c in range(1, ws_rec.max_column + 1):
            src_c = ws_rec.cell(active_row, c)
            dst_c = ws_rec.cell(rec_insert_at, c)
            if c == REC_COL_ID:
                dst_c.value = new_rec_id
            elif c == REC_COL_START:
                dst_c.value = new_start.strftime("%Y-%m-%d %H:%M:%S")
            elif c == REC_COL_END:
                dst_c.value = new_end.strftime("%Y-%m-%d %H:%M:%S")
            elif c == REC_COL_END_ACTIVATE:
                dst_c.value = new_end.strftime("%Y-%m-%d %H:%M:%S")
            elif c == REC_COL_FORMULA:
                dst_c.value = new_formula
            elif c == REC_COL_SWITCH:
                dst_c.value = "TRUE"
            elif c == REC_COL_THEME:
                dst_c.value = theme_tag
            else:
                dst_c.value = src_c.value
            copy_cell_format(src_c, dst_c)

        ws_rec.cell(active_row, REC_COL_SWITCH).value = "FALSE"

        # --- 折扣链常量: swap the outgoing triplet's IDs for the new ones ---
        for r in range(2, ws_const.max_row + 1):
            if str(ws_const.cell(r, CONST_COL_KEY).value or "").strip() == VALID_LIST_KEY:
                raw = str(ws_const.cell(r, CONST_COL_VALUE).value or "")
                ids = [x.strip() for x in raw.split(",") if x.strip()]
                old_id_strs = {str(x) for x in latest_ids}
                ids = [x for x in ids if x not in old_id_strs]
                ids.extend(str(x) for x in next_chain_ids)
                ws_const.cell(r, CONST_COL_VALUE).value = ",".join(ids)
                break

        wb.save(file_path)
        print(f"折扣链数据: Added triplet {next_chain_ids} ('{event_name}') at rows {new_rows}, cloned from slot template at rows {prev_triplet}.")
        print(f"折扣链推荐: Added ID {new_rec_id} at row {rec_insert_at}, {new_start.strftime('%Y-%m-%d')} -> {new_end.strftime('%Y-%m-%d')}, "
              f"theme='{theme_tag}', switch=TRUE. Demoted row {active_row} to switch=FALSE.")
        print("折扣链常量: ENDLESS_TREASURE_VALID_LIST updated.")
        print("endless_treasure.xlsx successfully updated.")
        return True

    except Exception as e:
        print(f"Error updating endless_treasure.xlsx: {e}")
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

    success = update_endless_treasure(
        design_path,
        payload.get("start_date", ""),
        payload.get("event_name", ""),
    )
    sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()
