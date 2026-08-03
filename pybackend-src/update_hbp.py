import sys
import os
import openpyxl
import json
import re
import copy as _copy

def copy_cell_format(src_cell, dst_cell):
    """Copy number format + style from src_cell to dst_cell (call after
    writing dst_cell.value). Plain value-only writes leave the destination at
    openpyxl's default "General" format, diverging from source cells
    formatted as e.g. Text."""
    dst_cell.number_format = src_cell.number_format
    if src_cell.has_style:
        dst_cell.font = _copy.copy(src_cell.font)
        dst_cell.fill = _copy.copy(src_cell.fill)
        dst_cell.border = _copy.copy(src_cell.border)
        dst_cell.alignment = _copy.copy(src_cell.alignment)
        dst_cell.protection = _copy.copy(src_cell.protection)

def get_last_data_row(ws, id_col=1):
    for r in range(ws.max_row, 0, -1):
        val = ws.cell(row=r, column=id_col).value
        if val is not None and str(val).strip() != '':
            return r
    return 1

def find_last_matching_block(ws, pattern, alias_col=2):
    """Scan from the bottom of the sheet for the last holiday cycle's rows:
    all rows whose alias (alias_col) matches `pattern` (e.g.
    r'^\\d{6}节日BP通行证') AND share the same leading 6-digit holiday id as the
    very last matching row. Consecutive cycles sit back-to-back with no
    separator row, so matching on `pattern` alone would sweep in every past
    cycle too — the id-prefix check is what actually bounds a single cycle.
    Returns row numbers top to bottom, or [] if no match. Used instead of
    hardcoded row numbers, which go stale every time a new cycle is appended."""
    regex = re.compile(pattern)
    last_match_row = None
    last_prefix = None
    for r in range(ws.max_row, 1, -1):
        alias = str(ws.cell(row=r, column=alias_col).value or "")
        m = regex.search(alias)
        if m:
            last_match_row = r
            last_prefix = alias[:6]
            break
    if last_match_row is None:
        return []
    rows = [last_match_row]
    r = last_match_row - 1
    while r >= 2:
        alias = str(ws.cell(row=r, column=alias_col).value or "")
        if regex.search(alias) and alias[:6] == last_prefix:
            rows.append(r)
            r -= 1
        else:
            break
    rows.reverse()
    return rows

def main():
    sys.stdout.reconfigure(encoding='utf-8')
    try:
        design_path = sys.argv[1]
        payload_path = sys.argv[2]
    except IndexError:
        print("Error: Missing args")
        sys.exit(1)

    with open(payload_path, 'r', encoding='utf-8-sig') as f:
        payload = json.load(f)

    holiday_id = payload['holiday_id'] # e.g. "260601"
    start_time = payload['start_time']
    end_time = payload['end_time']

    bp_path = os.path.join(design_path, "festivals_bp.xlsx")
    item_path = os.path.join(design_path, "item.xlsx")
    store_path = os.path.join(design_path, "store.xlsx")

    try:
        wb_bp = openpyxl.load_workbook(bp_path)
        
        # --- 1. festivals_bp sheet ---
        ws_bp = wb_bp['festivals_bp']
        last_bp_r = get_last_data_row(ws_bp)
        
        # Toggle old switches
        for r in range(2, last_bp_r + 1):
            ws_bp.cell(row=r, column=2).value = False

        # Title/token: verified against 12 real SVN revisions (cycles 27-32) —
        # the active cycle's title strictly ALTERNATES every single transition
        # (bare <-> "_2"), and every other (non-active) row is mass-reset in a
        # block to whatever the just-expired cycle itself used while active.
        # It is not a fixed "active=bare" convention — capture the old value
        # BEFORE overwriting anything, toggle it for the new row, and reset
        # every prior row to the old value.
        old_title = str(ws_bp.cell(row=last_bp_r, column=5).value or "HOLIDAY_BP_TITLE")
        old_token = str(ws_bp.cell(row=last_bp_r, column=6).value or "HOLIDAY_BP_TOKEN_NAME")
        if "_2" in old_title:
            new_title, new_token = "HOLIDAY_BP_TITLE", "HOLIDAY_BP_TOKEN_NAME"
        else:
            new_title, new_token = "HOLIDAY_BP_TITLE_2", "HOLIDAY_BP_TOKEN_NAME_2"

        for r in range(2, last_bp_r + 1):
            ws_bp.cell(row=r, column=5).value = old_title
            ws_bp.cell(row=r, column=6).value = old_token

        new_bp_r = last_bp_r + 1
        cycle_id = int(ws_bp.cell(row=last_bp_r, column=1).value or 0) + 1
        recommend_id = int(ws_bp.cell(row=last_bp_r, column=9).value or 18400) + 1

        # Duplicate last row values initially
        for c in range(1, ws_bp.max_column + 1):
            src_c = ws_bp.cell(row=last_bp_r, column=c)
            dst_c = ws_bp.cell(row=new_bp_r, column=c)
            dst_c.value = src_c.value
            copy_cell_format(src_c, dst_c)

        ws_bp.cell(row=new_bp_r, column=1).value = cycle_id
        ws_bp.cell(row=new_bp_r, column=2).value = True
        ws_bp.cell(row=new_bp_r, column=3).value = start_time
        ws_bp.cell(row=new_bp_r, column=4).value = end_time
        ws_bp.cell(row=new_bp_r, column=5).value = new_title
        ws_bp.cell(row=new_bp_r, column=6).value = new_token
        ws_bp.cell(row=new_bp_r, column=9).value = recommend_id

        # --- 2. 扩展表 (Extended Table) ---
        ws_ext = wb_bp['扩展表']
        last_ext_r = get_last_data_row(ws_ext)
        new_ext_start = last_ext_r + 1
        
        # Determine suffixes and last item id from previous group of 4
        prev_group_start = max(2, last_ext_r - 3)
        prev_item_id = int(ws_ext.cell(row=last_ext_r, column=5).value or 3310)
        new_item_id = prev_item_id + 1
        
        for i in range(4):
            src_r = prev_group_start + i
            tgt_r = new_ext_start + i
            for c in range(1, ws_ext.max_column + 1):
                src_c = ws_ext.cell(row=src_r, column=c)
                dst_c = ws_ext.cell(row=tgt_r, column=c)
                dst_c.value = src_c.value
                copy_cell_format(src_c, dst_c)
            
            # Alias logic: detect prefix
            old_alias = str(ws_ext.cell(row=src_r, column=2).value)
            match = re.match(r'(\d{6})(.*)', old_alias)
            suffix = match.group(2) if match else "节日BP" + str(i+1)
            ws_ext.cell(row=tgt_r, column=2).value = f"{holiday_id}{suffix}"
            
            # Increment numeric columns: 1 (Extension ID), 4 (Paid Products)
            ws_ext.cell(row=tgt_r, column=1).value = int(ws_ext.cell(row=src_r, column=1).value or 0) + 4 # Wait, increment by 4 since we add a group of 4? No, user says 'remaining columns incremented by 1'. 
            # Actually, looking at current data: 38074, 38075, 38076, 38077. So they increment by 4 from previous group's corresponding row?
            # User said: 'Remaining columns will be incremented by 1'. If Row 70 (38074) was duplicated to 74? 38074 + 4 = 38078. 
            # Let's use: prev_row_val + 4.
            ws_ext.cell(row=tgt_r, column=1).value = int(ws_ext.cell(row=src_r, column=1).value) + 4
            ws_ext.cell(row=tgt_r, column=4).value = int(ws_ext.cell(row=src_r, column=4).value) + 4
            ws_ext.cell(row=tgt_r, column=5).value = new_item_id

        # --- 3. chest_plan ---
        ws_cp = wb_bp['chest_plan']
        last_cp_r = get_last_data_row(ws_cp)
        new_cp_start = last_cp_r + 1
        
        prev_cp_group_start = max(2, last_cp_r - 3)
        old_holiday_prefix = ""
        for i in range(4):
            src_r = prev_cp_group_start + i
            tgt_r = new_cp_start + i
            for c in range(1, ws_cp.max_column + 1):
                src_c = ws_cp.cell(row=src_r, column=c)
                dst_c = ws_cp.cell(row=tgt_r, column=c)
                dst_c.value = src_c.value
                copy_cell_format(src_c, dst_c)
            
            ws_cp.cell(row=tgt_r, column=1).value = int(ws_cp.cell(row=src_r, column=1).value) + 4
            
            old_alias = str(ws_cp.cell(row=src_r, column=2).value)
            match = re.match(r'(\d{6})(.*)', old_alias)
            if match:
                old_holiday_prefix = match.group(1)
                suffix = match.group(2)
                ws_cp.cell(row=tgt_r, column=2).value = f"{holiday_id}{suffix}"
            
            # 积分奖励列表 string replacement
            rewards_str = str(ws_cp.cell(row=tgt_r, column=7).value)
            ws_cp.cell(row=tgt_r, column=7).value = rewards_str.replace(old_holiday_prefix, holiday_id)

        # --- 4. chest_conditions ---
        ws_cond = wb_bp['chest_conditions']
        # Locate last 4 groups (based on old_holiday_prefix)
        template_rows = []
        for r in range(4, ws_cond.max_row + 1):
            val = str(ws_cond.cell(row=r, column=2).value)
            if old_holiday_prefix and old_holiday_prefix in val:
                template_rows.append(r)
        
        if template_rows:
            next_cond_start = get_last_data_row(ws_cond) + 1
            # User example: if 260509 was 8601-8924. 
            # 8601 -> 8701 etc? 'Each group is incrementing by 100'.
            for i, src_r in enumerate(template_rows):
                tgt_r = next_cond_start + i
                for c in range(1, ws_cond.max_column + 1):
                    src_c = ws_cond.cell(row=src_r, column=c)
                    dst_c = ws_cond.cell(row=tgt_r, column=c)
                    dst_c.value = src_c.value
                    copy_cell_format(src_c, dst_c)
                
                # Alias replace
                alias = str(ws_cond.cell(row=tgt_r, column=2).value)
                ws_cond.cell(row=tgt_r, column=2).value = alias.replace(old_holiday_prefix, holiday_id)
                # ID offset: verified against every historical cycle transition
                # (250701 through 260828, dozens of them) — each new cycle's IDs
                # are always +400 from the same position in the immediately
                # preceding cycle. This matches the sheet's own original Excel
                # formulas (e.g. row 422 = "=A338+400", 84 rows = one cycle back).
                # +100 was wrong: it only advances one "sub-group slot" (BP1/2/3/4
                # each start 100 apart within a cycle), causing IDs to collide
                # with an existing cycle 2-3 holiday cycles later.
                ws_cond.cell(row=tgt_r, column=1).value = int(ws_cond.cell(row=src_r, column=1).value) + 400

        wb_bp.save(bp_path)

        # --- 5. item.xlsx ---
        wb_item = openpyxl.load_workbook(item_path)
        ws_item = wb_item['item']
        item_block = find_last_matching_block(ws_item, r'^\d{6}节日BP通行证')
        insert_item_r = (item_block[-1] + 1) if item_block else (ws_item.max_row + 1)
        ws_item.insert_rows(insert_item_r)
        ws_item.cell(row=insert_item_r, column=1).value = new_item_id
        ws_item.cell(row=insert_item_r, column=2).value = f"{holiday_id}节日BP通行证"
        ws_item.cell(row=insert_item_r, column=3).value = "解锁道具"
        if item_block:
            template_r = item_block[-1]
            for c in range(1, ws_item.max_column + 1):
                copy_cell_format(ws_item.cell(row=template_r, column=c), ws_item.cell(row=insert_item_r, column=c))
        wb_item.save(item_path)

        # --- 6. store.xlsx ---
        wb_store = openpyxl.load_workbook(store_path)

        # store_recharge: clone the most recent holiday-cycle block, wherever it
        # currently sits (never a hardcoded row — the sheet grows every cycle).
        ws_sr = wb_store['store_recharge']
        src_range_sr = find_last_matching_block(ws_sr, r'^\d{6}节日BP通行证')
        if not src_range_sr:
            raise ValueError("store_recharge: no previous 节日BP通行证 block found to clone from")
        insert_sr = src_range_sr[-1] + 1
        num_sr = len(src_range_sr)
        ws_sr.insert_rows(insert_sr, num_sr)
        for i, src_r in enumerate(src_range_sr):
            tgt_r = insert_sr + i
            for c in range(1, ws_sr.max_column + 1):
                src_c = ws_sr.cell(row=src_r, column=c)
                dst_c = ws_sr.cell(row=tgt_r, column=c)
                dst_c.value = src_c.value
                copy_cell_format(src_c, dst_c)

            # Store Item ID (Col 1): offset by the block size, matching the
            # spacing observed between consecutive real holiday cycles.
            ws_sr.cell(row=tgt_r, column=1).value = int(ws_sr.cell(row=src_r, column=1).value) + num_sr
            # Swap Alias (Col 2 & 3)
            ws_sr.cell(row=tgt_r, column=2).value = str(ws_sr.cell(row=src_r, column=2).value).replace(old_holiday_prefix, holiday_id)
            ws_sr.cell(row=tgt_r, column=3).value = str(ws_sr.cell(row=src_r, column=3).value).replace(old_holiday_prefix, holiday_id)

        # store_package: same dynamic-block approach as store_recharge.
        ws_sp = wb_store['store_package']
        src_range_sp = find_last_matching_block(ws_sp, r'^\d{6}节日BP通行证')
        if not src_range_sp:
            raise ValueError("store_package: no previous 节日BP通行证 block found to clone from")
        insert_sp = src_range_sp[-1] + 1
        num_sp = len(src_range_sp)
        ws_sp.insert_rows(insert_sp, num_sp)
        for i, src_r in enumerate(src_range_sp):
            tgt_r = insert_sp + i
            for c in range(1, ws_sp.max_column + 1):
                src_c = ws_sp.cell(row=src_r, column=c)
                dst_c = ws_sp.cell(row=tgt_r, column=c)
                dst_c.value = src_c.value
                copy_cell_format(src_c, dst_c)

            ws_sp.cell(row=tgt_r, column=1).value = int(ws_sp.cell(row=src_r, column=1).value) + num_sp
            ws_sp.cell(row=tgt_r, column=2).value = str(ws_sp.cell(row=src_r, column=2).value).replace(old_holiday_prefix, holiday_id)
            # Content String sync
            content = str(ws_sp.cell(row=tgt_r, column=3).value)
            ws_sp.cell(row=tgt_r, column=3).value = content.replace(old_holiday_prefix, holiday_id)

        # recommend_gift: clone the most recent 节日BP推荐 row, wherever it sits.
        ws_rg = wb_store['recommend_gift']
        rg_block = find_last_matching_block(ws_rg, r'^\d{6}节日BP推荐')
        if not rg_block:
            raise ValueError("recommend_gift: no previous 节日BP推荐 row found to clone from")
        prev_rg_r = rg_block[-1]
        insert_rg = prev_rg_r + 1
        ws_rg.insert_rows(insert_rg)
        for c in range(1, ws_rg.max_column + 1):
            src_c = ws_rg.cell(row=prev_rg_r, column=c)
            dst_c = ws_rg.cell(row=insert_rg, column=c)
            dst_c.value = src_c.value
            copy_cell_format(src_c, dst_c)

        ws_rg.cell(row=insert_rg, column=1).value = recommend_id
        ws_rg.cell(row=insert_rg, column=2).value = f"{holiday_id}节日BP推荐"
        # Recharge ID list — pulled from the just-inserted store_recharge rows
        recharge_ids = [str(int(ws_sr.cell(row=insert_sr + j, column=1).value)) for j in range(num_sr)]
        ws_rg.cell(row=insert_rg, column=3).value = ",".join(recharge_ids)

        wb_store.save(store_path)
        print("Successfully updated Holiday BattlePass configuration.")

    except Exception as e:
        print(f"Fatal Error: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
