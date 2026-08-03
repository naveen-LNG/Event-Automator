import sys
import os
import json
import openpyxl

def main():
    sys.stdout.reconfigure(encoding='utf-8')
    try:
        design_path = sys.argv[1]
    except IndexError:
        print(json.dumps({"error": "Missing design folder path"}))
        sys.exit(1)
        
    excel_path = os.path.join(design_path, "mini_mgr.xlsx")
    
    if not os.path.exists(excel_path):
        print(json.dumps({"error": f"mini_mgr.xlsx not found at {excel_path}"}))
        sys.exit(1)

    try:
        wb = openpyxl.load_workbook(excel_path, data_only=True)
    except Exception as e:
        print(json.dumps({"error": f"Failed to load workbook: {str(e)}"}))
        sys.exit(1)

    result = {
        "max_id": 0,
        "minigames": []
    }

    ws_data = wb["活动数据"] if "活动数据" in wb.sheetnames else None
    ws_info = wb["活动信息"] if "活动信息" in wb.sheetnames else None

    if not ws_data or not ws_info:
        print(json.dumps({"error": "Missing required sheets in mini_mgr.xlsx"}))
        sys.exit(1)

    # 1. Get max ID from 活动数据
    max_id = 0
    # Search from the bottom up to find the last valid numeric ID
    for r in range(ws_data.max_row, 1, -1):
        val = ws_data.cell(row=r, column=1).value
        try:
            val_int = int(val)
            max_id = val_int
            break
        except (ValueError, TypeError):
            continue
            
    result["max_id"] = max_id

    # 2. Build history dict from 活动数据 to find 折扣链ID and 双周活动ID requirement
    history_map = {}
    
    # Columns in 活动数据 (1-indexed, matches real header row):
    # column=6 is 活动类型 (Minigame Name), column=7 is 双周活动ID, column=8 is 折扣链ID
    for r in range(ws_data.max_row, 3, -1):
        mg_name = ws_data.cell(row=r, column=6).value
        if not mg_name or not isinstance(mg_name, str):
            continue

        discount_id = ws_data.cell(row=r, column=8).value
        double_week_id = ws_data.cell(row=r, column=7).value
        
        if mg_name not in history_map:
            history_map[mg_name] = {
                "discount_id": discount_id if discount_id is not None else -1,
                "requires_double_week": double_week_id is not None and str(double_week_id).strip() != "None" and str(double_week_id).strip() != ""
            }

    # 3. Get all available minigames from 活动信息 (Column 2 is Alias/Name, Column 3 is internal name)
    # The requirement is that dropdown shows names from "活动信息" (user probably meant Alias, which is column 3 (index 2))
    # Wait, the user said: "available minigames from the sheet 活动信息". Let's look at row 4: [3, '大富翁', 'gembrawl', 10186, 20009]
    # So column 2 is the Chinese Name (大富翁, 答题挑战, 挖蘑菇, etc.). We should use the Chinese name because that's what's written to 活动数据.
    
    for r in range(4, ws_info.max_row + 1):
        mg_name = ws_info.cell(row=r, column=2).value
        if not mg_name or not str(mg_name).strip():
            continue
        mg_name = str(mg_name).strip()

        # Merge heuristic data if available
        heuristic = history_map.get(mg_name, {"discount_id": -1, "requires_double_week": False})

        result["minigames"].append({
            "name": mg_name,
            "discount_id": heuristic["discount_id"],
            "requires_double_week": heuristic["requires_double_week"]
        })

    # 4. Detect the recurring 5-minigame set pattern (verified against real production history):
    # position 1 alternates between two minigames; positions 2-5 are always the same set,
    # each minigame running back-to-back (1 day gap) for a fixed duration.
    # ponytail: fixed template, not derived generically — matches the last ~5 real cycles;
    # revisit if the studio's rotation itself changes shape.
    SET_ALTERNATES = ("赢钱排行榜", "连胜空岛")
    SHARED_TAIL = [("挖蘑菇", 3), ("答题挑战", 1), ("大富翁", 3), ("Merge_Mini", 3)]

    last_position1 = None
    for r in range(ws_data.max_row, 3, -1):
        mg_name = ws_data.cell(row=r, column=6).value
        if mg_name in SET_ALTERNATES:
            last_position1 = mg_name
            break
    next_first = SET_ALTERNATES[1] if last_position1 == SET_ALTERNATES[0] else SET_ALTERNATES[0]

    last_quiz_dw = None
    for r in range(ws_data.max_row, 3, -1):
        if str(ws_data.cell(row=r, column=6).value or "").strip() == "答题挑战":
            dw_val = ws_data.cell(row=r, column=7).value
            if dw_val is not None and str(dw_val).strip() != "":
                last_quiz_dw = str(dw_val).strip()
                break

    result["next_set_template"] = [{"name": next_first, "duration_days": 3}] + [
        {"name": name, "duration_days": days} for name, days in SHARED_TAIL
    ]
    result["gap_days"] = 1
    result["last_quiz_dw"] = last_quiz_dw

    print(json.dumps(result, ensure_ascii=False))

if __name__ == "__main__":
    main()
