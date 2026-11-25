import sys
import os
import json
import openpyxl
from xlsx_safe import load_workbook_frozen, mini_mgr_columns, mini_mgr_ab_columns

def main():
    sys.stdout.reconfigure(encoding='utf-8')
    try:
        design_path = sys.argv[1]
        payload_path = sys.argv[2]
    except IndexError:
        print("Error: Missing args")
        sys.exit(1)
        
    excel_path = os.path.join(design_path, "mini_mgr.xlsx")
    
    if not os.path.exists(excel_path):
        print(f"Error: mini_mgr.xlsx not found at {excel_path}")
        sys.exit(1)

    try:
        with open(payload_path, 'r', encoding='utf-8-sig') as f:
            payload = json.load(f)
    except Exception as e:
        print(f"Error reading payload: {e}")
        sys.exit(1)

    try:
        wb = load_workbook_frozen(excel_path)
    except Exception as e:
        print(f"Error loading workbook: {str(e)}")
        sys.exit(1)

    ws_data = wb["活动数据"] if "活动数据" in wb.sheetnames else None
    if not ws_data:
        print("Error: Missing 活动数据")
        sys.exit(1)

    header = [c.value for c in ws_data[1]]
    try:
        cols = mini_mgr_columns(header)
    except ValueError as e:
        print(f"Error: {e}")
        sys.exit(1)
    ab_cols = mini_mgr_ab_columns(header)

    # Rows are hand-editable in the preview, so check them before touching the sheet.
    existing_ids = {str(ws_data.cell(r, cols["id"]).value).strip() for r in range(2, ws_data.max_row + 1)}
    for row_data in payload.get('rows', []):
        if str(row_data['id']).strip() in existing_ids:
            print(f"Error: ID {row_data['id']} already exists in 活动数据 — nothing was written.")
            sys.exit(1)
        unknown = [h for h in (row_data.get('ab') or {}) if h not in ab_cols]
        if unknown:
            print(f"Error: 活动数据 has no AB column {', '.join(unknown)} — nothing was written.")
            sys.exit(1)

    TEMPLATE_IDS = {101, 102, 103}

    # Find the real last row with data (scan backwards to skip ghost empty rows)
    def get_last_data_row(ws):
        for r in range(ws.max_row, 0, -1):
            # A row is considered 'real' if the ID column has a numeric value
            try:
                val = ws.cell(row=r, column=cols["id"]).value
                if val is not None and str(val).strip() != '':
                    int(val)  # must be numeric
                    return r
            except (TypeError, ValueError):
                continue
        return 1

    last_data_row = get_last_data_row(ws_data)

    # 1. Flip old switches to False — but NEVER touch template rows (101, 102, 103)
    for r in range(4, last_data_row + 1):
        row_id = ws_data.cell(row=r, column=cols["id"]).value
        try:
            if int(row_id) in TEMPLATE_IDS:
                continue  # Protect template rows — always stay True
        except (TypeError, ValueError):
            continue
        if ws_data.cell(row=r, column=cols["switch"]).value is not None:
            ws_data.cell(row=r, column=cols["switch"]).value = False

    # 2. Append new rows directly after the last real data row
    start_row = last_data_row + 1
    for row_data in payload.get('rows', []):
        ws_data.cell(row=start_row, column=cols["id"]).value = row_data['id']
        ws_data.cell(row=start_row, column=cols["switch"]).value = row_data.get('switch', True)
        # AB switches: True/False written, None left blank (the sheet's default).
        for h, v in (row_data.get('ab') or {}).items():
            if v is not None:
                ws_data.cell(row=start_row, column=ab_cols[h]).value = bool(v)
        ws_data.cell(row=start_row, column=cols["start"]).value = row_data['start_time']
        ws_data.cell(row=start_row, column=cols["end"]).value = row_data['end_time']
        ws_data.cell(row=start_row, column=cols["type"]).value = row_data['minigame']

        dw_val = row_data.get('double_week_id')
        if dw_val and str(dw_val).strip() != "":
            try:
                ws_data.cell(row=start_row, column=cols["double_week"]).value = int(dw_val)
            except ValueError:
                ws_data.cell(row=start_row, column=cols["double_week"]).value = dw_val

        if row_data.get('discount_id') != -1:
            ws_data.cell(row=start_row, column=cols["discount"]).value = row_data['discount_id']
        
        start_row += 1

    try:
        wb.save(excel_path)
    except PermissionError:
        print(f"ERROR: Cannot save '{excel_path}'.\nThe file is open in Excel or another program. Please close it and try again.")
        sys.exit(1)
    print(f"Successfully appended {len(payload.get('rows', []))} Minigame rows to Excel.")

if __name__ == "__main__":
    main()
