"""
PVP Bots Card Back Update — Task 2 of Decoration Shop Config.

Mechanism (reverse-engineered from Robot_Calc.xlsx's live RAND()/LOOKUP formulas,
confirmed against real data, not guessed):

Robot_Calc.xlsx holds two weighted-draw tables on its first sheet:
  - Frame table:    K11:K101ish (weight), L (cumulative boundary), M (name)
  - Card back table: P (weight), Q (cumulative boundary), R (name)
Each table always ends in a "dangling placeholder" row: boundary (L/Q) filled,
name (M/R) and weight (K/P) empty — reserved for the next skin to be added.

Excel's per-bot draw is: I = ROUND(RAND()*SUM(K:K)*2.5, 0); G = LOOKUP(I, L, M)
(blank if I falls in the placeholder's range, meaning "no change this cycle").
This is deterministic enough to reproduce directly in Python — bisect I against
the boundary list; a draw landing at/after the last real boundary leaves the
bot's existing assignment untouched (matches the documented "inflation factor
intentionally produces some blank draws" design).

This script:
  1. Appends the new skin(s) to Robot_Calc.xlsx's weight tables (fills the
     dangling placeholder, appends a fresh one for next cycle).
  2. Re-rolls every real bot in robot.xlsx's `robot` sheet using that mechanism,
     writing results via HEADER-NAME lookup (not a fixed column index — column
     position for 机器人头像框/机器人卡背 is confirmed to differ across branches).
"""
import sys
import os
import stat
import json
import random
import bisect
import openpyxl
from xlsx_safe import load_workbook_frozen


def get_real_last_row(ws, id_col=1):
    for r in range(ws.max_row, 0, -1):
        v = ws.cell(r, id_col).value
        if v is not None and str(v).strip() != "":
            return r
    return 1


def find_header_col(ws, header_name, header_row=1):
    for c in range(1, ws.max_column + 1):
        if str(ws.cell(header_row, c).value or "").strip() == header_name:
            return c
    raise ValueError(f"robot.xlsx: header '{header_name}' not found in row {header_row}")


def extend_weight_table(ws, ws_values, weight_col, boundary_col, name_col, new_names, weights, start_row=3):
    """Fill the dangling placeholder (boundary filled, name empty) with the first
    new name, then append a fresh row per additional name, ending with a new
    dangling placeholder for the next cycle. Returns (entries, k_sum, boundary_max)
    where entries is the full ascending (boundary, name) list including the ones
    just added, k_sum is the sum of all real weights, boundary_max is the final
    placeholder's boundary (the real/blank cutoff for the draw).

    `ws` is the editable (data_only=False) sheet — writes land here so existing
    formulas elsewhere in the workbook survive the save. `ws_values` is a
    data_only=True load of the same file, used only to resolve boundary/weight
    cells that are themselves formulas (confirmed: the placeholder boundary is
    typically "=L{prev}+K{prev}", not a literal)."""
    last_name_row = None
    for r in range(start_row, ws_values.max_row + 1):
        if ws_values.cell(r, name_col).value is not None:
            last_name_row = r
    if last_name_row is None:
        raise ValueError("weight table: no existing named entries found")

    ph_row = last_name_row + 1
    ph_boundary = ws_values.cell(ph_row, boundary_col).value
    if ph_boundary is None:
        # No dangling placeholder present — bootstrap one from the last real boundary.
        prev_boundary = ws_values.cell(last_name_row, boundary_col).value
        ph_boundary = prev_boundary + weights[0]
        ws.cell(ph_row, boundary_col).value = ph_boundary

    # Existing entries, read from the resolved (data_only) copy.
    entries = []
    k_sum = 0
    for r in range(start_row, last_name_row + 1):
        name = ws_values.cell(r, name_col).value
        boundary = ws_values.cell(r, boundary_col).value
        weight = ws_values.cell(r, weight_col).value
        if name is not None:
            entries.append((boundary, name))
        if weight is not None:
            k_sum += weight

    # Newly added entries, computed directly (no need to read them back).
    for name, weight in zip(new_names, weights):
        ws.cell(ph_row, name_col).value = name
        ws.cell(ph_row, weight_col).value = weight
        entries.append((ph_boundary, name))
        k_sum += weight
        next_row = ph_row + 1
        next_boundary = ph_boundary + weight
        ws.cell(next_row, boundary_col).value = next_boundary
        ph_row, ph_boundary = next_row, next_boundary

    return entries, k_sum, ph_boundary


def roll(entries, k_sum, boundary_max, inflation):
    """One weighted draw. Returns the drawn name, or None if the draw landed in
    the placeholder's blank range (meaning: leave the bot's value unchanged)."""
    if k_sum <= 0:
        return None
    i = round(random.random() * k_sum * inflation)
    if i >= boundary_max:
        return None
    boundaries = [b for b, _ in entries]
    idx = bisect.bisect_right(boundaries, i) - 1
    if idx < 0:
        return None
    return entries[idx][1]


def main():
    try:
        design_path = sys.argv[1]
        payload_path = sys.argv[2]
    except IndexError:
        print("Error: Missing args")
        sys.exit(1)

    with open(payload_path, "r", encoding="utf-8-sig") as f:
        payload = json.load(f)

    robot_calc_path = payload["robot_calc_path"]
    frame_aliases = [a for a in payload.get("frame_aliases", []) if a]
    cardback_alias = payload.get("cardback_alias") or None
    frame_weight = int(payload.get("frame_weight", 30))
    cardback_weight = int(payload.get("cardback_weight", 30))

    robot_path = os.path.join(design_path, "robot.xlsx")
    for p in [robot_calc_path, robot_path]:
        if not os.path.exists(p):
            print(f"Error: required file not found: {p}")
            sys.exit(1)
    os.chmod(robot_calc_path, stat.S_IWRITE | stat.S_IREAD)
    os.chmod(robot_path, stat.S_IWRITE | stat.S_IREAD)

    try:
        # Deliberately NOT load_workbook_frozen: Robot_Calc.xlsx is a live designer
        # calculator whose RAND()/LOOKUP formulas are the point of the file. Baking
        # them to literals would freeze the draw and break the sheet. The trade-off is
        # that this save drops the cached results, so the boundary values read on the
        # NEXT run come back empty until someone opens the file in Excel and saves.
        wb_calc = openpyxl.load_workbook(robot_calc_path, data_only=False)
        ws_calc = wb_calc[wb_calc.sheetnames[0]]
        # Separate data_only load to resolve boundary/weight cells that are formulas
        # (confirmed: the placeholder boundary is typically "=L{prev}+K{prev}").
        # Editing lands on ws_calc so every other formula in the sheet survives the save.
        wb_calc_values = openpyxl.load_workbook(robot_calc_path, data_only=True)
        ws_calc_values = wb_calc_values[wb_calc_values.sheetnames[0]]

        frame_entries, frame_k_sum, frame_boundary_max = None, None, None
        if frame_aliases:
            frame_entries, frame_k_sum, frame_boundary_max = extend_weight_table(
                ws_calc, ws_calc_values, weight_col=11, boundary_col=12, name_col=13,
                new_names=frame_aliases, weights=[frame_weight] * len(frame_aliases),
            )
            print(f"Robot_Calc.xlsx: added frame(s) {frame_aliases} at weight {frame_weight} each")

        cardback_entries, cardback_k_sum, cardback_boundary_max = None, None, None
        if cardback_alias:
            cardback_entries, cardback_k_sum, cardback_boundary_max = extend_weight_table(
                ws_calc, ws_calc_values, weight_col=16, boundary_col=17, name_col=18,
                new_names=[cardback_alias], weights=[cardback_weight],
            )
            print(f"Robot_Calc.xlsx: added card back '{cardback_alias}' at weight {cardback_weight}")

        wb_calc.save(robot_calc_path)

        wb_robot = load_workbook_frozen(robot_path)
        ws_robot = wb_robot["robot"]
        frame_col = find_header_col(ws_robot, "机器人头像框")
        cardback_col = find_header_col(ws_robot, "机器人卡背")
        id_col = find_header_col(ws_robot, "机器人数据ID")
        last_row = get_real_last_row(ws_robot, id_col=id_col)

        frame_changed = cardback_changed = 0
        for r in range(2, last_row + 1):
            if frame_entries is not None:
                new_frame = roll(frame_entries, frame_k_sum, frame_boundary_max, 2.5)
                if new_frame is not None:
                    ws_robot.cell(r, frame_col).value = new_frame
                    frame_changed += 1
            if cardback_entries is not None:
                new_cardback = roll(cardback_entries, cardback_k_sum, cardback_boundary_max, 2.1)
                if new_cardback is not None:
                    ws_robot.cell(r, cardback_col).value = new_cardback
                    cardback_changed += 1

        wb_robot.save(robot_path)
        print(f"robot.xlsx: re-rolled {last_row - 1} bots, "
              f"{frame_changed} frame changes, {cardback_changed} card back changes")
        print("Successfully updated PVP Bots Card Back configuration.")

    except Exception as e:
        print(f"Error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
