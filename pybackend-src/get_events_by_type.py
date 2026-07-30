import sys
import os
import re
import json
import openpyxl

# Fixed display order for the event-type picker. Any other type found in the
# sheet (Hunt, Drink, Clover, Garden, Snake, Clothes, Monopoly, ...) is still
# returned under "types" so the caller isn't silently missing data.
KNOWN_TYPES = ["Merge", "Sim", "Blocks", "Sort", "Dessert"]

ALIAS_RE = re.compile(r'^([A-Za-z]+)(\d{4})')


def main():
    sys.stdout.reconfigure(encoding='utf-8')
    try:
        design_path = sys.argv[1]
    except IndexError:
        print(json.dumps({"error": "Missing design folder path"}))
        sys.exit(1)

    events_path = os.path.join(design_path, "events.xlsx")
    if not os.path.exists(events_path):
        print(json.dumps({"error": f"events.xlsx not found at {events_path}"}))
        sys.exit(1)

    try:
        wb = openpyxl.load_workbook(events_path, data_only=True)
        ws = wb["events"]
    except Exception as e:
        print(json.dumps({"error": f"Failed to open events.xlsx: {e}"}))
        sys.exit(1)

    events_by_type = {}
    seen_types = []

    # Sheet rows are chronological (ascending), so the last row collected for
    # a type is that type's most recent event — matches the "latest" convention
    # already used elsewhere in this tool (get_last_numeric_id, etc.).
    for r in range(2, ws.max_row + 1):
        alias = ws.cell(r, 2).value
        if not alias:
            continue
        alias = str(alias).strip()
        m = ALIAS_RE.match(alias)
        if not m:
            continue

        type_name = m.group(1)
        code = m.group(2)
        event_name = f"{type_name.lower()}{code}"
        start_time = ws.cell(r, 8).value
        event_id = ws.cell(r, 1).value

        if type_name not in events_by_type:
            events_by_type[type_name] = []
            seen_types.append(type_name)

        events_by_type[type_name].append({
            "name": event_name,
            "alias": alias,
            "event_id": str(event_id) if event_id is not None else "",
            "start_time": str(start_time) if start_time is not None else ""
        })

    ordered_types = KNOWN_TYPES + [t for t in seen_types if t not in KNOWN_TYPES]

    print(json.dumps({
        "types": ordered_types,
        "events_by_type": events_by_type
    }, ensure_ascii=False))


if __name__ == "__main__":
    main()
