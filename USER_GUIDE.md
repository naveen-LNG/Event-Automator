# Event Automation Tool: User & Installation Guide

This guide provides everything the team needs to install, configure, and operate the SkipBo Event Automation Tool.

---

## 1. Prerequisites & Setup

### **Python Environment**
The backend logic is powered by Python. Ensure you have the following installed:
- **Python 3.8+**: [Download here](https://www.python.org/downloads/)
- **Required Library**: `openpyxl` (for Excel manipulation)
  ```bash
  pip install openpyxl
  ```

### **Unity Project Integration**
1. Copy the `EventAutomation` folder directly into your project's `Assets/Editor/` directory.
2. The tool should be accessible via the top menu: **Tools > Event Automator**.

---

## 2. Interface Overview

### **Global Settings**
- **Mode Selection**: 
  - `Reopen`: Use this when an event already exists but needs its dates/IDs refreshed. 
  - `New Event`: Use this to clone an old event into a completely new one (e.g., cloning `Sim2507` to create `Sim2605`).
- **Source Event**: The name of the reference event (usually the one you are cloning *from*).
- **Target Event**: The name of the new event being created.
- **Workspace Path**: The **absolute path** to your project root (e.g., `C:\SkipBo\branches\Misc`).

### **New Event Settings**
Only required when in `New Event` mode:
- **Event ID**: The unique 8-digit identifier from the design document (e.g., `20260501`).
- **Timestamps**: `Start`, `End`, `Near End`, and `Close` times. These follow the `YYYY-MM-DD HH:MM:SS` format.

---

## 3. The 21-Step Pipeline Breakdown

The tool automates these critical steps sequentially:

1.  **Configure `storagedata.proto`**: Injects the new event ID into the Protobuf definition.
2.  **Execute ProtoToLua**: Triggers the Unity build-in command to regenerate Lua proto files.
3.  **Add Save File**: Clones and renames the `LocalService.lua` file for the new event.
4.  **Register LocalService**: Injects the new event into `LocalServiceMgr.lua`.
5.  **Clone WorkSpace3**: Ensures `WorkSpcae_Design` is synchronized with the latest project data.
6.  **Clone Event Excels**: Copies the event-specific configuration tables (e.g., `Sim2XXXX.xlsx`).
7.  **Copy Descriptor Scripts**: Migrates the event descriptor JSONs to the client directory.
8.  **Remap Convert References**: Updates `convert_layout.lua` and `convert.json` for the new event name.
9.  **Update `bi.xlsx`**: Registers new BI tracking codes for the event.
10. **Update `events.xlsx`**: **Crucial Step.** Registers the event in the central registry and syncs the 8-digit ID.
11. **Update `event_shop.xlsx`**: Clones shop entries for the new event.
12. **Update `item.xlsx`**: Generates new event-specific item IDs (e.g., Tokens/Materials) and **automatically syncs them back** to `events.xlsx`.
13. **Update `icon.xlsx`**:
    *   Clones icon references from backup sheets to active sheets.
    *   Automatically duplicates Token/Material templates (Rows 69-70).
    *   Updates Sprite Name IDs to match the GUI Event ID.
14. **Update Localization**: Refreshes strings, Quiz data, and Answer Challenge tables for the new event dates.
15. **Update `asset_ref.xlsx`**: Remaps asset paths for prefabs and UI elements.
16. **Update `store.xlsx`**: Clones store packages and pricing for the new event.
17. **Update `card.xlsx`**: Reserved for card-specific event configurations.
18. **Update `pack.xlsx`**: Clones deal/pack configurations.
19. **Update `guide.xlsx`**: Sets up new guide/tutorial triggers for the event.
20. **Update `sys.xlsx`**: Registers the DeepLink and UI shortcut IDs in the system registry.
21. **Execute `coder_convert`**: Triggers the project's data conversion tool to bake all changes.

---

## 4. Special Automation Features

### **Automatic Item ID Sync**
When Step 12 runs, it calculates a new ID for items like `TOKEN_SIM2605`. It then automatically finds the corresponding row in `events.xlsx` and updates the `掉落的道具ID` column to match. **No manual ID copying required.**

### **Icon Template Logic**
In `icon_item`, the tool specifically targets rows 69 and 70 as "Templates." It creates a copy for your target event, handles the renaming, and inserts them **above** the `// BP相关` section to keep the spreadsheet organized.

---

## 5. Troubleshooting & Maintenance

- **"Permission Denied"**: Close your Excel files before running the tool. The script needs to "Save" the workbooks, which is blocked if they are open in Excel.
- **Log Files**: If a step fails, check the `PythonBackend/` folder for `.log` or `extra_args.json` files to see exactly what data was passed to Python.
- **Python Errors**: If a step returns an error about `openpyxl`, ensure you've run `pip install openpyxl`.
- **Manual Overrides**: If the automation fails on a specific step, you can fix the issue manually in Excel and then use the **"Run"** button in Unity to retry only that step.

---

> [!TIP]
> **Pro-Tip**: Always verify your Target Event name format (e.g., `Sim2605`) matches exactly with how other events are named in the shared spreadsheets to ensure pattern matching works perfectly.
