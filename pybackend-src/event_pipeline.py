import os
import sys
import shutil
import json
import traceback
import re
import copy as _copy
import builtins
from xlsx_safe import load_workbook_frozen, mini_mgr_columns


def open(file, mode='r', *args, **kwargs):  # noqa: A001 - deliberately shadows builtins.open
    """Every text-mode write in this module goes through here. Python's text mode on Windows
    turns each '\\n' into '\\r\\n', so re-saving an LF file (storagedata.proto; convert.json in
    WorkSpcae3/WorkSpcae_NE) silently converts ALL its line endings — SVN then shows a
    whole-file change. An existing file keeps its own line endings; new files are unchanged.
    ponytail: sniffs the first 64 KB; a file that mixes endings is treated as CRLF (the old behaviour)."""
    if 'w' in mode and 'b' not in mode and 'newline' not in kwargs and len(args) < 3:
        try:
            with builtins.open(file, 'rb') as f:
                head = f.read(65536)
            if b'\n' in head and b'\r\n' not in head:
                kwargs['newline'] = '\n'
        except OSError:
            pass
    return builtins.open(file, mode, *args, **kwargs)


WORKSPACE_NAME = "WorkSpcae_Design"
UNDO_LOG_FILE = "undo_log.json"

def copy_cell_format(src_cell, dst_cell):
    """Copy number format + style from src_cell to dst_cell. Call this after
    writing dst_cell.value when cloning a row — plain value-only writes leave
    the destination cell's format at openpyxl's default ("General"), which
    silently diverges from source cells formatted as e.g. Text."""
    dst_cell.number_format = src_cell.number_format
    if src_cell.has_style:
        dst_cell.font = _copy.copy(src_cell.font)
        dst_cell.fill = _copy.copy(src_cell.fill)
        dst_cell.border = _copy.copy(src_cell.border)
        dst_cell.alignment = _copy.copy(src_cell.alignment)
        dst_cell.protection = _copy.copy(src_cell.protection)

def cell_style_to_dict(cell):
    """JSON-serializable snapshot of a cell's style, for row extractors that must
    survive a round-trip through a cache file or an SVN-history byte blob (where
    the original Cell object is gone by the time the destination row is written)."""
    if not cell.has_style:
        return None
    f, fl, a, b, p = cell.font, cell.fill, cell.alignment, cell.border, cell.protection
    def rgb_of(color):
        # color.rgb can be a str-like RGB, an int (indexed color), or None — only the
        # str case round-trips through JSON and back into a usable Color(rgb=...).
        if color is None:
            return None
        v = color.rgb
        return str(v) if isinstance(v, str) else None
    def side(s):
        return {"style": s.style, "color": rgb_of(s.color)} if s else None
    return {
        "number_format": cell.number_format,
        "font": {"name": f.name, "size": f.size, "bold": f.bold, "italic": f.italic,
                 "underline": f.underline, "color": rgb_of(f.color)},
        "fill": {"patternType": fl.patternType, "fgColor": rgb_of(fl.fgColor), "bgColor": rgb_of(fl.bgColor)},
        "alignment": {"horizontal": a.horizontal, "vertical": a.vertical, "wrap_text": a.wrap_text},
        "border": {"left": side(b.left), "right": side(b.right), "top": side(b.top), "bottom": side(b.bottom)},
        "protection": {"locked": p.locked, "hidden": p.hidden},
    }

def apply_cell_style_dict(cell, style):
    """Inverse of cell_style_to_dict — writes a captured style snapshot onto a
    freshly-written destination cell."""
    if not style:
        return
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, Protection
    from openpyxl.styles.colors import Color
    def mk_color(rgb):
        return Color(rgb=rgb) if rgb else Color()
    cell.number_format = style.get("number_format") or "General"
    f = style.get("font") or {}
    cell.font = Font(name=f.get("name"), size=f.get("size"), bold=f.get("bold"),
                      italic=f.get("italic"), underline=f.get("underline"), color=mk_color(f.get("color")))
    fl = style.get("fill") or {}
    if fl.get("patternType"):
        cell.fill = PatternFill(patternType=fl.get("patternType"),
                                 fgColor=mk_color(fl.get("fgColor")), bgColor=mk_color(fl.get("bgColor")))
    a = style.get("alignment") or {}
    cell.alignment = Alignment(horizontal=a.get("horizontal"), vertical=a.get("vertical"), wrap_text=a.get("wrap_text"))
    b = style.get("border") or {}
    def mk_side(d):
        return Side(style=d.get("style"), color=mk_color(d.get("color"))) if d else Side()
    cell.border = Border(left=mk_side(b.get("left")), right=mk_side(b.get("right")),
                          top=mk_side(b.get("top")), bottom=mk_side(b.get("bottom")))
    pr = style.get("protection") or {}
    cell.protection = Protection(locked=pr.get("locked", True), hidden=pr.get("hidden", False))

def log(msg):
    try:
        print(f"[PythonBackend] {msg}", flush=True)
    except UnicodeEncodeError:
        safe_msg = str(msg).encode('ascii', errors='replace').decode('ascii')
        print(f"[PythonBackend] {safe_msg}", flush=True)

def load_undo_log(root_path):
    log_path = os.path.join(root_path, UNDO_LOG_FILE)
    if os.path.exists(log_path):
        with open(log_path, 'r') as f:
            return json.load(f)
    return {}

def save_undo_log(root_path, undo_data):
    log_path = os.path.join(root_path, UNDO_LOG_FILE)
    with open(log_path, 'w') as f:
        json.dump(undo_data, f, indent=4)

def record_file_creation(root_path, step_index, file_path):
    undo_data = load_undo_log(root_path)
    step_key = str(step_index)
    
    if step_key not in undo_data:
        undo_data[step_key] = {"created": [], "modified": []}
        
    if file_path not in undo_data[step_key]["created"]:
        undo_data[step_key]["created"].append(file_path)
        
    save_undo_log(root_path, undo_data)

def record_file_backup(root_path, step_index, file_path):
    """Back up a file this step is about to edit IN PLACE and register it for revert_step.
    Unlike the .bak-next-to-the-file convention, the copy goes in an "undo_backups~" folder:
    root_path is inside Assets/, and Unity never imports folders ending in "~" — a .bak copy
    of a prefab sitting next to it would be imported as an asset of its own."""
    undo_data = load_undo_log(root_path)
    entry = undo_data.setdefault(str(step_index), {"created": [], "modified": []})
    mods = entry.setdefault("modified", [])
    if any(isinstance(m, dict) and m.get("orig") == file_path for m in mods):
        return  # the first backup in a run holds the true pre-step state
    bdir = os.path.join(root_path, "undo_backups~", str(step_index))
    os.makedirs(bdir, exist_ok=True)
    backup = os.path.join(bdir, f"{len(mods)}_{os.path.basename(file_path)}")
    shutil.copy2(file_path, backup)
    mods.append({"orig": file_path, "backup": backup})
    save_undo_log(root_path, undo_data)

def revert_step(root_path, step_index):
    undo_data = load_undo_log(root_path)
    step_key = str(step_index)

    if step_key not in undo_data:
        log(f"No undo data found for Step {step_index}")
        return True

    files_created = undo_data[step_key].get("created", [])
    import stat

    for m in undo_data[step_key].get("modified", []):
        if isinstance(m, dict) and os.path.exists(m.get("backup", "")):
            if os.path.exists(m["orig"]):
                os.chmod(m["orig"], stat.S_IWRITE | stat.S_IREAD)
            shutil.copy2(m["backup"], m["orig"])
            os.remove(m["backup"])
            log(f"Reverted (Restored edited file): {os.path.basename(m['orig'])}")

    for f in files_created:
        if os.path.exists(f):
            if f.endswith(".bak"):
                # It's a backup file, restore it
                orig_file = f[:-4]
                if os.path.exists(orig_file):
                    # Remove read-only from original before overwrite
                    os.chmod(orig_file, stat.S_IWRITE | stat.S_IREAD)
                shutil.copy2(f, orig_file)
                os.chmod(f, stat.S_IWRITE)
                os.remove(f)
                log(f"Reverted (Restored from bak): {os.path.basename(orig_file)}")
            else:
                try:
                    if os.path.isdir(f):
                        shutil.rmtree(f)
                        log(f"Reverted (Deleted Dir): {os.path.basename(f)}")
                    else:
                        os.chmod(f, stat.S_IWRITE)
                        os.remove(f)
                        log(f"Reverted (Deleted File): {os.path.basename(f)}")
                except Exception as e:
                    log(f"Failed to revert {f}: {e}")
                
    # Remove entry
    del undo_data[step_key]
    save_undo_log(root_path, undo_data)
    return True

def get_real_max_row(ws, col=1):
    """Find the actual last row with data in a given column."""
    for r in range(ws.max_row, 0, -1):
        if ws.cell(r, col).value is not None and str(ws.cell(r, col).value).strip() != "":
            return r
    return 0

def get_real_max_col(ws, scan_rows=5, scan_limit=200):
    """Find the actual last used column by scanning the first few rows, bounded
    to scan_limit. Some sheets report a bloated ws.max_column (stray formatting
    applied to whole columns, no real data — e.g. bi.xlsx's 广告打点常量表 reports
    16382 columns with only 5 ever populated) that makes a full-width iteration
    or insert_rows call take tens of seconds for what should be instant. Every
    real sheet in this codebase has a handful of columns, so a bounded forward
    scan is both correct and cheap; a full-width reverse scan (mirroring
    get_real_max_row) would be just as slow as the bug it's fixing."""
    limit = min(ws.max_column, scan_limit)
    last_col = 0
    for r in range(1, min(scan_rows, ws.max_row) + 1):
        for c in range(1, limit + 1):
            val = ws.cell(r, c).value
            if val is not None and str(val).strip() != "" and c > last_col:
                last_col = c
    return last_col if last_col else limit

def get_last_numeric_id(ws, col=1, start_row=2):
    """Get the last numeric ID in a column."""
    last_id = 0
    real_max = get_real_max_row(ws, col)
    for r in range(start_row, real_max + 1):
        val = ws.cell(r, col).value
        if val is not None:
            try:
                num = int(val)
                if num > last_id:
                    last_id = num
            except (ValueError, TypeError):
                pass
    return last_id

def find_target_row_info(ws, tgt_name):
    """Find the row index and comment status for the target event."""
    tgt_lower = tgt_name.lower()
    # Some sheets report a bloated ws.max_column (stray formatting on whole columns,
    # e.g. bi.xlsx's 广告打点常量表 reports 16382 with only ~5 ever populated) — an
    # unbounded row x column scan there measured at ~40 SECONDS for a single call.
    # This function is called from check_and_handle_existing all over the codebase,
    # so the fix belongs here, not in each caller.
    real_max_col = get_real_max_col(ws)
    for r in range(2, ws.max_row + 1):
        for c in range(1, real_max_col + 1):
            val = ws.cell(r, c).value
            if val and tgt_lower in str(val).lower():
                col1_val = str(ws.cell(r, 1).value or "").strip()
                is_commented = col1_val.startswith("//") or col1_val.startswith("--")
                return r, is_commented
    return None, False

def check_and_handle_existing(ws, sheet_name, tgt_name):
    """
    Uncomments the row if it exists but is disabled. 
    Returns (True, row_idx) if it exists, (False, None) otherwise.
    """
    row_idx, is_commented = find_target_row_info(ws, tgt_name)
    if row_idx:
        if is_commented:
            # Uncomment
            col1_val = str(ws.cell(row_idx, 1).value).strip()
            if col1_val.startswith("//"): col1_val = col1_val[2:].strip()
            elif col1_val.startswith("--"): col1_val = col1_val[2:].strip()
            ws.cell(row_idx, 1, col1_val)
            log(f"{sheet_name}: Target existed but was commented. Uncommented row {row_idx}.")
            return True, row_idx
        else:
            log(f"{sheet_name}: Target already exists and is active.")
            return True, row_idx
    return False, None

def copy_and_rename_file(src_path, dest_path, old_name, new_name, root_path, step_index):
    """Copies a file and does a raw string replacement of old_name to new_name in its contents."""
    if not os.path.exists(src_path):
        log(f"Error: Source file not found: {src_path}")
        return False
    if os.path.exists(dest_path):
        # A reopen under the same name makes src and dest the SAME file. Rewriting it and
        # recording it as "created" meant Undo deleted the event's real, committed script.
        log(f"{os.path.basename(dest_path)} already exists — left untouched.")
        return True

    try:
        os.makedirs(os.path.dirname(dest_path), exist_ok=True)
        
        with open(src_path, 'r', encoding='utf-8', errors='ignore') as f:
            content = f.read()
            
        content = content.replace(old_name.lower(), new_name.lower())
        content = content.replace(old_name.capitalize(), new_name.capitalize())
        content = content.replace(old_name.upper(), new_name.upper())

        with open(dest_path, 'w', encoding='utf-8') as f:
            f.write(content)
            
        # Log for undo
        record_file_creation(root_path, step_index, dest_path)
            
        log(f"Successfully copied and patched {os.path.basename(dest_path)}")
        return True
    except Exception as e:
        log(f"Exception during file copy/patch: {e}")
        traceback.print_exc()
        return False

def inject_proto(proto_dir, src_evt, tgt_evt, root_path, step_index):
    log("Injecting Proto configs...")
    passed_proto = os.path.join(proto_dir, "passed_events.proto")
    storage_proto = os.path.join(proto_dir, "storagedata.proto")
    
    if not os.path.exists(passed_proto) or not os.path.exists(storage_proto):
        log("Proto files not found.")
        return False
        
    # Extract the full block belonging to the event (Bucket + any auxiliary messages)
    def extract_event_block(content, evt_name):
        msg_name = f"E{evt_name.capitalize()}Bucket"
        pattern = re.compile(r'message\s+' + msg_name + r'\s*\{')
        match = pattern.search(content)
        if not match: return None
        
        start_idx = match.start()
        
        # Look for the next event bucket to find the boundary of this event's code block
        # Matches newlines followed optionally by comments, then 'message E...Bucket'
        next_bucket_pattern = re.compile(r'\n\s*(?://.*?\n\s*)*message\s+E[A-Z]\w+Bucket\s*\{')
        next_match = next_bucket_pattern.search(content, match.end())
        
        if next_match:
            end_idx = next_match.start()
        else:
            end_idx = len(content)
            
        return content[start_idx:end_idx].strip()

    with open(passed_proto, 'r', encoding='utf-8') as f:
        passed_content = f.read()
        
    src_msg_name = f"E{src_evt.capitalize()}Bucket"
    tgt_msg_name = f"E{tgt_evt.capitalize()}Bucket"
    
    msg_block = extract_event_block(passed_content, src_evt)
    if not msg_block:
        log(f"Message {src_msg_name} not found in passed_events.proto. (Is it already in storagedata.proto?)")
        with open(storage_proto, 'r', encoding='utf-8') as f:
            msg_block = extract_event_block(f.read(), src_evt)
            if msg_block:
                log("Found source event in storagedata.proto instead. We will use that.")
            else:
                return False

    with open(storage_proto, 'r', encoding='utf-8') as f:
        storage_content = f.read()

    # If it already exists in target
    if extract_event_block(storage_content, tgt_evt):
        log(f"Target event message {tgt_msg_name} already exists in storagedata.proto. Skipping insertion.")
        return True

    if src_evt == tgt_evt:
        # Reopen Mode: Just append exactly as is
        new_msg_block = msg_block
    else:
        # Create New Event Mode: Rename inside the block
        new_msg_block = msg_block.replace(src_evt.lower(), tgt_evt.lower())
        new_msg_block = new_msg_block.replace(src_evt.capitalize(), tgt_evt.capitalize())
        new_msg_block = new_msg_block.replace(src_evt.upper(), tgt_evt.upper())
        
    # Create a backup before modifying
    shutil.copy2(storage_proto, storage_proto + ".bak")
    record_file_creation(root_path, step_index, storage_proto + ".bak") # Treat bak like a creation so it can be reverted
    
    # Append the block just before the end or at the end of the file
    storage_content += f"\n\n// Added by Event Automator: {tgt_evt}\n{new_msg_block}\n"
    
    with open(storage_proto, 'w', encoding='utf-8') as f:
        f.write(storage_content)
        
    # Also we should record the modification of storagedata.proto, but since we rely on the bak file for revert:
    log(f"Successfully injected proto message {tgt_msg_name} to storagedata.proto")
    return True

def _strip_lua_comment(line):
    """Drop a trailing `--` comment from one line of Lua, ignoring `--` inside strings."""
    in_str = None
    i = 0
    while i < len(line):
        c = line[i]
        if in_str:
            if c == '\\':
                i += 2
                continue
            if c == in_str:
                in_str = None
        elif c in '"\'':
            in_str = c
        elif line.startswith('--', i):
            return line[:i]
        i += 1
    return line


def _last_lua_code_index(text):
    """Index of the last character in text that is actual code (not whitespace or comment), or -1."""
    offset = len(text)
    for line in reversed(text.splitlines(keepends=True)):
        offset -= len(line)
        code = _strip_lua_comment(line.rstrip('\r\n')).rstrip()
        if code:
            return offset + len(code) - 1
    return -1


def find_lua_table_close(content, table_name):
    """
    Index of the `}` closing the table assigned by `<table_name> = {`, found by brace
    matching that skips strings and comments — so it doesn't care how the file is
    formatted. None if the assignment or its closing brace can't be found.
    """
    m = re.search(re.escape(table_name) + r'\s*=\s*\{', content)
    if not m:
        return None
    depth = 0
    in_str = None
    i = m.end() - 1
    while i < len(content):
        c = content[i]
        if in_str:
            if c == '\\':
                i += 2
                continue
            if c == in_str or c == '\n':  # an unterminated string can't span lines in Lua
                in_str = None
        elif content.startswith('--[[', i):
            end = content.find(']]', i + 4)
            if end == -1:
                return None
            i = end + 2
            continue
        elif content.startswith('--', i):
            nl = content.find('\n', i)
            if nl == -1:
                return None
            i = nl
            continue
        elif c in '"\'':
            in_str = c
        elif c == '{':
            depth += 1
        elif c == '}':
            depth -= 1
            if depth == 0:
                return i
        i += 1
    return None


def inject_local_service_mgr(proj_root, tgt_evt, root_path, step_index):
    mgr_path = os.path.join(proj_root, r"develop\client\Skipbo\Assets\Lua\Game\Module\LocalService\LocalServiceMgr.lua")
    if not os.path.exists(mgr_path):
        log(f"Error: LocalServiceMgr.lua not found at {mgr_path}")
        return False
        
    with open(mgr_path, 'r', encoding='utf-8') as f:
        content = f.read()

    tgt_cap = tgt_evt.capitalize()
    require_str_base = f'local E{tgt_cap}LocalService = require "Game/Module/LocalService/Event/E{tgt_cap}LocalService"'
    require_str = f'{require_str_base}\n'
    
    bucket_str_base = f'E{tgt_cap}Bucket = E{tgt_cap}LocalService.New(LocalServiceMgr, "E{tgt_cap}Bucket"),'
    bucket_str = f'        {bucket_str_base}\n'

    # 1. Check if actively registered (line starts directly with 'local' not '--')
    if re.search(r'^\s*' + re.escape(require_str_base), content, re.MULTILINE):
        log(f"LocalServiceMgr: {tgt_evt} already registered and active. Skipping.")
        return True

    # 2. Check if commented out (line starts with '--')
    commented_req_pattern = r'^\s*--\s*' + re.escape(require_str_base)
    if re.search(commented_req_pattern, content, re.MULTILINE):
        log(f"Found commented out registration for {tgt_evt}. Uncommenting...")
        content = re.sub(commented_req_pattern, require_str_base, content, flags=re.MULTILINE)
        
        # Pattern to find the commented bucket line (e.g. `--        E...`)
        commented_bucket_pattern = r'^\s*--\s*(?:.*?)' + re.escape(bucket_str_base)
        content = re.sub(commented_bucket_pattern, r'        ' + bucket_str_base, content, flags=re.MULTILINE)
        
        shutil.copy2(mgr_path, mgr_path + ".bak")
        record_file_creation(root_path, step_index, mgr_path + ".bak")
        with open(mgr_path, 'w', encoding='utf-8') as f:
            f.write(content)
        log(f"Successfully uncommented {tgt_evt} in LocalServiceMgr.lua")
        return True

    # 3. If neither active nor commented, insert it as a brand new event

    # Insert bucket inside ServiceDef. Located structurally, not by exact text: the file is
    # hand-edited between events, and r132944 joined the table's closing brace onto the next
    # statement ("}    LocalServiceMgr.LoginLocalService = ..."), which broke an exact-string anchor.
    close_idx = find_lua_table_close(content, "LocalServiceMgr.ServiceDef")
    if close_idx is None:
        log("Error: Could not find the LocalServiceMgr.ServiceDef = { ... } table in LocalServiceMgr.lua")
        return False
    line_start = content.rfind('\n', 0, close_idx) + 1
    if content[line_start:close_idx].strip() == "":
        # "}" starts its line: insert the new entry as the line above it.
        before, after = content[:line_start], content[line_start:]
    else:
        # "}" shares a line with the last entry: break the line before it.
        before, after = content[:close_idx].rstrip(' \t') + "\n", content[close_idx:]
    # The previous entry needs a trailing separator, or the new line is a Lua syntax error.
    last = _last_lua_code_index(before)
    if last >= 0 and before[last] not in ",;{":
        before = before[:last + 1] + "," + before[last + 1:]
    content = before + bucket_str + after

    # Insert require statement before the PBReg require (after the bucket edit, which is
    # located by searching forward from ServiceDef and so is unaffected by this insertion).
    pbreg = re.search(r'^local\s+PBReg\s*=.*$', content, re.MULTILINE)
    if not pbreg:
        log("Error: Could not find 'local PBReg =' in LocalServiceMgr.lua")
        return False
    content = content[:pbreg.start()] + require_str + content[pbreg.start():]

    shutil.copy2(mgr_path, mgr_path + ".bak")
    record_file_creation(root_path, step_index, mgr_path + ".bak")
    
    with open(mgr_path, 'w', encoding='utf-8') as f:
        f.write(content)
        
    log(f"Successfully registered {tgt_evt} in LocalServiceMgr.lua")
    return True

def clone_workspace3_to_design(proj_root, step_index):
    import subprocess
    
    design_dir = os.path.join(proj_root, r"design\DesignData")
    workspace3 = os.path.join(design_dir, "WorkSpcae3")
    workspace_design = os.path.join(design_dir, WORKSPACE_NAME)

    if not os.path.exists(workspace3):
        log(f"Error: WorkSpcae3 not found at {workspace3}")
        return False

    if os.path.normcase(os.path.abspath(workspace3)) == os.path.normcase(os.path.abspath(workspace_design)):
        log("Error: Active workspace is WorkSpcae3 itself. Refusing to delete/clone WorkSpcae3 onto itself — "
            "select WorkSpcae_Design, WorkSpcae_NE, or WorkSpcae_FC as the active workspace instead.")
        return False

    try:
        # 1. SVN Update design directory
        log("Running SVN Update on DesignData...")
        subprocess.run(["svn", "cleanup", "."], cwd=design_dir, shell=True, check=False, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        subprocess.run(["svn", "up", ".", "--non-interactive", "--trust-server-cert"], cwd=design_dir, shell=True, check=False, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        
        # 2. rmdir /s/q [CurrentWorkspace] (same as original bat)
        if os.path.exists(workspace_design):
            log(f"Removing old {WORKSPACE_NAME} directory...")
            subprocess.run(["cmd", "/c", "rmdir", "/s", "/q", workspace_design], check=False)
            
        # 3. xcopy WorkSpcae3 -> [CurrentWorkspace]
        log(f"Cloning WorkSpcae3 -> {WORKSPACE_NAME}...")
        shutil.copytree(workspace3, workspace_design)
        
        # Record creation for undo - FIX: Use provided root_path for central undo_log.json
        # NOTE: We need to pass the log_dir here, but it's not in the signature.
        # However, for Step 5 it might be called differently. Let's fix the call site or use a global.
        # For now, we'll assume the undo log is managed by the caller for this step,
        # or that the entire workspace_design can be removed if needed.
        # record_file_creation(design_dir, step_index, workspace_design) # Removed as per instruction
        
        # 4. SVN revert local utility bats in the new folder
        log("Reverting base bat files via SVN...")
        files_to_revert = [
            os.path.join(WORKSPACE_NAME, "PostConvert", "Lua", "lua_converter.lua"),
            os.path.join(WORKSPACE_NAME, "copy_data.bat"),
            os.path.join(WORKSPACE_NAME, "a_design_commit.bat"),
            os.path.join(WORKSPACE_NAME, "a_design_commit_with_convert.bat")
        ]
        for f in files_to_revert:
            subprocess.run(["svn", "revert", f], cwd=design_dir, shell=True, check=False, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            
        log("Successfully Cloned WorkSpace3 configuration to Design.")
        return True
    except Exception as e:
        log(f"Exception during Clone_2_Design: {e}")
        traceback.print_exc()
        return False

def detect_frozen_formula_risks(src_path, src_evt, wb_data_only):
    """
    Find formula cells in the SOURCE workbook whose cached/frozen result does not
    contain src_evt as text. clone_event_excel loads the target with data_only=True
    (required so coder_convert reads plain values, not formulas), which freezes every
    formula to whatever it last computed in the SOURCE file. The blanket src->tgt
    string-replace that follows only "fixes" a frozen cell by coincidence, when its
    frozen text happens to contain the source event name. Anything that doesn't is a
    silent risk: it ships the source's stale value with nothing to catch it.

    wb_data_only is the already-open, data_only=True copy of the target (pre-replace)
    — same cell coordinates as the source, since it's a fresh, unmodified copy at this
    point in clone_event_excel. Returns a list of (sheet_name, coordinate, formula, cached_value).
    """
    import openpyxl

    risky = []
    try:
        wb_formulas = openpyxl.load_workbook(src_path, data_only=False)
    except Exception as e:
        log(f"Warning: could not re-open {os.path.basename(src_path)} to check for formula risks: {e}")
        return risky

    src_variants = [src_evt.lower(), src_evt.capitalize(), src_evt.upper()]

    for ws_name in wb_formulas.sheetnames:
        if ws_name not in wb_data_only.sheetnames:
            continue
        ws_f = wb_formulas[ws_name]
        ws_v = wb_data_only[ws_name]
        for row in ws_f.iter_rows():
            for cell in row:
                v = cell.value
                if isinstance(v, str) and v.startswith('='):
                    cached = ws_v.cell(cell.row, cell.column).value
                    cached_str = str(cached) if cached is not None else ""
                    if not any(variant in cached_str for variant in src_variants):
                        risky.append((ws_name, cell.coordinate, v, cached_str))

    wb_formulas.close()
    return risky


def clone_event_excel(proj_root, src_evt, tgt_evt, root_path, step_index):
    """Step 6: Clone Event Excel configuration tables."""
    import openpyxl
    
    design_dir = os.path.join(proj_root, "design", "DesignData", WORKSPACE_NAME, "design")
    backup_dir = os.path.join(proj_root, os.path.join("design", "DesignData", WORKSPACE_NAME, "backup_xlsx"))
    
    src_xlsx = src_evt.lower() + ".xlsx"
    tgt_xlsx = tgt_evt.lower() + ".xlsx"
    
    tgt_path = os.path.join(design_dir, tgt_xlsx)
    
    # Check if target already exists in design
    if os.path.exists(tgt_path):
        if src_evt == tgt_evt:
            log(f"Reopen: {tgt_xlsx} already exists in design directory. Skipping.")
            return True
        else:
            log(f"New Event: {tgt_xlsx} already exists in design directory. Skipping.")
            return True
    
    # Find the source file
    src_path = os.path.join(design_dir, src_xlsx)
    if not os.path.exists(src_path):
        # Try backup_xlsx
        src_path = os.path.join(backup_dir, src_xlsx)
        if not os.path.exists(src_path):
            log(f"Error: Source Excel {src_xlsx} not found in design or backup_xlsx directories.")
            return False
        log(f"Found source Excel in backup_xlsx: {src_xlsx}")
    else:
        log(f"Found source Excel in design: {src_xlsx}")
    
    try:
        if src_evt == tgt_evt:
            # ---- REOPEN MODE ----
            # Just copy the file from backup to design if it was missing
            if os.path.dirname(src_path) != design_dir:
                shutil.copy2(src_path, tgt_path)
                record_file_creation(root_path, step_index, tgt_path)
                log(f"Reopen: Copied {src_xlsx} from backup_xlsx to design.")
            else:
                log(f"Reopen: {src_xlsx} already in design directory.")
            return True
        else:
            # ---- NEW EVENT MODE ----
            # Copy the file with the new name
            shutil.copy2(src_path, tgt_path)
            record_file_creation(root_path, step_index, tgt_path)
            log(f"Copied {src_xlsx} -> {tgt_xlsx}")
            
            # Open with openpyxl and replace content
            log(f"Replacing event references inside {tgt_xlsx}...")
            # Use data_only=True to ensure formulas (especially for IDs) are converted to static values
            # This is critical for coder_convert to read the IDs correctly.
            wb = openpyxl.load_workbook(tgt_path, data_only=True)

            # data_only=True freezes every formula to its cached SOURCE-computed value.
            # The string-replace pass below only "fixes" a frozen formula by coincidence,
            # when its cached result happens to contain the source event name as text.
            # Flag anything that doesn't — those ship the source's stale value silently.
            formula_risks = detect_frozen_formula_risks(src_path, src_evt, wb)
            if formula_risks:
                log(f"WARNING: {len(formula_risks)} formula cell(s) in {tgt_xlsx} were frozen to a value "
                    f"that does not contain '{src_evt}' — the rename below can't fix these, check manually:")
                for sheet_name, coord, formula, cached_val in formula_risks:
                    log(f"  [{sheet_name}!{coord}] formula={formula!r} frozen_value={cached_val!r}")

            replacements = [
                (src_evt.lower(), tgt_evt.lower()),
                (src_evt.capitalize(), tgt_evt.capitalize()),
                (src_evt.upper(), tgt_evt.upper()),
            ]
            
            cells_changed = 0
            sheets_renamed = 0
            
            for ws in wb.worksheets:
                # Rename sheet if it contains the source event name
                for old, new in replacements:
                    if old in ws.title:
                        ws.title = ws.title.replace(old, new)
                        sheets_renamed += 1
                
                # Replace cell values
                for row in ws.iter_rows():
                    for cell in row:
                        if cell.value and isinstance(cell.value, str):
                            original = cell.value
                            for old, new in replacements:
                                cell.value = cell.value.replace(old, new)
                            if cell.value != original:
                                cells_changed += 1
            
            wb.save(tgt_path)
            wb.close()
            
            log(f"Successfully cloned Excel: {tgt_xlsx} ({cells_changed} cells updated, {sheets_renamed} sheets renamed)")
            return True
    except Exception as e:
        log(f"Exception during Excel clone: {e}")
        traceback.print_exc()
        return False

def clone_event_descriptors(proj_root, src_evt, tgt_evt, root_path, step_index):
    """Step 7: Copy event descriptor JSON(s) to app_client directory."""
    import glob
    
    app_client_dir = os.path.join(proj_root, os.path.join("design", "DesignData", WORKSPACE_NAME, "descriptor", "app_client"))
    backup_dir = os.path.join(proj_root, os.path.join("design", "DesignData", WORKSPACE_NAME, "descriptor", "backup"))
    
    src_lower = src_evt.lower()
    tgt_lower = tgt_evt.lower()
    
    # 1. Find all source descriptor JSONs via glob
    src_pattern_app = os.path.join(app_client_dir, f"*{src_lower}*_convert.json")
    src_pattern_bak = os.path.join(backup_dir, f"*{src_lower}*_convert.json")
    
    src_files = glob.glob(src_pattern_app)
    source_location = "app_client"
    
    if not src_files:
        src_files = glob.glob(src_pattern_bak)
        source_location = "backup"
    
    if not src_files:
        log(f"Error: No descriptor JSONs matching '*{src_lower}*_convert.json' found in app_client or backup.")
        return False
    
    log(f"Found {len(src_files)} descriptor JSON(s) in {source_location}: {[os.path.basename(f) for f in src_files]}")
    
    try:
        for src_path in src_files:
            src_basename = os.path.basename(src_path)
            
            if src_evt == tgt_evt:
                # ---- REOPEN MODE ----
                tgt_basename = src_basename
                tgt_path = os.path.join(app_client_dir, tgt_basename)
                if os.path.exists(tgt_path):
                    log(f"Reopen: {tgt_basename} already in app_client. Skipping.")
                    continue
                shutil.copy2(src_path, tgt_path)
                record_file_creation(root_path, step_index, tgt_path)
                log(f"Reopen: Copied {src_basename} from {source_location} to app_client.")
            else:
                # ---- NEW EVENT MODE ----
                tgt_basename = src_basename.replace(src_lower, tgt_lower)
                tgt_path = os.path.join(app_client_dir, tgt_basename)
                
                if os.path.exists(tgt_path):
                    log(f"New Event: {tgt_basename} already exists. Skipping.")
                    continue
                
                # Read source, replace content, write target
                with open(src_path, 'r', encoding='utf-8') as f:
                    content = f.read()
                
                content = content.replace(src_evt.lower(), tgt_evt.lower())
                content = content.replace(src_evt.capitalize(), tgt_evt.capitalize())
                content = content.replace(src_evt.upper(), tgt_evt.upper())
                
                with open(tgt_path, 'w', encoding='utf-8') as f:
                    f.write(content)
                
                record_file_creation(root_path, step_index, tgt_path)
                log(f"Created {tgt_basename} with content renamed from {src_basename}")
        
        log("Successfully cloned all event descriptor JSON(s).")
        return True
    except Exception as e:
        log(f"Exception during descriptor clone: {e}")
        traceback.print_exc()
        return False

def cleanup_orphaned_convert_entries(design_base, root_path, step_index):
    """
    Audit convert.json's const_event.json ExcelSheets list against the design folder.
    Any event registered there whose <name>.xlsx no longer exists in design/ is stale
    (the Excel was deleted/moved but the config references were never cleaned up) —
    remove that event from both convert.json and convert_layout.lua so coder_convert
    doesn't choke on a reference to a file that isn't there.
    """
    import json
    import re

    convert_json_path = os.path.join(design_base, r"descriptor\app_client\convert.json")
    convert_layout_path = os.path.join(design_base, r"PostConvert\Lua\convert_layout.lua")
    design_dir = os.path.join(design_base, "design")

    if not os.path.exists(convert_json_path) or not os.path.exists(convert_layout_path):
        return True  # Let the caller's own existence checks handle this; nothing to audit yet.

    try:
        with open(convert_json_path, 'r', encoding='utf-8') as f:
            data = json.load(f)

        const_event_dir = None
        for director in data.get("Directors", []):
            if director.get("FileName") == "const_event.json":
                const_event_dir = director
                break

        if not const_event_dir:
            return True

        sheets = const_event_dir.get("ExcelSheets", [])
        orphaned = [
            s.get("ExcelName") for s in sheets
            if s.get("ExcelName") and not os.path.exists(os.path.join(design_dir, f"{s.get('ExcelName')}.xlsx"))
        ]

        if not orphaned:
            log("convert.json/convert_layout.lua audit: no orphaned event references found.")
            return True

        log(f"convert.json/convert_layout.lua audit: found {len(orphaned)} orphaned event(s) with no Excel file in design/: {orphaned}")

        # --- Remove from convert.json ---
        backup_json = convert_json_path + ".bak"
        shutil.copy2(convert_json_path, backup_json)
        record_file_creation(root_path, step_index, backup_json)

        const_event_dir["ExcelSheets"] = [s for s in sheets if s.get("ExcelName") not in orphaned]
        with open(convert_json_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
        log(f"convert.json: removed {len(orphaned)} orphaned entr{'y' if len(orphaned) == 1 else 'ies'}.")

        # --- Remove matching lines from convert_layout.lua ---
        backup_lua = convert_layout_path + ".bak"
        shutil.copy2(convert_layout_path, backup_lua)
        record_file_creation(root_path, step_index, backup_lua)

        with open(convert_layout_path, 'r', encoding='utf-8') as f:
            lines = f.read().split('\n')

        orphan_pattern = re.compile(r'json\s*=\s*"(?:' + '|'.join(re.escape(o) for o in orphaned) + r')_')
        kept_lines = [ln for ln in lines if not orphan_pattern.search(ln)]
        removed_count = len(lines) - len(kept_lines)

        with open(convert_layout_path, 'w', encoding='utf-8') as f:
            f.write('\n'.join(kept_lines))
        log(f"convert_layout.lua: removed {removed_count} orphaned line(s) for {orphaned}.")

        return True
    except Exception as e:
        log(f"Exception during orphaned convert entry cleanup: {e}")
        traceback.print_exc()
        return False

def parse_convert_layout_suffix_map(content, restrict_prefix=None):
    """
    Scan convert_layout.lua-style content and build:
      section_end_lines: section_name -> last line index containing an entry
      suffix_to_info:    file_suffix   -> (section_name, merge_suffix, sample_line)

    If restrict_prefix is given (e.g. "blocks2508"), only entries whose merge_as
    value starts with "<restrict_prefix>." are considered — used when scanning a
    historical snapshot that may contain many other events besides the one we
    actually care about.
    """
    import re as _re

    lines = content.split('\n')
    current_section = None
    section_end_lines = {}
    suffix_to_info = {}

    for i, line in enumerate(lines):
        stripped = line.strip()
        section_match = _re.match(r'(\w+)\s*=\s*\{', stripped)
        if section_match:
            current_section = section_match.group(1)

        entry_match = _re.match(r'.*json\s*=\s*"([^"]+)".*merge_as\s*=\s*"([^"]+)"', stripped)
        if entry_match and current_section and not stripped.startswith('--'):
            json_name = entry_match.group(1)
            merge_as = entry_match.group(2)

            dot_idx = merge_as.find('.')
            if dot_idx > 0:
                evt_prefix = merge_as[:dot_idx]
                merge_suffix = merge_as[dot_idx + 1:]

                if restrict_prefix and evt_prefix.lower() != restrict_prefix.lower():
                    continue

                if json_name.startswith(evt_prefix + '_'):
                    file_suffix = json_name[len(evt_prefix) + 1:]
                else:
                    file_suffix = json_name

                # Section end lines should reflect entries actually kept (matching
                # restrict_prefix when set), not every event's last line in the file.
                section_end_lines[current_section] = i

                if file_suffix not in suffix_to_info:
                    suffix_to_info[file_suffix] = (current_section, merge_suffix, line)

    return section_end_lines, suffix_to_info


def get_svn_log_revisions(file_path, limit=60):
    """Return revision numbers (newest first) that touched file_path, via `svn log -q`."""
    import subprocess
    import re as _re

    try:
        result = subprocess.run(
            ["svn", "log", "-q", "-l", str(limit), os.path.basename(file_path)],
            cwd=os.path.dirname(file_path), shell=True, capture_output=True, text=True, timeout=60,
            stdin=subprocess.DEVNULL
        )
    except Exception as e:
        log(f"svn log failed to run: {e}")
        return []

    if result.returncode != 0:
        log(f"svn log error: {result.stderr.strip()}")
        return []

    return _re.findall(r'^r(\d+)\s', result.stdout, flags=_re.MULTILINE)


def svn_cat_revision(file_path, revision):
    """Return file_path's content at a specific SVN revision, or None on failure."""
    import subprocess

    try:
        result = subprocess.run(
            ["svn", "cat", "-r", str(revision), os.path.basename(file_path)],
            cwd=os.path.dirname(file_path), shell=True, capture_output=True, text=True,
            timeout=30, encoding='utf-8', errors='ignore', stdin=subprocess.DEVNULL
        )
    except Exception as e:
        log(f"svn cat -r {revision} failed to run: {e}")
        return None

    if result.returncode != 0:
        return None
    return result.stdout


def svn_cat_revision_bytes(file_path, revision):
    """Same as svn_cat_revision but returns raw bytes — required for binary files like
    .xlsx (text=True mode would corrupt the zip container)."""
    import subprocess

    try:
        result = subprocess.run(
            ["svn", "cat", "-r", str(revision), os.path.basename(file_path)],
            cwd=os.path.dirname(file_path), shell=True, capture_output=True,
            timeout=60, stdin=subprocess.DEVNULL
        )
    except Exception as e:
        log(f"svn cat -r {revision} (binary) failed to run: {e}")
        return None

    if result.returncode != 0:
        return None
    return result.stdout


def get_svn_revisions_in_date_range(file_path, center_date_str, days_before=45, days_after=120):
    """
    Return revision numbers (newest first) that touched file_path within a date
    window bracketing center_date_str (the source event's own start_time).

    This is the whole point: a source event from 11 months ago can sit 400+
    revisions back on a file that's touched roughly every 2 weeks. A blind
    HEAD-relative scan either has to check hundreds of revisions (slow — each
    `svn cat` is a network round trip) or caps out and misses it entirely.
    Bracketing by the event's own known active dates instead finds it in a
    couple dozen revisions regardless of how long ago it ran.
    """
    import subprocess
    import re as _re
    from datetime import datetime, timedelta

    try:
        center = datetime.strptime(center_date_str.strip()[:10], "%Y-%m-%d")
    except Exception:
        return None  # Unparseable/missing hint — caller falls back to a revision-count scan.

    range_start = (center - timedelta(days=days_before)).strftime("%Y-%m-%d")
    range_end = (center + timedelta(days=days_after)).strftime("%Y-%m-%d")

    try:
        result = subprocess.run(
            ["svn", "log", "-q", "-r", f"{{{range_start}}}:{{{range_end}}}", os.path.basename(file_path)],
            cwd=os.path.dirname(file_path), shell=True, capture_output=True, text=True, timeout=60,
            stdin=subprocess.DEVNULL
        )
    except Exception as e:
        log(f"svn log (date range) failed to run: {e}")
        return None

    if result.returncode != 0:
        log(f"svn log (date range) error: {result.stderr.strip()}")
        return None

    revs = _re.findall(r'^r(\d+)\s', result.stdout, flags=_re.MULTILINE)
    return list(reversed(revs))  # svn log over an ascending date range returns oldest -> newest


def find_source_event_suffix_map_from_svn(convert_layout_path, src_lower, hint_date=None, limit=60):
    """
    Look through convert_layout.lua's SVN history for the most recent revision
    that actually contains entries for src_lower (e.g. "blocks2509").
    Returns (suffix_to_info, revision) for that revision, or (None, None) if SVN
    isn't usable or no matching revision was found.

    When hint_date (the source event's own start_time) is available, the search
    is bounded to a date window around it — see get_svn_revisions_in_date_range.
    Falls back to a bare HEAD-relative revision-count scan only if no hint_date
    was given or it couldn't be parsed; that fallback is a best-effort last
    resort, not reliable for events more than a few months old.
    """
    revisions = None
    used_date_bracket = False

    if hint_date:
        revisions = get_svn_revisions_in_date_range(convert_layout_path, hint_date)
        if revisions:
            used_date_bracket = True
            log(f"SVN history lookup: date-bracketed search around {hint_date.strip()[:10]} "
                f"found {len(revisions)} revision(s) to check (instead of a blind HEAD scan).")

    if not revisions:
        if hint_date:
            log("SVN history lookup: date-bracketed search found nothing usable — falling back to a HEAD-relative scan.")
        revisions = get_svn_log_revisions(convert_layout_path, limit=limit)

    if not revisions:
        log(f"SVN history lookup: no revision history available for {os.path.basename(convert_layout_path)}.")
        return None, None

    if not used_date_bracket:
        log(f"SVN history lookup: scanning the last {len(revisions)} revision(s) of "
            f"{os.path.basename(convert_layout_path)} for '{src_lower}_' entries (no reliable date hint available)...")

    for rev in revisions:
        content = svn_cat_revision(convert_layout_path, rev)
        if content is None:
            continue
        if f'"{src_lower}_' not in content.lower():
            continue

        _, suffix_map = parse_convert_layout_suffix_map(content, restrict_prefix=src_lower)
        if suffix_map:
            log(f"SVN history lookup: found {src_lower} entries at r{rev} ({len(suffix_map)} suffix pattern(s)).")
            return suffix_map, rev

    log(f"SVN history lookup: no revision within the {len(revisions)} checked contained active '{src_lower}_' entries.")
    return None, None


# ---------------- Local cache for convert_layout.lua source-event lookups ----------------
#
# The expensive part of the above is walking SVN history. But the common case (reopening
# the previous latest event) almost never needs it: that event's entries are usually still
# live in the local working copy right now — reading the file that's already on disk costs
# nothing. The cache only matters for events whose entries have already rolled off the live
# file (closed + cleaned up), and even then only needs to pay the SVN cost once per event.

def convert_layout_cache_path(log_dir):
    return os.path.join(log_dir, f"convert_layout_cache_{WORKSPACE_NAME}.json")


def load_convert_layout_cache(log_dir):
    p = convert_layout_cache_path(log_dir)
    if os.path.exists(p):
        try:
            with open(p, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            return {}
    return {}


def _json_safe_default(obj):
    """json.dump default= handler — cached row data can contain openpyxl's native
    datetime/date cell values, which aren't JSON-serializable on their own."""
    from datetime import datetime, date
    if isinstance(obj, (datetime, date)):
        return obj.isoformat()
    return str(obj)


def save_convert_layout_cache(log_dir, cache):
    p = convert_layout_cache_path(log_dir)
    with open(p, 'w', encoding='utf-8') as f:
        json.dump(cache, f, indent=2, ensure_ascii=False, default=_json_safe_default)


def _cache_timestamp():
    from datetime import datetime
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def get_suffix_map_for_source_event(convert_layout_path, src_lower, hint_date, log_dir):
    """
    Resolve the section/merge_as suffix map for src_lower with the minimum SVN cost:
      1. If src_lower's entries are still live in the current local file, use them
         directly (zero network cost) and refresh the cache with them — this is the
         common "reopen the latest event" path.
      2. Otherwise, if previously cached, use the cached result (zero cost).
      3. Otherwise, run the full SVN-history search (expensive) and cache the result
         so the next reopen of this same event is instant.
    Returns (suffix_map, source_label) — source_label is for logging/transparency,
    never for logic branching.
    """
    with open(convert_layout_path, 'r', encoding='utf-8') as f:
        live_content = f.read()

    _, live_map = parse_convert_layout_suffix_map(live_content, restrict_prefix=src_lower)
    if live_map:
        cache = load_convert_layout_cache(log_dir)
        cache[src_lower] = {"suffix_map": live_map, "source": "live", "cached_at": _cache_timestamp()}
        save_convert_layout_cache(log_dir, cache)
        return live_map, f"current live file ({len(live_map)} pattern(s))"

    cache = load_convert_layout_cache(log_dir)
    cached_entry = cache.get(src_lower)
    if cached_entry and cached_entry.get("suffix_map"):
        return cached_entry["suffix_map"], f"local cache (from {cached_entry.get('source', '?')}, cached {cached_entry.get('cached_at', '?')})"

    suffix_map, rev = find_source_event_suffix_map_from_svn(convert_layout_path, src_lower, hint_date=hint_date)
    if suffix_map:
        cache[src_lower] = {"suffix_map": suffix_map, "source": f"svn r{rev}", "cached_at": _cache_timestamp()}
        save_convert_layout_cache(log_dir, cache)
        return suffix_map, f"SVN history (r{rev}, freshly fetched and cached)"

    return None, None


def refresh_convert_layout_cache(proj_root, log_dir, event_names):
    """
    Manual 'Update from SVN' action. Re-verifies every event name given, plus
    whatever's already cached, against the current live file first (free) and,
    if not live, a fresh SVN-history search — updating the cache only where the
    resolved data actually differs from what's cached. Never blindly overwrites.
    """
    convert_layout_path = os.path.join(proj_root, "design", "DesignData", WORKSPACE_NAME, "PostConvert", "Lua", "convert_layout.lua")
    if not os.path.exists(convert_layout_path):
        log(f"Error: convert_layout.lua not found at {convert_layout_path}")
        return False

    with open(convert_layout_path, 'r', encoding='utf-8') as f:
        live_content = f.read()

    cache = load_convert_layout_cache(log_dir)
    all_names = sorted(set(n.lower() for n in event_names) | set(cache.keys()))

    changed, unchanged, missing = [], [], []

    for evt_lower in all_names:
        _, live_map = parse_convert_layout_suffix_map(live_content, restrict_prefix=evt_lower)
        if live_map:
            prior = cache.get(evt_lower, {}).get("suffix_map")
            if prior != live_map:
                cache[evt_lower] = {"suffix_map": live_map, "source": "live", "cached_at": _cache_timestamp()}
                changed.append(evt_lower)
            else:
                unchanged.append(evt_lower)
            continue

        fresh_map, rev = find_source_event_suffix_map_from_svn(convert_layout_path, evt_lower)
        if fresh_map:
            prior = cache.get(evt_lower, {}).get("suffix_map")
            if prior != fresh_map:
                cache[evt_lower] = {"suffix_map": fresh_map, "source": f"svn r{rev}", "cached_at": _cache_timestamp()}
                changed.append(evt_lower)
            else:
                unchanged.append(evt_lower)
        else:
            missing.append(evt_lower)
            log(f"Refresh: no data found for {evt_lower} — not live, and nothing in SVN history.")

    save_convert_layout_cache(log_dir, cache)
    log(f"Cache refresh complete: {len(changed)} updated {changed}, {len(unchanged)} unchanged, {len(missing)} not found {missing}.")
    return True


# ================================================================================
# icon.xlsx / item.xlsx source-event data: same live -> cache -> SVN-history
# pattern as convert_layout.lua, extended to three more data shapes:
#   - icon_item's "活动物品" group (variable row count per event)
#   - icon_event (checked via icon_备份 first, since that's a maintained in-repo
#     archive — free, no SVN needed — before falling back to SVN history)
#   - item.xlsx's 物品小/中/大ICON columns (7,8,9), which are only ever populated
#     for events that were registered in convert.json/convert_layout.lua at the
#     time — everything else has them blank and needs backfilling from the source.
# ================================================================================

def load_named_cache(log_dir, cache_name):
    p = os.path.join(log_dir, f"{cache_name}_{WORKSPACE_NAME}.json")
    if os.path.exists(p):
        try:
            with open(p, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            return {}
    return {}


def save_named_cache(log_dir, cache_name, cache):
    p = os.path.join(log_dir, f"{cache_name}_{WORKSPACE_NAME}.json")
    with open(p, 'w', encoding='utf-8') as f:
        json.dump(cache, f, indent=2, ensure_ascii=False, default=_json_safe_default)


def _strip_event_name(value, evt_name):
    """Normalize an alias/name by removing the event name in any case variant,
    leaving a generic suffix comparable across events (e.g. 'Merge2606能量' -> '能量')."""
    if not value or not evt_name:
        return str(value or "").strip()
    s = str(value)
    for variant in [evt_name.upper(), evt_name.lower(), evt_name[0].upper() + evt_name[1:]]:
        if variant:
            s = s.replace(variant, "")
    return s.strip()


def find_icon_item_group_bounds(ws, group_marker="活动物品"):
    """Locate the icon_item sheet's '// ...活动物品...' divider and the next divider
    after it (e.g. '// BP相关...'). Returns (start_row, end_row_exclusive), or
    (None, None) if the group marker isn't found."""
    start = None
    end = ws.max_row + 1
    for r in range(1, ws.max_row + 1):
        v = str(ws.cell(r, 1).value or "").strip()
        if v.startswith("//"):
            if group_marker in v:
                start = r + 1
            elif start is not None:
                end = r
                break
    return start, end


def extract_icon_item_group_rows(wb, evt_lower, group_marker="活动物品", sheet_name="icon_item", with_style=False):
    """Return raw row value-lists within a divider-bounded group whose col1 mentions
    evt_lower. Generalized beyond icon_item — also used for localize_event_fixed's
    '活动名' group, which follows the identical divider-bounded-accumulation shape.
    with_style=True (used by localize_event_fixed, which has real per-row highlight
    formatting) returns each row as a list of (value, style_dict) pairs instead of
    plain values, so formatting survives a round-trip through the JSON cache."""
    if sheet_name not in wb.sheetnames:
        return None
    ws = wb[sheet_name]
    start, end = find_icon_item_group_bounds(ws, group_marker)
    if start is None:
        return None
    matched = []
    for r in range(start, end):
        v = str(ws.cell(r, 1).value or "")
        if evt_lower in v.lower():
            if with_style:
                matched.append([(ws.cell(r, c).value, cell_style_to_dict(ws.cell(r, c))) for c in range(1, ws.max_column + 1)])
            else:
                matched.append([ws.cell(r, c).value for c in range(1, ws.max_column + 1)])
    return matched if matched else None


def event_ever_registered_in_convert_json(icon_or_item_path, evt_lower, hint_date):
    """
    Cheap pre-check before the expensive icon.xlsx/item.xlsx binary SVN search:
    has evt_lower EVER appeared in convert.json (current content, or its history)?

    Every closed event checked so far with no convert.json history also had zero
    icon_item/icon_event/item-icon data — the two are the same underlying fact
    (this pipeline only started populating those fields once events started being
    registered there). convert.json is small text, so `svn cat` per revision costs
    ~6s instead of icon.xlsx's ~14min for a full binary scan. `svn log --diff`
    looks cheaper on paper but the server-side diff computation timed out even on
    this small file — plain `svn cat` per revision is what's actually fast here.

    Returns True (proceed with the expensive search) if genuinely uncertain —
    this is a time-saving skip, never a correctness gate, so it fails open.
    """
    import subprocess

    design_base = os.path.dirname(os.path.dirname(icon_or_item_path))
    convert_json_path = os.path.join(design_base, r"descriptor\app_client\convert.json")
    if not os.path.exists(convert_json_path):
        return True

    try:
        with open(convert_json_path, 'r', encoding='utf-8') as f:
            if evt_lower in f.read().lower():
                return True
    except Exception:
        return True

    revisions = get_svn_revisions_in_date_range(convert_json_path, hint_date) if hint_date else None
    if not revisions:
        revisions = get_svn_log_revisions(convert_json_path, limit=30)
    if not revisions:
        return True  # Can't check — don't block a legitimate search.

    for rev in revisions:
        try:
            result = subprocess.run(
                ["svn", "cat", "-r", str(rev), os.path.basename(convert_json_path)],
                cwd=os.path.dirname(convert_json_path), shell=True, capture_output=True, text=True,
                timeout=30, encoding='utf-8', errors='ignore', stdin=subprocess.DEVNULL
            )
        except Exception:
            return True
        if result.returncode == 0 and evt_lower in result.stdout.lower():
            return True

    return False


def get_icon_item_rows_for_source(icon_path, src_lower, hint_date, log_dir):
    """
    (rows, source_label) for src_lower's icon_item '活动物品' group rows.
    Tier 1: live file (free). Tier 2: local cache (free). Tier 3: SVN history
    (date-bracketed by hint_date when available).
    """
    import openpyxl as _oxl
    import io as _io

    wb_live = _oxl.load_workbook(icon_path, data_only=True)
    live_rows = extract_icon_item_group_rows(wb_live, src_lower)
    if live_rows:
        cache = load_named_cache(log_dir, "icon_item_cache")
        cache[src_lower] = {"rows": live_rows, "source": "live", "cached_at": _cache_timestamp()}
        save_named_cache(log_dir, "icon_item_cache", cache)
        return live_rows, f"current live file ({len(live_rows)} row(s))"

    cache = load_named_cache(log_dir, "icon_item_cache")
    cached_entry = cache.get(src_lower)
    if cached_entry and cached_entry.get("rows"):
        return cached_entry["rows"], f"local cache (from {cached_entry.get('source', '?')}, cached {cached_entry.get('cached_at', '?')})"

    if not event_ever_registered_in_convert_json(icon_path, src_lower, hint_date):
        log(f"icon_item: {src_lower} never appears in convert.json (current or history) — skipping the expensive "
            f"icon.xlsx SVN search, since this pipeline never populates 活动物品 rows for events outside that set.")
        return None, None

    revisions = get_svn_revisions_in_date_range(icon_path, hint_date) if hint_date else None
    if not revisions:
        revisions = get_svn_log_revisions(icon_path, limit=60)
    if not revisions:
        return None, None

    for rev in revisions:
        content_bytes = svn_cat_revision_bytes(icon_path, rev)
        if content_bytes is None:
            continue
        try:
            wb_hist = _oxl.load_workbook(_io.BytesIO(content_bytes), data_only=True)
        except Exception:
            continue
        rows = extract_icon_item_group_rows(wb_hist, src_lower)
        if rows:
            cache[src_lower] = {"rows": rows, "source": f"svn r{rev}", "cached_at": _cache_timestamp()}
            save_named_cache(log_dir, "icon_item_cache", cache)
            return rows, f"SVN history (r{rev}, freshly fetched and cached)"

    return None, None


def get_icon_event_rows_for_source(icon_path, src_lower, hint_date, log_dir):
    """
    (rows, source_label) for src_lower's icon_event rows.
    Tier 1a: icon_备份 (a maintained in-repo archive — free). Tier 1b: icon_event
    itself, if the source is still currently active. Tier 2: local cache.
    Tier 3: SVN history of the icon_event sheet.
    """
    import openpyxl as _oxl
    import io as _io

    wb_live = _oxl.load_workbook(icon_path, data_only=True)

    def _rows_from_sheet(ws, match_all_cols):
        rows = []
        for r in range(2, ws.max_row + 1):
            c1 = str(ws.cell(r, 1).value or "").lower()
            if match_all_cols:
                c2 = str(ws.cell(r, 2).value or "").lower()
                c3 = str(ws.cell(r, 3).value or "").lower()
                hit = src_lower in c1 or src_lower in c2 or src_lower in c3
            else:
                hit = src_lower in c1
            if hit:
                rows.append([ws.cell(r, c).value for c in range(1, ws.max_column + 1)])
        return rows

    if "icon_备份" in wb_live.sheetnames:
        rows = _rows_from_sheet(wb_live["icon_备份"], match_all_cols=True)
        if rows:
            cache = load_named_cache(log_dir, "icon_event_cache")
            cache[src_lower] = {"rows": rows, "source": "icon_备份 (live)", "cached_at": _cache_timestamp()}
            save_named_cache(log_dir, "icon_event_cache", cache)
            return rows, f"icon_备份 backup sheet ({len(rows)} row(s))"

    if "icon_event" in wb_live.sheetnames:
        rows = _rows_from_sheet(wb_live["icon_event"], match_all_cols=False)
        if rows:
            cache = load_named_cache(log_dir, "icon_event_cache")
            cache[src_lower] = {"rows": rows, "source": "icon_event (live)", "cached_at": _cache_timestamp()}
            save_named_cache(log_dir, "icon_event_cache", cache)
            return rows, f"icon_event active sheet ({len(rows)} row(s))"

    cache = load_named_cache(log_dir, "icon_event_cache")
    cached_entry = cache.get(src_lower)
    if cached_entry and cached_entry.get("rows"):
        return cached_entry["rows"], f"local cache (from {cached_entry.get('source', '?')}, cached {cached_entry.get('cached_at', '?')})"

    # Weaker signal here than for icon_item/item_icon — icon_备份 already independently
    # covers most historical events regardless of convert.json status, so reaching this
    # point at all is rarer. Still worth skipping when we're confident, since the cost
    # is the same expensive binary scan either way.
    if not event_ever_registered_in_convert_json(icon_path, src_lower, hint_date):
        log(f"icon_event: {src_lower} never appears in convert.json (current or history) and wasn't in icon_备份 "
            f"or the live icon_event sheet — skipping the expensive SVN search.")
        return None, None

    revisions = get_svn_revisions_in_date_range(icon_path, hint_date) if hint_date else None
    if not revisions:
        revisions = get_svn_log_revisions(icon_path, limit=60)
    if not revisions:
        return None, None

    for rev in revisions:
        content_bytes = svn_cat_revision_bytes(icon_path, rev)
        if content_bytes is None:
            continue
        try:
            wb_hist = _oxl.load_workbook(_io.BytesIO(content_bytes), data_only=True)
        except Exception:
            continue
        if "icon_event" not in wb_hist.sheetnames:
            continue
        rows = _rows_from_sheet(wb_hist["icon_event"], match_all_cols=False)
        if rows:
            cache[src_lower] = {"rows": rows, "source": f"svn r{rev} (icon_event)", "cached_at": _cache_timestamp()}
            save_named_cache(log_dir, "icon_event_cache", cache)
            return rows, f"SVN history (r{rev}, freshly fetched and cached)"

    return None, None


def get_item_icon_columns_for_source(item_path, src_evt, log_dir, hint_date):
    """
    {suffix: (icon7, icon8, icon9)} for src_evt's items, keyed by the alias with the
    event name stripped (so it's comparable to a differently-named target event).
    Tier 1: live file. Tier 2: cache. Tier 3: SVN history.
    """
    import openpyxl as _oxl
    import io as _io
    src_lower = src_evt.lower()

    def _extract(wb):
        if "item" not in wb.sheetnames:
            return None
        ws = wb["item"]
        result = {}
        for r in range(2, ws.max_row + 1):
            alias = str(ws.cell(r, 2).value or "")
            if src_lower in alias.lower():
                icon7, icon8, icon9 = ws.cell(r, 7).value, ws.cell(r, 8).value, ws.cell(r, 9).value
                if icon7 or icon8 or icon9:
                    result[_strip_event_name(alias, src_evt)] = (icon7, icon8, icon9)
        return result if result else None

    wb_live = _oxl.load_workbook(item_path, data_only=True)
    live_map = _extract(wb_live)
    if live_map:
        cache = load_named_cache(log_dir, "item_icon_cache")
        cache[src_lower] = {"map": live_map, "source": "live", "cached_at": _cache_timestamp()}
        save_named_cache(log_dir, "item_icon_cache", cache)
        return live_map, f"current live file ({len(live_map)} item(s))"

    cache = load_named_cache(log_dir, "item_icon_cache")
    cached_entry = cache.get(src_lower)
    if cached_entry and cached_entry.get("map"):
        return cached_entry["map"], f"local cache (from {cached_entry.get('source', '?')}, cached {cached_entry.get('cached_at', '?')})"

    if not event_ever_registered_in_convert_json(item_path, src_lower, hint_date):
        log(f"item_icon_backfill: {src_lower} never appears in convert.json (current or history) — skipping the "
            f"expensive item.xlsx SVN search, since icon columns are only ever populated for registered events.")
        return None, None

    revisions = get_svn_revisions_in_date_range(item_path, hint_date) if hint_date else None
    if not revisions:
        revisions = get_svn_log_revisions(item_path, limit=60)
    if not revisions:
        return None, None

    for rev in revisions:
        content_bytes = svn_cat_revision_bytes(item_path, rev)
        if content_bytes is None:
            continue
        try:
            wb_hist = _oxl.load_workbook(_io.BytesIO(content_bytes), data_only=True)
        except Exception:
            continue
        hist_map = _extract(wb_hist)
        if hist_map:
            cache[src_lower] = {"map": hist_map, "source": f"svn r{rev}", "cached_at": _cache_timestamp()}
            save_named_cache(log_dir, "item_icon_cache", cache)
            return hist_map, f"SVN history (r{rev}, freshly fetched and cached)"

    return None, None


def backfill_item_icon_columns(proj_root, src_evt, tgt_evt, log_dir):
    """
    Populate the target event's item rows' 物品小/中/大ICON columns (7,8,9) from the
    source event's own values (live/cache/SVN) — these columns are only ever populated
    for events that were registered in convert.json/convert_layout.lua at clone time,
    so a source that wasn't active back then leaves the target's clone blank otherwise.
    """
    import openpyxl
    import stat as _stat

    item_path = os.path.join(proj_root, "design", "DesignData", WORKSPACE_NAME, "design", "item.xlsx")
    if not os.path.exists(item_path):
        return True

    tgt_lower = tgt_evt.lower()
    src_lower = src_evt.lower()
    tgt_pascal = tgt_evt[0].upper() + tgt_evt[1:] if tgt_evt else tgt_evt
    src_pascal = src_evt[0].upper() + src_evt[1:] if src_evt else src_evt

    def swap_name(value):
        if value is None:
            return None
        s = str(value)
        s = s.replace(src_evt.upper(), tgt_evt.upper())
        s = s.replace(src_lower, tgt_lower)
        s = s.replace(src_pascal, tgt_pascal)
        s = s.replace(src_evt, tgt_evt)
        return s

    extra = load_extra_args(log_dir)
    src_start_time = extra.get("src_start_time", "")

    src_icon_map, src_label = get_item_icon_columns_for_source(item_path, src_evt, log_dir, src_start_time)
    if not src_icon_map:
        log(f"item_icon_backfill: no icon column data found anywhere for {src_evt} (live, cache, or SVN history).")
        return True

    os.chmod(item_path, _stat.S_IWRITE | _stat.S_IREAD)
    wb = load_workbook_frozen(item_path)
    if "item" not in wb.sheetnames:
        return True
    ws = wb["item"]

    filled = 0
    for r in range(2, ws.max_row + 1):
        alias = str(ws.cell(r, 2).value or "")
        if tgt_lower not in alias.lower():
            continue
        if ws.cell(r, 7).value or ws.cell(r, 8).value or ws.cell(r, 9).value:
            continue  # already populated — don't clobber real data
        suffix = _strip_event_name(alias, tgt_evt)
        if suffix in src_icon_map:
            icon7, icon8, icon9 = src_icon_map[suffix]
            if icon7: ws.cell(r, 7, swap_name(icon7))
            if icon8: ws.cell(r, 8, swap_name(icon8))
            if icon9: ws.cell(r, 9, swap_name(icon9))
            filled += 1

    if filled:
        wb.save(item_path)
        log(f"item_icon_backfill: Populated icon columns for {filled} item(s) in {tgt_evt}, sourced from {src_label}.")
    else:
        log(f"item_icon_backfill: No item rows needed backfilling for {tgt_evt}.")
    return True


def refresh_icon_and_item_caches(proj_root, log_dir, event_names):
    """Pre-populate icon_item / icon_event / item_icon caches for the given event
    names (e.g. each type's current latest event) — same manual 'Update from SVN'
    role as refresh_convert_layout_cache, extended to these three data shapes."""
    icon_path = os.path.join(proj_root, "design", "DesignData", WORKSPACE_NAME, "design", "icon.xlsx")
    item_path = os.path.join(proj_root, "design", "DesignData", WORKSPACE_NAME, "design", "item.xlsx")

    stats = {}
    for evt in event_names:
        evt_lower = evt.lower()
        icon_item_rows, icon_item_label = get_icon_item_rows_for_source(icon_path, evt_lower, None, log_dir)
        icon_event_rows, icon_event_label = get_icon_event_rows_for_source(icon_path, evt_lower, None, log_dir)
        item_icon_map, item_icon_label = get_item_icon_columns_for_source(item_path, evt, log_dir, None)

        stats[evt_lower] = {
            "icon_item": {"count": len(icon_item_rows) if icon_item_rows else 0, "source": icon_item_label},
            "icon_event": {"count": len(icon_event_rows) if icon_event_rows else 0, "source": icon_event_label},
            "item_icon": {"count": len(item_icon_map) if item_icon_map else 0, "source": item_icon_label},
        }
        log(f"Cache warm [{evt_lower}]: icon_item={stats[evt_lower]['icon_item']['count']} "
            f"icon_event={stats[evt_lower]['icon_event']['count']} "
            f"item_icon={stats[evt_lower]['item_icon']['count']}")

    return stats


# ================================================================================
# localization.xlsx source-event data: same live -> cache -> SVN-history pattern,
# for the three sheets that actually need it — localize_event_fixed (divider-bounded
# 2-row group, same shape as icon_item), localize_event (backup-first, same shape as
# icon_event), and localize_server (pure accumulator, always live so far — no sheet
# ever observed missing an older event).
# ================================================================================

def get_localize_event_fixed_rows_for_source(loc_path, src_lower, hint_date, log_dir):
    """(rows, source_label) for src_lower's 2-row group in localize_event_fixed's
    '活动名 & 礼包名区 & 成就分享区' divider-bounded section."""
    import openpyxl as _oxl
    import io as _io

    wb_live = _oxl.load_workbook(loc_path, data_only=True)
    live_rows = extract_icon_item_group_rows(wb_live, src_lower, group_marker="活动名", sheet_name="localize_event_fixed", with_style=True)
    if live_rows:
        cache = load_named_cache(log_dir, "localize_event_fixed_cache")
        cache[src_lower] = {"rows": live_rows, "source": "live", "cached_at": _cache_timestamp()}
        save_named_cache(log_dir, "localize_event_fixed_cache", cache)
        return live_rows, f"current live file ({len(live_rows)} row(s))"

    cache = load_named_cache(log_dir, "localize_event_fixed_cache")
    cached_entry = cache.get(src_lower)
    if cached_entry and cached_entry.get("rows"):
        return cached_entry["rows"], f"local cache (from {cached_entry.get('source', '?')}, cached {cached_entry.get('cached_at', '?')})"

    if not event_ever_registered_in_convert_json(loc_path, src_lower, hint_date):
        log(f"localize_event_fixed: {src_lower} never appears in convert.json (current or history) — skipping the expensive SVN search.")
        return None, None

    revisions = get_svn_revisions_in_date_range(loc_path, hint_date) if hint_date else None
    if not revisions:
        revisions = get_svn_log_revisions(loc_path, limit=60)
    if not revisions:
        return None, None

    for rev in revisions:
        content_bytes = svn_cat_revision_bytes(loc_path, rev)
        if content_bytes is None:
            continue
        try:
            wb_hist = _oxl.load_workbook(_io.BytesIO(content_bytes), data_only=True)
        except Exception:
            continue
        rows = extract_icon_item_group_rows(wb_hist, src_lower, group_marker="活动名", sheet_name="localize_event_fixed", with_style=True)
        if rows:
            cache[src_lower] = {"rows": rows, "source": f"svn r{rev}", "cached_at": _cache_timestamp()}
            save_named_cache(log_dir, "localize_event_fixed_cache", cache)
            return rows, f"SVN history (r{rev}, freshly fetched and cached)"

    return None, None


def get_localize_event_rows_for_source(loc_path, src_lower, hint_date, log_dir):
    """(rows, source_label) for src_lower's localize_event rows — checked against
    localize_旧活动备份 first (a maintained in-repo archive, free), then cache, then
    SVN history of localize_event itself. Deliberately never falls back to a
    placeholder — if this returns (None, None), the caller must log that plainly."""
    import openpyxl as _oxl
    import io as _io

    backup_sheet_name = "localize_旧活动备份"
    active_sheet_name = "localize_event"

    wb_live = _oxl.load_workbook(loc_path, data_only=True)

    def _rows_from_sheet(ws):
        rows = []
        for r in range(2, ws.max_row + 1):
            c1 = str(ws.cell(r, 1).value or "").lower()
            if src_lower in c1:
                rows.append([(ws.cell(r, c).value, cell_style_to_dict(ws.cell(r, c))) for c in range(1, ws.max_column + 1)])
        return rows

    def _header(ws):
        return [h for h in (str(ws.cell(1, c).value or "").strip() for c in range(1, ws.max_column + 1)) if h]

    def _launch_rows():
        """src_lower's localize_event rows as the event shipped: the shipping workspace's
        (WorkSpcae3) file the day after the source started. The active workspace's own history
        can be stale — WorkSpcae_Design had no localization commit from 2025-09-18 until well
        after Merge2510 launched on 2025-10-16."""
        from datetime import datetime, timedelta
        try:
            day = (datetime.strptime(hint_date.strip()[:10], "%Y-%m-%d") + timedelta(days=1)).strftime("%Y-%m-%d")
        except Exception:
            return None, None
        design_data = os.path.dirname(os.path.dirname(os.path.dirname(loc_path)))
        for path in dict.fromkeys([os.path.join(design_data, "WorkSpcae3", "design", "localization.xlsx"), loc_path]):
            data = svn_cat_revision_bytes(path, "{%s}" % day) if os.path.exists(path) else None
            try:
                wb_hist = _oxl.load_workbook(_io.BytesIO(data), data_only=True) if data else None
            except Exception:
                wb_hist = None
            if not wb_hist or active_sheet_name not in wb_hist.sheetnames:
                continue
            ws_hist = wb_hist[active_sheet_name]
            # ponytail: rows are copied by position, so a column layout change since launch skips
            # the snapshot instead of shifting translations into the wrong language; remap by header if that ever bites.
            if active_sheet_name in wb_live.sheetnames and _header(ws_hist) != _header(wb_live[active_sheet_name]):
                log(f"localize_event: {os.path.basename(os.path.dirname(os.path.dirname(path)))} @ {day} has different "
                    f"columns than today's sheet — not used.")
                continue
            rows = _rows_from_sheet(ws_hist)
            if rows:
                return rows, f"{active_sheet_name} as shipped ({os.path.basename(os.path.dirname(os.path.dirname(path)))} @ {day})"
        return None, None

    def _loc_key(v):
        return str(v or "").replace("//", "").strip().upper()

    def _topped_up(rows, label):
        # 旧活动备份 keeps only part of an event — Merge2510's archive held 94 of the 151 rows it
        # shipped with (its 建筑名 / NPC对话 groups were deleted outright), so the rest come from the launch snapshot.
        snap, snap_label = _launch_rows()
        have = {_loc_key(r[0][0]) for r in rows}
        extra = [r for r in snap or [] if _loc_key(r[0][0]) not in have]
        if not extra:
            return rows, label
        log(f"localize_event: {len(extra)} row(s) of {src_lower} missing from {label} — recovered from {snap_label}.")
        return rows + extra, f"{label} + {len(extra)} from {snap_label}"

    if backup_sheet_name in wb_live.sheetnames:
        rows = _rows_from_sheet(wb_live[backup_sheet_name])
        if rows:
            rows, label = _topped_up(rows, f"{backup_sheet_name} backup sheet ({len(rows)} row(s))")
            cache = load_named_cache(log_dir, "localize_event_cache")
            cache[src_lower] = {"rows": rows, "source": label, "cached_at": _cache_timestamp()}
            save_named_cache(log_dir, "localize_event_cache", cache)
            return rows, label

    if active_sheet_name in wb_live.sheetnames:
        rows = _rows_from_sheet(wb_live[active_sheet_name])
        if rows:
            cache = load_named_cache(log_dir, "localize_event_cache")
            cache[src_lower] = {"rows": rows, "source": f"{active_sheet_name} (live)", "cached_at": _cache_timestamp()}
            save_named_cache(log_dir, "localize_event_cache", cache)
            return rows, f"{active_sheet_name} active sheet ({len(rows)} row(s))"

    cache = load_named_cache(log_dir, "localize_event_cache")
    cached_entry = cache.get(src_lower)
    if cached_entry and cached_entry.get("rows"):
        label = f"local cache (from {cached_entry.get('source', '?')}, cached {cached_entry.get('cached_at', '?')})"
        if "as shipped" in label:  # already topped up (or taken whole) from the launch snapshot
            return cached_entry["rows"], label
        return _topped_up(cached_entry["rows"], label)

    if not event_ever_registered_in_convert_json(loc_path, src_lower, hint_date):
        log(f"localize_event: {src_lower} never appears in convert.json (current or history) and wasn't in "
            f"{backup_sheet_name} or the live {active_sheet_name} sheet — skipping the expensive SVN search.")
        return None, None

    rows, label = _launch_rows()
    if rows:
        cache[src_lower] = {"rows": rows, "source": label, "cached_at": _cache_timestamp()}
        save_named_cache(log_dir, "localize_event_cache", cache)
        return rows, f"SVN {label}"

    revisions = get_svn_revisions_in_date_range(loc_path, hint_date) if hint_date else None
    if not revisions:
        revisions = get_svn_log_revisions(loc_path, limit=60)
    if not revisions:
        return None, None

    for rev in revisions:
        content_bytes = svn_cat_revision_bytes(loc_path, rev)
        if content_bytes is None:
            continue
        try:
            wb_hist = _oxl.load_workbook(_io.BytesIO(content_bytes), data_only=True)
        except Exception:
            continue
        if active_sheet_name not in wb_hist.sheetnames:
            continue
        rows = _rows_from_sheet(wb_hist[active_sheet_name])
        if rows:
            cache[src_lower] = {"rows": rows, "source": f"svn r{rev} ({active_sheet_name})", "cached_at": _cache_timestamp()}
            save_named_cache(log_dir, "localize_event_cache", cache)
            return rows, f"SVN history (r{rev}, freshly fetched and cached)"

    return None, None


def get_localize_server_rows_for_source(loc_path, src_lower):
    """(rows, source_label) for src_lower's localize_server rows — this sheet is a
    pure accumulator (every event checked, back to 2024, is still present live), so
    unlike the other two this never needs cache/SVN — just read the live sheet."""
    import openpyxl as _oxl
    wb_live = _oxl.load_workbook(loc_path, data_only=True)
    if "localize_server" not in wb_live.sheetnames:
        return None, None
    ws = wb_live["localize_server"]
    rows = []
    for r in range(2, ws.max_row + 1):
        c1 = str(ws.cell(r, 1).value or "").lower()
        if src_lower in c1:
            rows.append([(ws.cell(r, c).value, cell_style_to_dict(ws.cell(r, c))) for c in range(1, ws.max_column + 1)])
    return (rows, f"current live file ({len(rows)} row(s))") if rows else (None, None)


def refresh_localize_event_cache(proj_root, log_dir, event_names):
    """Pre-populate localize_event_fixed / localize_event caches for the given event
    names (e.g. each type's current latest event)."""
    loc_path = os.path.join(proj_root, "design", "DesignData", WORKSPACE_NAME, "design", "localization.xlsx")

    stats = {}
    for evt in event_names:
        evt_lower = evt.lower()
        fixed_rows, fixed_label = get_localize_event_fixed_rows_for_source(loc_path, evt_lower, None, log_dir)
        event_rows, event_label = get_localize_event_rows_for_source(loc_path, evt_lower, None, log_dir)

        stats[evt_lower] = {
            "localize_event_fixed": {"count": len(fixed_rows) if fixed_rows else 0, "source": fixed_label},
            "localize_event": {"count": len(event_rows) if event_rows else 0, "source": event_label},
        }
        log(f"Cache warm [{evt_lower}]: localize_event_fixed={stats[evt_lower]['localize_event_fixed']['count']} "
            f"localize_event={stats[evt_lower]['localize_event']['count']}")

    return stats


def inject_convert_references(proj_root, src_evt, tgt_evt, root_path, step_index):
    """Step 8: Modify convert_layout.lua and convert.json to register the new event."""
    import json

    design_base = os.path.join(proj_root, os.path.join("design", "DesignData", WORKSPACE_NAME))
    convert_json_path = os.path.join(design_base, r"descriptor\app_client\convert.json")
    convert_layout_path = os.path.join(design_base, r"PostConvert\Lua\convert_layout.lua")

    # Audit existing entries before registering the new event — remove anything whose
    # Excel file no longer exists in design/, so stale references don't break coder_convert.
    if not cleanup_orphaned_convert_entries(design_base, root_path, step_index):
        return False

    # The source event's own start_time (passed by the caller as src_start_time in
    # extra_args.json) lets the SVN-history lookup below bracket its search by date
    # instead of blindly walking back from HEAD — see find_source_event_suffix_map_from_svn.
    extra = load_extra_args(root_path)
    src_start_time = extra.get("src_start_time", "")

    src_lower = src_evt.lower()
    tgt_lower = tgt_evt.lower()
    
    success_json = False
    success_lua = False
    
    # ========== Part 1: convert.json ==========
    try:
        if not os.path.exists(convert_json_path):
            log(f"Error: convert.json not found at {convert_json_path}")
            return False
        
        with open(convert_json_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        # Find the const_event.json Director entry
        const_event_dir = None
        for director in data.get("Directors", []):
            if director.get("FileName") == "const_event.json":
                const_event_dir = director
                break
        
        if not const_event_dir:
            log("Warning: 'const_event.json' Director not found in convert.json. Skipping JSON injection.")
        else:
            sheets = const_event_dir.get("ExcelSheets", [])
            tgt_entry = {"ExcelName": tgt_lower, "SheetName": f"{tgt_lower}_const"}
            
            # Check if already exists
            already_exists = any(
                s.get("ExcelName") == tgt_lower for s in sheets
            )
            
            if already_exists:
                log(f"convert.json: {tgt_lower} already registered in const_event.json. Skipping.")
            else:
                # Backup
                backup_path = convert_json_path + ".bak"
                shutil.copy2(convert_json_path, backup_path)
                record_file_creation(root_path, step_index, backup_path)
                
                sheets.append(tgt_entry)
                
                with open(convert_json_path, 'w', encoding='utf-8') as f:
                    json.dump(data, f, indent=2, ensure_ascii=False)
                
                log(f"convert.json: Added {tgt_lower} to const_event.json ExcelSheets.")
            
            success_json = True
    except Exception as e:
        log(f"Exception updating convert.json: {e}")
        traceback.print_exc()
        return False
    
    # ========== Part 2: convert_layout.lua ==========
    try:
        import glob as _glob
        import re
        
        if not os.path.exists(convert_layout_path):
            log(f"Error: convert_layout.lua not found at {convert_layout_path}")
            return False
        
        with open(convert_layout_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # Check if target event already has active (uncommented) entries
        has_active = bool(re.search(rf'(?m)^\s*\{{.*json\s*=\s*"{tgt_lower}_.*"', content))
        has_commented = bool(re.search(rf'(?m)^\s*--.*{tgt_lower}_', content))
        
        if has_active:
            log(f"convert_layout.lua: {tgt_lower} entries already exist. Skipping.")
            success_lua = True
        elif has_commented:
            # Uncomment the lines
            lines = content.split('\n')
            count = 0
            for i, line in enumerate(lines):
                stripped = line.strip()
                if tgt_lower in stripped and stripped.startswith('--') and 'json' in stripped:
                    new_stripped = stripped.lstrip('-').lstrip()
                    lines[i] = line.replace(stripped, new_stripped)
                    count += 1
            
            content = '\n'.join(lines)
            backup_path = convert_layout_path + ".bak"
            shutil.copy2(convert_layout_path, backup_path)
            record_file_creation(root_path, step_index, backup_path)
            
            with open(convert_layout_path, 'w', encoding='utf-8') as f:
                f.write(content)
            
            log(f"convert_layout.lua: Uncommented {count} entries for {tgt_lower}.")
            success_lua = True
        else:
            # Need to generate entries — read the target's _convert.json
            app_client_dir = os.path.join(design_base, r"descriptor\app_client")
            backup_desc_dir = os.path.join(design_base, r"descriptor\backup")
            
            # Find the target _convert.json file(s)
            tgt_convert_files = _glob.glob(os.path.join(app_client_dir, f"*{tgt_lower}*_convert.json"))
            if not tgt_convert_files:
                tgt_convert_files = _glob.glob(os.path.join(backup_desc_dir, f"*{tgt_lower}*_convert.json"))
            
            if not tgt_convert_files:
                log(f"Warning: No _convert.json found for {tgt_lower}. Cannot generate convert_layout entries.")
                success_lua = True
            else:
                # Collect all FileNames from the target's _convert.json(s)
                tgt_filenames = []
                for cf in tgt_convert_files:
                    with open(cf, 'r', encoding='utf-8') as f:
                        cdata = json.load(f)
                    for director in cdata.get("Directors", []):
                        fn = director.get("FileName", "")
                        if fn:
                            tgt_filenames.append(fn)
                
                log(f"Read {len(tgt_filenames)} FileNames from target _convert.json(s).")
                
                # Preferred source: the source event's OWN placement — resolved via the
                # local cache first, falling back to SVN history only if genuinely needed.
                # See get_suffix_map_for_source_event for the live-file / cache / SVN order.
                history_suffix_map, history_source_label = get_suffix_map_for_source_event(convert_layout_path, src_lower, src_start_time, root_path)

                # Fallback source: whatever's currently in the live file, for any
                # suffix history didn't cover (unknown source, no SVN, new file type, etc).
                section_end_lines, current_suffix_map = parse_convert_layout_suffix_map(content)

                suffix_to_info = dict(current_suffix_map)
                if history_suffix_map:
                    suffix_to_info.update(history_suffix_map)  # history wins on overlap
                    # History entries carry a section that may not exist in the CURRENT
                    # file yet (e.g. the source's section got removed since). Only trust
                    # a history-provided section if it's actually present today; otherwise
                    # the insertion fallback below drops to default_map.
                log(f"Suffix lookup: {len(current_suffix_map)} pattern(s) from current file, "
                    f"{len(history_suffix_map) if history_suffix_map else 0} pattern(s) from {src_lower}'s own data"
                    f"{f' — {history_source_label}' if history_source_label else ''}.")

                # Now generate entries for each target FileName
                # Group by section for insertion
                section_entries = {}  # section -> list of new lines
                resolved_from_history = 0

                for fn in tgt_filenames:
                    # Merge events: <tgt>_machineItem.json is handled separately below by
                    # update_merge_item_to_machine_slot (single-slot replace, not append).
                    if tgt_lower.startswith("merge") and fn == f"{tgt_lower}_machineItem.json":
                        continue

                    # Derive the file suffix: strip target event prefix
                    if fn.startswith(tgt_lower + '_'):
                        file_suffix = fn[len(tgt_lower)+1:]  # e.g., "scene.json"
                    else:
                        file_suffix = fn

                    # Look up the suffix in our mapping (SVN history first, then current-file pattern)
                    if file_suffix in suffix_to_info:
                        section, merge_suffix, sample_line = suffix_to_info[file_suffix]
                        if history_suffix_map and file_suffix in history_suffix_map:
                            resolved_from_history += 1
                    else:
                        # Fallback: use default_map, derive merge_suffix from file_suffix
                        section = "default_map"
                        merge_suffix = file_suffix.replace('.json', '')

                    merge_as = f"{tgt_lower}.{merge_suffix}"
                    new_line = f'        {{json = "{fn}", merge_as = "{merge_as}"}},'

                    if section not in section_entries:
                        section_entries[section] = []
                    section_entries[section].append(new_line)
                
                # Insert entries into convert_layout.lua, at the end of each section
                lines = content.split('\n')
                offset = 0
                
                backup_path = convert_layout_path + ".bak"
                shutil.copy2(convert_layout_path, backup_path)
                record_file_creation(root_path, step_index, backup_path)
                
                total_injected = 0
                for section, entries in section_entries.items():
                    if section in section_end_lines:
                        insert_idx = section_end_lines[section] + 1 + offset
                        new_block = [''] + entries
                        for j, entry in enumerate(new_block):
                            lines.insert(insert_idx + j, entry)
                        offset += len(new_block)
                        total_injected += len(entries)
                        log(f"  Injected {len(entries)} entries into '{section}' section.")
                    else:
                        # Section not found, inject into default_map as fallback
                        if "default_map" in section_end_lines:
                            insert_idx = section_end_lines["default_map"] + 1 + offset
                            new_block = [''] + entries
                            for j, entry in enumerate(new_block):
                                lines.insert(insert_idx + j, entry)
                            offset += len(new_block)
                            total_injected += len(entries)
                            log(f"  Injected {len(entries)} entries into 'default_map' (fallback for '{section}').")
                
                content = '\n'.join(lines)
                
                with open(convert_layout_path, 'w', encoding='utf-8') as f:
                    f.write(content)
                
                log(f"convert_layout.lua: Injected {total_injected} total entries for {tgt_lower} "
                    f"({resolved_from_history} resolved from {src_lower}'s own data, {total_injected - resolved_from_history} from current-file pattern/default_map fallback).")
                success_lua = True
    except Exception as e:
        log(f"Exception updating convert_layout.lua: {e}")
        traceback.print_exc()
        return False

    # Merge events only: convert_merge2406_item_to_machine is a separate, independent
    # single-slot table — handled after the main entry logic regardless of which of the
    # three cases above fired (it's not part of the accumulating default_map convention).
    success_item_to_machine = True
    if tgt_lower.startswith("merge"):
        try:
            success_item_to_machine = update_merge_item_to_machine_slot(convert_layout_path, tgt_lower, root_path, step_index)
        except Exception as e:
            log(f"Exception updating convert_merge2406_item_to_machine: {e}")
            traceback.print_exc()
            success_item_to_machine = False

    return success_json and success_lua and success_item_to_machine


def update_merge_item_to_machine_slot(convert_layout_path, tgt_lower, root_path, step_index):
    """
    convert_merge2406_item_to_machine is a SINGLE-SLOT table, not an accumulating
    list like every other section in convert_layout.lua — its one entry gets
    REPLACED with each new merge event, not appended alongside old ones. The
    table's name is legacy (frozen at "2406" from whenever it was first created)
    but its content is actively maintained. Confirmed by cross-branch inspection
    and SVN history: merge2406 -> merge2603 -> merge2606 -> merge2608, always
    exactly one entry present at a time.
    """
    import re as _re

    if not os.path.exists(convert_layout_path):
        return True  # Earlier checks in inject_convert_references already failed loudly if this matters.

    with open(convert_layout_path, 'r', encoding='utf-8') as f:
        content = f.read()

    target_line = f'        {{json = "{tgt_lower}_machineItem.json", merge_as="{tgt_lower}.itemToMachine"}}'

    pattern = _re.compile(r'(convert_merge2406_item_to_machine\s*=\s*\{)(.*?)(\n?\s*\},)', _re.DOTALL)
    match = pattern.search(content)

    if not match:
        log("convert_layout.lua: no convert_merge2406_item_to_machine table found in this file — skipping.")
        return True

    current_inner = match.group(2)
    if f'"{tgt_lower}_machineItem.json"' in current_inner:
        log(f"convert_layout.lua: convert_merge2406_item_to_machine already holds {tgt_lower}'s entry. Skipping.")
        return True

    backup_path = convert_layout_path + ".bak"
    if not os.path.exists(backup_path):
        # Don't clobber a .bak already made earlier in this same step (Part 2's own
        # backup) — the first backup in a run should always capture the true pre-run state.
        shutil.copy2(convert_layout_path, backup_path)
        record_file_creation(root_path, step_index, backup_path)

    new_content = pattern.sub(lambda m: m.group(1) + "\n" + target_line + m.group(3), content, count=1)

    with open(convert_layout_path, 'w', encoding='utf-8') as f:
        f.write(new_content)

    old_display = current_inner.strip() or "(empty)"
    log(f"convert_layout.lua: replaced convert_merge2406_item_to_machine slot with {tgt_lower}'s entry (was: {old_display}).")
    return True

def update_bi_excel(proj_root, src_evt, tgt_evt, root_path, step_index):
    """Step 9: Update bi.xlsx with target event tracking entries."""
    import openpyxl
    import re
    
    design_dir = os.path.join(proj_root, "design", "DesignData", WORKSPACE_NAME, "design")
    bi_path = os.path.join(design_dir, "bi.xlsx")
    
    if not os.path.exists(bi_path):
        log(f"Error: bi.xlsx not found at {bi_path}")
        return False
    
    src_lower = src_evt.lower()
    tgt_lower = tgt_evt.lower()
    src_upper = src_evt.upper()
    tgt_upper = tgt_evt.upper()
    # PascalCase: first letter upper, rest as-is from the original
    src_pascal = src_evt[0].upper() + src_evt[1:] if src_evt else src_evt
    tgt_pascal = tgt_evt[0].upper() + tgt_evt[1:] if tgt_evt else tgt_evt
    
    def swap_event_name(value):
        """Replace all case variants of the source event name with the target."""
        if value is None:
            return None
        s = str(value)
        s = s.replace(src_upper, tgt_upper)
        s = s.replace(src_lower, tgt_lower)
        s = s.replace(src_pascal, tgt_pascal)
        # Also try the original casing
        s = s.replace(src_evt, tgt_evt)
        return s
    
        return False

    try:
        # Backup first
        backup_path = bi_path + ".bak"
        if os.path.exists(backup_path):
            import stat
            os.chmod(backup_path, stat.S_IWRITE)
            os.remove(backup_path)
        shutil.copy2(bi_path, backup_path)
        record_file_creation(root_path, step_index, backup_path)
        
        wb = load_workbook_frozen(bi_path)
        modified = False
        
        # ========== Sheet 1: bi_reason ==========
        if "bi_reason" in wb.sheetnames:
            ws = wb["bi_reason"]
            # check_and_handle_existing always returns a 2-tuple, so `if check_and_handle_existing(...)`
            # was always truthy regardless of the actual result — unpack it properly.
            exists, _ = check_and_handle_existing(ws, "bi_reason", tgt_evt)
            if exists:
                modified = True # Might have uncommented
            else:
                # Find source event row
                src_row = None
                for r in range(1, ws.max_row + 1):
                    val = ws.cell(r, 1).value
                    if val and src_lower in str(val).lower():
                        src_row = r
                        break
                
                if src_row:
                    new_row = ws.max_row + 1
                    for c in range(1, ws.max_column + 1):
                        src_c = ws.cell(src_row, c)
                        old_val = src_c.value
                        dst_c = ws.cell(new_row, c)
                        dst_c.value = swap_event_name(old_val)
                        copy_cell_format(src_c, dst_c)
                    log(f"bi_reason: Added row {new_row} for {tgt_evt}.")
                    modified = True
                else:
                    log(f"bi_reason: Source event {src_evt} not found. Skipping.")
        
        # ========== Sheet 2: 广告打点常量表 ==========
        ad_sheet_name = "广告打点常量表"
        if ad_sheet_name in wb.sheetnames:
            ws = wb[ad_sheet_name]
            exists, _ = check_and_handle_existing(ws, ad_sheet_name, tgt_evt)
            if exists:
                modified = True
            else:
                # ws.max_column reports 16382 here — stray formatting on whole columns,
                # only ~5 ever actually populated. Use the real width, not the phantom one
                # (see get_real_max_col) — otherwise every row read/write below iterates
                # 16382 cells instead of 5.
                real_max_col = get_real_max_col(ws)

                # Collect source rows grouped by prefix (SITE vs ENTRANCE), remembering
                # each group's last row so new rows land right after their own block
                # instead of at the sheet's unrelated tail.
                groups = {}  # prefix -> [(row_idx, [(value, cell)])]
                for r in range(2, ws.max_row + 1):
                    val = ws.cell(r, 1).value
                    if val and src_upper in str(val).upper():
                        enum_val = str(val)
                        # Determine group: AD_BI_SITE_ or AD_BI_ENTRANCE_
                        if 'AD_BI_SITE_' in enum_val:
                            group_key = 'AD_BI_SITE'
                        elif 'AD_BI_ENTRANCE_' in enum_val:
                            group_key = 'AD_BI_ENTRANCE'
                        else:
                            group_key = 'OTHER'

                        row_data = []
                        for c in range(1, real_max_col + 1):
                            row_data.append((ws.cell(r, c).value, ws.cell(r, c)))

                        if group_key not in groups:
                            groups[group_key] = []
                        groups[group_key].append((r, row_data))

                # Insert each group's new rows immediately after that group's own last
                # source row, preserving the sheet's SITE/ENTRANCE block layout. We avoid
                # ws.insert_rows (which shifts every one of the ~16382 phantom-formatted
                # columns and measured 15+ seconds for a single 3-row insert on this sheet)
                # by manually shifting only the real_max_col columns that actually hold
                # data. Groups are processed bottom-most-first so an earlier insertion
                # never invalidates a not-yet-processed group's row positions.
                for group_key, rows_data in sorted(groups.items(), key=lambda kv: kv[1][-1][0], reverse=True):
                    num_new = len(rows_data)
                    last_src_row = rows_data[-1][0]
                    insert_at = last_src_row + 1
                    old_max_row = ws.max_row

                    # Stray merged cells (part of this sheet's phantom formatting bloat,
                    # same cause as the bogus max_column) can fall inside the shift range
                    # and raise "MergedCell object attribute 'value' is read-only" when
                    # written to. Unmerge anything overlapping the affected columns first.
                    for mc in list(ws.merged_cells.ranges):
                        if mc.max_row >= insert_at and mc.min_col <= real_max_col:
                            ws.unmerge_cells(str(mc))

                    # Shift existing rows [insert_at, old_max_row] down by num_new,
                    # bottom-up so we never overwrite a row before it's been copied.
                    for r in range(old_max_row, insert_at - 1, -1):
                        for c in range(1, real_max_col + 1):
                            src_c = ws.cell(r, c)
                            dst_c = ws.cell(r + num_new, c)
                            dst_c.value = src_c.value
                            copy_cell_format(src_c, dst_c)

                    for i, (_, cell_values) in enumerate(rows_data):
                        for c in range(1, len(cell_values) + 1):
                            val, src_c = cell_values[c-1]
                            dst_c = ws.cell(insert_at + i, c)
                            dst_c.value = swap_event_name(val)
                            copy_cell_format(src_c, dst_c)

                    log(f"{ad_sheet_name}: Inserted {num_new} rows for {tgt_evt} in {group_key} group at row {insert_at}.")
                    modified = True
        
        # ========== Sheet 3: bi_event_names ==========
        if "bi_event_names" in wb.sheetnames:
            ws = wb["bi_event_names"]
            exists, _ = check_and_handle_existing(ws, "bi_event_names", tgt_evt)
            if exists:
                modified = True
            else:
                # Find source event row
                src_row = None
                for r in range(1, ws.max_row + 1):
                    val = ws.cell(r, 1).value
                    if val and src_lower in str(val).lower():
                        src_row = r
                        break
                
                if src_row:
                    new_row = ws.max_row + 1
                    for c in range(1, ws.max_column + 1):
                        src_c = ws.cell(src_row, c)
                        old_val = src_c.value
                        dst_c = ws.cell(new_row, c)
                        dst_c.value = swap_event_name(old_val)
                        copy_cell_format(src_c, dst_c)
                    log(f"bi_event_names: Added row {new_row} for {tgt_evt}.")
                    modified = True
                else:
                    log(f"bi_event_names: Source event {src_evt} not found. Skipping.")
        
        if modified:
            import stat
            os.chmod(bi_path, stat.S_IWRITE | stat.S_IREAD)
            wb.save(bi_path)
            log(f"bi.xlsx saved successfully.")
        else:
            log(f"bi.xlsx: No changes were needed.")
        
        return True
    except Exception as e:
        log(f"Exception updating bi.xlsx: {e}")
        traceback.print_exc()
        return False

def load_extra_args(root_path):
    """Load extra arguments from extra_args.json written by the C# GUI."""
    args_path = os.path.join(root_path, "extra_args.json")
    if os.path.exists(args_path):
        with open(args_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {}

def parse_buff_header(line):
    """
    Recognizes a "活动buff常量" biweekly-group header line and returns the event
    name it identifies, or None if the line isn't a group header at all.

    Confirmed against real historical data that these headers are NOT consistently
    formatted — some use "//正常双周活动(Merge2608)", others "//双周活动Sim2605"
    (no "正常" prefix, no parens at all). Anchoring on the "双周活动" marker itself,
    rather than requiring the full prefix + parens, handles both. A line missing
    even the "双周活动" marker (a genuine typo, seen once in real data) is left
    unrecognized — that's a data problem, not something a header parser should guess at.
    """
    if not line.startswith("//") or "双周活动" not in line:
        return None
    name = line.split("双周活动", 1)[1].strip()
    name = name.lstrip("(").rstrip(")").strip()
    return name or None


def rekey_id_prefixed(value, src_id, tgt_id, swap):
    """
    Re-keys a "<event id>_<key>" cell (活动buff常量 col 1) to the target event id and
    swaps names in the rest. The prefix is tested on the ORIGINAL value: swap's digit
    fallback rewrites the month ("202610" -> "202611"), so testing after the swap missed
    every cross-month clone — Blocks2610 -> Tiles2611 got 20261101_* (Merge2611's id)
    instead of the GUI's 20261102_*, and Tiles2611EventCtrl crashed on every game end.
    """
    s = "" if value is None else str(value)
    if src_id and tgt_id and s.startswith(src_id):
        return str(tgt_id) + swap(s[len(src_id):])
    return swap(value)


def update_events_excel(proj_root, src_evt, tgt_evt, root_path, step_index):
    """Step 10: Update events.xlsx with target event registration data."""
    import openpyxl
    import stat
    
    design_dir = os.path.join(proj_root, "design", "DesignData", WORKSPACE_NAME, "design")
    events_path = os.path.join(design_dir, "events.xlsx")
    
    if not os.path.exists(events_path):
        log(f"Error: events.xlsx not found at {events_path}")
        return False
    
    extra = load_extra_args(root_path)
    event_id = extra.get("event_id", "")
    start_time = extra.get("start_time", "")
    end_time = extra.get("end_time", "")
    near_end_time = extra.get("near_end_time", "")
    close_time = extra.get("close_time", "")
    is_reopen = extra.get("is_reopen", "0") == "1"
    
    src_lower = src_evt.lower()
    tgt_lower = tgt_evt.lower()
    src_upper = src_evt.upper()
    tgt_upper = tgt_evt.upper()
    src_pascal = src_evt[0].upper() + src_evt[1:] if src_evt else src_evt
    tgt_pascal = tgt_evt[0].upper() + tgt_evt[1:] if tgt_evt else tgt_evt
    
    def swap_name(value):
        if value is None:
            return None
        s = str(value)
        # Standard string swaps
        if src_evt:
            s = s.replace(src_upper, tgt_upper)
            s = s.replace(src_lower, tgt_lower)
            s = s.replace(src_pascal, tgt_pascal)
            s = s.replace(src_evt, tgt_evt)
        # Fallback for ID-like columns (e.g. 20250701 -> 20260501)
        src_digits = "".join(filter(str.isdigit, src_evt))
        tgt_digits = "".join(filter(str.isdigit, tgt_evt))
        if len(src_digits) >= 4 and len(tgt_digits) >= 4:
            s = s.replace("20" + src_digits[:4], "20" + tgt_digits[:4])
        return s
    
    def find_src_event_id(ws):
        """Find the Event ID used by the source event in a sheet."""
        for r in range(2, ws.max_row + 1):
            for c in range(1, ws.max_column + 1):
                val = ws.cell(r, c).value
                if val and src_lower in str(val).lower():
                    # Return col 1 value of this row as the event ID
                    return str(ws.cell(r, 1).value) if ws.cell(r, 1).value else None
        return None
    
        return False, None


    
    try:
        # Backup
        backup_path = events_path + ".bak"
        if os.path.exists(backup_path):
            os.chmod(backup_path, stat.S_IWRITE)
            os.remove(backup_path)
        shutil.copy2(events_path, backup_path)
        record_file_creation(root_path, step_index, backup_path)
        
        wb = load_workbook_frozen(events_path)
        modified = False
        src_event_id = None  # Will be discovered from data
        
        # ========== Sheet 1: events (main) ==========
        if "events" in wb.sheetnames:
            ws = wb["events"]
            exists, row_idx = check_and_handle_existing(ws, "events", tgt_evt)
            if exists:
                # Even if it exists, we want to update dates and IDs if provided
                target_row = row_idx
                modified = True
            else:
                src_row = None
                real_max = get_real_max_row(ws)
                for r in range(2, real_max + 1):
                    for c in range(1, ws.max_column + 1):
                        val = ws.cell(r, c).value
                        if val and src_lower in str(val).lower():
                            src_row = r
                            break
                    if src_row:
                        break
                
                if src_row:
                    src_event_id = str(ws.cell(src_row, 1).value)
                    last_row_idx = real_max
                    target_row = last_row_idx + 1
                    
                    # Set the last row's control switch to FALSE (col 6) - as it's the "previous" event
                    if last_row_idx >= 2:
                        ws.cell(last_row_idx, 6, "FALSE")
                        log(f"events: Set previous event (row {last_row_idx}) control switch to FALSE.")
                    
                    # Copy all columns from template source
                    for c in range(1, ws.max_column + 1):
                        src_c = ws.cell(src_row, c)
                        old_val = src_c.value
                        dst_c = ws.cell(target_row, c)
                        dst_c.value = swap_name(old_val)
                        copy_cell_format(src_c, dst_c)
                    
                    # Set previous event ref (col 23)
                    if last_row_idx >= 2:
                        last_alias = ws.cell(last_row_idx, 2).value
                        if last_alias:
                            ws.cell(target_row, 23, str(last_alias))
                    
                    log(f"events: Added row {target_row} for {tgt_evt}.")
                    modified = True
                else:
                    target_row = None
                    log(f"events: Source event {src_evt} not found.")

            # Apply overrides to the target_row (whether new or existing/uncommented)
            if target_row:
                # Override specific columns
                # Col 1: Event ID from GUI input
                if event_id:
                    ws.cell(target_row, 1, int(event_id))
                
                # Col 4: BI Title = first 6 digits of ID + target name
                if event_id and len(event_id) >= 6:
                    bi_title = event_id[:6] + tgt_pascal
                    ws.cell(target_row, 4, bi_title)
                
                # Col 5: Localization key swap
                ws.cell(target_row, 5, f"{tgt_upper}_EVENTNAME")
                
                # Col 6: Control switch = TRUE for new event
                ws.cell(target_row, 6, "TRUE")
                
                # Col 8-11: Dates from GUI
                if start_time:
                    ws.cell(target_row, 8, start_time)
                if end_time:
                    ws.cell(target_row, 9, end_time)
                if near_end_time:
                    ws.cell(target_row, 10, near_end_time)
                if close_time:
                    ws.cell(target_row, 11, close_time)
                
                modified = True
        
        # Discover src_event_id if not yet found
        if not src_event_id and "events" in wb.sheetnames:
            src_event_id = find_src_event_id(wb["events"])
        
        # ========== Sheet 2: event_type ==========
        if "event_type" in wb.sheetnames:
            ws = wb["event_type"]
            exists, _ = check_and_handle_existing(ws, "event_type", tgt_evt)
            if exists:
                modified = True
            else:
                src_row = None
                real_max = get_real_max_row(ws)
                for r in range(2, real_max + 1):
                    val = ws.cell(r, 3).value  # Col 3 is the enum
                    if val and src_lower in str(val).lower():
                        src_row = r
                        break
                
                if src_row:
                    new_row = real_max + 1
                    last_id = get_last_numeric_id(ws, 1)
                    
                    for c in range(1, ws.max_column + 1):
                        src_c = ws.cell(src_row, c)
                        old_val = src_c.value
                        dst_c = ws.cell(new_row, c)
                        dst_c.value = swap_name(old_val)
                        copy_cell_format(src_c, dst_c)
                    
                    # Col 1: Auto-increment type ID
                    ws.cell(new_row, 1, last_id + 1)
                    
                    log(f"event_type: Added row {new_row} (ID={last_id+1}) for {tgt_evt}.")
                    modified = True
        
        # ========== Sheet 3: event_entrance ==========
        if "event_entrance" in wb.sheetnames:
            ws = wb["event_entrance"]
            exists, _ = check_and_handle_existing(ws, "event_entrance", tgt_evt)
            if exists:
                modified = True
            else:
                src_row = None
                real_max = get_real_max_row(ws)
                for r in range(2, real_max + 1):
                    val = ws.cell(r, 2).value
                    if val and src_lower in str(val).lower():
                        src_row = r
                        break
                
                if src_row:
                    new_row = real_max + 1
                    last_id = get_last_numeric_id(ws, 1)
                    
                    for c in range(1, ws.max_column + 1):
                        src_c = ws.cell(src_row, c)
                        old_val = src_c.value
                        dst_c = ws.cell(new_row, c)
                        dst_c.value = swap_name(old_val)
                        copy_cell_format(src_c, dst_c)
                    
                    ws.cell(new_row, 1, last_id + 1)
                    
                    log(f"event_entrance: Added row {new_row} (ID={last_id+1}) for {tgt_evt}.")
                    modified = True
        
        # ========== Sheet 4: event_score ==========
        if "event_score" in wb.sheetnames:
            ws = wb["event_score"]
            exists, _ = check_and_handle_existing(ws, "event_score", tgt_evt)
            if exists:
                modified = True
            else:
                src_row = None
                real_max = get_real_max_row(ws)
                for r in range(2, real_max + 1):
                    val = ws.cell(r, 2).value
                    if val and src_lower in str(val).lower():
                        src_row = r
                        break
                
                if src_row:
                    new_row = real_max + 1
                    last_id = get_last_numeric_id(ws, 1)
                    
                    for c in range(1, ws.max_column + 1):
                        src_c = ws.cell(src_row, c)
                        old_val = src_c.value
                        dst_c = ws.cell(new_row, c)
                        dst_c.value = swap_name(old_val)
                        copy_cell_format(src_c, dst_c)
                    
                    ws.cell(new_row, 1, last_id + 1)
                    
                    log(f"event_score: Added row {new_row} (ID={last_id+1}) for {tgt_evt}.")
                    modified = True
        
        # ========== Sheet 5: event_pre ==========
        if "event_pre" in wb.sheetnames:
            ws = wb["event_pre"]
            exists, row_idx = check_and_handle_existing(ws, "event_pre", tgt_evt)
            if exists:
                # Update ID if it exists
                if event_id:
                    ws.cell(row_idx, 1, int(event_id))
                modified = True
            else:
                src_row = None
                real_max = get_real_max_row(ws)
                for r in range(2, real_max + 1):
                    val = ws.cell(r, 2).value
                    if val and src_lower in str(val).lower():
                        src_row = r
                        break
                
                if src_row:
                    new_row = real_max + 1
                    for c in range(1, ws.max_column + 1):
                        src_c = ws.cell(src_row, c)
                        old_val = src_c.value
                        dst_c = ws.cell(new_row, c)
                        dst_c.value = swap_name(old_val)
                        copy_cell_format(src_c, dst_c)
                    
                    # Col 1: Use new event ID
                    if event_id:
                        ws.cell(new_row, 1, int(event_id))
                    
                    log(f"event_pre: Added row {new_row} for {tgt_evt}.")
                    modified = True
        
        # ========== Sheet 6: event_post ==========
        if "event_post" in wb.sheetnames:
            ws = wb["event_post"]
            exists, row_idx = check_and_handle_existing(ws, "event_post", tgt_evt)
            if exists:
                # Update ID if it exists
                if event_id:
                    ws.cell(row_idx, 1, int(event_id))
                modified = True
            else:
                src_row = None
                real_max = get_real_max_row(ws)
                for r in range(2, real_max + 1):
                    val = ws.cell(r, 2).value
                    if val and src_lower in str(val).lower():
                        src_row = r
                        break
                
                if src_row:
                    new_row = real_max + 1
                    for c in range(1, ws.max_column + 1):
                        src_c = ws.cell(src_row, c)
                        old_val = src_c.value
                        dst_c = ws.cell(new_row, c)
                        dst_c.value = swap_name(old_val)
                        copy_cell_format(src_c, dst_c)
                    
                    if event_id:
                        ws.cell(new_row, 1, int(event_id))
                    
                    log(f"event_post: Added row {new_row} for {tgt_evt}.")
                    
                    # Update switches: last two = TRUE, rest = FALSE
                    # new_row is the last row.
                    for r in range(2, new_row + 1):
                        # Switch is in Column 3
                        # We want row new_row and new_row-1 to be TRUE
                        if r >= (new_row - 1):
                            ws.cell(r, 3, "TRUE")
                        else:
                            ws.cell(r, 3, "FALSE")
                    log(f"event_post: Set TRUE for last two rows in switch column.")
                    
                    modified = True
        
        # ========== Sheet 7: 活动其他 ==========
        other_sheet = "活动其他"
        if other_sheet in wb.sheetnames:
            ws = wb[other_sheet]
            exists, row_idx = check_and_handle_existing(ws, "activity_other", tgt_evt)
            if exists:
                if event_id:
                    ws.cell(row_idx, 1, int(event_id))
                modified = True
            else:
                src_row = None
                real_max = get_real_max_row(ws)
                for r in range(2, real_max + 1):
                    for c in range(1, ws.max_column + 1):
                        val = ws.cell(r, c).value
                        if val and src_lower in str(val).lower():
                            src_row = r
                            break
                    if src_row:
                        break
                
                if src_row:
                    new_row = real_max + 1
                    for c in range(1, ws.max_column + 1):
                        src_c = ws.cell(src_row, c)
                        old_val = src_c.value
                        dst_c = ws.cell(new_row, c)
                        dst_c.value = swap_name(old_val)
                        copy_cell_format(src_c, dst_c)
                    
                    if event_id:
                        ws.cell(new_row, 1, int(event_id))
                    
                    log(f"activity_other: Added row {new_row} for {tgt_evt}.")
                    
                    # New adjustment: Comment out all rows except the last two
                    # After adding, new_row is the last row.
                    # We need rows (new_row) and (new_row-1) uncommented.
                    # Everything from 2 up to (new_row-2) should be commented.
                    for r in range(2, new_row - 1):
                        val = ws.cell(r, 1).value
                        if val is not None:
                            s = str(val)
                            if not s.startswith("//"):
                                ws.cell(r, 1, "//" + s)
                                log(f"activity_other: Commented out old row {r}.")
                    
                    modified = True
        
        # ========== Sheet 8: 活动成就 ==========
        achieve_sheet = "活动成就"
        if achieve_sheet in wb.sheetnames:
            ws = wb[achieve_sheet]
            exists, row_idx = check_and_handle_existing(ws, "activity_achievement", tgt_evt)
            if exists:
                target_row = row_idx
                modified = True
            else:
                # For achievements, we need a reference row - try source first, then any recent event
                src_row = None
                real_max = get_real_max_row(ws)
                for r in range(2, real_max + 1):
                    for c in range(1, ws.max_column + 1):
                        val = ws.cell(r, c).value
                        if val and src_lower in str(val).lower():
                            src_row = r
                            break
                    if src_row:
                        break
                
                # If source not found, use the last row as template
                if not src_row and real_max >= 2:
                    src_row = real_max
                    log(f"activity_achievement: Source not found, using last row {src_row} as template.")
                
                if src_row:
                    new_row = real_max + 1
                    target_row = new_row
                    for c in range(1, ws.max_column + 1):
                        src_c = ws.cell(src_row, c)
                        old_val = src_c.value
                        dst_c = ws.cell(target_row, c)
                        dst_c.value = swap_name(old_val)
                        copy_cell_format(src_c, dst_c)
                    
                    # Col 3: Switch = TRUE
                    ws.cell(target_row, 3, "TRUE")
                    
                    # Col 4: Icon - swap name
                    ws.cell(target_row, 4, f"{tgt_pascal}活动成就" if not ws.cell(src_row, 4).value else swap_name(ws.cell(src_row, 4).value))
                    
                    # Col 5: Localization key
                    ws.cell(new_row, 5, f"ACHIEVEMENT_{tgt_upper}")
                    
                    # Col 6: DeepLink - auto increment
                    last_deeplink = get_last_numeric_id(ws, 6)
                    ws.cell(new_row, 6, last_deeplink + 1)
                    
                    log(f"activity_achievement: Added row {new_row} for {tgt_evt}.")
                    modified = True
        
        # ========== Sheet 9: \u6d3b\u52a8buff\u5e38\u91cf ==========
        buff_sheet = "\u6d3b\u52a8buff\u5e38\u91cf"
        buff_backup_sheet = "buff\u8001\u6d3b\u52a8\u5907\u4efd"
        if buff_sheet in wb.sheetnames and buff_backup_sheet in wb.sheetnames:
            ws = wb[buff_sheet]
            ws_backup = wb[buff_backup_sheet]

            # Step 1: Collect all event groups currently in the main sheet
            # A group consists of a header line and rows starting with 202xxxxx.
            # Real historical headers use several inconsistent formats \u2014 confirmed
            # against actual data: "//\u6b63\u5e38\u53cc\u5468\u6d3b\u52a8(Merge2608)",
            # "//\u53cc\u5468\u6d3b\u52a8Sim2605" (no "\u6b63\u5e38" prefix, no parens \u2014 this is what broke
            # Sim2605), even "//\u6b63\u5e38\u53cc\u5468\u6d3b\u52a8\u5496\u5561\u5e97\u7ecf\u8425)" (missing opening paren).
            # parse_buff_header handles all of these by anchoring on "\u53cc\u5468\u6d3b\u52a8" itself
            # rather than requiring the full prefix + parens.
            main_groups = []
            current_group = None

            for r in range(1, ws.max_row + 1):
                c1 = ws.cell(r, 1).value
                if not c1:
                    if current_group:
                        main_groups.append(current_group)
                        current_group = None
                    continue

                s1 = str(c1)
                evt_name = parse_buff_header(s1)
                if evt_name:
                    if current_group:
                        main_groups.append(current_group)
                    current_group = {
                        'name': evt_name,
                        'rows': [[(ws.cell(r, cc).value, ws.cell(r, cc)) for cc in range(1, ws.max_column + 1)]]
                    }
                elif current_group and s1.startswith("202"):
                    # Data row belonging to the event
                    current_group['rows'].append([(ws.cell(r, cc).value, ws.cell(r, cc)) for cc in range(1, ws.max_column + 1)])
                else:
                    # Some other row: beginner or something else
                    if current_group:
                        main_groups.append(current_group)
                        current_group = None

            if current_group:
                main_groups.append(current_group)

            # Step 2: Move missing event groups from main to backup
            # Collect existing event names in backup
            backup_event_names = set()
            for r in range(1, ws_backup.max_row + 1):
                c1 = ws_backup.cell(r, 1).value
                if c1:
                    evt_name = parse_buff_header(str(c1))
                    if evt_name:
                        backup_event_names.add(evt_name.lower())
            
            for group in main_groups:
                if group['name'].lower() not in backup_event_names:
                    log(f"buff_constants: Backing up event {group['name']} to {buff_backup_sheet}")
                    for row_vals in group['rows']:
                        next_bk_row = get_real_max_row(ws_backup) + 1
                        for c_idx, (val, src_c) in enumerate(row_vals):
                            dst_c = ws_backup.cell(next_bk_row, c_idx + 1)
                            dst_c.value = val
                            copy_cell_format(src_c, dst_c)
                    modified = True
            
            # Step 3: Clear Main sheet of all bi-weekly events before injecting target
            # Skip clearing/injection if target already in main sheet unless we want a rebuild
            exists_idx, _ = check_and_handle_existing(ws, "buff_constants", tgt_evt)
            if exists_idx:
                log("buff_constants: Target already exists in main sheet. Skipping injection.")
            else:
                rows_to_delete = []
                for r in range(1, ws.max_row + 1):
                    c1 = ws.cell(r, 1).value
                    if c1:
                        s1 = str(c1)
                        if parse_buff_header(s1) or s1.startswith("202"):
                            rows_to_delete.append(r)
                
                # Deleting from bottom up to preserve indices
                for r in sorted(rows_to_delete, reverse=True):
                    ws.delete_rows(r)
                log(f"buff_constants: Removed {len(rows_to_delete)} bi-weekly rows from main sheet.")
                
                # Step 4: Inject target event from backup
                src_template_rows = []
                src_id_prefix = None
                
                # Search backup for source event
                found_src = False
                for r in range(1, ws_backup.max_row + 1):
                    c1 = ws_backup.cell(r, 1).value
                    if c1 and parse_buff_header(str(c1)) and src_lower in str(c1).lower():
                        # Found it!
                        found_src = True
                        src_template_rows.append([(ws_backup.cell(r, cc).value, ws_backup.cell(r, cc)) for cc in range(1, ws_backup.max_column + 1)])
                        # Collect following 202xxxxx rows
                        nr = r + 1
                        while nr <= ws_backup.max_row:
                            nc1 = ws_backup.cell(nr, 1).value
                            if nc1 and str(nc1).startswith("202"):
                                data_row = [(ws_backup.cell(nr, cc).value, ws_backup.cell(nr, cc)) for cc in range(1, ws_backup.max_column + 1)]
                                if not src_id_prefix:
                                    src_id_prefix = str(nc1).split('_')[0]
                                src_template_rows.append(data_row)
                                nr += 1
                            else:
                                break
                        break
                
                if found_src and src_template_rows:
                    log(f"buff_constants: Templating {tgt_evt} from {src_evt} found in backup.")
                    
                    # Instead of append, use explicit row to avoid gaps from delete_rows residual max_row
                    curr_row = get_real_max_row(ws) + 1
                    
                    for row_data in src_template_rows:
                        for c_idx, (val, src_c) in enumerate(row_data):
                            # Col 1 (Program Enum with ID prefix) takes the GUI's event id
                            if c_idx == 0:
                                new_val = rekey_id_prefixed(val, src_id_prefix, event_id, swap_name)
                            else:
                                new_val = swap_name(val)
                            dst_c = ws.cell(curr_row, c_idx + 1)
                            dst_c.value = new_val
                            copy_cell_format(src_c, dst_c)
                        curr_row += 1
                    modified = True
                else:
                    log(f"buff_constants: Source event {src_evt} not found in backup sheet.")
        
        if modified:
            os.chmod(events_path, stat.S_IWRITE | stat.S_IREAD)
            wb.save(events_path)
            log(f"events.xlsx saved successfully.")
        else:
            log(f"events.xlsx: No changes were needed.")
        
        return True
    except Exception as e:
        log(f"Exception updating events.xlsx: {e}")
        traceback.print_exc()
        return False

def update_item_excel(proj_root, src_evt, tgt_evt, log_dir, step_index):
    """
    Step 12: Update item.xlsx (Item definitions)
    """
    log(f"Step 12: Updating item.xlsx for {tgt_evt} (cloning from {src_evt})")
    
    design_data_path = os.path.join(proj_root, "design", "DesignData", WORKSPACE_NAME, "design")
    file_path = os.path.join(design_data_path, "item.xlsx")
    
    if not os.path.exists(file_path):
        log(f"item.xlsx not found at {file_path}")
        return False
        
    try:
        import openpyxl
        # Backup
        bak_path = file_path + ".bak"
        if not os.path.exists(bak_path):
            shutil.copy2(file_path, bak_path)
            record_file_creation(log_dir, step_index, bak_path)
        
        # Make writable
        import stat
        os.chmod(file_path, stat.S_IWRITE | stat.S_IREAD)
        
        wb = load_workbook_frozen(file_path)
        modified = False
        
        # 1. Prepare name swap strings
        src_upper = src_evt.upper()
        tgt_upper = tgt_evt.upper()
        src_lower = src_evt.lower()
        tgt_lower = tgt_evt.lower()
        src_pascal = src_evt[0].upper() + src_evt[1:] if src_evt else ""
        tgt_pascal = tgt_evt[0].upper() + tgt_evt[1:] if tgt_evt else ""
        
        # Extract digits from src/tgt names (e.g. 2507, 2605)
        src_digits = "".join(filter(str.isdigit, src_evt))
        tgt_digits = "".join(filter(str.isdigit, tgt_evt))
        
        def swap_name(value):
            if value is None: return None
            s = str(value)
            if src_evt:
                s = s.replace(src_upper, tgt_upper)
                s = s.replace(src_lower, tgt_lower)
                s = s.replace(src_pascal, tgt_pascal)
                s = s.replace(src_evt, tgt_evt)
            if src_digits and tgt_digits:
                s = s.replace(src_digits, tgt_digits)
            return s

        # ========== Sheet: item ==========
        if "item" in wb.sheetnames:
            ws = wb["item"]
            real_max = get_real_max_row(ws)
            
            # Find source items
            src_rows = []
            for r in range(2, real_max + 1):
                alias_val = str(ws.cell(r, 2).value or "").lower()
                name_val = str(ws.cell(r, 5).value or "").lower()
                
                match = False
                if src_evt.lower() in alias_val or src_evt.lower() in name_val:
                    match = True
                
                if match:
                    src_rows.append(r)
            
            exists, _ = check_and_handle_existing(ws, "item", tgt_evt)
            if exists:
                log(f"item_sheet: Target items already exist for {tgt_evt}. Reconstructing id_map...")
                # Match by normalized suffix (event name/digits stripped from both sides), not raw
                # substring containment — "Merge2606能量" is never a substring of "Merge2608能量"
                # since the digits differ, so a plain `src_alias in tgt_alias` check never matches
                # any cross-event pair. Stripping the event name from both first ("能量" == "能量")
                # is what actually makes the match event-independent.
                def _strip_event_name(value, evt_name):
                    if not value:
                        return ""
                    s = str(value)
                    for variant in [evt_name.upper(), evt_name.lower(), (evt_name[0].upper() + evt_name[1:]) if evt_name else evt_name]:
                        if variant:
                            s = s.replace(variant, "")
                    return s.strip()

                id_map = {}
                for sr in src_rows:
                    src_id = str(ws.cell(sr, 1).value or "").replace("//", "").strip()
                    src_name_suffix = _strip_event_name(ws.cell(sr, 5).value, src_evt)
                    src_alias_suffix = _strip_event_name(ws.cell(sr, 2).value, src_evt)

                    for tr in range(2, real_max + 1):
                        tgt_id = str(ws.cell(tr, 1).value or "").replace("//", "").strip()
                        if tgt_id == src_id: continue # Same ID means it's likely the same row or a duplicate

                        tgt_name_suffix = _strip_event_name(ws.cell(tr, 5).value, tgt_evt)
                        tgt_alias_suffix = _strip_event_name(ws.cell(tr, 2).value, tgt_evt)

                        # Match by normalized suffix (exact, not substring — avoids "宝箱" matching "金币宝箱")
                        if (src_name_suffix and src_name_suffix == tgt_name_suffix) or \
                           (src_alias_suffix and src_alias_suffix == tgt_alias_suffix):
                            if src_id.isdigit() and tgt_id.isdigit():
                                id_map[int(src_id)] = int(tgt_id)
                                break
                
                if id_map:
                    log(f"item_sheet: Reconstructed id_map with {len(id_map)} entries.")
                    return True, id_map
                else:
                    log(f"item_sheet: Could not reconstruct id_map. Returning True without map.")
                    return True, {}
            
            if src_rows:
                log(f"item_sheet: Found {len(src_rows)} items for {src_evt}. Cloning...")
                
                # Group source items by their ID range and find current max ID in the sheet for those ranges
                max_ids_by_range = {}
                for r in range(2, real_max + 1):
                    val = ws.cell(r, 1).value
                    if val:
                        s_val = str(val).strip()
                        if s_val.startswith("//"): s_val = s_val[2:].strip()
                        if s_val.isdigit():
                            num = int(s_val)
                            range_key = num // 1000
                            if range_key not in max_ids_by_range or num > max_ids_by_range[range_key]:
                                max_ids_by_range[range_key] = num
                
                id_map = {}
                new_rows_data = []
                for src_r in src_rows:
                    row_data = []
                    src_cells = []
                    for c in range(1, ws.max_column + 1):
                        row_data.append(ws.cell(src_r, c).value)
                        src_cells.append(ws.cell(src_r, c))

                    old_id_str = str(row_data[0]).strip()
                    if old_id_str.startswith("//"): old_id_str = old_id_str[2:].strip()

                    new_id = -1
                    if old_id_str.isdigit():
                        old_id = int(old_id_str)
                        range_key = old_id // 1000
                        new_id = max_ids_by_range.get(range_key, old_id) + 1
                        max_ids_by_range[range_key] = new_id
                        row_data[0] = new_id
                        id_map[old_id] = new_id
                        log(f"item_sheet: Map ID {old_id} -> {new_id}")
                    else:
                        row_data[0] = swap_name(row_data[0])
                        new_id = 99999999 # Fallback

                    for i in range(1, len(row_data)):
                        row_data[i] = swap_name(row_data[i])

                    new_rows_data.append((new_id, row_data, src_cells))

                new_rows_data.sort(key=lambda x: x[0], reverse=True)
                for new_id, row_data, src_cells in new_rows_data:
                    best_row = 1
                    max_id_seen = -1
                    for r in range(2, ws.max_row + 1):
                        val = ws.cell(r, 1).value
                        if val is not None:
                            val_str = str(val).strip()
                            if val_str.startswith("//"): val_str = val_str[2:].strip()
                            if val_str.isdigit():
                                current_id = int(val_str)
                                if current_id < new_id and current_id > max_id_seen:
                                    max_id_seen = current_id
                                    best_row = r

                    insert_pos = best_row + 1
                    ws.insert_rows(insert_pos)
                    for c_idx, val in enumerate(row_data):
                        dst_c = ws.cell(insert_pos, c_idx + 1)
                        dst_c.value = val
                        copy_cell_format(src_cells[c_idx], dst_c)
                    log(f"item_sheet: Inserted item {new_id} at row {insert_pos}")
                    
                modified = True
            else:
                log(f"item_sheet: No source items found for {src_evt}.")
                id_map = {}
        
        if modified:
            wb.save(file_path)
            log(f"item.xlsx updated for {tgt_evt}.")
            return True, id_map
        else:
            log(f"item.xlsx: No changes were needed.")
            return True, {}
            
    except Exception as e:
        log(f"Error updating item.xlsx: {e}")
        import traceback
        traceback.print_exc()
        return False, {}


def sync_events_excel_item_ids(proj_root, tgt_evt, id_map, log_dir, step_index):
    """Update events.xlsx with the new '掉落的道具ID' from item.xlsx sync."""
    import openpyxl
    import os
    import stat
    
    events_path = os.path.join(proj_root, "design", "DesignData", WORKSPACE_NAME, "design", "events.xlsx")
    if not os.path.exists(events_path):
        log(f"Warning: events.xlsx not found for ID sync at {events_path}")
        return True
        
    try:
        os.chmod(events_path, stat.S_IWRITE | stat.S_IREAD)
        wb = load_workbook_frozen(events_path)
        if "events" not in wb.sheetnames:
            log("Warning: 'events' sheet not found in events.xlsx")
            return True
            
        ws = wb["events"]
        
        # Find Alias and 掉落的道具ID columns
        alias_col = -1
        drop_id_col = -1
        for c in range(1, ws.max_column + 1):
            h = str(ws.cell(1, c).value or "").lower()
            if "alias" in h: alias_col = c
            if "掉落的道具id" in h: drop_id_col = c
            
        if alias_col == -1 or drop_id_col == -1:
            log(f"Warning: Could not find columns in events.xlsx (Alias={alias_col}, DropID={drop_id_col})")
            return True
            
        modified = False
        tgt_lower = tgt_evt.lower()
        for r in range(2, ws.max_row + 1):
            alias_val = str(ws.cell(r, alias_col).value or "").lower()
            if tgt_lower in alias_val:
                current_id_v = ws.cell(r, drop_id_col).value
                if current_id_v:
                    try:
                        cid = int(float(str(current_id_v)))
                        if cid in id_map:
                            new_id = id_map[cid]
                            ws.cell(r, drop_id_col, new_id)
                            log(f"events.xlsx: Updated drop ID for {tgt_evt}: {cid} -> {new_id}")
                            modified = True
                    except (ValueError, TypeError):
                        pass
        
        if modified:
            wb.save(events_path)
            log(f"events.xlsx successfully synchronized with new item IDs for {tgt_evt}.")
        return True
    except Exception as e:
        log(f"Error syncing events.xlsx: {e}")
        return False
            
    except Exception as e:
        log(f"Error updating item.xlsx: {e}")
        traceback.print_exc()
        return False

def update_event_shop_excel(proj_root, src_evt, tgt_evt, root_path, step_index):
    """Step 11: Update event_shop.xlsx with target shop and items."""
    import openpyxl
    import stat
    
    design_dir = os.path.join(proj_root, "design", "DesignData", WORKSPACE_NAME, "design")
    shop_path = os.path.join(design_dir, "event_shop.xlsx")
    
    if not os.path.exists(shop_path):
        log(f"Error: event_shop.xlsx not found at {shop_path}")
        return False
    
    extra = load_extra_args(root_path)
    event_id = extra.get("event_id", "")
    log(f"event_shop: Received extra args: {extra}")
    
    src_lower = src_evt.lower()
    tgt_lower = tgt_evt.lower()
    src_pascal = src_evt[0].upper() + src_evt[1:] if src_evt else src_evt
    tgt_pascal = tgt_evt[0].upper() + tgt_evt[1:] if tgt_evt else tgt_evt
    src_upper = src_evt.upper()
    tgt_upper = tgt_evt.upper()

    def swap_name(value):
        if value is None: return None
        s = str(value)
        if src_evt:
            s = s.replace(src_upper, tgt_upper)
            s = s.replace(src_lower, tgt_lower)
            s = s.replace(src_pascal, tgt_pascal)
            s = s.replace(src_evt, tgt_evt)
        # Fallback for ID-like columns (e.g. 20250701 -> 20260501)
        src_digits = "".join(filter(str.isdigit, src_evt))
        tgt_digits = "".join(filter(str.isdigit, tgt_evt))
        if len(src_digits) >= 4 and len(tgt_digits) >= 4:
            s = s.replace("20" + src_digits[:4], "20" + tgt_digits[:4])
        return s

    def get_id_value(val):
        """Extract numeric ID from potentially commented string '//100001'"""
        if val is None: return None
        s = str(val).strip()
        if s.startswith("//"):
            s = s[2:].strip()
        try:
            return int(s)
        except ValueError:
            return None

    try:
        # Backup
        backup_path = shop_path + ".bak"
        if os.path.exists(backup_path):
            os.chmod(backup_path, stat.S_IWRITE)
            os.remove(backup_path)
        shutil.copy2(shop_path, backup_path)
        record_file_creation(root_path, step_index, backup_path)
        
        wb = load_workbook_frozen(shop_path)
        modified = False
        new_item_ids = []  # Referenced by Sheet 2 below even if 物品 is missing or has no rows for this event.

        # ========== Sheet 1: 物品 (Items) ==========
        item_sheet = "物品"
        if item_sheet in wb.sheetnames:
            ws_item = wb[item_sheet]
            real_max = get_real_max_row(ws_item)
            
            # 1. Find the highest ID and determine new group start
            max_id = 0
            for r in range(2, real_max + 1):
                id_val = get_id_value(ws_item.cell(r, 1).value)
                if id_val and id_val > max_id:
                    max_id = id_val
            
            new_group_start = (max_id // 100 + 1) * 100 + 1
            log(f"item_sheet: Last max ID was {max_id}. New group starts at {new_group_start}.")
            
            # 2. Check if target items already exist
            target_item_rows = []
            for r in range(2, real_max + 1):
                alias_val = str(ws_item.cell(r, 2).value or "")
                if tgt_lower in alias_val.lower():
                    target_item_rows.append(r)
            
            if target_item_rows:
                log(f"item_sheet: Found {len(target_item_rows)} existing items for {tgt_evt}. Checking comment status...")
                for tr in target_item_rows:
                    val = ws_item.cell(tr, 1).value
                    if val and str(val).strip().startswith("//"):
                        ws_item.cell(tr, 1, str(val).strip()[2:].strip())
                        log(f"item_sheet: Uncommented row {tr} for {tgt_evt}.")
                        modified = True
                
                # If target items exist, we skip cloning
                src_item_rows = []
                # Also need to collect these IDs for the shop row later
                new_item_ids = []
                for tr in target_item_rows:
                    id_val = get_id_value(ws_item.cell(tr, 1).value)
                    if id_val: new_item_ids.append(id_val)
            else:
                # 2b. Find source items (those linked to src_evt)
                src_item_rows = []
                new_item_ids = []
                for r in range(2, real_max + 1):
                    alias_val = str(ws_item.cell(r, 2).value or "")
                    if src_lower in alias_val.lower():
                        src_item_rows.append(r)

                if not src_item_rows:
                    log(f"item_sheet: No items found for source {src_evt}.")
            
            # 3. Clone source items
            # (new_item_ids is already set above — from the target's own existing items if
            # target_item_rows was non-empty, or stays [] here and gets filled by the clone below.)
            if src_item_rows:
                current_id = new_group_start
                insert_pos = real_max + 1
                for src_r in src_item_rows:
                    src_cells = [ws_item.cell(src_r, c) for c in range(1, ws_item.max_column + 1)]
                    row_data = [c.value for c in src_cells]
                    # Update ID (Column 1)
                    row_data[0] = current_id
                    # Swap names in ALL other columns (Column 2 onwards)
                    for i in range(1, len(row_data)):
                        if row_data[i] is not None:
                            row_data[i] = swap_name(row_data[i])

                    # Write to new row
                    for c_idx, val in enumerate(row_data):
                        dst_c = ws_item.cell(insert_pos, c_idx + 1)
                        dst_c.value = val
                        copy_cell_format(src_cells[c_idx], dst_c)

                    new_item_ids.append(current_id)
                    log(f"item_sheet: Added link item {current_id} (cloned from {src_r})")
                    current_id += 1
                    insert_pos += 1
                modified = True
            
            # 4. Commenting logic: only last two groups uncommented
            # A group is defined by its ID // 100 segment
            real_max = get_real_max_row(ws_item) # Refresh after adding
            # Find all group segments
            segments = []
            for r in range(2, real_max + 1):
                id_v = get_id_value(ws_item.cell(r, 1).value)
                if id_v:
                    seg = id_v // 100
                    if not segments or segments[-1] != seg:
                        segments.append(seg)
            
            if len(segments) > 2:
                # The last two groups, plus the target's own: a reopen under the same name keeps
                # its items in their original (older, mid-sheet) group, which "last two" alone
                # would comment straight back out right after uncommenting it (Merge2510: 1035).
                keep_segments = list(dict.fromkeys(segments[-2:] + [i // 100 for i in new_item_ids if isinstance(i, int)]))
                # ...and whatever an ACTIVE shop row still sells. WorkSpcae3 keeps every recent
                # event's shop live (13 groups, Blocks2604..Blocks2610); commenting their items left
                # e.g. Merge2608's shop pointing at 105501-105504, and c_data's reference check then
                # aborts the whole coder_convert ("value:105501 not exist in event_shop_sale_item").
                sold = set()
                if "event_shop" in wb.sheetnames:
                    ws_s = wb["event_shop"]
                    for sr in range(2, ws_s.max_row + 1):
                        if str(ws_s.cell(sr, 1).value or "").strip().startswith("//"):
                            continue
                        for sc in range(1, ws_s.max_column + 1):
                            sold.update(int(x) for x in re.findall(r'(?<!\d)\d{6}(?!\d)', str(ws_s.cell(sr, sc).value or "")))
                log(f"item_sheet: Keeping groups {keep_segments} active, plus any item an active shop still sells. Commenting others.")
                for r in range(2, real_max + 1):
                    val = ws_item.cell(r, 1).value
                    if val:
                        id_v = get_id_value(val)
                        if id_v and (id_v // 100) not in keep_segments and id_v not in sold:
                            s_val = str(val).strip()
                            if not s_val.startswith("//"):
                                ws_item.cell(r, 1, "//" + s_val)
                        elif id_v and ((id_v // 100) in keep_segments or id_v in sold):
                            s_val = str(val).strip()
                            if s_val.startswith("//"):
                                ws_item.cell(r, 1, s_val[2:].strip())
                modified = True

        # ========== Sheet 2: event_shop ==========
        shop_sheet = "event_shop"
        if shop_sheet in wb.sheetnames:
            ws_shop = wb[shop_sheet]
            real_max = get_real_max_row(ws_shop)
            
            # Check if target shop already exists
            exists, target_row_idx = check_and_handle_existing(ws_shop, "event_shop", tgt_evt)
            if exists:
                log(f"event_shop: Target shop already exists at row {target_row_idx}. Updating...")
                # Update Shop ID
                if event_id:
                    ws_shop.cell(target_row_idx, 1, int(event_id))
                # Update Item List
                if new_item_ids:
                    ws_shop.cell(target_row_idx, 4, ",".join(map(str, new_item_ids)))
                modified = True
                src_row = None # Skip cloning logic
            else:
                # Find source row
                src_row = None
                for r in range(2, real_max + 1):
                    val = str(ws_shop.cell(r, 2).value or "")
                    if src_lower in val.lower():
                        src_row = r
                        break
            
            if src_row:
                new_row = real_max + 1
                for c in range(1, ws_shop.max_column + 1):
                    src_c = ws_shop.cell(src_row, c)
                    old_val = src_c.value
                    dst_c = ws_shop.cell(new_row, c)
                    dst_c.value = swap_name(old_val)
                    copy_cell_format(src_c, dst_c)
                
                # Update Shop ID
                if event_id:
                    ws_shop.cell(new_row, 1, int(event_id))
                
                # Update Item List (Col 4)
                if new_item_ids:
                    ws_shop.cell(new_row, 4, ",".join(map(str, new_item_ids)))
                
                log(f"event_shop: Added shop row {new_row} for {tgt_evt}.")
                
                # Commenting logic: only last two rows uncommented
                # After adding, new_row is the latest
                for r in range(2, real_max + 2):
                    val = ws_shop.cell(r, 1).value
                    if val:
                        s_val = str(val).strip()
                        if r < (real_max + 2 - 2): # All but last two
                            if not s_val.startswith("//"):
                                ws_shop.cell(r, 1, "//" + s_val)
                        else: # Last two
                            if s_val.startswith("//"):
                                ws_shop.cell(r, 1, s_val[2:].strip())
                
                modified = True
            else:
                log(f"event_shop: Source event {src_evt} not found. Skipping shop creation.")

        if modified:
            os.chmod(shop_path, stat.S_IWRITE | stat.S_IREAD)
            wb.save(shop_path)
            log("event_shop.xlsx saved successfully.")
        else:
            log("event_shop.xlsx: No changes were needed.")
            
        return True
    except Exception as e:
        log(f"Exception updating event_shop.xlsx: {e}")
        traceback.print_exc()
        return False

def update_icon_excel(proj_root, src_evt, tgt_evt, log_dir, step_index):
    """Step 13: Update icon.xlsx and intelligently auto-generate missing icon hooks."""
    import openpyxl
    import stat
    import os
    
    design_dir = os.path.join(proj_root, "design", "DesignData", WORKSPACE_NAME, "design")
    icon_path = os.path.join(design_dir, "icon.xlsx")
    item_path = os.path.join(design_dir, "item.xlsx")
    
    if not os.path.exists(icon_path):
        log(f"Error: icon.xlsx not found at {icon_path}")
        return False
        
    src_lower = src_evt.lower()
    tgt_lower = tgt_evt.lower()
    src_pascal = src_evt[0].upper() + src_evt[1:] if src_evt else src_evt
    tgt_pascal = tgt_evt[0].upper() + tgt_evt[1:] if tgt_evt else tgt_evt
    src_upper = src_evt.upper()
    tgt_upper = tgt_evt.upper()

    def swap_name(value):
        if value is None: return None
        s = str(value)
        if src_evt:
            s = s.replace(src_upper, tgt_upper)
            s = s.replace(src_lower, tgt_lower)
            s = s.replace(src_pascal, tgt_pascal)
            s = s.replace(src_evt, tgt_evt)
        return s

    # 1. First, gather REQUIRED target icons from item.xlsx AND the target event excel
    required_tgt_icons = set()
    try:
        # A. From item.xlsx
        if os.path.exists(item_path):
            wb_item = openpyxl.load_workbook(item_path, data_only=True)
            if "item" in wb_item.sheetnames:
                ws_item = wb_item["item"]
                for r in range(2, ws_item.max_row + 1):
                    alias = str(ws_item.cell(r, 2).value or "")
                    name = str(ws_item.cell(r, 5).value or "")
                    if tgt_evt.lower() in alias.lower() or tgt_evt.lower() in name.lower():
                        icon7 = str(ws_item.cell(r, 7).value or "").strip()
                        icon8 = str(ws_item.cell(r, 8).value or "").strip()
                        icon9 = str(ws_item.cell(r, 9).value or "").strip()
                        if icon7: required_tgt_icons.add(icon7)
                        if icon8: required_tgt_icons.add(icon8)
                        if icon9: required_tgt_icons.add(icon9)
            wb_item.close()
            
        # B. From tgt_evt excel (e.g. sim2605.xlsx)
        tgt_evt_path = os.path.join(design_dir, f"{tgt_evt}.xlsx")
        if not os.path.exists(tgt_evt_path):
            # Try lowercase
            tgt_evt_path = os.path.join(design_dir, f"{tgt_evt.lower()}.xlsx")
            
        if os.path.exists(tgt_evt_path):
            wb_tgt_evt = openpyxl.load_workbook(tgt_evt_path, data_only=True)
            for ws_t in wb_tgt_evt.worksheets:
                for row in ws_t.iter_rows():
                    for cell in row:
                        val = str(cell.value or "")
                        # Match icons (e.g. SIM2605_DISH_ICON_1)
                        if "_ICON" in val.upper() and tgt_upper in val.upper():
                            # Clean up (sometimes values have suffixes or are part of strings)
                            match = re.search(rf'({tgt_upper}_[A-Za-z0-9_]+_ICON_[0-9]+)', val, re.IGNORECASE)
                            if match:
                                required_tgt_icons.add(match.group(1).upper())
                            elif "_" in val:
                                 # Fallback: if it has _ICON, add it if it looks like an identifier
                                 if re.match(r'^[A-Za-z0-9_]+$', val):
                                     required_tgt_icons.add(val.strip().upper())
            wb_tgt_evt.close()
            log(f"icon_sheet: Identified {len(required_tgt_icons)} required icons from item.xlsx and {tgt_evt}.xlsx")
            
    except Exception as e:
        log(f"icon_sheet: Failed to gather required icons for cross-validation: {e}")

    try:
        os.chmod(icon_path, stat.S_IWRITE | stat.S_IREAD)
        wb_icon = load_workbook_frozen(icon_path)
        
        modified = False
        gui_event_id = None
        try:
            extra_args_path = os.path.join(log_dir, "extra_args.json")
            if os.path.exists(extra_args_path):
                with open(extra_args_path, "r") as f:
                    extra = json.load(f)
                    gui_event_id = str(extra.get("event_id", ""))
        except: pass

        extra_for_hint = load_extra_args(log_dir)
        src_start_time = extra_for_hint.get("src_start_time", "")

        # 2a. icon_item's "活动物品" group — a variable-length run of rows (2, 3, however
        # many an event needs) bounded by dividers, NOT fixed row indices. Resolve the
        # source event's own rows via live file -> cache -> SVN history, then append
        # (never overwrite) at the end of the group, right before the next divider.
        if "icon_item" in wb_icon.sheetnames:
            ws_item = wb_icon["icon_item"]
            start, end = find_icon_item_group_bounds(ws_item)

            if start is None:
                log("icon_item: '活动物品' group marker not found — skipping icon_item update.")
            else:
                tgt_already_present = any(
                    tgt_lower in str(ws_item.cell(r, 1).value or "").lower()
                    for r in range(start, end)
                )
                if tgt_already_present:
                    log(f"icon_item: {tgt_evt} already has row(s) in the 活动物品 group. Skipping.")
                else:
                    src_item_rows, src_item_label = get_icon_item_rows_for_source(icon_path, src_lower, src_start_time, log_dir)
                    if src_item_rows:
                        log(f"icon_item: Resolved {len(src_item_rows)} row(s) for {src_evt} from {src_item_label}.")
                        insert_at = end
                        for row_vals in src_item_rows:
                            new_row = []
                            for c_idx, v in enumerate(row_vals):
                                nv = swap_name(v)
                                if c_idx == 2 and nv and gui_event_id:  # column 3: sprite name
                                    nv = re.sub(r'\d{8}', gui_event_id, str(nv))
                                new_row.append(nv)
                            ws_item.insert_rows(insert_at)
                            for c_idx, v in enumerate(new_row):
                                ws_item.cell(insert_at, c_idx + 1, v)
                            insert_at += 1
                        log(f"icon_item: Appended {len(src_item_rows)} row(s) for {tgt_evt} to the 活动物品 group.")
                        modified = True
                    else:
                        log(f"icon_item: WARNING — no data found for {src_evt} anywhere (not live, not cached, not in SVN history). Needs manual entry.")

        # 2b. icon_event — resolve source rows via icon_备份 (free) -> cache -> SVN
        # history of icon_event itself, then clone into the active sheet.
        backup_sheet_name = "icon_备份"
        active_sheet_name = "icon_event"

        if active_sheet_name in wb_icon.sheetnames:
            ws_active = wb_icon[active_sheet_name]

            existing_active_icons = {} # name -> row_idx
            for r in range(2, ws_active.max_row + 1):
                icon_val = str(ws_active.cell(r, 1).value or "").strip()
                if icon_val:
                    if icon_val.startswith("//"):
                        name = icon_val[2:].strip().upper()
                        existing_active_icons[name] = r
                    else:
                        existing_active_icons[icon_val.upper()] = r

            src_event_rows, src_event_label = get_icon_event_rows_for_source(icon_path, src_lower, src_start_time, log_dir)

            if src_event_rows:
                log(f"icon_event: Resolved {len(src_event_rows)} row(s) for {src_evt} from {src_event_label}.")
                divider_row = ws_active.max_row + 1
                ws_active.cell(divider_row, 1, f"// {tgt_upper} (CLONED FROM {src_upper})")

                cloned_count = 0
                for row_vals in src_event_rows:
                    new_icon_name = swap_name(row_vals[0])
                    name_upper = str(new_icon_name).upper().strip()
                    if new_icon_name and name_upper in existing_active_icons:
                        row_idx = existing_active_icons[name_upper]
                        val = str(ws_active.cell(row_idx, 1).value or "").strip()
                        if val.startswith("//"):
                            ws_active.cell(row_idx, 1, val[2:].strip())
                            log(f"icon_sheet: Uncommented existing icon {new_icon_name} in active event sheet.")
                            modified = True
                        continue

                    new_row_idx = ws_active.max_row + 1
                    for c_idx, old_val in enumerate(row_vals):
                        new_val = swap_name(old_val)
                        # Sync Event ID in Sprite Name (Col 3) as requested
                        if c_idx == 2 and new_val and gui_event_id:
                            new_val = re.sub(r'\d{8}', gui_event_id, str(new_val))
                        ws_active.cell(new_row_idx, c_idx + 1, new_val)
                    cloned_count += 1

                if cloned_count > 0:
                    log(f"icon_sheet: Cloned {cloned_count} rows for {tgt_evt} into active event sheet.")
                    modified = True
            else:
                log(f"icon_event: WARNING — no data found for {src_evt} anywhere (icon_备份, icon_event, cache, or SVN history). Needs manual entry.")

        # 3. Traditional sheet-by-sheet update
        for sheet_name in wb_icon.sheetnames:
            if sheet_name in [backup_sheet_name, active_sheet_name, "icon_item"]:
                continue
                
            ws = wb_icon[sheet_name]
            src_rows = []
            for r in range(2, ws.max_row + 1):
                c1 = str(ws.cell(r, 1).value or "").lower()
                c2 = str(ws.cell(r, 2).value or "").lower()
                c3 = str(ws.cell(r, 3).value or "").lower()
                if src_lower in c1 or src_lower in c2 or src_lower in c3:
                    src_rows.append(r)
            
            if src_rows:
                existing_in_sheet = {} # name -> row_idx
                for r in range(2, ws.max_row + 1):
                    v = str(ws.cell(r, 1).value or "").strip()
                    if v:
                        if v.startswith("//"):
                            name = v[2:].strip().upper()
                            existing_in_sheet[name] = r
                        else:
                            existing_in_sheet[v.upper()] = r

                divider_row = ws.max_row + 1
                ws.cell(divider_row, 1, f"// {tgt_upper}")
                
                cloned_in_sheet = 0
                for r_src in src_rows:
                    new_key = swap_name(ws.cell(r_src, 1).value)
                    name_upper = str(new_key).strip().upper()
                    if new_key and name_upper in existing_in_sheet:
                        row_idx = existing_in_sheet[name_upper]
                        val = str(ws.cell(row_idx, 1).value or "").strip()
                        if val.startswith("//"):
                            ws.cell(row_idx, 1, val[2:].strip())
                            log(f"icon_sheet: Uncommented existing icon {new_key} in sheet {sheet_name}.")
                            modified = True
                        continue
                    new_row_idx = ws.max_row + 1
                    for c in range(1, ws.max_column + 1):
                        src_c = ws.cell(r_src, c)
                        old_v = src_c.value
                        new_v = swap_name(old_v)
                        # Sync Event ID in Sprite Name (Col 3)
                        if c == 3 and new_v and gui_event_id:
                            new_v = re.sub(r'\d{8}', gui_event_id, str(new_v))
                        dst_c = ws.cell(new_row_idx, c)
                        dst_c.value = new_v
                        copy_cell_format(src_c, dst_c)
                    cloned_in_sheet += 1
                
                if cloned_in_sheet > 0:
                    log(f"icon_sheet: Cloned {cloned_in_sheet} rows in [{sheet_name}] for {tgt_evt}.")
                    modified = True

        # 4. Intelligent Missing Icon Generation
        all_final_icons = set()
        for sheet_name in wb_icon.sheetnames:
            if "备份" in sheet_name: continue
            ws_f = wb_icon[sheet_name]
            for r in range(2, ws_f.max_row + 1):
                # item.xlsx's 物品ICON columns name icons by ALIAS (col 2, e.g. "MERGE2510代币"), not
                # key — checking keys alone flagged all 17 of Merge2510's (all present) as missing art.
                for c in (1, 2):
                    v = str(ws_f.cell(r, c).value or "").strip().upper()
                    if v: all_final_icons.add(v)
                
        missing_icons = []
        for req_icon in required_tgt_icons:
            if req_icon.upper() not in all_final_icons:
                missing_icons.append(req_icon)

        if missing_icons:
            ws_target_fix = wb_icon["icon_event"] if "icon_event" in wb_icon.sheetnames else wb_icon.worksheets[0]
            source_to_clone = {}
            for sheet_name in wb_icon.sheetnames:
                ws_search = wb_icon[sheet_name]
                for r in range(2, ws_search.max_row + 1):
                    s1 = str(ws_search.cell(r, 1).value or "").strip().upper()
                    if s1 and s1 not in source_to_clone:
                        source_to_clone[s1] = (ws_search, r)

            recovered = []
            still_missing = []

            for m_icon in missing_icons:
                src_icon_name = m_icon.replace(tgt_upper, src_upper).replace(tgt_lower, src_lower).replace(tgt_pascal, src_pascal)
                if src_icon_name.upper() in source_to_clone:
                    ws_src, r_idx = source_to_clone[src_icon_name.upper()]
                    new_r = ws_target_fix.max_row + 1
                    for c in range(1, ws_src.max_column + 1):
                        src_c = ws_src.cell(r_idx, c)
                        old_v = src_c.value
                        new_v = swap_name(old_v)
                        if c == 3 and new_v and gui_event_id:
                            new_v = re.sub(r'\d{8}', gui_event_id, str(new_v))
                        dst_c = ws_target_fix.cell(new_r, c)
                        dst_c.value = new_v
                        copy_cell_format(src_c, dst_c)
                    log(f"icon_sheet: Successfully RECOVERED icon '{m_icon}' from {ws_src.title}")
                    recovered.append(m_icon)
                else:
                    # No fabrication — a fake placeholder (Atlas_Quest/_TBD) would silently
                    # ship broken art with no visible signal beyond a log line. This icon
                    # genuinely doesn't exist anywhere yet; that's real missing art, not
                    # something this step can invent correctly.
                    still_missing.append(m_icon)

            if recovered:
                modified = True
            if still_missing:
                log(f"icon_sheet: WARNING — {len(still_missing)} icon(s) still need real art, not created anywhere: {still_missing}")
        
        if modified:
            wb_icon.save(icon_path)
            log(f"icon.xlsx successfully updated for {tgt_evt}.")
            return True
        else:
            log(f"icon.xlsx: No changes were needed.")
            return True
    
    except Exception as e:
        log(f"Error updating icon.xlsx: {e}")
        import traceback
        traceback.print_exc()
        return False

def get_quiz_date_prefix(proj_root, event_name):
    """Derive quiz date prefix (YYMMDD) from events.xlsx start time for a given event."""
    import openpyxl
    import os
    
    events_path = os.path.join(proj_root, "design", "DesignData", WORKSPACE_NAME, "design", "events.xlsx")
    if not os.path.exists(events_path):
        log(f"Warning: events.xlsx not found at {events_path}")
        return None
    
    wb = openpyxl.load_workbook(events_path, data_only=True)
    ws = wb["events"]
    
    # Find Alias column and 开始时间 column
    alias_col = None
    start_time_col = None
    for c in range(1, ws.max_column + 1):
        header = str(ws.cell(1, c).value or "")
        if "Alias" in header or "alias" in header.lower():
            alias_col = c
        if "开始时间" in header:
            start_time_col = c
    
    if not alias_col or not start_time_col:
        log(f"Warning: Could not find Alias or 开始时间 column in events.xlsx")
        return None
    
    evt_lower = event_name.lower()
    for r in range(2, ws.max_row + 1):
        alias_val = str(ws.cell(r, alias_col).value or "").lower()  # type: ignore
        if evt_lower in alias_val:
            start_val = ws.cell(r, start_time_col).value  # type: ignore
            if start_val:
                # Could be datetime object or string like "2025-07-17 00:00:00"
                from datetime import datetime
                if isinstance(start_val, datetime):
                    return start_val.strftime("%y%m%d")
                else:
                    s = str(start_val).strip()
                    # Parse "2025-07-17 00:00:00" or similar
                    try:
                        dt = datetime.strptime(s[:10], "%Y-%m-%d")
                        return dt.strftime("%y%m%d")
                    except Exception:
                        log(f"Warning: Could not parse start time '{s}' for {event_name}")
                        return None
    
    log(f"Warning: Could not find event '{event_name}' in events.xlsx")
    return None


def update_localization_main_excel(proj_root, src_evt, tgt_evt, log_dir, step_index):
    """Step 14 Part A: Update localization.xlsx (main localization file only)."""
    import openpyxl
    import stat
    import os
    import re
    
    design_dir = os.path.join(proj_root, "design", "DesignData", WORKSPACE_NAME, "design")
    file_path = os.path.join(design_dir, "localization.xlsx")
    
    if not os.path.exists(file_path):
        log(f"Warning: localization.xlsx not found. Skipping.")
        return True
    
    src_lower = src_evt.lower()
    tgt_lower = tgt_evt.lower()
    src_pascal = src_evt[0].upper() + src_evt[1:] if src_evt else src_evt
    tgt_pascal = tgt_evt[0].upper() + tgt_evt[1:] if tgt_evt else tgt_evt
    src_upper = src_evt.upper()
    tgt_upper = tgt_evt.upper()

    src_num_match = re.search(r'\d+', src_evt)
    tgt_num_match = re.search(r'\d+', tgt_evt)
    src_digits = src_num_match.group() if src_num_match else ""
    tgt_digits = tgt_num_match.group() if tgt_num_match else ""

    # 1. First, gather REQUIRED target keys from the target event excel
    required_tgt_keys = set()
    try:
        tgt_evt_path = os.path.join(design_dir, f"{tgt_evt}.xlsx")
        if not os.path.exists(tgt_evt_path):
            tgt_evt_path = os.path.join(design_dir, f"{tgt_evt.lower()}.xlsx")
            
        if os.path.exists(tgt_evt_path):
            wb_tgt_evt = openpyxl.load_workbook(tgt_evt_path, data_only=True)
            for ws_t in wb_tgt_evt.worksheets:
                for row in ws_t.iter_rows():
                    for cell in row:
                        val = str(cell.value or "")
                        # Match localization keys: e.g. SIM2605_ROOMNAME_1 or LOC_SIM2605_...
                        # Usually uppercase identifiers starting with SIM or having the event digits
                        if (tgt_upper in val.upper() or (tgt_digits and tgt_digits in val)) and "_" in val:
                            match = re.search(rf'([A-Z0-9_]*{tgt_upper}[A-Z0-9_]*)', val.upper())
                            if match:
                                required_tgt_keys.add(match.group(1))
                            elif re.match(r'^[A-Z0-9_]+$', val.upper()):
                                required_tgt_keys.add(val.upper().strip())
            wb_tgt_evt.close()
            log(f"localization: Identified {len(required_tgt_keys)} required keys from {tgt_evt}.xlsx")
            
    except Exception as e:
        log(f"localization: Failed to gather required keys from {tgt_evt}.xlsx: {e}")

    def swap_name(value):
        if value is None: return None
        s = str(value)  # type: ignore
        if src_evt:
            s = s.replace(src_upper, tgt_upper)
            s = s.replace(src_lower, tgt_lower)
            s = s.replace(src_pascal, tgt_pascal)
            s = s.replace(src_evt, tgt_evt)
        if src_digits and tgt_digits:
            s = s.replace(src_digits, tgt_digits)
        return s

    try:
        os.chmod(file_path, stat.S_IWRITE | stat.S_IREAD)
        wb = load_workbook_frozen(file_path)
        
        modified = False
        extra_for_hint = load_extra_args(log_dir)
        src_start_time = extra_for_hint.get("src_start_time", "")

        # 1. Process each sheet efficiently — excludes the three sheets with their own
        # dedicated handling below (localize_event_fixed's divider group, localize_event's
        # backup-first lookup, localize_server's accumulator scan) and the backup sheet
        # itself (nothing should write into an archive via the generic pass).
        all_final_keys = set()
        for sheet_name in wb.sheetnames:
            if sheet_name in ["languages", "country", "localize_零使用",
                               "localize_event_fixed", "localize_event", "localize_server", "localize_旧活动备份"]:
                continue
                
            ws = wb[sheet_name]
            src_rows = []
            existing_keys = {} # key_upper -> row_idx
            
            # Single pass to build key map and find source rows
            for r in range(1, ws.max_row + 1):
                val = str(ws.cell(r, 1).value or "").strip()
                if not val: continue
                
                k_clean = val.replace("//", "").strip().upper()
                if k_clean:
                    all_final_keys.add(k_clean)
                    if k_clean not in existing_keys:
                        existing_keys[k_clean] = r
                
                # Find source rows
                if r >= 2:
                    c1 = val.lower()
                    c2 = str(ws.cell(r, 2).value or "").lower()
                    if src_lower in c1 or src_lower in c2 or (src_digits and src_digits in c1):
                        src_rows.append(r)
            
            if src_rows:
                cloned_in_sheet = 0
                for r in src_rows:
                    new_key = swap_name(ws.cell(r, 1).value)
                    new_key_upper = str(new_key).upper().strip()
                    
                    if new_key_upper in existing_keys:
                        row_idx = existing_keys[new_key_upper]
                        c1_val = str(ws.cell(row_idx, 1).value or "").strip()
                        if c1_val.startswith("//"):
                            ws.cell(row_idx, 1, c1_val[2:].strip())
                            modified = True
                        continue
                        
                    new_row_idx = ws.max_row + 1
                    for c in range(1, ws.max_column + 1):
                        src_c = ws.cell(r, c)
                        dst_c = ws.cell(new_row_idx, c)
                        dst_c.value = swap_name(src_c.value)
                        copy_cell_format(src_c, dst_c)

                    existing_keys[new_key_upper] = new_row_idx
                    all_final_keys.add(new_key_upper)
                    cloned_in_sheet += 1
                
                if cloned_in_sheet > 0:
                    log(f"localization: Cloned {cloned_in_sheet} rows in [{sheet_name}] for {tgt_evt}")
                    modified = True

        # 1b. localize_event_fixed — divider-bounded '活动名 & 礼包名区 & 成就分享区' group.
        # Append the target's 2 rows (<EVENT>_EVENTNAME, ACHIEVEMENT_<EVENT>) at the end
        # of the group, before the next divider — never at the end of the whole sheet.
        if "localize_event_fixed" in wb.sheetnames:
            ws_lef = wb["localize_event_fixed"]
            start, end = find_icon_item_group_bounds(ws_lef, "活动名")
            if start is None:
                log("localize_event_fixed: '活动名' group marker not found — skipping.")
            else:
                tgt_already_present = any(tgt_lower in str(ws_lef.cell(r, 1).value or "").lower() for r in range(start, end))
                if tgt_already_present:
                    log(f"localize_event_fixed: {tgt_evt} already has row(s) in the group. Skipping.")
                else:
                    lef_rows, lef_label = get_localize_event_fixed_rows_for_source(file_path, src_lower, src_start_time, log_dir)
                    if lef_rows:
                        log(f"localize_event_fixed: Resolved {len(lef_rows)} row(s) for {src_evt} from {lef_label}.")
                        insert_at = end
                        for row_pairs in lef_rows:
                            ws_lef.insert_rows(insert_at)
                            for c_idx, (v, style) in enumerate(row_pairs):
                                dst_c = ws_lef.cell(insert_at, c_idx + 1, swap_name(v))
                                apply_cell_style_dict(dst_c, style)
                            insert_at += 1
                        log(f"localize_event_fixed: Appended {len(lef_rows)} row(s) for {tgt_evt} to the 活动名 group.")
                        # insert_rows() shifts cell data down but never touches auto_filter.ref,
                        # leaving its row bound stale against the sheet's new dimension — that
                        # mismatch is exactly what triggers Excel's "we found a problem with
                        # content" repair prompt on open. Extend the filter's row bound (only)
                        # to match; the column bound (H vs the sheet's real R) is a pre-existing,
                        # long-standing mismatch in this file and is left untouched.
                        if ws_lef.auto_filter.ref:
                            from openpyxl.utils import range_boundaries
                            min_col, min_row, max_col, max_row = range_boundaries(ws_lef.auto_filter.ref)
                            if ws_lef.max_row > max_row:
                                ws_lef.auto_filter.ref = (f"{ws_lef.cell(min_row, min_col).coordinate}:"
                                                           f"{ws_lef.cell(ws_lef.max_row, max_col).coordinate}")
                        modified = True
                    else:
                        log(f"localize_event_fixed: WARNING — no data found for {src_evt} anywhere (not live, not cached, not in SVN history).")

        # 1c. localize_event — backup-first (localize_旧活动备份), then cache, then SVN
        # history. Never falls back to a placeholder — if nothing is found anywhere,
        # that's logged plainly and left for a human, not papered over.
        # Gap-aware rather than skip-if-any-present: a reopen (or a re-run) that already holds part
        # of the event gets only the rows it still lacks.
        if "localize_event" in wb.sheetnames:
            ws_le = wb["localize_event"]
            have = {str(ws_le.cell(r, 1).value or "").replace("//", "").strip().upper() for r in range(2, ws_le.max_row + 1)}
            le_rows, le_label = get_localize_event_rows_for_source(file_path, src_lower, src_start_time, log_dir)
            new_rows = [rp for rp in le_rows or []
                        if str(swap_name(rp[0][0]) or "").replace("//", "").strip().upper() not in have]
            if new_rows:
                log(f"localize_event: Resolved {len(le_rows)} row(s) for {src_evt} from {le_label}.")
                insert_at = ws_le.max_row + 1
                for row_pairs in new_rows:
                    for c_idx, (v, style) in enumerate(row_pairs):
                        dst_c = ws_le.cell(insert_at, c_idx + 1, swap_name(v))
                        apply_cell_style_dict(dst_c, style)
                    insert_at += 1
                log(f"localize_event: Appended {len(new_rows)} row(s) for {tgt_evt}"
                    f"{f' ({len(le_rows) - len(new_rows)} already present)' if len(new_rows) < len(le_rows) else ''}.")
                modified = True
            elif le_rows:
                log(f"localize_event: all {len(le_rows)} row(s) of {src_evt} already present for {tgt_evt}. Skipping.")
            else:
                log(f"localize_event: WARNING — no data found for {src_evt} anywhere (not live, not in "
                    f"localize_旧活动备份, not cached, not in SVN history). Needs manual entry — no placeholder created.")

        # 1d. localize_server — pure accumulator, always live so far. Simple scan + clone.
        if "localize_server" in wb.sheetnames:
            ws_ls = wb["localize_server"]
            tgt_already_present = any(tgt_lower in str(ws_ls.cell(r, 1).value or "").lower() for r in range(2, ws_ls.max_row + 1))
            if tgt_already_present:
                log(f"localize_server: {tgt_evt} already present. Skipping.")
            else:
                ls_rows, ls_label = get_localize_server_rows_for_source(file_path, src_lower)
                if ls_rows:
                    log(f"localize_server: Resolved {len(ls_rows)} row(s) for {src_evt} from {ls_label}.")
                    insert_at = ws_ls.max_row + 1
                    for row_pairs in ls_rows:
                        for c_idx, (v, style) in enumerate(row_pairs):
                            dst_c = ws_ls.cell(insert_at, c_idx + 1, swap_name(v))
                            apply_cell_style_dict(dst_c, style)
                        insert_at += 1
                    log(f"localize_server: Appended {len(ls_rows)} row(s) for {tgt_evt}.")
                    modified = True
                else:
                    log(f"localize_server: WARNING — no data found live for {src_evt}.")

        # 1e. The generic pass (section 1) never scans localize_event_fixed/localize_event/
        # localize_server — they have their own dedicated handling above — so any key that
        # lives only in one of those three is invisible to all_final_keys unless folded in
        # here. Without this, the check below treats every such key as "missing" even
        # though 1b/1c/1d just cloned it correctly, and duplicates it as a [FIX ME] stub.
        for special_sheet in ("localize_event_fixed", "localize_event", "localize_server"):
            if special_sheet in wb.sheetnames:
                ws_sp = wb[special_sheet]
                for r in range(1, ws_sp.max_row + 1):
                    v = str(ws_sp.cell(r, 1).value or "").strip()
                    if v:
                        all_final_keys.add(v.replace("//", "").strip().upper())

        # 2. Final verification against REQUIRED keys — warning-only. The detector below
        # matches any "SIM2609_..."-shaped string in the target event's own xlsx, which
        # catches plenty of non-localization content codes (furniture/room/minigame-item
        # IDs) alongside real keys — confirmed false positives even for the source event,
        # which never had matching localization.xlsx entries either. Auto-creating
        # [FIX ME] rows for these was pure noise (and used a "localize_room" sheet name
        # that doesn't exist in the real file, landing in whatever sheet came first
        # instead). Logging plainly, like the icon.xlsx placeholder removal, so a human
        # can judge which of these actually need a translation entry.
        if required_tgt_keys:
            missing = [k for k in required_tgt_keys if k not in all_final_keys]
            if missing:
                log(f"localization: {len(missing)} referenced key(s) not found in localization.xlsx "
                    f"(may include non-translatable content codes, not just real gaps) — review manually: "
                    f"{', '.join(sorted(missing)[:15])}{', ...' if len(missing) > 15 else ''}")

        if modified:
            wb.save(file_path)
            log(f"localization.xlsx successfully updated for {tgt_evt}.")
        else:
            log(f"localization.xlsx: No changes needed.")
            
    except Exception as e:
        log(f"Error updating localization.xlsx: {e}")
        import traceback
        traceback.print_exc()
        return False
        
    return True


def update_localization_quiz_excel(proj_root, src_evt, tgt_evt, log_dir, step_index):
    """Step 14 Part B: Update localization_quiz.xlsx with backup/clone workflow."""
    import openpyxl
    import stat
    import os
    import json
    
    design_dir = os.path.join(proj_root, "design", "DesignData", WORKSPACE_NAME, "design")
    file_path = os.path.join(design_dir, "localization_quiz.xlsx")
    
    if not os.path.exists(file_path):
        log(f"Warning: localization_quiz.xlsx not found. Skipping.")
        return True
    
    # --- Derive quiz date prefixes ---
    src_quiz_date = get_quiz_date_prefix(proj_root, src_evt)
    if not src_quiz_date:
        log(f"Warning: Could not derive source quiz date for {src_evt}. Checking extra_args...")
        return True  # Non-fatal: quiz may not apply to all events
    
    # Primary: derive from events.xlsx (Step 10 already populated correct dates)
    tgt_quiz_date = get_quiz_date_prefix(proj_root, tgt_evt)
    
    # Fallback: check extra_args.json start_time if events.xlsx doesn't have target yet
    if not tgt_quiz_date:
        extra_args_path = os.path.join(log_dir, "extra_args.json")
        if os.path.exists(extra_args_path):
            try:
                with open(extra_args_path, "r") as f:
                    extra = json.load(f)
                start_time_str = extra.get("start_time", "")
                if start_time_str:
                    from datetime import datetime
                    try:
                        dt = datetime.strptime(start_time_str[:10], "%Y-%m-%d")
                        tgt_quiz_date = dt.strftime("%y%m%d")
                    except Exception:
                        pass
            except Exception:
                pass
    
    if not tgt_quiz_date:
        log(f"Warning: Could not derive target quiz date for {tgt_evt}. Cannot process localization_quiz.xlsx.")
        return True  # Non-fatal
    
    src_quiz_prefix = f"QUIZ{src_quiz_date}"  # e.g., QUIZ250717
    tgt_quiz_prefix = f"QUIZ{tgt_quiz_date}"  # e.g., QUIZ260501
    
    log(f"Quiz date derivation: {src_evt} -> {src_quiz_prefix}, {tgt_evt} -> {tgt_quiz_prefix}")
    
    def swap_quiz(value):
        """Swap quiz date prefixes in cell values."""
        if value is None: return None
        s = str(value)  # type: ignore
        s = s.replace(src_quiz_prefix.upper(), tgt_quiz_prefix.upper())
        s = s.replace(src_quiz_prefix.lower(), tgt_quiz_prefix.lower())
        # Also handle mixed case
        s = s.replace(src_quiz_prefix, tgt_quiz_prefix)
        return s
    
    try:
        os.chmod(file_path, stat.S_IWRITE | stat.S_IREAD)
        wb = load_workbook_frozen(file_path)
        
        quiz_sheet_name = "Quiz"
        backup_sheet_name = "Quiz_备份"
        
        if quiz_sheet_name not in wb.sheetnames or backup_sheet_name not in wb.sheetnames:
            log(f"Warning: Required sheets '{quiz_sheet_name}' or '{backup_sheet_name}' not found.")
            return True
        
        ws_quiz = wb[quiz_sheet_name]
        ws_backup = wb[backup_sheet_name]
        
        # --- Phase 1: Archive current Quiz rows to Quiz_备份 ---
        # Collect existing backup keys for deduplication
        existing_backup_keys = set()
        for r in range(2, ws_backup.max_row + 1):
            key = str(ws_backup.cell(r, 1).value or "").strip()  # type: ignore
            if key:
                existing_backup_keys.add(key.lower())
        
        archived_count = 0
        for r in range(2, ws_quiz.max_row + 1):
            quiz_key = str(ws_quiz.cell(r, 1).value or "").strip()  # type: ignore
            if not quiz_key:
                continue
            
            # Only archive if not already in backup
            if quiz_key.lower() not in existing_backup_keys:
                new_backup_row = ws_backup.max_row + 1
                for c in range(1, ws_quiz.max_column + 1):
                    ws_backup.cell(new_backup_row, c, ws_quiz.cell(r, c).value)  # type: ignore
                archived_count += 1
        
        log(f"localization_quiz: Archived {archived_count} new rows from Quiz to Quiz_备份.")
        
        # --- Phase 2: Clear Quiz sheet (keep row 1 header) ---
        # Save headers first
        headers = []
        for c in range(1, ws_quiz.max_column + 1):
            headers.append(ws_quiz.cell(1, c).value)  # type: ignore
        
        # Delete all data rows
        if ws_quiz.max_row > 1:
            ws_quiz.delete_rows(2, ws_quiz.max_row - 1)
        
        log(f"localization_quiz: Cleared Quiz sheet (kept headers).")
        
        # --- Phase 3: Clone source quiz rows from backup ---
        src_prefix_lower = src_quiz_prefix.lower()
        cloned_count = 0
        
        for r in range(2, ws_backup.max_row + 1):
            val = str(ws_backup.cell(r, 1).value or "").lower()  # type: ignore
            if src_prefix_lower in val:
                new_row = ws_quiz.max_row + 1
                for c in range(1, ws_backup.max_column + 1):
                    old_val = ws_backup.cell(r, c).value  # type: ignore
                    ws_quiz.cell(new_row, c, swap_quiz(old_val))  # type: ignore
                cloned_count += 1
        
        log(f"localization_quiz: Cloned {cloned_count} rows from Quiz_备份 ({src_quiz_prefix} -> {tgt_quiz_prefix}) into Quiz sheet.")
        
        wb.save(file_path)
        log(f"localization_quiz.xlsx successfully updated.")
        return True
        
    except Exception as e:
        log(f"Error updating localization_quiz.xlsx: {e}")
        import traceback
        traceback.print_exc()
        return False


def update_answer_challenge_excel(proj_root, src_evt, tgt_evt, log_dir, step_index):
    """
    Step 14 Part C: Update answer_challenge.xlsx — stage (rewards) only.

    answer_challenge_list and 题目 (questions) are deliberately NOT handled here
    anymore — both are being rebuilt as a separate step later (they were quiz-date
    coupled, and quiz handling itself moved out of this step). This function no
    longer depends on quiz-date derivation at all, since 'stage' doesn't need it.
    """
    import openpyxl
    import stat
    import os
    import re

    design_dir = os.path.join(proj_root, "design", "DesignData", WORKSPACE_NAME, "design")
    file_path = os.path.join(design_dir, "answer_challenge.xlsx")

    if not os.path.exists(file_path):
        log(f"Warning: answer_challenge.xlsx not found. Skipping.")
        return True

    src_lower = src_evt.lower()
    tgt_lower = tgt_evt.lower()
    src_upper = src_evt.upper()
    tgt_upper = tgt_evt.upper()
    src_pascal = src_evt[0].upper() + src_evt[1:] if src_evt else src_evt
    tgt_pascal = tgt_evt[0].upper() + tgt_evt[1:] if tgt_evt else tgt_evt

    try:
        os.chmod(file_path, stat.S_IWRITE | stat.S_IREAD)
        wb = load_workbook_frozen(file_path)
        modified = False

        # === stage (Rewards) — auto-detect event reward references and replace ===
        if "stage" in wb.sheetnames:
            ws_stage = wb["stage"]
            stage_modified = 0
            
            # Auto-detect the current event name in stage rewards
            # Pattern: {EventName}{4digits}能量 or {EVENTNAME}{4digits}代币
            detected_event_name = None
            detected_event_digits = None
            for r in range(2, ws_stage.max_row + 1):
                for c in range(1, ws_stage.max_column + 1):
                    cell_val = str(ws_stage.cell(r, c).value or "")  # type: ignore
                    # Match patterns like "Merge2512能量" or "MERGE2512代币"
                    evt_match = re.search(r'([A-Za-z]+)(\d{4})(能量|代币)', cell_val)
                    if evt_match:
                        detected_event_name = evt_match.group(1)
                        detected_event_digits = evt_match.group(2)
                        break
                if detected_event_name:
                    break
            
            if detected_event_name and detected_event_digits:
                det_lower = detected_event_name.lower()
                det_upper = detected_event_name.upper()
                det_pascal = detected_event_name[0].upper() + detected_event_name[1:]

                # 'stage' is a single shared slot across EVERY event type and workspace —
                # whichever event last ran this step (in any workspace) owns it, regardless
                # of which event is being reopened here. So the detected name/digits are only
                # used to locate what to replace; the suffix must be forced from tgt_evt's type.
                # SVN history of trunk/WorkSpcae3 (40 revisions, 2512 → 2610) shows two
                # different conventions side by side:
                #   额外奖励内容 (item-alias Map, cols 9/14/19): Merge → 'Merge2608能量:30',
                #                                                 every other type → '代币'
                #   额外奖励图标列表 (icon alias, cols 10/15/20): '代币' for ALL types,
                #                                                 Merge included ('MERGE2608代币')
                # The icon alias must match icon.xlsx icon_item, which names Merge's token
                # icon 'MERGE2608代币' too — so icon columns are left to the plain name swap.
                is_merge_tgt = tgt_evt.lower().startswith("merge")
                correct_suffix = "能量" if is_merge_tgt else "代币"
                wrong_suffix = "代币" if is_merge_tgt else "能量"
                # Only the suffix glued to the target's own event code, and never the '能量'
                # inside a compound like Dessert's 'Dessert2608能量代币' item.
                suffix_re = re.compile(rf'({re.escape(tgt_evt)}){wrong_suffix}(?!代币)', re.IGNORECASE)
                reward_cols = {c for c in range(1, ws_stage.max_column + 1)
                               if "图标" not in str(ws_stage.cell(1, c).value or "")}

                log(f"answer_challenge [stage]: Detected '{detected_event_name}{detected_event_digits}' in rewards, replacing with '{tgt_pascal}'")

                suffix_corrections = 0
                for r in range(2, ws_stage.max_row + 1):
                    for c in range(1, ws_stage.max_column + 1):
                        cell_val = ws_stage.cell(r, c).value  # type: ignore
                        if cell_val is not None:
                            s = str(cell_val)
                            new_s = s
                            # Replace all case variants: Merge2512 -> Sim2605, MERGE2512 -> SIM2605
                            # Fixed replacement to avoid redundant digit concatenation
                            # If tgt_upper already includes tgt_digits (e.g. SIM2605), don't append it again.
                            new_s = new_s.replace(f"{det_upper}{detected_event_digits}", tgt_upper)
                            new_s = new_s.replace(f"{det_lower}{detected_event_digits}", tgt_lower)
                            new_s = new_s.replace(f"{det_pascal}{detected_event_digits}", tgt_pascal)
                            if c in reward_cols:
                                fixed_s = suffix_re.sub(lambda m: m.group(1) + correct_suffix, new_s)
                                if fixed_s != new_s:
                                    new_s = fixed_s
                                    suffix_corrections += 1
                            if new_s != s:
                                ws_stage.cell(r, c, new_s)  # type: ignore
                                stage_modified += 1

                if suffix_corrections > 0:
                    log(f"answer_challenge [stage]: Corrected {suffix_corrections} reward-currency suffix reference(s) "
                        f"from '{wrong_suffix}' to '{correct_suffix}' (forced by {tgt_evt}'s event type, not carried over "
                        f"from the detected '{detected_event_name}{detected_event_digits}' source).")

                if stage_modified > 0:
                    log(f"answer_challenge [stage]: Updated {stage_modified} reward references")
                    modified = True
            else:
                log(f"answer_challenge [stage]: No event reward references detected.")
        
        if modified:
            wb.save(file_path)
            log(f"answer_challenge.xlsx successfully updated.")
        else:
            log(f"answer_challenge.xlsx: No changes needed.")
        
        return True
        
    except Exception as e:
        log(f"Error updating answer_challenge.xlsx: {e}")
        import traceback
        traceback.print_exc()
        return False


def _asset_ref_divider_matches(normalized, event_upper):
    """True if a '//...' divider label in asset_ref.xlsx's key_asset sheet refers
    to event_upper. Divider labels don't always match the event code exactly —
    e.g. every Sim event's divider is written '//SIMS<digits>' (an extra S)
    while its item keys, event name, and every other file all use 'SIM<digits>'
    (confirmed on both SIM2605 and SIM2607 in trunk). Match on digit equality +
    letter-substring overlap instead of strict equality so that drift — or the
    similar '//EVENT活动' suffix variant — doesn't silently break block lookup."""
    if not normalized.startswith("//"):
        return False
    ev_digits_m = re.search(r'\d+', event_upper)
    ev_digits = ev_digits_m.group() if ev_digits_m else ""
    div_digits_m = re.search(r'\d+', normalized)
    div_digits = div_digits_m.group() if div_digits_m else ""
    # Exact equality on the digit run, not substring containment — event codes
    # sharing the same month digits but a different type (e.g. //SORT2606 vs
    # //MERGE2606) must never cross-match on a short/partial digit overlap.
    if ev_digits != div_digits:
        return False
    letters = re.sub(r'\d', '', event_upper)
    div_letters = re.sub(r'\d', '', normalized[2:])
    return bool(letters) and bool(div_letters) and (letters in div_letters or div_letters in letters)


def get_asset_ref_block_rows_for_source(file_path, src_upper, hint_date, log_dir):
    """(rows, source_label) for src_upper's '//EVENT' asset block in asset_ref.xlsx.
    Tier 1: live key_asset, then live key_asset备份 (free). Tier 2: local cache
    (free). Tier 3: SVN history of asset_ref.xlsx (date-bracketed by hint_date
    when available). Each row is a list of (value, style_dict) pairs — a plain
    cell reference doesn't survive the cache/SVN round-trip, so style has to be
    captured and reapplied the same way as icon_item / localize_event."""
    import openpyxl as _oxl
    import io as _io

    src_lower = src_upper.lower()

    def _extract_block(wb, sheet_name):
        """Find src_upper's data rows by item-key prefix (e.g. 'SIM2605_') — not by
        divider — since some branches have the block with no divider comment above
        it at all (confirmed: Misc's WorkSpcae_Design has SIM2605_* rows with zero
        divider, sitting directly after an unrelated theme's block). If a divider
        does immediately precede the first data row, include it too, so it gets
        renamed for the target event same as before."""
        if sheet_name not in wb.sheetnames:
            return None
        ws = wb[sheet_name]
        prefix = f"{src_upper}_"
        data_rows = [r for r in range(1, ws.max_row + 1)
                     if str(ws.cell(r, 1).value or "").strip().upper().startswith(prefix)]
        if not data_rows:
            return None
        row_indices = data_rows
        divider_r = data_rows[0] - 1
        if divider_r >= 1:
            v = str(ws.cell(divider_r, 1).value or "").strip().replace(" ", "").upper()
            if _asset_ref_divider_matches(v, src_upper):
                row_indices = [divider_r] + row_indices
        return [[(ws.cell(r, c).value, cell_style_to_dict(ws.cell(r, c))) for c in range(1, ws.max_column + 1)] for r in row_indices]

    wb_live = _oxl.load_workbook(file_path, data_only=True)
    rows = _extract_block(wb_live, "key_asset")
    tier = "key_asset (live)"
    if not rows:
        rows = _extract_block(wb_live, "key_asset备份")
        tier = "key_asset备份 (live)"
    if rows:
        cache = load_named_cache(log_dir, "asset_ref_block_cache")
        cache[src_lower] = {"rows": rows, "source": tier, "cached_at": _cache_timestamp()}
        save_named_cache(log_dir, "asset_ref_block_cache", cache)
        return rows, f"{tier}, {len(rows)} row(s)"

    cache = load_named_cache(log_dir, "asset_ref_block_cache")
    cached_entry = cache.get(src_lower)
    if cached_entry and cached_entry.get("rows"):
        return cached_entry["rows"], f"local cache (from {cached_entry.get('source', '?')}, cached {cached_entry.get('cached_at', '?')})"

    # No convert.json short-circuit here: that heuristic ("no convert.json history
    # -> no data") was verified against icon_item/icon_event/item-icon specifically
    # (see event_ever_registered_in_convert_json) — it hasn't been verified for
    # asset_ref.xlsx, so skipping the SVN search on it would just be guessing.
    revisions = get_svn_revisions_in_date_range(file_path, hint_date) if hint_date else None
    if not revisions:
        revisions = get_svn_log_revisions(file_path, limit=60)
    if not revisions:
        return None, None

    for rev in revisions:
        content_bytes = svn_cat_revision_bytes(file_path, rev)
        if content_bytes is None:
            continue
        try:
            wb_hist = _oxl.load_workbook(_io.BytesIO(content_bytes), data_only=True)
        except Exception:
            continue
        rows = _extract_block(wb_hist, "key_asset") or _extract_block(wb_hist, "key_asset备份")
        if rows:
            cache[src_lower] = {"rows": rows, "source": f"svn r{rev}", "cached_at": _cache_timestamp()}
            save_named_cache(log_dir, "asset_ref_block_cache", cache)
            return rows, f"SVN history (r{rev}, freshly fetched and cached)"

    return None, None


def update_asset_ref_excel(proj_root, src_evt, tgt_evt, log_dir, step_index):
    """Step 15: Update asset_ref.xlsx — clone asset blocks and achievement rows."""
    import openpyxl
    import stat
    import os
    import re
    
    design_dir = os.path.join(proj_root, "design", "DesignData", WORKSPACE_NAME, "design")
    file_path = os.path.join(design_dir, "asset_ref.xlsx")
    
    if not os.path.exists(file_path):
        log(f"Error: asset_ref.xlsx not found at {file_path}")
        return False
    
    # Create backup for undo
    backup_path = file_path + ".bak"
    import stat
    os.chmod(file_path, stat.S_IWRITE | stat.S_IREAD)
    shutil.copy2(file_path, backup_path)
    record_file_creation(log_dir, step_index, backup_path)
    
    # Name variants
    src_lower = src_evt.lower()
    src_upper = src_evt.upper()
    src_pascal = src_evt[0].upper() + src_evt[1:]
    tgt_lower = tgt_evt.lower()
    tgt_upper = tgt_evt.upper()
    tgt_pascal = tgt_evt[0].upper() + tgt_evt[1:]
    src_digits = re.search(r'\d+', src_evt)
    tgt_digits = re.search(r'\d+', tgt_evt)
    src_digits = src_digits.group() if src_digits else ""
    tgt_digits = tgt_digits.group() if tgt_digits else ""
    
    def swap_name(val):
        if val is None:
            return val
        s = str(val)
        if isinstance(val, (int, float)):
            return val
        s = s.replace(src_upper, tgt_upper)
        s = s.replace(src_lower, tgt_lower)
        s = s.replace(src_pascal, tgt_pascal)
        s = s.replace(src_evt, tgt_evt)
        if src_digits and tgt_digits:
            s = s.replace(src_digits, tgt_digits)
        return s
    
    try:
        os.chmod(file_path, stat.S_IWRITE | stat.S_IREAD)
        wb = load_workbook_frozen(file_path)
        modified = False

        ws_main = wb["key_asset"] if "key_asset" in wb.sheetnames else None

        if not ws_main:
            log("Error: key_asset sheet not found in asset_ref.xlsx")
            return False

        # === Part 1: Clone source event asset block ===
        def find_event_block(ws, event_upper):
            """Rows already present for event_upper, by item-key prefix (e.g.
            'SIM2605_') for the target-already-exists check — not by divider, since
            a previously-cloned target block may itself have no divider (when its
            source didn't either). Matching the same way get_asset_ref_block_rows_
            for_source does keeps this in sync and avoids duplicate re-inserts."""
            prefix = f"{event_upper}_"
            return [r for r in range(1, ws.max_row + 1)
                    if str(ws.cell(r, 1).value or "").strip().upper().startswith(prefix)]

        # Live key_asset -> live key_asset备份 -> local cache -> SVN history
        extra_for_hint = load_extra_args(log_dir)
        src_start_time = extra_for_hint.get("src_start_time", "")
        src_block_rows, block_source = get_asset_ref_block_rows_for_source(file_path, src_upper, src_start_time, log_dir)

        if src_block_rows:
            log(f"asset_ref [key_asset]: Found {len(src_block_rows)} row asset block for {src_evt} from {block_source}")

            # Check if target event block already exists
            existing_check = find_event_block(ws_main, tgt_upper)
            if existing_check:
                log(f"asset_ref [key_asset]: Target block //{tgt_upper} already exists. Skipping block clone.")
            else:
                # Find insertion point — before //活动成就资源 section
                insert_row = ws_main.max_row + 1
                for r in range(1, ws_main.max_row + 1):
                    cell1 = str(ws_main.cell(r, 1).value or "")  # type: ignore
                    if "活动成就" in cell1:
                        insert_row = r
                        break
                
                # Shift existing rows down to make space
                # NOTE: openpyxl's ws.cell(r, c, value) is a no-op when value is None
                # (it only sets cell.value when value is not None), so None values must
                # be assigned via `.value =` directly or stale data is left behind.
                num_new_rows = len(src_block_rows)
                for r in range(ws_main.max_row, insert_row - 1, -1):
                    for c in range(1, ws_main.max_column + 1):
                        src_c = ws_main.cell(r, c)
                        dst_c = ws_main.cell(r + num_new_rows, c)
                        dst_c.value = src_c.value
                        copy_cell_format(src_c, dst_c)

                # Clear original cells to avoid duplicates
                for r in range(insert_row, insert_row + num_new_rows):
                    for c in range(1, ws_main.max_column + 1):
                        ws_main.cell(r, c).value = None  # type: ignore

                # Insert cloned block
                for i, row_data in enumerate(src_block_rows):
                    target_r = insert_row + i
                    for c_idx, (val, style) in enumerate(row_data):
                        dst_c = ws_main.cell(target_r, c_idx + 1)
                        dst_c.value = swap_name(val)  # type: ignore
                        apply_cell_style_dict(dst_c, style)
                
                log(f"asset_ref [key_asset]: Inserted {num_new_rows} rows at R{insert_row} for //{tgt_upper}")
                modified = True
        else:
            log(f"asset_ref [key_asset]: No asset block found for {src_evt}. Normal for some event types.")
        
        # === Part 2: Clone achievement row ===
        achieve_src_row = None
        max_chronicle = 0
        
        for r in range(1, ws_main.max_row + 1):
            cell1 = str(ws_main.cell(r, 1).value or "")  # type: ignore
            if cell1.strip().upper() == f"ACHIEVEMENT_E_{src_upper}":
                achieve_src_row = r
            cell3 = str(ws_main.cell(r, 3).value or "")  # type: ignore
            chronicle_match = re.search(r'icon_Chronicle(\d+)', cell3, re.IGNORECASE)
            if chronicle_match:
                num = int(chronicle_match.group(1))
                if num > max_chronicle:
                    max_chronicle = num
        
        if achieve_src_row:
            # Check if target already exists
            tgt_achieve_exists = False
            for r in range(1, ws_main.max_row + 1):
                cell1 = str(ws_main.cell(r, 1).value or "").strip().upper()
                if cell1 == f"ACHIEVEMENT_E_{tgt_upper}" or (cell1.startswith("//") and f"ACHIEVEMENT_E_{tgt_upper}" in cell1):
                    tgt_achieve_exists = True
                    if cell1.startswith("//"):
                        ws_main.cell(r, 1, cell1[2:].strip())
                        log(f"asset_ref [achievement]: Uncommented ACHIEVEMENT_E_{tgt_upper}")
                        modified = True
                    break
            
            if tgt_achieve_exists:
                log(f"asset_ref [achievement]: ACHIEVEMENT_E_{tgt_upper} already exists. Skipping.")
            else:
                new_chronicle = max_chronicle + 1
                # Insert after the last achievement row (find //通用活动商店 section)
                insert_after = achieve_src_row
                for r in range(achieve_src_row + 1, ws_main.max_row + 1):
                    cell1 = str(ws_main.cell(r, 1).value or "")  # type: ignore
                    if cell1.startswith("ACHIEVEMENT_E_"):
                        insert_after = r
                    elif cell1.startswith("//"):
                        break
                
                new_row = insert_after + 1

                # Shift rows down (see note above re: openpyxl None no-op)
                for r in range(ws_main.max_row, new_row - 1, -1):
                    for c in range(1, ws_main.max_column + 1):
                        src_c = ws_main.cell(r, c)
                        dst_c = ws_main.cell(r + 1, c)
                        dst_c.value = src_c.value  # type: ignore
                        copy_cell_format(src_c, dst_c)

                # Clone the source achievement row with name swap
                for c in range(1, ws_main.max_column + 1):
                    src_c = ws_main.cell(achieve_src_row, c)
                    old_val = src_c.value  # type: ignore
                    new_val = swap_name(old_val)
                    if c == 3 and isinstance(new_val, str):
                        new_val = re.sub(r'icon_Chronicle\d+', f'icon_Chronicle{new_chronicle}', new_val)
                    dst_c = ws_main.cell(new_row, c)
                    dst_c.value = new_val  # type: ignore
                    copy_cell_format(src_c, dst_c)
                
                log(f"asset_ref [achievement]: Added ACHIEVEMENT_E_{tgt_upper} at R{new_row} with Chronicle{new_chronicle}")
                modified = True
        else:
            log(f"asset_ref [achievement]: No source achievement row found for {src_evt}")
        
        if modified:
            wb.save(file_path)
            log(f"asset_ref.xlsx successfully updated for {tgt_evt}.")
        else:
            log(f"asset_ref.xlsx: No changes needed.")
        
        return True
        
    except Exception as e:
        log(f"Error updating asset_ref.xlsx: {e}")
        import traceback
        traceback.print_exc()
        return False


def find_generic_gift_style_row_indices(ws_s):
    """
    event_gift_style's 12-row content (stage 1/2/3 x with/without item, 2 rows
    each) is identical across every event type — verified byte-for-byte against
    real Merge, Sim, Blocks, Sort, and Dessert entries; only column 1 (the gift
    ID) ever differs. Returns the row indices (ascending) of the most recent
    complete 12-row block in the live sheet, to use as a backfill template —
    no SVN needed, this template is always present in the live file already.
    """
    rows_with_id = []
    for r in range(ws_s.max_row, 1, -1):
        if ws_s.cell(r, 1).value is not None:
            rows_with_id.append(r)
        if len(rows_with_id) == 12:
            break
    if len(rows_with_id) < 12:
        return None
    rows_with_id.reverse()
    return rows_with_id


def update_store_excel(proj_root, src_evt, tgt_evt, log_dir, step_index):
    """Step 16: Update store.xlsx — clone event rows across 4 sheets with ID remapping."""
    import openpyxl
    import stat
    import os
    import re
    import shutil
    
    design_dir = os.path.join(proj_root, "design", "DesignData", WORKSPACE_NAME, "design")
    file_path = os.path.join(design_dir, "store.xlsx")
    
    if not os.path.exists(file_path):
        log(f"Error: store.xlsx not found at {file_path}")
        return False
    
    # Create backup for undo
    backup_path = file_path + ".bak"
    os.chmod(file_path, stat.S_IWRITE | stat.S_IREAD)
    shutil.copy2(file_path, backup_path)
    record_file_creation(log_dir, step_index, backup_path)
    
    # Name variants for swapping
    src_lower = src_evt.lower()
    src_upper = src_evt.upper()
    src_pascal = src_evt[0].upper() + src_evt[1:]
    tgt_lower = tgt_evt.lower()
    tgt_upper = tgt_evt.upper()
    tgt_pascal = tgt_evt[0].upper() + tgt_evt[1:]
    
    def swap_name(val):
        """Replace source event name references with target event name."""
        if val is None:
            return val
        if isinstance(val, (int, float)):
            return val
        s = str(val)
        s = s.replace(src_upper, tgt_upper)
        s = s.replace(src_lower, tgt_lower)
        s = s.replace(src_pascal, tgt_pascal)
        s = s.replace(src_evt, tgt_evt)
        return s
    
    def remap_ids_in_cell(val, id_map):
        """Replace old IDs with new IDs in a comma-separated cell value."""
        if val is None:
            return val
        if isinstance(val, (int, float)):
            old_id = str(int(val))
            return int(id_map[old_id]) if old_id in id_map else val
        s = str(val)
        parts = s.split(",")
        new_parts = []
        for part in parts:
            stripped = part.strip()
            if stripped in id_map:
                new_parts.append(id_map[stripped])
            else:
                new_parts.append(part)
        return ",".join(new_parts)
    
    # Load extra_args for event timing
    extra_path = os.path.join(log_dir, "extra_args.json")
    start_time = None
    end_time = None
    if os.path.exists(extra_path):
        import json
        with open(extra_path, "r", encoding="utf-8") as f:
            extra = json.load(f)
        start_time = extra.get("start_time", "")
        end_time = extra.get("end_time", "")
    
    try:
        wb = load_workbook_frozen(file_path)
        modified = False
        
        # =====================================================
        # Part 1: store_recharge — clone 4 event shop rows
        # =====================================================
        ws_r = wb["store_recharge"] if "store_recharge" in wb.sheetnames else None
        new_recharge_ids = []  # Track new IDs for store_package mirroring
        old_recharge_ids = []
        
        if ws_r:
            # Find source event rows
            src_recharge_rows = []
            for r in range(2, ws_r.max_row + 1):
                alias = str(ws_r.cell(r, 2).value or "")  # type: ignore
                if src_evt.upper() in alias.upper() or src_evt.lower() in alias.lower() or src_evt in alias:
                    # Only match event shop rows (商店礼包), not templates
                    if not alias.startswith("B-") and not alias.startswith("A-"):
                        src_recharge_rows.append(r)
                        old_recharge_ids.append(str(ws_r.cell(r, 1).value or ""))  # type: ignore
            
            if src_recharge_rows:
                exists, row_idx = check_and_handle_existing(ws_r, "store_recharge", tgt_evt)
                if exists:
                    log(f"store_recharge: Target event {tgt_evt} already exists.")
                    # If it exists (and was uncommented), we still need the IDs for mapping to other sheets
                    # Let's collect existing IDs in this section
                    # Search for rows matching tgt_evt that are active
                    for r in range(2, ws_r.max_row + 1):
                        alias_val = str(ws_r.cell(r, 2).value or "").strip()
                        if tgt_evt.lower() in alias_val.lower() and not alias_val.startswith("B-") and not alias_val.startswith("A-"):
                            id_val = str(ws_r.cell(r, 1).value or "").strip()
                            if id_val and not id_val.startswith("//"):
                                new_recharge_ids.append(id_val)
                    src_recharge_rows = [] # Skip cloning logic
                else:
                    # Find max ID in same event section as source
                    # Source rows define the section; find next section divider after them
                    last_src_row = src_recharge_rows[-1]
                    section_end = last_src_row
                    for r in range(last_src_row + 1, ws_r.max_row + 1):
                        id_val = str(ws_r.cell(r, 1).value or "")  # type: ignore
                        if id_val.startswith("//"):
                            section_end = r - 1
                            break
                        # Also stop if we hit an empty row
                        if not id_val:
                            section_end = r - 1
                            break
                        section_end = r
                    
                    # Find max ID within the event section (between // markers)
                    # Walk backward from source to find section start marker
                    section_start = 2
                    for r in range(src_recharge_rows[0] - 1, 1, -1):
                        id_val = str(ws_r.cell(r, 1).value or "")  # type: ignore
                        if id_val.startswith("//"):
                            section_start = r + 1
                            break
                    
                    max_id = 0
                    for r in range(section_start, section_end + 1):
                        id_val = ws_r.cell(r, 1).value  # type: ignore
                        if id_val:
                            try:
                                num = int(id_val)
                                if num > max_id:
                                    max_id = num
                            except (ValueError, TypeError):
                                pass
                    
                    # Insert right after last data row in the event section.
                    # This is NOT the end of the sheet — a shared template section
                    # (e.g. "// 活动进度礼包 - A") follows immediately, so existing
                    # rows must be shifted down first or they get silently overwritten.
                    insert_row = section_end + 1
                    num_new_rows = len(src_recharge_rows)
                    for r in range(ws_r.max_row, insert_row - 1, -1):
                        for c in range(1, ws_r.max_column + 1):
                            src_c = ws_r.cell(r, c)
                            dst_c = ws_r.cell(r + num_new_rows, c)
                            dst_c.value = src_c.value  # type: ignore
                            copy_cell_format(src_c, dst_c)

                    next_id = max_id + 1
                    for i, src_r in enumerate(src_recharge_rows):
                        new_id = next_id + i
                        new_recharge_ids.append(str(new_id))
                        for c in range(1, ws_r.max_column + 1):
                            src_c = ws_r.cell(src_r, c)
                            old_val = src_c.value  # type: ignore
                            new_val = swap_name(old_val)
                            if c == 1:
                                new_val = new_id
                            dst_c = ws_r.cell(insert_row + i, c)
                            dst_c.value = new_val  # type: ignore
                            copy_cell_format(src_c, dst_c)

                    log(f"store_recharge: Cloned {len(src_recharge_rows)} rows at R{insert_row}. IDs {new_recharge_ids[0]}-{new_recharge_ids[-1]}")
                    modified = True
            else:
                log(f"store_recharge: No event rows found for {src_evt}")
        
        # =====================================================
        # Part 2: store_package — clone with SAME IDs as recharge
        # =====================================================
        ws_p = wb["store_package"] if "store_package" in wb.sheetnames else None
        
        if ws_p and new_recharge_ids:
            src_package_rows = []
            for r in range(2, ws_p.max_row + 1):
                id_val = str(ws_p.cell(r, 1).value or "")  # type: ignore
                if id_val in old_recharge_ids:
                    src_package_rows.append(r)
            
            if src_package_rows:
                exists, row_idx = check_and_handle_existing(ws_p, "store_package", tgt_evt)
                if exists:
                    src_package_rows = [] # Skip cloning
                else:
                    # Insert after last source event row in the section
                    last_pkg_src = src_package_rows[-1]
                    pkg_section_end = last_pkg_src
                    for r in range(last_pkg_src + 1, ws_p.max_row + 1):
                        id_val = str(ws_p.cell(r, 1).value or "")  # type: ignore
                        if id_val.startswith("//") or not id_val:
                            pkg_section_end = r - 1
                            break
                        pkg_section_end = r
                    
                    insert_row = pkg_section_end + 1
                    num_new_rows = len(src_package_rows)
                    for r in range(ws_p.max_row, insert_row - 1, -1):
                        for c in range(1, ws_p.max_column + 1):
                            src_c = ws_p.cell(r, c)
                            dst_c = ws_p.cell(r + num_new_rows, c)
                            dst_c.value = src_c.value  # type: ignore
                            copy_cell_format(src_c, dst_c)

                    old_to_new_recharge = dict(zip(old_recharge_ids, new_recharge_ids))

                    for i, src_r in enumerate(src_package_rows):
                        for c in range(1, ws_p.max_column + 1):
                            src_c = ws_p.cell(src_r, c)
                            old_val = src_c.value  # type: ignore
                            new_val = swap_name(old_val)
                            if c == 1:
                                old_id = str(old_val)
                                new_val = int(old_to_new_recharge.get(old_id, old_val))
                            dst_c = ws_p.cell(insert_row + i, c)
                            dst_c.value = new_val  # type: ignore
                            copy_cell_format(src_c, dst_c)

                    log(f"store_package: Cloned {len(src_package_rows)} rows at R{insert_row} with matching IDs")
                    modified = True
            else:
                log(f"store_package: No rows found matching recharge IDs for {src_evt}")
        
        # =====================================================
        # Part 3: recommend_gift — clone 18 rows with ID remapping
        # =====================================================
        ws_g = wb["recommend_gift"] if "recommend_gift" in wb.sheetnames else None
        new_gift_ids = []
        old_gift_ids = []
        gift_id_map = {}  # old_id → new_id for internal remapping
        
        if ws_g:
            src_gift_rows = []
            for r in range(2, ws_g.max_row + 1):
                row_text = ""
                for c in [2, 21, 29]:  # Alias, 依赖活动, 推荐条件
                    row_text += " " + str(ws_g.cell(r, c).value or "")  # type: ignore
                if src_evt.upper() in row_text.upper() or src_evt in row_text:
                    id_val = ws_g.cell(r, 1).value  # type: ignore
                    if id_val and not str(id_val).startswith("//"):
                        src_gift_rows.append(r)
                        old_gift_ids.append(str(int(id_val)))
            
            if src_gift_rows:
                exists, row_idx = check_and_handle_existing(ws_g, "recommend_gift", tgt_evt)
                if exists:
                    # Reconstruct new_gift_ids from the target's OWN current rows (not
                    # old_gift_ids, which holds src_evt's — wrong ids in a true clone,
                    # only coincidentally right when src==tgt). Without this, Part 4
                    # (event_gift_style) sees an empty gift_id_map and skips completely
                    # — no log, no warning, forever, even if style rows are genuinely
                    # missing. Confirmed live: merge2608 has all 18 recommend_gift rows
                    # but zero event_gift_style rows, and every reopen silently re-skips.
                    for r in range(2, ws_g.max_row + 1):
                        row_text = ""
                        for c in [2, 21, 29]:
                            row_text += " " + str(ws_g.cell(r, c).value or "")
                        if tgt_evt.upper() in row_text.upper() or tgt_evt in row_text:
                            id_val = ws_g.cell(r, 1).value
                            if id_val and not str(id_val).startswith("//"):
                                new_gift_ids.append(str(int(id_val)))
                    src_gift_rows = [] # Skip cloning
                else:
                    # Next ID = max of the event-gift ID band (below 18000), NOT of a row
                    # range. Event gifts live in 9xxx-11xxx; 18xxx is the non-event block
                    # (周末礼包, 节日BP推荐 184xx, 全屏礼包 185xx...). Those non-event rows
                    # are interleaved with the event rows in some workspaces and the old
                    # "// 活动商店与进度礼包" marker this used to key off is absent in
                    # WorkSpcae_NE (and now WorkSpcae_Design), so a row-range max picked
                    # up 18518 and Merge2611 got 18519+ instead of 11609+.
                    # ponytail: fixed 18000 ceiling; revisit if event IDs ever reach 17xxx.
                    EVENT_GIFT_ID_CEILING = 18000
                    all_gift_ids = set()
                    max_gift_id = 0
                    for r in range(2, ws_g.max_row + 1):
                        try:
                            num = int(ws_g.cell(r, 1).value)  # type: ignore
                        except (ValueError, TypeError):
                            continue
                        all_gift_ids.add(num)
                        if num < EVENT_GIFT_ID_CEILING and num > max_gift_id:
                            max_gift_id = num

                    clash = [max_gift_id + 1 + i for i in range(len(old_gift_ids))
                             if max_gift_id + 1 + i in all_gift_ids]
                    if clash:
                        raise RuntimeError(f"recommend_gift: new IDs {clash} already used elsewhere in the sheet")

                    log(f"recommend_gift: Event-gift max ID (< {EVENT_GIFT_ID_CEILING}) = {max_gift_id}")
                    
                    # Build old→new ID map
                    for i, old_id in enumerate(old_gift_ids):
                        new_id = str(max_gift_id + 1 + i)
                        gift_id_map[old_id] = new_id
                        new_gift_ids.append(new_id)
                    
                    log(f"recommend_gift: ID map ({len(gift_id_map)} entries): {list(gift_id_map.items())[:4]}...")
                    
                    # Insert after the last event row (not at sheet end)
                    last_gift_src = src_gift_rows[-1]
                    insert_row = last_gift_src + 1
                    # Walk forward to find the actual last event row
                    for r in range(ws_g.max_row, last_gift_src, -1):
                        alias = str(ws_g.cell(r, 2).value or "")  # type: ignore
                        id_val = ws_g.cell(r, 1).value  # type: ignore
                        if id_val and not str(id_val).startswith("//") and alias:
                            insert_row = r + 1
                            break
                    
                    # Derive end_time from source event's actual end_time (23:44:59 pattern)
                    src_end_time = None
                    src_start_time = None
                    if src_gift_rows:
                        src_end_time = ws_g.cell(src_gift_rows[0], 13).value  # type: ignore
                        src_start_time = ws_g.cell(src_gift_rows[0], 12).value  # type: ignore
                    
                    # Columns that need internal ID remapping
                    remap_cols = {4, 7, 16}  # 合并购买关系, 互斥组, 属性继承
                    
                    for i, src_r in enumerate(src_gift_rows):
                        for c in range(1, ws_g.max_column + 1):
                            src_c = ws_g.cell(src_r, c)
                            old_val = src_c.value  # type: ignore

                            if c == 1:
                                # Replace ID
                                old_id_str = str(int(old_val))
                                new_val = int(gift_id_map.get(old_id_str, old_val))
                            elif c in remap_cols:
                                # Remap internal cross-references
                                new_val = remap_ids_in_cell(old_val, gift_id_map)
                            elif c == 12 and start_time:
                                # Start time from extra_args
                                new_val = start_time
                            elif c == 13 and src_end_time is not None:
                                # Derive end_time: use target event end date + source time portion
                                # Source end_time has the correct time portion (e.g. 23:44:59)
                                src_end_str = str(src_end_time)
                                if end_time and " " in src_end_str:
                                    # Use end date from extra_args + time from source
                                    src_time_part = src_end_str.split(" ")[-1]
                                    end_date = str(end_time).split(" ")[0] if " " in str(end_time) else str(end_time)
                                    new_val = f"{end_date} {src_time_part}"
                                else:
                                    new_val = swap_name(old_val)
                            else:
                                new_val = swap_name(old_val)

                            dst_c = ws_g.cell(insert_row + i, c)
                            dst_c.value = new_val  # type: ignore
                            copy_cell_format(src_c, dst_c)

                    log(f"recommend_gift: Cloned {len(src_gift_rows)} rows at R{insert_row}. IDs {new_gift_ids[0]}-{new_gift_ids[-1]}")
                    modified = True
            else:
                log(f"recommend_gift: No rows found for {src_evt}")
        
        # =====================================================
        # Part 4: event_gift_style — clone 12 rows (first 12 gift IDs)
        # =====================================================
        ws_s = wb["event_gift_style"] if "event_gift_style" in wb.sheetnames else None

        # Gated on new_gift_ids (always populated when recommend_gift has target rows,
        # whether freshly cloned or reconstructed from an existing target) rather than
        # gift_id_map (only populated on a fresh clone) — see the exists-branch fix
        # above for why: this used to skip silently and permanently whenever
        # recommend_gift already existed for the target.
        #
        # Wrapped in its own try/except: this part must never be able to abort the
        # function before wb.save() runs — a failure here previously discarded
        # Part 1-3's already-computed changes (e.g. recommend_gift's new rows),
        # even though the log had already reported them as cloned.
        try:
            if ws_s and new_gift_ids:
                main_new_ids = new_gift_ids[:12] if len(new_gift_ids) >= 12 else new_gift_ids

                if gift_id_map:
                    # Fresh clone: copy the source's own style rows, remapping IDs.
                    main_old_ids = old_gift_ids[:12] if len(old_gift_ids) >= 12 else old_gift_ids
                    src_style_rows = []
                    for r in range(2, ws_s.max_row + 1):
                        id_val = str(ws_s.cell(r, 1).value or "")  # type: ignore
                        if id_val in main_old_ids:
                            src_style_rows.append(r)

                    if src_style_rows:
                        exists, row_idx = check_and_handle_existing(ws_s, "event_gift_style", tgt_evt)
                        if exists:
                            src_style_rows = [] # Skip cloning
                        else:
                            insert_row = ws_s.max_row + 1
                            for r in range(ws_s.max_row, 1, -1):
                                if ws_s.cell(r, 1).value is not None:  # type: ignore
                                    insert_row = r + 1
                                    break

                            for i, src_r in enumerate(src_style_rows):
                                for c in range(1, ws_s.max_column + 1):
                                    src_c = ws_s.cell(src_r, c)
                                    old_val = src_c.value  # type: ignore
                                    if c == 1:
                                        old_id_str = str(old_val)
                                        new_val = int(gift_id_map.get(old_id_str, old_val))
                                    else:
                                        new_val = old_val  # Style values don't contain event names
                                    dst_c = ws_s.cell(insert_row + i, c)
                                    dst_c.value = new_val  # type: ignore
                                    copy_cell_format(src_c, dst_c)

                            log(f"event_gift_style: Cloned {len(src_style_rows)} rows")
                            modified = True
                    else:
                        log(f"event_gift_style: No style rows found for source gift IDs")
                else:
                    # recommend_gift already existed for the target (reopen / re-run) — no
                    # source style rows to clone from. Just verify the target's own gift
                    # IDs actually have style entries. Where they're missing, backfill using
                    # the universal stage template (see find_generic_gift_style_row_indices) —
                    # matched by POSITION (1st gift ID -> template row 0, etc.), not by any
                    # per-event source, since the template content itself never varies by event.
                    existing_style_ids = set()
                    for r in range(2, ws_s.max_row + 1):
                        id_val = ws_s.cell(r, 1).value  # type: ignore
                        if id_val is not None:
                            existing_style_ids.add(str(id_val))
                    missing_positions = [i for i, gid in enumerate(main_new_ids) if gid not in existing_style_ids]

                    if not missing_positions:
                        log(f"event_gift_style: All of {tgt_evt}'s gift IDs already have style entries.")
                    else:
                        template_rows = find_generic_gift_style_row_indices(ws_s)
                        if not template_rows:
                            log(f"event_gift_style: WARNING — {tgt_evt}'s gift ID(s) "
                                f"{[main_new_ids[i] for i in missing_positions]} have no style entry, and no complete "
                                f"12-row template was found in the live sheet to backfill from. Needs manual entry.")
                        else:
                            insert_row = ws_s.max_row + 1
                            filled_ids = []
                            for i in missing_positions:
                                if i >= len(template_rows):
                                    continue  # only positions 0-11 have a defined stage template
                                src_r = template_rows[i]
                                for c in range(1, ws_s.max_column + 1):
                                    src_c = ws_s.cell(src_r, c)
                                    new_val = int(main_new_ids[i]) if c == 1 else src_c.value  # type: ignore
                                    dst_c = ws_s.cell(insert_row, c)
                                    dst_c.value = new_val  # type: ignore
                                    copy_cell_format(src_c, dst_c)
                                filled_ids.append(main_new_ids[i])
                                insert_row += 1
                            if filled_ids:
                                log(f"event_gift_style: Backfilled {len(filled_ids)} row(s) for {tgt_evt} using the "
                                    f"universal stage template (gift IDs {filled_ids}).")
                                modified = True
        except Exception as e:
            log(f"event_gift_style: Error — {e}. Skipping this part; other sheets already updated are unaffected.")
            traceback.print_exc()
        
        # =====================================================
        # Part 5: featured combo-pack currency (fixed rows 19806-19808 in
        # store_package). Verified via SVN history that this slot has held
        # Sort2606, Merge2606, Dessert2608, and Sim2607's currency at various
        # points — it's not tied to one event type, and per instruction this
        # should be kept in sync with EVERY event created, not gated behind a
        # manual flag. The replacement string is derived from the target's own
        # item.xlsx data (already cloned by this point in the pipeline): the
        # FIRST item row for any event is always its primary currency —
        # confirmed empirically as "能量" for Merge, "代币" for every other
        # type — so no hardcoded type-to-suffix mapping is needed.
        # =====================================================
        # Wrapped in its own try/except for the same reason as Part 4: a failure here
        # must not discard Parts 1-3's already-computed changes before wb.save() runs.
        try:
            if ws_p is not None:
                item_path = os.path.join(design_dir, "item.xlsx")
                tgt_currency_alias = None
                if os.path.exists(item_path):
                    wb_item = openpyxl.load_workbook(item_path, data_only=True)
                    if "item" in wb_item.sheetnames:
                        ws_item = wb_item["item"]
                        for r in range(2, ws_item.max_row + 1):
                            alias = str(ws_item.cell(r, 2).value or "")
                            if tgt_lower in alias.lower():
                                tgt_currency_alias = alias.strip()
                                break  # First match = primary currency (item.xlsx clones in source order)
                    wb_item.close()

                if not tgt_currency_alias:
                    log(f"store_package (featured combo): could not find {tgt_evt}'s primary currency in item.xlsx — skipping.")
                else:
                    currency_pattern = re.compile(r'([A-Za-z]+\d{4}(?:能量|代币)):(\d+)')
                    combo_updated = 0
                    for r in range(2, ws_p.max_row + 1):
                        alias = str(ws_p.cell(r, 2).value or "")
                        if "三方支付组合礼包" not in alias:
                            continue  # only the shared featured-combo-pack slot, never per-event historical rows
                        content = ws_p.cell(r, 3).value  # type: ignore
                        if not content:
                            continue
                        content_str = str(content)
                        m = currency_pattern.search(content_str)
                        if m and m.group(1) != tgt_currency_alias:
                            old_ref, qty = m.group(1), m.group(2)
                            new_content = content_str.replace(f"{old_ref}:{qty}", f"{tgt_currency_alias}:{qty}")
                            ws_p.cell(r, 3, new_content)  # type: ignore
                            combo_updated += 1
                            log(f"store_package: Featured combo pack R{r} (ID {ws_p.cell(r, 1).value}) "
                                f"'{old_ref}' -> '{tgt_currency_alias}'")

                    if combo_updated:
                        log(f"store_package: Updated {combo_updated} featured combo-pack row(s) to feature {tgt_evt}.")
                        modified = True
                    else:
                        log(f"store_package (featured combo): already features {tgt_evt} (or no matching rows found).")
        except Exception as e:
            log(f"store_package (featured combo): Error — {e}. Skipping this part; other sheets already updated are unaffected.")
            traceback.print_exc()

        if modified:
            wb.save(file_path)
            log(f"store.xlsx successfully updated for {tgt_evt}.")
        else:
            log(f"store.xlsx: No changes needed.")

        return True
        
    except Exception as e:
        log(f"Error updating store.xlsx: {e}")
        traceback.print_exc()
        return False


CARD_COLLECTION_KINDS = "二星卡包|三星卡包|四星卡包|五星卡包|自选卡包"
CARD_COLLECTION_PATTERN_RE = None  # compiled lazily in get_latest_card_collection_yymm / swap_name


def get_latest_card_collection_yymm(design_dir):
    """
    item.xlsx runs an independent 'card collection' release series — periodic
    entries named '<YYMM>二星卡包' / '三星卡包' / '四星卡包' / '五星卡包' / '自选卡包'
    (e.g. 2408, 2410, ..., 2605, 2607, ...), released roughly every couple of
    months, NOT tied to individual event dates. pack.xlsx rewards reference
    whichever collection was current when a pack row was first written, and
    since it's a different number than the event's own digits, the ordinary
    event-name swap never touches it — it rides along unchanged forever unless
    something explicitly re-points it at the current latest collection.
    Returns the max YYMM string found (e.g. "2607"), or None if none found.
    """
    import openpyxl
    import re as _re

    item_path = os.path.join(design_dir, "item.xlsx")
    if not os.path.exists(item_path):
        return None

    try:
        wb = openpyxl.load_workbook(item_path, data_only=True)
    except Exception:
        return None
    if "item" not in wb.sheetnames:
        return None

    ws = wb["item"]
    pattern = _re.compile(rf'^(\d{{4}})({CARD_COLLECTION_KINDS})$')
    latest = None
    for r in range(2, ws.max_row + 1):
        alias = str(ws.cell(r, 2).value or "")
        m = pattern.match(alias)
        if m and (latest is None or m.group(1) > latest):
            latest = m.group(1)
    return latest


def update_pack_excel(proj_root, src_evt, tgt_evt, log_dir, step_index):
    """Step 18: Update pack.xlsx — clone pack reward rows and auto-increment IDs."""
    import openpyxl
    import stat
    import os
    import re
    import shutil

    design_dir = os.path.join(proj_root, "design", "DesignData", WORKSPACE_NAME, "design")
    file_path = os.path.join(design_dir, "pack.xlsx")

    if not os.path.exists(file_path):
        log(f"Error: pack.xlsx not found at {file_path}")
        return False

    backup_path = file_path + ".bak"
    os.chmod(file_path, stat.S_IWRITE | stat.S_IREAD)
    shutil.copy2(file_path, backup_path)
    record_file_creation(log_dir, step_index, backup_path)

    src_lower = src_evt.lower()
    src_upper = src_evt.upper()
    src_pascal = src_evt[0].upper() + src_evt[1:]
    tgt_lower = tgt_evt.lower()
    tgt_upper = tgt_evt.upper()
    tgt_pascal = tgt_evt[0].upper() + tgt_evt[1:]

    latest_card_collection = get_latest_card_collection_yymm(design_dir)
    card_collection_re = re.compile(rf'(\d{{4}})({CARD_COLLECTION_KINDS})')
    if latest_card_collection:
        log(f"pack.xlsx: Latest card collection in item.xlsx is {latest_card_collection} — "
            f"any stale 卡包 reference in cloned rows will be repointed to it.")

    def swap_name(val):
        """Replace source event name references with target event name, and repoint
        any stale card-collection (卡包) reference to the current latest one."""
        if val is None:
            return val
        if isinstance(val, (int, float)):
            return val
        s = str(val)
        s = s.replace(src_upper, tgt_upper)
        s = s.replace(src_lower, tgt_lower)
        s = s.replace(src_pascal, tgt_pascal)
        s = s.replace(src_evt, tgt_evt)
        if latest_card_collection:
            s = card_collection_re.sub(
                lambda m: m.group(1) + m.group(2) if m.group(1) == latest_card_collection
                else latest_card_collection + m.group(2),
                s
            )
        return s

    try:
        wb = load_workbook_frozen(file_path)
        modified = False
        
        ws = wb["pack"] if "pack" in wb.sheetnames else None
        if not ws:
            log(f"pack.xlsx: 'pack' sheet not found")
            return False
            
        src_rows = []
        for r in range(2, ws.max_row + 1):
            alias = str(ws.cell(r, 2).value or "")  # type: ignore
            if src_evt.upper() in alias.upper() or src_evt in alias:
                src_rows.append(r)
                
        if not src_rows:
            log(f"pack.xlsx: No source rows found for {src_evt}")
            return True
            
        exists, _ = check_and_handle_existing(ws, "pack", tgt_evt)
        if exists:
            # check_and_handle_existing may have just uncommented the row in-memory —
            # it doesn't report back whether it actually mutated anything, so we can't
            # tell "already active" from "just reactivated" here. Always save: an
            # unnecessary save of an unchanged file is free; skipping a real one
            # silently discards the reactivation and leaves the row commented on disk
            # forever, while the log claims it succeeded.
            wb.save(file_path)
            return True

        max_id = 0
        for r in range(2, ws.max_row + 1):
            id_val = ws.cell(r, 1).value  # type: ignore
            if id_val:
                try:
                    num = int(id_val)
                    if num > max_id:
                        max_id = num
                except (ValueError, TypeError):
                    pass
                    
        insert_row = ws.max_row + 1
        for r in range(ws.max_row, 1, -1):
            if ws.cell(r, 1).value is not None:  # type: ignore
                insert_row = r + 1
                break
                
        next_id = max_id + 1
        for i, src_r in enumerate(src_rows):
            new_id = next_id + i
            for c in range(1, ws.max_column + 1):
                src_c = ws.cell(src_r, c)  # type: ignore
                old_val = src_c.value
                new_val = swap_name(old_val)
                if c == 1:
                    new_val = new_id
                dst_c = ws.cell(insert_row + i, c)  # type: ignore
                dst_c.value = new_val
                copy_cell_format(src_c, dst_c)
                
        log(f"pack.xlsx: Cloned {len(src_rows)} rows. New IDs {next_id} to {next_id + len(src_rows) - 1}")
        
        wb.save(file_path)
        log(f"pack.xlsx successfully updated for {tgt_evt}.")
        return True
        
    except Exception as e:
        log(f"Error updating pack.xlsx: {e}")
        import traceback
        traceback.print_exc()
        return False


def get_event_time_range(proj_root, event_name):
    """Return (start_str, end_str) for event_name from events.xlsx cols H/I, or (None, None)."""
    import openpyxl

    events_path = os.path.join(proj_root, "design", "DesignData", WORKSPACE_NAME, "design", "events.xlsx")
    if not os.path.exists(events_path):
        return None, None
    try:
        wb = openpyxl.load_workbook(events_path, data_only=True)
        ws = wb["events"]
    except Exception:
        return None, None

    evt_lower = event_name.lower()
    for r in range(2, ws.max_row + 1):
        alias = str(ws.cell(r, 2).value or "").lower()
        if evt_lower in alias:
            return (str(ws.cell(r, 8).value or "").strip() or None,
                    str(ws.cell(r, 9).value or "").strip() or None)
    return None, None


def source_event_safe_to_comment(proj_root, src_evt, log_dir):
    """
    Decide whether the SOURCE event's guide rows should be commented out when its
    content is cloned to a new event.

    Verified against SVN history of trunk/WorkSpcae3/design/guide.xlsx: adding a new
    event NEVER comments other events in the same commit. The active set accumulates
    (observed 0 -> 1 -> 2 -> 3 concurrent events) and rows are commented separately,
    later, when an event actually ENDS. So commenting is lifecycle-driven, not
    clone-driven.

    Practical rule that matches that: only comment the source if it has already
    finished before the new event starts. If they overlap (the source is still live,
    which is the common case since the UI defaults the Source dropdown to the latest
    event of that type) leave it active — otherwise we'd disable the tutorial of a
    currently-running event. When either date is unknown, default to NOT commenting,
    since that's the behaviour the history shows.
    """
    from datetime import datetime

    def _parse(s):
        if not s:
            return None
        try:
            return datetime.strptime(str(s).strip()[:19], "%Y-%m-%d %H:%M:%S")
        except Exception:
            try:
                return datetime.strptime(str(s).strip()[:10], "%Y-%m-%d")
            except Exception:
                return None

    _, src_end_str = get_event_time_range(proj_root, src_evt)
    src_end = _parse(src_end_str)

    tgt_start = _parse(load_extra_args(log_dir).get("start_time", ""))

    if src_end is None or tgt_start is None:
        log(f"guide: source '{src_evt}' end date or target start date unknown — leaving source rows ACTIVE "
            f"(matches observed convention: adding an event never deactivates others).")
        return False

    if src_end < tgt_start:
        log(f"guide: source '{src_evt}' ended {src_end_str} before the new event starts — safe to comment its rows.")
        return True

    log(f"guide: source '{src_evt}' is still live (ends {src_end_str}, new event starts {tgt_start}) — "
        f"leaving its rows ACTIVE so the running event's tutorial keeps working.")
    return False


def update_guide_excel(proj_root, src_evt, tgt_evt, log_dir, step_index):
    """Step 19: Update guide.xlsx — clone, remap IDs, insert precisely after active events."""
    import openpyxl
    import stat
    import os
    import shutil
    
    design_dir = os.path.join(proj_root, "design", "DesignData", WORKSPACE_NAME, "design")
    file_path = os.path.join(design_dir, "guide.xlsx")
    
    if not os.path.exists(file_path):
        log(f"Error: guide.xlsx not found at {file_path}")
        return False
        
    backup_path = file_path + ".bak"
    os.chmod(file_path, stat.S_IWRITE | stat.S_IREAD)
    shutil.copy2(file_path, backup_path)
    record_file_creation(log_dir, step_index, backup_path)
    
    src_lower = src_evt.lower()
    src_upper = src_evt.upper()
    src_pascal = src_evt[0].upper() + src_evt[1:]
    tgt_lower = tgt_evt.lower()
    tgt_upper = tgt_evt.upper()
    tgt_pascal = tgt_evt[0].upper() + tgt_evt[1:]
    
    def swap_name(val):
        if val is None: return val
        if isinstance(val, (int, float)): return val
        s = str(val)
        s = s.replace(src_upper, tgt_upper)
        s = s.replace(src_lower, tgt_lower)
        s = s.replace(src_pascal, tgt_pascal)
        s = s.replace(src_evt, tgt_evt)
        return s

    try:
        wb = load_workbook_frozen(file_path)
        modified = False

        # Decide once, up front, whether the source event's own rows should be
        # deactivated — see source_event_safe_to_comment for the SVN-verified rule.
        comment_src = source_event_safe_to_comment(proj_root, src_evt, log_dir)

        # 1. guide_section
        old_id_to_new_id = {}
        ws_sec = wb["guide_section"] if "guide_section" in wb.sheetnames else None
        if ws_sec:
            # find max ID among the event sections (IDs < 100000)
            max_evt_id = 0
            insert_sec_row = ws_sec.max_row + 1

            for r in range(2, ws_sec.max_row + 1):
                id_val_raw = str(ws_sec.cell(r, 1).value or "")
                id_val = id_val_raw.replace("//", "")

                if id_val.isdigit():
                    num = int(id_val)
                    if num < 100000:
                        max_evt_id = max(max_evt_id, num)
                    elif num >= 100000 and insert_sec_row > ws_sec.max_row:
                        # Found the start of the core guides (100001)
                        # We will insert standard events right before this row, 
                        # but wait, let's insert right before the //新手固定活动 divider if it exists
                        # Let's peek backwards to see if there's a divider
                        insert_sec_row = r
                        if str(ws_sec.cell(r - 1, 1).value or "").startswith("//"):
                            insert_sec_row = r - 1
            
            src_rows = []
            existing_target_rows = []
            for r in range(2, ws_sec.max_row + 1):
                alias_val = str(ws_sec.cell(r, 2).value or "").strip()
                id_val_raw = str(ws_sec.cell(r, 1).value or "").strip()
                
                if src_evt.lower() in alias_val.lower():
                    src_rows.append(r)
                
                if tgt_evt.lower() in alias_val.lower():
                    existing_target_rows.append(r)
            
            exists, row_idx = check_and_handle_existing(ws_sec, "guide_section", tgt_evt)
            if exists:
                src_rows = [] # This effectively skips the cloning block below
            
            if src_rows:
                next_id = max_evt_id + 1
                for i, src_r in enumerate(src_rows):
                    old_id_str = str(ws_sec.cell(src_r, 1).value or "").replace("//", "")
                    if old_id_str.isdigit():
                        old_id_to_new_id[int(old_id_str)] = next_id + i
                        
                # Insert rows precisely before the core game guides (insert_sec_row)
                ws_sec.insert_rows(insert_sec_row, amount=len(src_rows))
                
                # Fill newly inserted rows
                for i, src_r in enumerate(src_rows):
                    actual_src_r = src_r if src_r < insert_sec_row else src_r + len(src_rows)
                    for c in range(1, ws_sec.max_column + 1):
                        src_c = ws_sec.cell(actual_src_r, c)
                        old_val = src_c.value
                        new_val = swap_name(old_val)
                        if c == 1 and new_val:
                            old_id_str = str(old_val).replace("//", "")
                            if old_id_str.isdigit() and int(old_id_str) in old_id_to_new_id:
                                parsed_id = int(old_id_to_new_id[int(old_id_str)])
                                new_val = str(parsed_id) if isinstance(old_val, str) else parsed_id
                        dst_c = ws_sec.cell(insert_sec_row + i, c)
                        dst_c.value = new_val
                        copy_cell_format(src_c, dst_c)
                        
                # Comment out ONLY the source event's own rows (being superseded by this
                # clone) — never a blanket range. This used to comment everything from a
                # hardcoded section ID ("619") through the new insertion point, which
                # meant creating any one event silently disabled the guide sections of
                # every OTHER currently-active event caught in that range. Confirmed via
                # a real run: creating a new dessert event wiped out all 16 of sim2607's
                # (an unrelated, still-open event) guide_section rows.
                # ...and only when the source has actually ENDED (comment_src) — history
                # shows commenting is lifecycle-driven, not triggered by cloning.
                commented_n = 0
                if comment_src:
                    for src_r in src_rows:
                        actual_src_r = src_r if src_r < insert_sec_row else src_r + len(src_rows)
                        val = ws_sec.cell(actual_src_r, 1).value
                        if val and not str(val).startswith("//"):
                            ws_sec.cell(actual_src_r, 1, f"//{val}")
                            commented_n += 1

                modified = True
                log(f"guide_section: Inserted at row {insert_sec_row}. Mapped {len(old_id_to_new_id)} IDs. "
                    f"Commented out {commented_n} source row(s).")
        
        # 2. guide and guide_backup
        ws_guide = wb["guide"] if "guide" in wb.sheetnames else None
        ws_gb = wb["guide_backup"] if "guide_backup" in wb.sheetnames else None
        
        if ws_guide and ws_gb and old_id_to_new_id:
            # Guide event IDs are sectionID * 100 + step. e.g. section 619 -> guide ID > 61900
            # Identify active event rows in guide (61900 <= ID < 10000000)
            rows_to_transfer = []
            insert_g_row = ws_guide.max_row + 1
            
            for r in range(2, ws_guide.max_row + 1):
                id_val_str = str(ws_guide.cell(r, 1).value or "").replace("//", "")
                if id_val_str.isdigit():
                    num = int(id_val_str)
                    if 61900 <= num < 10000000:
                        rows_to_transfer.append(r)
                    elif num >= 10000000 and insert_g_row > ws_guide.max_row:
                        insert_g_row = r
                        if str(ws_guide.cell(r - 1, 1).value or "").startswith("//"):
                            insert_g_row = r - 1
            
            # If no rows to transfer were found, we still need the insertion point
            if insert_g_row > ws_guide.max_row:
                for r in range(2, ws_guide.max_row + 1):
                    id_val_str = str(ws_guide.cell(r, 1).value or "").replace("//", "")
                    if id_val_str.isdigit() and int(id_val_str) >= 10000000:
                        insert_g_row = r
                        if str(ws_guide.cell(r - 1, 1).value or "").startswith("//"):
                            insert_g_row = r - 1
                        break
            
            if rows_to_transfer:
                gb_insert_row = ws_gb.max_row + 1
                for r in range(ws_gb.max_row, 1, -1):
                    if ws_gb.cell(r, 1).value is not None:
                        gb_insert_row = r + 1
                        break
                        
                for r in rows_to_transfer:
                    id_val = str(ws_guide.cell(r, 1).value or "")
                    exists_in_gb = False
                    # Simple check to avoid duplicates in backup
                    for b_row in range(gb_insert_row - 100, gb_insert_row):
                        if b_row > 1 and str(ws_gb.cell(b_row, 1).value or "") == id_val:
                            exists_in_gb = True
                            break
                    if not exists_in_gb and id_val:
                        for c in range(1, ws_guide.max_column + 1):
                            src_c = ws_guide.cell(r, c)
                            dst_c = ws_gb.cell(gb_insert_row, c)
                            dst_c.value = src_c.value
                            copy_cell_format(src_c, dst_c)
                        gb_insert_row += 1
                
                # Delete rows from guide
                # Do it in reverse order to preserve indices
                for r in reversed(rows_to_transfer):
                    ws_guide.delete_rows(r, 1)
                    if r < insert_g_row:
                        insert_g_row -= 1
            
            # Find source rows in guide_backup using old_id_to_new_id prefix matching
            src_gb_rows = []
            for r in range(2, ws_gb.max_row + 1):
                id_val_str = str(ws_gb.cell(r, 1).value or "").replace("//", "")
                if id_val_str.isdigit() and len(id_val_str) >= 3:
                    base_old = int(id_val_str[:-2])
                    if base_old in old_id_to_new_id:
                        src_gb_rows.append(r)
            
            if src_gb_rows:
                ws_guide.insert_rows(insert_g_row, amount=len(src_gb_rows))
                for i, src_r in enumerate(src_gb_rows):
                    for c in range(1, ws_gb.max_column + 1):
                        src_c = ws_gb.cell(src_r, c)
                        old_val = src_c.value
                        new_val = swap_name(old_val)
                        if c == 1 and new_val:
                            old_id_str = str(old_val).replace("//", "")
                            if old_id_str.isdigit() and len(old_id_str) >= 3:
                                base_old = int(old_id_str[:-2])
                                remng = old_id_str[-2:]
                                if base_old in old_id_to_new_id:
                                    mapped = int(f"{old_id_to_new_id[base_old]}{remng}")
                                    new_val = str(mapped) if isinstance(old_val, str) else mapped
                        dst_c = ws_guide.cell(insert_g_row + i, c)
                        dst_c.value = new_val
                        copy_cell_format(src_c, dst_c)
                modified = True
                log(f"guide & guide_backup: Transferred {len(rows_to_transfer)} rows, Cloned {len(src_gb_rows)} rows to row {insert_g_row}.")
                        
        # 3. guide_trigger
        ws_trig = wb["guide_trigger"] if "guide_trigger" in wb.sheetnames else None
        if ws_trig:
            # find original rows by name matching 
            src_rows = []
            existing_target_rows = []
            for r in range(2, ws_trig.max_row + 1):
                alias = str(ws_trig.cell(r, 2).value or "").lower()
                id_val = str(ws_trig.cell(r, 1).value or "").lower()
                if src_evt.lower() in alias or src_evt.lower() in id_val:
                    src_rows.append(r)
                if tgt_evt.lower() in alias or tgt_evt.lower() in id_val:
                    existing_target_rows.append(r)
                    
            exists, _ = check_and_handle_existing(ws_trig, "guide_trigger", tgt_evt)
            if exists:
                src_rows = [] # Skip cloning
            
            if src_rows:
                insert_row = ws_trig.max_row + 1
                for r in range(ws_trig.max_row, 1, -1):
                    if ws_trig.cell(r, 1).value is not None:
                        insert_row = r + 1
                        break
                
                # We append to the bottom for triggers 
                for i, src_r in enumerate(src_rows):
                    for c in range(1, ws_trig.max_column + 1):
                        src_c = ws_trig.cell(src_r, c)
                        old_val = src_c.value
                        new_val = swap_name(old_val)
                        if c == 1 and new_val:
                            new_val = str(new_val).replace("//", "") # Uncomment
                        dst_c = ws_trig.cell(insert_row + i, c)
                        dst_c.value = new_val
                        copy_cell_format(src_c, dst_c)
                        
                # Comment out ONLY the source event's own rows — a hardcoded row-148
                # blanket range used to be here, which (like guide_section) meant creating
                # any one event silently disabled every OTHER active event's triggers
                # caught in that range. Confirmed via a real run alongside the same bug
                # in guide_section: 31 unrelated sim2607 rows got wiped this way.
                # ...and only when the source has actually ENDED (comment_src).
                if comment_src:
                    for src_r in src_rows:
                        val = ws_trig.cell(src_r, 1).value
                        if val and not str(val).startswith("//"):
                            s_val = str(val).upper()
                            # DO NOT comment out shared MergeFix triggers or global triggers
                            if "MERGE_FIX" in s_val or "MERGEFIX" in s_val:
                                continue
                            ws_trig.cell(src_r, 1, f"//{val}")
                modified = True
                        
        # 4. guide_refui
        ws_ui = wb["guide_refui"] if "guide_refui" in wb.sheetnames else None
        if ws_ui:
            src_rows = []
            existing_target_rows = []
            for r in range(2, ws_ui.max_row + 1):
                alias = str(ws_ui.cell(r, 2).value or "").lower()
                id_val = str(ws_ui.cell(r, 1).value or "").lower()
                if src_evt.lower() in alias or src_evt.lower() in id_val:
                    src_rows.append(r)
                if tgt_evt.lower() in alias or tgt_evt.lower() in id_val:
                    existing_target_rows.append(r)
                    
            exists, _ = check_and_handle_existing(ws_ui, "guide_refui", tgt_evt)
            if exists:
                src_rows = [] # Skip cloning
            
            if src_rows:
                insert_row = ws_ui.max_row + 1
                for r in range(ws_ui.max_row, 1, -1):
                    if ws_ui.cell(r, 1).value is not None:
                        insert_row = r + 1
                        break
                        
                for i, src_r in enumerate(src_rows):
                    for c in range(1, ws_ui.max_column + 1):
                        src_c = ws_ui.cell(src_r, c)
                        old_val = src_c.value
                        new_val = swap_name(old_val)
                        if c == 1 and new_val:
                            new_val = str(new_val).replace("//", "") # Uncomment
                        dst_c = ws_ui.cell(insert_row + i, c)
                        dst_c.value = new_val
                        copy_cell_format(src_c, dst_c)
                        
                # Comment out ONLY the source event's own rows — same fix as guide_section
                # and guide_trigger. The old hardcoded row-97 blanket range wiped out every
                # other active event's UI refs caught in it (confirmed: 19 unrelated
                # sim2607 rows here in the same real-run test).
                # ...and only when the source has actually ENDED (comment_src).
                if comment_src:
                    for src_r in src_rows:
                        val = ws_ui.cell(src_r, 1).value
                        if val and not str(val).startswith("//"):
                            s_val = str(val).upper()
                            # DO NOT comment out shared MergeFix UI or global UI refs
                            if "MERGE_FIX" in s_val or "MERGEFIX" in s_val:
                                continue
                            ws_ui.cell(src_r, 1, f"//{val}")
                modified = True

        if modified:
            wb.save(file_path)
            log(f"guide.xlsx successfully updated for {tgt_evt}.")
        else:
            log(f"guide.xlsx: No changes needed.")
            
        return True
    
    except Exception as e:
        log(f"Error updating guide.xlsx: {e}")
        import traceback
        traceback.print_exc()
        return False


def update_localization_all(proj_root, src_evt, tgt_evt, log_dir, step_index):
    """
    Step 14: Orchestrator for localization + answer_challenge (stage rewards only).

    localization_quiz.xlsx (formerly Part B) is deliberately NOT called here — it's
    being rebuilt as a separate step later. update_localization_quiz_excel still
    exists in the codebase (not deleted) for when that work picks back up.

    answer_challenge_list and 题目 (formerly parts of Part C) were removed entirely
    from update_answer_challenge_excel, not just skipped — they'll be rebuilt fresh
    as a separate step, same as the quiz work.
    """

    # Part A: localization.xlsx
    log("=== Step 14 Part A: localization.xlsx ===")
    success_a = update_localization_main_excel(proj_root, src_evt, tgt_evt, log_dir, step_index)
    if not success_a:
        return False

    # Part C: answer_challenge.xlsx — stage (rewards) only
    log("=== Step 14 Part C: answer_challenge.xlsx (stage rewards only) ===")
    success_c = update_answer_challenge_excel(proj_root, src_evt, tgt_evt, log_dir, step_index)
    if not success_c:
        return False

    log("=== Step 14 Complete. ===")
    return True

def update_sys_excel(proj_root, src_evt, tgt_evt, log_dir, step_index):
    """Step 19.5/20: Update sys.xlsx — register DeepLink IDs in correct sections."""
    import openpyxl
    import stat
    import os
    import shutil
    
    design_dir = os.path.join(proj_root, "design", "DesignData", WORKSPACE_NAME, "design")
    sys_path = os.path.join(design_dir, "sys.xlsx")
    events_path = os.path.join(design_dir, "events.xlsx")
    tgt_evt_xlsx = os.path.join(design_dir, f"{tgt_evt}.xlsx")
    
    if not os.path.exists(sys_path):
        log(f"Error: sys.xlsx not found at {sys_path}")
        return False
        
    # 1. Fetch the New Achievement DeepLink ID from events.xlsx
    tgt_achievement_dl_id = None
    try:
        wb_evt = openpyxl.load_workbook(events_path, data_only=True)
        ws_a = wb_evt["活动成就"]
        for r in range(2, ws_a.max_row + 1):
            alias_val = str(ws_a.cell(r, 4).value or "").lower()
            if tgt_evt.lower() in alias_val:
                tgt_achievement_dl_id = ws_a.cell(r, 6).value
                break
        wb_evt.close()
    except Exception as e:
        log(f"sys.xlsx: Failed to read Achievement ID from events.xlsx: {e}")
        return False
        
    log(f"sys.xlsx: Target Achievement DeepLink ID: {tgt_achievement_dl_id}")

    # 2. Start Processing sys.xlsx
    backup_path = sys_path + ".bak"
    os.chmod(sys_path, stat.S_IWRITE | stat.S_IREAD)
    shutil.copy2(sys_path, backup_path)
    record_file_creation(log_dir, step_index, backup_path)
    
    src_lower, tgt_lower = src_evt.lower(), tgt_evt.lower()
    src_upper, tgt_upper = src_evt.upper(), tgt_evt.upper()
    src_pascal = src_evt[0].upper() + src_evt[1:] if src_evt else src_evt
    tgt_pascal = tgt_evt[0].upper() + tgt_evt[1:] if tgt_evt else tgt_evt

    def swap_name(val):
        if val is None or isinstance(val, (int, float)): return val
        s = str(val)
        s = s.replace(src_upper, tgt_upper).replace(src_lower, tgt_lower)
        s = s.replace(src_pascal, tgt_pascal).replace(src_evt, tgt_evt)
        return s

    try:
        # Step 2a: Scan for Max IDs and Dividers using data_only=True
        wb_scan = openpyxl.load_workbook(sys_path, data_only=True)
        ws_scan = wb_scan["DeepLink"]
        
        row_achieve_divider = None
        row_invite_divider = None
        max_share_id = 0
        src_share_rows = []
        src_achieve_row_idx = None

        for r in range(1, ws_scan.max_row + 1):
            val1 = ws_scan.cell(r, 1).value
            val2 = str(ws_scan.cell(r, 2).value or "") or ""
            
            # Robust ID extraction (handles both int and string IDs)
            id_num = None
            try:
                if val1 is not None:
                    id_num = int(float(str(val1)))
            except:
                pass
            
            # Normalize divider check
            s_val1 = str(val1).replace(" ", "") if val1 else ""
            if s_val1 == "//活动成就": row_achieve_divider = r
            if s_val1 == "//邀新": row_invite_divider = r
            
            # Find Source and Target Rows
            if src_evt.lower() in val2.lower():
                if id_num is not None:
                    if id_num < 20000: src_share_rows.append(r)
                    else: src_achieve_row_idx = r
            
            if tgt_evt.lower() in val2.lower():
                if id_num is not None:
                    # If target exists and is commented, we'll uncomment later
                    pass
            
            # Find Section Max IDs (only before the first divider)
            if id_num is not None:
                if row_achieve_divider is None and id_num < 20000:
                    max_share_id = max(max_share_id, id_num)
                    
        wb_scan.close()

        log(f"sys.xlsx Sections: ShareDivider={row_achieve_divider}, InviteDivider={row_invite_divider}")
        log(f"sys.xlsx MaxIDs: Share={max_share_id}")

        if not row_achieve_divider:
            log("Error: Could not find '// 活动成就' divider in sys.xlsx")
            return False

        # Step 2b: Open for writing
        wb = load_workbook_frozen(sys_path)
        ws = wb["DeepLink"]
        modified = False

        # A reopen under the same name (or a re-run) finds the target's Share group already in
        # sys.xlsx, and <tgt>.xlsx already uses exactly those IDs. Remapping it anyway pointed
        # every share link at IDs that are never created (Merge2510: 11810.. -> 12639.., none exist).
        tgt_share_exists = any(tgt_evt.lower() in str(ws.cell(r, 2).value or "").lower()
                               for r in range(1, row_achieve_divider + 1))
        if src_share_rows and tgt_share_exists:
            log(f"sys.xlsx: {tgt_evt}'s Share group already exists — {tgt_evt}.xlsx keeps its IDs (no remap).")

        # 3. Update Sim2605.xlsx if we have Share IDs to map
        if src_share_rows and not tgt_share_exists and not os.path.exists(tgt_evt_xlsx):
            log(f"Warning: {tgt_evt}.xlsx not found — skipping in-file Share ID remap. "
                f"That workbook will still reference the OLD (unshifted) DeepLink IDs until this step "
                f"is re-run after {tgt_evt}.xlsx exists (normally created by step 6).")

        if src_share_rows and not tgt_share_exists and os.path.exists(tgt_evt_xlsx):
            # Get the exact numeric values from scan result
            wb_ref = openpyxl.load_workbook(sys_path, data_only=True)
            ws_ref = wb_ref["DeepLink"]
            src_first_id = int(float(str(ws_ref.cell(src_share_rows[0], 1).value)))
            src_ids = [int(float(str(ws_ref.cell(r, 1).value))) for r in src_share_rows]
            wb_ref.close()

            new_first_id = max_share_id + 1
            offset = new_first_id - src_first_id
            
            # Use string versions for broad mapping in the event spreadsheet
            src_ids_str = [str(x) for x in src_ids]
            
            log(f"Mapping Share IDs: {src_ids[0]}..{src_ids[-1]} -> {new_first_id}.. (Offset: {offset})")
            
            wb_tgt = load_workbook_frozen(tgt_evt_xlsx)
            replaced_count = 0
            for sheet in wb_tgt.worksheets:
                for r in range(1, sheet.max_row + 1):
                    for c in range(1, sheet.max_column + 1):
                        v = sheet.cell(r, c).value
                        if v is not None:
                            try:
                                v_int = int(float(str(v)))
                                if v_int in src_ids:
                                    sheet.cell(r, c, v_int + offset)
                                    replaced_count += 1
                            except:
                                pass
            wb_tgt.save(tgt_evt_xlsx)
            log(f"Updated {tgt_evt}.xlsx: {replaced_count} IDs remapped.")

        # 4. Insert new rows into sys.xlsx
        # Part B: Achievement ID (before // 邀新)
        if row_invite_divider and src_achieve_row_idx:
            # Check if group already exists (normalize type for check)
            existing_row = None
            for r in range(row_achieve_divider, row_invite_divider + 1):
                val = str(ws.cell(r, 1).value or "").replace("//", "").strip()
                if val == str(tgt_achievement_dl_id):
                    existing_row = r
                    break
            
            if existing_row:
                val = str(ws.cell(existing_row, 1).value or "").strip()
                if val.startswith("//"):
                    ws.cell(existing_row, 1, val[2:].strip())
                    log(f"sys.xlsx: Uncommented Achievement ID {tgt_achievement_dl_id}")
                    modified = True
            else:
                ws.insert_rows(row_invite_divider)
                for c in range(1, ws.max_column + 1):
                    src_c = ws.cell(src_achieve_row_idx, c)
                    old_val = src_c.value
                    new_val = swap_name(old_val)
                    if c == 1: new_val = int(tgt_achievement_dl_id)
                    dst_c = ws.cell(row_invite_divider, c)
                    dst_c.value = new_val
                    copy_cell_format(src_c, dst_c)
                log(f"Inserted Achievement ID {tgt_achievement_dl_id} at row {row_invite_divider}")
                modified = True

        # Part A: Share Group (before // 活动成就)
        if row_achieve_divider and src_share_rows:
            new_first_id = max_share_id + 1
            # Check if group already exists (by event name in col 2)
            existing_target_share_rows = []
            for r in range(1, row_achieve_divider + 1):
                alias = str(ws.cell(r, 2).value or "").lower()
                if tgt_evt.lower() in alias:
                    existing_target_share_rows.append(r)
            
            if existing_target_share_rows:
                log(f"sys.xlsx: Target Share Group for {tgt_evt} already exists. Ensuring rows are uncommented.")
                for r in existing_target_share_rows:
                    val = str(ws.cell(r, 1).value or "").strip()
                    if val.startswith("//"):
                        ws.cell(r, 1, val[2:].strip())
                        modified = True
            else:
                wb_ref = openpyxl.load_workbook(sys_path, data_only=True)
                ws_ref = wb_ref["DeepLink"]
                src_ref_ids = [int(float(str(ws_ref.cell(r, 1).value))) for r in src_share_rows]
                offset = new_first_id - src_ref_ids[0]
                wb_ref.close()
                
                for i, r_idx in enumerate(src_share_rows):
                    new_row_pos = row_achieve_divider + i
                    ws.insert_rows(new_row_pos)
                    for c in range(1, ws.max_column + 1):
                        src_c = ws.cell(r_idx, c)
                        old_val = src_c.value
                        new_val = swap_name(old_val)
                        if c == 1: new_val = src_ref_ids[i] + offset
                        dst_c = ws.cell(new_row_pos, c)
                        dst_c.value = new_val
                        copy_cell_format(src_c, dst_c)
                log(f"Inserted {len(src_share_rows)} Share IDs starting at {new_first_id}")
                modified = True

        if modified:
            wb.save(sys_path)
            log("sys.xlsx successfully updated with sections and remapped IDs.")
        else:
            log("sys.xlsx: No changes were needed.")
        return True
    
    except Exception as e:
        log(f"Error updating sys.xlsx: {e}")
        import traceback
        traceback.print_exc()
        return False


# ===================== Trivia Challenge (答题挑战) cycle refresh =====================
# Unlike the numbered clone-event pipeline steps above, this is a standalone recurring
# maintenance action: it doesn't clone from a source event, it refreshes the quiz content
# for the upcoming 答题挑战 cycle using freshly-authored question data supplied as JSON.

def get_quiz_prefix_from_mini_mgr(proj_root):
    """Reads mini_mgr.xlsx's 活动数据 sheet and picks the 答题挑战 cycle to key the
    quiz content on, returning an ID prefix like 'QUIZ260827'.

    Preference order (dated rows only — the dateless template row never counts):
      1. a row that is switched on (控制开关=True): the next one on/after today,
         else the most recent one;
      2. otherwise the nearest upcoming row that is still switched off — the
         next cycle is normally scheduled before it is enabled, and refreshing
         its quiz content ahead of time is exactly the point of this step;
      3. otherwise None, with a message that says what to do about it.
    Returns None if no usable row is found."""
    import openpyxl
    from datetime import datetime, date

    design_dir = os.path.join(proj_root, "design", "DesignData", WORKSPACE_NAME, "design")
    file_path = os.path.join(design_dir, "mini_mgr.xlsx")
    if not os.path.exists(file_path):
        log(f"Error: mini_mgr.xlsx not found at {file_path}")
        return None

    wb = openpyxl.load_workbook(file_path, data_only=True)
    if "活动数据" not in wb.sheetnames:
        log("Error: 活动数据 sheet not found in mini_mgr.xlsx")
        return None
    ws = wb["活动数据"]
    try:
        cols = mini_mgr_columns([c.value for c in ws[1]])
    except ValueError as e:
        log(f"Error: {e}")
        return None

    def as_date(v):
        if isinstance(v, datetime):
            return v.date()
        if isinstance(v, date):
            return v
        return datetime.strptime(str(v).split(" ")[0], "%Y-%m-%d").date()

    dated = []  # (start_date, enabled, row)
    for r in range(1, ws.max_row + 1):
        start = ws.cell(r, cols["start"]).value
        if ws.cell(r, cols["type"]).value != "答题挑战" or start is None:  # type: ignore
            continue
        try:
            dated.append((as_date(start), ws.cell(r, cols["switch"]).value is True, r))  # type: ignore
        except ValueError:
            log(f"Warning: mini_mgr.xlsx row {r}: unreadable 答题挑战 start date {start!r}, skipped")

    if not dated:
        log("Error: mini_mgr.xlsx has no dated 答题挑战 rows at all — add the next minigame set (Minigame Configs tab) first")
        return None

    today = date.today()
    enabled = [d for d in dated if d[1]]
    if enabled:
        upcoming = [d for d in enabled if d[0] >= today]
        chosen = min(upcoming) if upcoming else max(enabled)
        note = "current"
    else:
        upcoming = [d for d in dated if d[0] >= today]
        if not upcoming:
            latest = max(dated)[0].strftime("%Y-%m-%d")
            log(f"Error: no upcoming 答题挑战 cycle scheduled in mini_mgr.xlsx — the latest is {latest} (already over, switched off). "
                f"Add the next minigame set (Minigame Configs tab) first, then run this again")
            return None
        chosen = min(upcoming)
        note = "upcoming, still switched off (控制开关=False)"

    dt, _, row = chosen
    prefix = f"QUIZ{dt.strftime('%y%m%d')}"
    log(f"mini_mgr.xlsx: {note} 答题挑战 cycle starts {dt.strftime('%Y-%m-%d')} (row {row}) -> prefix {prefix}")
    return prefix


# Column 9 of localization_quiz.xlsx's Quiz / Quiz_备份 sheets was always unused
# (verified against real data) — used here to stamp which pool batch number a cycle's
# content came from. Being a normal cell in the actual Excel file (not a side-channel
# state file), this marker is version-controlled and travels through SVN merges just
# like the questions themselves, so it stays meaningful across branches — unlike a
# local JSON file, which would only ever describe one machine's one checkout.
BATCH_MARKER_COL = 9


def _questions_module_for_batch(batch_number):
    """Batch content is deliberately NOT keyed by calendar date or QUIZ prefix —
    multiple branches (different teams) advance 答题挑战 cycles independently, on
    different real-world dates, so a date-keyed filename would force every branch to
    agree on a schedule. Content instead lives in a numbered pool (trivia_questions_
    batch_01.py, _02.py, ...); which number a given cycle uses is resolved by scanning
    every configured branch's own localization_quiz.xlsx (see resolve_questions_source),
    so branch A landing on batch 1 means branch B's next cycle gets batch 2, never a repeat —
    even though the two branches are on physically separate checkouts, possibly separate machines."""
    return f"trivia_questions_batch_{batch_number:02d}"


def batch_number_from_source(questions_source):
    """Batch number of a bundled module name ('trivia_questions_batch_05' -> 5), else None."""
    m = re.fullmatch(r"trivia_questions_batch_(\d+)", questions_source or "")
    return int(m.group(1)) if m else None


def available_batch_numbers():
    """Batch numbers that have a trivia_questions_batch_NN.py file next to this module."""
    import glob
    here = os.path.dirname(os.path.abspath(__file__))
    nums = (batch_number_from_source(os.path.splitext(os.path.basename(p))[0])
            for p in glob.glob(os.path.join(here, "trivia_questions_batch_*.py")))
    return sorted(n for n in nums if n is not None)


def scan_batch_usage(labeled_design_dirs):
    """{label: design_dir} -> [{label, live_prefix, live_batch, archived: [[prefix, batch], ...]}].
    live = the Quiz sheet's current cycle; archived = cycles moved to Quiz_备份. Batch
    markers sit on each cycle's first row (BATCH_MARKER_COL). Missing files are skipped."""
    import openpyxl
    # Cycles merged in without column 9 lose their marker; their first question title
    # still identifies the batch (unique per batch, and unchanged by option reshuffles).
    title_to_batch = {}
    for n in available_batch_numbers():
        q = load_trivia_questions(_questions_module_for_batch(n))[0]["question"]
        title_to_batch[q.get("en") if isinstance(q, dict) else q] = n
    usage = []
    for label, design_dir in labeled_design_dirs.items():
        path = os.path.join(design_dir, "localization_quiz.xlsx")
        if not os.path.exists(path):
            continue
        try:
            wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        except Exception as e:
            log(f"Warning: could not scan {path} for batch markers: {e}")
            continue
        entry = {"label": label, "live_prefix": None, "live_batch": None, "live_batch_inferred": False, "archived": []}
        for sheet_name in ("Quiz", "Quiz_备份"):
            if sheet_name not in wb.sheetnames:
                continue
            # read_only sheets are only fast when streamed (ws.cell() per row re-walks
            # the XML: 30s+ on the ~16k-row archive), so iter_rows in one pass.
            for row in wb[sheet_name].iter_rows(min_row=2, max_col=BATCH_MARKER_COL, values_only=True):
                m = re.match(r"(QUIZ\d+)_", str(row[0] or ""))
                marker = row[BATCH_MARKER_COL - 1] if len(row) >= BATCH_MARKER_COL else None
                batch = int(marker) if isinstance(marker, (int, float)) else None
                if sheet_name == "Quiz":
                    entry["live_prefix"] = m.group(1) if m else None
                    entry["live_batch"] = batch
                    if batch is None and row[1] in title_to_batch:
                        entry["live_batch"] = title_to_batch[row[1]]
                        entry["live_batch_inferred"] = True
                    break
                if batch is not None:
                    entry["archived"].append([m.group(1) if m else None, batch])
        usage.append(entry)
    return usage


def scan_max_used_batch_number(branch_design_dirs):
    """Highest batch marker in any given branch's Quiz / Quiz_备份, so a new cycle picks
    a batch nobody else has used yet. Accepts a list of dirs or a {label: dir} dict."""
    dirs = branch_design_dirs if isinstance(branch_design_dirs, dict) else {d: d for d in branch_design_dirs}
    return max_used_batch(scan_batch_usage(dirs))


def max_used_batch(usage):
    """Highest batch number in a scan_batch_usage() result (0 if none)."""
    used = [e["live_batch"] for e in usage if e["live_batch"]]
    used += [b for e in usage for _, b in e["archived"]]
    return max(used, default=0)


def quiz_cycle_is_current(design_dir, prefix, batch_number, questions=None):
    """True when the live Quiz sheet already holds `prefix` with translations, the same
    batch marker (any marker, when batch_number is None) and — when `questions` is given —
    the same English text, so an edited batch file (e.g. reshuffled options) still counts
    as needing a rewrite."""
    import openpyxl
    path = os.path.join(design_dir, "localization_quiz.xlsx")
    if not os.path.exists(path):
        return False
    ws = openpyxl.load_workbook(path, read_only=True, data_only=True)["Quiz"]
    n_rows = 3 * len(questions) if questions else 1
    rows = [list(r) + [None] * (BATCH_MARKER_COL - len(r))
            for r in ws.iter_rows(min_row=2, max_row=1 + n_rows, max_col=BATCH_MARKER_COL, values_only=True)]
    if not rows:
        return False
    m = re.match(r"(QUIZ\d+)_", str(rows[0][0] or ""))
    marker = rows[0][BATCH_MARKER_COL - 1]
    if not (m and m.group(1) == prefix and rows[0][2] and isinstance(marker, (int, float))):
        return False
    if batch_number is not None and int(marker) != batch_number:
        return False
    if questions:
        en = lambda v: v.get("en", "") if isinstance(v, dict) else v
        expected = [en(q[k]) for q in questions for k in ("question", "opt1", "opt2")]
        if [r[1] for r in rows] != expected:
            return False
    return True


def resolve_questions_source(own_design_dir, all_branch_design_dirs, prefix):
    """Returns (module_name, batch_number) for this cycle. If this branch's OWN live
    Quiz sheet already shows this exact prefix (re-run, or backfilling a missing
    translation column), reuses whatever batch it used last time via that row's own
    marker — it does NOT re-scan, so a repeat call never advances. Only a genuinely
    new prefix triggers scan_max_used_batch_number() across every configured branch
    to pick the next never-used-by-anyone batch."""
    import openpyxl
    import re

    quiz_path = os.path.join(own_design_dir, "localization_quiz.xlsx")
    if os.path.exists(quiz_path):
        wb = openpyxl.load_workbook(quiz_path, read_only=True, data_only=True)
        ws = wb["Quiz"]
        live_id = str(ws.cell(2, 1).value or "")
        live_prefix_match = re.match(r"(QUIZ\d+)_", live_id)
        if live_prefix_match and live_prefix_match.group(1) == prefix:
            marker = ws.cell(2, BATCH_MARKER_COL).value
            if isinstance(marker, (int, float)):
                batch_number = int(marker)
                return _questions_module_for_batch(batch_number), batch_number
            # Content for this exact prefix already exists but predates the batch-marker
            # column (written by an older version of this tool) — it's still batch 1,
            # the only batch that existed before this feature. Don't treat it as new.
            log(f"{prefix} content has no batch marker (pre-dates batch tracking) — assuming batch 1.")
            return _questions_module_for_batch(1), 1

    batch_number = scan_max_used_batch_number(all_branch_design_dirs) + 1
    return _questions_module_for_batch(batch_number), batch_number


def load_trivia_questions(questions_source):
    """Loads a question list. `questions_source` is either a path to a JSON file
    (schema: {"questions": [{"category", "question", "opt1", "opt2",
    "correct_index" (1 or 2), "tier" (1/2/3)}, ...]}), or the bare module name of a
    bundled trivia_questions_batch_NN.py module exposing a QUESTIONS list in that
    same per-question shape. Modules are loaded via a plain dynamic import — this
    tool's Python backend ships as raw .py files next to a real interpreter, not a
    frozen PyInstaller executable, so a normal filesystem-based import just works."""
    if questions_source.endswith(".json"):
        with open(questions_source, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data["questions"]
    import importlib
    try:
        module = importlib.import_module(questions_source)
    except ModuleNotFoundError:
        raise ModuleNotFoundError(
            f"{questions_source}.py not found — author this batch's 100 fresh trivia "
            f"questions into a module of that name (see trivia_questions_batch_01.py for "
            f"the expected shape) before running this tool."
        )
    return module.QUESTIONS


def update_localization_quiz_excel(proj_root, log_dir, step_index, questions_source=None,
                                    all_branch_design_dirs=None, batch_number=None, force=False):
    """Refreshes localization_quiz.xlsx's 'Quiz' sheet for the upcoming 答题挑战 cycle:
    archives the current live batch into 'Quiz_备份' (skipped if already archived),
    then writes the fresh batch (all 7 language columns) under the new cycle's ID prefix,
    stamping BATCH_MARKER_COL with the batch number used. `questions_source`/`batch_number`
    default to this cycle's next never-used-anywhere batch (see resolve_questions_source) —
    pass them explicitly only to override (e.g. testing, or a manual re-run)."""
    import openpyxl

    prefix = get_quiz_prefix_from_mini_mgr(proj_root)
    if not prefix:
        return False
    design_dir = os.path.join(proj_root, "design", "DesignData", WORKSPACE_NAME, "design")
    if not questions_source:
        questions_source, batch_number = resolve_questions_source(design_dir, all_branch_design_dirs or [], prefix)

    try:
        questions = load_trivia_questions(questions_source)
    except ModuleNotFoundError as e:
        log(f"Error: {e}")
        return False

    file_path = os.path.join(design_dir, "localization_quiz.xlsx")
    if not os.path.exists(file_path):
        log(f"Error: localization_quiz.xlsx not found at {file_path}")
        return False

    import stat
    backup_path = file_path + ".bak"
    os.chmod(file_path, stat.S_IWRITE | stat.S_IREAD)
    shutil.copy2(file_path, backup_path)
    record_file_creation(log_dir, step_index, backup_path)

    try:
        wb = load_workbook_frozen(file_path)
        if "Quiz" not in wb.sheetnames or "Quiz_备份" not in wb.sheetnames:
            log("Error: 'Quiz' or 'Quiz_备份' sheet not found in localization_quiz.xlsx")
            return False
        ws_live = wb["Quiz"]
        ws_backup = wb["Quiz_备份"]

        live_id = str(ws_live.cell(2, 1).value or "")
        live_prefix_match = re.match(r"(QUIZ\d+)_", live_id)
        live_prefix = live_prefix_match.group(1) if live_prefix_match else None

        if live_prefix == prefix:
            # Content already matches this cycle — but if it was written by an older
            # version of this tool (missing translations and/or the batch marker column),
            # re-run only the write step (same content, same row positions) to backfill,
            # rather than silently leaving those columns incomplete forever.
            has_translations = bool(ws_live.cell(2, 3).value)  # German column
            live_marker = ws_live.cell(2, BATCH_MARKER_COL).value
            has_marker = isinstance(live_marker, (int, float))
            same_batch = batch_number is None or (has_marker and int(live_marker) == batch_number)
            if has_translations and has_marker and same_batch and not force:
                log(f"Quiz sheet already contains {prefix} data with translations and batch marker. Nothing to do.")
                return True
            if force and has_marker and same_batch:
                log(f"Quiz sheet has {prefix} as batch {batch_number} but its text differs from the batch file — rewriting.")
            elif has_marker and not same_batch:
                log(f"Quiz sheet has {prefix} as batch {int(live_marker)} — replacing it with batch {batch_number}.")
            else:
                log(f"Quiz sheet has {prefix} data but is missing translations and/or the batch marker — backfilling.")
        elif live_prefix:
            already_archived = any(
                str(ws_backup.cell(r, 1).value or "").startswith(live_prefix + "_")
                for r in range(2, ws_backup.max_row + 1)
            )
            if already_archived:
                log(f"{live_prefix} already present in Quiz_备份. Skipping archive.")
            else:
                num_cols = max(ws_live.max_column, BATCH_MARKER_COL)
                insert_at = ws_backup.max_row + 1
                copied = 0
                for r in range(2, ws_live.max_row + 1):
                    if ws_live.cell(r, 1).value is None:
                        continue
                    for c in range(1, num_cols + 1):
                        src_c = ws_live.cell(r, c)
                        dst_c = ws_backup.cell(insert_at + copied, c)
                        dst_c.value = src_c.value
                        copy_cell_format(src_c, dst_c)
                    copied += 1
                log(f"Archived {copied} rows of {live_prefix} to Quiz_备份 starting at row {insert_at}.")
        else:
            log("Quiz sheet has no existing ID prefix to archive (empty or unrecognized). Proceeding.")

        # Clear the live sheet's data rows, then write the fresh batch.
        for r in range(2, ws_live.max_row + 1):
            for c in range(1, max(ws_live.max_column, BATCH_MARKER_COL) + 1):
                ws_live.cell(r, c).value = None

        # Sheet columns: 1=key, 2=English, 3=German, 4=Spanish, 5=Dutch, 6=Swedish, 7=French, 8=Danish.
        LANG_COLS = [(2, "en"), (3, "de"), (4, "es"), (5, "nl"), (6, "sv"), (7, "fr"), (8, "da")]

        def cell_text(field_value, lang):
            # Supports both the translated {"en": ..., "de": ..., ...} shape and a
            # plain string (English-only, left blank in other languages) for backward compat.
            if isinstance(field_value, dict):
                return field_value.get(lang) or field_value.get("en", "")
            return field_value if lang == "en" else None

        row = 2
        for i, q in enumerate(questions):
            qnum = i + 1
            ws_live.cell(row, 1).value = f"{prefix}_Q_{qnum}_TITLE"
            ws_live.cell(row + 1, 1).value = f"{prefix}_Q_{qnum}_OPT1"
            ws_live.cell(row + 2, 1).value = f"{prefix}_Q_{qnum}_OPT2"
            for col, lang in LANG_COLS:
                ws_live.cell(row, col).value = cell_text(q["question"], lang)
                ws_live.cell(row + 1, col).value = cell_text(q["opt1"], lang)
                ws_live.cell(row + 2, col).value = cell_text(q["opt2"], lang)
            row += 3

        # Stamp which pool batch this cycle used, so resolve_questions_source (this
        # branch re-running the same cycle) and scan_max_used_batch_number (a different
        # branch picking its own next cycle) can both find it later.
        if batch_number is not None:
            ws_live.cell(2, BATCH_MARKER_COL).value = batch_number

        wb.save(file_path)
        log(f"localization_quiz.xlsx: wrote {len(questions)} fresh questions under {prefix} (batch {batch_number}).")
        return True

    except Exception as e:
        log(f"Error updating localization_quiz.xlsx: {e}")
        import traceback
        traceback.print_exc()
        return False


def update_answer_challenge_quiz_excel(proj_root, log_dir, step_index, questions_source=None, force=False):
    """Refreshes answer_challenge.xlsx's '题目' sheet for the upcoming 答题挑战 cycle:
    replaces the question list with categories/correct answers matching the freshly
    written localization_quiz.xlsx batch, preserving the tier ID-range convention
    (难度库: tier1=1001+, tier2=2001+, tier3=3001+) and appending in tier order.
    Called via run_trivia_challenge_cycle, which always resolves and passes an explicit
    `questions_source` so both files use the exact same batch — the fallback here (this
    branch only, no cross-branch scan) only matters for a standalone/manual call."""
    import openpyxl

    prefix = get_quiz_prefix_from_mini_mgr(proj_root)
    if not prefix:
        return False
    design_dir = os.path.join(proj_root, "design", "DesignData", WORKSPACE_NAME, "design")
    if not questions_source:
        questions_source, _ = resolve_questions_source(design_dir, [], prefix)

    try:
        questions = load_trivia_questions(questions_source)
    except ModuleNotFoundError as e:
        log(f"Error: {e}")
        return False

    file_path = os.path.join(design_dir, "answer_challenge.xlsx")
    if not os.path.exists(file_path):
        log(f"Error: answer_challenge.xlsx not found at {file_path}")
        return False

    import stat
    backup_path = file_path + ".bak"
    os.chmod(file_path, stat.S_IWRITE | stat.S_IREAD)
    shutil.copy2(file_path, backup_path)
    record_file_creation(log_dir, step_index, backup_path)

    try:
        wb = load_workbook_frozen(file_path)
        if "题目" not in wb.sheetnames:
            log("Error: '题目' sheet not found in answer_challenge.xlsx")
            return False
        ws = wb["题目"]

        # force: localization_quiz.xlsx was just rewritten, so correct_index must follow it
        # even when the prefix is unchanged (another batch, or a backfill).
        existing_id = str(ws.cell(2, 6).value or "")
        if existing_id.startswith(prefix + "_") and not force:
            log(f"题目 sheet already references {prefix}. Nothing to do.")
            return True

        # Preserve header (row 1), replace all data rows.
        for r in range(2, ws.max_row + 1):
            for c in range(1, ws.max_column + 1):
                ws.cell(r, c).value = None

        tier_id_base = {1: 1001, 2: 2001, 3: 3001}
        tier_counters = {1: 0, 2: 0, 3: 0}
        ordered = sorted(enumerate(questions, start=1), key=lambda iq: iq[1]["tier"])

        row = 2
        sort_order = 1
        for qnum, q in ordered:
            tier = q["tier"]
            question_id = tier_id_base[tier] + tier_counters[tier]
            tier_counters[tier] += 1
            ws.cell(row, 1).value = question_id
            ws.cell(row, 2).value = sort_order
            ws.cell(row, 3).value = q["category"]
            ws.cell(row, 4).value = tier
            ws.cell(row, 5).value = q["correct_index"]
            ws.cell(row, 6).value = f"{prefix}_Q_{qnum}_TITLE"
            ws.cell(row, 7).value = f"{prefix}_Q_{qnum}_OPT1"
            ws.cell(row, 8).value = f"{prefix}_Q_{qnum}_OPT2"
            row += 1
            sort_order += 1

        wb.save(file_path)
        log(f"answer_challenge.xlsx [题目]: wrote {len(questions)} rows referencing {prefix}.")
        return True

    except Exception as e:
        log(f"Error updating answer_challenge.xlsx: {e}")
        import traceback
        traceback.print_exc()
        return False


def run_trivia_challenge_cycle(proj_root, log_dir, step_index, questions_source=None,
                                all_branch_design_dirs=None):
    """Orchestrates the full Trivia Challenge cycle refresh: localization_quiz.xlsx
    then answer_challenge.xlsx, both keyed off the same mini_mgr-derived prefix and
    the same resolved batch. `all_branch_design_dirs` should list every configured
    branch's design directory (trunk/Misc/Feature/...) so a genuinely new cycle picks
    the next batch nobody else has used yet — omit it and only this branch's own
    history is considered (fine for a single-branch setup, wrong for a multi-team one)."""
    prefix = get_quiz_prefix_from_mini_mgr(proj_root)
    if not prefix:
        return False
    design_dir = os.path.join(proj_root, "design", "DesignData", WORKSPACE_NAME, "design")
    if not questions_source:
        questions_source, batch_number = resolve_questions_source(design_dir, all_branch_design_dirs or [], prefix)
    else:
        # A picked pool batch is stamped like an auto one; a JSON file stays untracked.
        batch_number = batch_number_from_source(questions_source)

    try:
        questions = load_trivia_questions(questions_source)
    except ModuleNotFoundError as e:
        log(f"Error: {e}")
        return False

    # Both workbooks are skipped or rewritten together: a half rewrite would leave
    # 题目's correct_index pointing at the other option.
    if quiz_cycle_is_current(design_dir, prefix, batch_number, questions):
        log(f"{prefix} already uses batch {batch_number}. Nothing to do.")
        return True
    if not update_localization_quiz_excel(proj_root, log_dir, step_index, questions_source,
                                          batch_number=batch_number, force=True):
        return False
    if not update_answer_challenge_quiz_excel(proj_root, log_dir, step_index, questions_source, force=True):
        return False
    return True


# ================================================================================
# Step 22 (docx "Reopening" Step 11): register the event across the Lua config
# files EventConfig.lua, BIDef.lua, UIManager.lua, and PanelConfig.lua. Each file
# was inspected against a real event (merge2608) before writing this — every one
# is a flat, single-block-per-event structure with no cross-references between
# files, so each can be handled independently and idempotently.
#
# GlobalEvent.lua (docx also names it) is deliberately NOT automated here — the
# docx itself says "For Production events, you need to add..."; checked both
# merge2606 and merge2608 (real dev/test events) and neither has an entry there,
# confirming it's conditional/production-only, not a per-event requirement.
# ================================================================================

def find_lua_comment_block_bounds(lines, marker_text):
    """Find a '-- <marker_text>' comment line and the block of entries that follows
    it, bounded by the next blank line or the next '--' comment line. Returns
    (marker_line_idx, block_start_idx, block_end_idx_exclusive), or (None, None, None)
    if no such marker line exists."""
    marker_line_idx = None
    for i, line in enumerate(lines):
        stripped = line.strip()
        if stripped.startswith('--') and stripped.lstrip('-').strip() == marker_text:
            marker_line_idx = i
            break
    if marker_line_idx is None:
        return None, None, None
    block_start = marker_line_idx + 1
    block_end = len(lines)
    for i in range(block_start, len(lines)):
        s = lines[i].strip()
        if s == '' or s.startswith('--'):
            block_end = i
            break
    return marker_line_idx, block_start, block_end


def update_event_config_lua(proj_root, src_evt, tgt_evt, log_dir, step_index):
    """EventConfig.lua: <Event> = { Name = "...", CtrlPath = "...", Enable = ... }"""
    import re

    file_path = os.path.join(proj_root, "develop", "client", "Skipbo", "Assets", "Lua", "Game", "Module", "Event", "EventConfig.lua")
    if not os.path.exists(file_path):
        log(f"Warning: EventConfig.lua not found at {file_path}. Skipping.")
        return True

    src_pascal = src_evt[0].upper() + src_evt[1:] if src_evt else src_evt
    tgt_pascal = tgt_evt[0].upper() + tgt_evt[1:] if tgt_evt else tgt_evt
    src_upper, tgt_upper = src_evt.upper(), tgt_evt.upper()
    src_lower, tgt_lower = src_evt.lower(), tgt_evt.lower()

    def swap_name(s):
        if s is None:
            return None
        s = str(s)
        s = s.replace(src_upper, tgt_upper)
        s = s.replace(src_lower, tgt_lower)
        s = s.replace(src_pascal, tgt_pascal)
        s = s.replace(src_evt, tgt_evt)
        return s

    with open(file_path, 'r', encoding='utf-8') as f:
        content = f.read()

    existing = re.search(rf'(?m)^[ \t]*{re.escape(tgt_pascal)}\s*=\s*\{{[^}}]*\}}', content)
    if existing:
        # Reopening an event under its own name finds its old entry, which was switched off
        # when the event closed (Merge2510: Enable = false). Same docx rule as a new event.
        if not re.search(r'Enable\s*=\s*false', existing.group(0)):
            log(f"EventConfig.lua: {tgt_pascal} already registered and enabled. Skipping.")
            return True
        backup_path = file_path + ".bak"
        shutil.copy2(file_path, backup_path)
        record_file_creation(log_dir, step_index, backup_path)
        block = re.sub(r'Enable\s*=\s*false', 'Enable = true', existing.group(0))
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(content[:existing.start()] + block + content[existing.end():])
        log(f"EventConfig.lua: {tgt_pascal} was already registered but disabled — set Enable = true.")
        return True

    match = re.search(rf'([ \t]*{re.escape(src_pascal)}\s*=\s*\{{[^}}]*\}},?)', content, re.DOTALL)
    if not match:
        log(f"Warning: EventConfig.lua has no '{src_pascal} = {{...}}' block for source event. Skipping.")
        return True

    new_block = swap_name(match.group(1))
    # Docx is explicit: "make sure the Enable setting is set to true" — the source's
    # own value is often false (a closed event), so this must be forced, not inherited.
    new_block = re.sub(r'Enable\s*=\s*(true|false)', 'Enable = true', new_block)

    backup_path = file_path + ".bak"
    shutil.copy2(file_path, backup_path)
    record_file_creation(log_dir, step_index, backup_path)

    insert_pos = match.end()
    new_content = content[:insert_pos] + "\n" + new_block + content[insert_pos:]
    with open(file_path, 'w', encoding='utf-8') as f:
        f.write(new_content)

    log(f"EventConfig.lua: Registered {tgt_pascal} (cloned from {src_pascal}).")
    return True


def update_bidef_lua(proj_root, src_evt, tgt_evt, log_dir, step_index):
    """BIDef.lua: Event_<Event> = "Event_<Event>","""
    import re

    file_path = os.path.join(proj_root, "develop", "client", "Skipbo", "Assets", "Lua", "Game", "Module", "BI", "BIDef.lua")
    if not os.path.exists(file_path):
        log(f"Warning: BIDef.lua not found at {file_path}. Skipping.")
        return True

    src_pascal = src_evt[0].upper() + src_evt[1:] if src_evt else src_evt
    tgt_pascal = tgt_evt[0].upper() + tgt_evt[1:] if tgt_evt else tgt_evt

    with open(file_path, 'r', encoding='utf-8') as f:
        content = f.read()

    if re.search(rf'(?m)^\s*Event_{re.escape(tgt_pascal)}\s*=', content):
        log(f"BIDef.lua: Event_{tgt_pascal} already registered. Skipping.")
        return True

    match = re.search(rf'([ \t]*)Event_{re.escape(src_pascal)}\s*=\s*"Event_{re.escape(src_pascal)}",', content)
    if not match:
        log(f"Warning: BIDef.lua has no 'Event_{src_pascal}' entry for source event. Skipping.")
        return True

    indent = match.group(1)
    new_line = f'{indent}Event_{tgt_pascal} = "Event_{tgt_pascal}",'

    backup_path = file_path + ".bak"
    shutil.copy2(file_path, backup_path)
    record_file_creation(log_dir, step_index, backup_path)

    insert_pos = match.end()
    new_content = content[:insert_pos] + "\n" + new_line + content[insert_pos:]
    with open(file_path, 'w', encoding='utf-8') as f:
        f.write(new_content)

    log(f"BIDef.lua: Registered Event_{tgt_pascal} (cloned from Event_{src_pascal}).")
    return True


def _special_panel_lines_from_svn(file_path, src_pascal, hint_date):
    """The source's SpecialPanelMap line(s) from UIManager.lua's SVN history. They're removed
    once an event closes (Merge2510's was), so reopening an older event has nothing live to
    clone. Date-bracketed by the source's start, newest revision first."""
    revs = get_svn_revisions_in_date_range(file_path, hint_date) if hint_date else None
    for rev in revs or get_svn_log_revisions(file_path, limit=60) or []:
        old = svn_cat_revision(file_path, rev) or ''
        m = re.search(r'local\s+SpecialPanelMap\s*=\s*\{(.*?)\n[ \t]*\}', old, re.DOTALL)
        lines = re.findall(rf'[ \t]*\["UI{re.escape(src_pascal)}\w*"\]\s*=\s*true,', m.group(1)) if m else []
        if lines:
            log(f"UIManager.lua: recovered UI{src_pascal}*'s SpecialPanelMap line(s) from r{rev}.")
            return lines
    return []


def update_ui_manager_special_panel_map(proj_root, src_evt, tgt_evt, log_dir, step_index):
    """UIManager.lua's SpecialPanelMap: ["UI<Event>...Panel"] = true, — suffix
    (LoadingPanel vs StageLoadingPanel etc.) varies per event, so clone whatever
    line(s) the source actually has rather than assuming one fixed suffix."""
    import re

    file_path = os.path.join(proj_root, "develop", "client", "Skipbo", "Assets", "Lua", "Framework", "Base", "UIManager.lua")
    if not os.path.exists(file_path):
        log(f"Warning: UIManager.lua not found at {file_path}. Skipping.")
        return True

    src_pascal = src_evt[0].upper() + src_evt[1:] if src_evt else src_evt
    tgt_pascal = tgt_evt[0].upper() + tgt_evt[1:] if tgt_evt else tgt_evt

    with open(file_path, 'r', encoding='utf-8') as f:
        content = f.read()

    table_match = re.search(r'(local\s+SpecialPanelMap\s*=\s*\{)(.*?)(\n[ \t]*\})', content, re.DOTALL)
    if not table_match:
        log("Warning: UIManager.lua has no 'SpecialPanelMap' table. Skipping.")
        return True

    table_body = table_match.group(2)

    if re.search(rf'\["UI{re.escape(tgt_pascal)}\w*"\]', table_body):
        log(f"UIManager.lua: UI{tgt_pascal}*Panel already registered in SpecialPanelMap. Skipping.")
        return True

    src_lines = re.findall(rf'[ \t]*\["UI{re.escape(src_pascal)}\w*"\]\s*=\s*true,', table_body)
    if not src_lines:
        src_lines = _special_panel_lines_from_svn(file_path, src_pascal, load_extra_args(log_dir).get("src_start_time", ""))
    if not src_lines:
        log(f"Warning: UIManager.lua's SpecialPanelMap has no entry for UI{src_pascal}* (live or SVN). Skipping.")
        return True

    new_lines = [line.replace(src_pascal, tgt_pascal) for line in src_lines]

    backup_path = file_path + ".bak"
    shutil.copy2(file_path, backup_path)
    record_file_creation(log_dir, step_index, backup_path)

    insert_pos = table_match.start(3)  # right before the closing "\n\t}"
    new_content = content[:insert_pos] + "\n" + "\n".join(new_lines) + content[insert_pos:]
    with open(file_path, 'w', encoding='utf-8') as f:
        f.write(new_content)

    log(f"UIManager.lua: Registered {len(new_lines)} SpecialPanelMap entr{'y' if len(new_lines) == 1 else 'ies'} for {tgt_pascal}.")
    return True


def _extract_panel_config_entries(content, token):
    """
    Extract every PanelConfig.lua entry whose key embeds `token` (e.g. 'Sim2607'),
    from BOTH real table shapes in this file:
      - PanelConfig.Id: 1-line entries, ["UI<token>XPanel"] = "UI<token>XPanel",
      - PanelConfig.Info: 4-line stanzas with ResPath/ViewPath.
    (?!\\d) after the token stops 'Sim260' from matching inside 'Sim2607'.
    Returns (id_entries, info_entries) — lists of exact source text per entry.
    """
    import re as _re
    token_re = _re.escape(token)
    id_pattern = _re.compile(rf'^[ \t]*\["UI{token_re}(?!\d)[A-Za-z]*"\]\s*=\s*"UI{token_re}(?!\d)[A-Za-z]*",?[ \t]*$', _re.MULTILINE)
    info_pattern = _re.compile(rf'[ \t]*\["UI{token_re}(?!\d)[A-Za-z]*"\]\s*=\s*\{{\s*\n\s*ResPath[^\n]*\n\s*ViewPath[^\n]*\n\s*\}},?', _re.MULTILINE)
    return id_pattern.findall(content), info_pattern.findall(content)


def _find_panel_config_table_close(lines, table_name):
    """Locate the standalone '}' line that closes 'PanelConfig.<table_name> = { ... }',
    by finding the table's own start and the next top-level 'PanelConfig.<X> = ' after
    it, then walking back from there to the nearest bare '}'. Returns a 0-indexed line
    number, or None if the table can't be found."""
    import re as _re
    start_idx = None
    for i, line in enumerate(lines):
        if _re.match(rf'^PanelConfig\.{table_name}\s*=\s*\{{', line):
            start_idx = i + 1
            break
    if start_idx is None:
        return None
    next_decl_idx = len(lines)
    for i in range(start_idx, len(lines)):
        if _re.match(r'^PanelConfig\.\w+\s*=', lines[i]):
            next_decl_idx = i
            break
    for i in range(next_decl_idx - 1, start_idx - 1, -1):
        if lines[i].strip() == '}':
            return i
    return None


def get_panel_config_entries_for_source(file_path, src_evt, hint_date, log_dir):
    """
    Resolve src_evt's PanelConfig.lua entries with the minimum SVN cost — the same
    live -> cache -> SVN-history pattern as convert_layout.lua's resolver:
      1. Still live in the current file? Use it directly, refresh the cache.
      2. Previously cached? Use that.
      3. Otherwise, search SVN history (date-bracketed by hint_date when available).

    This matters specifically for Blocks/Sim: unlike Merge/Dessert/Sort (whose old
    blocks stay in the file even after a newer event registers), Blocks/Sim entries
    get renamed in place each time the family cycles — confirmed: Sim2605 has zero
    entries left anywhere in the live file, only Sim2607's. Reopening an older
    Blocks/Sim event has nothing live to clone from and needs its SVN history.
    """
    src_pascal = src_evt[0].upper() + src_evt[1:] if src_evt else src_evt

    with open(file_path, 'r', encoding='utf-8') as f:
        live_content = f.read()

    id_entries, info_entries = _extract_panel_config_entries(live_content, src_pascal)
    if id_entries or info_entries:
        cache = load_named_cache(log_dir, "panelconfig_cache")
        cache[src_evt.lower()] = {"id_entries": id_entries, "info_entries": info_entries, "source": "live", "cached_at": _cache_timestamp()}
        save_named_cache(log_dir, "panelconfig_cache", cache)
        return id_entries, info_entries, "current live file"

    cache = load_named_cache(log_dir, "panelconfig_cache")
    cached = cache.get(src_evt.lower())
    if cached and (cached.get("id_entries") or cached.get("info_entries")):
        return (cached.get("id_entries", []), cached.get("info_entries", []),
                f"local cache (from {cached.get('source', '?')}, cached {cached.get('cached_at', '?')})")

    revisions = get_svn_revisions_in_date_range(file_path, hint_date) if hint_date else None
    used_date_bracket = bool(revisions)
    if not revisions:
        revisions = get_svn_log_revisions(file_path, limit=60)
    if not revisions:
        log(f"PanelConfig.lua: no SVN history available to search for {src_pascal}.")
        return [], [], None

    if used_date_bracket:
        log(f"PanelConfig.lua: {src_pascal} not live — date-bracketed SVN search around {hint_date.strip()[:10]} ({len(revisions)} revision(s) to check).")
    else:
        log(f"PanelConfig.lua: {src_pascal} not live — scanning the last {len(revisions)} revision(s) (no reliable date hint available).")

    for rev in revisions:
        hist_content = svn_cat_revision(file_path, rev)
        if hist_content is None:
            continue
        hist_id, hist_info = _extract_panel_config_entries(hist_content, src_pascal)
        if hist_id or hist_info:
            cache[src_evt.lower()] = {"id_entries": hist_id, "info_entries": hist_info, "source": f"svn r{rev}", "cached_at": _cache_timestamp()}
            save_named_cache(log_dir, "panelconfig_cache", cache)
            log(f"PanelConfig.lua: found {src_pascal} at r{rev} ({len(hist_id)} Id entries, {len(hist_info)} Info entries).")
            return hist_id, hist_info, f"SVN history (r{rev}, freshly fetched and cached)"

    log(f"PanelConfig.lua: no revision within the {len(revisions)} checked contained {src_pascal} entries.")
    return [], [], None


def update_panel_config_lua(proj_root, src_evt, tgt_evt, log_dir, step_index):
    """
    PanelConfig.lua: registers the target event across BOTH of its tables —
    PanelConfig.Id (key list) and PanelConfig.Info (ResPath/ViewPath). A prior
    version of this function found its clone source via a '-- <Event>' comment
    marker, which only ever appears once per event (in the Id table) — so it
    silently cloned Id entries only and never touched Info, meaning a cloned
    event would get a panel ID with no ResPath/ViewPath behind it. Fixed by
    extracting matching entries from both tables directly by digit-token, which
    also works for Blocks/Sim, whose entries were never given a comment marker
    at all.

    Source lookup now goes live -> local cache -> SVN history (see
    get_panel_config_entries_for_source) instead of failing outright when the
    source event's entries have already been renamed away by a later event.

    Note: the docx names PanelConfig_Design.lua for this, but that file is
    actually an empty, unused override stub (AppendId/AppendInfo both {}) —
    verified against the real file. The real per-event blocks live in the base
    PanelConfig.lua.
    """
    file_path = os.path.join(proj_root, "develop", "client", "Skipbo", "Assets", "Lua", "Global", "PanelConfig.lua")
    if not os.path.exists(file_path):
        log(f"Warning: PanelConfig.lua not found at {file_path}. Skipping.")
        return True

    src_pascal = src_evt[0].upper() + src_evt[1:] if src_evt else src_evt
    tgt_pascal = tgt_evt[0].upper() + tgt_evt[1:] if tgt_evt else tgt_evt
    src_upper, tgt_upper = src_evt.upper(), tgt_evt.upper()
    src_lower, tgt_lower = src_evt.lower(), tgt_evt.lower()

    def swap_name(s):
        if s is None:
            return None
        s = str(s)
        s = s.replace(src_upper, tgt_upper)
        s = s.replace(src_lower, tgt_lower)
        s = s.replace(src_pascal, tgt_pascal)
        s = s.replace(src_evt, tgt_evt)
        return s

    with open(file_path, 'r', encoding='utf-8') as f:
        content = f.read()

    existing_id, existing_info = _extract_panel_config_entries(content, tgt_pascal)
    if existing_id or existing_info:
        log(f"PanelConfig.lua: {tgt_pascal} already present ({len(existing_id)} Id, {len(existing_info)} Info entries). Skipping.")
        return True

    extra = load_extra_args(log_dir)
    hint_date = extra.get("src_start_time", "")

    src_id_entries, src_info_entries, source_label = get_panel_config_entries_for_source(file_path, src_evt, hint_date, log_dir)
    if not src_id_entries and not src_info_entries:
        log(f"Warning: PanelConfig.lua has no {src_pascal} entries — not live, and none found in SVN history. Skipping.")
        return True

    new_id_lines = [swap_name(e) for e in src_id_entries]
    new_info_blocks = [swap_name(e) for e in src_info_entries]

    lines = content.split('\n')
    id_close = _find_panel_config_table_close(lines, "Id")
    info_close = _find_panel_config_table_close(lines, "Info")
    if id_close is None or info_close is None:
        log("Warning: PanelConfig.lua's Id/Info table boundaries couldn't be located. Skipping.")
        return True

    backup_path = file_path + ".bak"
    shutil.copy2(file_path, backup_path)
    record_file_creation(log_dir, step_index, backup_path)

    # Info's insertion point sits at a later line than Id's — insert there first so
    # the already-computed id_close index isn't shifted out from under it.
    if new_info_blocks:
        lines = lines[:info_close] + ['', f"    -- {tgt_pascal}"] + new_info_blocks + lines[info_close:]
    if new_id_lines:
        lines = lines[:id_close] + ['', f"-- {tgt_pascal}"] + new_id_lines + lines[id_close:]

    with open(file_path, 'w', encoding='utf-8') as f:
        f.write('\n'.join(lines))

    log(f"PanelConfig.lua: registered {tgt_pascal} — {len(new_id_lines)} Id entries, {len(new_info_blocks)} Info entries (source: {source_label}).")
    return True


def update_guide_refui_mgr(proj_root, src_evt, tgt_evt, log_dir, step_index):
    """
    Step 13 (docx): GuideRefuiMgr.lua's guide UI-hook functions.

    Two different real mechanisms exist in this file (confirmed by inspecting it
    directly — the docx doesn't describe this split):

      - Merge: one shared "mergeFix" region with generically-named functions
        (MERGE_FIX_*) whose bodies hardcode whichever merge event is CURRENTLY
        active (e.g. PanelConfig.Id.UIMerge2608MergePanel). Registering a new
        merge event means repointing this single slot, not adding a region —
        the same single-slot-rotation risk already found and fixed in
        convert_layout.lua's item_to_machine table.

      - Every other family (Blocks/Sim/Dessert/Sort, ...): one region named
        "--region <Family><Digits>" whose function names AND panel-ID
        references both embed the event digits. Only the latest event's region
        exists per family at any time (e.g. only Blocks2607 is present, not
        also an older Blocks2601) — so registering a new event of that family
        renames the existing region in place, mirroring how its guide_refui
        Excel rows get cloned-and-the-old-commented rather than accumulated.

      - A family with no existing region at all (never seen in this file
        before) has no template to clone from — flagged, not guessed at.
    """
    import re as _re

    path = os.path.join(proj_root, r"develop\client\Skipbo\Assets\Lua\Game\Panel\UIGuidePanel\GuideRefuiMgr.lua")
    if not os.path.exists(path):
        log(f"GuideRefuiMgr.lua not found at {path} — skipping.")
        return True

    with open(path, 'r', encoding='utf-8') as f:
        content = f.read()

    tgt_pascal = tgt_evt[0].upper() + tgt_evt[1:]
    tgt_upper = tgt_evt.upper()
    tgt_lower = tgt_evt.lower()

    fam_match = _re.match(r'^([A-Za-z]+?)(\d+)$', tgt_evt)
    if not fam_match:
        log(f"GuideRefuiMgr.lua: couldn't parse family/digits from '{tgt_evt}' — skipping.")
        return True
    family = fam_match.group(1)

    backup_path = path + ".bak"
    def _ensure_backup():
        if not os.path.exists(backup_path):
            shutil.copy2(path, backup_path)
            record_file_creation(log_dir, step_index, backup_path)

    if family.lower() == "merge":
        region_match = _re.search(r'(--region\s+mergeFix\b.*?\n)(.*?)(\n--endregion)', content, _re.DOTALL | _re.IGNORECASE)
        if not region_match:
            log("GuideRefuiMgr.lua: no 'mergeFix' region found — skipping.")
            return True
        block = region_match.group(2)
        # Auto-detect whichever merge event is currently hardcoded rather than
        # trusting src_evt — this slot gets repointed by whoever last touched
        # it, which may not be src_evt if it's drifted out of sync.
        tokens = set(_re.findall(r'PanelConfig\.Id\.UI(Merge\d+)\w*', block))
        if not tokens:
            log("GuideRefuiMgr.lua: mergeFix region has no PanelConfig.Id.UIMerge<digits> references to repoint — skipping.")
            return True
        if tokens == {tgt_pascal}:
            log(f"GuideRefuiMgr.lua: mergeFix already points at {tgt_pascal}. Skipping.")
            return True
        if len(tokens) > 1:
            log(f"GuideRefuiMgr.lua: mergeFix region references multiple merge events ({sorted(tokens)}) — expected exactly one active slot. Repointing all of them to {tgt_pascal}.")
        _ensure_backup()
        new_block = block
        for tok in tokens:
            new_block = new_block.replace(tok, tgt_pascal)
        new_content = content[:region_match.start(2)] + new_block + content[region_match.end(2):]
        with open(path, 'w', encoding='utf-8') as f:
            f.write(new_content)
        log(f"GuideRefuiMgr.lua: repointed mergeFix region from {sorted(tokens)} to {tgt_pascal}.")
        return True

    # Non-merge families: rename the existing "<Family><Digits>" region in place.
    region_header_re = _re.compile(rf'--region\s+({family}\d+)\b', _re.IGNORECASE)
    header_match = region_header_re.search(content)
    if not header_match:
        log(f"GuideRefuiMgr.lua: no existing '--region {family}<digits>' block found for family '{family}' — this looks like a brand-new family with no template to clone from. Guide UI-hook functions must be added manually.")
        return True

    old_token = header_match.group(1)  # e.g. "Blocks2607", cased as written in the file
    if old_token.lower() == tgt_evt.lower():
        log(f"GuideRefuiMgr.lua: region already registered for {tgt_evt}. Skipping.")
        return True

    region_re = _re.compile(rf'(--region\s+{_re.escape(old_token)}\b.*?\n)(.*?)(\n--endregion)', _re.DOTALL)
    region_match = region_re.search(content)
    if not region_match:
        log(f"GuideRefuiMgr.lua: found region header for {old_token} but couldn't find its matching --endregion — skipping to avoid a bad edit.")
        return True

    old_upper = old_token.upper()
    old_pascal = old_token[0].upper() + old_token[1:]
    old_lower = old_token.lower()

    full_block = region_match.group(0)
    new_block = full_block.replace(old_upper, tgt_upper).replace(old_pascal, tgt_pascal).replace(old_lower, tgt_lower)

    _ensure_backup()
    new_content = content[:region_match.start()] + new_block + content[region_match.end():]
    with open(path, 'w', encoding='utf-8') as f:
        f.write(new_content)
    log(f"GuideRefuiMgr.lua: renamed region {old_token} -> {tgt_pascal}.")
    return True


def _regenerate_meta_guid(meta_path):
    """Unity .meta files carry a unique GUID identifying that specific asset.
    Copying a .meta file verbatim to a new asset would give two different
    files the same GUID, which corrupts Unity's asset reference resolution —
    so every cloned .meta gets a freshly generated one instead."""
    import re as _re
    import uuid
    with open(meta_path, 'r', encoding='utf-8') as f:
        content = f.read()
    new_content, n = _re.subn(r'guid:\s*[0-9a-fA-F]{32}', f'guid: {uuid.uuid4().hex}', content, count=1)
    if n == 0:
        log(f"Warning: no 'guid:' line found in {meta_path} — GUID not regenerated.")
        return
    with open(meta_path, 'w', encoding='utf-8') as f:
        f.write(new_content)


def _rename_tree_deepest_first(root, swap):
    """Rename every file/dir under root whose name changes under swap(). Returns the
    number renamed. Paths are snapshotted in one upfront pass, then renamed
    deepest-first: os.walk's own dirnames list goes stale if directories are renamed
    mid-walk (it hands back a name that's already been changed out from under it),
    while a fixed snapshot processed deepest-to-shallowest never touches an ancestor
    before its children are done."""
    all_paths = []
    for dirpath, dirnames, filenames in os.walk(root):
        for name in filenames + dirnames:
            all_paths.append(os.path.join(dirpath, name))
    all_paths.sort(key=lambda p: p.count(os.sep), reverse=True)

    renamed = 0
    for old_path in all_paths:
        base = os.path.basename(old_path)
        new_base = swap(base)
        if new_base != base:
            os.rename(old_path, os.path.join(os.path.dirname(old_path), new_base))
            renamed += 1
    return renamed


def _clone_one_event_code_folder(base_dir, src_folder_name, tgt_folder_name, src_evt, tgt_evt, log_dir, step_index):
    """Clones a single Core<Event> or UI<Event>Panel folder tree (Step 10a):
    copy the folder + its sibling .meta, rename every nested file/dir whose
    name embeds the source event token, rewrite Lua content via a multi-case
    swap, and regenerate a fresh GUID for every cloned .meta file."""
    src_dir = os.path.join(base_dir, src_folder_name)
    tgt_dir = os.path.join(base_dir, tgt_folder_name)
    src_meta = src_dir + ".meta"
    tgt_meta = tgt_dir + ".meta"

    if not os.path.isdir(src_dir):
        log(f"{src_folder_name}: source folder not found at {src_dir} — skipping.")
        return True
    if os.path.exists(tgt_dir):
        log(f"{tgt_folder_name}: target folder already exists — skipping (idempotent).")
        return True

    shutil.copytree(src_dir, tgt_dir)
    record_file_creation(log_dir, step_index, tgt_dir)

    if os.path.exists(src_meta):
        shutil.copy2(src_meta, tgt_meta)
        _regenerate_meta_guid(tgt_meta)
        record_file_creation(log_dir, step_index, tgt_meta)
    else:
        log(f"Warning: {src_folder_name}.meta not found — cloned folder {tgt_folder_name} has no folder-level .meta.")

    src_pascal = src_evt[0].upper() + src_evt[1:]
    tgt_pascal = tgt_evt[0].upper() + tgt_evt[1:]
    src_upper, tgt_upper = src_evt.upper(), tgt_evt.upper()
    src_lower, tgt_lower = src_evt.lower(), tgt_evt.lower()

    def swap(s):
        return s.replace(src_upper, tgt_upper).replace(src_pascal, tgt_pascal).replace(src_lower, tgt_lower)

    _rename_tree_deepest_first(tgt_dir, swap)

    # Now rewrite content: Lua files get the same multi-case swap, .meta files get a
    # fresh GUID instead (their content otherwise carries no event-specific text).
    file_count = 0
    for dirpath, _dirnames, filenames in os.walk(tgt_dir):
        for name in filenames:
            fpath = os.path.join(dirpath, name)
            if name.endswith(".meta"):
                _regenerate_meta_guid(fpath)
            else:
                with open(fpath, 'r', encoding='utf-8') as f:
                    content = f.read()
                new_content = swap(content)
                if new_content != content:
                    with open(fpath, 'w', encoding='utf-8') as f:
                        f.write(new_content)
                file_count += 1

    log(f"Cloned {src_folder_name} -> {tgt_folder_name} ({file_count} code files, GUIDs regenerated).")
    return True


def clone_event_code_folders(proj_root, src_evt, tgt_evt, log_dir, step_index):
    """
    Step 10a (docx): copy the event's Lua code folders for a reopen/new event.

    Confirmed real scope (merge2608): Core<Event> = 5 files, UI<Event>Panel = 63
    files across nested subfolders. Unlike GuideRefuiMgr.lua's per-family regions,
    these folders accumulate — every past event's folder is still present on disk
    (CoreMerge2606, CoreMerge2608, ... all coexist) — so this is a pure clone, never
    an overwrite-in-place.

    The event's ResourcesAb/Events asset folder is cloned separately by
    clone_event_resources (step 24). The "Generate UI Code" Unity codegen for
    genuinely new events is NOT covered — it invokes a Unity Editor prefab tool.
    """
    base_dir = os.path.join(proj_root, r"develop\client\Skipbo\Assets\Lua\Game\Panel\UIEventHub")
    src_pascal = src_evt[0].upper() + src_evt[1:]
    tgt_pascal = tgt_evt[0].upper() + tgt_evt[1:]

    pairs = [
        (f"Core{src_pascal}", f"Core{tgt_pascal}"),
        (f"UI{src_pascal}Panel", f"UI{tgt_pascal}Panel"),
    ]
    results = [_clone_one_event_code_folder(base_dir, s, t, src_evt, tgt_evt, log_dir, step_index) for s, t in pairs]
    return all(results)


_META_GUID_LINE = re.compile(rb'^guid:\s*([0-9a-f]{32})', re.M)
# Bare 32-hex match (not just "guid: x") so GUIDs stored under other keys are caught
# too; a false hit would need a random 128-bit value to equal one of the mapped GUIDs.
_HEX32 = re.compile(rb'(?<![0-9a-f])[0-9a-f]{32}(?![0-9a-f])')


def _meta_guid(meta_path):
    with open(meta_path, 'rb') as f:
        m = _META_GUID_LINE.search(f.read())
    return m.group(1) if m else None


def _fresh_guid_map(files):
    """old GUID -> brand-new GUID for every .meta among files. Freshly copied .meta files
    still carry the source asset's GUID, so this is exactly source GUID -> clone GUID."""
    import uuid
    guid_map = {}
    for p in files:
        if p.endswith(".meta"):
            g = _meta_guid(p)
            if g:
                guid_map[g] = uuid.uuid4().hex.encode()
    return guid_map


def _rewrite_unity_text_files(files, guid_map, swap_bytes):
    """In every .meta and Unity YAML (%YAML header) file: replace mapped GUIDs, then apply
    swap_bytes for names. Bytes-level, so LF endings and UTF-8 survive exactly; every other
    file (png/txt/bytes/...) is left untouched. Returns (guid refs remapped, files
    rewritten, list of the non-YAML files that were left alone)."""
    remapped = 0
    def sub(m):
        nonlocal remapped
        new = guid_map.get(m.group(0))
        if new is None:
            return m.group(0)
        remapped += 1
        return new

    rewritten, other_files = 0, []
    for p in files:
        with open(p, 'rb') as f:
            data = f.read()
        if p.endswith(".meta") or data.startswith(b'%YAML'):
            new = swap_bytes(_HEX32.sub(sub, data))
            if new != data:
                with open(p, 'wb') as f:
                    f.write(new)
                rewritten += 1
        else:
            other_files.append(p)
    return remapped, rewritten, other_files


def _find_event_backup_dir(proj_root):
    """docx Step 10: a closed event's resources may have been moved out of Assets into
    <repo>/common/program/event_backup (same Cell/Icon/Prefabs/Texture layout). That
    folder sits at the repo root, a sibling of trunk/ and branches/, so walk up from
    the branch root to find it rather than assuming how deep the branch is."""
    d = os.path.abspath(proj_root)
    for _ in range(4):
        cand = os.path.join(d, "common", "program", "event_backup")
        if os.path.isdir(cand):
            return cand
        d = os.path.dirname(d)
    return None


def clone_event_resources(proj_root, src_evt, tgt_evt, log_dir, step_index):
    """
    Step 24 (docx Step 10b): duplicate Assets/ResourcesAb/Events/<Src> as <Tgt> and
    repoint every reference to the new event.

    Unity-specific traps, measured on real event folders (Sim2607/Merge2608/Sort2609):
      - Every asset's .meta carries a GUID, and assets inside the folder reference each
        other by GUID — 9,328 such internal references in Sim2607 alone. A verbatim copy
        gives two assets the same GUID (Unity silently reassigns one on import, and the
        clone's prefabs keep pointing at the SOURCE's sprites); regenerating GUIDs
        without remapping breaks every one of them. So each .meta gets a new GUID and
        every occurrence of an old in-folder GUID is rewritten to its new one.
        References to GUIDs outside the folder (shared fonts/materials/scripts, ~8k per
        event) are left alone — they must keep pointing at the shared asset.
      - Event names only ever appear in dir/.prefab/.anim/.controller names and inside
        Unity YAML (m_Name, LocalizeKey) — never in png/txt/bytes/mat, across all five
        families. So only YAML files (%YAML header) and .meta get content edits; every
        other file is copied byte-for-byte (Spine .txt/.bytes pairs in particular must
        never be swapped independently of each other).
      - All edits are bytes-level, so Unity's LF line endings and UTF-8 text survive exactly.
    """
    import stat

    src_pascal = src_evt[0].upper() + src_evt[1:]
    tgt_pascal = tgt_evt[0].upper() + tgt_evt[1:]
    pairs = [(src_evt.upper(), tgt_evt.upper()), (src_pascal, tgt_pascal), (src_evt.lower(), tgt_evt.lower())]
    bpairs = [(a.encode(), b.encode()) for a, b in pairs]

    def swap(s):
        for a, b in pairs:
            s = s.replace(a, b)
        return s

    def swap_bytes(b):
        for x, y in bpairs:
            b = b.replace(x, y)
        return b

    events_dir = os.path.join(proj_root, "develop", "client", "Skipbo", "Assets", "ResourcesAb", "Events")
    tgt_dir = os.path.join(events_dir, tgt_pascal)
    tgt_meta = tgt_dir + ".meta"
    if os.path.exists(tgt_dir):
        # Existing is not the same as done. An empty <Tgt> folder - left by a reverted
        # run, or recreated by SVN/Unity - used to count as "already cloned", so nothing
        # was copied and the event shipped with no resources at all. That is the exact
        # mistake chapter update's dynamic_map copy made three runs in a row. Only real
        # assets make the folder somebody's work; orphan .meta files don't.
        assets = [n for _, _, fn in os.walk(tgt_dir) for n in fn if not n.endswith(".meta")]
        if assets:
            log(f"ResourcesAb/Events/{tgt_pascal} already holds {len(assets)} asset file(s) — "
                f"skipping (never overwritten).")
            return True
        # Cleared rather than copied into. What a reverted clone leaves behind is orphan
        # .meta files that ALREADY carry the target's names (bg_sim2609.png.meta), so the
        # rename pass below collides with them and Windows refuses the rename. With no
        # assets in the folder they are metadata for nothing - Unity deletes them on
        # import, and every GUID is regenerated here anyway - so nothing is lost.
        # ponytail: Undo then removes the folder even though it existed (empty) before.
        # Recording files individually would go stale the moment the rename pass runs,
        # and an empty folder is not worth that.
        for dp, _, fn in os.walk(tgt_dir):   # svn:needs-lock files are read-only, and
            for n in fn:                     # rmtree cannot delete those on Windows
                os.chmod(os.path.join(dp, n), stat.S_IWRITE | stat.S_IREAD)
        shutil.rmtree(tgt_dir)
        log(f"ResourcesAb/Events/{tgt_pascal} existed but held no assets (only orphan .meta "
            f"files, if any) — cleared it and cloning from {src_pascal}.")

    src_dir = os.path.join(events_dir, src_pascal)
    if not os.path.isdir(src_dir):
        backup = _find_event_backup_dir(proj_root)
        if backup and os.path.isdir(os.path.join(backup, src_pascal)):
            src_dir = os.path.join(backup, src_pascal)
            log(f"ResourcesAb/Events/{src_pascal} not in Assets — using the archived copy at {src_dir}.")
        else:
            log(f"Warning: no ResourcesAb/Events/{src_pascal} folder (checked Assets and common/program/event_backup). Skipping.")
            return True

    shutil.copytree(src_dir, tgt_dir)
    record_file_creation(log_dir, step_index, tgt_dir)
    if os.path.exists(src_dir + ".meta"):
        shutil.copy2(src_dir + ".meta", tgt_meta)
        record_file_creation(log_dir, step_index, tgt_meta)
    else:
        log(f"Warning: {src_pascal}.meta not found — Unity will generate a folder .meta for {tgt_pascal} on import.")

    # Working-copy files can be read-only (svn:needs-lock); copy2 carries that over, which
    # would block the rewrite below and the undo's rmtree.
    all_files = [os.path.join(dp, n) for dp, _, fn in os.walk(tgt_dir) for n in fn]
    for p in all_files + ([tgt_meta] if os.path.exists(tgt_meta) else []):
        os.chmod(p, stat.S_IWRITE | stat.S_IREAD)

    renamed = _rename_tree_deepest_first(tgt_dir, swap)

    all_files = [os.path.join(dp, n) for dp, _, fn in os.walk(tgt_dir) for n in fn]
    if os.path.exists(tgt_meta):
        all_files.append(tgt_meta)
    guid_map = _fresh_guid_map(all_files)
    remapped, rewritten, other_files = _rewrite_unity_text_files(all_files, guid_map, swap_bytes)

    src_tokens = [x for x, _ in bpairs]
    untouched_with_token = []
    for p in other_files:
        with open(p, 'rb') as f:
            data = f.read()
        if any(t in data for t in src_tokens):
            untouched_with_token.append(os.path.relpath(p, tgt_dir))

    log(f"ResourcesAb/Events: cloned {src_pascal} -> {tgt_pascal} ({len(all_files)} files, {renamed} renamed, "
        f"{len(guid_map)} GUIDs regenerated, {remapped} GUID references remapped, {rewritten} files rewritten).")
    if untouched_with_token:
        log(f"Warning: {len(untouched_with_token)} non-YAML file(s) contain '{src_pascal}' but were copied unchanged "
            f"(binary-safe) — check whether they need updating: {untouched_with_token[:10]}")
    return True


def _lookup_event_id_and_start(proj_root, evt_name, exclude_id=None, workspace=None):
    """(event_id, start_time) of evt_name's LATEST row in the workspace's events.xlsx
    (rows are chronological), skipping exclude_id — so a reopen that reuses the source's
    name still resolves to the source's own row, not the target row step 10 just added."""
    import openpyxl
    path = os.path.join(proj_root, "design", "DesignData", workspace or WORKSPACE_NAME, "design", "events.xlsx")
    if not os.path.exists(path):
        return None, None
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    found = (None, None)
    alias_re = re.compile(rf'^{re.escape(evt_name)}(?!\d)', re.IGNORECASE)
    for row in wb["events"].iter_rows(min_row=2, max_col=8, values_only=True):
        eid = str(row[0] or "").replace("//", "").strip()
        if eid and alias_re.match(str(row[1] or "").strip()) and eid != exclude_id:
            found = (eid, str(row[7] or ""))
    wb.close()
    return found


def retired_export_ids(ws_dir, evt, event_id=None):
    """{table.json: [ids]} of evt's own rows whose IDs PostConvert's default_filter_map.lua drops
    from the client export. The team raises those ranges as events retire (DeepLink: up to 12300 in
    May 2026, 12526 in Aug, 12567 in Sep), so an old event reopened under its own name keeps IDs the
    game never loads — Merge2510's event row, entrance, score, 31 gift/pack rows and 48 share links
    all sit inside them. A row is evt's own when its name is in the ID/alias columns (another event's
    标签刷新前置活动 column can mention it further right). event_id stands in for the rows step 10
    re-keys to it (events / event_pre / event_post)."""
    import glob as _glob
    import openpyxl
    lua = os.path.join(ws_dir, "PostConvert", "Lua", "convert", "default_filter_map.lua")
    if not os.path.exists(lua):
        return {}
    ranges = {}
    for m in re.finditer(r'\["([\w.]+\.json)"\]\s*=\s*\{(.*?)\n\s*\},', open(lua, encoding="utf-8").read(), re.S):
        body = re.sub(r'--[^\n]*', '', m.group(2))
        ranges[m.group(1)] = ([(int(a), int(b)) for a, b in re.findall(r'\{\s*(\d+)\s*,\s*(\d+)\s*\}', body)] +
                              [(int(n), int(n)) for n in re.findall(r'\d+', re.sub(r'\{[^{}]*\}', '', body))])

    def retired(j, k):
        return any(a <= k <= b for a, b in ranges[j])

    evt_lower, out = evt.lower(), {}
    for p in _glob.glob(os.path.join(ws_dir, "descriptor", "app_client", "*.json")):
        try:
            directors = json.load(open(p, encoding="utf-8")).get("Directors") or []
        except Exception:
            continue
        for d in directors:
            j = d.get("FileName")
            if j not in ranges:
                continue
            if event_id and j in ("events.json", "event_pre.json", "event_post.json"):
                if retired(j, int(event_id)):
                    out[j] = [int(event_id)]
                continue
            for es in d.get("ExcelSheets", []):
                path = os.path.join(ws_dir, "design", es["ExcelName"] + ".xlsx")
                if not os.path.exists(path):
                    continue
                wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
                if es["SheetName"] in wb.sheetnames:
                    rows = wb[es["SheetName"]].iter_rows(values_only=True)
                    head = [str(h or "").strip() for h in next(rows, ())]
                    ki = head.index(d.get("Key")) if d.get("Key") in head else 0
                    for r in rows:
                        if not any(evt_lower in str(v).lower() for v in r[:2] if v is not None):
                            continue
                        try:
                            k = int(float(r[ki]))
                        except (TypeError, ValueError, IndexError):
                            continue  # commented-out (//) or blank key
                        if retired(j, k):
                            out.setdefault(j, []).append(k)
                wb.close()
    return out


def _guids_in_use(assets_dir, skip_paths):
    """Every asset GUID under assets_dir, ignoring the files in skip_paths."""
    skip = {os.path.normcase(p) for p in skip_paths}
    used = set()
    for dp, _, fn in os.walk(assets_dir):
        for n in fn:
            if n.endswith(".meta"):
                p = os.path.join(dp, n)
                if os.path.normcase(p) not in skip:
                    g = _meta_guid(p)
                    if g:
                        used.add(g)
    return used


def _svn(args, cwd, timeout=120):
    import subprocess
    try:
        r = subprocess.run(["svn"] + args, cwd=cwd, shell=True, capture_output=True, text=True,
                           timeout=timeout, encoding='utf-8', errors='replace')
    except Exception as e:
        log(f"svn {args[0]} failed to run: {e}")
        return None
    return r.stdout if r.returncode == 0 else None


def _svn_export_deleted(wc_dir, name_prefix, hint_date, dest_dir):
    """
    Recover files/dirs under wc_dir whose name starts with name_prefix and that have since
    been deleted, by finding each one's deletion commit and exporting it from the revision
    just before. Returns {name: exported_path}.

    Deleted paths are matched by their location relative to the ResourcesAb tree, not by
    the current branch URL: /branches/Misc was re-created from trunk on 2026-09-24
    (r133037), so every older deletion is recorded under /trunk/... in this path's history.
    """
    from datetime import datetime, timedelta
    root_url = (_svn(["info", "--show-item", "repos-root-url", "."], wc_dir) or "").strip()
    if not root_url:
        return {}

    marker = "/Assets/ResourcesAb/"
    rel_dir = wc_dir.replace("\\", "/").split(marker, 1)[-1].rstrip("/")
    want_parent = marker + rel_dir + "/"

    def scan(rev_args):
        out = _svn(["log", "-v", "-q"] + rev_args + ["."], wc_dir, timeout=300) or ""
        deleted, rev = {}, None
        for line in out.splitlines():
            m = re.match(r'^r(\d+) \|', line)
            if m:
                rev = int(m.group(1))
                continue
            m = re.match(r'^\s+D (/\S.*)$', line)
            if m and rev and want_parent in m.group(1):
                path = m.group(1)
                name = path.rsplit("/", 1)[-1]
                if path.endswith(want_parent + name) and name.startswith(name_prefix):
                    if name not in deleted or rev > deleted[name][0]:
                        deleted[name] = (rev, path)
        return deleted

    deleted = {}
    try:
        center = datetime.strptime(hint_date.strip()[:10], "%Y-%m-%d")
        deleted = scan(["-r", f"{{{(center - timedelta(days=45)):%Y-%m-%d}}}:{{{(center + timedelta(days=240)):%Y-%m-%d}}}"])
    except (ValueError, AttributeError):
        pass
    if not deleted:
        deleted = scan(["-l", "500"])  # ponytail: no/odd date hint -> last 500 revs of this dir

    exported = {}
    os.makedirs(dest_dir, exist_ok=True)  # a folder export creates its own path; a single-file export doesn't
    for name, (rev, path) in deleted.items():
        dest = os.path.join(dest_dir, name)
        if _svn(["export", "-q", "--force", f"{root_url}{path}@{rev - 1}", dest], wc_dir) is not None and os.path.exists(dest):
            exported[name] = dest
            log(f"  recovered {name} from SVN ({path}@{rev - 1}, deleted in r{rev})")
        else:
            log(f"  Warning: {name} was deleted in r{rev} but couldn't be exported from r{rev - 1}.")
    return exported


def _chronicle_numbers(proj_root, src_evt, tgt_evt):
    """(source, target) chronicle numbers as written, e.g. ("90", "92"), from asset_ref.xlsx
    key_asset rows ACHIEVEMENT_E_<EVENT> -> .../EventAchieve/icon_ChronicleNN.png."""
    import openpyxl
    path = os.path.join(proj_root, "design", "DesignData", WORKSPACE_NAME, "design", "asset_ref.xlsx")
    if not os.path.exists(path):
        return None, None
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    found = {}
    if "key_asset" in wb.sheetnames:
        want = {f"ACHIEVEMENT_E_{src_evt.upper()}": "src", f"ACHIEVEMENT_E_{tgt_evt.upper()}": "tgt"}
        for row in wb["key_asset"].iter_rows(min_col=1, max_col=3, values_only=True):
            key = str(row[0] or "").replace("//", "").strip().upper()
            m = re.search(r'icon_Chronicle(\d+)', str(row[2] or ""), re.IGNORECASE)
            if m:
                for k, role in want.items():
                    if key == k:
                        found[role] = m.group(1)
    wb.close()
    return found.get("src"), found.get("tgt")


def _clone_chronicle_icon(proj_root, src_evt, tgt_evt, log_dir, step_index):
    """
    Texture/EventAchieve/icon_ChronicleNN.png — the event's chronicle (album) icon, loaded by
    path from asset_ref's ACHIEVEMENT_E_<EVENT> row. Numbers are sequential with one NEW
    number per event, reopens included (Sim2605=82, Sim2607=86, Sim2609=90). Step 15 already
    gives the target its number (max+1) in asset_ref.xlsx but nothing creates the PNG, so the
    new row points at a missing file. Both numbers are read back from those rows rather than
    recounted, so this can never disagree with step 15. Returns the files it created.
    """
    import stat
    import tempfile

    if src_evt.lower() == tgt_evt.lower():
        log(f"EventAchieve: reopening {tgt_evt} under its own name keeps its existing chronicle icon — nothing to copy.")
        return []
    src_n, tgt_n = _chronicle_numbers(proj_root, src_evt, tgt_evt)
    if not tgt_n:
        log(f"EventAchieve: no ACHIEVEMENT_E_{tgt_evt.upper()} row in {WORKSPACE_NAME}/asset_ref.xlsx yet — run step 15 first (it assigns the number). Skipped.")
        return []
    if not src_n:
        log(f"EventAchieve: no ACHIEVEMENT_E_{src_evt.upper()} row for the source in asset_ref.xlsx — skipped.")
        return []
    if src_n == tgt_n:
        log(f"EventAchieve: {tgt_evt} and {src_evt} both point at icon_Chronicle{tgt_n} in asset_ref.xlsx — nothing to copy.")
        return []

    folder = os.path.join(proj_root, "develop", "client", "Skipbo", "Assets", "ResourcesAb", "Texture", "EventAchieve")
    src_name, tgt_name = f"icon_Chronicle{src_n}.png", f"icon_Chronicle{tgt_n}.png"
    dst = os.path.join(folder, tgt_name)
    if os.path.exists(dst):
        log(f"EventAchieve: {tgt_name} already exists — left untouched.")
        return []

    tmp = tempfile.mkdtemp(prefix="chron_svn_")
    try:
        sources = {n: os.path.join(folder, n) for n in (src_name, src_name + ".meta") if os.path.isfile(os.path.join(folder, n))}
        if src_name not in sources:  # never happened so far (86 of 86 present), but cheap to cover
            log(f"EventAchieve: {src_name} not in the live folder — checking SVN history...")
            sources = {**sources, **_svn_export_deleted(folder, src_name, "", tmp)}
        if src_name not in sources:
            log(f"Warning: EventAchieve: {src_name} not found (live or SVN) — {tgt_name} not created.")
            return []
        created = []
        for n, p in sources.items():
            d = os.path.join(folder, tgt_name + n[len(src_name):])
            shutil.copy2(p, d)
            os.chmod(d, stat.S_IWRITE | stat.S_IREAD)
            record_file_creation(log_dir, step_index, d)
            created.append(d)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
    _rewrite_unity_text_files(created, _fresh_guid_map(created), lambda b: b)
    log(f"EventAchieve: {src_name} -> {tgt_name} (matches asset_ref.xlsx ACHIEVEMENT_E_{tgt_evt.upper()}; new GUID).")
    return created


def clone_event_id_resources(proj_root, src_evt, tgt_evt, log_dir, step_index):
    """
    Step 25 (docx Step 10b): the event resources keyed by EVENT ID (yyyymmnn from
    events.xlsx) rather than event name, all under ResourcesAb/Events/Common:
      - Cell/Entrance/<id>TabEventBtn.prefab — the HUD tab button. UIEventTabWidget.lua loads
        "<eventId>TabEventBtn" with no fallback, so without it the event has no HUD entry.
      - UI/<id>TabEventBtn/ — that prefab's sprites, linked by GUID (14 of its 45 references).
      - Icon/CommonEventImg/<id>_* — buff/token images. Step 13 already rewrites icon.xlsx
        sprite names to "<newId>_img_*", so these files must exist or those icons are blank.
      - Icon/PropIcon/<id>_* — the event token / rescue-token / tool item icons (icon_item rows
        with no atlas), same naming and lifecycle as CommonEventImg.
    Plus Texture/EventAchieve/icon_ChronicleNN.png, keyed by a sequential number rather than
    the event ID (see _clone_chronicle_icon) — run first, since it doesn't need event IDs.

    Each piece comes from the live folder, else from SVN (these get deleted once an event
    ends). Nothing is cached: an SVN lookup here costs seconds, not the minutes the xlsx
    history scans take — and the log dir sits inside Assets/, where Unity would import
    cached PNGs/prefabs as real assets.
    """
    import stat
    import tempfile

    _clone_chronicle_icon(proj_root, src_evt, tgt_evt, log_dir, step_index)

    extra = load_extra_args(log_dir)
    tgt_id = str(extra.get("event_id", "")).strip()
    if not re.fullmatch(r'\d{8}', tgt_id):
        log(f"Warning: no valid target event ID supplied (got '{tgt_id}') — skipping event-ID resources.")
        return True
    src_id, src_start = _lookup_event_id_and_start(proj_root, src_evt, exclude_id=tgt_id)
    if not src_id and WORKSPACE_NAME != "WorkSpcae3":
        # Step 10 rewrites the event's own row with the target ID, so for a same-name reopen
        # with a new ID the original ID is only still visible in WorkSpcae3 (never edited here).
        src_id, src_start = _lookup_event_id_and_start(proj_root, src_evt, exclude_id=tgt_id, workspace="WorkSpcae3")
    restore = False
    if not src_id and src_evt.lower() == tgt_evt.lower():
        # Same name AND same ID: nothing to rename. The event's own ID-keyed files were deleted
        # when it closed (Merge2510's all were), so bring them back in place from SVN instead.
        src_id, restore = tgt_id, True
    if not src_id:
        log(f"Warning: {src_evt} not found in {WORKSPACE_NAME}/events.xlsx — can't tell which event-ID resources to copy. Skipping.")
        return True
    src_start = extra.get("src_start_time") or src_start  # the source's ORIGINAL start; step 10 may have rewritten the row
    log(f"Event-ID resources: {src_evt} ({src_id}) -> {tgt_evt} ({tgt_id})" + (" — restoring in place" if restore else ""))

    common = os.path.join(proj_root, "develop", "client", "Skipbo", "Assets", "ResourcesAb", "Events", "Common")
    pieces = [  # (label, folder, name prefix shared by every item of the piece)
        ("tab button prefab", os.path.join(common, "Cell", "Entrance"), f"{src_id}TabEventBtn"),
        ("tab button sprites", os.path.join(common, "UI"), f"{src_id}TabEventBtn"),
        ("CommonEventImg images", os.path.join(common, "Icon", "CommonEventImg"), f"{src_id}_"),
        # icon.xlsx icon_item rows with no atlas (<EVENT>_TOKEN -> "<id>_img_token", _RESCUE_TOKEN
        # -> "<id>_img_prop", plus tools like _img_hammer) resolve here, and are deleted with the event.
        ("PropIcon item icons", os.path.join(proj_root, "develop", "client", "Skipbo", "Assets", "ResourcesAb", "Icon", "PropIcon"), f"{src_id}_"),
    ]

    # Only whole-name uses of the ID: Unity files are full of long fileIDs and hex GUIDs
    # that can contain the same 8 digits.
    id_re = re.compile(rb'(?<![0-9A-Za-z])' + src_id.encode() + rb'(?=TabEventBtn|_)')
    def swap_bytes(b):
        return id_re.sub(tgt_id.encode(), b)
    def tgt_name(name):
        return tgt_id + name[len(src_id):]

    def has_content(p):  # a leftover .meta or an emptied folder (e.g. Merge2608's today) isn't the asset
        if p.endswith(".meta"):
            return False
        return os.path.isfile(p) or (os.path.isdir(p) and any(not n.endswith(".meta") for n in os.listdir(p)))

    tmp = tempfile.mkdtemp(prefix="evtid_svn_")
    created, guid_map, summary = [], {}, []
    try:
        for label, folder, prefix in pieces:
            if not os.path.isdir(folder):
                log(f"Warning: {folder} not found — skipping {label}.")
                continue
            live = {n: os.path.join(folder, n) for n in os.listdir(folder) if n.startswith(prefix)}
            if not any(has_content(p) for p in live.values()):
                log(f"{label}: not in the live folder — checking SVN history...")
                recovered = _svn_export_deleted(folder, prefix, src_start, os.path.join(tmp, os.path.basename(folder)))
                # Keep a leftover live .meta the deletion missed (Misc still has an orphaned
                # UI/20260901TabEventBtn.meta) rather than letting Unity invent a new one.
                live = {**{n: p for n, p in live.items() if n.endswith(".meta")}, **recovered} if recovered else {}
            if not live:
                log(f"Warning: no {label} found for {src_id} (live or SVN) — skipped.")
                summary.append(f"{label}: none")
                continue

            # Only real content counts as "already exists": a leftover .meta or an emptied folder
            # (what an event's cleanup commit leaves behind) must not block the copy/restore.
            existing = [n for n in live if has_content(os.path.join(folder, tgt_name(n)))]
            if existing:
                # Something (usually art) already made the target's version — never overwrite.
                # Still link it: map each source GUID to the existing target file's GUID by name,
                # so a freshly cloned prefab points at these sprites rather than the source's.
                for n in existing:
                    src_p, tgt_p = live[n], os.path.join(folder, tgt_name(n))
                    pairs = [(src_p, tgt_p)] if os.path.isfile(src_p) else [
                        (os.path.join(dp, f), os.path.join(tgt_p, os.path.relpath(os.path.join(dp, f), src_p)))
                        for dp, _, fn in os.walk(src_p) for f in fn]
                    for s, t in pairs:
                        if s.endswith(".meta") and os.path.exists(t):
                            gs, gt = _meta_guid(s), _meta_guid(t)
                            if gs and gt:
                                guid_map[gs] = gt
                log(f"{label}: {tgt_id} version already exists — left untouched, linked by name.")
                summary.append(f"{label}: already existed")
                continue

            for n, src_p in live.items():
                dst = os.path.join(folder, tgt_name(n))
                if dst.endswith(".meta") and os.path.exists(dst):
                    # Keep a .meta that's already there - and record whose GUID it stands in
                    # for. The asset copied beside it takes THIS file's GUID, so every link to
                    # the source's GUID (in prefabs copied here, and in the event's own prefabs
                    # repointed below) has to move to it; otherwise they keep pointing at the
                    # source event's image. Same-ID restore: both GUIDs are the original, so
                    # this maps a GUID to itself and changes nothing.
                    gs, gt = _meta_guid(src_p), _meta_guid(dst)
                    if gs and gt and gs != gt:
                        guid_map[gs] = gt
                    continue
                if os.path.isdir(src_p):
                    shutil.copytree(src_p, dst, dirs_exist_ok=True)  # may land in an emptied leftover folder
                    new_files = [os.path.join(dp, f) for dp, _, fn in os.walk(dst) for f in fn]
                else:
                    shutil.copy2(src_p, dst)
                    new_files = [dst]
                record_file_creation(log_dir, step_index, dst)
                for p in new_files:
                    os.chmod(p, stat.S_IWRITE | stat.S_IREAD)  # svn:needs-lock files copy over read-only
                created += new_files
            summary.append(f"{label}: {len(live)} item(s) {'restored' if restore else 'copied'}")
    finally:
        shutil.rmtree(tmp, ignore_errors=True)

    fresh = _fresh_guid_map(created)
    if restore and created:   # nothing restored = nothing to check; skip the 30s scan
        # Restored files keep their ORIGINAL GUIDs, so every existing link to them in the event's
        # own prefabs resolves again untouched. The one exception: a GUID something else now
        # owns (the team renames these files forward, e.g. <id>_img_token -> the next event's,
        # which keeps the GUID) — a duplicate would confuse Unity, so that file gets a new one.
        # Scans every .meta under Assets: ~67k GUIDs, around 30s on a dev machine. Said up
        # front, or the step sits silent for half a minute and looks hung.
        log(f"Checking {len(created)} restored file(s) against every GUID in Assets "
            f"(reads each .meta once - this takes about half a minute)...")
        in_use = _guids_in_use(os.path.join(proj_root, "develop", "client", "Skipbo", "Assets"), set(created))
        fresh = {g: n for g, n in fresh.items() if g in in_use}
        if fresh:
            log(f"{len(fresh)} restored file(s) had a GUID another asset now uses — gave them new ones.")
    guid_map.update(fresh)
    remapped, rewritten, _ = _rewrite_unity_text_files(created, guid_map, swap_bytes)
    log(f"Event-ID resources {src_id} -> {tgt_id}: {'; '.join(summary)}. {len(created)} files created, "
        f"{remapped} GUID references remapped, {rewritten} files rewritten.")

    # Repoint the new event's own prefabs (the folder step 24 cloned) from the source's
    # ID-keyed images to the target's. The clone still links the source's GUIDs: fine while
    # the source's images are live, broken once they're cleaned up — Merge2611 inherited 56
    # such links to Merge2608's token/prop/buff images, deleted in r132788. The team renames
    # these files forward, which keeps the GUID; this tool copies, so the links must move too.
    evt_dir = os.path.join(proj_root, "develop", "client", "Skipbo", "Assets", "ResourcesAb", "Events", tgt_evt[0].upper() + tgt_evt[1:])
    moves = {old: new for old, new in guid_map.items() if old != new}
    if moves and os.path.isdir(evt_dir):
        touched = []
        for dp, _, fn in os.walk(evt_dir):
            for n in fn:
                if n.endswith(".meta"):
                    continue
                p = os.path.join(dp, n)
                with open(p, 'rb') as f:
                    data = f.read()
                if data.startswith(b'%YAML') and any(old in data for old in moves):
                    touched.append(p)
        for p in touched:
            os.chmod(p, stat.S_IWRITE | stat.S_IREAD)
            record_file_backup(log_dir, step_index, p)
        relinked, n_files, _ = _rewrite_unity_text_files(touched, moves, lambda b: b)
        log(f"Relinked {relinked} reference(s) in {n_files} {os.path.basename(evt_dir)} file(s) from {src_id}'s images to {tgt_id}'s.")
    return True


def register_event_lua_configs(proj_root, src_evt, tgt_evt, log_dir, step_index):
    """Step 22: registers the target event across EventConfig.lua, BIDef.lua,
    UIManager.lua, PanelConfig.lua, and GuideRefuiMgr.lua. Each file is
    independent — all five are attempted regardless of individual failures,
    overall success requires all five."""
    results = [
        update_event_config_lua(proj_root, src_evt, tgt_evt, log_dir, step_index),
        update_bidef_lua(proj_root, src_evt, tgt_evt, log_dir, step_index),
        update_ui_manager_special_panel_map(proj_root, src_evt, tgt_evt, log_dir, step_index),
        update_panel_config_lua(proj_root, src_evt, tgt_evt, log_dir, step_index),
        update_guide_refui_mgr(proj_root, src_evt, tgt_evt, log_dir, step_index),
    ]
    return all(results)


def main():
    # Log messages include Chinese sheet/column names; when stdout isn't attached to a
    # real UTF-8 console (e.g. piped from Electron's spawn on a non-UTF-8 Windows locale)
    # the default encoding can be cp1252 and crash on write. Force UTF-8 defensively.
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

    if len(sys.argv) < 6:
        log("Usage: python event_pipeline.py <step_index> <source_event> <target_event> <project_root> <is_revert>")
        sys.exit(1)

    step_index = int(sys.argv[1])
    src_evt = sys.argv[2]
    tgt_evt = sys.argv[3]
    proj_root = sys.argv[4]
    global WORKSPACE_NAME
    WORKSPACE_NAME = sys.argv[6] if len(sys.argv) > 6 else WORKSPACE_NAME
    
    # Optional 5th arg indicating revert mode
    is_revert = True if sys.argv[5] == "1" else False

    log_dir = os.path.join(proj_root, r"develop\client\Skipbo\Assets\Editor\EventAutomation\PythonBackend")
    os.makedirs(log_dir, exist_ok=True)

    if is_revert:
        log(f"Reverting Python Step {step_index} | {src_evt} -> {tgt_evt}")
        success = revert_step(log_dir, step_index)
        sys.exit(0 if success else 1)
        
    log(f"Executing Python Step {step_index} | {src_evt} -> {tgt_evt}")

    if step_index == 0:
        # Same-name reopen: stop before anything is touched if the event's own IDs are ones the
        # client export drops. Reads WorkSpcae3 — what step 5 copies in.
        if src_evt.lower() == tgt_evt.lower():
            ws3 = os.path.join(proj_root, "design", "DesignData", "WorkSpcae3")
            retired = retired_export_ids(ws3, tgt_evt, load_extra_args(log_dir).get("event_id"))
            if retired:
                log(f"Reopen blocked: reopening {tgt_evt} under its own name keeps its old rows, and these IDs are inside the "
                    f"ranges PostConvert/Lua/convert/default_filter_map.lua drops from the client export — the game would "
                    f"never load them (a missing entrance is a Lua error in EventMgr):")
                for j, ids in retired.items():
                    log(f"  {j}: {ids if len(ids) <= 3 else f'{len(ids)} IDs, {min(ids)}-{max(ids)}'}")
                log(f"Nothing was changed. Either run it as a new event (source {src_evt} -> a new name), which gives every row a "
                    f"fresh ID, or take these IDs out of default_filter_map.lua in WorkSpcae3 and run again.")
                sys.exit(1)
        proto_dir = os.path.join(proj_root, r"develop\protocol")
        success = inject_proto(proto_dir, src_evt, tgt_evt, log_dir, step_index)
        if not success: sys.exit(1)
    elif step_index == 2:
        # Step 3: Add active save file in local service
        src_file = os.path.join(proj_root, rf"develop\client\Skipbo\Assets\Lua\Game\Module\LocalService\Event\E{src_evt.capitalize()}LocalService.lua")
        tgt_file = os.path.join(proj_root, rf"develop\client\Skipbo\Assets\Lua\Game\Module\LocalService\Event\E{tgt_evt.capitalize()}LocalService.lua")
        success = copy_and_rename_file(src_file, tgt_file, src_evt, tgt_evt, log_dir, step_index)
        if not success: sys.exit(1)
    elif step_index == 3:
        # Step 4: Register in LocalServiceMgr.lua
        success = inject_local_service_mgr(proj_root, tgt_evt, log_dir, step_index)
        if not success: sys.exit(1)
    elif step_index == 4:
        # Step 5: Clone WorkSpace3 configuration to Design
        success = clone_workspace3_to_design(proj_root, step_index)
        if success:
            design_dir = os.path.join(proj_root, r"design\DesignData")
            workspace_design = os.path.join(design_dir, WORKSPACE_NAME)
            record_file_creation(log_dir, step_index, workspace_design)
        if not success: sys.exit(1)
    elif step_index == 5:
        # Step 6: Clone Event Excel configuration tables
        success = clone_event_excel(proj_root, src_evt, tgt_evt, log_dir, step_index)
        if not success: sys.exit(1)
    elif step_index == 6:
        # Step 7: Copy event descriptor JSON(s) to app_client
        success = clone_event_descriptors(proj_root, src_evt, tgt_evt, log_dir, step_index)
        if not success: sys.exit(1)
    elif step_index == 7:
        # Step 8: Modify convert_layout.lua and convert.json
        success = inject_convert_references(proj_root, src_evt, tgt_evt, log_dir, step_index)
        if not success: sys.exit(1)
    elif step_index == 8:
        # Step 9: Update bi.xlsx
        success = update_bi_excel(proj_root, src_evt, tgt_evt, log_dir, step_index)
        if not success: sys.exit(1)
    elif step_index == 9:
        # Step 10: Update events.xlsx
        success = update_events_excel(proj_root, src_evt, tgt_evt, log_dir, step_index)
        if not success: sys.exit(1)
    elif step_index == 10:
        # Step 11: Update event_shop.xlsx
        success = update_event_shop_excel(proj_root, src_evt, tgt_evt, log_dir, step_index)
        if not success: sys.exit(1)
    elif step_index == 11:
        # Step 12: Update item.xlsx
        success, id_map = update_item_excel(proj_root, src_evt, tgt_evt, log_dir, step_index)
        if success and id_map:
            sync_events_excel_item_ids(proj_root, tgt_evt, id_map, log_dir, step_index)
        if success:
            backfill_item_icon_columns(proj_root, src_evt, tgt_evt, log_dir)
        if not success: sys.exit(1)
    elif step_index == 12:
        # Step 13: Update icon.xlsx
        success = update_icon_excel(proj_root, src_evt, tgt_evt, log_dir, step_index)
        if not success: sys.exit(1)
    elif step_index == 13:
        # Step 14: Update localization files, quiz, and answer_challenge
        success = update_localization_all(proj_root, src_evt, tgt_evt, log_dir, step_index)
        if not success: sys.exit(1)
    elif step_index == 14:
        # Step 15: Update asset_ref.xlsx
        success = update_asset_ref_excel(proj_root, src_evt, tgt_evt, log_dir, step_index)
        if not success: sys.exit(1)
    elif step_index == 15:
        # Step 16: Update store.xlsx
        success = update_store_excel(proj_root, src_evt, tgt_evt, log_dir, step_index)
        if not success: sys.exit(1)
    elif step_index == 16:
        # Step 17: Update card.xlsx (Placeholder)
        log("card.xlsx automation skipped (no event info present).")
    elif step_index == 17:
        # Step 18: Update pack.xlsx
        success = update_pack_excel(proj_root, src_evt, tgt_evt, log_dir, step_index)
        if not success: sys.exit(1)
    elif step_index == 18:
        # Step 19: Update guide.xlsx
        success = update_guide_excel(proj_root, src_evt, tgt_evt, log_dir, step_index)
        if not success: sys.exit(1)
    elif step_index == 19:
        # Step 20: Update sys.xlsx
        success = update_sys_excel(proj_root, src_evt, tgt_evt, log_dir, step_index)
        if not success: sys.exit(1)
    elif step_index == 21:
        # Step 22: Register event in EventConfig.lua, BIDef.lua, UIManager.lua, PanelConfig.lua, GuideRefuiMgr.lua
        success = register_event_lua_configs(proj_root, src_evt, tgt_evt, log_dir, step_index)
        if not success: sys.exit(1)
    elif step_index == 22:
        # Step 23: Copy & rename event code folders (Core<Event> + UI<Event>Panel)
        success = clone_event_code_folders(proj_root, src_evt, tgt_evt, log_dir, step_index)
        if not success: sys.exit(1)
    elif step_index == 23:
        # Step 24: Copy & rename ResourcesAb/Events/<Event> (GUIDs regenerated + remapped)
        success = clone_event_resources(proj_root, src_evt, tgt_evt, log_dir, step_index)
        if not success: sys.exit(1)
    elif step_index == 24:
        # Step 25: Copy & rename event-ID resources (TabEventBtn prefab + sprites, CommonEventImg)
        success = clone_event_id_resources(proj_root, src_evt, tgt_evt, log_dir, step_index)
        if not success: sys.exit(1)
    else:
        log(f"Step {step_index} execution mock succeeded.")
        # Optional mock file creation to test undoing
        # mock_file = os.path.join(log_dir, f"mock_file_step_{step_index}.txt")
        # open(mock_file, 'w').close()
        # record_file_creation(log_dir, step_index, mock_file)
    
    sys.exit(0)

if __name__ == "__main__":
    main()
