#!/usr/bin/env python3
"""
xlsx_merge.py - 3-way merge of DesignData xlsx (and descriptor/PostConvert text) changes
from a sprint-branch workspace into trunk/WorkSpcae3, without disturbing anything else in trunk.

    python xlsx_merge.py                     # dry-run report (default: Misc/WorkSpcae_Design -> trunk/WorkSpcae3)
    python xlsx_merge.py --apply             # write into trunk via Excel COM (lossless), then verify
    python xlsx_merge.py --remap remap.csv   # rename colliding IDs first (see remap_suggest.csv from a dry run)

BASE = branch revision the last merge covered (state.json), else the branch copy point.
Row identity = descriptor "Key" column (row-1 header, first line). Comment rows ("//...") match by content.
Cells are read as their cached VALUES (what c_data reads): a key like "=A1366+400" is the number it computes,
so rows never look new just because a formula was re-based. Formula cells are therefore written as values.
Trunk is only touched where the branch changed something AND trunk still equals BASE there; everything
else is reported as a conflict for a human, unless a policy says who wins:
  --prefer-branch / --prefer-trunk   per sheet: both-sides-changed cells and same-key collisions take that side
  --yield-trunk                      per sheet: a genuinely new branch row that clashes keeps its ID; trunk's row
                                     (and every trunk reference to it) is renamed instead
  --auto-ids                         apply the tool's next-free-ID suggestions instead of writing them to remap_suggest.csv
  --scope / --scope-include          merge only the rows, files and layout entries that belong to the named events
  --switch-off                       new event rows go in with their toggle columns set to FALSE
Trunk-only rows/columns/sheets are never modified or deleted.
"""
import argparse, csv, datetime, glob, json, math, os, re, shutil, subprocess, sys, urllib.parse
from collections import Counter, defaultdict

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
import openpyxl

TOOL_DIR = os.path.dirname(os.path.abspath(__file__))
# cache / state.json / out live outside the script folder: the desktop app rebuilds its Python
# bundle from scratch, and a lost state.json would silently reset the base to the copy point.
DATA_DIR = os.path.join(os.environ.get("LOCALAPPDATA") or TOOL_DIR, "XlsxMerge")
DEF_BRANCH = r"C:\SkipBo\branches\Misc\design\DesignData\WorkSpcae_Design"
DEF_TRUNK = r"C:\SkipBo\trunk\design\DesignData\WorkSpcae3"
SKIP_PREFIXES = ("convert/", "backup_xlsx/", "descriptor/backup/")   # generated output / archives: never merged
DESC_FAMILIES = ("app_client", "server", "gm_server", "h5_client")   # Key lookup priority
SWITCH_COLS = "活动开关,控制开关,后处理开关,开关,活动开启"              # toggle columns set FALSE by --switch-off
LAYOUT_LUA = "postconvert/lua/convert_layout.lua"
CONVERT_JSON = "descriptor/app_client/convert.json"

# ----------------------------------------------------------------------------- helpers
def run(cmd):
    p = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace")
    if p.returncode != 0:
        raise RuntimeError(f"{' '.join(cmd)}\n{p.stderr.strip()}")
    return p.stdout

def svn_item(path, item):
    return run(["svn", "info", "--show-item", item, path]).strip()

def norm(v):
    if v is None: return None
    if isinstance(v, str): return v if v != "" else None
    if isinstance(v, float) and v.is_integer(): return int(v)
    return v

def _numlike(v):
    """'3000' and 3000 are the same cell to c_data (it reads text); '0012' is not 12."""
    if isinstance(v, str) and re.fullmatch(r"-?\d+", v.strip()) and str(int(v)) == v.strip(): return int(v)
    if isinstance(v, str) and re.fullmatch(r"-?\d+\.\d+", v.strip()):
        try:
            f = float(v)
            if str(f) == v.strip(): return int(f) if f.is_integer() else f
        except ValueError: pass
    return v

def eq(a, b):
    a, b = norm(a), norm(b)
    if isinstance(a, bool) != isinstance(b, bool): return False
    if a == b: return True
    return type(a) is not type(b) and _numlike(a) == _numlike(b) and not isinstance(_numlike(a), str)

def first_line(v):
    return None if v is None else str(v).split("\n")[0].strip()

def is_comment(cells):
    c0 = cells[0] if cells else None
    return isinstance(c0, str) and c0.lstrip().startswith("//")

def is_formula(v):
    return isinstance(v, str) and v.startswith("=")

def blank_row(cells):
    return all(c is None or (isinstance(c, str) and c.strip() == "") for c in cells)

def key_str(k):
    return " | ".join(str(c) for c in k[1] if c is not None)[:40] if k[0] == "#" else str(k[1])

def parse_sheet_specs(spec):
    """'*' | 'events.xlsx' | 'events' | 'events.xlsx!event_type', comma separated -> list of (xlsx, sheet|None|'*')"""
    out = []
    for part in (spec or "").split(","):
        part = part.strip()
        if not part: continue
        if part == "*": out.append(("*", None)); continue
        x, _, s = part.partition("!")
        out.append((x[:-5] if x.lower().endswith(".xlsx") else x, s or None))
    return out

def sheet_matches(specs, xlsx, sheet):
    return any(x == "*" or (x == xlsx and (s is None or s == sheet)) for x, s in specs)

# ----------------------------------------------------------------------------- descriptors
def lenient_json(path):
    txt = open(path, encoding="utf-8-sig").read()
    try: return json.loads(txt)
    except json.JSONDecodeError:
        txt = re.sub(r"//[^\n]*", "", txt)
        txt = re.sub(r",(\s*[}\]])", r"\1", txt)
        return json.loads(txt)

class Descriptors:
    """(xlsx, sheet) -> key column name; and Ref graph for ID propagation."""
    def __init__(self, *workspaces):
        self.key = {}                       # (xlsx, sheet) -> key header name
        self.by_json = defaultdict(set)     # "foo.json" -> {(xlsx, sheet)} (multi-sheet directors feed one json)
        self.refs = defaultdict(set)        # (xlsx, sheet, column) -> {(xlsx, sheet)} referenced
        self.bad = []
        fields = []
        for ws in workspaces:
            for fam in DESC_FAMILIES:
                for p in sorted(glob.glob(os.path.join(ws, "descriptor", fam, "*.json"))):
                    try: d = lenient_json(p)
                    except Exception as e:
                        self.bad.append(f"{fam}/{os.path.basename(p)}: {e}"); continue
                    for dr in d.get("Directors", []):
                        sheets = [(s.get("ExcelName"), s.get("SheetName")) for s in dr.get("ExcelSheets", [])]
                        k = dr.get("Key")
                        for xs in sheets:
                            if k and xs not in self.key: self.key[xs] = k
                            if dr.get("FileName"): self.by_json[dr["FileName"]].add(xs)
                        for f in dr.get("Fields", []):
                            if f.get("Ref") and f.get("Name"):
                                for xs in sheets: fields.append((xs, first_line(f["Name"]), f["Ref"]))
        for xs, col, ref in fields:
            for tgt in self.by_json.get(ref, ()): self.refs[(xs[0], xs[1], col)].add(tgt)
        self.referrers = defaultdict(set)   # (xlsx, sheet) -> {(xlsx2, sheet2, column)} that point at it
        for src, tgts in self.refs.items():
            for t in tgts: self.referrers[t].add(src)

# ----------------------------------------------------------------------------- sheet model
class SheetData:
    def __init__(self, name, headers, raw_headers, key_idx, rows, order, last_row):
        self.name, self.headers, self.raw_headers, self.key_idx = name, headers, raw_headers, key_idx
        self.formulas = set()   # (excel_row, col_idx) of formula cells; only filled for the branch, for reporting
        self.rows, self.order, self.last_row = rows, order, last_row   # rows: key -> {"row": excel_row, "vals": {hdr: val}}
        self.named = [h for h in headers if not h.startswith("_col")]   # unnamed columns are not exported by c_data: ignored
    @property
    def key_col(self):
        return self.headers[self.key_idx] if self.key_idx is not None else None
    def rekey(self, new_key):
        """new_key: {old k -> new key value}. One pass, occurrence numbers recomputed."""
        new_rows, new_order, occ = {}, [], Counter()
        for k in self.order:
            b = ("K", str(norm(new_key[k])).strip()) if k in new_key else k[:2]
            occ[b] += 1; nk = b + (occ[b],)
            new_rows[nk] = self.rows[k]; new_order.append(nk)
        self.rows, self.order = new_rows, new_order

def load_workbook_sheets(path, desc, xlsx_name, want_formulas=False):
    # data_only: cached values. Keys/compare/write all use what c_data will read; formula text is
    # never used for identity (it shifts whenever rows move) and never copied into trunk.
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    out = {}
    for ws in wb.worksheets:
        out[ws.title] = load_sheet(ws, desc.key.get((xlsx_name, ws.title)))
    wb.close()
    if want_formulas:   # second pass only to know which cells were formulas (reported, and checked for missing caches)
        wf = openpyxl.load_workbook(path, read_only=True)
        for ws in wf.worksheets:
            sd = out.get(ws.title)
            if sd is None: continue
            for r_i, cells in enumerate(ws.iter_rows(values_only=True), start=1):
                for c_i, v in enumerate(cells):
                    if is_formula(v): sd.formulas.add((r_i, c_i))
        wf.close()
    return out

def load_sheet(ws, key_name):
    it = ws.iter_rows(values_only=True)
    hdr = list(next(it, None) or ())
    width = max(len(hdr), ws.max_column or 0)
    hdr += [None] * (width - len(hdr))
    headers, seen = [], Counter()
    for i, h in enumerate(hdr):
        n = first_line(h) or f"_col{i+1}"
        seen[n] += 1
        headers.append(n if seen[n] == 1 else f"{n}#{seen[n]}")
    key_idx = headers.index(key_name) if key_name in headers else (0 if headers else None)
    rows, order, occ, last_row = {}, [], Counter(), 1
    for r_i, cells in enumerate(it, start=2):
        cells = tuple(cells or ())
        if blank_row(cells): continue
        cells = cells[:width] + (None,) * (width - len(cells))
        last_row = r_i
        kv = cells[key_idx] if key_idx is not None else None
        if is_comment(cells) or kv is None or str(kv).strip() == "":
            content = tuple(norm(c) for c in cells)
            while content and content[-1] is None: content = content[:-1]   # width-independent: trunk may have extra columns
            base = ("#", content)
        else:
            base = ("K", str(norm(kv)).strip())
        occ[base] += 1
        k = base + (occ[base],)
        rows[k] = {"row": r_i, "vals": {h: norm(v) for h, v in zip(headers, cells)}}
        order.append(k)
    return SheetData(ws.title, headers, hdr, key_idx, rows, order, last_row)

def same_row(a, b, cols):
    return all(eq(a["vals"].get(h), b["vals"].get(h)) for h in cols)

def empty_ops():
    return {"add": [], "mod": [], "del": [], "skip": 0, "conflict": [], "new_cols": [], "alias_dup": [], "prefer": Counter()}

# ----------------------------------------------------------------------------- scope
def row_in_scope(k, vals, tokens):
    text = (key_str(k) + " " + " ".join(str(v) for v in vals.values() if v is not None and not isinstance(v, bool))).lower()
    return any(t in text for t in tokens)

def scope_filter(br, base, tokens):
    """Neutralise branch rows that do not mention a scope token: revert them to their base content (so they
    show no change) or drop them when the branch added them; put back base rows the branch deleted.
    Returns the number of rows taken out of scope."""
    out = 0
    for k in list(br.order):
        if k[0] != "K" and base and k in base.rows: continue       # unchanged comment rows: irrelevant either way
        if row_in_scope(k, br.rows[k]["vals"], tokens): continue
        if base and k in base.rows:
            if not same_row(br.rows[k], base.rows[k], br.named):
                br.rows[k]["vals"] = dict(base.rows[k]["vals"]); out += 1
        else:
            br.order.remove(k); del br.rows[k]; out += 1
    if base:
        for k, r in base.rows.items():
            if k not in br.rows and not row_in_scope(k, r["vals"], tokens):
                br.rows[k] = {"row": r["row"], "vals": dict(r["vals"])}; br.order.append(k)   # deletion out of scope
    return out

def switch_off(br, ops, cols, tokens=()):
    """--switch-off: added rows of OTHER events go in with their toggle columns FALSE (they are merged before they
    go live); rows mentioning a scope token belong to the event being released and keep their switches."""
    n = 0
    for k, _ in ops["add"]:
        if tokens and row_in_scope(k, br.rows[k]["vals"], tokens): continue
        for c in cols:
            v = br.rows[k]["vals"].get(c)
            if c in br.named and v is not None and not (v is False or str(v).strip().upper() == "FALSE"):
                br.rows[k]["vals"][c] = "FALSE" if isinstance(v, str) else False; n += 1
    return n

def expand_scope(tokens, branch_ws):
    """Rows of the release event rarely spell its name: they carry its id (20260902) or Alias. Add the id and
    Alias of every events.xlsx!events row that matches a token (and the event_type/entrance/score rows' aliases)."""
    p = os.path.join(branch_ws, "design", "events.xlsx")
    if not os.path.isfile(p): return tokens
    out = list(tokens)
    try:
        wb = openpyxl.load_workbook(p, read_only=True, data_only=True)
        for sn in ("events", "event_type", "event_entrance", "event_score"):
            if sn not in wb.sheetnames: continue
            for cells in wb[sn].iter_rows(min_row=2, values_only=True):
                if not cells or cells[0] is None or is_comment(cells): continue
                text = " ".join(str(c) for c in cells if c is not None).lower()
                if any(t in text for t in tokens):
                    for v in cells[:2]:
                        if v is not None and str(norm(v)).strip() and str(norm(v)).strip().lower() not in out: out.append(str(norm(v)).strip().lower())
        wb.close()
    except Exception: pass
    return out

# ----------------------------------------------------------------------------- 3-way diff
def diff_sheet(base, br, tr, add_missing=False, on_collision="conflict", prefer=None):
    """Return dict of ops for one sheet. base may be None (sheet new in branch). tr must exist.
    on_collision: branch added a key trunk already has with different content: 'conflict' | 'skip' | 'overwrite'.
    prefer: None | 'branch' | 'trunk' - who wins both-sides-changed cells and same-key collisions on this sheet."""
    ops = empty_ops()
    if prefer == "branch": on_collision = "overwrite"
    elif prefer == "trunk": on_collision = "skip"
    elif prefer is not None: raise ValueError(prefer)
    brows = base.rows if base else {}
    ops["new_cols"] = [h for h in br.named if h not in tr.headers]
    shared = [h for h in br.named if h in tr.headers]
    tr_alias = {r["vals"].get("Alias") for r in tr.rows.values() if r["vals"].get("Alias") is not None} if "Alias" in tr.headers else set()
    for k, brow in brows.items():
        if k in br.rows or k not in tr.rows: continue
        if same_row(tr.rows[k], brow, base.named):
            ops["del"].append(k)
        else:
            ops["conflict"].append(("del-vs-trunk-mod", k, "branch deleted row; trunk changed it since base"))
    prev = None
    for k in br.order:
        brow = br.rows[k]
        if k not in brows:
            if k in tr.rows:
                if same_row(brow, tr.rows[k], shared):
                    fill = {h: brow["vals"][h] for h in ops["new_cols"] if brow["vals"].get(h) is not None}
                    if fill: ops["mod"].append((k, fill))   # already in trunk; only the new columns need values
                    else: ops["skip"] += 1
                else:
                    diffc = [h for h in shared if not eq(brow["vals"].get(h), tr.rows[k]["vals"].get(h))]
                    b_al, t_al = brow["vals"].get("Alias"), tr.rows[k]["vals"].get("Alias")
                    same_row_ = "Alias" not in br.named or (b_al is not None and t_al is not None and str(b_al).strip() == str(t_al).strip())
                    policy = on_collision if (prefer is None or same_row_) else "conflict"   # a different Alias is a real ID clash: never decided by --prefer
                    if policy == "skip": ops["skip"] += 1; ops["prefer"]["trunk"] += 1 if prefer else 0
                    elif policy == "overwrite":
                        ops["mod"].append((k, {h: brow["vals"].get(h) for h in br.named if not eq(brow["vals"].get(h), tr.rows[k]["vals"].get(h))}))
                        ops["prefer"]["branch"] += 1 if prefer else 0
                    else: ops["conflict"].append(("add-collision", k, f"key already in trunk with different content; differing cols: {diffc[:6]}"))
            elif k[0] == "K" and k[2] > 1 and ("K", k[1], 1) in tr.rows:
                # duplicate key inside the branch sheet (usually a remap that renamed a row onto a sibling's key):
                # adding it would give trunk two rows with one key, which c_data rejects.
                if on_collision == "skip": ops["skip"] += 1
                else: ops["conflict"].append(("add-collision", k, f"branch has {k[2]} rows with key {k[1]} (after remap); trunk already has one"))
            else:
                ops["add"].append((k, prev))
                al = brow["vals"].get("Alias")
                if k[0] == "K" and al is not None and al in tr_alias: ops["alias_dup"].append((k, al))
        else:
            changed = {h: brow["vals"].get(h) for h in br.named if not eq(brow["vals"].get(h), brows[k]["vals"].get(h))}
            if changed:
                if k not in tr.rows:
                    if add_missing: ops["add"].append((k, prev))
                    else: ops["conflict"].append(("mod-vs-trunk-missing", k, "branch modified a row trunk does not have (use --add-missing to add it)"))
                else:
                    trow = tr.rows[k]
                    trunk_changed = {h for h in changed if not eq(trow["vals"].get(h), brows[k]["vals"].get(h))}
                    clash = sorted(h for h in trunk_changed if not eq(trow["vals"].get(h), changed[h]))
                    if clash and prefer is None:
                        ops["conflict"].append(("mod-overlap", k, f"both sides changed cols {clash[:6]}"))
                    else:
                        if clash: ops["prefer"][prefer] += len(clash)
                        cells = {h: v for h, v in changed.items() if h not in trunk_changed or prefer == "branch"}
                        if cells: ops["mod"].append((k, cells))
        prev = k
    return ops

# ----------------------------------------------------------------------------- lint (what c_data would reject)
def merged_view(plan_entry):
    """Rows of a sheet as they will be after apply: trunk rows + modified cells + added rows - deleted rows."""
    ops, br, tr = plan_entry["ops"], plan_entry["br"], plan_entry["tr"]
    if tr is None:
        return {k: dict(r["vals"]) for k, r in br.rows.items()}, set(br.rows)
    rows = {k: dict(r["vals"]) for k, r in tr.rows.items()}
    touched = set()
    for k in ops["del"]: rows.pop(k, None)
    for k, cells in ops["mod"]:
        rows.setdefault(k, {}).update(cells); touched.add(k)
    for k, _ in ops["add"]:
        rows[k] = dict(br.rows[k]["vals"]); touched.add(k)
    return rows, touched

def _tokens(v):
    if v is None or isinstance(v, bool): return []
    return [t.split(":")[0].strip() for t in re.split(r"[,;|\n]", str(norm(v))) if t.strip()]

def lint_merged(plans, desc, trunk_ws):
    """Duplicate keys / Aliases introduced by the merge, and references from merged rows that resolve nowhere.
    Returns conflict-style rows (xlsx, sheet, kind, key, detail)."""
    views = {}
    for rel, plan in plans.items():
        x = os.path.basename(rel)[:-5]
        for sname, p in plan.items():
            if p["ops"] is None and p["tr"] is None:  # whole new sheet
                views[(x, sname)] = ({k: dict(r["vals"]) for k, r in p["br"].rows.items()}, set(p["br"].rows))
            elif p["ops"] is not None:
                views[(x, sname)] = merged_view(p)
    cache = {}
    def target_ids(xs):
        if xs in views: rows = views[xs][0]
        else:
            if xs[0] not in cache:
                path = os.path.join(trunk_ws, "design", xs[0] + ".xlsx")
                cache[xs[0]] = load_workbook_sheets(path, desc, xs[0]) if os.path.isfile(path) else {}
            sd = cache[xs[0]].get(xs[1])
            rows = {k: r["vals"] for k, r in sd.rows.items()} if sd else {}
        keys = {k[1] for k in rows if k[0] == "K"}
        aliases = {str(v.get("Alias")).strip() for v in rows.values() if v.get("Alias") is not None}
        return keys, aliases
    findings = []
    for (x, sname), (rows, touched) in views.items():
        if not touched: continue
        by_key, by_alias = defaultdict(list), defaultdict(list)
        for k, v in rows.items():
            if k[0] != "K": continue                       # comment rows are not exported
            by_key[k[1]].append(k)
            if v.get("Alias") is not None: by_alias[str(v.get("Alias")).strip()].append(k)
        for key, ks in by_key.items():
            if len(ks) > 1 and any(k in touched for k in ks):
                findings.append((x + ".xlsx", sname, "dup-key", key, f"{len(ks)} rows share this key after merge"))
        for al, ks in by_alias.items():
            if len(ks) > 1 and any(k in touched for k in ks):
                findings.append((x + ".xlsx", sname, "dup-alias", key_str(ks[-1]), f"Alias {al!r} used by {len(ks)} rows after merge: {[key_str(k) for k in ks][:4]}"))
        cols = {h: tgts for (xx, ss, h), tgts in desc.refs.items() if xx == x and ss == sname}
        if not cols: continue
        tid = {}
        for k in touched:
            v = rows.get(k)
            if not v or k[0] != "K": continue
            for h, tgts in cols.items():
                for tok in _tokens(v.get(h)):
                    if re.fullmatch(r"-?\d+(\.\d+)?", tok) and float(tok) == 0: continue
                    ok = False
                    for xs in tgts:
                        if xs not in tid: tid[xs] = target_ids(xs)
                        if tok in tid[xs][0] or tok in tid[xs][1]: ok = True; break
                    if not ok:
                        findings.append((x + ".xlsx", sname, "dangling-ref", key_str(k), f"{h} = {tok!r} not found in {', '.join(f'{a}!{b}' for a, b in sorted(tgts))}"))
    return findings

# ----------------------------------------------------------------------------- remap (ID changes)
def load_remap(path):
    m = defaultdict(dict)
    with open(path, encoding="utf-8-sig", newline="") as f:
        for row in csv.reader(f):
            if not row or row[0].startswith("#") or row[0].lower() == "xlsx": continue
            xlsx, sheet, old, new = [c.strip() for c in row[:4]]
            if new: m[(xlsx.replace(".xlsx", ""), sheet)][old] = new
    return m

def _retype(new, like):
    return int(new) if isinstance(like, (int, float)) and not isinstance(like, bool) and re.fullmatch(r"-?\d+", str(new)) else new

def _replace_tokens(v, table):
    if v is None or isinstance(v, bool): return v, False
    if isinstance(v, (int, float)):
        s = str(norm(v))
        return (_retype(table[s], v), True) if s in table else (v, False)
    if isinstance(v, str) and not is_formula(v):
        parts = v.split(",")
        hit = False
        for i, p in enumerate(parts):
            t = p.strip()
            if t in table: parts[i] = p.replace(t, str(table[t])); hit = True
        return (",".join(parts), hit)
    return v, False

def apply_remap(xlsx, sheets_br, sheets_base, remap, desc, log):
    """Rename keys of branch-added rows and rewrite references in rows that are part of the delta.
    All rewrites are computed against the ORIGINAL values and applied as one simultaneous substitution,
    then rows are re-keyed once. Doing it stepwise chains remaps (6149->6150->106150 when the table holds
    both 6149->6150 and 6150->106150) and lets a renamed key clobber a sibling row."""
    for sname, br in sheets_br.items():
        base = sheets_base.get(sname) if sheets_base else None
        table = remap.get((xlsx, sname), {})
        key_col = br.key_col
        new_key = {}   # k -> new key value
        # 1. sheet-level renames of branch-added rows (marked; their key cell is final)
        if table and key_col:
            for k in br.order:
                if k[0] == "K" and k[1] in table:
                    if base and k in base.rows:
                        raise SystemExit(f"remap {xlsx}!{sname} {k[1]}->{table[k[1]]}: source key exists in BASE (not a branch-added row); refusing")
                    r = br.rows[k]
                    r["vals"][key_col] = _retype(table[k[1]], r["vals"][key_col]); new_key[k] = table[k[1]]
                    log.append(f"remap key {xlsx}!{sname}: {k[1]} -> {table[k[1]]}")
        # 2. reference propagation: columns whose descriptor Ref points at a remapped sheet
        for h in br.headers:
            for tgt in desc.refs.get((xlsx, sname, h), ()):
                t = remap.get(tgt)
                if not t: continue
                for k, r in br.rows.items():
                    if h == key_col and k in new_key: continue      # already renamed above; never chain
                    in_delta = not base or k not in base.rows or not same_row(r, base.rows[k], br.named)
                    if not in_delta: continue
                    nv, hit = _replace_tokens(r["vals"].get(h), t)
                    if hit:
                        log.append(f"remap ref {xlsx}!{sname}[{key_str(k)}].{h}: {r['vals'].get(h)!r} -> {nv!r}")
                        r["vals"][h] = nv
                        if h == key_col and k[0] == "K":            # foreign-key key column: identity follows
                            new_key[k] = nv; log.append(f"rekey {xlsx}!{sname}: {k[1]} -> {nv} (key column follows {h} remap)")
        # 3. one re-keying pass, occurrence numbers recomputed so nothing is clobbered or duplicated
        if new_key: br.rekey(new_key)

# ----------------------------------------------------------------------------- ID suggestions
def _group_modulus(existing):
    """Largest M in (10000, 1000, 100) at which the sheet's numeric keys look like <group>*M + <step>:
    steps start near 0 and stay well below M. None = plain sequential IDs."""
    for M in (10000, 1000, 100):
        if len(existing) < 2: break
        steps = [k % M for k in existing]
        groups = Counter(k // M for k in existing)
        populated = [g for g, n in groups.items() if n >= 2]
        if len(populated) >= 2 and min(steps) < M // 10 and max(steps) < M // 2 and len(existing) / len(groups) >= 1.5:
            return M
    return None

def _next_in_block(n, ex):
    """Smallest free number above the largest existing one that shares n's digit count and leading digit
    (used for group prefixes: 1159 -> 1178, never 6301)."""
    block = [k for k in ex if len(str(k)) == len(str(n)) and str(k)[0] == str(n)[0]]
    c = (max(block) if block else n) + 1
    while c in ex: c += 1
    return c

def _next_after(o, ex):
    """Smallest free id above the clashing one with the same digit count (3345 -> 3346, 9885 -> 9888 when
    9886/9887 clash too) - what designers type by hand; falls back to the block rule if the range overflows."""
    n = o + 1
    while n in ex: n += 1
    return n if len(str(n)) == len(str(o)) else _next_in_block(o, ex)

def suggest_ids(old_ids, existing, taken, existing_str=()):
    """Next-free IDs for colliding numeric keys, the way designers number by hand.
    old_ids: colliding branch keys (str); existing: all numeric keys of the merged sheet (trunk + kept branch rows);
    taken: set updated with every id handed out (so several sheets / passes never collide).
    Returns {old: new} (str -> str) plus a one-line rule description."""
    out = {}
    # string keys ending in a number (CARD_SKIN_MID_062, FRAME_SHOP_95): bump to the next free number, same width
    str_olds = [o for o in old_ids if not re.fullmatch(r"-?\d+", o) and re.search(r"\d+$", o)]
    if str_olds:
        ex_str = set(existing_str) if isinstance(existing_str, (set, list)) else set()
        for o in sorted(str_olds):
            pre, num = re.match(r"^(.*?)(\d+)$", o).groups()
            same = [int(re.search(r"\d+$", e).group()) for e in ex_str if e.startswith(pre) and re.fullmatch(re.escape(pre) + r"\d+", e)]
            n = (max(same) if same else int(num)) + 1
            while f"{pre}{n:0{len(num)}d}" in ex_str: n += 1
            out[o] = f"{pre}{n:0{len(num)}d}"; ex_str.add(out[o])
    olds = sorted({int(o) for o in old_ids if re.fullmatch(r"-?\d+", o)})
    if not olds: return out, ("string keys: next free numeric suffix" if out else "")
    ex = set(existing) | taken
    unit = 0
    for k in existing: unit = math.gcd(unit, k)
    if unit >= 10 and len(existing) >= 3:
        # every id is a multiple of `unit` (card_skin_atlas: 0, 10000, ..., 620000): the next id is the next free multiple
        for o in olds:
            n = (o // unit + 1) * unit
            while n in ex: n += unit
            out[str(o)] = str(n); ex.add(n)
        taken.update(int(v) for v in out.values() if re.fullmatch(r"-?\d+", v))
        return out, f"ids step by {unit}: next free multiple"
    M = _group_modulus(sorted(existing))
    if M:
        by_group = defaultdict(list)
        for o in olds: by_group[o // M].append(o)
        rule = f"groups of {M}: "
        prefixes = {k // M for k in ex}
        for g, members in sorted(by_group.items()):
            in_trunk = {k for k in ex if k // M == g}
            collided = len({o for o in members} & in_trunk)
            if in_trunk and collided / max(len(in_trunk), 1) < 0.5:
                # the branch appended a few rows onto a group trunk also extended: continue after the clash
                for o in members:
                    n = _next_after(o, ex); out[str(o)] = str(n); ex.add(n)
                rule += f"{g}->continue; "
            else:
                # the whole group clashes (both sides created it): the group moves to the next free prefix in
                # its own block (1159 -> 1178, not 6301), steps kept
                next_prefix = _next_in_block(g, prefixes); prefixes.add(next_prefix)
                for o in members:
                    n = next_prefix * M + o % M
                    while n in ex: n += 1
                    out[str(o)] = str(n); ex.add(n)
                rule += f"{g}->{next_prefix}; "
    else:
        rule = "next free id after the clashing one"
        for o in olds:
            n = _next_after(o, ex); out[str(o)] = str(n); ex.add(n)
    taken.update(int(v) for v in out.values() if re.fullmatch(r"-?\d+", v))
    return out, rule.rstrip("; ")

# ----------------------------------------------------------------------------- svn / base
def branch_changes(branch_ws, base_rev):
    out = run(["svn", "diff", "--summarize", "-r", str(base_rev), branch_ws])
    changes = []
    root = os.path.abspath(branch_ws)
    for line in out.splitlines():
        if not line.strip(): continue
        st, path = line[0], line[8:].strip()
        rel = os.path.relpath(os.path.abspath(path), root).replace("\\", "/")   # relpath keeps the repo's case
        if rel == "." or rel.lower().startswith(SKIP_PREFIXES) or os.path.basename(rel).startswith("~$"): continue
        if os.path.isdir(os.path.join(branch_ws, rel)): continue
        changes.append((st, rel))
    return changes

def fetch_base(branch_url, base_rev, rel):
    out = os.path.join(DATA_DIR, "cache", "base", str(base_rev), rel)
    if os.path.isfile(out) and os.path.getsize(out) > 0: return out
    os.makedirs(os.path.dirname(out), exist_ok=True)
    p = subprocess.run(["svn", "cat", "-r", str(base_rev), f"{branch_url}/{urllib.parse.quote(rel)}@{base_rev}"], capture_output=True)
    if p.returncode != 0 or not p.stdout:
        return None
    open(out, "wb").write(p.stdout)
    return out

def copy_point(branch_url):
    out = run(["svn", "log", "-q", "--stop-on-copy", branch_url])
    revs = [int(l.split()[0][1:]) for l in out.splitlines() if l.startswith("r")]
    return min(revs) if revs else None

def state_path():
    return os.path.join(DATA_DIR, "state.json")

def load_state():
    p = state_path()
    return json.load(open(p, encoding="utf-8")) if os.path.isfile(p) else {}

# ----------------------------------------------------------------------------- excel COM apply
XL_DOWN, XL_FMT_ABOVE, XL_MANUAL, XL_AUTO = -4121, 0, -4135, -4105

class Excel:
    def __init__(self):
        try: import win32com.client
        except ImportError: raise SystemExit("--apply needs pywin32 (pip install pywin32) and a local Excel; dry-run works without them")
        self.app = win32com.client.DispatchEx("Excel.Application")
        self.app.Visible = False; self.app.DisplayAlerts = False; self.app.AskToUpdateLinks = False
        self.app.ScreenUpdating = False; self.app.EnableEvents = False
    def open(self, path, read_only=False):
        wb = self.app.Workbooks.Open(os.path.abspath(path), 0, read_only)
        if not read_only and wb.ReadOnly:
            wb.Close(False); raise RuntimeError(f"{path} opened read-only (file locked or open elsewhere)")
        return wb
    def quit(self):
        try: self.app.Quit()
        except Exception: pass

def write_cell(ws, r, c, v):
    cell = ws.Cells(r, c)
    if v is None: cell.ClearContents(); return
    if isinstance(v, str): cell.NumberFormat = "@"   # ponytail: text stays text ("0012", "TRUE", "=x", "1/2" never get coerced)
    cell.Value2 = v

def apply_workbook(xl, trunk_path, branch_path, plan, report):
    """plan: {sheet: {"ops", "br", "tr"}}; writes trunk_path in place."""
    wb = xl.open(trunk_path)
    src_wb = None
    try:
        xl.app.Calculation = XL_MANUAL
        for sname, p in plan.items():
            ops, br, tr = p["ops"], p["br"], p["tr"]
            if tr is None:  # whole new sheet: copy from branch workbook
                src_wb = src_wb or xl.open(branch_path, read_only=True)
                src_wb.Worksheets(sname).Copy(None, wb.Worksheets(wb.Worksheets.Count))
                report.append(f"  + sheet {sname} copied from branch")
                continue
            ws = wb.Worksheets(sname)
            col = {h: i + 1 for i, h in enumerate(tr.headers)}
            for h in ops["new_cols"]:
                c = len(col) + 1; col[h] = c
                ws.Cells(1, c).Value2 = br.raw_headers[br.headers.index(h)]
            for k, cells in ops["mod"]:
                r = tr.rows[k]["row"]
                for h, v in cells.items():
                    write_cell(ws, r, col[h], v)
            # group adds by the trunk row they follow (walk branch order so consecutive adds stay together)
            groups, anchor_row = defaultdict(list), {}
            for k, prev in ops["add"]:
                if prev is None: a = 1
                elif prev in anchor_row: a = anchor_row[prev]
                elif prev in tr.rows: a = tr.rows[prev]["row"]
                else: a = tr.last_row  # predecessor unknown in trunk: append at end
                groups[a].append(k); anchor_row[k] = a
            del_rows = sorted((tr.rows[k]["row"] for k in ops["del"]), reverse=True)
            actions = [(a + 1, "ins", ks) for a, ks in groups.items()] + [(r, "del", None) for r in del_rows]
            for pos, kind, ks in sorted(actions, key=lambda t: (t[0], t[1] == "ins"), reverse=True):
                if kind == "del":
                    ws.Rows(pos).Delete(); continue
                ws.Rows(f"{pos}:{pos + len(ks) - 1}").Insert(XL_DOWN, XL_FMT_ABOVE)
                for i, k in enumerate(ks):
                    r, vals = pos + i, br.rows[k]["vals"]
                    for h, c in col.items():
                        if not h.startswith("_col") and vals.get(h) is not None:
                            write_cell(ws, r, c, vals[h])
            report.append(f"  ~ {sname}: +{len(ops['add'])} rows, ~{len(ops['mod'])} rows, -{len(ops['del'])} rows, +{len(ops['new_cols'])} cols")
        if src_wb:
            links = wb.LinkSources(1)
            for link in (links or ()): wb.BreakLink(link, 1)
        xl.app.Calculation = XL_AUTO
        wb.Save()
    finally:
        wb.Close(False)
        if src_wb: src_wb.Close(False)

def verify_workbook(trunk_path, plan, desc, xlsx):
    after = load_workbook_sheets(trunk_path, desc, xlsx)
    problems = []
    for sname, p in plan.items():
        if p["tr"] is None:
            if sname not in after: problems.append(f"{sname}: sheet missing after copy")
            continue
        ta, br, ops = after.get(sname), p["br"], p["ops"]
        if ta is None: problems.append(f"{sname}: sheet vanished"); continue
        for k, _ in ops["add"]:
            if k not in ta.rows: problems.append(f"{sname}: added row {key_str(k)} not found"); continue
            for h in br.named:
                v = br.rows[k]["vals"].get(h)
                if not eq(v, ta.rows[k]["vals"].get(h)): problems.append(f"{sname}[{key_str(k)}].{h}: wrote {v!r}, read {ta.rows[k]['vals'].get(h)!r}")
        for k, cells in ops["mod"]:
            kk = k
            if p["tr"].key_col in cells and k not in ta.rows:        # a renamed trunk row (--yield-trunk) is read back under its new key
                kk = ("K", str(norm(cells[p["tr"].key_col])).strip(), 1)
            for h, v in cells.items():
                got = ta.rows[kk]["vals"].get(h) if kk in ta.rows else "<row missing>"
                if not eq(v, got): problems.append(f"{sname}[{key_str(k)}].{h}: wrote {v!r}, read {got!r}")
        for k in ops["del"]:
            if k in ta.rows: problems.append(f"{sname}: deleted row {key_str(k)} still present")
        exp = len(p["tr"].rows) + len(ops["add"]) - len(ops["del"])
        if len(ta.rows) != exp: problems.append(f"{sname}: row count {len(ta.rows)} != expected {exp}")
    return problems

# ----------------------------------------------------------------------------- text files
def _nl(b): return b"\r\n" if b"\r\n" in b else b"\n"

def merge_convert_layout(tr_bytes, br_bytes, base_bytes, scope_tokens):
    """Union of {json = ...} entries per section: branch-added entries (not in base, not in trunk) are appended
    at the end of the same section in trunk, in trunk's line endings. Returns (bytes, added_count) or None."""
    nl = _nl(tr_bytes)
    ent = re.compile(rb'json\s*=\s*"([^"]+)"')
    def sections(text):
        out, cur, name = {}, [], None
        for line in text.replace(b"\r\n", b"\n").split(b"\n"):
            m = re.match(rb"^    (\w+)\s*=\s*\{", line)
            if m and not ent.search(line): name = m.group(1); out[name] = []
            elif name is not None and ent.search(line): out[name].append(line)
        return out
    tr_s, br_s = sections(tr_bytes), sections(br_bytes)
    base_names = set(ent.findall(base_bytes)) if base_bytes else set()
    tr_names = set(ent.findall(tr_bytes))
    if not tr_s: return None
    lines = tr_bytes.split(nl)
    added = 0
    for name, entries in br_s.items():
        new = [l for l in entries if (n := ent.search(l).group(1)) not in tr_names and n not in base_names
               and (not scope_tokens or any(t in n.decode("utf-8", "replace").lower() for t in scope_tokens))]
        if not new: continue
        # find the section's closing "    }" in trunk
        start = next((i for i, l in enumerate(lines) if re.match(rb"^    " + re.escape(name) + rb"\s*=\s*\{", l)), None)
        if start is None: return None
        end = next((i for i in range(start + 1, len(lines)) if re.match(rb"^    \},?\s*$", lines[i])), None)
        if end is None: return None
        indent = b"        "
        for i in range(end - 1, start, -1):
            if ent.search(lines[i]): indent = re.match(rb"^\s*", lines[i]).group(0); break
        block = [indent + l.strip() for l in new]
        lines[end:end] = block; added += len(block)
    return nl.join(lines), added

def merge_convert_json(tr_bytes, br_bytes, base_bytes, scope_tokens):
    """Union of ExcelSheets for directors present on both sides (const_event & co); text insertion keeps
    trunk's formatting and line endings. Returns (bytes, added_count) or None when directors differ otherwise."""
    def lenient(b):
        t = b.decode("utf-8-sig")
        try: return json.loads(t)
        except json.JSONDecodeError: return json.loads(re.sub(r",(\s*[}\]])", r"\1", re.sub(r"//[^\n]*", "", t)))
    try: tj, bj, basej = lenient(tr_bytes), lenient(br_bytes), (lenient(base_bytes) if base_bytes else {"Directors": []})
    except Exception: return None
    tmap = {d["FileName"]: d for d in tj["Directors"]}; bmap = {d["FileName"]: d for d in bj["Directors"]}
    basemap = {d["FileName"]: d for d in basej["Directors"]}
    if set(bmap) - set(tmap) - set(basemap): return None            # branch added a whole director: hand-merge
    text = tr_bytes.decode("utf-8-sig"); bom = tr_bytes.startswith(b"\xef\xbb\xbf")
    nl = "\r\n" if "\r\n" in text else "\n"
    added = 0
    for fn, bd in bmap.items():
        td = tmap.get(fn)
        if not td: continue
        base_sheets = basemap.get(fn, {}).get("ExcelSheets", [])
        missing = [s for s in bd.get("ExcelSheets", []) if s not in td.get("ExcelSheets", []) and s not in base_sheets
                   and (not scope_tokens or any(t in (s.get("ExcelName", "") + " " + s.get("SheetName", "")).lower() for t in scope_tokens))]
        if not missing: continue
        if [f for f in bd.get("Fields", []) if f not in td.get("Fields", []) and f not in basemap.get(fn, {}).get("Fields", [])]:
            return None                                              # branch changed Fields too: hand-merge
        i = text.index(f'"FileName": "{fn}"'); j = text.index('"ExcelSheets"', i); k = text.index("]", j)
        last = text.rfind("}", j, k)
        indent = re.search(r"\n([ \t]*)\{", text[j:k]).group(1); inner = re.search(r"\n([ \t]*)\"ExcelName\"", text[j:k]).group(1)
        ins = "".join(f",{nl}{indent}{{{nl}{inner}\"ExcelName\": \"{s['ExcelName']}\",{nl}{inner}\"SheetName\": \"{s['SheetName']}\"{nl}{indent}}}" for s in missing)
        text = text[:last + 1] + ins + text[last + 1:]; added += len(missing)
    out = text.encode("utf-8")
    return ((b"\xef\xbb\xbf" if bom else b"") + out), added

def merge_text(rel, base_path, br_path, tr_path, out_dir, apply, scope_tokens=()):
    br_bytes = open(br_path, "rb").read()
    if not os.path.isfile(tr_path):
        if apply: os.makedirs(os.path.dirname(tr_path), exist_ok=True); shutil.copy2(br_path, tr_path)
        return "copy-new"
    tr_bytes = open(tr_path, "rb").read()
    if tr_bytes == br_bytes: return "identical"
    base_bytes = open(base_path, "rb").read() if base_path else None
    low = rel.lower()
    if low.endswith(LAYOUT_LUA) or low == CONVERT_JSON:
        # the two files every sprint appends to: merge by meaning (union of entries / sheets), not by lines
        res = (merge_convert_layout if low.endswith(LAYOUT_LUA) else merge_convert_json)(tr_bytes, br_bytes, base_bytes, scope_tokens)
        if res is not None:
            new_bytes, added = res
            if added == 0: return "nothing to add (trunk already has every branch entry in scope)"
            if apply: open(tr_path, "wb").write(new_bytes)
            return f"union-merged: +{added} entries"
    if base_bytes is None: return "conflict (no base; trunk differs)"
    if tr_bytes == base_bytes:
        if apply: shutil.copy2(br_path, tr_path)
        return "copy (trunk unchanged since base)"
    if b"\0" in br_bytes[:4096] or b"\0" in tr_bytes[:4096]: return "conflict (binary, both changed)"
    try: p = subprocess.run(["git", "merge-file", "-p", "-L", "trunk", "-L", "base", "-L", "branch", tr_path, base_path, br_path], capture_output=True)
    except FileNotFoundError: return "conflict (both changed; git not installed, no 3-way text merge)"
    if p.returncode == 0:
        if apply: open(tr_path, "wb").write(p.stdout)
        return "3-way merged"
    out = os.path.join(out_dir, "text_conflicts", rel)
    os.makedirs(os.path.dirname(out), exist_ok=True); open(out, "wb").write(p.stdout)
    return f"conflict ({p.returncode} hunks) -> text_conflicts/{rel}"

# ----------------------------------------------------------------------------- planning pass
class Ctx:
    """Everything a planning pass needs; built once in main()."""
    pass

def plan_all(ctx, remap, trunk_renames):
    """One planning pass over every changed workbook. trunk_renames: {(xlsx, sheet): {old: new}} rows of TRUNK
    to rename (--yield-trunk), applied to the loaded trunk sheets before diffing; their references are
    rewritten across every trunk workbook. Returns a result dict; nothing is written."""
    a, desc = ctx.a, ctx.desc
    R = {"plans": {}, "conflicts": [], "suggest": [], "collide": [], "new_files": [], "remap_log": [], "alias_warn": [],
         "lines": [], "totals": Counter({k: 0 for k in ("add", "mod", "del", "new_cols", "new_sheets", "conflicts", "formulas_as_values",
                                                         "out_of_scope_rows", "out_of_scope_files", "switched_off", "prefer_branch_cells", "prefer_trunk_cells", "trunk_renamed")}),
         "loaded_tr": {}, "loaded_br": {}}
    totals = R["totals"]
    renamed_refs = []   # ((xlsx, sheet), old, new) for the trunk-wide reference rewrite after the loop
    for status, rel in ctx.xlsx_changes:
        name = os.path.basename(rel); xlsx = name[:-5]
        if ctx.only and name not in ctx.only: continue
        br_path, tr_path = os.path.join(ctx.branch_ws, rel), os.path.join(ctx.trunk_ws, rel)
        if status == "D":
            R["lines"].append(f"## {name}: deleted in branch -> " + ("trunk still has it (NOT deleted; remove by hand if wanted)" if os.path.isfile(tr_path) else "already absent in trunk")); continue
        if not os.path.isfile(tr_path):
            if ctx.scope and not (any(t in name.lower() for t in ctx.scope) or sheet_matches(ctx.scope_include, xlsx, None)):
                R["lines"].append(f"## {name}: new file, out of scope -> not copied"); totals["out_of_scope_files"] += 1; continue
            R["new_files"].append(rel); R["lines"].append(f"## {name}: new file -> copy to trunk (svn add)"); continue
        base_path = fetch_base(ctx.branch_url, ctx.base_rev, rel) if status == "M" else None
        if status == "M" and base_path is None: R["lines"].append(f"## {name}: WARNING no base at r{ctx.base_rev}; treating all branch rows as added")
        sheets_br = load_workbook_sheets(br_path, desc, xlsx, want_formulas=True)
        sheets_tr = load_workbook_sheets(tr_path, desc, xlsx)
        sheets_base = load_workbook_sheets(base_path, desc, xlsx) if base_path else {}
        R["loaded_tr"][xlsx] = sheets_tr; R["loaded_br"][xlsx] = sheets_br
        if ctx.scope:
            for sname, br in sheets_br.items():
                if sheet_matches(ctx.scope_include, xlsx, sname): continue
                totals["out_of_scope_rows"] += scope_filter(br, sheets_base.get(sname), ctx.scope)
        if remap: apply_remap(xlsx, sheets_br, sheets_base, remap, desc, R["remap_log"])
        extra_mods = defaultdict(list)   # sheet -> [(k, cells)] for renamed trunk rows
        for sname, table in ((s, t) for (x, s), t in trunk_renames.items() if x == xlsx):
            tr = sheets_tr.get(sname)
            if not tr or not tr.key_col: continue
            new_key = {}
            for k in tr.order:
                if k[0] == "K" and k[1] in table:
                    nv = _retype(table[k[1]], tr.rows[k]["vals"][tr.key_col])
                    tr.rows[k]["vals"][tr.key_col] = nv; new_key[k] = nv
                    R["remap_log"].append(f"trunk yields {xlsx}!{sname}: trunk row {k[1]} renamed -> {nv}")
                    renamed_refs.append(((xlsx, sname), k[1], str(nv)))
            if new_key:
                tr.rekey(new_key)
                for k, r in tr.rows.items():
                    if k[0] == "K" and k[1] in {str(norm(v)).strip() for v in new_key.values()}:
                        extra_mods[sname].append((k, {tr.key_col: r["vals"][tr.key_col]}))
                totals["trunk_renamed"] += len(new_key)
        plan, lines = {}, []
        for sname, br in sheets_br.items():
            tr = sheets_tr.get(sname)
            if tr is None:
                if ctx.scope and not sheet_matches(ctx.scope_include, xlsx, sname) and not any(row_in_scope(k, r["vals"], ctx.scope) for k, r in br.rows.items()):
                    lines.append(f"  + sheet `{sname}` new in branch, out of scope -> not copied"); continue
                plan[sname] = {"ops": None, "br": br, "tr": None}; lines.append(f"  + sheet `{sname}` new in branch -> copy whole sheet"); totals["new_sheets"] += 1; continue
            prefer = "branch" if sheet_matches(ctx.prefer_branch, xlsx, sname) else "trunk" if sheet_matches(ctx.prefer_trunk, xlsx, sname) else None
            ops = diff_sheet(sheets_base.get(sname), br, tr, a.add_missing, a.on_collision, prefer)
            ops["mod"] = extra_mods.get(sname, []) + ops["mod"]
            totals["prefer_branch_cells"] += ops["prefer"]["branch"]; totals["prefer_trunk_cells"] += ops["prefer"]["trunk"]
            if not a.apply_deletes and ops["del"]:
                lines.append(f"  - `{sname}`: {len(ops['del'])} rows deleted in branch, NOT deleted in trunk (use --apply-deletes): {[key_str(k) for k in ops['del']][:8]}")
                ops["del"] = []
            if ctx.switch_cols: totals["switched_off"] += switch_off(br, ops, ctx.switch_cols, ctx.scope)
            tr_alias_key = None
            for kind, k, why in ops["conflict"]:
                R["conflicts"].append((name, sname, kind, key_str(k), why))
                if kind == "add-collision":
                    if tr_alias_key is None:
                        tr_alias_key = {}
                        for kk, r in tr.rows.items():
                            al = r["vals"].get("Alias")
                            if kk[0] == "K" and al is not None: tr_alias_key.setdefault(str(al).strip(), kk[1])
                    al = br.rows[k]["vals"].get("Alias")
                    hit = tr_alias_key.get(str(al).strip()) if al is not None else None
                    key_ref = desc.refs.get((xlsx, sname, br.key_col)) if br.key_col else None
                    if key_ref and hit is None:
                        tgt = ", ".join(f"{x}!{sn}" for x, sn in sorted(key_ref))
                        R["suggest"].append([name, sname, key_str(k), "", f"key column references {tgt}: follows that sheet's remap (leave new_id blank; otherwise --on-collision)"])
                        totals["collision_key_is_ref"] += 1
                    elif hit is not None and hit == k[1]:
                        R["suggest"].append([name, sname, key_str(k), "", f"same row on both sides (Alias {al!r}), content differs: leave new_id blank, use --on-collision / --prefer"])
                        totals["collision_same_row"] += 1
                    elif hit is not None:
                        R["suggest"].append([name, sname, key_str(k), hit, f"trunk row {hit} has the same Alias {al!r}"])
                        totals["collision_alias_prefilled"] += 1
                    else:
                        R["suggest"].append([name, sname, key_str(k), "", "no Alias match in trunk: choose a free ID"])
                        R["collide"].append((xlsx, sname, k[1], br, tr)); totals["collision_needs_id"] += 1
            for k, al in ops["alias_dup"]:
                R["alias_warn"].append(f"{name}!{sname} [{key_str(k)}] Alias {al!r} already used in trunk")
                R["conflicts"].append((name, sname, "alias-dup", key_str(k), f"Alias {al!r} already used by another trunk row (row still added)"))
            if ops["add"] or ops["mod"] or ops["del"] or ops["new_cols"]:
                plan[sname] = {"ops": ops, "br": br, "tr": tr}
                cidx = {h: i for i, h in enumerate(br.headers)}
                written = [(br.rows[k]["row"], cidx[h], br.rows[k]["vals"].get(h)) for k, _ in ops["add"] for h in br.named] + \
                          [(br.rows[k]["row"], cidx[h], v) for k, cells in ops["mod"] if k in br.rows for h, v in cells.items() if h in cidx]
                fcells = [(r, c, v) for r, c, v in written if (r, c) in br.formulas]
                nocache = [(r, c) for r, c, v in fcells if v is None]
                totals["formulas_as_values"] += len(fcells)
                pref = f", {ops['prefer']['branch']} cells branch-preferred" if ops["prefer"]["branch"] else (f", {ops['prefer']['trunk']} cells trunk-preferred" if ops["prefer"]["trunk"] else "")
                lines.append(f"  ~ `{sname}`: +{len(ops['add'])} add, ~{len(ops['mod'])} modify, -{len(ops['del'])} delete, +{len(ops['new_cols'])} new cols{' ' + str(ops['new_cols']) if ops['new_cols'] else ''}, {ops['skip']} already in trunk, {len(ops['conflict'])} conflicts" + pref + (f", {len(fcells)} formula cells written as values" if fcells else ""))
                if nocache: lines.append(f"  ! `{sname}`: {len(nocache)} branch formula cells have NO cached value (workbook not saved by Excel?) - they will be written empty: rows {sorted({r for r, c in nocache})[:8]}")
                for kk in ("add", "mod", "del", "new_cols"): totals[kk] += len(ops[kk])
            elif ops["conflict"] or ops["skip"]:
                lines.append(f"  = `{sname}`: nothing to apply ({ops['skip']} already in trunk, {len(ops['conflict'])} conflicts)")
            totals["conflicts"] += len(ops["conflict"])
        for sname in sheets_tr:
            if sname not in sheets_br and sheets_base.get(sname) is not None: lines.append(f"  - sheet `{sname}` deleted in branch; kept in trunk")
        R["lines"].append(f"## {name}" + ("" if lines else ": no row-level differences")); R["lines"].extend(lines)
        if plan: R["plans"][rel] = plan
    # --yield-trunk: every trunk row anywhere that points at a renamed trunk row must follow it
    for tgt, old, new in renamed_refs:
        for (x2, s2, col) in desc.referrers.get(tgt, ()):
            if x2 not in R["loaded_tr"]:
                p2 = os.path.join(ctx.trunk_ws, "design", x2 + ".xlsx")
                R["loaded_tr"][x2] = load_workbook_sheets(p2, desc, x2) if os.path.isfile(p2) else {}
            tr2 = R["loaded_tr"][x2].get(s2)
            if not tr2 or col not in tr2.headers: continue
            rel2 = f"design/{x2}.xlsx"
            for k, r in tr2.rows.items():
                nv, hit = _replace_tokens(r["vals"].get(col), {old: new})
                if not hit: continue
                r["vals"][col] = nv
                entry = R["plans"].setdefault(rel2, {}).setdefault(s2, {"ops": empty_ops(), "br": None, "tr": tr2})
                if entry["ops"] is None: continue
                entry["ops"]["mod"].append((k, {col: nv})); totals["mod"] += 1
                R["remap_log"].append(f"trunk yields: {x2}!{s2}[{key_str(k)}].{col} {old} -> {new}")
    return R

def _parent_rule(xlsx, sname, olds, br, desc, M, siblings=None):
    """Keys derived from another sheet: (1) a Ref column holds old_key // M for every clashing row, or
    (2) old_key // M are themselves keys of a sibling sheet in the same workbook (guide 115901 -> guide_section 1159,
    no Ref declared). Returns (parent (xlsx, sheet), how) or None."""
    if not M: return None
    num = [o for o in olds if re.fullmatch(r"-?\d+", o)]
    if not num: return None
    for col in br.named:
        tgts = desc.refs.get((xlsx, sname, col))
        if not tgts or col == br.key_col: continue
        ok = True
        for o in num:
            v = br.rows.get(("K", o, 1), {}).get("vals", {}).get(col)
            if v is None or not re.fullmatch(r"-?\d+", str(norm(v))) or int(norm(v)) != int(o) // M: ok = False; break
        if ok: return next(iter(tgts)), col
    prefixes = {str(int(o) // M) for o in num}
    for psn, psd in (siblings or {}).items():
        if psn == sname: continue
        pkeys = {k[1] for k in psd.rows if k[0] == "K"}
        if prefixes and len(prefixes & pkeys) >= max(1, int(0.9 * len(prefixes))): return (xlsx, psn), f"key // {M}"
    return None

def auto_remap_from(R, ctx, remap, taken):
    """Turn 'needs a free ID' collisions of pass 1 into a remap table (branch rows renamed) - or, for
    --yield-trunk sheets, into trunk renames. Returns (remap2, trunk_renames, notes)."""
    remap2 = defaultdict(dict, {k: dict(v) for k, v in remap.items()})
    trunk_renames = defaultdict(dict); notes = []
    by_sheet = defaultdict(list)
    for xlsx, sname, old, br, tr in R["collide"]: by_sheet[(xlsx, sname)].append((old, br, tr))
    def parent_of(item):   # parents (guide_section) before children (guide) so derived ids can follow them
        (xlsx, sname), items = item
        olds = [o for o, _, _ in items]; br = items[0][1]
        ex_ = {int(k[1]) for k in items[0][2].rows if k[0] == "K" and re.fullmatch(r"-?\d+", k[1])}
        return _parent_rule(xlsx, sname, olds, br, ctx.desc, _group_modulus(sorted(ex_)), R["loaded_br"].get(xlsx))
    for (xlsx, sname), items in sorted(by_sheet.items(), key=lambda it: 1 if parent_of(it) else 0):
        br, tr = items[0][1], items[0][2]
        olds = [o for o, _, _ in items]
        existing = {int(k[1]) for k in tr.rows if k[0] == "K" and re.fullmatch(r"-?\d+", k[1])}
        existing |= {int(k[1]) for k in br.rows if k[0] == "K" and re.fullmatch(r"-?\d+", k[1]) and k[1] not in olds}
        existing_str = {k[1] for k in tr.rows if k[0] == "K" and not re.fullmatch(r"-?\d+", k[1])} | {k[1] for k in br.rows if k[0] == "K" and not re.fullmatch(r"-?\d+", k[1]) and k[1] not in olds}
        M = _group_modulus(sorted(existing)); parent = _parent_rule(xlsx, sname, olds, br, ctx.desc, M, R["loaded_br"].get(xlsx))
        if parent and not sheet_matches(ctx.yield_trunk, xlsx, sname):
            ptable = remap2.get(parent[0], {})
            sug = {o: str(int(ptable[str(int(o) // M)]) * M + int(o) % M) for o in olds if re.fullmatch(r"-?\d+", o) and str(int(o) // M) in ptable}
            if sug:
                rule = f"follows {parent[0][0]}!{parent[0][1]} via {parent[1]} (x{M} + step)"
                taken.update(int(v) for v in sug.values())
                for o, n in sug.items(): remap2[(xlsx, sname)][o] = n
                notes.append(f"{xlsx}!{sname}: {len(sug)} branch IDs -> {rule}")
                for row in R["suggest"]:
                    if row[0] == xlsx + ".xlsx" and row[1] == sname and row[2] in sug and not row[3]: row[3] = sug[row[2]]; row[4] = f"suggested ({rule})"
                olds = [o for o in olds if o not in sug]
                if not olds: continue
        if sheet_matches(ctx.yield_trunk, xlsx, sname):
            sug, rule = suggest_ids(olds, existing, taken, existing_str)
            for o, n in sug.items(): trunk_renames[(xlsx, sname)][o] = n
            if sug: notes.append(f"{xlsx}!{sname}: trunk yields, its rows {list(sug)[:5]}{'...' if len(sug) > 5 else ''} renamed ({rule})")
        else:
            sug, rule = suggest_ids(olds, existing, taken, existing_str)
            for o, n in sug.items(): remap2[(xlsx, sname)][o] = n
            if sug: notes.append(f"{xlsx}!{sname}: {len(sug)} branch IDs -> {rule}")
        # prefill remap_suggest.csv rows with the suggestion so the user sees the rule even without --auto-ids
        for row in R["suggest"]:
            if row[0] == xlsx + ".xlsx" and row[1] == sname and row[2] in sug and not row[3]:
                row[3] = sug[row[2]] if not sheet_matches(ctx.yield_trunk, xlsx, sname) else ""
                row[4] = (f"suggested ({rule})" if not sheet_matches(ctx.yield_trunk, xlsx, sname) else f"trunk yields: trunk row renamed to {sug[row[2]]}")
    for row in R["suggest"]:
        if row[3] and row[4].startswith("trunk row"): remap2[(row[0][:-5], row[1])].setdefault(row[2], row[3])   # alias-prefilled: apply too
    return remap2, trunk_renames, notes

# ----------------------------------------------------------------------------- main
def main():
    global DATA_DIR
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--branch-ws", default=DEF_BRANCH); ap.add_argument("--trunk-ws", default=DEF_TRUNK)
    ap.add_argument("--base", type=int, help="branch revision to diff from (default: state.json, else copy point)")
    ap.add_argument("--files", help="comma list of xlsx basenames to restrict to")
    ap.add_argument("--remap", help="CSV: xlsx,sheet,old_id,new_id")
    ap.add_argument("--apply", action="store_true", help="write changes (default is dry-run report)")
    ap.add_argument("--apply-deletes", action="store_true", help="also delete rows the branch deleted (default: report only)")
    ap.add_argument("--add-missing", action="store_true", help="rows the branch modified but trunk lacks: add them instead of conflicting")
    ap.add_argument("--on-collision", choices=["conflict", "skip", "overwrite"], default="conflict", help="branch-added key already in trunk with different content: report / keep trunk / write branch")
    ap.add_argument("--prefer-branch", default="", help="sheets (xlsx[!sheet], comma list, or *) where both-sides-changed cells and same-key collisions take the BRANCH value")
    ap.add_argument("--prefer-trunk", default="", help="sheets where they keep the TRUNK value (no conflict reported)")
    ap.add_argument("--yield-trunk", default="", help="sheets where a genuinely new branch row that clashes keeps its ID and TRUNK's row (plus every trunk reference to it) is renamed")
    ap.add_argument("--auto-ids", action="store_true", help="apply the tool's next-free-ID suggestions for clashing branch rows (else they only go to remap_suggest.csv)")
    ap.add_argument("--scope", default="", help="comma list of event tokens (e.g. sim2609,ChocolateFactory): only rows, new files and layout entries mentioning one of them are merged")
    ap.add_argument("--scope-include", default="", help="sheets (xlsx[!sheet]) merged whole regardless of --scope, e.g. pve_layout.xlsx,robot.xlsx")
    ap.add_argument("--switch-off", nargs="?", const=SWITCH_COLS, default="", help=f"set these toggle columns to FALSE on every added row (default list: {SWITCH_COLS})")
    ap.add_argument("--unlock-local", action="store_true", help="chmod +w read-only trunk xlsx (svn:needs-lock) before writing")
    ap.add_argument("--no-text", action="store_true", help="skip descriptor/PostConvert text-file merging")
    ap.add_argument("--no-verify", action="store_true"); ap.add_argument("--offline", action="store_true", help="skip HEAD check")
    ap.add_argument("--out", help="report dir (default: <data-dir>/out/<timestamp>)")
    ap.add_argument("--data-dir", help=f"root for cache/, state.json and out/ (default {DATA_DIR})")
    a = ap.parse_args()
    if a.data_dir: DATA_DIR = os.path.abspath(a.data_dir)

    ctx = Ctx(); ctx.a = a
    ctx.branch_ws, ctx.trunk_ws = os.path.abspath(a.branch_ws), os.path.abspath(a.trunk_ws)
    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    out_dir = a.out or os.path.join(DATA_DIR, "out", ts); os.makedirs(out_dir, exist_ok=True)
    rep = [f"# xlsx merge report {ts}", f"branch: {ctx.branch_ws}", f"trunk:  {ctx.trunk_ws}", ""]
    def log(s=""):
        print(s, flush=True); rep.append(s)   # flushed: the desktop app streams these lines live

    ctx.branch_url = svn_item(ctx.branch_ws, "url"); branch_rev = int(svn_item(ctx.branch_ws, "revision"))
    state = load_state()
    try: trunk_url = svn_item(ctx.trunk_ws, "url")
    except Exception: trunk_url = ctx.trunk_ws   # unversioned target (e.g. a test copy)
    skey = f"{ctx.branch_url} -> {trunk_url}"
    ctx.base_rev = a.base or state.get(skey, {}).get("merged_branch_rev") or copy_point(ctx.branch_url)
    if not ctx.base_rev: raise SystemExit("cannot determine base revision; pass --base")
    ctx.prefer_branch, ctx.prefer_trunk, ctx.yield_trunk = parse_sheet_specs(a.prefer_branch), parse_sheet_specs(a.prefer_trunk), parse_sheet_specs(a.yield_trunk)
    ctx.scope = [t.strip().lower() for t in a.scope.split(",") if t.strip()]
    ctx.scope_include = parse_sheet_specs(a.scope_include)
    ctx.switch_cols = [c.strip() for c in a.switch_off.split(",") if c.strip()] if a.switch_off else []
    ctx.only = set(a.files.split(",")) if a.files else None
    log(f"base rev: {ctx.base_rev}   branch WC rev: {branch_rev}   mode: {'APPLY' if a.apply else 'DRY-RUN'}")
    if ctx.scope:
        ctx.scope = expand_scope(ctx.scope, ctx.branch_ws)
        log(f"scope: {ctx.scope} (whole: {a.scope_include or '-'})")
    if ctx.switch_cols and not ctx.scope: log("WARNING --switch-off needs --scope to know which rows are the release event; nothing will be switched off")
    if a.prefer_branch or a.prefer_trunk or a.yield_trunk: log(f"prefer branch: {a.prefer_branch or '-'} | prefer trunk: {a.prefer_trunk or '-'} | trunk yields: {a.yield_trunk or '-'}")
    st = run(["svn", "status", ctx.branch_ws]).strip()
    if st: log(f"WARNING branch WC has local modifications (merging working files as they are):\n{st}")
    if not a.offline and trunk_url != ctx.trunk_ws:
        try:
            head_lc = int(run(["svn", "info", "--show-item", "last-changed-revision", "-r", "HEAD", ctx.trunk_ws]).strip())
            if head_lc > int(svn_item(ctx.trunk_ws, "revision")): log(f"WARNING trunk WC is behind HEAD (HEAD last change r{head_lc}); run `svn up` first")
        except Exception as e: log(f"WARNING could not check trunk HEAD: {e}")

    ctx.desc = Descriptors(ctx.branch_ws, ctx.trunk_ws)
    for b in ctx.desc.bad: log(f"note: descriptor unparsable, skipped: {b}")
    remap = load_remap(a.remap) if a.remap else {}
    changes = branch_changes(ctx.branch_ws, ctx.base_rev)
    ctx.xlsx_changes = [(s, r) for s, r in changes if r.startswith("design/") and r.lower().endswith(".xlsx")]
    xset = set(ctx.xlsx_changes)
    text_changes = [(s, r) for s, r in changes if (s, r) not in xset]
    log(f"branch changed since r{ctx.base_rev}: {len(ctx.xlsx_changes)} xlsx, {len(text_changes)} other files"); log()

    # pass 1: find the collisions; pass 2 (only when asked) re-plans with suggested IDs / trunk renames applied
    R = plan_all(ctx, remap, {})
    remap2, trunk_renames, notes = auto_remap_from(R, ctx, remap, set())
    if (a.auto_ids or ctx.yield_trunk) and (notes or trunk_renames):
        if not a.auto_ids:   # only the yield part is wanted
            remap2 = defaultdict(dict, {k: dict(v) for k, v in remap.items()})
            for row in R["suggest"]:
                if row[3] and row[4].startswith("trunk row"): remap2[(row[0][:-5], row[1])].setdefault(row[2], row[3])
        R = plan_all(ctx, remap2, trunk_renames)
        log("## automatic ID resolution"); [log("  " + n) for n in notes]; log()
    for l in R["lines"]: log(l)
    totals, conflicts, suggest, new_files, plans = R["totals"], R["conflicts"], R["suggest"], R["new_files"], R["plans"]
    log(); log(f"TOTAL: +{totals['add']} rows, ~{totals['mod']} rows, -{totals['del']} rows, +{totals['new_cols']} cols, {totals['new_sheets']} new sheets, {len(new_files)} new files, {totals['conflicts']} conflicts, {totals['formulas_as_values']} formula cells written as values")
    extras = [f"{totals['prefer_branch_cells']} cells taken from branch by --prefer-branch" if totals["prefer_branch_cells"] else "",
              f"{totals['prefer_trunk_cells']} cells kept from trunk by --prefer-trunk" if totals["prefer_trunk_cells"] else "",
              f"{totals['trunk_renamed']} trunk rows renamed by --yield-trunk" if totals["trunk_renamed"] else "",
              f"{totals['out_of_scope_rows']} branch rows and {totals['out_of_scope_files']} files left out by --scope" if ctx.scope else "",
              f"{totals['switched_off']} toggle cells set FALSE on added rows" if totals["switched_off"] else ""]
    for e in extras:
        if e: log("  " + e)
    if R["alias_warn"]:
        log(); log(f"{len(R['alias_warn'])} added rows reuse an Alias already in trunk (c_data may reject; full list in conflicts.csv):")
        for w in R["alias_warn"][:10]: log("  " + w)
    if R["remap_log"]:
        log(); log(f"Remap applied ({len(R['remap_log'])} changes; first 40):")
        for l in R["remap_log"][:40]: log("  " + l)

    text_results = []
    if not a.no_text:
        log(); log("## non-xlsx files (descriptors, PostConvert, ...)")
        for status, rel in text_changes:
            br_path, tr_path = os.path.join(ctx.branch_ws, rel), os.path.join(ctx.trunk_ws, rel)
            if status == "D": res = "deleted in branch -> " + ("trunk still has it (not deleted)" if os.path.isfile(tr_path) else "already absent")
            elif ctx.scope and not os.path.isfile(tr_path) and not any(t in os.path.basename(rel).lower() for t in ctx.scope):
                res = "new file, out of scope -> not copied"
            else:
                base_path = fetch_base(ctx.branch_url, ctx.base_rev, rel) if status == "M" else None
                res = merge_text(rel, base_path, br_path, tr_path, out_dir, a.apply, ctx.scope)
                if res == "copy-new": new_files.append(rel)
            log(f"  {status} {rel}: {res}"); text_results.append([status, rel, res])

    lint = lint_merged(plans, ctx.desc, ctx.trunk_ws)
    log(); log(f"## lint of the merged result (what c_data would reject): {len(lint)} findings" + ("" if lint else " - clean"))
    for kind, n in sorted(Counter(f[2] for f in lint).items()): log(f"  {n} x {kind}")
    for f in lint[:15]: log(f"  {f[2]} {f[0]}!{f[1]} [{f[3]}]: {f[4]}")
    if len(lint) > 15: log(f"  ... {len(lint) - 15} more in conflicts.csv")
    conflicts.extend(lint)
    totals["lint"] = len(lint)
    if lint and a.apply: log("WARNING: applying anyway - fix the lint findings (remap / --prefer / --on-collision) before running coder_convert")

    # machine-readable twin of the report, read by the Event Automator "Merge to Trunk" tab
    summary = {"mode": "apply" if a.apply else "dry-run", "base_rev": ctx.base_rev, "branch_rev": branch_rev,
               "totals": dict(totals), "n_conflicts": len(conflicts), "conflicts_by_kind": dict(Counter(c[2] for c in conflicts)),
               "n_collisions": len(suggest), "n_lint": len(lint), "new_files": new_files, "text_results": text_results,
               "auto_id_notes": notes, "applied": False, "verify_problems": None, "state_updated": False,
               "out_dir": out_dir, "report_path": os.path.join(out_dir, "report.md")}
    if conflicts:
        with open(os.path.join(out_dir, "conflicts.csv"), "w", encoding="utf-8-sig", newline="") as f:
            w = csv.writer(f); w.writerow(["xlsx", "sheet", "kind", "key", "detail"]); w.writerows(conflicts)
    if suggest:
        with open(os.path.join(out_dir, "remap_suggest.csv"), "w", encoding="utf-8-sig", newline="") as f:
            w = csv.writer(f); w.writerow(["xlsx", "sheet", "old_id", "new_id", "note"]); w.writerows(suggest)
        log(); log(f"{len(suggest)} ID collisions -> {out_dir}\\remap_suggest.csv: {totals['collision_alias_prefilled']} prefilled with trunk's ID (same Alias already in trunk), "
                   f"{totals['collision_same_row']} are the same row with different content (--prefer / --on-collision), {totals['collision_key_is_ref']} have a foreign-key key column (follow the referenced sheet), "
                   f"{totals['collision_needs_id']} need a free ID" + (" (applied automatically: --auto-ids)" if a.auto_ids else " (suggested in new_id; rerun with --remap or --auto-ids)"))

    if a.apply:
        targets = [os.path.join(ctx.trunk_ws, rel) for rel in plans]
        ro = [p for p in targets if not os.access(p, os.W_OK)]
        if ro and a.unlock_local:
            for p in ro: os.chmod(p, 0o666)
        elif ro:
            log(); log("ABORT: read-only trunk files (svn:needs-lock). `svn lock` them, or rerun with --unlock-local:")
            for p in ro: log("  " + p)
            summary["error"] = "read-only trunk files"; write_report(out_dir, rep, summary); sys.exit(2)
        log(); log("## applying")
        xl = Excel(); problems_all = []
        try:
            for rel, plan in plans.items():
                log(f"{rel}"); apply_workbook(xl, os.path.join(ctx.trunk_ws, rel), os.path.join(ctx.branch_ws, rel), plan, rep)
                if not a.no_verify:
                    probs = verify_workbook(os.path.join(ctx.trunk_ws, rel), plan, ctx.desc, os.path.basename(rel)[:-5])
                    log(f"  verify: {'OK' if not probs else str(len(probs)) + ' PROBLEMS'}")
                    for p in probs: log("    ! " + p)
                    problems_all += probs
        finally:
            xl.quit()
        summary["applied"] = True; summary["verify_problems"] = problems_all
        for rel in [r for r in new_files if r.lower().endswith(".xlsx")]:
            dst = os.path.join(ctx.trunk_ws, rel); os.makedirs(os.path.dirname(dst), exist_ok=True); shutil.copy2(os.path.join(ctx.branch_ws, rel), dst)
            if not os.access(dst, os.W_OK): os.chmod(dst, 0o666)
        if new_files:
            with open(os.path.join(out_dir, "svn_after.txt"), "w", encoding="utf-8") as f:
                for rel in new_files:
                    f.write(f'svn add "{os.path.join(ctx.trunk_ws, rel)}"\n')
                    if rel.lower().endswith(".xlsx"): f.write(f'svn propset svn:needs-lock "*" "{os.path.join(ctx.trunk_ws, rel)}"\n')
            log(f"new files copied; run the commands in {out_dir}\\svn_after.txt")
        if not problems_all:
            state[skey] = {"base_rev": ctx.base_rev, "merged_branch_rev": branch_rev, "date": ts, "conflicts_left": len(conflicts)}
            os.makedirs(DATA_DIR, exist_ok=True)
            json.dump(state, open(state_path(), "w", encoding="utf-8"), indent=1)
            summary["state_updated"] = True
            log(f"state.json updated: next merge will use base r{branch_rev}")
        else:
            log(f"VERIFY FAILED ({len(problems_all)} problems) -- state.json NOT updated; review trunk with svn diff / revert")
    write_report(out_dir, rep, summary)
    print(f"\nreport: {os.path.join(out_dir, 'report.md')}", flush=True)

def write_report(out_dir, rep, summary=None):
    open(os.path.join(out_dir, "report.md"), "w", encoding="utf-8").write("\n".join(rep) + "\n")
    if summary is not None:
        json.dump(summary, open(os.path.join(out_dir, "summary.json"), "w", encoding="utf-8"), indent=1, ensure_ascii=False)

if __name__ == "__main__":
    main()
