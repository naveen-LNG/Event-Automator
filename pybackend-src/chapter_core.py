"""
Chapter creation automation — the verified core, shared by two front ends:

  * tools/EventAutomationDesktop      (Chapter Update tab, via chapter_update.py)
  * tools/ChapterCreator              (standalone .exe, via chapter_creator.py)

Only the steps that are pure openpyxl / file-copy live here, so the Electron
app's bundled python.org interpreter (openpyxl only — no pandas/numpy) can run
every one of them. The Unity / SVN / simulation steps (1, 2, 7, 8, 9) stay in
the standalone tool because they shell out to project scripts that need a full
system Python, Unity, or the RunCardEx2 binary.

DO NOT fork this logic. It was verified column-by-column against real
production workbooks (chest / item / store / notification / asset_ref / icon /
pve_chapter); a second copy means the next fix has to land twice.

Chapter numbers appear in three forms and every one of them has to be replaced:
  * internal map id   88000 -> 89000        (old_ch * 1000)
  * Chinese numerals  七十八 -> 七十九        (Alias / PIC columns)
  * display number    78 -> 79              (standalone digits only)
"""
import ast
import collections
import glob
import json
import os
import re
import shutil
import tempfile
import subprocess
import sys
from datetime import datetime as _datetime, timedelta as _timedelta
from pathlib import Path

import openpyxl
from openpyxl.styles import PatternFill, Font
from openpyxl.formatting.rule import CellIsRule


class Cfg:
    """The subset of the desktop app's AppConfig these steps actually read.

    branch_root is a checkout root (…/trunk, …/Misc), i.e. the folder holding
    design/ and develop/ — the same value the Electron app already tracks as
    the active branch. The standalone tool's AppConfig exposes the identical
    attribute names, so both front ends hand the same shape to these steps.
    """

    def __init__(self, branch_root, workspace, old_ch, display_offset=10,
                 sim_path=None, lvl_py=None, map_builder=None,
                 sim_repeat=500, sim_p_num=8, sim_tag=""):
        self.branch_root = Path(branch_root)
        self.workspace = workspace
        self.old_ch = int(old_ch)
        self.display_offset = int(display_offset)
        self._sim_path = Path(sim_path) if sim_path else None
        # tools/ lives OUTSIDE the branch checkout and its position relative to
        # a branch differs (trunk/ is one level down, branches/Misc/ is two), so
        # these cannot be derived reliably — they are configured, with the
        # tools-next-to-the-repo layout as the fallback guess.
        self._lvl_py = Path(lvl_py) if lvl_py else None
        self._map_builder = Path(map_builder) if map_builder else None
        self.sim_repeat = sim_repeat
        self.sim_p_num = sim_p_num
        self.sim_tag = sim_tag

    @property
    def _tools_guess(self):
        """<repo>/tools, found by walking up from the branch checkout.

        trunk/ sits one level below the repo root but branches/<name>/ sits two,
        so a fixed offset is wrong for half the branches. Walk up until a real
        tools/ turns up; if none does, the paths are simply unknown and the step
        that needs one says so with the path it tried.
        """
        here = self.branch_root
        for _ in range(4):
            here = here.parent
            if (here / "tools").is_dir():
                return here / "tools"
            if here == here.parent:      # hit the drive root
                break
        return self.branch_root.parent / "tools"

    @property
    def lvl_py(self):
        return self._lvl_py or (self._tools_guess
                                / "Level_Builder3" / "Magic_Builder" / "Assets" / "Py")

    @property
    def map_builder(self):
        return self._map_builder or (self._tools_guess / "Map_Builder" / "Magic_Builder")

    @property
    def map_scenes(self):
        return self.map_builder / "Assets" / "Scenes"

    @property
    def map_py(self):
        return self.map_builder / "Assets" / "Py"

    @property
    def map_output(self):
        return self.map_builder / "__Output"

    def design_root(self):
        return self.branch_root / "design" / "DesignData" / self.workspace

    @property
    def new_ch(self):
        return self.old_ch + 1

    @property
    def disp_old(self):
        return self.old_ch - self.display_offset

    @property
    def disp_new(self):
        return self.new_ch - self.display_offset

    def design(self):
        return self.branch_root / "design" / "DesignData" / self.workspace / "design"

    @property
    def client_map(self):
        return self.branch_root / "develop" / "client" / "Skipbo" / "Assets" / "Map"

    @property
    def client_tex(self):
        return (self.branch_root / "develop" / "client" / "Skipbo"
                / "Assets" / "ResourcesAb" / "Texture")

    @property
    def sim_path(self):
        # Simulation output lives outside the branch checkout (tools/ is a
        # sibling of trunk/), so it is passed in rather than derived.
        return self._sim_path or (self.branch_root / "develop" / "tools" / "RunCardEx2")


# The step functions below are annotated `cfg: AppConfig` (they were lifted
# verbatim out of the standalone tool). Python <3.14 evaluates annotations
# eagerly, so the name has to resolve here — and Cfg is exactly the shape they
# read. ponytail: alias, not a second class.
AppConfig = Cfg


def _python_exe():
    """A real Python interpreter for the project's own .py scripts.

    Step 1's conf_fill_ex.py needs pandas/numpy, which the bundled
    openpyxl-only interpreter does not carry, so that step has to borrow a
    system Python. Returns None when there is none to borrow.
    """
    if not getattr(sys, "frozen", False) and "pybackend" not in sys.executable.lower():
        return sys.executable
    for cand in ("python", "python3", "py"):
        found = shutil.which(cand)
        if found:
            return found
    for pat in (r"C:\Python3*\python.exe",
                r"C:\Program Files\Python3*\python.exe",
                os.path.expanduser(r"~\AppData\Local\Programs\Python\Python3*\python.exe")):
        hits = sorted(glob.glob(pat), reverse=True)
        if hits:
            return hits[0]
    return None


_NO_PYTHON_MSG = (
    "No system Python found.\n\n"
    "This step runs a SkipBo project script that needs pandas/numpy, which the "
    "app's bundled interpreter does not carry. Install Python 3.9+ from "
    "python.org (tick 'Add Python to PATH'), then retry."
)


_MISSING_MOD_RE = re.compile(r"ModuleNotFoundError: No module named ['\"]([^'\"]+)['\"]")


# conf_fill_ex.py hardcodes the design folder it operates on, resolved against
# the CWD, and offers no way to override it. Read that literal out of the
# script so this tool checks the same workbooks the script will, instead of
# duplicating a path that could drift.
_DESIGN_PATH_RE = re.compile(
    r"""design_path\s*=\s*os\.path\.abspath\(\s*['"]([^'"]+)['"]""")


def _same_dir(a, b):
    """Whether two paths name the same folder, short/long form aside."""
    try:
        return Path(a).resolve() == Path(b).resolve()
    except Exception:
        strip = chr(92) + "/"
        return str(a).rstrip(strip).lower() == str(b).rstrip(strip).lower()


def _script_design_path(script_path, cwd):
    """The design folder a Level Builder script will actually read and write."""
    try:
        src = Path(script_path).read_text("utf-8", errors="replace")
    except Exception:
        return None
    m = _DESIGN_PATH_RE.search(src)
    if not m:
        return None
    # resolve() so this is comparable to cfg.design(): a path can arrive in
    # 8.3 short form (NAVEEN~1) while the other side is long form, which would
    # otherwise read as a workspace mismatch that is not real.
    try:
        return Path(os.path.join(str(cwd), m.group(1))).resolve()
    except Exception:
        return Path(os.path.abspath(os.path.join(str(cwd), m.group(1))))


# Errors these project scripts raise that mean something specific and
# actionable. Their own messages are accurate but assume you know the tool's
# internals, so they get translated rather than passed through raw.
_REFILL_CONFLICT_RE = re.compile(r"refill level id conflict:\s*(\d+)")


def _explain_script_error(script_name, stderr, script_path=None, cfg=None):
    """Turn a known project-script error into next steps, or None."""
    m = _REFILL_CONFLICT_RE.search(stderr or "")
    if m:
        lvl = int(m.group(1))
        internal = lvl // 1000
        # Level ids are built from the INTERNAL chapter number; the number
        # everyone else says out loud is the display one, so give both rather
        # than quoting a chapter the user has never heard of.
        disp = internal - cfg.display_offset if cfg is not None else None
        which = (f"chapter {disp} (internal {internal})" if disp is not None
                 else f"internal chapter {internal}")
        first, last = internal * 1000 + 1, internal * 1000 + 200
        where = f"\n  {Path(script_path).parent}" if script_path else ""
        return (
            f"Magic Builder stopped: level {lvl} already exists in the Level "
            f"Builder's own source data, so it will not overwrite it.\n\n"
            f"That means {which} - levels {first}-{last} - has already been "
            f"built there, at least partly. Event Automator's own pre-check "
            f"looks at pve_level.xlsx / pve_layout.xlsx in the design folder "
            f"and found nothing, so the leftover rows are on the Level Builder "
            f"side:{where}\n\n"
            f"Usually one of two things:\n"
            f"  1. An earlier Chapter Update run got part-way through. Revert "
            f"the Level Builder working copy (svn revert) and re-run.\n"
            f"  2. {which} was genuinely already built - by someone else, or "
            f"already committed. Check before deleting anything; if it is "
            f"done, this chapter needs no run.\n\n"
            f"Nothing was written to the design folder by this step."
        )
    return None

def _script_failure(script_name, result, py_exe, script_path=None, cfg=None):
    """A RuntimeError that says what to actually do about it.

    These are SkipBo Level Builder / Map Editor scripts, not ours. When the
    interpreter running one is missing a dependency, the child's traceback is
    the only clue and 'exited with code 1' buries it - so name the package and
    the exact interpreter to install it into. PATH can resolve to a different
    Python than the user assumes, which is why py_exe is spelled out.
    """
    stderr = result.stderr or ""
    known = _explain_script_error(script_name, stderr, script_path, cfg)
    if known:
        return RuntimeError(known)
    m = _MISSING_MOD_RE.search(stderr)
    if not m:
        return RuntimeError(f"{script_name} exited with code {result.returncode}")
    mod = m.group(1)

    # A module sitting next to the script is one of ITS OWN files, not
    # something to pip install - telling the user otherwise sends them off to
    # install an unrelated package off PyPI, or a nonexistent one.
    if script_path:
        d = Path(script_path).parent
        if (d / f"{mod}.py").exists() or (d / mod).is_dir():
            return RuntimeError(
                f"{script_name} could not import '{mod}', which is one of its own "
                f"files in:\n  {d}\n\n"
                f"That is a path problem rather than a missing package - the "
                f"script's own folder was not on sys.path. Report this as an "
                f"Event Automator bug; it is not something to fix on your "
                f"machine.\n\nInterpreter used: {py_exe}"
            )
    return RuntimeError(
        f"{script_name} needs the '{mod}' package, which is missing from the "
        f"Python it was run with.\n\n"
        f"Interpreter used: {py_exe}\n\n"
        f"Install it there, then re-run this step:\n"
        f'  "{py_exe}" -m pip install {mod}\n\n'
        f"({script_name} is a SkipBo project script, not part of Event "
        f"Automator. The app runs it on its own bundled interpreter when that "
        f"can satisfy the script, and borrows a system Python otherwise.)"
    )

def _bundled_python():
    """The interpreter this backend is running on (the app's bundled one)."""
    return sys.executable


def _script_imports(script_path):
    """Third-party module names a project script needs, by reading its source.

    The Level Builder / Map Editor scripts are not ours and their dependencies
    change without notice, so the dependency list is discovered from the actual
    file on this machine rather than hardcoded here. Stdlib and modules sitting
    next to the script resolve on any interpreter, so both are dropped.

    Returns None when the source cannot be read or parsed - callers then fall
    back to trying interpreters rather than claiming to know what is needed.
    """
    try:
        src = Path(script_path).read_text("utf-8", errors="replace")
        tree = ast.parse(src)
    except Exception:
        return None

    roots = set()
    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            for a in node.names:
                roots.add(a.name.split(".")[0])
        elif isinstance(node, ast.ImportFrom):
            # level > 0 is an explicit relative import - always local
            if node.level == 0 and node.module:
                roots.add(node.module.split(".")[0])

    script_dir = Path(script_path).parent
    stdlib = getattr(sys, "stdlib_module_names", frozenset())
    out = set()
    for r in roots:
        if not r or r in stdlib or r == "__future__":
            continue
        if (script_dir / (r + ".py")).exists() or (script_dir / r).is_dir():
            continue  # sibling module, travels with the script
        out.add(r)
    return out


def _missing_modules(py_exe, modules):
    """Which of `modules` the given interpreter cannot import.

    Uses find_spec rather than a real import: some of these packages are slow
    to import and we only need to know they are installed. Returns all of them
    when the probe itself fails, so a broken interpreter is never treated as
    satisfying anything.
    """
    if not modules:
        return set()
    # Built line-by-line with no escape sequences: an escaped newline inside
    # this source got interpreted here rather than in the child, producing a
    # probe that failed to parse - which read as "every module is missing".
    probe = (
        "import importlib.util, sys\n"
        "for m in sys.argv[1:]:\n"
        "    try:\n"
        "        ok = importlib.util.find_spec(m) is not None\n"
        "    except Exception:\n"
        "        ok = False\n"
        "    if not ok:\n"
        "        print(m)\n"
    )
    try:
        r = subprocess.run([py_exe, "-c", probe, *sorted(modules)],
                           capture_output=True, text=True, timeout=120)
    except Exception:
        return set(modules)
    if r.returncode != 0:
        return set(modules)
    return {l.strip() for l in r.stdout.splitlines() if l.strip()}


def _resolve_interpreter(script_path, log, label=None):
    """Pick an interpreter that can actually run this project script.

    Prefers the app's own bundled interpreter, so a machine with no system
    Python still works. Falls back to a system Python when the script needs
    something the bundle does not carry - which is the honest default, because
    these scripts belong to another tool and we cannot know their full
    dependency set ahead of time.
    """
    name = label or Path(script_path).name
    needed = _script_imports(script_path)

    candidates = [("bundled", _bundled_python())]
    sys_py = _python_exe()
    if sys_py and Path(sys_py).resolve() != Path(_bundled_python()).resolve():
        candidates.append(("system", sys_py))

    shortfall = {}
    for kind, exe in candidates:
        if not exe:
            continue
        missing = _missing_modules(exe, needed) if needed is not None else set()
        if not missing:
            log(f"  {name}: using {kind} Python ({exe})")
            return exe
        shortfall[exe] = missing
        log(f"  {name}: {kind} Python lacks {', '.join(sorted(missing))} - trying next", "WARN")

    if not shortfall:
        raise RuntimeError(_NO_PYTHON_MSG)

    lines = [
        f"{name} cannot run: no available Python has everything it needs.",
        "",
        f"{name} is a SkipBo project script (Level Builder / Map Editor), not "
        "part of Event Automator, and it imports packages the app does not "
        "bundle.",
        "",
    ]
    for exe, missing in shortfall.items():
        lines.append(f"{exe}")
        lines.append(f"  missing: {', '.join(sorted(missing))}")
        lines.append(f'  fix:     "{exe}" -m pip install {" ".join(sorted(missing))}')
        lines.append("")
    lines.append("Install into whichever of those you maintain, then re-run.")
    raise RuntimeError("\n".join(lines))


# Ordinary "python script.py" puts the script's own folder on sys.path, so
# these project scripts can import their siblings (lb_helper next to
# conf_fill_ex.py, for instance). The app's bundled interpreter does NOT: its
# ._pth file makes sys.path fully explicit, PYTHONPATH is ignored, and
# sys.path[0] ends up being python314.zip. Running an external script with it
# directly therefore breaks every sibling import.
#
# So the script is launched through runpy with its folder inserted first, which
# reproduces normal invocation semantics on any interpreter, bundled or system.
_BOOTSTRAP = (
    "import sys, os, runpy\n"
    "script = sys.argv[1]\n"
    "sys.argv = sys.argv[1:]\n"
    "sys.path.insert(0, os.path.dirname(os.path.abspath(script)))\n"
    "runpy.run_path(script, run_name='__main__')\n"
)


# conf_fill_ex.py and the conf_complex_gen.py it imports both hardcode
#
#     design_path = os.path.abspath(".../trunk/design/DesignData/WorkSpcae3/design")
#
# with no argument or config to override it, so Magic Builder would always
# read and write WorkSpcae3 no matter which workspace was chosen here - while
# every other step, and the QA validation at the end, use the chosen one.
#
# Level_Builder3 is a separate SVN working copy, shared with the rest of the
# project and used standalone, so it is never modified. The scripts are also
# never relocated: they anchor everything on their own location (self_path =
# dirname(realpath(__file__)) followed by os.chdir(self_path)), so a copy
# elsewhere takes CondFilterConf.json and the __WorkSpace/__BackUp folder with
# it. Instead a small launcher patches the source in memory and execs it with
# __file__ still pointing at the original file.
_REDIRECT_LAUNCHER = '"""Run a Level Builder script with its hardcoded design_path redirected.\n\nWritten to a temp file and invoked as:\n    python <this> <design_dir> <script> [args...]\n\nThe scripts are NOT copied. They anchor everything on their own location -\n    self_path = os.path.dirname(os.path.realpath(__file__))\n    os.chdir(self_path)\n- so relocating them moves CondFilterConf.json and the __WorkSpace/__BackUp\nfolder with them. Instead the source is patched in memory and executed with\n__file__ still pointing at the original file, which keeps self_path, the chdir\nand every sibling path exactly as a normal run would have them.\n"""\nimport os\nimport re\nimport sys\nimport types\n\nASSIGN = re.compile(\n    r"""^(\\s*design_path\\s*=\\s*)os\\.path\\.abspath\\(\\s*[\'"][^\'"]*[\'"]\\s*\\)""",\n    re.MULTILINE)\n\nIMPORT_RE = re.compile(r"^\\s*(?:import|from)\\s+([A-Za-z_]\\w*)", re.MULTILINE)\n\n\ndef patch(src, design_dir):\n    return ASSIGN.subn(lambda m: m.group(1) + repr(design_dir), src)\n\n\ndef read(path):\n    with open(path, "r", encoding="utf-8", errors="replace") as f:\n        return f.read()\n\n\ndef main():\n    design_dir = sys.argv[1]\n    script = os.path.abspath(sys.argv[2])\n    args = sys.argv[3:]\n    script_dir = os.path.dirname(script)\n\n    sys.path.insert(0, script_dir)\n    sys.argv = [script] + args\n\n    entry_src = read(script)\n    entry_src, n = patch(entry_src, design_dir)\n    if not n:\n        sys.stderr.write(\n            "REDIRECT_FAILED: no \'design_path = os.path.abspath(...)\' in "\n            + os.path.basename(script) + "\\n")\n        sys.exit(3)\n\n    # Only siblings the entry actually imports, and only those that carry the\n    # same hardcoded assignment. Pre-loading anything else could run import\n    # side effects the real invocation never triggers.\n    patched = [os.path.basename(script)]\n    for name in sorted(set(IMPORT_RE.findall(entry_src))):\n        sib = os.path.join(script_dir, name + ".py")\n        if not os.path.exists(sib) or name in sys.modules:\n            continue\n        sib_src = read(sib)\n        sib_src, sn = patch(sib_src, design_dir)\n        if not sn:\n            continue\n        mod = types.ModuleType(name)\n        mod.__file__ = sib\n        sys.modules[name] = mod\n        try:\n            exec(compile(sib_src, sib, "exec"), mod.__dict__)\n        except Exception:\n            del sys.modules[name]\n            raise\n        patched.append(name + ".py")\n\n    sys.stderr.write("REDIRECT_OK: " + ", ".join(patched) + "\\n")\n    g = {"__name__": "__main__", "__file__": script, "__builtins__": __builtins__}\n    exec(compile(entry_src, script, "exec"), g)\n\n\nif __name__ == "__main__":\n    main()\n'


def _write_redirect_launcher():
    """Drop the launcher in a temp file; returns (path, tempdir)."""
    tmp = Path(tempfile.mkdtemp(prefix="ea_redirect_"))
    f = tmp / "_ea_redirect.py"
    f.write_text(_REDIRECT_LAUNCHER, encoding="utf-8")
    return f, tmp


def _run_project_script(py_exe, script, args=(), cwd=None, timeout=None,
                        design_override=None, log=None):
    """Run one of the SkipBo project scripts the way a normal python would.

    design_override redirects the design folder those scripts hardcode. The
    script itself is neither modified nor moved; see _REDIRECT_LAUNCHER. A
    redirect that cannot be applied is fatal, because carrying on would write
    the chapter into the wrong workspace.
    """
    tmp = None
    try:
        if design_override is not None:
            launcher, tmp = _write_redirect_launcher()
            cmd = [py_exe, str(launcher), str(design_override), str(script),
                   *[str(a) for a in args]]
        else:
            cmd = [py_exe, "-c", _BOOTSTRAP, str(script),
                   *[str(a) for a in args]]
        result = subprocess.run(
            cmd,
            cwd=str(cwd) if cwd else None,
            capture_output=True, text=True, encoding="utf-8", errors="replace",
            timeout=timeout,
        )
        if design_override is not None:
            stderr = result.stderr or ""
            if "REDIRECT_FAILED" in stderr or result.returncode == 3:
                raise RuntimeError(
                    "Could not point " + Path(script).name + " at the selected "
                    "workspace: its 'design_path = os.path.abspath(...)' line "
                    "was not found, so it would have written to its own "
                    "hardcoded workspace instead."
                    "\n\n"
                    "The Level Builder scripts have probably changed shape. "
                    "Report this rather than running the step - the run would "
                    "target the wrong design folder."
                )
            # The launcher reports what it patched on stderr; surface it in the
            # log and keep it out of the script's own error output.
            kept = []
            for line in stderr.splitlines():
                if line.startswith("REDIRECT_OK:"):
                    if log:
                        log(f"  Redirected design_path -> {design_override}")
                        log(f"  (patched: {line.split(':', 1)[1].strip()})")
                else:
                    kept.append(line)
            result = subprocess.CompletedProcess(
                result.args, result.returncode, result.stdout, "\n".join(kept))
        return result
    finally:
        if tmp is not None:
            shutil.rmtree(tmp, ignore_errors=True)


def preflight_project_scripts(cfg, log):
    """Check every borrowed-interpreter script BEFORE the run starts.

    These scripts sit at steps 1, 2 and 4 of Chapter Update, each mutating
    design files. Discovering a missing dependency at step 4 leaves a
    half-applied chapter needing an svn revert, so the dependency check is
    hoisted to the top where failing costs nothing.

    Returns a list of human-readable problems; empty means good to go.
    """
    problems = []
    targets = []
    lvl = getattr(cfg, "lvl_py", None)
    if lvl:
        targets += [Path(lvl) / n for n in
                    ("conf_fill_ex.py", "gen_captures.py", "gen_captures_others.py")]
    mp = getattr(cfg, "map_py", None)
    if mp:
        targets.append(Path(mp) / "save_map_data.py")

    for t in targets:
        if not t.exists():
            continue  # absent scripts are skipped at run time too
        try:
            _resolve_interpreter(t, log)
        except RuntimeError as e:
            problems.append(str(e))

    # Same reasoning as the dependency checks: Magic Builder refuses to
    # overwrite existing levels, and finding that out at step 1 means the
    # design files have already been touched by then.
    try:
        check_level_conflicts(cfg, log)
    except RuntimeError as e:
        problems.append(str(e))
    except Exception as e:
        log(f"  ⚠  Could not pre-check level conflicts: {e}", "WARN")
    return problems


def _ch_sub(text, old_val, new_val):
    """Replace old_val with new_val only when not adjacent to other digits."""
    if text is None:
        return text
    pattern = rf'(?<![0-9]){re.escape(str(old_val))}(?![0-9])'
    return re.sub(pattern, str(new_val), str(text))

def _cell_replace(val, old_ch, new_ch, old_disp, new_disp):
    """Replace all chapter references in a cell value."""
    if val is None:
        return val
    if isinstance(val, (int, float)):
        iv = int(val)
        if iv == old_ch * 1000:
            return new_ch * 1000
        return val
    if isinstance(val, str):
        s = val
        s = _ch_sub(s, f"章节{old_disp}", f"章节{new_disp}")
        s = _ch_sub(s, f"第{old_disp}章", f"第{new_disp}章")
        s = _ch_sub(s, f"章节{old_ch}关卡", f"章节{new_ch}关卡")
        s = _ch_sub(s, str(old_ch * 1000), str(new_ch * 1000))
        return s
    return val

def _insert_row_at(ws, row_idx, values):
    """Insert a blank row at row_idx (1-based) and write values into it."""
    ws.insert_rows(row_idx)
    for col, val in enumerate(values, start=1):
        ws.cell(row=row_idx, column=col, value=val)

def _insert_block_after(ws, after_row, rows_values):
    """Insert a block of rows directly after `after_row` (1-based), in order."""
    n = len(rows_values)
    if n == 0:
        return
    ws.insert_rows(after_row + 1, amount=n)
    for r_off, vals in enumerate(rows_values):
        for col, val in enumerate(vals, start=1):
            ws.cell(row=after_row + 1 + r_off, column=col, value=val)

def _cn_num(n):
    """Convert an integer (0-999) to simplified-Chinese numerals, chapter style.
    e.g. 78 -> 七十八, 79 -> 七十九, 100 -> 一百, 101 -> 一百零一."""
    d = "零一二三四五六七八九"
    n = int(n)
    if n < 0 or n > 999:
        return str(n)
    if n < 10:
        return d[n]
    if n < 20:                      # 10-19 -> 十, 十一 … 十九
        return "十" + (d[n % 10] if n % 10 else "")
    if n < 100:                     # 20-99 -> 二十, 七十八 …
        t, o = divmod(n, 10)
        return d[t] + "十" + (d[o] if o else "")
    h, rem = divmod(n, 100)         # 100-999
    s = d[h] + "百"
    if rem == 0:
        return s
    t, o = divmod(rem, 10)
    if t == 0:
        return s + "零" + d[o]
    return s + d[t] + "十" + (d[o] if o else "")

def _parse_noti_dt(v):
    """notification.xlsx dates arrive as datetime or 'YYYY/MM/DD HH:MM:SS' text."""
    if isinstance(v, _datetime):
        return v
    for fmt in ("%Y/%m/%d %H:%M:%S", "%Y-%m-%d %H:%M:%S", "%Y/%m/%d", "%Y-%m-%d"):
        try:
            return _datetime.strptime(str(v).strip(), fmt)
        except (ValueError, TypeError):
            continue
    return None


def _fmt_noti_dt(d, like):
    """Write the new date back in whatever form the sheet already uses."""
    return d if isinstance(like, _datetime) else d.strftime("%Y/%m/%d %H:%M:%S")


def _swap_disp(val, old_disp, new_disp):
    """Swap a cell that is nothing but the display chapter number.

    chest_rewards column E is a bare 84 -- no 章节 prefix, no map id -- so
    _cell_replace's patterns (which all require context to avoid mangling
    unrelated numbers) correctly refuse to touch it. Those columns have to be
    named explicitly instead.
    """
    if isinstance(val, int) and val == old_disp:
        return new_disp
    if isinstance(val, str) and val.strip() == str(old_disp):
        return str(new_disp)
    return None      # not a bare chapter number -- caller falls back


def _insert_rows_by_prefix(ws, old_ch, new_ch, old_disp, new_disp,
                            series_prefixes, disp_cols=None):
    """For each series prefix, collect ALL matching rows and insert them right after the group."""
    # prefix -> {"last_row_num": int, "rows": [values_list, ...]}
    prefix_data: dict = {}
    for row in ws.iter_rows(min_row=2):
        val = row[0].value
        if val is None:
            continue
        a_str = str(val)
        for pfx in series_prefixes:
            if a_str.startswith(f"{pfx}{old_ch:02d}"):
                if pfx not in prefix_data:
                    prefix_data[pfx] = {"last_row_num": row[0].row, "rows": []}
                prefix_data[pfx]["last_row_num"] = row[0].row
                prefix_data[pfx]["rows"].append([c.value for c in row])
                break

    if not prefix_data:
        return 0

    count = 0
    # Process prefixes from bottom to top so row insertions don't shift later groups
    for pfx in sorted(prefix_data, key=lambda p: prefix_data[p]["last_row_num"], reverse=True):
        last_row_num = prefix_data[pfx]["last_row_num"]
        rows_to_insert = prefix_data[pfx]["rows"]
        # Insert in REVERSE row order so that inserting all at last_row_num+1
        # results in the correct final order: [A', B', C'] after the group.
        for row_vals in reversed(rows_to_insert):
            new_row = []
            for i, val in enumerate(row_vals):
                col = i + 1
                if col == 1:
                    s = str(val).replace(f"{pfx}{old_ch:02d}", f"{pfx}{new_ch:02d}", 1)
                    try:
                        val = int(s)
                    except ValueError:
                        val = s
                elif disp_cols and col in disp_cols:
                    swapped = _swap_disp(val, old_disp, new_disp)
                    val = (swapped if swapped is not None
                           else _cell_replace(val, old_ch, new_ch, old_disp, new_disp))
                else:
                    val = _cell_replace(val, old_ch, new_ch, old_disp, new_disp)
                new_row.append(val)
            _insert_row_at(ws, last_row_num + 1, new_row)
            count += 1

    return count

def run_step3_level_validation(cfg: AppConfig, log):
    """Build validation workbook from pve_level + pve_layout."""
    design = cfg.design()
    lvl_path = design / "pve_level.xlsx"
    layout_path = design / "pve_layout.xlsx"
    out_path = design / f"validation_chapter_{cfg.new_ch}.xlsx"

    for p in (lvl_path, layout_path):
        if not p.exists():
            raise FileNotFoundError(f"Not found: {p}")

    log(f"Reading pve_level.xlsx …")
    wb_lvl = openpyxl.load_workbook(str(lvl_path), read_only=True, data_only=True)
    ws_lvl = wb_lvl["level"]
    lvl_rows = {r[0]: r[3] for r in ws_lvl.iter_rows(min_row=2, values_only=True)
                if r[0] and cfg.new_ch * 1000 < int(r[0]) <= cfg.new_ch * 1000 + 200}
    wb_lvl.close()
    log(f"  Found {len(lvl_rows)} level rows for chapter {cfg.new_ch}")

    log("Reading pve_layout.xlsx …")
    wb_lay = openpyxl.load_workbook(str(layout_path), read_only=True, data_only=True)
    ws_lay = wb_lay["layout"]
    lay_rows = {r[0]: r[2] for r in ws_lay.iter_rows(min_row=2, values_only=True)
                if r[0] and cfg.new_ch * 1000 < int(r[0]) <= cfg.new_ch * 1000 + 200}
    wb_lay.close()
    log(f"  Found {len(lay_rows)} layout rows for chapter {cfg.new_ch}")

    log("Building validation workbook …")
    wb_out = openpyxl.Workbook()
    ws = wb_out.active
    ws.title = "Validation"

    # Headers
    hdr_fill = PatternFill("solid", fgColor="1a3a5c")
    hdr_font = Font(bold=True, color="F0D060")
    for col, title in enumerate(["Level ID", "收集卡 Count (col D)", "Layout Data (col C)", "Status"], 1):
        c = ws.cell(row=1, column=col, value=title)
        c.fill = hdr_fill; c.font = hdr_font

    all_ids = sorted(set(lvl_rows) | set(lay_rows))
    issues = 0
    for row_idx, level_id in enumerate(all_ids, start=2):
        lvl_val = lvl_rows.get(level_id, "")
        lay_val = lay_rows.get(level_id, "")

        # Count 收集卡:1 in lvl_val
        collect_count = 0
        if isinstance(lvl_val, str):
            collect_count = lvl_val.count("收集卡:1")
        elif isinstance(lvl_val, (int, float)):
            collect_count = int(lvl_val)

        # Count card":14, in layout
        card14_count = 0
        if isinstance(lay_val, str):
            card14_count = lay_val.count('card":14,') + lay_val.count('card": 14,')

        ok = (collect_count == 0) or (collect_count == card14_count)
        status = "✓ OK" if ok else f"✗ MISMATCH: need {collect_count}, found {card14_count}"
        if not ok:
            issues += 1

        ws.cell(row=row_idx, column=1, value=level_id)
        ws.cell(row=row_idx, column=2, value=str(lvl_val)[:500] if lvl_val else "")
        ws.cell(row=row_idx, column=3, value=str(lay_val)[:200] if lay_val else "")
        ws.cell(row=row_idx, column=4, value=status)

        if not ok:
            err_fill = PatternFill("solid", fgColor="5c1a1a")
            for col in range(1, 5):
                ws.cell(row=row_idx, column=col).fill = err_fill

    # Conditional formatting for column B
    red_fill = PatternFill("solid", fgColor="5c1a1a")
    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 40
    ws.column_dimensions["C"].width = 60
    ws.column_dimensions["D"].width = 30

    wb_out.save(str(out_path))
    log(f"  ✓ Validation saved to {out_path.name}")
    log(f"  {'✓ All OK' if issues == 0 else f'⚠ {issues} mismatches found'}")
    return str(out_path)

def run_step4_reward_config(cfg: AppConfig, log):
    """Configure Chest, Item, Store, Notification for new chapter."""
    design = cfg.design()
    old_ch, new_ch = cfg.old_ch, cfg.new_ch
    old_d, new_d = cfg.disp_old, cfg.disp_new

    # ── CHEST ────────────────────────────────────────────────────────────────
    chest_path = design / "chest.xlsx"
    if not chest_path.exists():
        raise FileNotFoundError(f"chest.xlsx not found: {chest_path}")
    log(f"Opening chest.xlsx …")
    wb = openpyxl.load_workbook(str(chest_path))

    ws_cond = wb["chest_conditions"]
    n = _insert_rows_by_prefix(
        ws_cond, old_ch, new_ch, old_d, new_d,
        series_prefixes=["60", "70", "80"],
        # No column is skipped by series. Col E (免费宝箱) carries 章节NN in ALL
        # of 60/70/80; col F (额外付费宝箱) carries it only in 60/70 and is
        # empty for 80 — which _cell_replace already handles by leaving None
        # alone. The old skip_col_e_for=["80"] came from a misreading of the
        # design doc and left 49 rows on the previous chapter number.
    )
    log(f"  chest_conditions: inserted {n} rows in-place")

    ws_rew = wb["chest_rewards"]
    n = _insert_rows_by_prefix(
        ws_rew, old_ch, new_ch, old_d, new_d,
        series_prefixes=["60", "70", "80"],
        # chest_rewards has a different shape to chest_conditions: col D is
        # reward content ("1000,200") and must NOT be touched, while col E is
        # the bare chapter number with no 章节 prefix for _cell_replace to key
        # off. Per the design doc: "it only has chapter number, simply change
        # it to new chapter number".
        disp_cols={5},
    )
    log(f"  chest_rewards: inserted {n} rows in-place")
    wb.save(str(chest_path))
    log("  ✓ chest.xlsx saved")

    # ── ITEM ─────────────────────────────────────────────────────────────────
    item_path = design / "item.xlsx"
    if not item_path.exists():
        raise FileNotFoundError(f"item.xlsx not found: {item_path}")
    log("Opening item.xlsx (item sheet) …")
    # Load data_only first to resolve any formula values in column A
    wb2_ro = openpyxl.load_workbook(str(item_path), data_only=True)
    ws_ro = wb2_ro["item"]
    last_item_row_num = None
    last_item_vals_ro = None
    for row in ws_ro.iter_rows(min_row=2):
        val = row[0].value
        if val is None:
            continue
        try:
            id_int = int(float(val))
        except (ValueError, TypeError):
            continue
        if 4000 <= id_int <= 4199:
            last_item_row_num = row[0].row
            last_item_vals_ro = [c.value for c in row]
    wb2_ro.close()

    wb2 = openpyxl.load_workbook(str(item_path))
    ws_item = wb2["item"]
    if last_item_vals_ro is not None:
        new_item = [_cell_replace(v, old_ch, new_ch, old_d, new_d) for v in last_item_vals_ro]
        new_item[0] = int(float(last_item_vals_ro[0])) + 1
        _insert_row_at(ws_item, last_item_row_num + 1, new_item)
        log(f"  item: inserted row {new_item[0]} after sheet row {last_item_row_num}")
    else:
        log(f"  item: no rows found in 4000-4199 range")
    wb2.save(str(item_path))
    log("  ✓ item.xlsx saved")

    # ── STORE ────────────────────────────────────────────────────────────────
    store_path = design / "store.xlsx"
    if not store_path.exists():
        raise FileNotFoundError(f"store.xlsx not found: {store_path}")
    log("Opening store.xlsx …")
    wb3 = openpyxl.load_workbook(str(store_path))

    for sheet_name in ["store_recharge", "store_package"]:
        ws_s = wb3[sheet_name]
        # Find the last row whose ID is in the 40001-40999 chapter-store range
        last_store_row_num = None
        last_store_vals = None
        for row in ws_s.iter_rows(min_row=2):
            val = row[0].value
            if val is None:
                continue
            try:
                id_int = int(val)
            except (ValueError, TypeError):
                continue
            if 40001 <= id_int <= 40999:
                last_store_row_num = row[0].row
                last_store_vals = [c.value for c in row]
        if last_store_vals:
            new_r = [_cell_replace(v, old_ch, new_ch, old_d, new_d) for v in last_store_vals]
            new_r[0] = int(last_store_vals[0]) + 1  # increment ID
            _insert_row_at(ws_s, last_store_row_num + 1, new_r)
            log(f"  {sheet_name}: inserted row {new_r[0]} after row {last_store_row_num}")
        else:
            log(f"  {sheet_name}: no rows found in 40001-40999 range")
    wb3.save(str(store_path))
    log("  ✓ store.xlsx saved")

    # ── NOTIFICATION ─────────────────────────────────────────────────────────
    noti_path = design / "notification.xlsx"
    if not noti_path.exists():
        raise FileNotFoundError(f"notification.xlsx not found: {noti_path}")
    log("Opening notification.xlsx …")
    wb4 = openpyxl.load_workbook(str(noti_path))
    ws_noti = wb4["server_notification"]

    # Keep row numbers: the previous chapter's switch has to be turned off in
    # place, not just read.
    noti_all = [(r[0].row, [c.value for c in r]) for r in ws_noti.iter_rows(min_row=2)]
    ch_noti = [(rn, v) for rn, v in noti_all
               if len(v) > 1 and v[1] and f"{old_d}章节" in str(v[1])]
    prior = [(rn, v) for rn, v in noti_all
             if len(v) > 1 and v[1] and "章节上线" in str(v[1])]

    if ch_noti:
        src_rn, src = ch_noti[-1]
        last_noti = list(src)
        try:
            last_noti[0] = max(r[0] for _, r in noti_all
                               if r[0] and isinstance(r[0], int)) + 1
        except (ValueError, TypeError):
            pass
        last_noti[1] = f"{new_d}章节上线"
        last_noti[2] = "新关卡章节开放 "
        # D (开关): exactly one chapter notification is live at a time — the
        # newest. Verified across chapters 75-84: every older row is False and
        # only the latest is True. So the new row switches on and the row it
        # was cloned from switches off.
        last_noti[3] = True

        # E/F (start/end): the chapter cadence is a fixed 14 days (verified
        # unbroken across 75-84) and the window is always start+2 (3 days live,
        # counting both ends). Copying the source row verbatim would publish a
        # notification dated in the past, so advance by the interval the sheet
        # itself shows between the last two chapters.
        step_days = 14
        if len(prior) >= 2:
            d_prev, d_last = _parse_noti_dt(prior[-2][1][4]), _parse_noti_dt(prior[-1][1][4])
            if d_prev and d_last and (d_last - d_prev).days > 0:
                step_days = (d_last - d_prev).days
        start_dt = _parse_noti_dt(src[4]) if len(src) > 4 else None
        end_dt = _parse_noti_dt(src[5]) if len(src) > 5 else None
        if start_dt:
            new_start = start_dt + _timedelta(days=step_days)
            last_noti[4] = _fmt_noti_dt(new_start, src[4])
            span = (end_dt - start_dt).days if end_dt else 2
            last_noti[5] = _fmt_noti_dt(new_start + _timedelta(days=span), src[5])
            log(f"  server_notification: window {last_noti[4]} → {last_noti[5]} "
                f"(+{step_days}d from chapter {old_d}) — confirm against the "
                f"actual go-live date")
        else:
            log("  ! server_notification: could not parse the start date — "
                "set the window manually", "WARN")

        # Alternate PUSH_NOTE_NEW_CHAPTER2 / PUSH_NOTE_NEW_CHAPTER25 from what
        # the previous row really uses (a hand-edit once broke the run, so an
        # absolute rule would stay broken forever after).
        prev_h = str(last_noti[7]) if len(last_noti) > 7 else ""
        last_noti[7] = ("PUSH_NOTE_NEW_CHAPTER25" if "CHAPTER2" in prev_h
                         and "CHAPTER25" not in prev_h else "PUSH_NOTE_NEW_CHAPTER2")

        ws_noti.append(last_noti[:9])
        for rn, v in prior:
            if len(v) > 3 and str(v[3]).strip().lower() == "true":
                ws_noti.cell(row=rn, column=4, value=False)
                log(f"  server_notification: chapter row {v[1]} switch True -> False")
        log(f"  server_notification: added row for 章节{new_d}, "
            f"switch=True, push={last_noti[7]}")
    else:
        log(f"  server_notification: no chapter {old_d} notification found")
    wb4.save(str(noti_path))
    log("  ✓ notification.xlsx saved")
    log("  ✓ All reward config completed")

def run_step5_art_assets(cfg: AppConfig, log):
    """Duplicate and rename art asset folders for new chapter."""
    old_ch, new_ch = cfg.old_ch, cfg.new_ch
    old_map_id = old_ch * 1000
    new_map_id = new_ch * 1000

    # ── Dynamic map folder ───────────────────────────────────────────────────
    src_map = cfg.client_map / f"dynamic_map_{old_map_id}"
    dst_map = cfg.client_map / f"dynamic_map_{new_map_id}"
    if not src_map.exists():
        raise FileNotFoundError(f"Source map folder not found: {src_map}")
    if dst_map.exists():
        log(f"  dynamic_map_{new_map_id} already exists, skipping copy")
    else:
        shutil.copytree(str(src_map), str(dst_map))
        log(f"  ✓ Copied dynamic_map_{old_map_id} → dynamic_map_{new_map_id}")

    # Rename files inside (replace old_map_id with new_map_id in filenames)
    for fpath in dst_map.rglob("*"):
        if fpath.is_file():
            old_name = fpath.name
            new_name = old_name.replace(str(old_map_id), str(new_map_id))
            new_name = new_name.replace(f"_{old_ch}_", f"_{new_ch}_")
            if new_name != old_name:
                fpath.rename(fpath.parent / new_name)
    log(f"  ✓ Renamed files inside dynamic_map_{new_map_id}")

    # Delete meta files that Unity will regenerate
    for meta in dst_map.rglob("*.meta"):
        meta.unlink()
    log("  ✓ Removed stale .meta files (Unity will regenerate)")

    # ── AdventureMapBg placeholder ───────────────────────────────────────────
    adv_bg_dir = cfg.client_tex / "AdventureMapBg"
    old_bg = adv_bg_dir / f"img_chapter{cfg.disp_old}.png"
    new_bg = adv_bg_dir / f"img_chapter{cfg.disp_new}.png"
    if old_bg.exists() and not new_bg.exists():
        shutil.copy2(str(old_bg), str(new_bg))
        log(f"  ✓ AdventureMapBg: copied placeholder img_chapter{cfg.disp_new}.png")
    else:
        log(f"  – AdventureMapBg: {new_bg.name} already exists or source missing")

    # ── ChapeDoneBg placeholder ──────────────────────────────────────────────
    done_bg_dir = cfg.client_tex / "ChapeDoneBg"
    old_done = done_bg_dir / f"img_chapter_unlock_{cfg.disp_old}.png"
    new_done = done_bg_dir / f"img_chapter_unlock_{cfg.disp_new}.png"
    if old_done.exists() and not new_done.exists():
        shutil.copy2(str(old_done), str(new_done))
        log(f"  ✓ ChapeDoneBg: copied placeholder img_chapter_unlock_{cfg.disp_new}.png")
    else:
        log(f"  – ChapeDoneBg: {new_done.name} already exists or source missing")

    log("  ✓ Art assets prepared (replace placeholder images with real art)")

def run_step6_asset_config(cfg: AppConfig, log):
    """Configure asset_ref, icon, and pve_chapter workbooks for the new chapter.

    Every block is copied from the previous chapter and inserted directly after
    it (never appended at the end), with chapter numbers replaced in three forms:
      • internal map id   (old_ch*1000 -> new_ch*1000, e.g. 88000 -> 89000)
      • Chinese numerals  (七十八 -> 七十九)  — used in Alias/PIC columns
      • display number     (78 -> 79)         — standalone-digit replacement
    """
    design   = cfg.design()
    old_ch, new_ch = cfg.old_ch, cfg.new_ch
    old_d,  new_d  = cfg.disp_old, cfg.disp_new
    old_mapid, new_mapid = old_ch * 1000, new_ch * 1000
    cn_old, cn_new = _cn_num(old_d), _cn_num(new_d)

    def tf(v, mapid=True, chinese=True, display=True):
        """Transform a single cell value across the three chapter-number forms."""
        if not isinstance(v, str):
            return v
        s = v
        if mapid:
            s = s.replace(str(old_mapid), str(new_mapid))
        if chinese:
            s = s.replace(cn_old, cn_new)
        if display:
            s = _ch_sub(s, old_d, new_d)
        return s

    # ── pve_chapter.xlsx (process first — provides fruit numbering) ───────────
    pvc_path = design / "pve_chapter.xlsx"
    if not pvc_path.exists():
        raise FileNotFoundError(f"pve_chapter.xlsx not found: {pvc_path}")
    log("Opening pve_chapter.xlsx …")
    wb3 = openpyxl.load_workbook(str(pvc_path))

    # Guard: refuse to run if the new chapter already exists (partial prior run).
    # Detect by the internal map id in col A (set even by the old broken code) or
    # the display number in col C, so corrupted leftovers are still caught.
    ws_ch = wb3["chapter"]
    if any((r[0].value == new_mapid) or (r[2].value == new_d)
           for r in ws_ch.iter_rows(min_row=2, max_col=3) if r[0].value is not None):
        raise RuntimeError(
            f"pve_chapter.xlsx already contains chapter {new_d} (map id {new_mapid}).\n\n"
            f"A previous run left chapter-{new_d} data behind. Revert the design "
            f"files first:\n"
            f"  svn revert asset_ref.xlsx icon.xlsx pve_chapter.xlsx "
            f"item.xlsx store.xlsx chest.xlsx\n"
            f"…then re-run step 6."
        )

    def rows_of(ws, ncol):
        return [(r[0].row, [c.value for c in r]) for r in
                ws.iter_rows(min_row=2, max_col=ncol) if r[0].value is not None]

    # Sheet: reward_zone — 4 zones/chapter; new ids follow, fruit refs advance
    ws_rz = wb3["reward_zone"]
    rz = rows_of(ws_rz, 7)
    old_zones = [(rn, v) for rn, v in rz if v[1] and f"章节{old_d}关卡" in str(v[1])]
    fruits_per_chapter = len(old_zones) or 4
    max_zone_id = max((v[0] for rn, v in rz if isinstance(v[0], int)), default=0)
    max_fruit   = max((v[0] for rn, v in rows_of(wb3["harvest"], 6)
                       if isinstance(v[0], int)), default=0)
    new_fruits  = [max_fruit + i + 1 for i in range(fruits_per_chapter)]
    new_zone_ids = [max_zone_id + i + 1 for i in range(len(old_zones))]
    new_rz = []
    for i, (rn, v) in enumerate(old_zones):
        nv = list(v)
        nv[0] = new_zone_ids[i]
        nv[1] = tf(v[1], mapid=False, chinese=False, display=True)   # 章节78关卡 -> 79
        nv[2] = tf(v[2], mapid=False, chinese=False, display=True)
        if isinstance(v[4], int):
            nv[4] = new_fruits[i]                                    # collectible -> new fruit
        new_rz.append(nv)
    if old_zones:
        _insert_block_after(ws_rz, old_zones[-1][0], new_rz)
        log(f"  reward_zone: inserted {len(new_rz)} zones "
            f"({new_zone_ids[0]}–{new_zone_ids[-1]}), fruits {new_fruits[0]}–{new_fruits[-1]}")

    # Sheet: harvest — one row per new fruit; atlas = MAP_FRUIT_{id//10}
    ws_hv = wb3["harvest"]
    hv = rows_of(ws_hv, 6)
    old_fruits = [nf - fruits_per_chapter for nf in new_fruits]
    old_hv = sorted([(rn, v) for rn, v in hv if v[0] in old_fruits],
                    key=lambda x: x[1][0])
    new_hv = []
    for (rn, v), nf in zip(old_hv, new_fruits):
        nv = list(v)
        nv[0] = nf
        for c in (1, 2, 3):
            nv[c] = re.sub(r"FRUIT\d+", f"FRUIT{nf}", str(v[c])) if v[c] else v[c]
        nv[5] = f"MAP_FRUIT_{nf // 10}"                              # atlas (decade-packed)
        new_hv.append(nv)
    if old_hv:
        _insert_block_after(ws_hv, old_hv[-1][0], new_hv)
        log(f"  harvest: inserted {len(new_hv)} fruits {new_fruits[0]}–{new_fruits[-1]} "
            f"(atlas {new_hv[0][5]}…{new_hv[-1][5]})")

    # Sheet: chapter — single row, per-column rules
    ch = rows_of(ws_ch, 17)
    old_row = next(((rn, v) for rn, v in ch if v[2] == old_d), None)
    if old_row:
        rn, v = old_row
        nv = list(v)
        nv[0]  = new_mapid                                          # A internal id
        nv[1]  = tf(v[1], mapid=False, chinese=True,  display=False) # B 第七十八章节
        nv[2]  = new_d                                              # C display number
        nv[3]  = tf(v[3], mapid=False, chinese=True,  display=False) # D 第七十八章节地图
        nv[4]  = f"{new_zone_ids[0]},{new_zone_ids[-1]}" if new_zone_ids else v[4]  # E zones
        nv[5]  = tf(v[5], mapid=False, chinese=False, display=True)  # F 章节局内背景图78
        # G: the 1-2-3 background-music rotation, advanced from the PREVIOUS
        # row's actual value rather than computed from the chapter number.
        # An absolute formula matches history right up until someone edits a
        # row by hand (chapter 84 was set to 音效1 where the formula says 3) —
        # after that it repeats the value it should have advanced past.
        # Reading what is really there keeps the rotation intact either way.
        _prev_sound = re.search(r"(\d+)\s*$", str(v[6] or ""))
        _n = int(_prev_sound.group(1)) if _prev_sound else 0
        nv[6]  = f"PVE游戏内背景音效{(_n % 3) + 1}"
        for i in (7, 8, 9, 10, 12, 13, 14, 15, 16):                 # H,I,J,K,M,N,O,P,Q
            nv[i] = tf(v[i], mapid=False, chinese=False, display=True)
        # L (11) stays constant
        _insert_block_after(ws_ch, rn, [nv])
        log(f"  chapter: inserted chapter {new_d} (mapid {new_mapid}, {nv[6]})")

    # Sheet: chest_plan — 3 variant groups (60.. / 70.. / 80..), one row each
    ws_cp = wb3["chest_plan"]
    cp = rows_of(ws_cp, 5)
    old_cp = [(rn, v) for rn, v in cp if v[1] and f"第{old_d}章" in str(v[1])]
    cp_count = 0
    for rn, v in sorted(old_cp, key=lambda x: x[0], reverse=True):
        nv = list(v)
        if v[0] is not None:                                        # A plan id (mapid)
            s = str(v[0]).replace(str(old_mapid), str(new_mapid))
            nv[0] = int(s) if isinstance(v[0], int) and s.isdigit() else s
        nv[1] = tf(v[1], mapid=False, chinese=False, display=True)  # B 第78章
        nv[4] = tf(v[4], mapid=False, chinese=False, display=True)  # E 章节78星…
        _insert_block_after(ws_cp, rn, [nv])
        cp_count += 1
    log(f"  chest_plan: inserted {cp_count} plan rows")

    # Sheet: map — 13 rows (geometry filled later by step 7)
    ws_map = wb3["map"]
    mp = rows_of(ws_map, 10)
    old_map = sorted([(rn, v) for rn, v in mp
                      if isinstance(v[0], int) and old_d * 1000 < v[0] < old_d * 1000 + 1000],
                     key=lambda x: x[1][0])
    new_map = []
    for rn, v in old_map:
        nv = list(v)
        nv[0] = v[0] + 1000                                         # A 78001 -> 79001
        nv[1] = tf(v[1], mapid=False, chinese=True,  display=False) # B 第七十八章节地图n
        nv[2] = tf(v[2], mapid=False, chinese=True,  display=False) # C 场景七十八地图片n
        nv[3] = tf(v[3], mapid=False, chinese=False, display=True)  # D 章节78关卡
        # J: direction-key type (0 down, 1 up, 2 Tip/"coming soon"). The last
        # chapter in the game ends on a Tip because there is nothing after it;
        # the moment a newer chapter exists that Tip has to become an up-arrow.
        # So the new block takes 2 and the block it was cloned from drops to 1.
        if v[9] in (1, 2, "1", "2"):
            nv[9] = 2
        # E–I (coordinates / fruit lists) copied as placeholders for step 7
        new_map.append(nv)
    if old_map:
        _insert_block_after(ws_map, old_map[-1][0], new_map)
        # Demote the previous chapter's Tip. Its row numbers are still valid:
        # the new block was inserted after the last of them.
        for rn, v in old_map:
            if v[9] in (2, "2"):
                ws_map.cell(row=rn, column=10, value=1)
                log(f"  map: chapter {old_d} row {v[0]} direction 2 -> 1 "
                    f"(no longer the last chapter)")
        log(f"  map: inserted {len(new_map)} rows ({new_map[0][0]}–{new_map[-1][0]})")

    wb3.save(str(pvc_path))
    log("  ✓ pve_chapter.xlsx saved")

    # ── asset_ref.xlsx / key_asset ───────────────────────────────────────────
    ar_path = design / "asset_ref.xlsx"
    if not ar_path.exists():
        raise FileNotFoundError(f"asset_ref.xlsx not found: {ar_path}")
    log("Opening asset_ref.xlsx (key_asset sheet) …")
    wb = openpyxl.load_workbook(str(ar_path))
    ws = wb["key_asset"]
    ka = rows_of(ws, 3)

    # 6 display-numbered groups: (label, old-pattern, new-pattern)
    display_groups = [
        ("MAP scene",           rf"^MAP_{old_d}_\d+$",            rf"^MAP_{new_d}_\d+$"),
        ("PVE_BG",              rf"^PVE_BG_{old_d}$",             rf"^PVE_BG_{new_d}$"),
        ("CHAPTER_UNLOCK",      rf"^CHAPTER_UNLOCK_{old_d}$",     rf"^CHAPTER_UNLOCK_{new_d}$"),
        ("STARPASS_CHAPTER_BG", rf"^STARPASS_CHAPTER_BG_{old_d}$", rf"^STARPASS_CHAPTER_BG_{new_d}$"),
        ("STARPASS_ENTRY_ICON", rf"^STARPASS_ENTRY_ICON_{old_d}$", rf"^STARPASS_ENTRY_ICON_{new_d}$"),
        ("CHAPTER_TITLE_BG",    rf"^CHAPTER_TITLE_BG_{old_d}$",   rf"^CHAPTER_TITLE_BG_{new_d}$"),
    ]
    planned = []   # (after_row, [rows], label)
    for label, old_pat, new_pat in display_groups:
        if any(re.match(new_pat, str(v[0])) for rn, v in ka if v[0]):
            log(f"  key_asset/{label}: chapter {new_d} already present — skipped")
            continue
        grp = [(rn, v) for rn, v in ka if v[0] and re.match(old_pat, str(v[0]))]
        if not grp:
            log(f"  key_asset/{label}: no chapter {old_d} rows found — skipped")
            continue
        new_rows = [[tf(c) for c in v] for rn, v in grp]
        planned.append((grp[-1][0], new_rows, f"{label} (+{len(new_rows)})"))

    # MAP_FRUIT atlas — data-driven: add only atlases the new fruits actually need
    fruit_re = re.compile(r"^MAP_FRUIT_(\d+)$")
    existing_atlas = {int(m.group(1)) for rn, v in ka
                      if v[0] and (m := fruit_re.match(str(v[0])))}
    fruit_rows = [(rn, v) for rn, v in ka if v[0] and fruit_re.match(str(v[0]))]
    needed = sorted({nf // 10 for nf in new_fruits} - existing_atlas)
    if needed and fruit_rows:
        last_fr = fruit_rows[-1][0]
        atlas_rows = [[f"MAP_FRUIT_{a}", f"MAP_FRUIT_{a}",
                       f"Assets/ResourcesAb/Icon/MapIcon{a}"] for a in needed]
        planned.append((last_fr, atlas_rows, f"MAP_FRUIT (+{len(needed)}: {needed})"))
    else:
        log(f"  key_asset/MAP_FRUIT: no new atlas needed "
            f"(fruits {new_fruits[0]}–{new_fruits[-1]} fit MAP_FRUIT_{new_fruits[0]//10})")

    # Apply bottom-up so earlier row numbers stay valid
    for after_row, new_rows, label in sorted(planned, key=lambda x: x[0], reverse=True):
        _insert_block_after(ws, after_row, new_rows)
        log(f"  key_asset/{label}: inserted after row {after_row}")
    wb.save(str(ar_path))
    log("  ✓ asset_ref.xlsx saved")

    # ── icon.xlsx / icon_chapter ─────────────────────────────────────────────
    icon_path = design / "icon.xlsx"
    if not icon_path.exists():
        raise FileNotFoundError(f"icon.xlsx not found: {icon_path}")
    log("Opening icon.xlsx (icon_chapter sheet) …")
    wb2 = openpyxl.load_workbook(str(icon_path))
    ws_icon = wb2["icon_chapter"]
    ic = rows_of(ws_icon, 5)
    if any(v[1] and ("场景" + cn_new) in str(v[1]) for rn, v in ic):
        log(f"  icon_chapter: chapter {new_d} already present — skipped")
    else:
        old_block = [(rn, v) for rn, v in ic
                     if v[1] and ("场景" + cn_old) in str(v[1])]
        new_icon = []
        for rn, v in old_block:
            m = re.match(r"FRUIT(\d+)_(\d+)", str(v[0]))
            if not m:
                continue
            nfid = int(m.group(1)) + fruits_per_chapter
            state = m.group(2)
            nv = list(v)
            nv[0] = f"FRUIT{nfid}_{state}"
            nv[1] = f"场景{cn_new}收获物{nfid}状态{state}"
            nv[2] = f"fruit{nfid}_{state}"
            new_icon.append(nv)
        if old_block and new_icon:
            _insert_block_after(ws_icon, old_block[-1][0], new_icon)
            log(f"  icon_chapter: inserted {len(new_icon)} rows "
                f"(fruits {new_fruits[0]}–{new_fruits[-1]}, 场景{cn_new})")
        else:
            log(f"  icon_chapter: no chapter {old_d} (场景{cn_old}) block found — skipped")
    wb2.save(str(icon_path))
    log("  ✓ icon.xlsx saved")
    log("  ✓ All asset config completed")

def run_step10_level_balancing(cfg: AppConfig, log):
    """Build level balancing workbook from simulation output."""
    out_dir = cfg.sim_path / "Output"
    if not out_dir.exists():
        raise FileNotFoundError(f"Output dir not found: {out_dir}")

    ch_name = f"Chapter_{cfg.new_ch}"
    ch_folders = sorted([d for d in out_dir.iterdir() if d.is_dir() and ch_name in d.name], reverse=True)
    if not ch_folders:
        raise FileNotFoundError(f"No {ch_name} output folder found in {out_dir}")
    run_folder = ch_folders[0]
    log(f"Reading from: {run_folder.name}")

    # Find summary analysis file
    summary_files = list(run_folder.glob("summary_analysis*.xlsx")) + list(run_folder.glob("*.xlsx"))
    if not summary_files:
        raise FileNotFoundError(f"No xlsx files in {run_folder}")
    src_path = summary_files[0]
    log(f"Source file: {src_path.name}")

    wb_src = openpyxl.load_workbook(str(src_path), read_only=True, data_only=True)
    ws_src = wb_src.active
    all_rows = list(ws_src.iter_rows(values_only=True))
    wb_src.close()
    src_head = [str(h or "") for h in (all_rows[0] if all_rows else [])]
    src_data = all_rows[1:]

    # summary_analysis columns are "关卡ID,机器人ID," + the simulator's own
    # analysis2 header, i.e. E=通关率 (win rate) and H=三星率 (3-star rate).
    # The design doc warns the position can move ("please read 3 star rate"),
    # so match the header text and fall back to the documented letters.
    def _find_col(names, fallback):
        for i, h in enumerate(src_head):
            if any(n in h for n in names):
                return i
        log(f"  ! {names[0]} column not found in header — assuming column "
            f"{chr(ord('A') + fallback)}", "WARN")
        return fallback

    i_star3 = _find_col(["三星率", "3星率", "3★"], 7)   # H
    i_win = _find_col(["通关率"], 4)                    # E
    log(f"  Columns: 3-star={chr(ord('A') + i_star3)}, win rate={chr(ord('A') + i_win)}")

    # Create balancing workbook
    out_path = run_folder / f"level_balance_{cfg.new_ch}.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Balance"

    hdr_fill = PatternFill("solid", fgColor="1a3a5c")
    hdr_font = Font(bold=True, color="F0D060")
    headers = ["Level ID", "3★ Rate (raw)", "Win Rate", "Actual 3★ Rate"]
    for col, h in enumerate(headers, 1):
        c = ws.cell(row=1, column=col, value=h)
        c.fill = hdr_fill; c.font = hdr_font

    n_rows = 0
    for row in src_data:
        if row[0] is None:
            continue
        level_id = row[0]
        star3_raw = row[i_star3] if len(row) > i_star3 else None
        win_rate  = row[i_win] if len(row) > i_win else None
        ws.cell(row=n_rows + 2, column=1, value=level_id)
        ws.cell(row=n_rows + 2, column=2, value=star3_raw)
        ws.cell(row=n_rows + 2, column=3, value=win_rate)
        ws.cell(row=n_rows + 2, column=4, value=f"=B{n_rows+2}/C{n_rows+2}" if win_rate else None)
        n_rows += 1

    if n_rows > 0:
        # Format C, D as percentage
        from openpyxl.styles import numbers
        for row in ws.iter_rows(min_row=2, max_row=n_rows+1, min_col=3, max_col=4):
            for cell in row:
                cell.number_format = "0.0%"

        # Conditional formatting – Win Rate (col C)
        warn_fill = PatternFill("solid", fgColor="5c3a1a")
        bad_fill  = PatternFill("solid", fgColor="5c1a1a")
        c_range = f"C2:C{n_rows+1}"
        ws.conditional_formatting.add(c_range,
            CellIsRule(operator="between", formula=["0.05","0.12"], fill=warn_fill))
        ws.conditional_formatting.add(c_range,
            CellIsRule(operator="lessThan", formula=["0.05"], fill=bad_fill))

        # Conditional formatting – Actual 3★ Rate (col D)
        d_range = f"D2:D{n_rows+1}"
        ws.conditional_formatting.add(d_range,
            CellIsRule(operator="lessThan", formula=["0.1"], fill=bad_fill))
        ws.conditional_formatting.add(d_range,
            CellIsRule(operator="greaterThan", formula=["0.9"], fill=bad_fill))

    for col in ["A", "B", "C", "D"]:
        ws.column_dimensions[col].width = 16

    wb.save(str(out_path))
    log(f"  ✓ Level balance workbook saved: {out_path.name}")
    log(f"  {n_rows} levels processed")
    return str(out_path)


# ══════════════════════════════════════════════════════════════════════════════
#  QA VALIDATION  (Chapter Creation & Level Delivery – Review Checklist)
# ══════════════════════════════════════════════════════════════════════════════
# Runs last and re-derives, from the saved workbooks, whatever the review
# checklist can be checked mechanically. It deliberately re-reads the files
# rather than trusting what the earlier steps believed they wrote — a check that
# shares state with the thing it checks proves nothing.
#
# Statuses: PASS, FAIL (blocks sign-off), WARN (look at it), MANUAL (a human has
# to eyeball it — Unity scenes, fruit placement, screenshots).

VALID_DIFFICULTY = {"简单", "中简", "中等", "中难", "困难"}
MIN_HANDS, MAX_HANDS = 4, 7


class _Report:
    def __init__(self, log):
        self._log, self.rows = log, []

    def add(self, section, check, status, detail=""):
        self.rows.append((section, check, status, detail))
        if status != "PASS":
            self._log(f"    {status}: {check}" + (f" — {detail}" if detail else ""))

    def ok(self, section, check, cond, detail=""):
        self.add(section, check, "PASS" if cond else "FAIL", "" if cond else detail)

    def count(self, status):
        return sum(1 for r in self.rows if r[2] == status)


def _rows_of(path, sheet, max_col=None):
    if not path.exists():
        return None
    wb = openpyxl.load_workbook(str(path), data_only=True, read_only=True)
    try:
        if sheet not in wb.sheetnames:
            return None
        return [list(r) for r in wb[sheet].iter_rows(min_row=1, max_col=max_col,
                                                     values_only=True)]
    finally:
        wb.close()


def _col_by_header(header, *names, default=None):
    """Locate a column by its Chinese header.

    The review checklist's column letters are shifted by two against the real
    sheet (its "Column Q" is 星分数条件, actually column P), so the header text
    is the only identifier worth trusting.
    """
    for i, h in enumerate(header):
        t = str(h or "")
        if any(n in t for n in names):
            return i
    return default


def _nums(v):
    # `str(v or "")` would be wrong here: 0 is falsy, and 0 is a meaningful
    # value in these sheets (direction type 0 = down, difficulty 0 = normal).
    return [int(float(x)) for x in
            re.findall(r"-?\d+(?:\.\d+)?", "" if v is None else str(v))]


def run_qa_validation(cfg: AppConfig, log):
    """Validate the finished chapter against the review checklist."""
    design = cfg.design()
    new_ch, old_ch = cfg.new_ch, cfg.old_ch
    new_d, old_d = cfg.disp_new, cfg.disp_old
    lo, hi = new_ch * 1000 + 1, new_ch * 1000 + 200
    rp = _Report(log)

    # ── 2. Level generation ──────────────────────────────────────────────────
    log("  [2] Level generation …")
    lvl = _rows_of(design / "pve_level.xlsx", "level")
    lay = _rows_of(design / "pve_layout.xlsx", "layout")
    lvl_hdr = lvl[0] if lvl else []
    lvl_ids = [r[0] for r in (lvl[1:] if lvl else []) if isinstance(r[0], int)]
    lay_ids = [r[0] for r in (lay[1:] if lay else []) if isinstance(r[0], int)]
    new_lvl = [i for i in lvl_ids if lo <= i <= hi]
    new_lay = [i for i in lay_ids if lo <= i <= hi]

    if lvl is None:
        rp.add("2 Level Generation", "pve_level.xlsx readable", "FAIL", "file/sheet missing")
    else:
        rp.ok("2 Level Generation", f"pve_level has all 200 IDs ({lo}-{hi})",
              len(set(new_lvl)) == 200, f"found {len(set(new_lvl))}")
        rp.ok("2 Level Generation", "no duplicate level IDs",
              len(new_lvl) == len(set(new_lvl)),
              f"{len(new_lvl) - len(set(new_lvl))} duplicate(s)")
    if lay is None:
        rp.add("2 Level Generation", "pve_layout.xlsx readable", "FAIL", "file/sheet missing")
    else:
        rp.ok("2 Level Generation", "pve_layout carries the same 200 IDs",
              set(new_lay) == set(new_lvl),
              f"layout {len(set(new_lay))} vs level {len(set(new_lvl))}")

    # ── 3. Level objective validation (收集卡:1 vs card":14) ───────────────────
    log("  [3] Level objectives …")
    if lvl and lay:
        lmap = {r[0]: r[3] for r in lvl[1:] if isinstance(r[0], int) and len(r) > 3}
        ymap = {r[0]: r[2] for r in lay[1:] if isinstance(r[0], int) and len(r) > 2}
        bad = []
        for lid in sorted(set(new_lvl)):
            need = str(lmap.get(lid) or "").count("收集卡:1")
            if not need:
                continue
            got = str(ymap.get(lid) or "")
            have = got.count('card":14,') + got.count('card": 14,')
            if need != have:
                bad.append(f"{lid}: need {need} got {have}")
        rp.ok("3 Objective Validation",
              "收集卡 count matches the card\":14 count in every layout",
              not bad, f"{len(bad)} mismatch(es): " + "; ".join(bad[:5]))

    # ── 4-7. Star Pass: chest / item / store / notification ──────────────────
    log("  [4-7] Star Pass config …")
    chest = design / "chest.xlsx"
    for sheet, disp_col in (("chest_conditions", None), ("chest_rewards", 4)):
        rows = _rows_of(chest, sheet, max_col=6)
        if rows is None:
            rp.add("4 Chest", f"{sheet} readable", "FAIL", "file/sheet missing")
            continue
        got = {p: [r for r in rows[1:] if str(r[0] or "").startswith(f"{p}{new_ch:02d}")]
               for p in ("60", "70", "80")}
        for p in ("60", "70", "80"):
            rp.ok("4 Chest", f"{sheet}: {p}-series rows created", bool(got[p]), "none found")
        stale = [r[0] for rs in got.values() for r in rs
                 if any(f"章节{old_d}" in str(c) for c in r if c is not None)
                 or any(c == old_ch * 1000 for c in r)]
        rp.ok("4 Chest", f"{sheet}: no old-chapter reference in the new rows",
              not stale, f"{len(stale)} row(s) e.g. {stale[:3]}")
        if disp_col is not None:
            vals = {r[disp_col] for rs in got.values() for r in rs}
            rp.ok("4 Chest", "chest_rewards col E carries the new chapter number",
                  bool(vals) and all(str(v) == str(new_d) for v in vals),
                  f"got {sorted(map(str, vals))[:4]}")

    item = _rows_of(design / "item.xlsx", "item", max_col=2)
    if item is None:
        rp.add("5 Item", "item.xlsx readable", "FAIL", "file/sheet missing")
    else:
        ids = []
        for r in item[1:]:
            if not r or r[0] is None:
                continue
            try:
                ids.append(int(float(r[0])))
            except (TypeError, ValueError):
                continue
        chap = sorted(i for i in ids if 4000 <= i <= 4199)
        rp.ok("5 Item", "a 40-series row exists for the new chapter", bool(chap))
        def _item_id(marker):
            for r in item[1:]:
                if r and r[1] and marker in str(r[1]):
                    try:
                        return int(float(r[0]))
                    except (TypeError, ValueError):
                        return None
            return None
        new_id, prev_id = _item_id(f"章节{new_d}通行证"), _item_id(f"章节{old_d}通行证")
        rp.ok("5 Item", f"a row exists naming 章节{new_d}通行证", new_id is not None)
        rp.ok("5 Item", "col A is the previous chapter's number + 1",
              bool(new_id and prev_id) and new_id == prev_id + 1,
              f"new {new_id} vs previous {prev_id}")

    for sheet in ("store_recharge", "store_package"):
        rows = _rows_of(design / "store.xlsx", sheet)
        if rows is None:
            rp.add("6 Store", f"{sheet} readable", "FAIL", "file/sheet missing")
            continue
        chap = [r for r in rows[1:] if isinstance(r[0], int) and 40001 <= r[0] <= 40999]
        newr = [r for r in chap if any(f"章节{new_d}" in str(c) for c in r if c is not None)]
        rp.ok("6 Store", f"{sheet}: new 40-series row for 章节{new_d}", bool(newr))
        # The 40-series is not contiguous in production (40025 -> 40065 ->
        # 40079 -> 40084), so "no gaps" would be a false alarm. The checklist
        # only requires the new row to take the next number after the previous
        # chapter's row.
        prev_r = [r for r in chap if any(f"章节{old_d}" in str(c)
                                         for c in r if c is not None)]
        rp.ok("6 Store", f"{sheet}: new ID is the previous chapter's + 1",
              bool(newr and prev_r) and max(r[0] for r in newr) ==
              max(r[0] for r in prev_r) + 1,
              f"new {[r[0] for r in newr]} vs previous "
              f"{[r[0] for r in prev_r]}")
        stale = [r[0] for r in newr
                 if any(f"章节{old_d}" in str(c) for c in r if c is not None)]
        rp.ok("6 Store", f"{sheet}: no old-chapter reference in the new row",
              not stale, str(stale))

    noti = _rows_of(design / "notification.xlsx", "server_notification", max_col=9)
    if noti is None:
        rp.add("7 Notification", "notification.xlsx readable", "FAIL", "file/sheet missing")
    else:
        chap = [r for r in noti[1:] if len(r) > 1 and r[1] and "章节上线" in str(r[1])]
        new_rows = [r for r in chap if str(r[1]).startswith(f"{new_d}章节")]
        rp.ok("7 Notification", f"row added for {new_d}章节上线", bool(new_rows))
        if new_rows:
            r = new_rows[-1]
            rp.ok("7 Notification", "new chapter switch (col D) is True",
                  str(r[3]).strip().lower() == "true", f"got {r[3]!r}")
            on = [str(o[1]) for o in chap
                  if o is not r and str(o[3]).strip().lower() == "true"]
            rp.ok("7 Notification", "every previous chapter switch is False",
                  not on, f"still on: {on}")
            s, e = _parse_noti_dt(r[4]), _parse_noti_dt(r[5])
            rp.ok("7 Notification", "notification window is 3 days (start..start+2)",
                  bool(s and e) and (e - s).days == 2,
                  f"{s} → {e}" if s and e else "unparsable dates")
            prev = [o for o in chap if str(o[1]).startswith(f"{old_d}章节")]
            if prev:
                rp.ok("7 Notification", "PUSH_NOTE alternates from the previous chapter",
                      str(r[7]) != str(prev[-1][7]), f"both {r[7]}")
            rp.add("7 Notification", "start date is 1 day after the event go-live date",
                   "MANUAL",
                   f"currently {s.date() if s else '?'} — confirm against the roadmap")

    # ── 9-10. asset_ref + icon ───────────────────────────────────────────────
    log("  [9-10] Asset references …")
    ka = _rows_of(design / "asset_ref.xlsx", "key_asset", max_col=3)
    if ka is None:
        rp.add("9 asset_ref", "asset_ref.xlsx readable", "FAIL", "file/sheet missing")
    else:
        names = {str(r[0]): r for r in ka[1:] if r and r[0]}
        groups = {
            "MAP": [f"MAP_{new_d}_{i}" for i in range(1, 14)],
            "PVE_BG": [f"PVE_BG_{new_d}"],
            "CHAPTER_UNLOCK": [f"CHAPTER_UNLOCK_{new_d}"],
            "STARPASS_CHAPTER_BG": [f"STARPASS_CHAPTER_BG_{new_d}"],
            "STARPASS_ENTRY_ICON": [f"STARPASS_ENTRY_ICON_{new_d}"],
            "CHAPTER_TITLE_BG": [f"CHAPTER_TITLE_BG_{new_d}"],
        }
        for g, keys in groups.items():
            missing = [k for k in keys if k not in names]
            rp.ok("9 asset_ref", f"{g}: new chapter row(s) exist", not missing,
                  f"missing {missing[:4]}")
        stale = [k for keys in groups.values() for k in keys
                 if k in names and any(str(old_ch * 1000) in str(c) or f"_{old_d}" in str(c)
                                       for c in names[k][1:] if c)]
        rp.ok("9 asset_ref", "new rows carry no old chapter/map id in Alias or path",
              not stale, str(stale[:4]))
        # "Asset references point to the correct assets" — resolve them on disk.
        client = cfg.branch_root / "develop" / "client" / "Skipbo"
        missing_files, checked = [], 0
        for keys in groups.values():
            for k in keys:
                row = names.get(k)
                p = str(row[2] or "") if row and len(row) > 2 else ""
                if not p.startswith("Assets/"):
                    continue
                checked += 1
                if not (client / p).exists():
                    missing_files.append(p)
        rp.add("9 asset_ref", f"asset paths resolve on disk ({checked} checked)",
               "PASS" if not missing_files else "WARN",
               f"{len(missing_files)} missing, e.g. {missing_files[:3]}" if missing_files else "")

    icon = _rows_of(design / "icon.xlsx", "icon_chapter", max_col=3)
    if icon is None:
        rp.add("10 Icon", "icon.xlsx readable", "FAIL", "file/sheet missing")
    else:
        cn_new = _cn_num(new_d)
        blk = [r for r in icon[1:] if len(r) > 1 and r[1] and f"场景{cn_new}" in str(r[1])]
        rp.ok("10 Icon", "12 icon_chapter rows for the new chapter (4 fruits x 3 states)",
              len(blk) == 12, f"found {len(blk)}")
        fids = sorted({int(m.group(1)) for r in blk
                       for m in [re.match(r"FRUIT(\d+)_", str(r[0] or ""))] if m})
        rp.ok("10 Icon", "fruit numbering is sequential with no gaps",
              bool(fids) and fids == list(range(fids[0], fids[-1] + 1)), str(fids))
        states = collections.Counter(str(r[0]).rsplit("_", 1)[0] for r in blk)
        rp.ok("10 Icon", "every fruit has all 3 states",
              bool(states) and all(v == 3 for v in states.values()), str(dict(states)))

    # ── 11. pve_chapter ──────────────────────────────────────────────────────
    log("  [11] pve_chapter …")
    pvc = design / "pve_chapter.xlsx"
    ch = _rows_of(pvc, "chapter", max_col=17)
    if ch is None:
        rp.add("11 pve_chapter", "chapter sheet readable", "FAIL", "missing")
    else:
        row = next((r for r in ch[1:] if r[0] == new_ch * 1000), None)
        rp.ok("11 pve_chapter", "chapter row created", row is not None)
        if row:
            rp.ok("11 pve_chapter", "col C display number updated", row[2] == new_d, str(row[2]))
            rp.ok("11 pve_chapter", "cols B/D use the new Chinese numeral",
                  _cn_num(new_d) in str(row[1]) and _cn_num(new_d) in str(row[3]),
                  f"{row[1]!r} / {row[3]!r}")
            rp.ok("11 pve_chapter", "col G holds a 1/2/3 loop value",
                  bool(_nums(row[6])) and _nums(row[6])[-1] in (1, 2, 3), str(row[6]))
            e = _nums(row[4])
            rp.ok("11 pve_chapter", "col E is a 4-wide reward-zone range",
                  len(e) == 2 and e[1] - e[0] == 3, str(row[4]))
            stale = [i for i, c in enumerate(row)
                     if c is not None and (f"章节{old_d}" in str(c)
                                           or _cn_num(old_d) in str(c)
                                           or str(c) == str(old_ch * 1000))]
            rp.ok("11 pve_chapter", "no old-chapter leftovers in the chapter row",
                  not stale, f"columns {[chr(65 + i) for i in stale]}")

    cp = _rows_of(pvc, "chest_plan", max_col=5)
    if cp:
        new_plans = [r for r in cp[1:] if r[1] and f"第{new_d}章" in str(r[1])]
        rp.ok("11 pve_chapter", "chest_plan: all 3 StarPass variants (60/70/80)",
              len(new_plans) == 3, f"found {len(new_plans)}")
        stale = [r[0] for r in new_plans if f"章节{old_d}" in str(r[4] or "")]
        rp.ok("11 pve_chapter", "chest_plan: col E 章节 references updated",
              not stale, str(stale))

    mp = _rows_of(pvc, "map", max_col=10)
    if mp:
        newm = [r for r in mp[1:] if isinstance(r[0], int)
                and new_d * 1000 < r[0] < new_d * 1000 + 1000]
        oldm = [r for r in mp[1:] if isinstance(r[0], int)
                and old_d * 1000 < r[0] < old_d * 1000 + 1000]
        rp.ok("11 pve_chapter", "map: 13 rows for the new chapter",
              len(newm) == 13, f"found {len(newm)}")
        rp.ok("11 pve_chapter", "map: first row direction type is 0",
              bool(newm) and _nums(newm[0][9]) == [0], str(newm[0][9]) if newm else "-")
        tips = [r[0] for r in newm if str(r[9]).strip() == "2"]
        rp.ok("11 pve_chapter", "map: exactly one Tip (col J = 2) in the new chapter",
              len(tips) == 1, f"rows {tips}")
        old_tips = [r[0] for r in oldm if str(r[9]).strip() == "2"]
        rp.ok("11 pve_chapter", "map: the previous chapter no longer holds a Tip",
              not old_tips, f"rows {old_tips} still on 2")
        rp.ok("11 pve_chapter", "map: cols B/C use the new Chinese numeral",
              bool(newm) and all(_cn_num(new_d) in str(r[1]) and _cn_num(new_d) in str(r[2])
                                 for r in newm),
              "some rows still carry the old numeral")

    rz = _rows_of(pvc, "reward_zone", max_col=7)
    hv = _rows_of(pvc, "harvest", max_col=6)
    if rz and hv:
        newz = [r for r in rz[1:] if r[1] and f"章节{new_d}关卡" in str(r[1])]
        rp.ok("11 pve_chapter", "reward_zone: 4 zones for the new chapter",
              len(newz) == 4, f"found {len(newz)}")
        zids = sorted(r[0] for r in newz if isinstance(r[0], int))
        rp.ok("11 pve_chapter", "reward_zone: zone IDs sequential",
              bool(zids) and zids == list(range(zids[0], zids[-1] + 1)), str(zids))
        fruit_ids = [r[4] for r in newz if isinstance(r[4], int)]
        rp.ok("11 pve_chapter", "reward_zone: col E holds 3-digit fruit numbers",
              bool(fruit_ids) and all(99 < f < 1000 for f in fruit_ids), str(fruit_ids))
        hv_ids = [r[0] for r in hv[1:] if isinstance(r[0], int)]
        new_h = [r for r in hv[1:] if r[0] in fruit_ids]
        rp.ok("11 pve_chapter", "harvest: a row exists for every new fruit",
              len(new_h) == len(fruit_ids), f"{len(new_h)}/{len(fruit_ids)}")
        rp.ok("11 pve_chapter", "harvest: fruit IDs sequential with no gaps",
              bool(fruit_ids) and sorted(fruit_ids) == list(range(min(fruit_ids),
                                                                 max(fruit_ids) + 1))
              and len(hv_ids) == len(set(hv_ids)), str(fruit_ids))
        badatlas = [(r[0], r[5]) for r in new_h if str(r[5]) != f"MAP_FRUIT_{r[0] // 10}"]
        rp.ok("11 pve_chapter", "harvest: col F atlas matches fruit_id // 10",
              not badatlas, str(badatlas))
        badname = [r[0] for r in new_h
                   if not all(str(r[c]) == f"FRUIT{r[0]}_{c}" for c in (1, 2, 3))]
        rp.ok("11 pve_chapter", "harvest: cols B/C/D name the new fruit files",
              not badname, str(badname))

    # ── 19. Level balancing (pve_level, the new chapter's rows) ──────────────
    log("  [19] Level balancing …")
    if lvl and new_lvl:
        i_hands = _col_by_header(lvl_hdr, "限制回合数", default=6)
        i_diff = _col_by_header(lvl_hdr, "动态方案", default=12)
        i_ui = _col_by_header(lvl_hdr, "困难程度", default=14)
        i_star = _col_by_header(lvl_hdr, "星分数条件", default=15)
        i_max = _col_by_header(lvl_hdr, "星最大分数", default=16)
        body = [r for r in lvl[1:] if isinstance(r[0], int) and lo <= r[0] <= hi]

        bad = [r[0] for r in body
               if not (MIN_HANDS <= (_nums(r[i_hands]) or [0])[0] <= MAX_HANDS)]
        rp.ok("19 Balancing", f"col G hands within {MIN_HANDS}-{MAX_HANDS} for every level",
              not bad, f"{len(bad)} level(s) e.g. {bad[:5]}")

        bad = [r[0] for r in body if r[i_ui] is not None and str(r[i_ui]) != "None"
               and (_nums(r[i_ui]) or [9])[0] not in (0, 1, 2)]
        rp.ok("19 Balancing", "col O difficulty UI type is 0/1/2", not bad, f"e.g. {bad[:5]}")

        bad = [r[0] for r in body if r[i_diff] is not None and str(r[i_diff]) != "None"
               and str(r[i_diff]).strip() not in VALID_DIFFICULTY]
        rp.ok("19 Balancing", "col M difficulty text is a known value", not bad, f"e.g. {bad[:5]}")

        nostar = [r[0] for r in body if len(_nums(r[i_star])) != 3]
        rp.ok("19 Balancing", "col P holds three star thresholds (1★/2★/3★)",
              not nostar, f"{len(nostar)} level(s) e.g. {nostar[:5]}")
        notasc = [r[0] for r in body if len(_nums(r[i_star])) == 3
                  and _nums(r[i_star]) != sorted(set(_nums(r[i_star])))]
        rp.ok("19 Balancing", "star thresholds strictly increase (no overlap on the bar)",
              not notasc, f"e.g. {notasc[:5]}")

        badmax = []
        for r in body:
            st, mx = _nums(r[i_star]), _nums(r[i_max])
            if len(st) == 3 and mx and mx[0] != round(1.2 * st[2]):
                badmax.append((r[0], mx[0], round(1.2 * st[2])))
        rp.ok("19 Balancing", "col Q max score = round(1.2 x 3★ score)",
              not badmax, f"{len(badmax)} level(s) e.g. {badmax[:3]}")

        # U..AB must still carry the previous chapter's values for the matching
        # level — the generator re-rolls them and they have to be flattened back.
        i_u = _col_by_header(lvl_hdr, "星基础奖励$E4-E6", default=20)
        prev = {r[0]: r for r in lvl[1:] if isinstance(r[0], int)
                and old_ch * 1000 < r[0] <= old_ch * 1000 + 200}
        drift = []
        for r in body:
            p = prev.get(r[0] - 1000)
            if not p:
                continue
            for c in range(i_u, min(i_u + 8, len(r), len(p))):
                if str(r[c]) != str(p[c]):
                    drift.append(r[0])
                    break
        rp.add("19 Balancing", "cols U-AB match the previous chapter's level",
               "PASS" if not drift else "WARN",
               f"{len(drift)} level(s) differ, e.g. {drift[:5]}" if drift else "")

    # ── What a human still has to sign off ───────────────────────────────────
    for section, item in [
        ("2 Screenshots", "screenshots generated and uploaded for all 200 levels"),
        ("8 Unity Assets", "dynamic_map folder holds the real art (the tool copies placeholders)"),
        ("8 Unity Assets", "AdventureMapBg / ChapeDoneBg art replaced, Unity generated the metas"),
        ("12 Map Builder", "scene duplicated, fruits at scale 1.5, no fruit overlapping a level number"),
        ("13 Fruit Position", "map_data_temp cols E/F/G copied into PVE_Chapter map G/F/H"),
        ("14 Unity Import", "coder_convert.bat ran and the levels appear in the editor"),
        ("16 Simulation QA", "error file reviewed, every level has a result, no unexpected -1 deciles"),
        ("18 Balancing", "hard levels 7-12% win rate, nightmare levels 5-10%"),
        ("19 Balancing", "hard/nightmare alternate with roughly 6-7 levels between them"),
    ]:
        rp.add(section, item, "MANUAL")

    # ── Report ───────────────────────────────────────────────────────────────
    out_path = design / f"qa_validation_chapter_{new_ch}.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "QA Validation"
    hdr_fill = PatternFill("solid", fgColor="1a3a5c")
    hdr_font = Font(bold=True, color="F0D060")
    for c, t in enumerate(["Section", "Check", "Status", "Detail"], 1):
        cell = ws.cell(row=1, column=c, value=t)
        cell.fill, cell.font = hdr_fill, hdr_font
    fills = {"FAIL": PatternFill("solid", fgColor="5c1a1a"),
             "WARN": PatternFill("solid", fgColor="5c3a1a"),
             "MANUAL": PatternFill("solid", fgColor="24304a")}
    for i, (sec, chk, st, det) in enumerate(rp.rows, start=2):
        for c, v in enumerate((sec, chk, st, det), 1):
            cell = ws.cell(row=i, column=c, value=v)
            if st in fills:
                cell.fill = fills[st]
    for col, w in zip("ABCD", (22, 62, 9, 60)):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A2"
    wb.save(str(out_path))

    p, f, w, m = rp.count("PASS"), rp.count("FAIL"), rp.count("WARN"), rp.count("MANUAL")
    log("")
    log(f"  QA: {p} passed, {f} failed, {w} warning(s), {m} manual check(s)")
    log(f"  ✓ Report saved: {out_path.name}")
    if f:
        log(f"  ✗ {f} check(s) FAILED — the chapter is not ready to sign off.")
    else:
        log("  ✓ Every automated check passed. Work through the MANUAL rows before sign-off.")
    return str(out_path)

# ══════════════════════════════════════════════════════════════════════════════
#  Steps that drive the external toolchain
# ══════════════════════════════════════════════════════════════════════════════
# These shell out to the SkipBo project scripts. All of them run on the
# bundled openpyxl-only interpreter EXCEPT step 1: conf_fill_ex.py imports
# pandas/numpy, so it needs a real system Python (see _python_exe).


def check_level_conflicts(cfg, log):
    """Raise if the new chapter's level ids already exist.

    Checks the workspace selected in the app, which is where Magic Builder now
    writes: _run_project_script redirects the design_path those scripts
    hardcode (see _REDIRECT_LAUNCHER). Before that redirect existed this
    check and the script disagreed - it passed on the selected workspace while
    the script raised a conflict from WorkSpcae3.

    Split out of run_step1_magic_builder so the preflight can run it before
    any step writes, and so it is testable without driving the real Magic
    Builder against real project files.
    """
    new_first = cfg.new_ch * 1000 + 1
    new_last = cfg.new_ch * 1000 + 200
    target = cfg.design()

    if not Path(target).is_dir():
        log(f"  ⚠  Selected design folder does not exist: {target}", "WARN")
        return

    # Each workbook keeps its ids in its own sheet: pve_layout.xlsx has no
    # 'level' sheet at all, so falling back to the first sheet only happened
    # to read the right one.
    for name, sheet_name in (("pve_level.xlsx", "level"), ("pve_layout.xlsx", "layout")):
        xlsx_path = Path(target) / name
        if not xlsx_path.exists():
            log(f"  ⚠  {name} not found in {target} - skipping its pre-check", "WARN")
            continue
        try:
            import openpyxl as _oxl
            wb = _oxl.load_workbook(str(xlsx_path), read_only=True, data_only=True)
            if sheet_name not in wb.sheetnames:
                sheets = ", ".join(wb.sheetnames)
                wb.close()
                log(f"  ⚠  {name} has no '{sheet_name}' sheet - skipping pre-check "
                    f"(sheets: {sheets})", "WARN")
                continue
            sheet = wb[sheet_name]
            keys = {row[0].value for row in sheet.iter_rows(min_row=2)
                    if isinstance(row[0].value, int)}
            wb.close()
            conflict = keys & set(range(new_first, new_last + 1))
            if conflict:
                raise RuntimeError(
                    f"Chapter {cfg.disp_new} (internal {cfg.new_ch}) levels "
                    f"already exist in {name}.\n\n"
                    f"Found {len(conflict)} conflicting level(s) "
                    f"({min(conflict)}-{max(conflict)}) in:\n  {xlsx_path}\n\n"
                    f"Magic Builder has already been run for this chapter in "
                    f"this workspace. To re-run it, first remove rows "
                    f"{new_first}-{new_last} from the '{sheet_name}' sheet of "
                    f"pve_level.xlsx and pve_layout.xlsx in:\n  {target}\n\n"
                    f"If that chapter was genuinely already built, it needs no "
                    f"run - check before deleting anything."
                )
        except RuntimeError:
            raise
        except Exception as e:
            log(f"  ⚠  Could not pre-check {name}: {e}", "WARN")


def run_step1_magic_builder(cfg: AppConfig, log):
    """Update CondFilterConf.json and run conf_fill_ex.py directly."""
    conf_path = cfg.lvl_py / "CondFilterConf.json"
    if not conf_path.exists():
        raise FileNotFoundError(f"CondFilterConf.json not found at:\n{conf_path}")

    check_level_conflicts(cfg, log)

    log(f"Loading {conf_path.name}…")
    raw = conf_path.read_text("utf-8")
    data = json.loads(raw)

    # Add old chapter range to layout_diff (exclude from source pool)
    old_range = [cfg.old_ch * 1000 + 1, cfg.old_ch * 1000 + 200]
    layout_diff = data.get("layout_diff", [])
    if old_range not in layout_diff:
        layout_diff.append(old_range)
        data["layout_diff"] = layout_diff
        log(f"  ✓ Added {old_range} to layout_diff")
    else:
        log(f"  – {old_range} already in layout_diff")

    # Set new chapter range
    new_range = [[cfg.new_ch * 1000 + 1, cfg.new_ch * 1000 + 200]]
    data["new_level_range"] = new_range
    log(f"  ✓ Set new_level_range → {new_range}")

    conf_path.write_text(json.dumps(data, indent=4, ensure_ascii=False), "utf-8")
    log("  ✓ CondFilterConf.json saved")

    # Run the Python script
    py_script = cfg.lvl_py / "conf_fill_ex.py"
    py_exe = _resolve_interpreter(py_script, log)
    log(f"\nRunning conf_fill_ex.py -c CondFilterConf.json …")
    result = _run_project_script(py_exe, py_script,
                                 ["-c", "CondFilterConf.json"], cwd=cfg.lvl_py,
                                 design_override=cfg.design(), log=log)
    if result.stdout: log(result.stdout.strip())
    if result.stderr: log(f"[STDERR] {result.stderr.strip()}")
    if result.returncode != 0:
        raise _script_failure("conf_fill_ex.py", result, py_exe, py_script, cfg)
    log("  ✓ Magic Builder completed successfully")

def run_step2_upload(cfg: AppConfig, log):
    """Generate capture index pages and commit screenshots + output to SVN."""
    py_dir = cfg.lvl_py
    workspace = py_dir.parent.parent / "__WorkSpace"
    scr_path = workspace / "__ScreenShot"
    out_path = workspace / "__Output"

    # Fail loudly rather than reporting success having committed nothing: svn
    # returns 1 for "not a working copy", which the loop below treats as benign,
    # so without these guards a wrong Level Builder path looks like a clean run.
    if not py_dir.is_dir():
        raise FileNotFoundError(
            f"Level Builder Py folder not found: {py_dir} — set it in "
            f"Settings under 'Level Builder (Magic Builder) Py folder'.")
    missing = [n for n in ("gen_captures.py", "gen_captures_others.py")
               if not (py_dir / n).exists()]
    if missing:
        raise FileNotFoundError(
            f"{', '.join(missing)} not found in {py_dir} — that is not the "
            f"Level Builder Py folder.")
    if not workspace.is_dir():
        raise FileNotFoundError(
            f"__WorkSpace not found: {workspace} — capture the screenshots "
            f"in Unity first.")
    if not scr_path.is_dir():
        raise FileNotFoundError(
            f"No screenshots to upload: {scr_path} does not exist.")

    def svn(args, label):
        cmd = ["svn"] + args
        log(f"  svn {' '.join(args[:2])}…")
        r = subprocess.run(cmd, capture_output=True, text=True,
                           encoding="utf-8", errors="replace", timeout=60)
        if r.stdout.strip(): log(r.stdout.strip())
        if r.stderr.strip(): log(f"  [svn] {r.stderr.strip()}")
        if r.returncode not in (0, 1):  # svn returns 1 on "nothing to do"
            log(f"  ⚠  {label} exited {r.returncode} (continuing)", "WARN")

    # 1. SVN update workspace
    log("Step 1 / 4 — SVN update workspace…")
    svn(["up", str(workspace), "--force"], "svn up")

    # 2. Generate HTML index pages from screenshots
    log("Step 2 / 4 — Generating capture index pages…")
    for script_name in ("gen_captures.py", "gen_captures_others.py"):
        script = py_dir / script_name
        if not script.exists():
            log(f"  ⚠  {script_name} not found — skipping", "WARN")
            continue
        py_exe = _resolve_interpreter(script, log)
        r = _run_project_script(py_exe, script, cwd=py_dir, timeout=60)
        if r.stdout.strip(): log(r.stdout.strip())
        if r.stderr.strip(): log(f"  [STDERR] {r.stderr.strip()}")
        if r.returncode != 0:
            raise _script_failure(script_name, r, py_exe, script)
    log("  ✓ Index pages generated")

    # 3. SVN add new files
    log("Step 3 / 4 — SVN add new files…")
    for path in (scr_path, out_path):
        if path.exists():
            svn(["add", str(path), "--force"], f"svn add {path.name}")

    # 4. SVN commit
    log("Step 4 / 4 — SVN commit…")
    for path in (scr_path, out_path):
        if path.exists():
            svn(["ci", str(path), "-m", "自动上传库存与截图"], f"svn ci {path.name}")

    log("  ✓ Screenshots committed to SVN")

def run_step7_map_builder(cfg: AppConfig, scene_file: str, log):
    """Parse Unity scene YAML, extract fruit/map positions, run save_map_data.py."""
    scenes_dir = cfg.map_scenes
    # Find scene file
    if scene_file:
        scene_path = Path(scene_file)
    else:
        unity_files = list(scenes_dir.glob("*.unity"))
        # Prefer scene with latest chapter number
        unity_files.sort(key=lambda p: p.name, reverse=True)
        scene_path = unity_files[0] if unity_files else None
    if not scene_path or not scene_path.exists():
        raise FileNotFoundError(f"No .unity scene found in {scenes_dir}")

    log(f"Parsing scene: {scene_path.name} …")
    with open(str(scene_path), "r", encoding="utf-8", errors="replace") as f:
        scene_text = f.read()

    # Parse GameObjects with their RectTransforms
    # We need: FruitPanel children (fruits) and LevelPanel children (level buttons)
    # Each "--- !u!1 &{id}" block is a GameObject
    # We build a map: fileID -> {name, children, parent, anchored_pos, scale}

    go_pattern    = re.compile(r'--- !u!1 &(\d+)\nGameObject:(.*?)(?=\n--- |$)', re.DOTALL)
    rt_pattern    = re.compile(r'--- !u!224 &(\d+)\nRectTransform:(.*?)(?=\n--- |$)', re.DOTALL)

    go_map = {}   # fileID -> {name, components, children_from_rt}
    rt_map = {}   # fileID -> {go_file_id, anchored, scale, children, parent, name}

    for m in go_pattern.finditer(scene_text):
        fid = m.group(1)
        body = m.group(2)
        name_m = re.search(r'm_Name: (.+)', body)
        comp_m = re.findall(r'component: \{fileID: (\d+)\}', body)
        go_map[fid] = {
            "name": name_m.group(1).strip() if name_m else "",
            "comps": comp_m,
        }

    for m in rt_pattern.finditer(scene_text):
        fid = m.group(1)
        body = m.group(2)
        go_m   = re.search(r'm_GameObject: \{fileID: (\d+)\}', body)
        ap_m   = re.search(r'm_AnchoredPosition: \{x: ([\d.\-]+), y: ([\d.\-]+)\}', body)
        sc_m   = re.search(r'm_LocalScale: \{x: ([\d.\-]+), y: ([\d.\-]+)', body)
        ch_m   = re.findall(r'\{fileID: (\d+)\}', re.search(r'm_Children:(.*?)m_Father', body, re.DOTALL).group(1) if re.search(r'm_Children:(.*?)m_Father', body, re.DOTALL) else "")
        pa_m   = re.search(r'm_Father: \{fileID: (\d+)\}', body)

        rt_map[fid] = {
            "go_id":    go_m.group(1) if go_m else None,
            "ax": float(ap_m.group(1)) if ap_m else 0.0,
            "ay": float(ap_m.group(2)) if ap_m else 0.0,
            "sx": float(sc_m.group(1)) if sc_m else 1.0,
            "sy": float(sc_m.group(2)) if sc_m else 1.0,
            "children": ch_m,
            "parent":   pa_m.group(1) if pa_m else None,
        }

    # Build name lookup: go_id -> name
    go_name = {fid: go_map[fid]["name"] for fid in go_map}
    # Reverse: rt_fid -> go_name
    rt_name = {fid: go_name.get(rt_map[fid]["go_id"], "") for fid in rt_map}

    def get_children_named(parent_rt_id):
        return rt_map[parent_rt_id]["children"] if parent_rt_id in rt_map else []

    # Locate MapItemRoot panels (top-level map blocks)
    # Each map block has a FruitPanel and LevelPanel child
    fruit_panels = {fid for fid, nm in rt_name.items() if nm == "FruitPanel"}
    level_panels = {fid for fid, nm in rt_name.items() if nm == "LevelPanel"}

    if not fruit_panels:
        raise ValueError("No FruitPanel found in scene — is this the correct scene file?")

    log(f"  Found {len(fruit_panels)} FruitPanel(s), {len(level_panels)} LevelPanel(s)")

    # For each map block (sibling panels share a parent)
    map_blocks = []  # list of (level_data, fruit_data, dir_data)

    # Group FruitPanels by their parent (= map block RT)
    parent_to_panels = {}
    for fp_id in fruit_panels:
        p = rt_map[fp_id]["parent"]
        if p not in parent_to_panels:
            parent_to_panels[p] = {}
        parent_to_panels[p]["fruit"] = fp_id
    for lp_id in level_panels:
        p = rt_map[lp_id]["parent"]
        if p in parent_to_panels:
            parent_to_panels[p]["level"] = lp_id

    log(f"  Parsing {len(parent_to_panels)} map blocks …")

    level_pos_parts = []
    fruit_id_parts  = []
    fruit_pos_parts = []
    fruit_sc_parts  = []
    dir_parts       = []
    map_parts       = []

    for block_id, panels in sorted(parent_to_panels.items(), reverse=True):
        fp_id = panels.get("fruit")
        lp_id = panels.get("level")

        # Level positions
        level_buttons = get_children_named(lp_id) if lp_id else []
        lev_count = len(level_buttons)
        lev_poses = []
        min_lvl_id = None
        for btn_id in level_buttons:
            rt = rt_map.get(btn_id, {})
            lev_poses.append(f"{round(rt.get('ax',0))},{-round(rt.get('ay',0))}")
            btn_name = rt_name.get(btn_id, "")
            nums = re.findall(r'\d+', btn_name)
            if nums:
                lvl_id = int(nums[-1])
                if min_lvl_id is None or lvl_id < min_lvl_id:
                    min_lvl_id = lvl_id
        level_pos_parts.append(f"{lev_count}|{'*'.join(lev_poses)}")
        if min_lvl_id:
            map_parts.append(f"章节N关卡{min_lvl_id},章节N关卡{min_lvl_id + lev_count - 1}")

        # Fruit positions
        fruit_children = get_children_named(fp_id) if fp_id else []
        fr_count = len(fruit_children)
        fr_poses, fr_scales, fr_ids = [], [], []
        for fr_id in fruit_children:
            rt = rt_map.get(fr_id, {})
            fr_poses.append(f"{round(rt.get('ax',0))},{-round(rt.get('ay',0))}")
            fr_scales.append(f"{abs(rt.get('sx',1.5))},{abs(rt.get('sy',1.5))}")
            # Fruit ID from name like "fruit322 (3)" → 322
            fn = rt_name.get(fr_id, "")
            id_m = re.search(r'fruit(\d+)', fn, re.IGNORECASE)
            fr_ids.append(f"{id_m.group(1)}," if id_m else "0,")
        fruit_pos_parts.append(f"{fr_count}|{'*'.join(fr_poses)}")
        fruit_sc_parts.append(f"{fr_count}|{'*'.join(fr_scales)}")
        fruit_id_parts.append(f"{fr_count}|{''.join(fr_ids)}")

        # Direction markers
        blk_go = go_map.get(rt_map.get(block_id, {}).get("go_id", ""), {})
        dir_str = ""
        for rt_child_id in rt_map.get(block_id, {}).get("children", []):
            nm = rt_name.get(rt_child_id, "")
            rt = rt_map.get(rt_child_id, {})
            if nm == "CellNextLevel2":
                dir_str += f"#1|{round(rt.get('ax',0))},{round(rt.get('ay',0))}"
            elif nm == "CellNextLevel1":
                dir_str += f"#0|{round(rt.get('ax',0))},{round(rt.get('ay',0))}"
            elif nm == "CellNextComingSoon":
                dir_str += f"#2|{round(rt.get('ax',0))},{round(rt.get('ay',0))}"
        dir_parts.append(dir_str or "#")

    n_blocks = len(parent_to_panels)
    level_data  = f"{n_blocks}#" + "#".join(level_pos_parts)
    fruit_id    = "#".join(fruit_id_parts)
    fruit_pos   = "#".join(fruit_pos_parts)
    fruit_scale = "#".join(fruit_sc_parts)
    dir_data    = "#".join(dir_parts)
    map_data    = f"{n_blocks}#" + "#".join(map_parts)

    log(f"  Extracted {n_blocks} map blocks")

    # Run save_map_data.py
    save_py = cfg.map_py / "save_map_data.py"
    if not save_py.exists():
        raise FileNotFoundError(f"save_map_data.py not found: {save_py}")

    py_exe = _resolve_interpreter(save_py, log)
    log("Running save_map_data.py …")
    result = _run_project_script(
        py_exe, save_py,
        ["map_data_temp.xlsx", level_data, fruit_id, fruit_pos, fruit_scale, dir_data, map_data],
        cwd=cfg.map_py)
    if result.stdout: log(result.stdout.strip())
    if result.stderr: log(f"[STDERR] {result.stderr.strip()}")
    if result.returncode != 0:
        raise _script_failure("save_map_data.py", result, py_exe, save_py)
    log(f"  ✓ map_data_temp.xlsx written to {cfg.map_output}")

    # Auto-copy columns from map_data_temp → pve_chapter.map sheet
    map_temp_path = cfg.map_output / "map_data_temp.xlsx"
    pvc_path = cfg.design() / "pve_chapter.xlsx"
    if map_temp_path.exists() and pvc_path.exists():
        log("Copying map positions to pve_chapter.map sheet …")
        wb_tmp = openpyxl.load_workbook(str(map_temp_path), read_only=True, data_only=True)
        tmp_rows = list(wb_tmp["map_data"].iter_rows(min_row=2, values_only=True))
        wb_tmp.close()

        wb_pvc = openpyxl.load_workbook(str(pvc_path))
        ws_map = wb_pvc["map"]
        map_rows = list(ws_map.iter_rows(min_row=2, values_only=True))
        new_ch_rows = [i+2 for i, r in enumerate(map_rows)
                       if r[0] and str(new_ch*1000) in str(r[0])]
        if new_ch_rows:
            for i, row_num in enumerate(new_ch_rows):
                if i < len(tmp_rows):
                    tmp = tmp_rows[i]
                    # E→G, F→F, G→H in pve_chapter map
                    if len(tmp) > 4: ws_map.cell(row=row_num, column=7, value=tmp[4])  # E
                    if len(tmp) > 5: ws_map.cell(row=row_num, column=6, value=tmp[5])  # F
                    if len(tmp) > 6: ws_map.cell(row=row_num, column=8, value=tmp[6])  # G
            log(f"  ✓ Pasted map positions into {len(new_ch_rows)} pve_chapter rows")
        else:
            log("  ⚠ No new chapter rows found in pve_chapter.map to paste into")
        wb_pvc.save(str(pvc_path))
        log("  ✓ pve_chapter.xlsx map data updated")

def run_step8_import_config(cfg: AppConfig, log):
    """Run coder_convert.bat for the active workspace."""
    bat = cfg.design_root() / "coder_convert.bat"
    if not bat.exists():
        raise FileNotFoundError(f"coder_convert.bat not found: {bat}")
    log(f"Running coder_convert.bat in {cfg.workspace} …")
    result = subprocess.run(
        str(bat), shell=True,
        cwd=str(bat.parent),
        capture_output=True, text=True, encoding="utf-8", errors="replace"
    )
    if result.stdout: log(result.stdout.strip())
    if result.stderr: log(f"[STDERR] {result.stderr.strip()}")
    if result.returncode != 0:
        raise RuntimeError(f"coder_convert.bat exited with code {result.returncode}")
    log("  ✓ Config import completed")

def run_step9_simulation(cfg: AppConfig, log):
    """Update run_conf_2.json and start simulation."""
    sim_dir = cfg.sim_path
    conf_path = sim_dir / "run_conf_2.json"
    if not conf_path.exists():
        raise FileNotFoundError(f"run_conf_2.json not found: {conf_path}")

    log("Updating run_conf_2.json …")
    raw = conf_path.read_text("utf-8")
    data = json.loads(raw)
    data["levels"] = [[cfg.new_ch * 1000 + 1, cfg.new_ch * 1000 + 200]]
    data["dir_name"] = f"Chapter_{cfg.new_ch}"
    data["repeat_num"] = int(cfg.sim_repeat)
    data["p_num"] = int(cfg.sim_p_num)
    tag = str(cfg.sim_tag).strip()
    data["tags"] = [tag] if tag else [""]
    conf_path.write_text(json.dumps(data, indent=4, ensure_ascii=False), "utf-8")
    log(f"  ✓ Levels set to [{cfg.new_ch*1000+1}, {cfg.new_ch*1000+200}]")
    log(f"  ✓ dir_name = Chapter_{cfg.new_ch}")

    bat = sim_dir / "start_run_2.bat"
    if not bat.exists():
        raise FileNotFoundError(f"start_run_2.bat not found: {bat}")
    log("Launching simulation (this may take several minutes) …")
    proc = subprocess.Popen(
        str(bat), shell=True, cwd=str(sim_dir),
        stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
        text=True, encoding="utf-8", errors="replace"
    )
    for line in proc.stdout:
        stripped = line.rstrip()
        if stripped:
            log(stripped)
    proc.wait()
    if proc.returncode != 0:
        raise RuntimeError(f"Simulation exited with code {proc.returncode}")
    log("  ✓ Simulation completed")
