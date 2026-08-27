"""Checks for the borrowed-interpreter resolver in chapter_core.

Run: python test_interpreter_resolve.py

Covers the failure that prompted it: Chapter Update step 1 died because the
system Python it borrowed had no pandas, and the message did not say so.
"""
import os
import sys
import tempfile
from pathlib import Path

import chapter_core as cc

FAILS = []


def check(cond, label, detail=""):
    if cond:
        print(f"OK   {label}")
    else:
        FAILS.append(label)
        print(f"FAIL {label}" + (f"  ({detail})" if detail else ""))


def write(d, name, body):
    p = Path(d) / name
    p.write_text(body, encoding="utf-8")
    return p


def quiet(msg, level=None):
    pass


def main():
    with tempfile.TemporaryDirectory() as d:
        # --- _script_imports: stdlib and siblings must not be reported ---
        write(d, "sibling_mod.py", "x = 1\n")
        p = write(d, "a.py", "import os, sys, json\nimport sibling_mod\nimport openpyxl\n")
        got = cc._script_imports(p)
        check(got == {"openpyxl"}, "stdlib + sibling imports are excluded", f"got {got}")

        p = write(d, "b.py", "from pandas import DataFrame\nfrom . import thing\nimport numpy.linalg\n")
        got = cc._script_imports(p)
        check(got == {"pandas", "numpy"},
              "from-imports and dotted imports reduce to root package", f"got {got}")

        p = write(d, "c.py", "this is not valid python (\n")
        check(cc._script_imports(p) is None,
              "unparseable script returns None rather than claiming to know")

        # --- _missing_modules against the interpreter we are running on ---
        check(cc._missing_modules(sys.executable, {"os", "sys"}) == set(),
              "stdlib modules are not reported missing")
        check(cc._missing_modules(sys.executable, {"definitely_absent_pkg_zzz"}) ==
              {"definitely_absent_pkg_zzz"},
              "an absent module is reported missing")
        check(cc._missing_modules(sys.executable, set()) == set(),
              "empty request needs no subprocess")

        # --- _resolve_interpreter ---
        p = write(d, "stdlib_only.py", "import os\n")
        check(cc._resolve_interpreter(p, quiet) == sys.executable,
              "a stdlib-only script runs on the bundled interpreter")

        p = write(d, "impossible.py", "import definitely_absent_pkg_zzz\n")
        try:
            cc._resolve_interpreter(p, quiet)
            check(False, "an unsatisfiable script raises")
        except RuntimeError as e:
            msg = str(e)
            check("definitely_absent_pkg_zzz" in msg, "error names the missing package")
            check("-m pip install" in msg, "error gives a runnable pip command")
            check(sys.executable in msg, "error names the exact interpreter, not 'python'")

        # --- sibling imports must work on ANY interpreter -----------------
        # Regression: the bundled interpreter's ._pth makes sys.path explicit
        # and does NOT add the script's own folder, so running an external
        # script with it broke every sibling import (conf_fill_ex.py importing
        # lb_helper). PYTHONPATH is ignored too, so the launcher has to insert
        # the folder itself.
        proj = Path(d) / "proj"
        proj.mkdir()
        write(proj, "helper_mod.py", "MARK = 'helper ok'\n")
        write(proj, "tool.py",
              "import sys, helper_mod\n"
              "print('args:', sys.argv[1:])\n"
              "print(helper_mod.MARK)\n")
        r = cc._run_project_script(sys.executable, proj / "tool.py", ["-c", "conf.json"], cwd=proj)
        check(r.returncode == 0, "external script runs", (r.stderr or "")[:120])
        check("helper ok" in (r.stdout or ""), "script can import its sibling module")
        check("['-c', 'conf.json']" in (r.stdout or ""),
              "arguments reach the script unchanged", (r.stdout or "").strip())

        # --- a missing sibling is not reported as a pip problem -----------
        class Res:
            returncode = 1
            stderr = "ModuleNotFoundError: No module named 'helper_mod'"

        msg = str(cc._script_failure("tool.py", Res(), "X:/py.exe", proj / "tool.py"))
        check("one of its own files" in msg,
              "a sibling module that fails to import is called a path problem")
        check("pip install" not in msg,
              "no pip advice for a module that is part of the script itself")

        class Res2:
            returncode = 1
            stderr = "ModuleNotFoundError: No module named 'scipy'"

        msg = str(cc._script_failure("tool.py", Res2(), "X:/py.exe", proj / "tool.py"))
        check("pip install scipy" in msg, "a real third-party gap still gets pip advice")

        # --- known project-script errors get translated ------------------
        class Refill:
            returncode = 1
            stderr = ('  File "conf_fill_ex.py", line 252, in refill_spec\n'
                      '    raise RuntimeError("refill level id conflict: " + str(levelid))\n'
                      'RuntimeError: refill level id conflict: 95001\n')

        class FakeCfg:
            display_offset = 10

        msg = str(cc._script_failure("conf_fill_ex.py", Refill(), "X:/py.exe",
                                     proj / "tool.py", FakeCfg()))
        check("chapter 85" in msg,
              "refill conflict is reported as the DISPLAY chapter, not the internal one",
              msg.split("\n")[2] if len(msg.split("\n")) > 2 else msg)
        check("95001-95200" in msg, "refill conflict names the level range")
        check("exited with code" not in msg,
              "refill conflict replaces the bare exit-code message")
        # Without cfg it must not guess a display number.
        msg2 = str(cc._script_failure("conf_fill_ex.py", Refill(), "X:/py.exe"))
        check("internal chapter 95" in msg2 and "chapter 85" not in msg2,
              "without cfg it says internal chapter rather than inventing one")

        # --- sibling package directories count as local too ---------------
        (proj / "pkg_dir").mkdir()
        p = write(proj, "uses_pkg.py", "import pkg_dir\nimport pandas\n")
        check(cc._script_imports(p) == {"pandas"},
              "a sibling package directory is excluded from required packages")

        # --- the design folder is read from the script, not assumed ---------
        # conf_fill_ex.py hardcodes ../../../../../trunk/.../WorkSpcae3/design
        # and ignores the workspace chosen in the app, which is why checking
        # cfg.design() passed clean while the script itself found a conflict.
        lb = Path(d) / "lb" / "Py"
        lb.mkdir(parents=True)
        dsn = Path(d) / "lb" / "dsn"
        dsn.mkdir()
        write(lb, "conf_fill_ex.py",
              'import os\ndesign_path = os.path.abspath("../dsn")\n')
        got = cc._script_design_path(lb / "conf_fill_ex.py", lb)
        check(cc._same_dir(got, dsn),
              "design path is resolved from the script's own literal", f"got {got}")
        check(cc._script_design_path(lb / "nope.py", lb) is None,
              "an unreadable script yields no design path")
        write(lb, "no_path.py", "x = 1\n")
        check(cc._script_design_path(lb / "no_path.py", lb) is None,
              "a script without the idiom yields no design path")

        # --- level-conflict pre-check --------------------------------------
        import openpyxl

        def make_book(path, sheet, ids):
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = sheet
            ws.append(["id"])
            for i in ids:
                ws.append([i])
            wb.save(str(path))

        make_book(dsn / "pve_level.xlsx", "level", range(94001, 94201))
        make_book(dsn / "pve_layout.xlsx", "layout", range(94001, 94201))

        class LvlCfg:
            display_offset = 10
            lvl_py = lb

            def __init__(self, new):
                self._n = new

            @property
            def new_ch(self):
                return self._n

            @property
            def disp_new(self):
                return self._n - self.display_offset

            def design(self):
                return dsn

        cc.check_level_conflicts(LvlCfg(95), quiet)
        check(True, "a chapter absent from the workbooks passes the pre-check")

        try:
            cc.check_level_conflicts(LvlCfg(94), quiet)
            check(False, "an already-built chapter is blocked")
        except RuntimeError as e:
            msg = str(e)
            check("Chapter 84 (internal 94)" in msg,
                  "conflict names display and internal chapter", msg.splitlines()[0])
            check("94001-94200" in msg, "conflict names the level range")

        # pve_layout.xlsx has no 'level' sheet, so reading wb.active instead of
        # the named sheet only ever worked by accident.
        make_book(dsn / "pve_layout.xlsx", "layout_z", range(94001, 94201))
        (dsn / "pve_level.xlsx").unlink()
        warns = []
        cc.check_level_conflicts(LvlCfg(94), lambda m, lvl=None: warns.append(str(m)))
        check(any("no 'layout' sheet" in w for w in warns),
              "a workbook missing the expected sheet warns instead of guessing",
              "; ".join(warns)[:120])

        # --- preflight only inspects the scripts it actually runs ---
        lvl = Path(d) / "lvl"; lvl.mkdir()
        mp = Path(d) / "map"; mp.mkdir()
        write(lvl, "conf_fill_ex.py", "import os\n")
        write(lvl, "unrelated_tool.py", "import definitely_absent_pkg_zzz\n")

        class Cfg:
            lvl_py = lvl
            map_py = mp

        check(cc.preflight_project_scripts(Cfg(), quiet) == [],
              "preflight ignores scripts the tool never runs")

        write(lvl, "gen_captures.py", "import definitely_absent_pkg_zzz\n")
        probs = cc.preflight_project_scripts(Cfg(), quiet)
        check(len(probs) == 1 and "gen_captures.py" in probs[0],
              "preflight reports a real dependency gap before any step runs")

    print("")
    if FAILS:
        print(f"{len(FAILS)} CHECK(S) FAILED")
        sys.exit(1)
    print("INTERPRETER RESOLVE OK")


if __name__ == "__main__":
    main()
