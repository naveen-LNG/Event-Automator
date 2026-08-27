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
