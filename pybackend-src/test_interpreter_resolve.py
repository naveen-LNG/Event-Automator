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
              "the path a script hardcodes can still be read (diagnostics)",
              f"got {got}")
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

        # --- design_path redirect -----------------------------------------
        # conf_fill_ex.py and the conf_complex_gen.py it imports both hardcode
        # design_path, so Magic Builder used to ignore the chosen workspace.
        rd = Path(d) / "rd"
        (rd / "Py").mkdir(parents=True)
        (rd / "chosen").mkdir()
        write(rd / "Py", "conf_complex_gen.py",
              'import os\ndesign_path = os.path.abspath("../hardcoded")\n'
              'def where(): return design_path\n')
        write(rd / "Py", "conf_fill_ex.py",
              'import os\ndesign_path = os.path.abspath("../hardcoded")\n'
              'import conf_complex_gen\n'
              'print("ENTRY:" + design_path)\n'
              'print("SIB:" + conf_complex_gen.where())\n'
              'print("CWD:" + os.getcwd())\n')
        r = cc._run_project_script(sys.executable, rd / "Py" / "conf_fill_ex.py",
                                   cwd=rd / "Py", design_override=rd / "chosen",
                                   log=quiet)
        out = r.stdout or ""
        check(r.returncode == 0, "redirected script runs", (r.stderr or "")[:150])
        check("ENTRY:" + str(rd / "chosen") in out,
              "entry script's design_path points at the chosen workspace")
        check("SIB:" + str(rd / "chosen") in out,
              "imported sibling's design_path is redirected too")
        check("CWD:" + str(rd / "Py") in out,
              "working directory stays the real script folder, so its other "
              "relative paths still resolve")
        check("hardcoded" not in out,
              "the hardcoded workspace is not used anywhere")
        # The originals must never be touched - they are shared SVN files.
        original = (rd / "Py" / "conf_fill_ex.py").read_text(encoding="utf-8")
        check('os.path.abspath("../hardcoded")' in original,
              "the script on disk is left unmodified")

        # A script without the assignment must fail loudly, not silently run
        # against its own hardcoded folder.
        write(rd / "Py", "no_assign.py", "print('hi')\n")
        try:
            cc._run_project_script(sys.executable, rd / "Py" / "no_assign.py",
                                   cwd=rd / "Py", design_override=rd / "chosen",
                                   log=quiet)
            check(False, "an un-redirectable script is refused")
        except RuntimeError as e:
            check("would have written to its own hardcoded workspace" in str(e),
                  "an un-redirectable script is refused rather than misdirected")

        # --- redirect must not move the script -----------------------------
        # Regression: the first attempt copied the scripts to a temp folder.
        # These scripts do
        #     self_path = os.path.dirname(os.path.realpath(__file__))
        #     os.chdir(self_path)
        # and read CondFilterConf.json from self_path, so a copy elsewhere
        # broke that read - and would have written their __BackUp folder into
        # a temp dir that then got deleted.
        anchored = Path(d) / "anchored"
        (anchored / "Py").mkdir(parents=True)
        (anchored / "chosen").mkdir()
        write(anchored / "Py", "sidecar.json", '{"ok": true}\n')
        write(anchored / "Py", "conf_complex_gen.py",
              'import os\ndesign_path = os.path.abspath("../hardcoded")\n'
              'self_path = os.path.dirname(os.path.realpath(__file__))\n'
              'def where(): return design_path\n'
              'def sidecar(): return open(os.path.join(self_path, "sidecar.json")).read()\n')
        write(anchored / "Py", "conf_fill_ex.py",
              'import os\n'
              'design_path = os.path.abspath("../hardcoded")\n'
              'self_path = os.path.dirname(os.path.realpath(__file__))\n'
              'os.chdir(self_path)\n'
              'import conf_complex_gen\n'
              'print("DESIGN:" + design_path)\n'
              'print("SELF:" + self_path)\n'
              'print("SIDECAR:" + conf_complex_gen.sidecar().strip())\n'
              'print("SIBDESIGN:" + conf_complex_gen.where())\n')
        r = cc._run_project_script(sys.executable,
                                   anchored / "Py" / "conf_fill_ex.py",
                                   cwd=anchored / "Py",
                                   design_override=anchored / "chosen",
                                   log=quiet)
        out = r.stdout or ""
        check(r.returncode == 0, "anchored script runs", (r.stderr or "")[:200])
        check("DESIGN:" + str(anchored / "chosen") in out,
              "design_path is redirected to the chosen workspace")
        check("SIBDESIGN:" + str(anchored / "chosen") in out,
              "the imported sibling is redirected as well")
        # realpath() expands 8.3 short names, so compare as paths not strings.
        self_reported = out.split("SELF:")[1].splitlines()[0]
        check(cc._same_dir(self_reported, anchored / "Py"),
              "self_path still points at the script's REAL folder, not a copy",
              f"got {self_reported}")
        check('SIDECAR:{"ok": true}' in out,
              "a data file beside the script is still found (CondFilterConf.json case)")
        check("REDIRECT_OK" not in (r.stderr or ""),
              "the launcher's own marker is stripped from stderr")

        # --- non-ASCII output must not kill a script -----------------------
        # The child's stdout is a pipe, so print() uses the locale codec
        # (cp1252 here). save_map_data.py prints the map data, which contains
        # Chinese, so it died with UnicodeEncodeError before doing its work.
        write(proj, "chatty.py", "print('\u6536\u96c6\u5361')\n")
        r = cc._run_project_script(sys.executable, proj / "chatty.py", cwd=proj)
        check(r.returncode == 0, "a script printing non-ASCII exits cleanly",
              (r.stderr or "")[:160])
        check("\u6536\u96c6\u5361" in (r.stdout or ""),
              "the non-ASCII output survives the round trip")

        # --- map scene is chosen by content, not by filename ---------------
        # The old rule sorted names descending and took the first, which picked
        # BuilderScene.unity (an empty template) over BuilderScene-testCh1.unity
        # because '.' sorts above '-', and then failed with "No FruitPanel".
        scenes = Path(d) / "scenes"
        scenes.mkdir()
        (scenes / "BuilderScene.unity").write_text("GameObject:\n  m_Name: Canvas\n",
                                                   encoding="utf-8")
        (scenes / "BuilderScene-testCh1.unity").write_text(
            "GameObject:\n  m_Name: FruitPanel\nGameObject:\n  m_Name: FruitPanel\n",
            encoding="utf-8")
        picked = cc._pick_map_scene(scenes, "", quiet)
        check(picked.name == "BuilderScene-testCh1.unity",
              "the scene containing map blocks is chosen, not the alphabetical last",
              f"picked {picked.name}")

        explicit = cc._pick_map_scene(scenes, str(scenes / "BuilderScene.unity"), quiet)
        check(explicit.name == "BuilderScene.unity",
              "an explicitly chosen scene overrides the content check")

        empty = Path(d) / "scenes_empty"
        empty.mkdir()
        (empty / "OnlyTemplate.unity").write_text("GameObject:\n  m_Name: Canvas\n",
                                                  encoding="utf-8")
        try:
            cc._pick_map_scene(empty, "", quiet)
            check(False, "a folder with no usable scene raises")
        except ValueError as e:
            check("no map blocks" in str(e).lower(),
                  "the error explains that no scene has map blocks")
            check("OnlyTemplate.unity" in str(e),
                  "the error lists the scenes it looked at")

        # --- step 8 / coder_convert -----------------------------------------
        if os.name == "nt":
            conv = Path(d) / "conv"
            conv.mkdir()

            class ConvCfg:
                workspace = "WorkSpcae_NE"

                def design_root(self):
                    return conv

            def bat(name, body):
                (conv / name).write_text(body, encoding="utf-8", newline="\r\n")

            # Preferred variant wins, and each line is logged separately rather
            # than as one blob at the end - the console pane is fed live.
            bat("coder_convert_not_pause.bat",
                "@echo off\necho start convert table\necho copy data\n")
            bat("coder_convert.bat", "@echo off\necho WRONG SCRIPT\n")
            msgs = []
            cc.run_step8_import_config(ConvCfg(), lambda m, lvl=None: msgs.append(str(m)))
            joined = "\n".join(msgs)
            check("coder_convert_not_pause.bat" in joined,
                  "the not_pause variant is preferred over coder_convert.bat")
            check("WRONG SCRIPT" not in joined,
                  "the pausing variant is not run when the other exists")
            convert_lines = [m for m in msgs if "[convert]" in m]
            check(len(convert_lines) >= 2,
                  "each output line is logged on its own, not buffered into one",
                  f"got {len(convert_lines)} convert lines")

            # A trailing @pause must not stop the step completing.
            bat("coder_convert_not_pause.bat",
                "@echo off\necho converting\n@pause\n")
            msgs = []
            cc.run_step8_import_config(ConvCfg(), lambda m, lvl=None: msgs.append(str(m)))
            check(any("Config import completed" in m for m in msgs),
                  "a script ending in @pause still completes (stdin is /dev/null)")

            # Failure is surfaced, naming the script that failed.
            bat("coder_convert_not_pause.bat", "@echo off\necho nope\nexit /b 3\n")
            try:
                cc.run_step8_import_config(ConvCfg(), quiet)
                check(False, "a failing convert script raises")
            except RuntimeError as e:
                check("exited with code 3" in str(e),
                      "the convert script's exit code is reported")

            # Neither script present -> a clear error, not a crash.
            (conv / "coder_convert_not_pause.bat").unlink()
            (conv / "coder_convert.bat").unlink()
            try:
                cc.run_step8_import_config(ConvCfg(), quiet)
                check(False, "a missing convert script raises")
            except FileNotFoundError as e:
                check("coder_convert_not_pause.bat" in str(e),
                      "the missing-script error names what it looked for")

        # --- preflight only checks what the selected steps need -------------
        # Regression: the level-conflict check ran regardless of the steps
        # chosen, so once Magic Builder had done its job you could no longer
        # re-run the later steps - the normal way to resume a part-finished
        # chapter.
        pf = Path(d) / "pf"
        (pf / "lvl").mkdir(parents=True)
        (pf / "map").mkdir()
        (pf / "design").mkdir()
        write(pf / "lvl", "conf_fill_ex.py", "import os\n")
        make_book(pf / "design" / "pve_level.xlsx", "level", range(95001, 95201))
        make_book(pf / "design" / "pve_layout.xlsx", "layout", range(95001, 95201))

        class PfCfg:
            display_offset = 10
            new_ch = 95
            disp_new = 85
            lvl_py = pf / "lvl"
            map_py = pf / "map"

            def design(self):
                return pf / "design"

        blocked = cc.preflight_project_scripts(PfCfg(), quiet, ["magic"])
        check(any("already exist" in p for p in blocked),
              "with magic selected, existing levels block the run")
        later = cc.preflight_project_scripts(PfCfg(), quiet,
                                             ["validate", "reward", "asset", "qa"])
        check(later == [],
              "without magic, existing levels do NOT block the later steps",
              f"got {later}")

        # --- workbooks that cannot be written are caught before any step -----
        # Several design workbooks carry svn:needs-lock and sit read-only until
        # locked. Reward Config used to crash with a bare Errno 13 after it had
        # already saved chest.xlsx, leaving the chapter half-applied.
        import stat as _stat
        wbk = pf / "design" / "item.xlsx"
        make_book(wbk, "item", [1, 2, 3])
        os.chmod(wbk, _stat.S_IREAD)
        try:
            problem = cc.check_workbooks_writable(PfCfg(), quiet, ["reward"])
            check(problem is not None and "item.xlsx" in problem,
                  "a read-only workbook a step must write is reported up front")
            check("svn:needs-lock" in (problem or ""),
                  "the message explains why it is read-only and how to fix it")
            check(cc.check_workbooks_writable(PfCfg(), quiet, ["validate"]) is None,
                  "a step that does not write that workbook is not blocked by it")
        finally:
            os.chmod(wbk, _stat.S_IWRITE | _stat.S_IREAD)

        # --- preflight only inspects the scripts it actually runs ---
        lvl = Path(d) / "lvl"; lvl.mkdir()
        mp = Path(d) / "map"; mp.mkdir()
        write(lvl, "conf_fill_ex.py", "import os\n")
        write(lvl, "unrelated_tool.py", "import definitely_absent_pkg_zzz\n")

        class Cfg:
            lvl_py = lvl
            map_py = mp

            def design(self):
                return Path(d)

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
