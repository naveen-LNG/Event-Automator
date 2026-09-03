"""
Chapter Update backend for the Electron app's "Chapter Update" tab.

Two entry points, matching the convention every other tab already uses:

  get_chapter_info <designPath>
      -> one JSON blob on stdout. Detects the newest configured chapter
         straight out of pve_chapter.xlsx, so nobody has to type chapter
         numbers (or get the internal/display offset wrong) by hand.

  update_chapter <branchRoot> <payload.json>
      -> plain-text progress log on stdout, non-zero exit on failure.

All the actual work lives in chapter_core.py, shared verbatim with the
standalone ChapterCreator tool. Nothing here reimplements a step.
"""
import json
import sys
import traceback
from pathlib import Path

import openpyxl

import chapter_core
from chapter_core import Cfg

# The step logs carry ✓ / → / ⚠ and Chinese sheet names. Piped stdout on Windows
# defaults to cp1252, which raises UnicodeEncodeError on all of them — the GUI
# tool never hit this because it logged into a Tk widget instead of stdout.
# main.js decodes the pipe as UTF-8, so this is the matching half.
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

# Ordered exactly as the workflow runs them. Keys are what the UI sends.
STEPS = [
    ("magic",    "Magic Builder",     chapter_core.run_step1_magic_builder),
    ("upload",   "Screenshot Upload", chapter_core.run_step2_upload),
    ("validate", "Level Validation",  chapter_core.run_step3_level_validation),
    ("reward",   "Reward Config",     chapter_core.run_step4_reward_config),
    ("art",      "Art Assets",        chapter_core.run_step5_art_assets),
    ("asset",    "Asset Config",      chapter_core.run_step6_asset_config),
    ("map",      "Map Builder",       chapter_core.run_step7_map_builder),
    ("import",   "Import Config",     chapter_core.run_step8_import_config),
    ("sim",      "Simulation",        chapter_core.run_step9_simulation),
    ("balance",  "Level Balancing",   chapter_core.run_step10_level_balancing),
    # Always last: it re-reads the saved workbooks and scores them against the
    # Chapter Creation & Level Delivery review checklist.
    ("qa",       "QA Validation",     chapter_core.run_qa_validation),
]
STEP_MAP = {key: (label, fn) for key, label, fn in STEPS}

# Steps with a prerequisite the app cannot provide. Surfaced in get_chapter_info
# so the tab can warn up front instead of failing halfway through a run.
STEP_NEEDS = {
    "magic":  "a system Python with pandas/numpy (conf_fill_ex.py)",
    "upload": "the svn command-line client, and screenshots already captured in Unity",
    "map":    "the Unity scene for this chapter, already built and saved",
    "import": "coder_convert.bat in the workspace",
    "sim":    "RunCardEx2 (start_run_2.bat) — takes a long time to finish",
}


def _detect_latest_chapter(design_path):
    """Newest row in pve_chapter.xlsx's `chapter` sheet.

    Column A is the internal map id (88000), column C the display number (78).
    Reading both is what makes the internal/display offset a fact rather than a
    setting somebody can get wrong.
    """
    pvc = Path(design_path) / "pve_chapter.xlsx"
    if not pvc.exists():
        return {"error": f"pve_chapter.xlsx not found: {pvc}"}

    wb = openpyxl.load_workbook(str(pvc), read_only=True, data_only=True)
    try:
        ws = wb["chapter"]
        best = None
        for row in ws.iter_rows(min_row=2, max_col=3, values_only=True):
            map_id, _alias, disp = (list(row) + [None, None, None])[:3]
            if map_id is None:
                continue
            try:
                map_id = int(map_id)
                disp = int(disp)
            except (TypeError, ValueError):
                continue
            if best is None or map_id > best[0]:
                best = (map_id, disp)
    finally:
        wb.close()

    if best is None:
        return {"error": "No usable rows in pve_chapter.xlsx / chapter sheet."}

    map_id, disp = best
    old_ch = map_id // 1000
    return {
        "old_chapter": old_ch,
        "new_chapter": old_ch + 1,
        "display_offset": old_ch - disp,
        "display_old": disp,
        "display_new": disp + 1,
        "old_map_id": map_id,
        "new_map_id": (old_ch + 1) * 1000,
        "level_range": [(old_ch + 1) * 1000 + 1, (old_ch + 1) * 1000 + 200],
        "steps": [{"key": k, "label": lbl, "needs": STEP_NEEDS.get(k, "")}
                  for k, lbl, _ in STEPS],
        "has_system_python": bool(chapter_core._python_exe()),
    }


def info_main():
    if len(sys.argv) < 2:
        print(json.dumps({"error": "usage: get_chapter_info <designPath>"}))
        sys.exit(1)
    try:
        print(json.dumps(_detect_latest_chapter(sys.argv[1]), ensure_ascii=False))
    except Exception as e:
        print(json.dumps({"error": str(e)}, ensure_ascii=False))
        sys.exit(1)


def run_main():
    if len(sys.argv) < 3:
        print("usage: update_chapter <branchRoot> <payload.json>")
        sys.exit(1)

    branch_root = sys.argv[1]
    payload = json.loads(Path(sys.argv[2]).read_text(encoding="utf-8"))

    requested = payload.get("steps") or []
    unknown = [s for s in requested if s not in STEP_MAP]
    if unknown:
        print(f"Unknown step(s): {', '.join(unknown)}")
        sys.exit(1)
    if not requested:
        print("No steps selected — nothing to do.")
        sys.exit(1)

    # Told to us by main.js, because the backend cannot find this out for
    # itself: in an installed app package.json lives inside app.asar, so there
    # is nothing on disk beside resources\pybackend that names the build.
    # Without it a pasted log cannot answer "is this install current?".
    if payload.get("app_version"):
        chapter_core.APP_VERSION = str(payload["app_version"])

    cfg = Cfg(
        branch_root=branch_root,
        workspace=payload["workspace"],
        old_ch=payload["old_chapter"],
        display_offset=payload.get("display_offset", 10),
        sim_path=payload.get("sim_path") or None,
        lvl_py=payload.get("lvl_py") or None,
        map_builder=payload.get("map_builder") or None,
        sim_repeat=payload.get("sim_repeat", 500),
        sim_p_num=payload.get("sim_p_num"),   # unset = decide from the machine
        sim_tag=payload.get("sim_tag", ""),
    )

    def log(msg, level="INFO"):
        # Flush per line: main.js streams this straight into the console pane,
        # and a buffered 5-step run would otherwise look frozen.
        print(msg, flush=True)

    print(f"Chapter {cfg.disp_old} -> {cfg.disp_new}  "
          f"(internal {cfg.old_ch} -> {cfg.new_ch}, map id {cfg.new_ch * 1000})")
    print(f"Workspace: {cfg.workspace}")
    print(f"Design:    {cfg.design()}")
    print("")

    # Dependency preflight before anything is written. Steps 1, 2 and 4 shell
    # out to Level Builder / Map Editor scripts, and a missing package used to
    # surface only when that step ran - step 4 failing leaves a half-applied
    # chapter needing an svn revert. Checking up front costs a few seconds and
    # fails while the design files are still untouched.
    problems = chapter_core.preflight_project_scripts(cfg, log, requested)
    if problems:
        print("=== Dependency check ===", flush=True)
        for p in problems:
            print(p, flush=True)
            print("", flush=True)
        print("Nothing has been changed. Resolve the above and re-run.", flush=True)
        sys.exit(1)

    failed = 0
    for key in [k for k, _, _ in STEPS if k in requested]:   # keep workflow order
        label, fn = STEP_MAP[key]
        print(f"=== {label} ===", flush=True)
        try:
            if key == "map":
                # Step 7 also takes the scene to parse ("" = newest .unity).
                fn(cfg, payload.get("scene_file", ""), log)
            else:
                fn(cfg, log)
        except Exception as e:
            failed += 1
            print(f"  X {label} FAILED: {e}", flush=True)
            traceback.print_exc(file=sys.stdout)
            # Later steps build on earlier ones (asset config reads the fruit
            # numbering reward config just wrote), so stopping is the only safe
            # thing to do — a half-applied chapter is worse than none.
            if key == "qa":
                # QA only reads the workbooks; a failure here means a
                # broken report, not a half-applied chapter.
                print("\nQA validation could not complete. The chapter data is untouched.")
            else:
                print("\nStopped: later steps depend on this one. "
                      "Fix the cause, svn revert the design files, and re-run.")
            break
        print("", flush=True)

    if failed:
        sys.exit(1)
    print("All selected steps completed.")


def main():
    # Kept so `python chapter_update.py info|run ...` works during development;
    # the dispatcher calls info_main/run_main directly.
    mode = sys.argv[1] if len(sys.argv) > 1 else ""
    sys.argv = [sys.argv[0]] + sys.argv[2:]
    if mode == "info":
        info_main()
    elif mode == "run":
        run_main()
    else:
        print("usage: chapter_update.py info|run ...")
        sys.exit(1)


if __name__ == "__main__":
    main()
