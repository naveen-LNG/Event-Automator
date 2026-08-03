import sys
import os
import json

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import event_pipeline as ep


def main():
    sys.stdout.reconfigure(encoding='utf-8')
    if len(sys.argv) < 4:
        print("Usage: trivia_challenge_cycle.py <proj_root> <workspace> <mode> [all_branch_design_dirs_json] [questions_source]")
        print("  mode: 'run' refreshes both files; 'prefix' only prints the computed QUIZ ID prefix")
        print("  all_branch_design_dirs_json: JSON array of every configured branch's design dir —")
        print("  used to find the next batch nobody else has used yet; pass '[]' for single-branch use")
        print("  questions_source: JSON path, or bundled batch module name — overrides auto-resolution")
        sys.exit(1)

    proj_root = sys.argv[1]
    ep.WORKSPACE_NAME = sys.argv[2]
    mode = sys.argv[3]
    all_branch_design_dirs = json.loads(sys.argv[4]) if len(sys.argv) > 4 and sys.argv[4] else []
    questions_source = sys.argv[5] if len(sys.argv) > 5 and sys.argv[5] else None

    if mode == "prefix":
        prefix = ep.get_quiz_prefix_from_mini_mgr(proj_root)
        sys.exit(0 if prefix else 1)

    log_dir = os.path.join(proj_root, r"develop\client\Skipbo\Assets\Editor\EventAutomation\PythonBackend")
    os.makedirs(log_dir, exist_ok=True)

    ok = ep.run_trivia_challenge_cycle(proj_root, log_dir, "trivia_challenge", questions_source, all_branch_design_dirs)
    sys.exit(0 if ok else 1)


if __name__ == "__main__":
    main()
