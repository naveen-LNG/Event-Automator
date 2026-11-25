import sys
import os
import json

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import event_pipeline as ep


def main():
    sys.stdout.reconfigure(encoding='utf-8')
    # This script's stdout must be pure JSON for the Electron UI to parse — but
    # ep.get_quiz_prefix_from_mini_mgr() and friends call ep.log(), which prints to
    # stdout too. Redirect it to stderr here so log lines still surface (main.js
    # forwards stderr the same as stdout for other calls) without corrupting the JSON.
    ep.log = lambda msg: print(f"[PythonBackend] {msg}", file=sys.stderr, flush=True)

    if len(sys.argv) < 3:
        print(json.dumps({"error": "Usage: get_trivia_challenge_info.py <proj_root> <workspace> [all_branch_design_dirs_json] [questions_source]"}))
        sys.exit(1)

    proj_root = sys.argv[1]
    ep.WORKSPACE_NAME = sys.argv[2]
    # A list of design dirs, or {label: design_dir} so the usage table can name branches.
    all_branch_design_dirs = json.loads(sys.argv[3]) if len(sys.argv) > 3 and sys.argv[3] else []
    questions_source = sys.argv[4] if len(sys.argv) > 4 and sys.argv[4] else None
    labeled_dirs = all_branch_design_dirs if isinstance(all_branch_design_dirs, dict) \
        else {d: d for d in all_branch_design_dirs}

    result = {}

    prefix = ep.get_quiz_prefix_from_mini_mgr(proj_root)
    result["prefix"] = prefix
    if not prefix:
        result["error"] = ("No upcoming 答题挑战 cycle is scheduled in mini_mgr.xlsx (活动数据) for this branch/workspace — "
                           "add the next minigame set on the Minigame Configs tab first, then refresh here. "
                           "See the Pipeline Output for the latest cycle found.")
        print(json.dumps(result, ensure_ascii=False))
        sys.exit(1)

    design_dir = os.path.join(proj_root, "design", "DesignData", ep.WORKSPACE_NAME, "design")

    usage = ep.scan_batch_usage(labeled_dirs)
    result["batch_usage"] = usage
    result["available_batches"] = ep.available_batch_numbers()
    result["next_unused_batch"] = ep.max_used_batch(usage) + 1

    if questions_source:
        batch_number = ep.batch_number_from_source(questions_source)
    else:
        questions_source, batch_number = ep.resolve_questions_source(design_dir, labeled_dirs, prefix)
    result["questions_module"] = questions_source
    result["batch_number"] = batch_number

    if questions_source.endswith(".json") and not os.path.exists(questions_source):
        result["error"] = f"Questions file not found: {questions_source}"
        print(json.dumps(result, ensure_ascii=False))
        sys.exit(1)

    try:
        questions = ep.load_trivia_questions(questions_source)
    except Exception as e:
        result["error"] = f"Failed to load questions from '{questions_source}': {e}"
        print(json.dumps(result, ensure_ascii=False))
        sys.exit(1)
    result["question_count"] = len(questions)
    tiers = {1: 0, 2: 0, 3: 0}
    for q in questions:
        tiers[q["tier"]] = tiers.get(q["tier"], 0) + 1
    result["tier_counts"] = tiers
    result["correct_counts"] = {i: sum(1 for q in questions if q["correct_index"] == i) for i in (1, 2)}

    result["already_current"] = ep.quiz_cycle_is_current(design_dir, prefix, batch_number, questions)

    print(json.dumps(result, ensure_ascii=False))


if __name__ == "__main__":
    main()
