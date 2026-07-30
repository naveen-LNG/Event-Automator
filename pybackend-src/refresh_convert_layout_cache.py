import sys
import os
import json

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import event_pipeline as ep


def main():
    sys.stdout.reconfigure(encoding='utf-8')
    if len(sys.argv) < 5:
        print("Usage: refresh_convert_layout_cache.py <proj_root> <log_dir> <workspace> <event_names_json_array>")
        sys.exit(1)

    proj_root = sys.argv[1]
    log_dir = sys.argv[2]
    ep.WORKSPACE_NAME = sys.argv[3]
    event_names = json.loads(sys.argv[4])

    os.makedirs(log_dir, exist_ok=True)
    ok = ep.refresh_convert_layout_cache(proj_root, log_dir, event_names)
    sys.exit(0 if ok else 1)


if __name__ == "__main__":
    main()
