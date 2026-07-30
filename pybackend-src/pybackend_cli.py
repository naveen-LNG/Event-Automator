"""
Single dispatcher entry point compiled by PyInstaller into one bundled executable
(with a shared runtime, not 9 separate ~40MB standalone exes). The standalone
Electron app calls this instead of `python <script>.py <args>` — same scripts,
same argv contract, just invoked via `pybackend_cli.exe <script_name> <args...>`
so end users never need Python installed at all.
"""
import sys


def main():
    if len(sys.argv) < 2:
        print("Usage: pybackend_cli.exe <script_name> [args...]")
        sys.exit(1)

    script = sys.argv[1]
    # Downstream scripts read sys.argv as if they were invoked directly
    # (argv[0] = script name is never actually used by any of them, but keep
    # the shape consistent with normal `python script.py arg1 arg2` invocation).
    sys.argv = [script] + sys.argv[2:]

    if script == "event_pipeline":
        import event_pipeline
        event_pipeline.main()
    elif script == "get_minigames_info":
        import get_minigames_info
        get_minigames_info.main()
    elif script == "update_minigames":
        import update_minigames
        update_minigames.main()
    elif script == "get_bp_info":
        import get_bp_info
        get_bp_info.main()
    elif script == "update_bp":
        import update_bp
        update_bp.main()
    elif script == "get_hbp_info":
        import get_hbp_info
        get_hbp_info.main()
    elif script == "update_hbp":
        import update_hbp
        update_hbp.main()
    elif script == "get_events_by_type":
        import get_events_by_type
        get_events_by_type.main()
    elif script == "refresh_convert_layout_cache":
        import refresh_convert_layout_cache
        refresh_convert_layout_cache.main()
    else:
        print(f"Unknown script: {script}")
        sys.exit(1)


if __name__ == "__main__":
    main()
