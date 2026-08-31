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
    elif script == "update_decoration_shop":
        import update_decoration_shop
        update_decoration_shop.main()
    elif script == "update_robot_pvp":
        import update_robot_pvp
        update_robot_pvp.main()
    elif script == "update_resident_task":
        import update_resident_task
        update_resident_task.main()
    elif script == "get_events_by_type":
        import get_events_by_type
        get_events_by_type.main()
    elif script == "refresh_convert_layout_cache":
        import refresh_convert_layout_cache
        refresh_convert_layout_cache.main()
    elif script == "refresh_all_caches":
        import refresh_all_caches
        refresh_all_caches.main()
    elif script == "update_weekly_gift":
        import update_weekly_gift
        update_weekly_gift.main()
    elif script == "get_full_screen_gift_info":
        import get_full_screen_gift_info
        get_full_screen_gift_info.main()
    elif script == "update_full_screen_gift":
        import update_full_screen_gift
        update_full_screen_gift.main()
    elif script == "trivia_challenge_cycle":
        import trivia_challenge_cycle
        trivia_challenge_cycle.main()
    elif script == "get_trivia_challenge_info":
        import get_trivia_challenge_info
        get_trivia_challenge_info.main()
    elif script == "get_endless_treasure_info":
        import get_endless_treasure_info
        get_endless_treasure_info.main()
    elif script == "update_endless_treasure":
        import update_endless_treasure
        update_endless_treasure.main()
    elif script == "ideation_generate":
        import ideation_generator
        ideation_generator.generate_main()
    elif script == "ideation_regenerate":
        import ideation_generator
        ideation_generator.regenerate_main()
    elif script == "ideation_entries":
        import ideation_generator
        ideation_generator.entries_main()
    elif script == "ideation_retract":
        import ideation_generator
        ideation_generator.retract_main()
    elif script == "sprint_board":
        import sprint_board
        sprint_board.main()
    elif script == "profile_save":
        import ideation_generator
        ideation_generator.profile_save_main()
    elif script == "profiles_read":
        import ideation_generator
        ideation_generator.profiles_read_main()
    elif script == "doc_convert":
        import doc_convert
        doc_convert.convert_main()
    elif script == "md_to_docx":
        import doc_convert
        doc_convert.docx_main()
    elif script == "artreq_generate":
        import art_requirement_generator
        art_requirement_generator.generate_main()
    elif script == "artreq_regenerate":
        import art_requirement_generator
        art_requirement_generator.regenerate_main()
    elif script == "artreq_save_final":
        import art_requirement_generator
        art_requirement_generator.save_final_main()
    elif script == "artreq_mcp_save":
        import art_requirement_generator
        art_requirement_generator.mcp_save_main()
    elif script == "artreq_schemas":
        import art_requirement_generator
        art_requirement_generator.schemas_main()
    elif script == "artreq_preview":
        import art_requirement_generator
        art_requirement_generator.preview_main()
    elif script == "artreq_images":
        import artreq_images
        artreq_images.images_main()
    elif script == "artreq_plan":
        import artreq_images
        artreq_images.plan_main()
    elif script == "comfy_check":
        import artreq_images
        artreq_images.check_main()
    elif script == "artreq_regen_image":
        import artreq_images
        artreq_images.regenerate_main()
    elif script == "sequence_review":
        import sequence_review
        sequence_review.main()
    elif script == "get_chapter_info":
        import chapter_update
        chapter_update.info_main()
    elif script == "update_chapter":
        import chapter_update
        chapter_update.run_main()
    else:
        print(f"Unknown script: {script}")
        sys.exit(1)


if __name__ == "__main__":
    main()
