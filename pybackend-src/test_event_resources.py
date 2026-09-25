"""Step 24 (clone_event_resources): every state the target folder can be in.

An existing ResourcesAb/Events/<Tgt> used to count as "already cloned" even when
it was empty - the same mistake that broke chapter update's dynamic_map copy three
runs in a row. Only real assets make the folder somebody's work.

Uses a synthetic Unity folder whose prefab references its sprite by GUID, so the
GUID regeneration + remapping is checked on the fill path too, not just the
fresh-clone path.
"""
import os
import re
import stat
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
import event_pipeline as ep

FAILS = []
SRC_GUID = "a" * 32


def check(ok, msg, detail=""):
    print(("OK   " if ok else "FAIL ") + msg)
    if not ok:
        if detail:
            print("     " + str(detail)[:300])
        FAILS.append(msg)


def build(root, state):
    """A branch with ResourcesAb/Events/Sim2607, and the target in `state`."""
    events = os.path.join(root, "develop", "client", "Skipbo", "Assets", "ResourcesAb", "Events")
    src = os.path.join(events, "Sim2607")
    os.makedirs(os.path.join(src, "Prefabs"))
    with open(os.path.join(src, "Prefabs", "UISim2607Panel.prefab"), "w", newline="\n") as f:
        f.write("%YAML 1.1\n--- !u!1 &1\nGameObject:\n  m_Name: UISim2607Panel\n"
                f"  m_Sprite: {{fileID: 21300000, guid: {SRC_GUID}, type: 3}}\n")
    with open(os.path.join(src, "Prefabs", "UISim2607Panel.prefab.meta"), "w", newline="\n") as f:
        f.write("fileFormatVersion: 2\nguid: " + "b" * 32 + "\n")
    with open(os.path.join(src, "bg_sim2607.png"), "wb") as f:
        f.write(b"\x89PNG-sim2607-bytes")
    with open(os.path.join(src, "bg_sim2607.png.meta"), "w", newline="\n") as f:
        f.write(f"fileFormatVersion: 2\nguid: {SRC_GUID}\n")

    tgt = os.path.join(events, "Sim2609")
    if state == "empty shell":
        os.makedirs(tgt)
    elif state == "shell with orphan .meta":
        # What a reverted clone really leaves: .meta files already carrying the
        # TARGET's name (they collide with the rename pass), read-only because
        # svn:needs-lock carries over (rmtree cannot delete those on Windows).
        os.makedirs(os.path.join(tgt, "Prefabs"))
        for rel in ("bg_sim2609.png.meta", os.path.join("Prefabs", "UISim2609Panel.prefab.meta")):
            p = os.path.join(tgt, rel)
            with open(p, "w") as f:
                f.write("fileFormatVersion: 2\nguid: " + "c" * 32 + "\n")
            os.chmod(p, stat.S_IREAD)
    elif state == "real assets present":
        os.makedirs(tgt)
        with open(os.path.join(tgt, "bg_sim2609.png"), "wb") as f:
            f.write(b"REAL ART")
    return tgt


def main():
    print("--- step 24: clone_event_resources across target states ---")
    for state in ("absent", "empty shell", "shell with orphan .meta"):
        with tempfile.TemporaryDirectory() as td:
            tgt = build(td, state)
            log_dir = os.path.join(td, "log")
            os.makedirs(log_dir)
            ep.clone_event_resources(td, "sim2607", "sim2609", log_dir, 23)

            prefab = os.path.join(tgt, "Prefabs", "UISim2609Panel.prefab")
            png = os.path.join(tgt, "bg_sim2609.png")
            check(os.path.isfile(prefab) and os.path.isfile(png),
                  f"{state}: the event's resources are cloned and renamed",
                  sorted(os.listdir(tgt)) if os.path.isdir(tgt) else "no target")
            if not os.path.isfile(prefab):
                continue
            body = open(prefab, encoding="utf-8").read()
            new_guid = re.search(r"guid:\s*([0-9a-f]{32})",
                                 open(png + ".meta", encoding="utf-8").read()).group(1)
            check(new_guid != SRC_GUID, f"{state}: the sprite gets a fresh GUID")
            check(f"guid: {new_guid}" in body and SRC_GUID not in body,
                  f"{state}: the prefab is relinked to the clone's sprite, not the source's",
                  body)
            check("m_Name: UISim2609Panel" in body,
                  f"{state}: names inside the prefab are renamed", body)
            check(open(png, "rb").read() == b"\x89PNG-sim2607-bytes",
                  f"{state}: binary files are copied byte-for-byte")

    with tempfile.TemporaryDirectory() as td:
        tgt = build(td, "real assets present")
        log_dir = os.path.join(td, "log")
        os.makedirs(log_dir)
        ep.clone_event_resources(td, "sim2607", "sim2609", log_dir, 23)
        check(open(os.path.join(tgt, "bg_sim2609.png"), "rb").read() == b"REAL ART",
              "real assets present: nothing is overwritten")
        check(not os.path.exists(os.path.join(tgt, "Prefabs")),
              "real assets present: the folder is left exactly as it was")

    step25()

    print("")
    if FAILS:
        print(f"{len(FAILS)} CHECK(S) FAILED")
        sys.exit(1)
    print("EVENT RESOURCES OK")


S_GUID, O_GUID = "5" * 32, "0" * 32


def _branch_for_step25(root, rows, tgt_id):
    """events.xlsx with `rows`, extra_args naming tgt_id, and an empty Common tree."""
    import json
    import openpyxl
    design = os.path.join(root, "design", "DesignData", ep.WORKSPACE_NAME, "design")
    os.makedirs(design)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "events"
    ws.append(["ID", "Alias", "", "", "", "", "", "Start"])
    for eid, alias, start in rows:
        ws.append([eid, alias, "", "", "", "", "", start])
    wb.save(os.path.join(design, "events.xlsx"))
    log_dir = os.path.join(root, "log")
    os.makedirs(log_dir)
    with open(os.path.join(log_dir, "extra_args.json"), "w") as f:
        json.dump({"event_id": tgt_id, "src_start_time": ""}, f)
    events = os.path.join(root, "develop", "client", "Skipbo", "Assets", "ResourcesAb", "Events")
    img = os.path.join(events, "Common", "Icon", "CommonEventImg")
    os.makedirs(img)
    return log_dir, events, img


def _write(p, text):
    os.makedirs(os.path.dirname(p), exist_ok=True)
    with open(p, "w", newline="\n") as f:
        f.write(text)


def step25():
    print("\n--- step 25: event-ID resources ---")

    # New ID, with a leftover target-named .meta already in place. The copied image takes
    # that .meta's GUID, so the event's own prefab must be repointed to it - not left on
    # the source event's image.
    with tempfile.TemporaryDirectory() as td:
        log_dir, events, img = _branch_for_step25(
            td, [(20260801, "Merge2608活动", "2026-08-06")], "20261101")
        with open(os.path.join(img, "20260801_img_token.png"), "wb") as f:
            f.write(b"PNG-token")
        _write(os.path.join(img, "20260801_img_token.png.meta"), f"fileFormatVersion: 2\nguid: {S_GUID}\n")
        _write(os.path.join(img, "20261101_img_token.png.meta"), f"fileFormatVersion: 2\nguid: {O_GUID}\n")
        prefab = os.path.join(events, "Merge2611", "Prefabs", "UIMerge2611Panel.prefab")
        _write(prefab, f"%YAML 1.1\n--- !u!1 &1\n  m_Sprite: {{fileID: 21300000, guid: {S_GUID}, type: 3}}\n")

        ep.clone_event_id_resources(td, "merge2608", "merge2611", log_dir, 24)

        check(os.path.isfile(os.path.join(img, "20261101_img_token.png")),
              "new ID: the image is copied under the new ID")
        meta = open(os.path.join(img, "20261101_img_token.png.meta")).read()
        check(O_GUID in meta, "new ID: the leftover .meta is kept as it was", meta)
        body = open(prefab).read()
        check(f"guid: {O_GUID}" in body and S_GUID not in body,
              "new ID: the event's prefab now points at the copied image, not the source's",
              body)

    # Same name, same ID: the event's own deleted files come back from SVN in place.
    with tempfile.TemporaryDirectory() as td:
        log_dir, events, img = _branch_for_step25(
            td, [(20251002, "Merge2510活动", "2025-10-16")], "20251002")
        staged = os.path.join(td, "svn_copy")
        os.makedirs(staged)
        with open(os.path.join(staged, "20251002_img_token.png"), "wb") as f:
            f.write(b"PNG-original")
        _write(os.path.join(staged, "20251002_img_token.png.meta"), f"fileFormatVersion: 2\nguid: {S_GUID}\n")

        def fake_svn(folder, prefix, hint, dest):
            if "CommonEventImg" not in folder:
                return {}
            return {n: os.path.join(staged, n) for n in os.listdir(staged) if n.startswith(prefix)}

        real_svn, real_log = ep._svn_export_deleted, ep.log
        seen = []
        ep._svn_export_deleted = fake_svn
        ep.log = lambda m, *a, **k: seen.append(str(m))
        try:
            ep.clone_event_id_resources(td, "merge2510", "merge2510", log_dir, 24)
        finally:
            ep._svn_export_deleted, ep.log = real_svn, real_log

        check(os.path.isfile(os.path.join(img, "20251002_img_token.png")),
              "same ID: the deleted image is restored in place")
        meta = open(os.path.join(img, "20251002_img_token.png.meta")).read()
        check(S_GUID in meta, "same ID: it keeps its ORIGINAL GUID, so old links resolve", meta)
        check(any("Checking" in m and "GUID in Assets" in m for m in seen),
              "same ID: the long GUID scan is announced before it starts",
              [m for m in seen if "GUID" in m])


if __name__ == "__main__":
    main()
