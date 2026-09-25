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

    print("")
    if FAILS:
        print(f"{len(FAILS)} CHECK(S) FAILED")
        sys.exit(1)
    print("EVENT RESOURCES OK")


if __name__ == "__main__":
    main()
