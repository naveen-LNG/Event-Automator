import json, os, shutil, subprocess, sys

ROOT = r"C:\SkipBo\tools\EventAutomation\PythonBackend\_cli_test_root"
if os.path.isdir(ROOT):
    shutil.rmtree(ROOT)

HERE = os.path.dirname(os.path.abspath(__file__))
os.chdir(HERE)
sys.path.insert(0, HERE)
import ideation_store as st

# Seed two authors' feedback via the store (simulating two regenerate runs).
e1 = st.write_feedback_entry(ROOT, "Merge", "ideation", "Merge2609", "Sunday Garden Club",
                             "More vibrant please", "Default hero palettes to saturated, not pastel.",
                             r"AzureAD\NaveenKumar", "Naveen Kumar", revised_title="Sunday Garden Club (Vivid)")
e2 = st.write_feedback_entry(ROOT, "Merge", "ideation", "Merge2609", "Sunday Garden Club (Vivid)",
                             "Add anniversary props", "Reference the live anniversary event in props.",
                             r"AzureAD\Priya", "Priya S")

PY = sys.executable  # run me with the bundled pybackend_dist\python\python.exe

def run(script, payload):
    with open("_p.json", "w", encoding="utf-8") as f:
        json.dump(payload, f)
    r = subprocess.run([PY, "pybackend_cli.py", script, "_p.json"],
                       capture_output=True, text=True, encoding="utf-8")
    return r.returncode, json.loads(r.stdout) if r.stdout.strip() else {"_raw": r.stdout, "_err": r.stderr}

# --- read the thread ---
code, out = run("ideation_entries", {"events_root": ROOT, "event_type": "Merge",
                                     "doc_kind": "ideation", "ref_name": "Merge2609"})
assert code == 0, out
assert len(out["entries"]) == 2, out
assert [e["author_name"] for e in out["entries"]] == ["Naveen Kumar", "Priya S"], out
assert len(out["lessons"]) == 2, out
assert all(e["retracted"] is False for e in out["entries"])
print(f"OK  ideation_entries -> {len(out['entries'])} entries by {[e['author_name'] for e in out['entries']]}, {len(out['lessons'])} active lessons")

# --- retract one author's lesson as a DIFFERENT user ---
code, out = run("ideation_retract", {"events_root": ROOT, "event_type": "Merge",
                                     "entry_id": e1["entry_id"],
                                     "author_id": r"AzureAD\Priya", "author_name": "Priya S",
                                     "reason": "Too prescriptive for seasonal themes"})
assert code == 0 and out.get("ok"), out
assert len(out["lessons"]) == 1, out
print(f"OK  ideation_retract -> lesson removed, {len(out['lessons'])} lesson remains in prompt")

# --- thread still shows both, one flagged retracted ---
code, out = run("ideation_entries", {"events_root": ROOT, "event_type": "Merge"})
retracted = [e for e in out["entries"] if e["retracted"]]
assert len(out["entries"]) == 2 and len(retracted) == 1, out
assert retracted[0]["retracted_by"] == "Priya S", retracted
print(f"OK  thread preserved: 2 entries, 1 flagged retracted by {retracted[0]['retracted_by']}")

# --- human-readable digest regenerated on disk, shows the strikethrough ---
digest = os.path.join(ROOT, "Merge", "Feedback", "Merge2609 Ideation Feedback.md")
assert os.path.exists(digest), "digest not written"
text = open(digest, encoding="utf-8").read()
assert "Naveen Kumar" in text and "Priya S" in text, text
assert "RETRACTED" in text and "~~" in text, text
assert "Revised to:" in text, text
print("OK  digest .md regenerated with both authors + retraction marked")

os.remove("_p.json")
shutil.rmtree(ROOT)
print("\nALL CLI TESTS PASSED")
