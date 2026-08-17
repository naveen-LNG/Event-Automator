import sys, os, shutil
sys.path.insert(0, ".")
import ideation_store as st

ROOT = r"C:\SkipBo\tools\EventAutomation\PythonBackend\_store_test_root"
if os.path.isdir(ROOT):
    shutil.rmtree(ROOT)

# Two different authors giving feedback on the same doc -> must be 2 files.
e1 = st.write_feedback_entry(ROOT, "Merge", "ideation", "Merge2609", "Sunday Garden Club",
                             "Make colours more vibrant", "Default hero palettes to saturated, not pastel.",
                             r"AzureAD\NaveenKumar", "Naveen Kumar")
e2 = st.write_feedback_entry(ROOT, "Merge", "ideation", "Merge2609", "Sunday Garden Club",
                             "Anniversary props please", "Reference the live anniversary event in props.",
                             r"AzureAD\Priya", "Priya S")
entry_dir = os.path.join(ROOT, "Merge", "Feedback", "entries")
files = os.listdir(entry_dir)
assert len(files) == 2, f"expected 2 entry files, got {files}"
assert e1["entry_id"] != e2["entry_id"]
print(f"OK  two authors -> {len(files)} distinct files, no collision")

# Lessons aggregate from both authors.
lessons = st.active_lessons(ROOT, "Merge")
assert len(lessons) == 2, lessons
print(f"OK  active lessons = {len(lessons)} ({[l['author_name'] for l in lessons]})")

# Retract one -> drops out of prompt, other survives.
st.write_retraction(ROOT, "Merge", e1["entry_id"], r"AzureAD\Priya", "Priya S", "Too prescriptive")
lessons2 = st.active_lessons(ROOT, "Merge")
assert len(lessons2) == 1, lessons2
assert lessons2[0]["entry_id"] == e2["entry_id"]
block = st.lessons_prompt_block(ROOT, "Merge")
assert "saturated" not in block, "retracted lesson leaked into the prompt!"
assert "anniversary" in block.lower()
print("OK  retraction removes only that lesson from the prompt block")

# Feedback thread still shows BOTH (history is preserved, not deleted).
thread = st.read_feedback_entries(ROOT, "Merge", doc_kind="ideation", ref_name="Merge2609")
assert len(thread) == 2, thread
print(f"OK  thread still shows all {len(thread)} entries (retraction != deletion)")

# Theme history across two separate runs by different people.
st.record_generated_themes(ROOT, "Merge", [{"theme_name": "Sunday Garden Club", "internal_reference_name": "Merge2609"}],
                           r"AzureAD\NaveenKumar", "Naveen Kumar")
st.record_generated_themes(ROOT, "Merge", [{"theme_name": "Cozy Book Nook", "internal_reference_name": "Merge2610"}],
                           r"AzureAD\Priya", "Priya S")
hist_dir = os.path.join(ROOT, "_theme_history", "Merge")
assert len(os.listdir(hist_dir)) == 2, os.listdir(hist_dir)
pt = st.past_themes(ROOT, "Merge")
assert len(pt) == 2, pt
assert "Cozy Book Nook" in st.past_themes_block(ROOT, "Merge")
print(f"OK  theme history = {len(os.listdir(hist_dir))} run files, {len(pt)} themes merged for de-dup")

# Corrupt file must not break aggregation (half-synced OneDrive file).
with open(os.path.join(entry_dir, "99999999T000000Z_broken.json"), "w", encoding="utf-8") as f:
    f.write("{ this is not valid json")
assert len(st.active_lessons(ROOT, "Merge")) == 1, "corrupt file broke aggregation"
print("OK  corrupt/half-synced file skipped, aggregation survives")

# Empty root = first run ever, must not crash.
assert st.active_lessons(ROOT, "Sort") == []
assert "nothing to avoid" in st.past_themes_block(ROOT, "Sort")
print("OK  empty/first-run state handled")

shutil.rmtree(ROOT)
print("\nALL STORE TESTS PASSED")
