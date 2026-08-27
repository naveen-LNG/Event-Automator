# Trigger eval set

`trigger-eval.json` is 20 realistic queries — 9 that should pull this skill in,
11 near-misses that should not. The negatives are the valuable half: they share
vocabulary with the skill ("run", "launch", the same filenames) but need
something else — the Python suites, `pybackend_cli`, the build, the Unity client,
or just an edit. A negative that is obviously irrelevant tests nothing.

## Running it

The measured optimisation loop needs the **Claude Code CLI** (`claude -p`), which
was not installed on this machine when the set was written — only Claude Desktop,
which is a different binary. Nothing here has therefore been measured; the
current description was reasoned against these queries, not scored on them.

Once the CLI is available:

```bash
cd <skill-creator dir>
python -m scripts.run_loop \
  --eval-set  C:/SkipBo/tools/EventAutomationDesktop/.claude/skills/run-event-automator/evals/trigger-eval.json \
  --skill-path C:/SkipBo/tools/EventAutomationDesktop/.claude/skills/run-event-automator \
  --model <the model id running the session> \
  --max-iterations 5 --verbose
```

It splits 60/40 into train and held-out test, runs each query three times for a
stable trigger rate, and picks the winner on the **test** split so a description
cannot win by overfitting the queries it was tuned against. Take
`best_description` from the JSON it returns and paste it into the frontmatter.

## What the current description was tuned for

Two failure modes the queries exposed on reading:

- **Over-trigger on edits.** The first draft said "after touching
  renderer/app.js, main.js, preload.js or style.css", which reads as a match for
  *"add a Characters tab to the sidebar"* or *"make the counts bold"* — both
  edits, not verification. Now the description leads with VERIFY and ends with an
  explicit "Not for editing UI code".
- **Under-trigger on diagnosis.** *"the app opens on a blank white window"* is
  squarely this skill's job and nothing in the first draft said so. "diagnosing a
  window that opens blank" is now named.

If you add queries, keep them concrete — real paths, real symptoms, the odd
lowercase or typo. Abstract prompts like "run the app" are too easy to
discriminate and make the set look better than it is.
