# Event Final Spec -- Execution Reference

## Purpose & How to Use

This is the single source of truth the team builds from. The event is already
approved -- this document is not the place to re-justify the "why." Instead, it
tells each discipline exactly what to ship, by when, and to what quality bar.

How to use it:

1. One owner per section. The discipline lead fills in their own section.
2. Fill the placeholders written `[like this]`. Delete the prompts once you have answered them.
3. Link out -- do not duplicate. If a value lives in the config sheet or Jira, link to it rather than copying it.
4. Every section carries an `Owner:` and a `Status:`. Section status moves through:
   `TODO -> WIP -> For_Internal_Review -> Addressing_Feedbacks -> For_Internal_Review -> final_feedback_addressal->Locked`
5. Acceptance Criteria are the contract. If a criterion is not met, the section is not done.
6. Target read time is about 15 minutes start to finish. Keep entries tight.

---

## 0. Document Header / At-a-Glance
Approval confirmed: [ ] Yes -- this spec assumes a green-lit event.

Owner: [Designer/Lead]   Status: []

| Field 							| Value 
|-------							|-------
| Event Name (public) 				| [e.g., "Mansion Makeover"] 
| Event Name (internal) 			| [<Type><yy><mm>, e.g., `Sort2606`] 
| Event ID 							| [<yyyy><mm><##>, e.g., `20260601`] 
| Event Type 						| [Merge / Block / Sort / Hotel Sim / Management / Other] 
| Patch 							| [e.g., 8.06 refer to roadmap sheets] 
| Duration 							| [14 days standard / 21d/other -- if not 14, link justification] 
| Live Start -> End 				| [Date] -> [Date] 
| Proposal (approved) 				| [Link to approved Event Proposal msg] 
| Jira Epic 						| [Link] 	
| Configs / Economy-sheets 			| [Links] 
| Last similar event (apk + video) 	| [Links for reference] 

---

## 1. Scope & Non-Goals

Owner: [Designer]   Status: [ ]

### Decisions
- In scope: [Bullet the concrete deliverables this event ships -- e.g., "5-zone scene, 2 cosmetic frames, themed event packs, trivia set."]
- Out of scope / Non-Goals: [What this event explicitly does NOT include -- e.g., "No new core mechanic. No store redesign. No new tutorial."]
- Reused vs. new: [What is config-only refresh vs. genuinely new build. New = risk; call it out.]

### Acceptance Criteria
- [ ] Every in-scope item maps to a Jira ticket.
- [ ] Non-goals are explicit so no discipline silently expands scope.
- [ ] All "new" items are flagged (they drive risk in Section 10).

---

## 2. Design / Creative Spec

Owner: [Designer / Design Lead]   Status: [ ]

### Decisions
- Theme & narrative hook: [One line -- approved theme. Link to mood board.]
- Core mechanic: [Standard for this event type -- link the common mechanic page, do not re-explain.]
- Unique twist / refresh this time: [What differs from the last event of this type. If nothing, say "config refresh only."]
- Difficulty & pacing: [How difficulty scales; where the walls are; intended completion curve.-- link the common mechanic page, do not re-explain if no changes.]
- Any A/B planned? what & why

### Specs & Links
- Level / chapter count: [#]
- Scene zones: [# -- typically 4-5]
- Buildings per zone: [# -- typically 7-9]
- UX wireframes / prototype: [Link -- clickable prototype preferred]
- All texts for localization text source: [Link to copy doc / sheet]

### Asset Lists & details for them in Art requirements doc 
						   
| Asset 						| Type 			| New/Reuse 	| Description | Reference (optional		
|-------						|------			|-----------	| ---------------------------------		
| Event icon (bottom center) 	| UI 			| [New/Reuse] 	| 			  |
| Background art 				| Scene 		| [New/Reuse] 	|			  |
| Building / decoration sprites | Scene 		| [New] 		|			  |
| Card frame (cosmetic) 		| Cosmetic 		| [New] 		|			  |
| Card back (cosmetic) 			| Cosmetic 		| [New] 		|			  |
| Event pack art 				| Store 		| [New/Reuse] 	|			  |
| Unlock / reward screens 		| UI 			| [New/Reuse] 	|			  |

todo: add/update missing items here

### Acceptance Criteria
- [ ] All asset requirements are shared in the art requirements doc
- [ ] Mechanic flow documented or linked; no ambiguity for Engineering to config.
- [ ] All screens have wireframes or reference screenshots attached.
- [ ] Difficulty curve reviewed against the target archetype's skill level.

---

## 3. Engineering Spec

Owner: [Engineering Lead]   Status: [ ]

### Decisions
- Build type: [ ] Config-only  [ ] New code  [ ] Both
- Branch / environment: [Branch name + env]
- Feature flags / kill-switch: [Flag name; rollback plan -- required for any new mechanic]
- A/B setup: [If experimenting -- variant split, target segment]

### Specs & Links
- Config schema / keys: [Link to config sheet; list new keys if any]
- Event ID wiring: [Confirm Event ID `[YYYYMMDD]` set in config]
- Dependencies on other systems: [e.g., "Requires Chapter 82 live first"]
- Telemetry events to fire: [List event names -- coordinate with Section 5]
- Build / submission constraints: [SDK/XCode version, store submission lead time, server deps etc]

### Acceptance Criteria
- [ ] Kill-switch / rollback verified for every new mechanic.
- [ ] All config keys present in the shared sheet and named per convention.
- [ ] Telemetry events agreed with Product (Section 5) and fire in smoke test.
- [ ] Branch builds clean; no blocking platform/SDK issue.
- [ ] Any other tech callouts.
---

## 4. Art / Visual Spec

Owner: [Art Lead]   Status: [ ]

### Decisions
- Visual direction: [Approved mood/aesthetic -- one line. Link mood board.]
- New vs. reskin: [Which assets are net-new vs. recolor/reuse.]
- Animation scope: [Building placement, cosmetic reveals, transitions -- or "none new".]
- vfx scope: [Building placement, cosmetic reveals, transitions -- or "none new".]


### Acceptance Criteria
- [ ] Every asset in the list has an owner and a due date inside the timeline.
- [ ] Naming / export specs follow the standard handoff format.
- [ ] Feasibility confirmed by the Art Lead -- no asset is a "maybe."

---

## 5. Product / Analytics Spec

Owner: [PM]   Status: [ ]

### Decisions
- Primary KPI this event moves: [Engagement / Monetization / Balanced]
- Success targets (from approved Proposal -- do not re-derive):

| Metric 					| Baseline 		| Target 		| Source/calcualtions |
|--------					|----------		|--------		|--------|
| [D7 Retention] 			| [e.g., 35%] 	| [e.g., 38%] 	| [Link] |
| [Event engagement rate] 	| [TBD] 		| [TBD] 		| [Link] |
| [ARPU during event] 		| [TBD] 		| [+8%] 		| [Link] |
[todo : add relevant ones ...]

### Specs & Links
- Telemetry events required: [List -- must match Section 3 wiring.]
- Dashboard / report: [Link to performance dashboard or live report.]
- Post-launch review window: [Event start -> end + 2-3 days.]

### Acceptance Criteria
- [ ] Every target has a baseline and a measurement source.
- [ ] Telemetry list reconciled with Engineering (Section 3).
- [ ] Post-launch review is scheduled and owned.

---

## 6. Economy / Balance Spec

Owner: [Senior Economy Designer]   Status: [ ]

### Decisions
- Economy change vs. baseline: [ ] No change (standard rates)  [ ] Changes (list below)
- If changes: [Token earn rate, EUT costs, pricing deltas -- one line each, with reason.]

### Specs & Links
- Event token earn rate: [Tokens/level; estimated tokens/day; IAP bypass price] -- [Link to sheet]
- Scene building cost: [Buildings #, upgrades/building, EUT/building, total EUT to complete]
- Cosmetics & rewards: [Cosmetic list, unlock method, ruby/coin rewards + milestones]
- Balance sheet: [Link]

### Acceptance Criteria (Economy Health Check)
- [ ] Inflation risk: acceptable / flagged
- [ ] Progression pacing: smooth (not grindy, not trivial)
- [ ] Cosmetics pricing fair vs. gameplay value
- [ ] Economy Design reviewed and dated: [Date]

---

## 7. QA Spec / Test Plan

Owner: [QA Lead]   Status: [ ]

### Decisions
- Test scope: [Full pass vs. smoke-only -- justify if smoke-only.]
- Device / OS coverage: [List tiers/devices.]
- Regression scope: [What existing systems this could touch.]

### Test Cases (core)

| Area | Test | Expected | Status |
|------|------|----------|--------|
| Event entry | Tap event icon | Event loads, correct theme | [ ] |
| Token earn | Complete a level | Tokens credited per config | [ ] |
| Scene build | Spend EUT on building | Building upgrades, EUT deducted | [ ] |
| Cosmetic unlock | Hit milestone | Frame/back unlocked & equippable | [ ] |
| Store / packs | Open event pack | Correct contents, correct price | [ ] |
| Kill-switch | Toggle flag off | Event hidden cleanly, no crash | [ ] |
| End-of-event | Reach end date | Event closes, rewards settle | [ ] |

### Acceptance Criteria
- [ ] All core cases pass on covered devices.
- [ ] Edge cases (mid-event purchase, offline, timezone rollover) listed and tested.
- [ ] No open Sev-1/Sev-2 bugs at PO signoff.

---

## 8. Localization Spec (conditional)

Owner: [Designer / Loc coordinator]   Status: [ ] / N/A

> Fill only if this event ships new player-facing copy. If all strings are reused, mark N/A and skip.

### Decisions
- New strings this event: [ ] Yes  [ ] No (-> mark section N/A)
- Languages: [List target locales.]

### Specs & Links
- String sheet: [Link -- event title, zone names, flavor text, notifications.]
- Loc deadline: [Date -- must precede final build, fit in timeline.]

### Acceptance Criteria
- [ ] All new strings extracted to the loc sheet (no hardcoded text).
- [ ] Translations returned and integrated before final QA pass.
- [ ] Text fits UI in longest-language layouts.

---

## 9. Timeline, Dependencies & Capacity

Owner: [Prj.M]   Status: [ ]

### Milestones

| Milestone | Owner | Due | Status |
|-----------|-------|-----|--------|
| Design -> Art handoff | [Designer] | [Date] | [ ] |
| Art in-progress review | [Art Lead] | [Date] | [ ] |
| Design -> Eng handoff | [Designer] | [Date] | [ ] |
| Config locked | [Engineering Lead] | [Date] | [ ] |
| Loc returned (if applicable) | [Loc] | [Date] | [ ] |
| QA smoke test | [QA Lead] | [Date] | [ ] |
| Full QA pass | [QA Lead] | [Date] | [ ] |
| Cross-discipline review | [All Leads] | [Date] | [ ] |
| PO signoff | [Product Owner] | [Date] | [ ] |
| Store submission (if needed) | [Engineering] | [Date] | [ ] |
| Go-live | [Engineering] | [Date] | [ ] |

[Todo L review these milestones and update]

### Dependencies
- Internal: [e.g., "Chapter 82 must be live first." Each dependency = a Jira link.]
- External: [HQ approval; store review (2-3 wks); SDK/platform.]

### Capacity
- At-risk allocations / leaves overlapping this build: [Note who, when.]

### Acceptance Criteria
- [ ] Every milestone has an owner and date.
- [ ] Each dependency has a Jira ticket with the dependency marked.
- [ ] Store submission planned ~4 weeks before go-live (if a binary submission is needed).

---

## 10. Risks & Open Questions

Owner: [PM / Designer]   Status: [ ]

### Risks

| Risk | Impact | Mitigation | Owner | Status |
|------|--------|-----------|-------|--------|
| [Description] | [H/M/L] | [Plan -- e.g., A/B, kill-switch, rollback] | [Name] | Active |
| [Description] | [H/M/L] | [Plan] | [Name] | Active |

### Open Questions
- [ ] [Question] -- Owner: [Name] -- Needed by: [Date]
- [ ] [Question] -- Owner: [Name] -- Needed by: [Date]

---

## 11. Sign-off & Change Log

Owner: [PM]   Status: [ ]

### Review Ring (5-step -- required before handoff)
- [ ] 1. Self-review -- [Designer], [Date]
- [ ] 2. Peer review (optional) -- [Name], [Date]
- [ ] 3. Lead review -- Design Lead, [Date]
- [ ] 4. Cross-discipline -- Art [Date] / Engineering [Date] / Economy [Date] / QA [Date] / Product [Date]
- [ ] 5. PO signoff -- Product Owner, [Date] -- Approved? [Yes / No / Needs Changes]

### Change Log

| Date | Author | Change | Section |
|------|--------|--------|---------|
| [Date] | [Name] | [What changed and why] | [#] |

---

Template Version: 1.0
Last Updated: 2026-06-28
