# Session Handoff: Stage 4 MACalc Batch

Version 1.0 | Prepared 3 November 2026 | Next session: 4 November 2026

## Tomorrow's Goal

The user's instruction at the end of today's session:

> Stage 4 will be just exploring the ma_calc_batch model and make it run with
> the output handed off by Stage 3.

Start by understanding the consuming model, then connect its existing execution
path to the validated Stage 3 output. Do not redesign its financial calculations
or add unrelated functionality.

**Stage numbering and permission:** older Stage 3 documents and Decision 51 call
the consuming-model run "Stage 5". The user now explicitly calls this work
**Stage 4**. Decision 51 kept the model outside Stage 3's scope; tomorrow's request
authorises exploring it and implementing its Stage 3 handoff as Stage 4. No model
inspection, changes or execution were performed while preparing this handoff.
Reconcile the older numbering in the plan and decision log during Stage 4
planning; do not treat the Stage 3 restriction as a prohibition on tomorrow's work.

## Read First

- [Stage 3 plan](plan-stage-three.md), version 0.15, particularly Slices C and D
  and their implementation addenda.
- [Decisions](decisions.md), especially 49-54: the list contract, consuming-model
  boundary, CSV rating suffix and unique filename.
- [Stage 3 handoff](session-handoff-2026-11-02-stage-three.md). Its opening status
  is current; later sections describe older snapshots and are not current release
  instructions.
- [Setup guide](bonds-lite-setup.md), including deployment, Stage 3 handoff and
  ISIN list alignment.

## End-of-Day State

- Stage 3 Slices A-D are implemented and accepted.
- Final Slice D commit: `0fdf81a feat: Give each MACalc input a unique filename`,
  including the Draft, tests, documentation and the 218-character path limit.
- Earlier commits today: `ce799db` (controller-only Stage 3 hidden import) and
  `d4e94c9` (Slice D plan and v6 documentation).
- **Active central release:** `20261103-cache-v8`. Promotion was explicitly
  approved and the pointer was verified after the update.
- **Rollback release:** `20261103-cache-v7`. Do not roll back without approval.
- The user's packaged Stages 1-3 test passed on v7. v8 carries the subsequent
  218-character limit and passed build checks, but a user-run v8 workflow has not
  yet been confirmed. Excel updates the LocalAppData cache when it next checks
  the pointer; do not assume that the user's local cache is already v8.
- v8 model and controller self-tests passed. Compared with v6, the file and frozen
  module sets are identical: controller 82 files / 469 modules; model 869 files /
  1,770 modules. The Bonds Lite model does not bundle Stage 3.
- Full suite after the path-limit change: **885 passed, 85 skipped, 2 known
  failures**:
  - `test_frozen_main_validates_cached_bundle_before_stage_dispatch`
  - `test_model_results_transaction_rollback_failure_retains_recoverable_backup`
- All four opt-in Excel tests passed earlier in Slice D, before the final
  path-limit adjustment. Three live source-mode Stage 3 runs also passed.
- A packaged Stage 1 MAP import once failed with a `HeaderRowRange` error.
  Six repeat MAP imports succeeded and the user's retry passed. No Stage 1 fix
  was made; the cause was not proven. Investigate separately only if it recurs.

Central package root:

```text
\\server090\internal model$\29 Financial Resilience\Sandbox\SST spreadsheet executable file
```

## Stage 3 Handoff Contract

The input path must come from the Draft's **`SST_Stage3_LastSuccessfulOutput`**.
Never reconstruct the filename, choose by modification time or assume that
`MACalc input.xlsb` is the current input.

- Stage 3 names each output `MACalc input_<Stage 1 run ID>_<16 hex>.xlsb`.
  Every rerun gets a fresh token and keeps earlier successful outputs.
- Relevant audit names are `SST_Stage3_Status`, `SST_Stage3_InputRun` and
  `SST_Stage3_LastSuccessfulOutput`. Determine tomorrow how Stage 4 should handle
  a later unsuccessful Stage 3 attempt that has retained an older successful
  path; do not silently invent a stale-input policy.
- Stage 3 saves and closes the workbook, reopens it read-only for verification,
  and publishes it only after verification passes.
- Both the output path and staged path are limited to 218 characters.
- Staging is under `<output root>\Stage 3 staging\<token>\`, outside the run
  folder, and is removed afterwards. Do not use staged or incomplete files.

The latest packaged-run folder inspected on disk was `20261103_183144`, containing:

```text
Bonds Lite output\20261103_183144\MACalc input_20261103_183144_97a820f94d0861f6.xlsb
```

This is a **candidate for the first live test**, not a hard-coded handoff. The
Draft's audit was not reread when this document was created; resolve and validate
the actual audit path at the start of tomorrow's session.

**Folder-scanning risk:** the user confirmed that MACalc Batch scans the folder
holding the supplied input and does not ignore dot-prefixed files. A run folder
also contains Bonds/MAP input workbooks and can contain several tokenised MACalc
outputs. Explore its actual selection rules before running it. Supplying one path
must not accidentally run every earlier output or another workbook. Do not delete
earlier outputs to work around this.

## Model Location And Deferred Contract Checks

Decision 51 records the model directory as:

```text
C:\Users\Joao Branco\OneDrive - Just Group\Financial Resilience Team - General\00 Ad Hoc\MACalc_Batch
```

This is the recorded location, not a newly inspected directory. Confirm it tomorrow
and identify the actual entry point; `ma_calc_batch` is the user's name for the
model, not a verified Python module or command. The old `MACalc.py` in `v_3_3`
is not the authoritative consuming model.

Before the first run, verify the assumptions deferred by Decision 51:

- `tbl_Assets` is read dynamically, including all funded and asset rows.
- The following lists join to `I. Assets` column D by exact key:
  `tbl_ForcedCompAperc`, `tbl_MinCompAperc`, `lst_pre16assets` and `lst_REITs`.
- List lookup, duplicate-key and unmatched-key behaviour agree with the Stage 3
  contract. `lst_noPRoC` is unchanged by Stage 3.
- Rating-suffixed keys use the row's CSV `Final_Rating` (column H), not
  `Final_Notched_Rating` (column I). Stage 3 still writes the notched rating into
  `I. Assets` column I; do not strip `+`/`-` or infer the suffix.

Any mismatch needs a reported finding and an agreed correction, not a silent
change to the accepted Stage 3 output.

## First-Session Checklist

1. Inspect the current worktree and active/cache release without altering them.
2. Explore the MACalc Batch entry point, configuration, dependency manifests,
   workbook selection, other required inputs, output locations and side effects.
   Read any accompanying specifications or reference results.
3. Ask the user to confirm the execution surface: a source-mode proof first or
   immediate controller/button integration. The current controller CLI supports
   only stages 1, 2 and 3; Stage 4 dispatch is not implemented.
4. Confirm which scenarios to run, where outputs should go, whether existing
   outputs may be replaced, and what reference result demonstrates correctness.
   Identify required Excel/add-in/runtime dependencies rather than guessing.
5. Verify the deferred list contract and single-input selection rules above.
6. Agree the smallest integration plan. Keep financial model changes out of scope
   unless a specific incompatibility is found and the user approves its repair.
7. Write failing tests for the agreed handoff and launcher behaviour before
   implementation. Cover the exact audited filename, invalid/missing inputs,
   stale-run policy, extra workbooks/earlier outputs in the same folder, explicit
   model failure reporting and the expected output artefacts.
8. Run targeted tests, then a controlled live run on the audited Stage 3 output.
   Check expected financial results with the user's reference, not merely a zero
   exit code. Confirm that Stage 3's input and earlier outputs remain intact.
9. Only expand into packaging, Draft/VBA changes or release work if agreed.
   Preserve the controller/Bonds Lite model separation if packaging is required.

Use base coding and testing standards; no team-specific skills are active.
Suggested workflow: Supir Explore for bounded model discovery, then
Plan -> Test -> Implement -> Review -> Commit. The planning agent must hand off
production implementation rather than writing source code itself.

## Working Tree And Permissions

These tracked documentation files have **uncommitted** build/promotion updates
after `0fdf81a`; preserve them:

- [Setup guide](bonds-lite-setup.md)
- [Stage 3 plan](plan-stage-three.md)
- [Stage 3 handoff](session-handoff-2026-11-02-stage-three.md)

This new handoff is also uncommitted. Existing tracked deletions under
`Bonds Lite output\`, modified bytecode, caches, Draft backups and untracked
`.xlsb` templates are not part of a new change. Do not revert, delete or stage them.

Do not commit, push, promote or roll back without explicit approval. Use
Supir Commit when a commit is approved, with the required Co-authored-by trailer.
Do not change the Draft's controls/names or VBA without discussing the Stage 4
integration scope. Today's task only creates this handoff; it does not start
Stage 4 execution.

For the existing source test baseline:

```powershell
python -m pytest -q -p no:cacheprovider
```
