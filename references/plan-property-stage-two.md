# Plan: Property Stage 2 - Run LTM Flash

Version 1.0 | 7 November 2026 | Scope decisions confirmed; implementation and release approval remain separate

## Overview

Property Stage 2 runs the LTM Flash model against the workbook committed by
Property Stage 1. It opens that workbook in a separate visible Excel instance,
recalculates without updating external links, and invokes the LTM macro used by
the existing multi-run button's option 2. The model updates tabs in the same
workbook; Stage 2 creates no output workbook or folder and saves the workbook
only after the agreed success checks pass.

## Verified LTM Flash Run Contract

The `.xlsb` workbook was inspected read-only via Excel on a temporary copy. Its
original and the Stage 1 output were not changed during discovery.

- The multi-run button is `CommandButton2` on `I. Control`. Its event procedure
  displays the 0/1/2 prompt; option 2 calls
  `Module1.CallRunFrozenPython "flipper"`.
- Stage 2 must **not** click the button or reproduce the prompt. It should call
  the public `CallRunFrozenPython` macro once with `"flipper"`, which runs the
  same frozen LTM executable path and flipper run as option 2.
- The macro sets `RunStartTime`, clears prior progress/output content, runs
  synchronously, and then sets `RunEndTime`. It clears `O. CFOutput`,
  `O. PVOutput`, `O. ConcRisk`, the `O. NoteValRebase` table from `A4`, and
  several result ranges in `O. NoteVal`.
- The macro itself does not save the workbook and does not provide a reliable
  success return to Stage 2. A non-zero executable exit displays a VBA error
  dialog; after it is dismissed, the macro can return normally. Therefore,
  Stage 2 must validate the run timestamps, progress log, and outputs rather
  than trust the VBA call's return value.
- The current LTM macro uses
  `\\server090\internal model$\7 ALM\2. Model Management\Model Repository\LTMflashArunmanager_v3.2a(d9bb4fbe).exe`.
  The executable was reachable during discovery. Stage 2 must check that the
  expected executable is available before opening the LTM workbook.
- The Stage 1 output currently has `prm_LocalRun=TRUE` and
  `prm_DataSource=SQLserver`; these settings avoid the macro's remote-run
  rejection path.
- The inspected Stage 1 output had 16 flipper scenarios and took approximately
  45 minutes for a prior run. This is an observed runtime, not a guarantee.
- The template contains external workbook links and JRC-backed formulas. Stage
  2 must open with link updates disabled, then recalculate in the new Excel
  instance before running the macro.

## Confirmed Requirements And Decisions

1. Use a new, separate **visible** Excel instance for the LTM workbook; leave
   the Draft's Excel instance open and usable.
2. Read the input only from the latest successful Stage 1 handoff,
   `SST_PropStage1_LastSuccessfulOutput`. Require the Property flag to be
   Excel `TRUE` and Stage 1 status to be `Succeeded`.
3. Recalculate without updating external links before calling the LTM macro.
4. Run option 2 by invoking `CallRunFrozenPython("flipper")` directly. Do not
   click `CommandButton2`, show its input prompt, or implement another model
   runner.
5. On a macro error or failed success check, close the LTM workbook without
   saving. On success, save the same workbook in place and verify that Excel
   reports it saved and its file modification time changed.
6. Accept these success checks, in addition to requiring the expected LTM
   workbook remains open and readable:
   - `RunStartTime` is fresh for this invocation and `RunEndTime` is at or
     after it. Capture the invocation time from the same Excel instance so
     Excel's timestamp precision is handled consistently.
   - The progress log contains `Run N - End of run!` for every scenario from
     1 through the current `prm_NoOfScens`, and its last non-blank message is
     exactly `End of multirun!`.
   - `O. CFOutput`, `O. PVOutput`, and `O. ConcRisk` each contain at least one
     non-blank cell after the run. Do not add a row-1 run-number check yet.
   - After saving, `Workbook.Saved` is true and the file modification time is
     later than the pre-run value.
7. Set no Stage 2 timeout. The visible Excel instance may display a model error
   dialog that needs the user's attention; do not force-close Excel or kill
   the model process.
8. Allow the macro's normal clearing of `O. NoteVal` and `O. NoteValRebase`.
   Stage 2 does not snapshot, restore, validate, or populate those areas.
9. Add the Property Stage 2 button and audit panel below the Property Stage 1
   controls, following the Credit Stage 2 placement, occupancy, and
   move/resize-safe patterns.
10. Use CLI token `P2`, Draft macro `RunLTMFlash`, button `btnRunLTMFlash`,
    caption `Run LTM Flash`, panel `SST_PropStage2_Panel`, and panel heading
    `Property Stage 2 - Run LTM Flash`.
11. Add named audit fields
    `SST_PropStage2_{LastRun, Start, Finish, Elapsed, User, Status,
    InputWorkbook, LastSuccessfulRun}`. `Input Workbook` is a link to the
    Stage 1 workbook. `Last Successful Run` is a date/time, updated only after
    success and retained after a failed attempt.
12. Keep Property stages in their own modules, share the Draft run guard, and
    leave packaging/promotion for a separate, explicitly approved combined
    Stage 1 + Stage 2 release.

## Steps

1. **Write failing contract tests first.** [✓ Implemented — see addendum]
   Add `tests/test_sst_property_stage_two.py` using the existing fake workbook,
   audit, and run-guard patterns. Add only the CLI/launcher/package assertions
   that belong in `tests/test_bonds_lite_input.py`.
   - Cover flag and Stage 1 readiness gates; workbook path resolution and
     missing/unreadable/unsupported workbooks; the model executable preflight;
     fresh/stale run timestamps; missing scenario completion markers; wrong
     final progress message; empty output sheets; save verification; audit
     persistence; and failure cleanup.
   - Fake the Excel application and workbook; assert Stage 2 creates a distinct
     visible Excel instance, opens with `UpdateLinks=0`, recalculates before
     invoking exactly `CallRunFrozenPython("flipper")`, saves only after all
     success checks, and closes without saving on failure.
   - Assert `O. NoteVal` and `O. NoteValRebase` are not part of the success
     checks or rollback snapshot.
   - Pass criterion: new tests fail because Stage 2 behavior is absent, not
     because fixtures or Excel dependencies are unavailable.

2. **Implement Property Stage 2 input and audit contracts** in a new
   `src/sst_property_stage_two.py`.
   [✓ Implemented — see addendum]
   - Reuse `PROPERTY_FLAG`, `read_property_flag`,
     `require_property_selected`, `require_property_stage_one_panel`, and
     shared helpers from `sst_property_stage_one.py` and
     `bonds_lite_input.py` rather than duplicating their validation.
   - Add Stage 2 constants, field names, labels, and panel validation. Resolve
     the Stage 1 output path from the named last-successful-output field; do
     not infer the path from the output folder or newest file.
   - Require the Stage 1 status to be `Succeeded`, the handoff to be a
     non-empty existing workbook, and a macro-enabled format (`.xlsb` or
     `.xlsm`). Fail explicitly before launching Excel if a prerequisite is
     missing.
   - Check the expected LTM runner executable before opening the workbook.
     Also fail fast if the workbook's `prm_LocalRun` / `prm_DataSource`
     settings would cause `CallRunFrozenPython` to reject the run.
   - Use the shared Draft run guard. Persist a `Running` audit before starting
     model work; retain the previous `LastSuccessfulRun` unless a complete
     verified run succeeds.
   - Pass criterion: unit tests prove invalid flags, stale Stage 1 status,
     invalid handoffs, missing runner, and duplicate concurrent runs never
     invoke the model macro.

3. **Implement the separate Excel run and verification lifecycle** in
   `sst_property_stage_two.py`.
   [✓ Implemented — see addendum]
   - Create a new visible Excel application instance; do not reuse, close, or
     change the Draft's Excel application. Ensure the LTM workbook is not
     already open in another Excel instance before opening it for writing.
   - Enable VBA for this explicitly approved local model workbook, open it
     read/write with external-link updates disabled, and perform the requested
     full recalculation before the macro call. Preserve user visibility and
     normal Excel dialogs.
   - Capture the file modification time and Excel invocation time. Call the
     public workbook macro once with `"flipper"`; do not invoke the button
     event procedure.
   - On return, inspect `RunStartTime`, `RunEndTime`,
     `prm_NoOfScens`, the progress region anchored by `ProgressFirstCell`,
     and non-blank content on the three agreed output sheets. Ignore the
     `NoteVal` and `NoteValRebase` ranges; the macro is allowed to clear them.
   - If a macro exception or any check fails, close without saving and record
     `Failed` or `Interrupted`. The Stage 1 workbook on disk must remain
     unchanged. There is deliberately no timeout or automated process kill.
   - If checks pass, save in place, confirm `Workbook.Saved`, and confirm the
     file modification time advanced before recording `Succeeded` and updating
     `LastSuccessfulRun`.
   - If final Draft audit persistence fails after the verified LTM workbook
     was saved, report the two states explicitly, leave the verified workbook
     intact, do not claim a persisted Stage 2 success, and preserve the prior
     `LastSuccessfulRun`; follow the Stage 1 precedent of not discarding a
     completed output solely because its final audit save failed.
   - Always close only the owned LTM workbook/application and release the Draft
     run guard after cleanup. Do not close the Draft or unrelated Excel books.
   - Pass criterion: tests prove operation ordering, no-save rollback,
     output/audit status consistency, explicit exceptions, and cleanup on
     success, failure, and interruption.

4. **Install the Stage 2 controls idempotently** in the new Property module.
   [✓ Implemented — see addendum]
   - Add the named audit panel and its eight named fields below
     `SST_PropStage1_Panel`, at the same relative column and with a similar
     vertical gap and 11-by-8 geometry to Credit Stage 2. Place the new button
     in the matching button column and preserve the Stage 1 button's relative
     row/size.
   - Calculate placement from the current named Stage 1 panel and
     `btnPrepareLTMFlash`, not fixed coordinates. Check cells, merges, and
     shapes for collisions before writing.
   - Add or refresh `btnRunLTMFlash` and assign it to the new public Draft
     macro. Preserve a complete user-moved/resized Stage 2 panel and button on
     reinstall; reject partial named installations rather than create
     duplicates.
   - Insert `RunLTMFlash` into `modBondsLiteLauncher`; it must dispatch `P2`
     through the existing launcher. Keep the existing Draft backup-before-
     modification and closed-workbook installation rules.
   - Pass criterion: fake and opt-in Excel installer tests prove correct
     geometry, no-overwrite collision behavior, backup creation, idempotency,
     save/reopen, and preservation of moved/resized Stage 1 and Stage 2
     controls.

5. **Wire source and packaged dispatch** without changing existing stage
   behavior. [✓ Implemented — see addendum]
   - Add `--install-property-stage-two` and `P2` dispatch to
     `src/bonds_lite_input.py`; log the stage as Property Stage 2 and report
     that the existing LTM workbook was updated in place.
   - Add the `RunLTMFlash` launcher to both source and packaged VBA launchers.
   - Mirror `P2` and the installer command in `tools/controller_entry.py` and
     add `sst_property_stage_two` as a hidden import in
     `tools/build_packaged_stages.py`.
   - Add tests proving default dispatch remains Credit Stage 1, `P1` and
     stages 1-4 remain unchanged, `P2` dispatches only to Property Stage 2,
     both VBA variants contain the new entry point, and the packaged build
     includes the new module.

6. **Run targeted and regression validation.** [✓ Implemented — see addendum]
   - Run the focused Property Stage 2 tests, related Stage 1/launcher tests,
     and panel installer tests first; then run the default full suite and
     `git diff --check`.
   - Add an opt-in Excel test guarded by an environment variable, following
     `test_excel_property_stage_one_uses_temporary_draft_and_template_copies`.
     It should install/reinstall controls on a temporary Draft copy and verify
     names, geometry, launcher text, and save/reopen behavior. It must not
     launch the production LTM model.
   - Preserve Stage 1 coverage and the existing user-edited tests. Do not
     silently fix the two unrelated full-suite failures recorded in the
     handoff; report whether they remain blockers for the combined release.
   - Pass criterion: all new and related focused tests pass; the full suite has
     no newly introduced failures; any known baseline failures are unchanged
     and documented.

7. **Source-mode and combined-release acceptance** (explicit user check-in
   before any real LTM run or live Draft modification).
   - Back up the live Draft and eligible Stage 1 workbook. Install the Stage 2
     button only after automated tests pass.
   - With the Property flag TRUE and Stage 1 status `Succeeded`, use the
     installed Stage 2 button to run the real macro once. Verify the visible
     second Excel instance, full progress markers, three outputs, saved
     workbook, and Stage 2 audit. Do not start a run from planning or without
     the user's explicit go-ahead.
   - Test flag FALSE and stale/failed Stage 1 handoff gates without running the
     LTM macro. Restore the flag and confirm the existing Stage 1 output and
     audit remain intact.
   - Separately test the installed Stage 1 button path, which remains an
     outstanding Stage 1 acceptance item in the handoff.
   - Build the combined package from a clean, identified source state; verify
     both Property modules are included and test on representative copies.
     Obtain explicit approval before installing/promoting that package.
   - Pass criterion: actual button-launched Stage 1 and Stage 2 acceptance
     succeeds, combined package tests pass, and promotion has explicit
     approval. A source-mode run alone is not packaged-release approval.

## Test Strategy

1. **Unit tests first:** fake workbook/Excel objects test all gates, macro
   invocation arguments and ordering, timestamp parsing, scenario-count
   completion, terminal progress message, output non-emptiness, and save
   verification. Include success, VBA exception, partial run, dismissed-error
   shape (macro returns but markers are missing), interruption, and audit
   persistence failures.
2. **Installer tests:** use existing Stage 1 and Credit Stage 2 installer
   fakes for dynamic layout, collision checks, moved/resized controls, partial
   name rejection, and idempotent backed-up installation.
3. **Opt-in Excel tests:** use temporary Draft copies only; validate the
   controls and launcher without calling the production model executable.
4. **Live integration:** a real LTM run is slow and depends on Excel, the JRC
   add-in, external data, and the network executable. Run it only after
   automated gates and separate user approval; verify workbook state rather
   than treating process launch as success.
5. **Regression:** preserve all current Stage 1 and Credit test behavior.
   There is no model algorithm change or new numerical golden fixture in this
   stage, so do not claim actuarial output parity from the operational success
   checks.

## Affected Files

- `src/sst_property_stage_two.py` — new Property Stage 2 input validation,
  visible Excel lifecycle, macro invocation, success checks, audit, and
  installer.
- `tests/test_sst_property_stage_two.py` — new unit and fake-Excel contract
  tests.
- `src/bonds_lite_input.py` — `P2` dispatch, installer option, and source /
  packaged VBA launcher entry points.
- `tests/test_bonds_lite_input.py` — CLI, launcher, installer, and packaging
  regressions.
- `tools/controller_entry.py` — packaged `P2` and installer dispatch.
- `tools/build_packaged_stages.py` — hidden import for the new stage module.
- `references/bonds-lite-setup.md` — setup and run instructions, visible
  second Excel instance, external-link/recalculation behavior, no timeout,
  macro-cleared NoteVal ranges, success checks, and recovery guidance.
- `references/decisions.md` — Decision 97, confirmed scope and behavior.
- `references/handoff-property-stage-two-and-promotion.md` — link the Stage 2
  plan and replace the unresolved planning prompts with the agreed scope.
- `Draft Input spreadsheet for SST.xlsm` — changed only through the backed-up,
  tested installer; no live installation during planning.

## Decisions

- User-confirmed requirements are recorded in
  [`decisions.md`](./decisions.md), Decision 97.
- No team-specific conventions apply to this Stage 2 plan, as reconfirmed by
  the user. No team skills were applied.

## Risks And Unknowns

- Recalculation in a separate Excel instance depends on the required JRC add-in
  being available there. Validate this in the opt-in/live acceptance path;
  never update external links as a workaround.
- The external executable path is version-specific and hard-coded in the
  inspected template macro. Keep the preflight path aligned with the actual
  workbook macro and fail clearly if either changes.
- A model failure may present a modal dialog in the visible Excel instance.
  With the agreed no-timeout behavior, the user must dismiss or resolve it;
  Stage 2 must not kill Excel or the executable.
- The current full suite has two unrelated failures documented in the
  handoff. Decide separately whether they block combined release; do not fold
  unrelated fixes into this scope.
- The worktree contains pre-existing workbook, Credit-output, implementation,
  and live-test changes. Review `git status --short` before any implementation
  or installation; preserve unrelated changes and do not perform broad cleanup.

## Skills Consulted

- questioning-patterns
- document-extraction
- excel-analysis
- testing-strategy
- regression-testing
- decision-logging
- developer-guides

## Addendum: Implementation Progress

- Steps 1-5 are implemented. The new fake-Excel tests cover Stage 2 gates,
  timestamp/progress/output validation, invocation and save ordering,
  no-save failure handling, restoration when save verification fails, audit
  states, and source/package dispatch. The named panel installer, VBA entry
  point, and temporary-Draft Excel test are also implemented.
- Step 6 Property Stage 2 tests passed (**30 passed, 1 skipped**). The focused
  Property Stage 1/2 and launcher suites completed with **934 passed, 86
  skipped, 2 failed**. The full suite (with the repository root on
  `PYTHONPATH`) completed with **1,247 passed, 91 skipped, 2 failed**. The two
  failures are the existing unrelated failures
  recorded in the handoff: `test_frozen_main_validates_cached_bundle_before_stage_dispatch`
  and `test_model_results_transaction_rollback_failure_retains_recoverable_backup`.
  No new failures were observed. `black --check`, Python compilation, and
  `git diff --check` passed.
- The opt-in Excel installer test was attempted on a temporary Draft copy and
  did not complete: Excel automation terminated Python with COM error
  `0x80010108` while entering the installer. No production model was launched
  and the live Draft was not used. Rerun this test in a clean, responsive
  Excel session before relying on live installation.
- The run keeps a temporary pre-run copy outside the output tree so a failed
  post-save verification can restore the Stage 1 workbook byte-for-byte. This
  supports the requirement that failed Stage 2 attempts do not damage the
  Stage 1 handoff; it is removed during cleanup.
- Step 7 remains pending: no real LTM run, live Draft installation, or
  packaged build/promotion was performed. These require the separate explicit
  user approval stated in the plan. No deviation from Decision 97.
