# Session Handoff: Stage 2 Debugging, Read-Only Cleanup Fix And Live Model Output

Version 1.0 | Prepared 4 November 2026 | Next: commit the read-only fix, then implement `plan-live-model-output.md`

> **Start here in the next session:** read "Pending Actions" below, in order.
> HEAD is `9ac5979` (release v10 docs). Nothing from this session is committed.

## Where We Stopped (4 November 2026, ~17:00 UTC)

1. **Stage 2 failure diagnosed (no code change).** Both model runs completed, but
   "Validate both model results" failed with
   `Missing or empty Bonds model artifact: ...\JRL Bonds\py_bonds_input_file_.xlsx`.
   - Cause: the saved inputs' `output_suffix` (`I. Control!D25`) was **blank**.
     The model reads blank as `None` and writes `py_bonds_input_file_None.xlsx`
     (`BondsModelHandler.run_full_process`). The controller (`read_model_inputs`)
     turns blank into `""` and expects `py_bonds_input_file_.xlsx`.
   - The failure triggered the Stage 2 rollback, so that attempt's results were
     removed and any earlier results restored.
   - Reproduced with the real saved inputs and the model source, whose hash
     matches the v10 `release.json`.
   - **User's workaround:** set a non-blank output suffix before the next run.
2. **Live model output planned and approved.** See
   [plan-live-model-output.md](plan-live-model-output.md) v0.2 and Decisions
   78-81. Not implemented yet.
3. **Stage 1 "replaced output could not be deleted" warning: diagnosed and fixed
   (uncommitted).**
   - Cause: OneDrive sets `ReadOnly` on synced folders. Plain `shutil.rmtree`
     gets `WinError 5` on `os.rmdir` of a read-only directory. It was not an
     open file: the folder was empty and renameable, and no Excel workbook was
     open from it.
   - Fix: `discard_replaced_stress_folder` and `remove_incomplete_stress_folder`
     in `src/bonds_lite_input.py` now call
     `shutil.rmtree(..., onexc=_retry_readonly_backup_removal)`, the handler
     Stage 2 already uses.
   - Tests: two new regression tests (red before the fix, green after):
     `test_discard_replaced_stress_folder_removes_read_only_directories` and
     `test_remove_incomplete_stress_folder_removes_read_only_directories`. Also
     changed: the `denied_cleanup` stub in
     `test_paired_cleanup_failure_preserves_generation_error_and_blocks_complete_pair`
     now accepts `*args, **kwargs`, because the new `onexc=` keyword broke the
     one-argument stub.

## Test Status (Verified)

- `tests/test_bonds_lite_input.py`: **819 passed, 84 skipped, 2 failed**.
- Both failures **also fail on clean HEAD** (checked in a temporary HEAD copy),
  so they are known baseline failures unrelated to this session:
  - `test_frozen_main_validates_cached_bundle_before_stage_dispatch`
    (`main() takes 0 positional arguments`);
  - `test_model_results_transaction_rollback_failure_retains_recoverable_backup`.
    It passes alone and fails only in the full-module run, so it depends on
    test order.
- `tests/test_sst_stage_three.py` and `tests/test_sst_stage_four.py` were not
  re-run this session.

## Uncommitted Changes From This Session

| File | Change |
|---|---|
| `src/bonds_lite_input.py` | Read-only-aware `rmtree` in the two Stage 1 cleanup functions |
| `tests/test_bonds_lite_input.py` | Two new regression tests; `denied_cleanup` stub signature |
| `references/plan-live-model-output.md` | New plan, v0.2 approved |
| `references/decisions.md` | Decisions 78-81 appended |
| `references/session-handoff-2026-11-04-live-output.md` | This file |

The `__pycache__` changes and the untracked `references/backups/*.xlsm` predate
this session and must not be committed with it.

## Pending Actions (In Order)

1. **Commit the read-only cleanup fix.** It is independent of the plan.
   - Suggested message: `fix: Clear OneDrive read-only folders when Stage 1 discards backups`.
   - Stage only `src/bonds_lite_input.py` and `tests/test_bonds_lite_input.py`.
     Optionally commit the plan, decisions and this handoff as a separate `docs:` commit.
2. **Delete the stale Stage 1 backups by hand** once the current Stage 2 run
   has finished. These are safe to delete; the current `Credit` folder is the
   latest run.
   - `SST Output_Test\~replacing Credit 20261104_163035`
   - `SST Output_Test\~replacing Credit 20261104_163237`
   - `SST Output_Test\~replacing Credit 20261104_165352`: left by the user's
     16:53 rerun, which used packaged v10 without the fix.
   - **Do not touch `SST Output_Test\.bonds-rollback-Credit-wumdo2t2` while
     Stage 2 is running.** It is Stage 2's rollback copy and is removed when
     Stage 2 finishes. If it remains after Stage 2 has ended, review it before
     deleting.
   - Source-mode runs no longer leave these backups. Packaged runs will until
     v11 is released.
   - Earlier stray `.bonds-lite-ekjlk9_a` working files in `Credit` were removed
     by the 16:53 rerun; no action needed.
3. **Implement `plan-live-model-output.md` Steps 1-6** (Decisions 78-81):
   - Step 1: `configure_console_streams()` in `src/bonds_lite_model_entry.py`.
   - Steps 2-3: `last_activity` and the 30-second quiet-heartbeat rule in `src/bonds_lite_input.py`.
   - Step 4: the same line buffering in `src/macalc_batch_entry.py` (approved, not optional).
   - Steps 5-6: tests, then a source-mode Stage 2 run with a non-blank suffix.
   - Write the tests in the plan's Test Strategy table first.
4. **Build and promote release `20261104-cache-v11`** (plan Step 7; needs separate
   approval). It bundles items 1 and 3, so acceptance covers the Stage 1 cleanup,
   live Stage 2 output, and Stage 4, whose v10 button acceptance is still pending.
   Rollback: write `20261104-cache-v10` to `current.txt`.
5. **Update `references/bonds-lite-setup.md`** (plan Step 8): v11 status, the live
   output note, and that `output_suffix` must not be blank (see item 6).

## Open Follow-Ups (Not Yet Approved)

6. **Blank `output_suffix` guard.** Optionally reject a blank suffix in
   `read_model_inputs` (Stage 2 preflight), or in Stage 1 validation, so the run
   fails in seconds instead of after the ~6-minute batch. Alternative: have the
   controller expect `py_bonds_input_file_None.xlsx` to match the unchanged
   model. This needs a user decision.
7. **Pre-existing test failures** listed above. Not caused by this session;
   investigate separately if wanted.

## Skills Consulted

- python-fact-grounded-coding, coding-standards, testing-strategy,
  regression-testing, decision-logging, document-extraction, questioning-patterns.
