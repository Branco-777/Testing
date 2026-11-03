# Session Handoff: Bonds Lite (2 November 2026)

## User Goal

Make the Stage 1 Bonds, MAP and shared-template path controls movable without
breaking Stage 1 or its saved Stage 2 handoff. The user does not need to create
workbook names manually. Stage 3 design is deferred until the user supplies its
requirements.

## Implemented

- Added workbook names `SST_Stage1_BondsDataPath`,
  `SST_Stage1_MAPDataPath` and `SST_Stage1_InputTemplatePath`.
- Stage 1 resolves paths only through those names; there is no runtime fallback
  to D39, D40 or D42.
- The installer migrates a legacy Draft only if all three exact adjacent labels
  match. It creates all names together, rolls back partial additions, preserves
  a complete existing set of moved names, and rejects partial/invalid names.
- The installer writes default Bonds/template paths through their names and
  preserves existing values.
- Update the layout by using Excel Cut/Paste so the defined name follows the
  cell. Copying a cell does not move its name; do not delete or rename names.

## Verification

- Focused named-control tests: **22 passed**.
- Disposable Draft installer Excel test: **2 passed** according to JUnit. Excel
  printed the known `0x800706ba`/`0x800706be` teardown diagnostics afterward.
- Full non-Excel suite: **664 passed, 8 skipped, 2 failed**. The unchanged
  failures are `test_frozen_main_validates_cached_bundle_before_stage_dispatch`
  (`main()` argument mismatch) and
  `test_model_results_transaction_rollback_failure_retains_recoverable_backup`
  (recovery-path assertion/capture).
- Static diagnostics and `git diff --check` passed.
- No Stage 1 or Stage 2 run was performed for this change.

## Deployment

- Replacement release `20261102-cache-v4` is active in both central and
  LocalAppData pointers.
- The 984-file package is 332,231,992 bytes. Robocopy reported 984 copied,
  zero mismatches, failures or extras. Controller/model executable hashes
  match the manifest; both v4 self-tests passed from staging and LocalAppData
  using the packaged clean environment. Model-source hashes are unchanged.
- The six relevant controller functions embedded in v4 match tested source.
- Draft migration succeeded and preserved existing D39/D40/D42 values. Backup:
  `references/backups/Draft Input spreadsheet for SST_20261102_102626_818697.xlsm`.
- Prior releases v1-v3 remain available. No financial model batch was run.

## Repository State

- Base implementation was committed earlier as
  `66659f31254cc8576b9c7c63b1d7e3c3a9b7bb43`
  (`feat: Add portable Bonds Lite Stage 1 and 2 runtime`); it was not pushed.
- Named-control source, test, Draft and documentation changes were made after
  that commit and are uncommitted. Check `git status` before further staging.
- Preserve unrelated/generated files. Do not delete or revert Bonds Lite output
  folders, `.pytest_cache`, `__pycache__`, inspected `REM_v1.10.xlsb`, or
  timestamped Draft backups without explicit authorization.

## Next Session

1. Check `git status --short` and inspect the named-control diff; preserve all
   existing outputs and backups.
2. Decide whether to commit the named-control change separately from the prior
   portable-runtime commit. Do not rebuild/re-promote v4 unless source changes
   after this handoff require it.
3. If continuing toward Stage 3, first obtain the user's detailed requirements.
   Keep Stage 3 business logic in a separate module and consume the validated
   Stage 2 run-folder contract; do not implement from assumptions.
