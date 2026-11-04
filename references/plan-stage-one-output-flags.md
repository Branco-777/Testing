# Plan: Stage 1 Stress Flags And SST Output Folder

Version 0.2 | 4 November 2026 | Revised after user review. Decisions 58 and 59 replace parts of Decisions 55 and 56; see "Revision 0.2".

## Overview

Stage 1 now reads two named credit flags from the Draft and checks them against the `I. Credit Stresses` grids before it writes anything. Each run writes to `<Draft folder>\SST Output_<suffix>\<stress folder>`, which replaces `Bonds Lite output\<timestamp>`. Today only `Credit` exists. A `Credit` folder is handled exactly as a timestamped run folder was: Stage 1 fills it, Stage 2 runs the Bonds Lite batch in it and Stage 3 writes the MACalc input to it. Stage 4 planning is paused, and its five open questions are retained.

## Confirmed Requirements (User, 4 November 2026)

1. Draft names `SST_Credit_DG_standalone` (`I. Control!D16`) and `SST_Credit_spread_standalone` (`I. Control!D17`) are manual TRUE/FALSE flags. Both TRUE means a run with both downgrade and spread stresses.
2. If the downgrade flag is TRUE and **any** of the four transition matrices in `I. Credit Stresses` is the identity matrix, stop.
3. If the spread flag is TRUE and **none** of the eight spread grids has a non-zero value, stop.
4. Each run creates `SST Output_<suffix>` beside the Draft. The suffix comes from `SST_output_folder_suffix` (`I. Control!D66`).
5. Inside it, create one folder for each stress whose flags are TRUE. For now, either credit flag TRUE creates `Credit`. Later flags (Interest, Inflation, Property, FX: `I. Control` D14, D15, D18, D19) will add more folders.
6. Treat `Credit` as the old timestamped run folder in all later stages.

## Steps

1. **Tests first (Supir Test/Implement).** [✓ Implemented — see addendum] Add the failing tests in "Test Strategy" to `tests/test_bonds_lite_input.py` and `tests/test_sst_stage_three.py`. Pass: the new tests fail for the expected reason; existing tests are unchanged.
2. **Flag and stress-selection validation** in `src/bonds_lite_input.py`: [✓ Implemented — see addendum]
   - Add `STRESS_OUTPUT_FOLDERS = {"Credit": ("SST_Credit_DG_standalone", "SST_Credit_spread_standalone")}` and `OUTPUT_SUFFIX_NAME = "SST_output_folder_suffix"`.
   - Add `TRANSITION_BLOCKS` (C:J for rows 9, 20, 31, 42) and `SPREAD_BLOCKS` (M:DI and DL:HH for those rows), both derived from `STRESS_BLOCKS`.
   - Add a pure `validate_stress_selection(flags, stresses)`. Stop if no flag is TRUE. If the downgrade flag is TRUE, stop when any transition block is exactly the identity: 1.0 on the seven rating diagonals and 0.0 everywhere else, including the D column. Name the offending blocks. If the spread flag is TRUE, stop when every value in all eight spread blocks is exactly 0.0.
   - Read each flag through its workbook name. Stop if a name is missing or its value is not an Excel boolean.
   - Extend `StageOneInputs` with `stress_flags` and `output_suffix`.
   - Run all checks in `validate_inputs`, before any folder is created. Pass: unit tests pass, and no folder is created when validation fails.
3. **Output folder allocation** (depends on step 2): [✓ Implemented — see addendum]
   - Replace `allocate_run_folder(output_root, timestamp)` with `allocate_stress_folder(caller_folder, suffix, folder_name)`.
   - Validate the suffix. It must be non-blank text, or an integer-valued number (written without `.0`). It must not contain `<>:"/\|?*` or control characters, must not end in a dot or a space, and the resulting path must stay inside the Draft folder.
   - Create `SST Output_<suffix>` if missing. Create `<stress folder>` exclusively; if it already exists, stop before any write and ask the user to change the suffix. Earlier outputs are never overwritten.
   - On failure, remove the incomplete stress folder as today. Also remove the root if this attempt created it and it is now empty.
   - `_generate_bonds_lite_input` uses `Credit`. Pass: unit tests cover a new root, an existing root without `Credit`, an existing `Credit` and cleanup.
4. **Handoff validation** (depends on step 3): change `portable_output_path`. [✓ Implemented — see addendum]
   - Accept only a real, non-symlinked folder whose name is a key of `STRESS_OUTPUT_FOLDERS`. Its parent must be named `SST Output_<non-empty suffix>` and must sit directly in the Draft folder. Keep the Bonds/MAP completeness check.
   - Legacy `Bonds Lite output\<timestamp>` handoffs are rejected with a message to rerun Stage 1.
   - Stage 2 and Stage 3 resolution (`resolve_stage_input`) keep working unchanged through this function. Pass: Stage 2/3 resolution tests use the new layout.
5. **Stage 3 filename** in `src/sst_stage_three.py` (depends on step 4): [✓ Implemented — see addendum] `output_filename` builds `MACalc input_<suffix>_<16 hex>.xlsb`, taking the suffix from the parent folder name. Remove `RUN_ID_PATTERN`. The staging folder becomes `SST Output_<suffix>\Stage 3 staging\<token>`, outside `Credit`. The 218-character limit and the fresh-token behaviour are unchanged. Pass: filename and path-limit tests pass.
6. **Documentation** (parallel with step 5): [✓ Implemented — see addendum] update `references/bonds-lite-setup.md` (configuration table, output location, Stage 1-3 handoff) and this plan's progress addendum.
7. **Verification:** [✓ Automated suite run — live Excel run pending; see addendum] run the targeted tests, then the full suite. Expected baseline: 885 passed, 85 skipped and the same 2 known failures, plus the new tests. A live Excel run on the Draft needs the user's approval.

## Test Strategy

- `validate_stress_selection`:
  - both flags FALSE → stop;
  - downgrade TRUE with one identity block (blocks 1-4 in turn) → stop, naming that block;
  - downgrade TRUE with a near-identity matrix (one off-diagonal 0.001, or D column non-zero) → pass;
  - spread TRUE with all eight blocks zero → stop;
  - spread TRUE with one non-zero cell in any one of the eight blocks → pass;
  - downgrade FALSE with identity matrices → pass;
  - spread FALSE with zero grids → pass.
- Flag reading: a missing name, a text "TRUE", a number 1 or a blank → stop; real booleans → accepted.
- Suffix: blank, `a/b`, `a:`, `name.`, `name ` and `..` → stop; `a`, `Q2 2026` and `7` (from the float 7.0) → accepted.
- `allocate_stress_folder`, using `tmp_path`: creates `SST Output_a\Credit`; reuses an existing root; stops on an existing `Credit` and leaves its contents untouched; cleanup removes only the folders it created.
- `portable_output_path`: accepts `SST Output_a/Credit` with both inputs present. It rejects a legacy timestamp folder, a non-stress folder name, a wrong parent prefix, a nested location, a symlink and a missing input.
- Stage 3: `output_filename(Path(".../SST Output_a/Credit"), token)` returns `MACalc input_a_<token>.xlsb`; a bad token still stops.
- Data: synthetic grids built from the 7×8 identity and 7×101 zero blocks. No real data is needed.

## Affected Files

- `src/bonds_lite_input.py`: flags, suffix, selection validation, folder allocation and handoff validation.
- `src/sst_stage_three.py`: output filename and run-folder assumptions.
- `tests/test_bonds_lite_input.py`, `tests/test_sst_stage_three.py`, `tests/fixtures/*`: new and updated tests.
- `references/bonds-lite-setup.md`, `references/decisions.md`, this plan.
- Not changed: the Draft workbook, the VBA launcher, packaging and release pointers.

## Decisions

See Decisions 55-57 in `decisions.md` (agent-made, pending user approval).

## Risks And Unknowns

- *(Superseded by Revision 0.2: FALSE flags are now checked and warned about; grids are still copied unchanged.)*
- *(Superseded by Revision 0.2: reusing a suffix now overwrites.)*
- Existing Draft audits that point to `Bonds Lite output\...` stop being valid handoffs, so Stage 1 must be rerun once.
- MACalc Batch skips filenames containing `_opt`, so a suffix containing `_opt` would hide the Stage 3 output from it. This is a Stage 4 concern; record it in the Stage 4 plan.
- The packaged release (v8) does not include this change. A rebuild or promotion needs separate approval. (Resolved: released as `20261104-cache-v9`; see "Commit and release".)

## Addendum: Implementation Progress

**Implementation session (Supir Implement, autonomous).** Steps 1-6 completed; step 7 automated verification completed.

- **Tests first:** new tests were added and seen to fail before the code changed: `validate_stress_selection` (no flag, each identity block, near-identity, all-zero spreads, one non-zero cell per spread block, unselected neutral grids); flag reading (missing name, text "TRUE", 1, 0, 1.0, blank, real booleans); suffix validation (blank, `a/b`, `a:`, `name.`, `name `, `..`, 7.5, TRUE and accepted `a`, `Q2 2026`, 7.0 → `7`); `allocate_stress_folder` (new root, existing root, existing `Credit` untouched, unsafe suffix, unregistered folder, failure removes only a new root); `remove_incomplete_stress_folder`; `portable_output_path` (accepts `SST Output_<suffix>/Credit`; rejects legacy timestamp, non-stress name, wrong prefix, nested, symlink, missing input); invalid selection creates no folder; and Stage 3 `output_filename`.
- **Code:** `src/bonds_lite_input.py` gained `TRANSITION_BLOCKS`, `SPREAD_BLOCKS`, `CREDIT_DG_FLAG`, `CREDIT_SPREAD_FLAG`, `STRESS_OUTPUT_FOLDERS`, `OUTPUT_SUFFIX_NAME`, `OUTPUT_ROOT_PREFIX`, `validate_stress_selection`, `validate_output_suffix`, `read_stress_flags`, `read_output_suffix`, `allocate_stress_folder`, `remove_incomplete_stress_folder` and `output_root_suffix`. `StageOneInputs` gained `stress_flags` and `output_suffix`. `allocate_run_folder` was removed. `src/sst_stage_three.py` drops `RUN_ID_PATTERN` and names outputs `MACalc input_<suffix>_<token>.xlsb`.
- **Existing tests:** fixtures and tests that assumed `Bonds Lite output\<timestamp>` or run-ID filenames now use `SST Output_<suffix>\Credit`. The opt-in live Excel caller fixture now defines the two flag names and the suffix name in its disposable workbook. These live tests were not run.
- **Verification:** full suite gave 977 passed, 87 skipped and the 2 known failures. The 2 extra skips are directory-symlink tests that this Windows profile cannot create.
- **Deviations from the plan (implementation detail only, no design change):**
  - `allocate_stress_folder` returns `(folder, root_created)`, and a separate `remove_incomplete_stress_folder(folder, root_created)` handles cleanup. This was needed so that cleanup removes the root only when this attempt created it (Decision 56).
  - Flags and the suffix must be single cells on this Draft's `I. Control`, matching the existing Stage 1 path-name ownership checks.
  - The already-unused `choose_output_path` and `publish_output` were left in place because this change did not make them unused.
- **Remaining:** user approval of Decisions 55-57; a live Excel run on the Draft (needs approval); packaging/promotion (separate approval).

## Revision 0.2: Warn-And-Ask Consistency And Overwrite (User, 4 November 2026)

User review outcome:
- **Approved:** both flags FALSE stops the run; exact comparisons; the suffix rules; rejecting legacy handoffs (the user will rerun every stage at the end).
- **Changed:** inconsistencies now produce a warning and a prompt instead of stopping the run, and FALSE flags are checked too (Decision 58). Reusing a suffix now overwrites the stress folder (Decision 59).

### Steps

8. **Tests first** for steps 9-10, in `tests/test_bonds_lite_input.py`. [✓ Implemented — see addendum] Pass: the new tests fail for the expected reason before the code changes.
9. **Consistency check** in `src/bonds_lite_input.py`: [✓ Implemented — see addendum]
   - Split `validate_stress_selection` into two parts. A stop-only check keeps the two hard stops: both flags FALSE, and flag reading failures.
   - Add a pure function `find_stress_inconsistencies(flags, stresses) -> list[str]`. It returns one readable line for each Decision 58 inconsistency, naming the flag and the block addresses.
   - Add `confirm_stress_inconsistencies(lines) -> bool`. With no lines it returns True and does not prompt. Otherwise it logs one warning listing the lines, then asks `Proceed with these inconsistencies? [y/N]: ` through `input()`, following the `confirm_model_overwrite` pattern. Only `y` or `yes` (case-insensitive) returns True. `EOFError` returns False.
   - In `run_stage_one`, run validation and the prompt before the `Running` audit write.
     - On cancel, return `None` without creating, renaming or deleting any folder and without writing the audit. `main` then logs `Stage 1 cancelled; no output was written`, as it does for Stages 2 and 3.
     - Validation errors still record a `Failed` Stage 1 audit, as today.
     - Generation must use the validated `StageOneInputs` snapshot instead of reading the Draft again. That guarantees the confirmed stresses are the ones written, even if the user edits the sheet while the prompt is open.
   - Pass:
     - unit tests for all four inconsistency kinds, for consistent inputs giving no lines, and for each block named;
     - prompt tests for `y`, `YES`, blank, `n` and `EOFError`;
     - a `run_stage_one` test proving that cancel leaves the folders and the audit unchanged.
10. **Overwrite** in `allocate_stress_folder` and the cleanup: [✓ Implemented — see addendum]
    - If `<root>\<stress folder>` exists, it must be a real, non-symlinked folder; otherwise stop. Rename it to `<root>\~replacing <stress folder> <YYYYMMDD_HHMMSS>`, adding `_NNN` if that name is taken, then create a fresh folder.
    - If the rename fails with `OSError`, stop before any write. The message names the folder and asks the user to close files open from it and rerun.
    - On success, delete the backup. If that deletion fails, log a warning naming the backup; the run still succeeds.
    - On failure or `BaseException`, remove the new folder and rename the backup back. If restoring fails, log an error naming the backup path.
    - Return the backup path, or `None`, so that `_generate_bonds_lite_input` can commit or roll back.
    - Pass: tests for replacing an existing `Credit` that holds files, success deleting the backup, failure restoring the original contents byte for byte, a blocked rename stopping with the old folder untouched, a symlinked or file `Credit` being rejected, and a new root still being removed on failure.
11. **Downstream safety** after a replace: [✓ Implemented — see addendum] confirm with a test that Stage 3 refuses a `Credit` folder whose Stage 2 results were removed by a Stage 1 replace, so stale results can never be used. Make its message say to rerun Stage 2. Stage 2 needs no change: its existing overwrite consent finds no earlier results in a fresh folder.
12. **Docs:** [✓ Implemented — see addendum] in `references/bonds-lite-setup.md`, update the configuration and troubleshooting text: the inconsistency prompt, and that the same suffix overwrites, so change the suffix or archive the folder to keep earlier outputs. Remove the "existing Credit stops Stage 1" troubleshooting row.
13. **Verification:** [✓ Automated suite run — live run pending; see addendum] targeted tests, then the full suite. Expected: 977 passed, 87 skipped and the same 2 known failures, plus the new tests. The live run happens when the user reruns everything at the end.

### Risks
- Overwriting deletes earlier Stage 2 results and Stage 3 MACalc inputs in that folder. This is intended (Decision 59).
- On a OneDrive folder, a sync client briefly holding a file can make the rename fail. Stage 1 then stops safely, and the user reruns.
- The Stage 2 and Stage 3 audit panels can still show "last output" paths that a later Stage 1 has replaced. Rerunning the stages refreshes them.

### Revision 0.2 implementation

**Session (Supir Implement, autonomous), 4 November 2026.** Steps 8-13 completed; the live run is left for the user's end-to-end rerun.

- **Tests first:** the new tests were seen to fail (81 failures: missing functions, old tuple shape, old stop behaviour) before the code changed.
- **Code:** `src/bonds_lite_input.py` gained `find_stress_inconsistencies`, `confirm_stress_inconsistencies`, `discard_replaced_stress_folder` and the private `_replacement_backup_path` and `_restore_replaced_stress_folder`. `validate_stress_selection` now stops only when both flags are FALSE. `allocate_stress_folder` returns `(folder, root_created, backup)`. `remove_incomplete_stress_folder` takes an optional `backup` and restores it. `_generate_bonds_lite_input(caller, settings=None)` uses the supplied snapshot. `run_stage_one` validates and prompts before the `Running` write and returns `None` on cancel; `main` logs `Stage 1 cancelled; no output was written`. `src/sst_stage_three.py`: the missing base-scenario CSV message now says to rerun Stage 2.
- **Tests changed:** the identity/all-zero stop tests became `find_stress_inconsistencies` tests; the "existing `Credit` stops Stage 1" test was replaced by replace/restore tests; the invalid-selection test keeps only the hard stops; the `stage_attempt` fixture now patches `validate_inputs` and checks that the backend receives the validated snapshot.
- **Verification:** full suite gave 1022 passed, 88 skipped and the 2 known failures. The extra skip is the new symlinked-`Credit` test, which this Windows profile cannot run.
- **Deviations (implementation detail only):**
  - The direct API `generate_bonds_lite_input` (no audit, not used by the launcher) still validates without the inconsistency prompt, so it stays non-interactive and its signature is unchanged. Only `run_stage_one`, the console path, prompts.
  - A `KeyboardInterrupt` during validation records an `Interrupted` audit, as a validation failure records `Failed`. An interrupt at the prompt itself writes nothing, like a cancel.
  - If another process creates `Credit` between the rename and the new `mkdir`, Stage 1 stops with "already exists; another run may be writing it" and logs the backup path if it cannot be renamed back.
  - The four new logging tests use an isolated recorder fixture rather than `caplog`, because earlier `main()` tests leave the module logger with `propagate=False` (the same cause as the known `rollback_failure` failure).

### Commit and release (4 November 2026)

- Committed as `84e2dae` (with the docs commit `2bd2f0e`) and pushed to `origin/main`.
- Built `20261104-cache-v9` from `84e2dae` without `--promote`, using the `controller-py313` and `model-py311` build environments and the same model folder as v8. Both self-tests passed (`MODEL_SELF_TEST_OK`, `CONTROLLER_SELF_TEST_OK`). Compared with v8, both bundles have the same file names and frozen module sets (controller 82 files, model 869 files). The v9 controller's `bonds_lite_input` contains `find_stress_inconsistencies`, `confirm_stress_inconsistencies` and `allocate_stress_folder`.
- The user approved the release; the central `current.txt` now points to `20261104-cache-v9`. Rollback is to `20261103-cache-v8`. The LocalAppData cache is expected to update when Excel next checks the pointer.
- Still to do: the user's end-to-end Stage 1-3 rerun, which also replaces the legacy `Bonds Lite output` handoff.
