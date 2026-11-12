# Plan: Property Stage 4 - Run NoteVal

## Objective

Run NoteVal's `modControl.Run` macro on the workbook produced by the latest successful Property Stage 3 run. Verify that the macro completed, save the workbook in place, and record the run in a Property Stage 4 audit panel on the Draft. This follows Decision 113: Stage 3 prepares NoteVal and Stage 4 runs it, using the Property Stage 2 pattern.

## NoteVal `Run` macro (v1.40, read from a macro-disabled temporary copy)

`modControl.Run` makes these changes, in this order:
1. Unprotects 27 sheets.
2. Sets manual calculation.
3. Writes `I_StartTime_1` (`'I. Control'!D22`) = `Now()`.
4. Runs the five solvers:
   1. `LTMNotePricingExHedge` (GoalSeeks on `C. Note Valuation ex Hedge`).
   2. `LTMNotePricing` (GoalSeeks on `C. Note Valuation`, `P_NoOfNotes` + 1 columns).
   3. `NNEGOverlayRatings` (a `Do While G29 > 0` loop).
   4. `SingleEquivalentRating` (a `Do While F29 > 0` loop).
   5. `MACalcSolverOutputs` (a GoalSeek of `O. MAC MI!D12`).
5. Writes `I_EndTime_1` (`'I. Control'!D23`) = `Now()`.
6. Runs `Application.Calculate`, reprotects the sheets, then restores the status bar and calculation mode.

Neither rating loop can exit once the CQS reaches `P_Min_CQS_Mapping` while the residual stays positive. A run can therefore hang indefinitely, so Stage 4 needs a time-limit watchdog.

## Design

- **Input (Decision 126):** Stage 4 runs only when all of these hold:
  - `SST_PropStage3_Status` = `Succeeded`.
  - `SST_PropStage3_LastSuccessfulOutput` resolves to an existing workbook.
  - Its Stage 3 input (`SST_PropStage3_InputWorkbook`) equals the Stage 2 input workbook.
  - `SST_PropStage2_LastSuccessfulRun` is no later than `SST_PropStage3_Finish`. If it is later, Flash was rerun after NoteVal was prepared, and Stage 3 must be run again.
- **Run in place with restore (Decision 127):** Stage 4 follows the Stage 2 pattern:
  - Copy the NoteVal workbook to a temporary backup.
  - Open it in a separate, visible Excel instance owned by Stage 4, with macros enabled, AutoSave off and `update_links=False`.
  - Run the macro, verify the result, then save. If any failure leaves the file changed, restore the backup and check its hash.
- **Links (Decision 128):** Stage 4 does not update external links. The macro uses the Flash values cached and verified by Stage 3, and the cached values for the other five sources. This matches Stage 2 and avoids OneDrive and permission prompts during an unattended run.
- **Completion checks (Decision 129):**
  - `I_StartTime_1` and `I_EndTime_1` are dates.
  - `I_StartTime_1` is no earlier than Excel `Now()` captured immediately before the call, less a two-second tolerance.
  - `I_EndTime_1` is no earlier than `I_StartTime_1`.
  - Each GoalSeek target cell is a finite number and not an Excel error:
    - `O. MAC MI!D12`.
    - On `C. Note Valuation ex Hedge`: `F12`, `H12`, `AB31:AM31`, `AF45`, `AQ33:AQ36` and `AB57:AM57`.
    - On `C. Note Valuation`: `T23`, `T50` and `T56`.

  Residual tolerances are not checked yet: the monetary scale is unknown, so a fixed tolerance could cause false failures. Residuals are logged for review.
- **Watchdog (Decision 130):**
  - The time limit is 30 minutes; a typical run takes about 9.
  - An `SST_P4_TIME_LIMIT_MINUTES` override is available only when running from source.
  - Every 5 minutes it logs a heartbeat showing elapsed time.
  - At the limit it prompts "[W]ait 30 more minutes or [A]bort". Abort terminates only the Excel PID owned by Stage 4, and the backup is then restored.
  - It reuses the Stage 2 console-prompt and PID-termination helpers.
- **Controls (Decision 131):**
  - Panel `SST_PropStage4_Panel` = `'I. Control'!AB48:AI58`.
  - Fields `SST_PropStage4_{LastRun, Start, Finish, Elapsed, User, Status, InputWorkbook, LastSuccessfulRun}`.
  - Button `btnRunNoteVal` ("Run NoteVal"), placed beside the panel in the Stage 2/3 button column. Its VBA `RunNoteVal` calls `LaunchBondsLite "P4"`.
  - CLI options `--stage P4` and `--install-property-stage-four`.
  - Packaging hidden import `sst_property_stage_four`.

## Steps

1. **Tests first** (`tests/test_sst_property_stage_four.py`), all with test doubles: [✓ Implemented — see addendum]
   - Input resolution (T1-T5): accepts a valid handoff; rejects a non-Succeeded Stage 3, a missing output, a Stage 3 input that does not match Stage 2, and Stage 2 rerun after Stage 3.
   - Completion checks (T6-T9): accepts valid times and residuals; rejects a stale start time, an end before the start, and an error or non-numeric residual.
   - Run sequence (T10-T13):
     - Success saves and audits `Succeeded`.
     - A macro failure restores the backup and audits `Failed`.
     - A watchdog abort audits `Timed Out` without closing the terminated Excel.
     - A guard conflict is reported.
   - Watchdog time limit and override (T14-T15).
   - Installer panel, button and macro idempotence (T16).
   - CLI `--stage P4` dispatch (T17).
2. **Implement `src/sst_property_stage_four.py`** until T1-T16 pass. [✓ Implemented — see addendum]
3. **Wire the CLI and packaging** (`bonds_lite_input.py`, `tools/controller_entry.py`, `tools/run_stage_from_source.py`, `tools/build_packaged_stages.py`) until T17 passes. [✓ Implemented, built and promoted — see addendum]
4. **Opt-in Excel test (T18):** run P4 on a temporary copy of the real Stage 3 output and a temporary Draft copy. It is gated by `BONDS_LITE_RUN_PROPERTY_STAGE_FOUR_EXCEL_TESTS=1`. The measured run took 55 seconds and was not run against live inputs. [✓ Implemented and passed on temporary copies — see addendum]
5. **Install:** install on a temporary Draft copy first, then on the live Draft. The installer creates a backup automatically. [✓ Installed and verified — see addendum]
6. **Docs:** update `references/bonds-lite-setup.md` and this addendum. [✓ Updated — see addendum]

## Out of scope

- Residual tolerance thresholds (see Decision 129).

## Addendum: Implementation Progress

- **Steps 1–3 done.**
  - `src/sst_property_stage_four.py` has been added, with unit tests in `tests/test_sst_property_stage_four.py`. P4 dispatch, P4 installation and console-message tests were added to `tests/test_bonds_lite_input.py`.
  - The CLI, controller entry, source runner and packaged hidden imports have been wired up.
  - Packaged release `20261111-property-p4-v1` was built from the isolated Python 3.13/3.11 environments and promoted to the central pointer. The controller, Model and MACalc self-tests passed; unchanged Model/MACalc components were reused. The build used the current working tree, not a Git commit.
  - Shared Stage 2 helper messages now accept stage/workbook descriptions so Stage 3 and Stage 4 errors no longer incorrectly identify Flash or NoteVal as an LTM workbook or a Stage 2 workbook.
  - Full suite after the changes: 1,497 passed, 95 skipped, 2 failed. The failures are the same unrelated `bonds_lite_input` failures: the frozen-main test calls `main()` with an argument although its signature accepts none, and the rollback test does not find the recovery path in captured output.
- **Deviations:**
  - The generic console waiting message is emitted every 60 seconds, as for P2, rather than every 5 seconds; the watchdog heartbeat remains every five minutes.
- **Step 4:** The opt-in T18 Excel test passed on 11 November 2026 using a temporary Draft, Flash and NoteVal copy. It verified a successful Stage 4 audit and confirmed the original workbooks were unchanged. The run took about 55 seconds. The test removes the Excel instances it owns.
  - The earlier manual run on copies took 61 seconds, with `I_StartTime_1` at 22:08:24 and `I_EndTime_1` at 22:09:03. Largest absolute residuals were 2.7e-5 on `O. MAC MI`, 7.4e-4 on `C. Note Valuation ex Hedge`, and 3.2e-4 on `C. Note Valuation`.
- **Step 5:** Stage 4 controls were installed in the live Draft on 11 November 2026. Backup: `references/backups/Draft Input spreadsheet for SST_20261111_221839_451460.xlsm`. A read-only validation confirmed the Stage 4 panel at `AB48:AI58`, button `btnRunNoteVal` calling `RunNoteVal`, and the existing Stage 3 panel at `AB35:AI45`.
- **Step 6:** `references/bonds-lite-setup.md` now documents the P3/P4 handoff, installation, run and safeguards.
- **Excel cleanup:** COM `Quit` left the owned NoteVal Excel process running in this environment. Stage 4 now verifies the exact owned PID after quit and terminates only that PID if needed. The gated test confirmed no Excel process created by it remained.
- **Decisions 126–131** were approved by the user on 11 November 2026; see the approval record in `references/decisions.md`.
