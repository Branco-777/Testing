# Plan: Property Stage 2 Long-Run Safeguards

Status: Implementation, Decision 102 and release approved; build and promotion in progress; live acceptance outstanding | 9 November 2026 | Decisions 99-103

## Overview

Property Stage 2 waits for the LTM macro with no time limit (Decision 97). A
16-scenario run normally takes 1-2 hours. If it stalls, or Excel is waiting
behind an error dialog, the console gives no sign. This plan adds a watcher
thread that reports progress, warns about stalls and likely dialogs, and offers
an abort at a per-scenario time limit. The LTM macro and model are unchanged.

## Confirmed facts

- `_invoke_ltm` calls `workbook.macro("Module1.CallRunFrozenPython")("flipper")`
  synchronously; the main Python thread blocks until the macro returns.
- `CallRunFrozenPython` calls xlwings `RunFrozenPython`, which runs
  `cmd.exe /C pushd <UNC> & "LTMflashArunmanager_v3.2a(d9bb4fbe).exe" ...`
  via `WScript.Shell.Run(..., WindowStyle=0, WaitOnReturn=True)`. The runner
  console is hidden and stderr goes to `%APPDATA%\xlwings.log`.
- On a non-zero runner exit code, xlwings `ShowError` shows a message box, and
  the macro (and therefore Stage 2) waits until someone clicks it.
- Stage 2 owns its Excel instance (`xw.App(visible=True, add_book=False)`), so
  the Excel PID is known and separate from the Draft's Excel.
- The runner writes progress (`Run N - End of run!`) into the workbook through
  COM while the macro waits.
- The packaged controller includes `pywin32==308`, not `psutil`.

## Behaviour (Decision 99)

| Signal | Rule | Action |
|---|---|---|
| Heartbeat | Every 5 minutes | Log `LTM running | elapsed 0:45 | Scenario 7 of 16 complete` (progress omitted if unreadable) |
| Stall | Runner tree CPU time unchanged for 10 minutes | Warn once per stall episode; reset when activity resumes |
| Likely dialog | Runner exited, macro not returned after 2 minutes | Warn: check the LTM Excel window for an error dialog |
| Time limit | `max(30, 10 x prm_NoOfScens)` minutes (16 scenarios: 160) | Console prompt: `[W]ait 30 more minutes or [A]bort?`; no answer means keep waiting |
| Abort | User answers A | End runner tree, then owned Excel PID; status `Timed Out`; no save; restore Stage 1 file if changed |

## Steps

1. **Write failing tests first** (see Test Strategy). Depends on nothing.
   [✓ Implemented — see addendum]
2. **New module `src/sst_property_stage_two_watchdog.py`** (depends on 1).
   - `ltm_time_limit(scenarios: int) -> timedelta`.
   - `count_completed_scenarios(messages: list[str]) -> int` from
     `Run N - End of run!` markers.
   - `ProcessProbe`: a Toolhelp32 snapshot (`ctypes`) for `(pid, ppid, exe)`, a
     search for the runner executable under the Excel PID, and the total
     CPU time of the runner tree (`win32api.OpenProcess` +
     `win32process.GetProcessTimes`). It returns `None` if the runner cannot be
     found, so the watcher falls back to elapsed time only and logs this once.
   - `ProgressReader`: run each read-only `ProgressFirstCell` COM query on a
     dedicated daemon worker with `pythoncom.CoInitialize()`, using
     `xw.apps[excel_pid]` and the workbook name. The watcher waits at most 5
     seconds. If Excel rejects the query, report no progress for that update.
     If it does not return within the deadline, disable further progress
     queries for this run; keep the other safeguards active.
   - `LtmRunWatchdog` (a `threading.Thread`, daemon) with injected clock,
     probe, reader, prompt and logger for testing. It exposes `stop()`,
     `aborted: bool` and `join()`. The state machine follows the table above.
   - `terminate_owned_run(excel_pid)`: ends the runner tree (deepest first),
     then Excel. It never touches PIDs outside the owned Excel tree.
   - Pass: watchdog unit tests are green.
   [✓ Implemented — see addendum]
3. **Wire into `run_property_stage_two`** in `src/sst_property_stage_two.py`
   (depends on 2).
   - Start the watcher immediately before `_invoke_ltm` with `application.pid`,
     the workbook name and `scenarios`; stop and join it in a `finally`
     around the macro call.
   - If `watchdog.aborted`, raise `PropertyStageTwoTimeoutError` (a subclass of
     `PropertyStageTwoError`) regardless of the COM error the macro call
     raises, and record status `Timed Out` in the `except` block.
   - When aborted, skip `workbook.api.Close` / `application.quit()` (the
     process is gone). Keep the existing backup restore, temporary-directory
     clean-up and run-guard release unchanged.
   - Pass: run-harness tests are green; the existing Stage 2 tests still pass.
   [✓ Implemented — see addendum]
4. **Packaging** (parallel with 3): add `sst_property_stage_two_watchdog` to
   the controller `--hidden-import` list in `tools/build_packaged_stages.py`
   (around line 132). Pass: the controller archive includes the module and
   the packaged self-tests pass.
   [✓ Implemented — see addendum]
5. **Documentation** (depends on 3): update the Property Stage 2 section of
   `references/bonds-lite-setup.md`: heartbeat, warnings, time-limit prompt,
   what abort does, and the `Timed Out` status. Replace "There is no automated
   timeout".
   [✓ Implemented — see addendum]
6. **Live acceptance** (depends on 3; needs user approval): source-mode run
   with a reduced `prm_NoOfScens` (e.g. 2) on a disposable Stage 1 copy.
   - Confirm that the heartbeat shows the scenario count rising. This verifies
     the bounded COM progress read from its worker thread.
   - With a temporary test-only limit override (environment variable
     `SST_P2_TIME_LIMIT_MINUTES`, source mode only), confirm that the prompt
     appears, that `W` continues, and that `A` ends the runner and Excel and
     records `Timed Out` with the Stage 1 file hash unchanged.
7. **Release** (depends on 5; user approved on 9 November 2026):
   commit, build a new release ID without `--promote`, check the self-tests
   and hashes, then promote. Decision 103 authorises release before step 6;
   live acceptance remains outstanding.

## Test Strategy

Write the tests before the code in `tests/test_sst_property_stage_two_watchdog.py`,
using a fake clock, probe, reader and prompt with no Excel or real processes:

1. `test_ltm_time_limit_*`: 1 scenario gives 30 min, 3 give 30 min, 4 give
   40 min, 16 give 160 min; zero or negative raises.
2. `test_count_completed_scenarios_*`: counts distinct `Run N - End of run!`
   markers and ignores other lines.
3. `test_watchdog_heartbeat_every_five_minutes`: four log lines at
   5/10/15/20 min. The scenario text appears when the reader returns
   messages and is left out when it returns `None`.
4. `test_watchdog_warns_once_per_stall`: CPU unchanged for 10 minutes gives one
   warning; activity then another 10-minute stall gives a second warning.
5. `test_watchdog_warns_when_runner_exited_but_macro_pending`: the runner
   disappears and 2 minutes later a dialog warning is logged.
6. `test_watchdog_prompts_at_limit_and_waits_on_w`: re-prompts 30 min later and
   never aborts.
7. `test_watchdog_aborts_on_a`: calls `terminate_owned_run` once and sets
   `aborted`.
8. `test_watchdog_without_runner_uses_elapsed_time_only`: logs that the runner
   was not found once, and the heartbeat and limit still work.
9. `test_find_runner_under_excel_pid`: in a fake snapshot, finds the runner
   under Excel through `cmd.exe` and ignores an identically named runner under
   another Excel PID.
10. `test_terminate_owned_run_only_targets_owned_tree`.
11. `test_progress_reader_returns_none_on_rejected_call`.
12. `test_progress_reader_disables_a_com_call_that_exceeds_its_deadline`.
13. `test_watchdog_stop_cancels_console_prompt`: stopping the watchdog while
    the console is waiting for input releases the prompt thread.
14. `test_terminate_owned_run_keeps_excel_open_if_runner_cannot_be_terminated`:
    do not terminate Excel if a runner process could not be ended.

In `tests/test_sst_property_stage_two.py`, extend `_run_harness`:

15. `test_run_records_timed_out_and_restores_when_aborted`: the fake macro
    blocks until the fake watchdog aborts. Expect status `Timed Out`, no save,
    no `Close`/`quit` calls, the backup restored if the file changed, and the run
    guard released.
16. `test_run_stops_watchdog_on_success_and_failure`.

Test data is synthetic: progress strings copied from the LTM macro's
format. The live check uses a disposable copy of a real Stage 1 workbook.

## Affected Files

- `src/sst_property_stage_two_watchdog.py`: new; watcher, process probe,
  progress reader, terminate helper.
- `src/sst_property_stage_two.py`: start/stop the watcher around `_invoke_ltm`,
  `Timed Out` status, abort-aware clean-up.
- `tests/test_sst_property_stage_two_watchdog.py`: new unit tests.
- `tests/test_sst_property_stage_two.py`: harness tests for abort and
  watcher lifecycle.
- `tests/test_bonds_lite_input.py`: controller-only hidden-import contract.
- `tools/build_packaged_stages.py`: hidden import.
- `references/bonds-lite-setup.md`: user-facing behaviour.
- `references/decisions.md`: Decisions 98-102 (added).

## Risks and Unknowns

- **Progress read from a second thread is unverified.** If Excel rejects or
  queues it, the heartbeat still shows elapsed time and runner state; only the
  scenario count is lost. Confirm in step 6 before relying on it.
- **False stall warnings** while the runner waits on SQL Server with low CPU.
  These are warnings only. Increase the threshold if they are noisy.
- **Runner process name or tree changes** in a future LTM version; the
  watcher degrades to elapsed time only and says so.
- **Console prompt input** while the main thread is blocked: the controller
  console is interactive (it already waits for Enter), but confirm in step 6.
- **Force-ending Excel** may leave an `~$` owner file next to the Stage 1
  workbook. Stage 2 restores the workbook content; check for and remove a
  stale owner file in step 6 if one appears.
- Ctrl+C during the macro call is still delivered only after the call returns
  (existing behaviour, out of scope).

## Addendum: Implementation Progress

- Steps 1-5 completed on 9 November 2026. Watchdog unit tests, Stage 2
  run-harness tests, packaging hidden-import checks, user documentation and
  decisions are updated.
- Added a bounded, single-worker progress COM query after code inspection
  showed that an unbounded query could itself block the watchdog. Decision 102
  records this conservative refinement, approved by the user in Decision 103.
  The approved run limits and other behaviour are unchanged.
- Added tests that verify the temporary time-limit override is source-only
  and ignored by a frozen controller.
- Added safety coverage for cancelling the console prompt and leaving Excel
  open if the runner cannot be terminated. Final focused validation:
  69 passed, 1 skipped.
- Added a regression test for answering Wait after leaving the prompt open.
  The extra 30 minutes now starts when the user answers, not when the original
  prompt appeared.
- Step 6 (live acceptance) remains pending separate user approval. The user
  approved step 7 (build and promotion) on 9 November 2026 before live
  acceptance; this authorised sequence change is recorded in Decision 103.
- Key decisions: 99 (user-approved run limits and prompts), 100 (watcher
  technical approach), 101 (approval to implement; live acceptance and release
  remain separate gates), 102 (bounded progress query, approved), and 103
  (approval to build and promote before live acceptance).
