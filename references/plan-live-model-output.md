# Plan: Live Bonds Lite Batch Output In Stage 2

Version 0.2 | 4 November 2026 | Approved by user. Decisions 78-80 confirmed by Decision 81.

## Overview

During Stage 2, the console shows only `Waiting: Run Bonds Lite batch` heartbeats. The model's own progress messages (`Opening workbook...`, `Solving idiosyncratic spreads...`, etc.) appear only after the batch ends. The controller already streams the child's stdout line by line (`run_model_batch` in `src/bonds_lite_input.py`), so the delay comes from the **packaged model buffering its own stdout** when writing to a pipe. Fix this in the model entry point without changing the Bonds Lite model, and show heartbeats only when the model has been silent for a while.

## Verified Facts (4 November 2026)

- The 16:39 run printed every model line after `stage 384.4s`, just before the batch ended at `388.3s`. This is block buffering of a piped stdout.
- `_model_environment` sets `PYTHONUNBUFFERED=1` for the packaged model. PyInstaller's documentation ("Specifying Python Interpreter Options") states that frozen apps ignore Python environment variables, so this setting has no effect.
- The model's `print()` calls in `BondsModelHandler.py` and `run_bonds_batch.py` do not pass `flush=True`. The model must stay unchanged; its sources are hash-pinned in `release.json`.
- `src/bonds_lite_model_entry.py` is the frozen entry point of `Model/BondsLiteModel.exe`. It is ours, so it can be changed.
- `_report_waiting` (heartbeat thread) and `stream_output` (model reader thread) run on separate threads. `threading.Thread` does not inherit `ContextVar` values on Python 3.13, so `_progress_state.get()` returns `None` inside a new thread. The state must be captured in the calling thread and passed in.
- The existing test `test_run_model_batch_streams_progress_before_child_can_finish` already covers the controller relaying lines before the child exits.

## Steps

1. **Line-buffer the packaged model's console** (`src/bonds_lite_model_entry.py`).
   - Add `configure_console_streams()`, which calls `reconfigure(line_buffering=True)` on `sys.stdout` and `sys.stderr` when they support it. It does nothing otherwise, for example when PyInstaller has set a stream to `None`.
   - Call it first in `main()`, before argument parsing, so it covers `--preflight`, `--self-test` and the batch.
   - **Verify:** the Step-1 tests pass; `--preflight` output is unchanged.

2. **Record model activity for the heartbeat** (`src/bonds_lite_input.py`). Can be done in parallel with Step 1.
   - Add `last_activity: float | None = None` to `_ProgressState`.
   - In `run_model_batch`, capture `state = _progress_state.get()` **before** starting the reader thread. In `stream_output`, after each printed line, set `state.last_activity = time.perf_counter()` under `state.lock`, if state is not `None`.
   - **Verify:** Step-2 tests pass; the existing `run_model_batch` tests still pass unchanged.

3. **Show heartbeats only when the model is quiet** (`_report_waiting`). Depends on Step 2.
   - Add the module constant `MODEL_QUIET_HEARTBEAT_SECONDS = 30.0`.
   - If `last_activity` is set, which happens only during a model batch, skip the `Waiting:` line until `MODEL_QUIET_HEARTBEAT_SECONDS` have passed since the later of the last model line and the last `Waiting:` line.
   - If `last_activity` is `None`, keep the current 5-second behaviour. This covers Stage 1 and other stages unchanged.
   - Reset `last_activity` to `None` when the batch finishes (in `run_model_batch`'s `finally`), so later stages return to 5-second heartbeats.
   - **Verify:** Step-3 tests pass; `test_report_waiting_uses_current_stage_clock_and_stops_without_sleeping` still passes.

4. **Apply the Step-1 fix to the MACalc entry point** (`src/macalc_batch_entry.py`). Can be done in parallel with Steps 1-3.
   - The packaged `MACalc/MACalcModel.exe` has the same issue for MACalc's internal `print()` calls. Our own messages there already use `flush=True`.
   - Call the same `reconfigure(line_buffering=True)` first in `main()`. Use a small local helper to avoid a cross-entry import that PyInstaller would need to bundle.
   - **Verify:** the MACalc variant of the Step-1 subprocess test passes.

5. **Run the targeted test suites.** Depends on Steps 1-4.
   - Run `pytest tests/test_bonds_lite_input.py tests/test_sst_stage_four.py -q` under the controller Python 3.13.
   - **Result:** focused live-output tests passed. The full targeted suites had unrelated baseline failures (45 failed, 701 passed, 81 skipped); the Stage 1 cleanup module run later had 3 unrelated baseline failures (671 passed, 81 skipped).

6. **Source-mode live check.** Depends on Step 5.
   - In source mode, Stage 2 runs `python -u run_bonds_batch.py`, so output was already live there. This step only confirms the heartbeat change: run Stage 2 on the existing `SST Output_Test\Credit` folder, with an output suffix set as the user intends.
   - **Status:** pending a live run from the Draft in Excel.

7. **Build, verify and promote the packaged release.** Originally gated on Step 6 and separate approval; the user approved build and promotion on 5 November while the live Stage 2 checks remain pending.
   - Build with `tools/build_packaged_stages.py --release-id 20261104-cache-v11` (no `--promote` at first). Run the model's `--self-test` and `--preflight` from the built folder.
   - Run one packaged Stage 2 from the Draft button using v11 (point `current.txt` to v11 only after approval, then restart Excel).
   - **Build and promotion:** v11 was built from a stale working tree that had lost the SST Output layout and Stages 3/4; the user saw Stage 2 write to `Bonds Lite output`. `current.txt` was rolled back to v10, the source was restored from HEAD with only the live-output and read-only cleanup changes reapplied, and `20261105-cache-v12` was built, checked to contain `allocate_stress_folder` and the heartbeat, validated on the central share and promoted. Never promote v11.
   - **Acceptance:** packaged Draft-button Stage 2 on v12 remains pending. Restart Excel before testing; output must land in `SST Output_<suffix>\Credit` and model lines should appear live. Rollback means writing `20261104-cache-v10` back to `current.txt`.

8. **Update documentation.** Depends on Step 7.
   - In `references/bonds-lite-setup.md`, update the deployment status to v11 and add a one-line note that Stage 2 shows live model progress, with heartbeats only after 30 seconds of silence.
   - **Status:** setup guidance updated. Keep this plan open until the packaged Stage 2 acceptance run passes.

## Test Strategy (Write First)

All tests use synthetic, in-process or `python -c` children; no Excel and no real model run.

| Step | Test (in `tests/test_bonds_lite_input.py` unless stated) | Expectation |
|---|---|---|
| 1 | `test_model_entry_line_buffers_piped_stdout_before_exit` | Start a child Python with `PYTHONUNBUFFERED` removed from its environment. It calls `configure_console_streams()`, prints a line **without** `flush=True`, then blocks on stdin. The parent reads the line within 5 seconds, **before** releasing the child. Control: without the call, the read times out. This proves the cause and the fix. |
| 1 | `test_model_entry_console_configuration_tolerates_missing_streams` | With `sys.stdout`/`sys.stderr` monkeypatched to `None` or to an object without `reconfigure`, the function does not raise. |
| 1 | `test_model_entry_main_configures_streams_before_any_mode` | Monkeypatch `configure_console_streams` and `preflight`; `main(["--preflight"])` calls the configuration first. |
| 2 | `test_run_model_batch_records_model_activity_from_reader_thread` | Inside `run_progress()`, a fake child emits two lines. `last_activity` becomes non-`None` during streaming and is reset to `None` after `run_model_batch` returns. |
| 3 | `test_report_waiting_suppressed_while_model_recently_active` | Use a fake clock (`time.perf_counter` monkeypatched) and a stop event that returns `False` a fixed number of times, as in the existing heartbeat test. Activity within 30 seconds produces no `Waiting:` line; after 30 seconds of silence, one line appears; the next appears only 30 seconds later. |
| 3 | `test_report_waiting_unchanged_without_model_activity` | With `last_activity=None`, heartbeats still appear every interval. |
| 4 | `test_macalc_entry_line_buffers_piped_stdout_before_exit` (in `tests/test_sst_stage_four.py`) | Same pattern as the first Step-1 test, for `macalc_batch_entry`. |

## Affected Files

- `src/bonds_lite_model_entry.py`: add `configure_console_streams()`; call it first in `main()`.
- `src/bonds_lite_input.py`: `_ProgressState.last_activity`, `MODEL_QUIET_HEARTBEAT_SECONDS`, reader-thread activity updates in `run_model_batch`, and the quiet-period rule in `_report_waiting`.
- `src/macalc_batch_entry.py` (Step 4): line-buffered console at the start of `main()`.
- `tests/test_bonds_lite_input.py`, `tests/test_sst_stage_four.py`: new tests above.
- `references/bonds-lite-setup.md`: deployment status and progress note.
- Not changed: Bonds Lite model sources (`run_bonds_batch.py`, `BondsModelHandler.py`, `pyBondsHMutils.py`), MACalc sources, `tools/build_packaged_stages.py`.

## Decisions

- Decision 78: fix buffering in our frozen entry point, not via PyInstaller `OPTIONS` or the model sources (confirmed by the user in Decision 81).
- Decision 79: during a model batch, heartbeat only after 30 seconds of model silence (confirmed by the user in Decision 81).
- Decision 80: include the MACalc entry point in the same change (confirmed by the user in Decision 81).

## Risks And Unknowns

- **Output from native code.** numba or other compiled code that writes directly to the C-level stdout is not affected by `reconfigure`. None was seen in the 16:39 log; every model line there is a Python `print`.
- **Console volume.** Live output adds no lines; it only moves them earlier. It may interleave with `Waiting:` lines, which is acceptable.
- **Blank output suffix (separate issue).** A blank `output_suffix` makes the model write `py_bonds_input_file_None.xlsx`, while the controller expects `py_bonds_input_file_.xlsx`. As a result, Stage 2 fails only after the full 6-minute batch. The user will set a suffix. A follow-up could reject a blank suffix in `read_model_inputs` before the batch starts. That is out of scope here unless requested.
- **Release timing.** Step 7 needs a new release; Stage 4 v10 acceptance is still pending. Building v11 from the same commit plus this change bundles both, so acceptance testing for v11 should cover both Stage 2 and Stage 4.
