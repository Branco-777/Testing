# Plan: Stage 2 Bonds Lite Batch

Version 1.2 | 28 October 2026 | Stage 2 implemented and installed; one authorised real batch attempted and rolled back after failure

## Overview

Add Run Bonds Lite as an independent development-stage button that passes the committed Stage 1 run folder to the existing external batch script once. Retain Python orchestration, the console and movable named audit fields; validate both successful model results before recording Stage 2 success.

The external model remains unchanged. This plan adds no run-all-stages command, no model optimisation and no history sheet. Stage 1 remains independently runnable.

## Confirmed Decisions

- Correct batch invocation is python run_bonds_batch.py <Stage 1 run folder>, not one invocation per input file.
- Use a separate per-user model interpreter. The supplied environment folder contains python.exe: C:/Users/Joao Branco/AppData/Local/anaconda3/envs/py311/python.exe.
- Add the button and a similar named/movable Stage 2 panel below the current Stage 1 controls, subject to occupancy checks. Preserve moved/resized Stage 1 layout.
- Ask before replacing existing Bonds/MAP results. Back up previous results and restore them if a confirmed rerun fails; remove the backup only after verified success and successful audit persistence.
- On failure/interruption/timeout, remove outputs created by the failed attempt. Never delete Stage 1 inputs or earlier successful results.
- The Stage 2 panel links to the run folder only, not the summary CSV. Keep console-only operational reporting; no new text logs.
- Timeout is 60 minutes for the complete two-input batch.
- After implementation and isolated checks pass, the user authorises one real batch verification on the latest eligible audited Stage 1 run. Do not start it during planning or select a folder by directory timestamp alone.

## Verified Runner Contract

External model directory supplied by the user ends in 02 Sensitivities/2025/05 ORSA/Models/Bonds Lite Batch/bonds-hm-cst. Its run_bonds_batch.py and BondsModelHandler.py were read without executing the model.

### Batch Behaviour

- run_bonds_batch.py::discover_input_files scans only direct-child *.xlsx files, excluding Excel lock files beginning with ~$; it does not recurse into the model-output folders.
- run_batch reads output_folder_path and output_suffix from every input in an extra Excel open/close preflight, checks target collisions, then calls BondsModelHandler.run_full_process sequentially in sorted filename order.
- A caught per-input exception records a failed summary row and stops the batch, leaving remaining inputs unprocessed. The current CLI can still exit 0. Exit code alone is therefore not a success signal.
- The runner writes bonds_batch_summary_YYYYMMDD_HHMMSS.csv in the input run folder. Columns are input_file, status, seconds, output_folder_path, output_suffix and error. Success values are success; failure values are failed.
- The summary is a model output artifact, not a new text-log file. Keep a successful summary in the run folder but expose only the folder hyperlink in the panel. Remove the failed attempt's summary during rollback after reporting its error.

### Model Outputs

BondsModelHandler.read_inputs reads output_folder_path/output_suffix and the P_tbl_OutputFilenames mapping. Its save_results defines names using the configured mapping or defaults, then appends the suffix and .csv.

- Always-produced core result types: ISINs, Grouped and BaseScenario.
- Always-produced BaseScenario nozeros filename is derived from the configured BaseScenario base name, preserving its trailing-underscore convention.
- InvExp_Base is always written, but can represent an empty DataFrame if bool_inv_exp is false; do not require data rows in that case.
- GroupedNoZeros is conditional: it is written only if zero-sum grouped columns are removed.
- Failed_zspreads files are diagnostic/conditional, not proof the batch failed.
- run_full_process also copies the input to its output folder as py_bonds_input_file_<output_suffix>.xlsx.
- Derive expected required names from each saved input's mapping/suffix, not from example filenames in USER_GUIDE.txt. No actuarial numerical-parity certification is implied by file presence checks.

### Lifecycle Risks

- run_full_process creates its own Excel application but uses xw.Book for the input; its early workbook reading/copying has no protecting finally block. Failure before its explicit wb.close/app_excel.quit can leave an orphaned Excel process.
- Do not close or kill the user's Draft Excel instance. Timeout cleanup must be scoped to the model child and demonstrably owned descendants; unknown Excel ownership is a recovery blocker, not permission to kill all Excel.
- The model itself opens each workbook again with link updates enabled. Stage 2 passes the inputs as they are; do not change their curves, stresses, precision settings or controls merely to accelerate the run.
- The guide quotes several minutes per input. The batch's existing extra preflight opens are not removed by this orchestration feature; no seconds-level runtime promise is made.

## Steps

### Addendum: Implementation Progress

- Step 1: Added 180 Stage 2 contract cases and a controlled subprocess runner fixture; no real-model execution. Preserved the existing user-edited Stage 1 suite.
- Steps 2-3: Implemented separate model interpreter configuration and a bounded Python 3.11/import preflight; added read-only named workbook metadata, exact pair membership, safe mapped result names and input hashes. Added openpyxl as a runtime dependency.
- Steps 5-7: Implemented explicit overwrite confirmation, same-volume result/summary backup and rollback, one streamed subprocess with a 60-minute deadline and Windows job containment, plus strict summary/artifact/input-hash validation.
- Step 4: Implemented named/movable Stage 2 audit APIs, panel and button installer; Stage 1 launcher remains intact and its existing button geometry is preserved on reinstall.
- Step 8: Implemented guarded Stage 2 orchestration, terminal audit states and run-identity readiness reader; CLI defaults to Stage 1 and dispatches Stage 2 only with `--stage 2`.
- Stage 1/2 OneDrive-XLL follow-up: Stage 1 opens read-only sources from byte-verified local `%LOCALAPPDATA%/Temp` copies and rejects nonnumeric base-curve caches. Stage 2 runs byte-verified local copies of both inputs, freezes the cached values in the three XLL-backed `P. Base Curves` columns (`GBP SONIA`, `GBP GOV`, `GBP RPI`) using the caller's add-in-enabled Excel application, and retains the original embedded output paths. The original Stage 1 workbooks remain formula-based and unchanged. Hidden model Excel recalculated `_xll.JRC_SpotAnnual(...)` to `None`, causing `1 + nparray_swaprate` to fail; a direct hidden-Excel open confirmed this. Focused refresh tests passed (25), input metadata tests passed (29), curve-freezing/batch/orchestration tests passed (31), and the complete non-Excel suite passed (600 passed, 8 skipped, 71 Excel tests deselected).
- Isolated verification: Default suite 593 passed, 79 skipped before the final scenario-prerequisite check; the focused read-only metadata suite then passed 28 cases. Eight controlled subprocess tests and result/transaction checks passed. The two movable/resized panel tests, five invalid-name tests and installer test each passed individually according to their XML reports. Supplied Python 3.11 preflight passed (Python 3.11.13; NumPy 1.26.4, pandas 2.2.3, xlwings 0.36.14, SciPy 1.14.1). The actual saved Stage 1 audit resolves by moved names to Succeeded, and both inputs, their distinct output folders, scenario directory and GBP_RPI.csv passed read-only checks. Excel emits native RPC shutdown diagnostics during test teardown; the terminal became occupied by an opt-in Excel run and subsequent commands did not execute.
- Operational deployment: Saved `BONDS_LITE_MODEL_PYTHON` and `BONDS_LITE_MODEL_DIR` in the current user's environment. Installed the Stage 2 panel and button into the closed original Draft; backup is `references/backups/Draft Input spreadsheet for SST_20261028_165507_606908.xlsm`. Verified the installed names and unchanged successful Stage 1 handoff read-only.
- Model failure diagnosis: The user's batch error `unsupported operand type(s) for +: 'int' and 'NoneType'` is raised in external `BondsModelHandler.solve_spreads_for_output` when adding 1 to `GBP SONIA`. In the Stage 1 output, all 1,201 cached values in `P. Base Curves[GBP SONIA]` were `#NAME?`; the shared template's same values were numeric. The Stage 1 diagnostic had created a separate hidden Excel instance without the JRC add-in, then calculated every worksheet. The table's 19 blank rows 1007-1025 were outside `I_tbl_Monthly_Bonds` (`B8:AUN1006`) and are not the cause. No external model file or original input was changed.
- Successful rerun: Stage 1 succeeded on `Bonds Lite output/20261028_181415`. Stage 2 used byte-verified local inputs and froze numeric caches for the three XLL-backed curve columns in those temporary copies. The saved audit records Succeeded for the same run (18:53:29-18:56:12; 162.685 seconds). Summary `bonds_batch_summary_20261028_185611.csv` passed the production validator with exactly two success rows, correct output folders/suffixes, all required artifacts in Bonds and MAP, and unchanged original input hashes. No external model files or original Stage 1 inputs were changed.
- Remaining: None for the approved Stage 2 verification. The earlier failed run `20261028_174512` remains an invalid historical attempt; the current saved handoff/results are `20261028_181415`.
- Documentation updated in `references/bonds-lite-setup.md`; this progress record supersedes the older Stage 1-only description in the prior plan addendum.

### 1. Establish The Test Contract First

Affected file: tests/test_bonds_lite_input.py, reading its current user-edited contents before changing it.

Reuse existing fake clocks, stage audit boundaries, temporary paths and isolated Excel fixtures. Add initially failing tests for CLI stage dispatch, model configuration, exact one-child invocation, summary/result validation, transactional overwrite and named Stage 2 audit. Ordinary tests must not start the real model or access production scenarios.

Use a tiny controlled subprocess fixture to emit progress, write summaries and create temporary model outputs. Reuse an existing fixture location; create a dedicated fixture script only if needed to test real child lifecycle and silent timeout. Do not add a framework or duplicate large workbooks.

Pass criteria: failures demonstrate the missing Stage 2 behaviour; user test edits are preserved; existing Stage 1 source immutability and movable-panel coverage remain intact.

### 2. Configure A Separate Model Interpreter

Affected files: src/bonds_lite_input.py and references/bonds-lite-setup.md.

Add load_model_config using a per-user BONDS_LITE_MODEL_PYTHON setting. Require an actual python.exe; accept a supplied environment directory only by resolving its python.exe explicitly. Never silently fall back to the controller's Python 3.13 environment.

Resolve the model directory from a configurable BONDS_LITE_MODEL_DIR; default to the supplied relative project location ../../02 Sensitivities/2025/05 ORSA/Models/Bonds Lite Batch/bonds-hm-cst from the caller's local folder, not Joao's absolute user profile. Verify run_bonds_batch.py exists. Keep the controller's existing BONDS_LITE_PYTHON role unchanged.

Use bounded preflight import/version checks for the model environment and declared dependencies. Do not import model numerical dependencies into the controller package, install into another environment without authorisation, or run the model as a dependency probe.

Pass criteria: missing/invalid settings fail clearly; paths with spaces work; Sarah's local directory resolves without a hard-coded Joao path; no controller/model-interpreter confusion.

### 3. Validate The Exact Stage 1 Handoff And Saved Inputs

Affected file: src/bonds_lite_input.py, reusing resolve_stage_input/resolve_stage_inputs.

Hold the same workbook guard for Stage 2. Require the committed saved Succeeded Stage 1 handoff and exactly the two expected direct-child input workbooks; reject an extra processable .xlsx because the external runner would also execute it. Require the two inputs to be closed in Excel before launch so model preflight does not attach to a user's open input workbook.

Read saved input metadata without saving or changing inputs: entity, output_folder_path, output_suffix, bool_inv_exp and P_tbl_OutputFilenames. Resolve names rather than assumed cell addresses. Validate each configured output folder belongs to the selected run and matches its dataset, both folders are distinct, and expected output names cannot escape them through absolute paths or traversal. Validate required model scenario paths are accessible; do not change Read/Create flags or scenario parameters.

If Stage 1 was created under another user's absolute execution paths, report the mismatch and require a correct regenerated Stage 1 run. Do not silently rewrite the inputs when the user requested running them as they are. Hash input files before launch and verify unchanged bytes afterward.

Pass criteria: no selection of newest uncommitted folders, no silent third input, no incorrect output destination, no changed inputs. A failed/missing Stage 1 handoff stops Stage 2 before result modification.

### 4. Add The Named Stage 2 Panel And Button

Affected files: src/bonds_lite_input.py and Draft Input spreadsheet for SST.xlsm via backed-up installer.

Extract only small private field-map-driven panel helpers if needed; keep existing Stage 1 public APIs, names and legacy upgrade behaviour stable. Add an independent Stage 2 name map with panel, LastRun, Start, Finish, Elapsed, User, Status, InputRun and LastSuccessfulResults fields. Do not reuse SST_Stage2_Input for a new panel field: it is already Stage 1's active handoff. Use SST_Stage2_InputRun for the Stage 2 attempt snapshot.

Locate free cells below the current named Stage 1 panel and button at installation time. Proposed visual grouping mirrors Stage 1's coral labels/pale values, wrapped path fields, full timestamp storage and [h]:mm:ss duration. Preflight cells, merges, protection, shape collisions and panel overlap before mutation. Preserve the actual Stage 1 button position, including on reinstall; the current fixed-anchor recreation must not move user controls.

Add a Form Control named Run Bonds Lite, wired to a new VBA RunBondsLite launcher. Reuse caller FullName/HWND matching, configured controller interpreter, -u, visible asynchronous console and --pause. Add CLI --stage with Stage 1 as the backwards-compatible default and Stage 2 as an explicit dispatch.

Pass criteria: default Stage 1 button still works; Stage 2 panel is movable/resizable through names; idempotent install preserves both panels and entered controls; no occupied region is overwritten. Save/close Draft and create a backup before installation.

### 5. Confirm Reruns And Stage Existing Results Safely

Affected file: src/bonds_lite_input.py, proposed model_results_transaction helper.

Before a model launch, inspect both expected result directories and old batch summaries. If results already exist, ask for explicit replacement confirmation in the visible console, defaulting to No. Declining must launch no child and alter no results; report cancellation rather than success. Keep the workbook guard while the decision is active.

For an accepted attempt, create a private same-volume rollback directory outside the batch's direct input-file selection. Back up existing contents of both result folders and summaries before removing/moving anything permanently; retain empty original result directories so unchanged input paths remain valid. Track exact ownership and before-state. Do not move the Stage 1 inputs or unrelated run-root files.

On success, discard the rollback copy only after summary/result validation AND the successful Stage 2 audit save. On failure, stop/reap the model first, remove only new attempt outputs/summary and restore prior directories/summaries. If restoration fails, retain the backup and print its location; do not claim clean rollback or delete the only recovery copy.

Pass criteria: refusal is non-destructive; first-run partial outputs are removed; failed rerun restores earlier bytes in both folders; previous successful results survive all error paths. A failure during backup preparation restores what was moved and never launches the model.

### 6. Launch One Bounded Batch And Stream Progress

Affected file: src/bonds_lite_input.py, proposed run_model_batch.

Start the configured model interpreter with an argument list equivalent to python -u run_bonds_batch.py <run-folder>, shell=False and cwd set to the model directory. Use one invocation for the complete folder, not two single-file calls. Keep the child environment independent of the controller while ensuring the configured environment's dependencies/DLL paths are available.

Stream merged stdout/stderr into the existing console with flushing. A reader thread may handle output but must not access Excel/COM. Enforce the 3,600-second deadline independently of output reads; a silent child or partial line must not defeat the timeout. Reuse five-second elapsed/waiting reporting and do not invent percentage completion.

Handle nonzero exit, timeout and Ctrl+C. Stop and reap the owned child/descendants before rollback. Scope any Windows process-tree/job handling to this model attempt; never issue a global taskkill for EXCEL.EXE or terminate the caller's Excel instance. If model-owned Excel remains locked and cannot be identified safely, report recovery blocked and preserve rollback files.

Pass criteria: one child receives the exact folder and cwd; spaces/non-ASCII paths survive; output is visible before completion; silent timeout and cancellation are bounded; all reporter/process resources close before the Enter prompt.

### 7. Validate Summary And Fresh Results

Affected file: src/bonds_lite_input.py, proposed validate_model_results.

The backed-up/empty destinations establish freshness without trusting old file timestamps alone. Identify the newly created summary from the attempt's file-set change; reject absent, duplicate or stale-only candidates. Parse CSV with the standard library's structured reader, validate columns, and require exactly two distinct expected input filenames with status success, empty error values and matching configured targets/suffixes. A failed or missing MAP row is a Stage 2 failure even if the child exits 0.

For each destination require newly produced nonempty core ISINs, Grouped, BaseScenario, derived BaseScenario nozeros and the model's input-copy artifact using the verified filename mapping. Check the always-written investment-expense file exists, permitting its documented empty-DataFrame form when expenses are disabled. Treat GroupedNoZeros and Failed_zspreads as conditional outputs, not universal requirements. Do not certify actuarial values from file presence or silently adjust numerical tolerances.

Pass criteria: stale files, an input-copy alone, a summary alone, partial datasets and wrong-folder output cannot produce Succeeded. Both input hashes remain unchanged. Output validation failures enter rollback with a readable per-file reason in the console.

### 8. Persist Stage 2 Audit And Readiness

Affected file: src/bonds_lite_input.py, proposed run_stage_two.

Resolve the saved Stage 1 handoff before dirtying Draft. Validate configuration/targets and obtain any overwrite approval before destructive changes. Record the accepted attempt's Running state, Windows user, full start timestamp and captured relative InputRun; preserve its last-successful-results link. Automatically save Draft under the existing approved autosave policy. If the initial audit save fails, launch no child.

After batch and artifact validation, write finish/duration and Succeeded with the run-folder link; save Draft before discarding rollback data or returning success. If success-audit saving fails, roll back new model outputs, restore old results where applicable, clear Stage 2 readiness and report the persistence failure. Record Failed/Interrupted/Timed Out where writable, preserving the original failure if audit recording also fails. Report preflight or declined-rerun outcomes clearly without falsely recording a completed batch.

Keep the latest Stage 2 attempt's InputRun explicit. When Stage 1 later creates another run, an old Stage 2 success must not authorise downstream use of that new folder: any future readiness reader must compare Stage 2's captured run identity with the current committed Stage 1 handoff, not status alone. Do not silently rewrite old successful result links.

Panel duration excludes the user's overwrite decision, Python startup and final Enter wait; console total remains the broader elapsed figure. Guard release and process/file cleanup must finish before the console pauses.

Pass criteria: coherent saved timestamps/status/user/folder; output validation and final audit commit precede success; failed audit/timeout/rerun cannot authorise downstream stages. No run history or summary-specific panel link is created.

### 9. Verify, Install And Run The Approved Real Batch Once

Affected files: tests/test_bonds_lite_input.py, references/bonds-lite-setup.md, references/decisions.md and references/plan-bonds-lite-input.md.

Run the complete default suite, isolated subprocess tests, and explicitly opted-in Excel installer/panel/audit cases. Preserve user edits in the current test file. Install into the saved/closed backed-up Draft only after these gates pass. Document both interpreter settings, model directory, expected runtime, overwrite prompt, timeout, rollback, result checks and failure recovery.

Then perform the user's authorised one real batch using the latest eligible named Stage 1 handoff, configured py311 interpreter and original model script. Inspect/resolve model prerequisites first, but do not alter inputs or model code to bypass failures. If no eligible successful Stage 1 folder exists, stop and ask the user to run Stage 1; never infer or regenerate it silently. Obtain the ordinary overwrite confirmation if needed. Record actual complete runtime, two success rows, fresh files in both output folders, unchanged inputs and saved Stage 2 audit. Do not automatically retry real model failures.

Pass criteria: Run Bonds Lite button launches one real folder batch; both output folders contain verified results; the saved named audit represents the same run; the console remains readable until Enter. Report any native Excel lifecycle or model-data issue separately, without claiming readiness from test or process launch alone.

## Test Strategy

1. Unit tests first: config paths, separate interpreter selection, exact command/cwd, input-folder membership, named metadata ownership, safe filenames, summary schema/rows/statuses, required/conditional artifacts and run-identity readiness.
2. Isolated subprocess: a controlled tiny runner covers immediate output, silence, unterminated lines, nonzero exit, zero exit plus failed/missing MAP summary, timeout and interruption. No real model dependency or scenario network is involved.
3. Transaction fixtures: both old result sets restored byte-for-byte, first-run partial files removed, unrelated run-root files/Stage 1 inputs untouched, backup and restoration errors visible with backups retained.
4. Audit tests: accepted start, initial save failure, verified success, partial model failure, timeout/interruption, final audit-save failure/rollback, duplicate-run guard, declined overwrite and cleanup-before-pause. Reuse Stage 1 clock/guard fixtures where appropriate.
5. Opt-in Excel tests: named Stage 2 panel, geometry validation/move/resize/save-reopen, preserved moved Stage 1 controls and idempotent backed-up installation. Existing input/source immutability tests remain unchanged unless their own dispatch boundary legitimately changes.
6. Real model verification: one run only after isolated gates and valid audited inputs. This is operational integration verification, not actuarial model regression certification. Reuse the model's existing regression suite separately if its numerical implementation changes; no model algorithm change is planned here.

## Affected Files

- src/bonds_lite_input.py: stage dispatch, model configuration, guarded Stage 2 wrapper, subprocess lifecycle, metadata/result validation, transaction and named panel/button installation.
- tests/test_bonds_lite_input.py: current-file-preserving Stage 2 coverage and existing Stage 1 regression expectations.
- tests/fixtures/: one small controlled runner only if necessary for actual subprocess behaviour; no production-data duplication.
- Draft Input spreadsheet for SST.xlsm: backed-up new button/panel and names; preserve existing controls and Stage 1 metadata.
- references/bonds-lite-setup.md: per-user Stage 2 configuration and use/recovery instructions.
- references/decisions.md and references/plan-bonds-lite-input.md: confirmed choices and implementation progress/link.
- External run_bonds_batch.py, BondsModelHandler.py and pyBondsHMutils.py: inspected integration dependencies, not edit targets.

## Risks And Unknowns

- Attempting a saved-Draft handoff read in this planning turn raised PermissionError. Current eligible run status remains unverified; read/validate it during implementation before real execution.
- The current external runner opens Excel twice per input for its own collision and model passes. The controller cannot remove those costs without a separate approved model change.
- A Windows-session mutex does not provide a cross-user/coauthoring lock. Local rollback moves do not imply atomic cloud synchronisation.
- An early model read failure can orphan Excel. Never endanger user workbooks to satisfy cleanup; preserve recovery backups and report blocked restoration if ownership/locks cannot be resolved.
- Saved inputs may include network scenario dependencies and absolute output paths. Validate rather than silently rewrite them when running as-is.
- The model guide's file count/naming examples are not fully authoritative; actual save_results and input mappings control required results.
- The model environment's executable exists, but imports, input prerequisites and numerical run compatibility have not yet been executed or certified.
- No production code, workbook panel, button or model run has been modified/performed in this planning turn.

## Skills Consulted

Questioning-patterns, logging, testing-strategy, regression-testing, performance-optimisation, decision-logging, coding standards and document-extraction conventions. No team-specific conventions are active.