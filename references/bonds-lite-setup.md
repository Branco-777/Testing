# Bonds Lite Input Setup And Use

Version 0.8 | 28 October 2026

## Overview

Generate Bonds Lite Input on the Draft workbook's I. Control sheet starts the Python workflow through a small VBA launcher. Stage 1 prepares two validated inputs, Bonds and MAP, using one template and one captured set of date/entity/stresses. Both must succeed before the run folder becomes eligible for Stage 2.

Python orchestrates native Excel copying, destination-only formula refresh and console progress. The movable Stage 1 panel links to the completed timestamped run folder. Run Bonds Lite invokes the unchanged folder-based model batch from a separately configured Python 3.11 environment. The default suite passed 593 tests before the final scenario-prerequisite check; the focused metadata suite then passed 28 tests. Isolated Excel assertions passed individually, though native RPC shutdown diagnostics appeared during test cleanup.

## Prerequisites

- Windows desktop Excel; Excel in a browser is not supported.
- An Anaconda Python 3.13 environment. The verified environment is Python 3.13.5 with xlwings 0.32.1.
- The Bonds Lite model's separate Python 3.11 environment, with NumPy, pandas, xlwings, SciPy and python-dateutil available. Imports and version are checked before the audit or results change.
- The existing Bonds Lite Batch project, including run_bonds_batch.py; model files are not changed by this workflow.
- The JRC Excel add-in that provides `_xll.JRC_SpotAnnual` must be available in the Excel session used to run Stage 1. Use the installed Draft button from that session; a separate automation Excel instance may not have the add-in.
- The project folder, including pyproject.toml and src, available at a stable local or synchronised location.
- Access to the selected data/template files and all saved external workbooks required by the destination formulas.
- Permission to run the known Draft macro under your organisation's Excel policy. No xlwings Excel add-in is required. Do not enable all macros or weaken global security settings.

## Installation

Run these commands in PowerShell from the project folder, adjusting the Anaconda path if needed:

```powershell
$python = "$env:LOCALAPPDATA\anaconda3\python.exe"
& $python -m pip install -e ".[dev]"
[Environment]::SetEnvironmentVariable("BONDS_LITE_PYTHON", $python, "User")
[Environment]::SetEnvironmentVariable("BONDS_LITE_MODEL_PYTHON", "$env:LOCALAPPDATA\anaconda3\envs\py311\python.exe", "User")
```

After changing either model environment variable, fully close and restart Excel before clicking Run Bonds Lite. A running Excel process retains the environment it inherited when it started, so the launched controller may otherwise report the model interpreter or directory as missing.

If the model project is elsewhere, set the optional per-user `BONDS_LITE_MODEL_DIR` to the folder containing `run_bonds_batch.py`. Otherwise it resolves relative to Draft as `../../02 Sensitivities/2025/05 ORSA/Models/Bonds Lite Batch/bonds-hm-cst`. There is no fallback to the controller's Python 3.13 environment.

The editable installation points to this project folder. Keep the folder in place; repeat installation if it moves. Each colleague installs into their own environment and sets their own interpreter path.

The supplied Draft workbook already has the button and audit panel installed. If installation or reinstallation is required, close Draft in Excel and run:

```powershell
& $python -m bonds_lite_input --install-button "Draft Input spreadsheet for SST.xlsm"
```

Installation requires Excel to allow programmatic access to the VBA project under your organisation's policy. If access is blocked, ask your authorised administrator or workbook maintainer; the installer does not change security policy. Original Draft copies are retained under references/backups before installation changes.

The Stage 1 panel was initially installed at I. Control!H73:O83. Stage 2 adds its own named panel below the current Stage 1 panel and controls. Installation checks occupied cells, merges and shape overlap and preserves the existing Stage 1 button geometry on reinstall. Keep Draft saved and closed; a timestamped backup is made before installation.

## Configuration

| Setting | Type | Value And Effect |
| --- | --- | --- |
| BONDS_LITE_PYTHON | Per-user environment variable | Absolute python.exe path; required, no built-in default |
| BONDS_LITE_MODEL_PYTHON | Per-user environment variable | Separate model environment's Python 3.11 python.exe; required for Stage 2 |
| BONDS_LITE_MODEL_DIR | Per-user environment variable | Optional model project folder containing run_bonds_batch.py |
| Draft I. Control!D39 | Text | Data workbook path; installer fills the supplied data filename only when blank |
| Draft I. Control!D40 | Text | Required MAP data workbook path; deliberately left blank by the installer if blank |
| Draft I. Control!D42 | Text | Shared input template path for both datasets; installer supplies its default only when blank |
| Draft I. Control!D11 | Excel date | Valuation date copied to output I. Control!D10 |
| Draft I. Control!D9 | Text | Entity copied to output I. Control!D13 |
| Draft I. Credit Stresses | Numeric cells | All 12 blocks required; decimals, copied without scaling |
| Output location | Derived, not an input cell | Bonds Lite output under the saved Draft workbook's folder |

Paths may be absolute or filenames relative to Draft. Use the locally synchronised OneDrive path, not an HTTPS link. Draft, Bonds data, MAP data and the shared template must all be distinct files. D41 is no longer the template-path control after the inserted MAP row; its contents are not used by Stage 1.

Each data source must contain I. Bonds Data and its I_tbl_Monthly_Bonds table, beginning at B8 and ending in column ATT: 1,215 columns. Its headers must match the destination's B:ATT prefix in order. The input template retains its calculated tail through AUN. Bonds and MAP row counts may differ; missing, reordered or additional raw columns are rejected. The old full-width AUN data-source format is no longer accepted.

The supplied source was rescanned as B8:ATT1066, or 1,058 bond rows; the supplied input table was B8:AUN1006, or 998 seed rows. These row counts are not limits: the output table resizes to the source records.

## Use

1. Save Draft in its local or UNC folder and open it in desktop Excel.
2. Set D39 to Bonds data, D40 to MAP data and D42 to the shared template. Close those files and required external dependency workbooks in this Excel instance before running. Dependencies must contain the saved data you intend to use. A blank D40 blocks generation; no MAP filename is assumed.
3. Confirm date/entity and populate every required numeric stress cell. Genuine zero is valid; a blank is not zero.
4. On I. Credit Stresses, use C:J, M:DI and DL:HH for each row group 9:15, 20:26, 31:37 and 42:48. These are the four transition matrices and the two spread grids per currency/sector.
5. Click Generate Bonds Lite Input on I. Control near row 73. Keep Draft open while the asynchronous Python process runs. Do not edit the inputs during a run.
6. Follow the stages in the visible console. Wait for the saved-path or failure summary and total processing time. Audit persistence, cleanup and guard release finish before the console asks you to press Enter to close it. A second simultaneous run for the same Draft is rejected without changing the active run's audit. A failure message is not a successful output.
7. After Stage 1 succeeds and its saved handoff is current, click **Run Bonds Lite**. Stage 2 passes that run folder to the batch script once. Close both generated inputs in Excel first. If old results exist, answer `y` to the console overwrite question; blank or any other answer defaults to No. Stage 2 allows up to 60 minutes for the complete batch.

The workflow generates two workbooks using all credit blocks, regardless of scenario toggles. Date, entity and stresses are captured once, so both outputs use the same settings even if controls are subsequently edited. Stage 2 executes the Bonds Lite model only after a saved Stage 1 handoff is validated. Entity must be nonempty Windows-safe text with no surrounding whitespace because it is used unchanged in directory names.

## Stage 2 Audit And Results

Stage 2 has its own movable panel and does not change `SST_Stage2_Input`, the active Stage 1 handoff. Its Input Run records the captured relative run folder; Last Successful Results is set only after both batch rows, required files, unchanged input hashes and the saved success audit pass. A later Stage 1 run does not make an older Stage 2 result current; readiness compares both run identities.

The batch summary must contain exactly one successful row for each expected input, with matching output destinations and suffixes. Required mapped ISINs, Grouped, BaseScenario, BaseScenario nozeros, model input copy and InvExp_Base files must be present in both result folders. InvExp_Base may be empty when investment expenses are disabled. Conditional diagnostic files are not universal requirements.

On an approved rerun, previous result-folder contents and old batch summaries are backed up outside the selected run folder on the same volume. A failed, interrupted, timed-out, invalid or unauditable attempt removes new artifacts and restores the old files. If recovery is blocked, the console identifies the retained backup. Inputs and unrelated run-folder files are not moved. Do not retry a real model failure automatically; resolve its reported prerequisite first.

Earlier verification attempts failed because hidden Excel could not access the OneDrive inputs and recalculated JRC XLL curves to `None`. The current workflow stages byte-verified local copies and freezes only the validated GBP SONIA, GBP GOV and GBP RPI formula caches in those copies; the original Stage 1 workbooks and embedded output paths stay unchanged. The successful end-to-end rerun used `Bonds Lite output/20261028_181415`; its saved audit and two-row summary both report success, and required artifacts were validated in both output folders. The summary is `bonds_batch_summary_20261028_185611.csv` in the run folder.

## Stage 1 Audit And Stage 2 Input

The coral/pale-pink panel beside the button starts as Not Run and is populated by the next audited run:

| Field | Initial Value Cell | Meaning |
| --- | --- | --- |
| Last Run | L74 | Date of the latest accepted attempt |
| Start | L75 | Start time; the stored value includes its full date |
| Finish | L76 | Finish time after generation cleanup; full date retained |
| Time | L77 | Monotonic elapsed duration, displayed as [h]:mm:ss |
| User | L78 | Windows login running Python |
| Status | L79 | Not Run, Running, Succeeded, Failed or Interrupted |
| Last Successful Output | L80 | Clickable relative path to the last successfully committed run folder |
| Stage 2 Input | L82 | Active relative run-folder handoff, set only after both inputs and the audit save succeed |

The two paths use wrapped merged cells through rows 81 and 83. For example, Bonds Lite output/20261028_140506 is resolved relative to the current user's local Draft folder, not the user who created the workbook. The panel remains compact: there are no separate Bonds/MAP audit fields.

### Moving And Resizing The Panel

Runtime reads, writes and validation now follow Excel names, not the initial addresses listed above.

1. Select SST_Stage1_Panel in Excel's Name Box to select the complete panel, including both output-path rows.
2. Use Cut, select an empty destination on I. Control, then Paste. Do not copy/paste a duplicate and expect the code to choose the new copy. Keep clear of existing inputs, buttons and other shapes.
3. Adjust column widths and row heights as needed. Inserting rows or columns within the panel is supported when Excel updates the names and keeps the value fields intact.
4. Save Draft before running the next stage.

For a structural redesign, use Formulas > Name Manager to ensure SST_Stage1_Panel encloses the complete panel and each field name below identifies its correct value cell. A field may name one cell, a cell inside one merged value area, or the complete merged value area. The code writes to that merged area's top-left cell.

| Field | Workbook Name |
| --- | --- |
| Last Run | SST_Stage1_LastRun |
| Start | SST_Stage1_Start |
| Finish | SST_Stage1_Finish |
| Time | SST_Stage1_Elapsed |
| User | SST_Stage1_User |
| Status | SST_Stage1_Status |
| Last Successful Output | SST_Stage1_LastSuccessfulOutput |
| Stage 2 Input | SST_Stage2_Input |

Keep the panel and all fields on I. Control. Do not delete or rename these names. Deleted/#REF references, references to another workbook/sheet, multi-area references, overlapping fields and fields outside the named panel are rejected before audit writes or saves. The code never falls back to the old L74/L82 cells when names are broken. Do not move or resize the panel while a run is active.

The existing four-name panel was upgraded without resetting its values, hyperlinks, colours or position. A legacy panel that has already been structurally redesigned cannot always be inferred safely; installation will refuse it rather than guess which cells contain the audit fields.

**Draft is saved automatically at run start and at its terminal audit update. This also saves all other pending edits in Draft.** The source data and original Bonds Lite input template are never saved by this process.

Before generating, Stage 1 records Running and clears Stage 2 Input. If that first save fails, generation does not begin. On success, the complete run folder is committed to both path fields. If either dataset fails or is interrupted, only the new incomplete attempt folder is removed; Stage 2 Input stays blank and previous runs/last-successful links are retained. If Windows refuses cleanup, the console identifies the leftover incomplete folder and preserves the original error; it is not an authorised handoff. This is a latest-run panel, not append-only history.

If both valid input files were created but the final audit save fails, the complete run folder is retained. The console reports its folder and audit failure rather than announcing success. The active handoff is cleared in memory and recovery is saved where possible. This differs from a dataset-generation failure, which removes the incomplete attempt. If recovery cannot be saved, do not assume the workbook on disk contains the latest failure record. Resolve save access and rerun Stage 1.

Panel duration includes the initial audit-save and generation work but excludes Python startup, final audit-save time and the Enter wait. Console total processing time is the broader measure. Midnight crossings and durations greater than 24 hours are supported.

Workbook names SST_Stage1_Status, SST_Stage1_LastSuccessfulOutput and SST_Stage2_Input identify the handoff contract. Stage 2 requires a saved caller, successful status, matching folder links and both fixed nonempty input files inside that direct-child timestamped run folder. Missing members, legacy flat-file links, temporary folders, absolute handoffs and traversal are rejected. There is no fallback to the newest folder or to Bonds alone. `resolve_stage_input` returns the run folder; `resolve_stage_inputs` returns labelled Bonds and MAP file paths.

Stage 2 is installed as an independent **Run Bonds Lite** button. It invokes `python run_bonds_batch.py <local staged input folder>` once from the configured model directory using the separate Python 3.11 environment. The two workbooks are copied byte-for-byte to local temporary storage for Excel COM access; only the three JRC-backed curve formula columns are frozen to their already validated cached numeric values in those temporary copies. Their embedded JRL output paths are unchanged, so result files still go to the corresponding folders inside the audited run. The controller checks saved output paths and scenario files as-is; if those absolute paths do not match the current run/user, regenerate Stage 1 rather than silently rewriting inputs. The temporary batch summary is validated and moved back to the audited run folder only for the existing validation/audit transaction. Original Stage 1 workbooks remain unchanged.

## Console Progress And Timing

The button starts a visible console asynchronously, with unbuffered Python output. It processes Bonds then MAP, opening only the current data file and a unique private template copy for each. It retains formula recalculation and required-link refresh, but does not open the original template for duplicate checks or reopen saved outputs in Excel. No log files are created. Console stages identify which dataset is being processed.

- Start identifies the operation about to run, including each workbook open/close, data import, dependency refresh, worksheet calculation, save and verification.
- Done shows the completed stage's duration and total elapsed processing time.
- Waiting appears every five seconds during longer operations and shows the active stage and elapsed stage/total time. It means Python is waiting; it is not proof Excel is responsive or making progress.
- Failed identifies the stage that raised an error. The final message explains the failure; it is not a successful export.
- Total processing time excludes Python startup and the time spent waiting for Enter. It therefore is not an exact measurement from the physical button click.

Timing values in this example are illustrative, not a measured benchmark:

```text
Start: Open workbook: working.xlsx | total 2.0s
Waiting: Open workbook: working.xlsx | stage 5.0s | total 7.0s
Done: Open workbook: working.xlsx | stage 6.0s | total 8.0s
```

The console stays open after success or failure until Enter. Closing its window forcibly is not guaranteed to be graceful cancellation. Ordinary command-line calls do not pause unless --pause is passed; the workbook launcher adds that flag automatically. All Excel/COM work stays on the main Python thread; the waiting reporter uses timing information only.

## Native Copy And Refresh

Excel copies contiguous raw-input blocks through ATT with values-only PasteSpecial. Python does not receive the large raw-data arrays. Excel-side scalar checks verify exact values and distinguish text, numbers and blanks; real Excel errors are rejected. Destination formula columns are excluded from every source copy. In the supplied templates this means copying B:I and L:ATT, retaining destination formulas in J:K and ATU:AUN. Table row counts remain dynamic, and only newly added rows need formula filling.

The data file still contains formulas in J:K; they are neither imported nor explicitly refreshed. The source is opened read-only, with macros/link updates disabled and manual calculation in effect, and is closed without saving. All formula filling, required-link updates and explicit worksheet calculation target only the generated working copy, never the source data workbook or original input template.

The revised input template's escalation lookups reference its own P. Bond Escalation file sheet. The scan found values, not formulas, on that sheet and no required external formula dependencies for the current bond calculation chain. Literal network-path metadata on that sheet is not treated as a formula dependency. Future templates with genuine external formula dependencies still use the existing destination-only refresh checks.

This operation uses Excel's clipboard. Do not copy/paste elsewhere while the transfer is running. Copy mode is cleared on success or failure, but previous clipboard contents are not restored.

The workflow explicitly updates required links and calculates the worksheets. It waits only while Excel is actively calculating; pending work in manual calculation mode is not treated as an active calculation to wait for indefinitely. Stage 1 also verifies that every saved GBP SONIA value in P. Base Curves is numeric after calculation. If `_xll.JRC_SpotAnnual` is unresolved, Stage 1 fails before publishing the pair rather than handing invalid curve caches to Stage 2.

If Excel still reports link status 3 (old) after successful refresh and calculation, the console prints a warning and formula-error validation continues, as approved by the user. Other bad statuses, missing files, update failures and calculated formula errors stop publication. This does not independently prove that source-workbook formulas were recalculated recently: keep those sources current and saved upstream.

Before each save, the generated workbook's output_folder_path defined name is assigned the final absolute path of its matching model-output directory. After saving, the workflow checks Saved, table/formula/stress/control state and that named path, then closes and validates the file before moving it to its fixed final filename. The pair is authorised only when both inputs pass. Persisted-output comparisons remain integration-test responsibilities.

## Outputs

Each attempt reserves a new YYYYMMDD_HHMMSS folder under Bonds Lite output beside Draft. Collisions receive numeric suffixes such as _001; preceding runs are never overwritten. For entity JRL:

```text
Bonds Lite output/
	20261028_140506/
		Bonds Lite input_Bonds.xlsx
		Bonds Lite input_MAP.xlsx
		JRL Bonds/
		JRL MAP/
```

The two input filenames are fixed within each run. Both workbooks are closed after saving and verification. The two entity-labelled directories are empty model-output destinations; Stage 1 does not run the model or write its results.

In the Bonds input, output_folder_path points to the final current-user absolute path of JRL Bonds; the MAP input points to JRL MAP. The template's name currently resolves to I. Control!D19, but the code uses the name rather than that address. Both folders are created before either generated workbook is saved, so embedded paths never point at a temporary staging location.

Audit links are relative and portable. The paths embedded inside generated workbooks are intentionally absolute execution paths for the user who created the run. Stage 2 validates that each saved path points to the correct dataset folder in this run. If another user reuses the inputs, moving the audit link does not rewrite those paths; regenerate Stage 1 under that user's local folder.

The destination tab is I. Sensitivity, singular. Raw inputs replace the resized table body, formula columns are retained/recalculated, and shared date/entity/stresses are applied. Other template controls remain unchanged except output_folder_path. Original datasets and template are not saved. Private working files are removed after each input; prior completed runs survive failures.

## Example

Set D39 to Bonds Lite data template.xlsx, D40 to Bonds Lite data template_MAP.xlsx and D42 to Bonds Lite input template.xlsx. Enter the intended entity/date and a complete approved stress scenario. D40 is deliberately not filled automatically. Missing data paths, unsafe entity text or blank required stresses report validation errors before creating a run folder.

After validation, a run at 14:05:06 on 28 October 2026 would create the paired tree shown above and audit Bonds Lite output/20261028_140506. Do not reuse example stresses merely to bypass validation unless those are the intended scenario.

## Troubleshooting

| Message Or Symptom | Action |
| --- | --- |
| Set BONDS_LITE_PYTHON to your Anaconda python.exe before running | Configure the per-user variable and ensure that executable exists |
| No module named bonds_lite_input | Install this project into the exact configured interpreter |
| Bonds/MAP/template path required | Fill D39, D40 and D42. D41 is unused; D40 is not defaulted by the installer |
| Stage output must be a timestamped run folder | Rerun Stage 1 to create the paired output; legacy single-workbook audit links are not accepted |
| Stage run folder has no complete Bonds/MAP input | Do not run Stage 2 on a partial pair. Recreate both inputs with Stage 1 |
| Incomplete run folder could not be removed | Close any file opened from that incomplete attempt and remove it after resolving access. Stage 1 remains failed; prior runs are not deleted |
| Invalid credit stresses | Complete every reported cell with a real finite number; text numbers and booleans are rejected |
| Bond column headers do not match the template schema | Use the correct data table layout; do not silently drop or reorder columns |
| Close the input workbook before running | Close the specified source in the calling Excel instance |
| Close the required dependency before running | Save intended dependency changes yourself, then close that source before retrying |
| Required formula dependency is unavailable | Check the named saved file and your network permissions |
| Required formula dependency could not be refreshed | Check the named file and referenced sheet/ranges in Excel; no stale-value fallback is used |
| Warning: Excel reports a dependency as old (status 3) | The update succeeded but Excel retains its old-status flag. Keep source calculations current upstream; formula-error validation still runs before publication |
| Native copy verification failed | Avoid changing the clipboard during copying and check source/destination values and cell types; publication is stopped |
| Excel formula errors after recalculation | Repair the listed calculated columns or their source dependencies before retrying |
| P. Base Curves[GBP SONIA] is nonnumeric after calculation | Open Draft in the normal Excel session with the JRC add-in loaded, then run Stage 1 using the installed button; do not generate inputs from a separate hidden Excel instance |
| Stage 2 reports nonnumeric GBP SONIA cached values | Do not retry Stage 2 on that handoff. Run Stage 1 from the installed button in Draft's add-in-enabled Excel session, then verify the new audit and input pair |
| Use a local synchronised path, not a workbook URL | Open Draft from the locally synchronised OneDrive folder |
| A Bonds Lite run is already in progress | Let the first run finish; do not keep clicking |
| Install the Stage 1 audit panel with Draft closed before running | Close Draft and run the installer; the Python button entry point now requires the installed panel and names |
| Draft is read-only / I. Control is protected | Resolve workbook access/protection through the authorised owner; no stage is started without a writable audit |
| Generated output retained ... audit/handoff could not be saved | The generated file is retained, but Stage 2 is not authorised to use it. Fix Draft save access and rerun Stage 1 |
| Save Draft's successful Stage 1 audit before running Stage 2 | Save the workbook and confirm its successful path fields; unsaved audit values are not an eligible handoff |
| No progress console appears | Check the configured interpreter/package and reinstall the updated workbook launcher with Draft closed; changing Python's VBA string alone does not update the embedded macro |
| Repeated Waiting messages for one stage | Read the stage name and elapsed duration. A slow or blocked Excel call is possible; a waiting message is not a completion percentage or a watchdog |
| Installer emits an RPC/OLE warning on shutdown | Check the saved button/module and retained backup; an installer regression saved them successfully despite a native shutdown warning |

## Verification

Ordinary tests must not start Excel or refresh network data:

```powershell
& $python -m pytest -q -m "not excel and not performance"
```

Temporary-copy Excel tests use a synthetic local dependency by default:

```powershell
$env:BONDS_LITE_RUN_EXCEL_TESTS = "1"
& $python -m pytest -q tests/test_bonds_lite_input.py::test_excel_full_workflow_publishes_validated_copy --tb=short
& $python -m pytest -q -m excel --tb=short
Remove-Item Env:BONDS_LITE_RUN_EXCEL_TESTS
```

Real-network verification additionally requires BONDS_LITE_RUN_NETWORK_TESTS=1. Enable it only intentionally, with permission to refresh the production dependency. Do not confuse an opt-in skip, an unresponsive terminal or a process launch with a passing test.

Current paired-run verification: 421 default tests passed, eight environment/layout cases skipped and 63 opt-in Excel/performance cases deselected in 4.13 seconds. Ten isolated named-output-folder cases passed. The user confirmed the corrected full paired export and paired audited-run/failed-rerun checks passed, using distinct dataset records and row counts. Reopened inputs were checked for their raw data, shared controls/stresses, retained formulas and matching final output_folder_path. Source bytes and earlier runs remained unchanged. No production model execution or whole-run timing guarantee is implied.

Historical raw-only verification: 243 Excel-independent tests passed in 0.94 seconds, including 27 prefix-schema cases; 24 opt-in cases were deselected. The read-only scan of both supplied templates passed. The isolated full export passed in 11.82 seconds including fixture setup and teardown, with source-link-update/source-calculation guards, unchanged source snapshots and calculated destination lookup checks. Native Windows RPC diagnostics appeared during test cleanup after the passing assertion result; the cleanup warning remains a residual risk.

Stage-audit verification: 287 default tests passed in 1.14 seconds, one read-only layout test skipped and 26 opt-in Excel cases deselected. The blank-workbook panel test, temporary Draft installer, actual installed-panel read-only check, and user-run full audited export all passed. The full audited export test saved/reopened the audit, verified the portable handoff, then checked that a failed rerun blocks Stage 2 and retains the successful output without changing data/template sources. The live test took 11.33 seconds including fixtures; native RPC cleanup diagnostics remain unresolved.

Movable-panel verification: 295 default tests passed in 1.58 seconds, with one read-only layout case skipped and 48 opt-in cases deselected. The live movement/resizing and invalid-layout cases passed; all six legacy migration cases passed after clarifying the invalid-name error. The real workbook's saved named-field upgrade passed its read-only check. Test coverage includes writes leaving the old panel area untouched, row/column insertion, merged-field resizing and save/reopen handoff. Native Excel RPC teardown diagnostics remain a separate unresolved issue.

The run guard is local to the Windows session, not a cross-user/coauthoring lock. Do not have colleagues run stages simultaneously against coauthored copies of the same workbook and assume the Last Run panel is a distributed audit system.

The earlier native-copy benchmark and row-count results are retained in the plan as historical evidence. Revised raw-only row-count/performance fixtures cover the ATT source and wider destination, but were not all rerun this session. A complete intended-data runtime and clean multi-user Excel lifecycle remain separate gates; no whole-run speedup is claimed. This Python-only change does not require reinstalling the button.

## Glossary

- Caller: the open Draft workbook whose button was clicked.
- Dependency: a saved external workbook referenced by a retained formula column.
- Table body: bond records excluding the header row; unlike UsedRange, it determines the actual import size.
- xlwings: the Python library controlling the installed desktop Excel application.