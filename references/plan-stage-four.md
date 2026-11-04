# Plan: Stage 4 MACalc Batch On The Stage 3 Output

Version 0.6 | 4 November 2026 | Step 13 packaging implementation verified; commit, build, promotion and packaged acceptance remain gated.

## Overview

Stage 4 runs the existing MACalc Batch model (`MACalcrunmanager.py` in `00 Ad Hoc\MACalc_Batch`) on the MACalc input that Stage 3 last published successfully. Like Stages 1-3, it is started from a **Run MACalc** button on the Draft's `I. Control` sheet, through `bonds_lite_input --stage 4`, and records an `SST_Stage4_*` audit panel (Decision 76). It reads the live Stage 3 audit from the open Draft, refuses stale or ambiguous inputs, and runs MACalc Batch in place in the stress folder (Decision 61) under the model's own Python 3.11 environment. It answers MACalc's confirmation prompt only when MACalc's own list is exactly the audited workbook (Decision 62). MACalc Batch is not changed. The Draft changes only when the Stage 4 panel and button are installed (backed up first), and from then on in its `SST_Stage4_*` cells. Success means a run without errors that produces the expected artefacts (Decision 64). A packaged release containing Stage 4 is built and promoted only after the source-mode live run passes, with separate approval.

## Confirmed Requirements (User, 4 November 2026)

| Ref | Requirement | Decision |
|---|---|---|
| Q4-1 | Source-mode launcher first; button, audit cells and packaging later (partly superseded by Decision 76) | 60 |
| Q4-2 | Run in place on the audited workbook in `SST Output_<suffix>\<stress folder>` | 61 |
| Confirm | Answer `y` only if MACalc lists exactly one workbook, the audited one | 62 |
| Q4-3 | `is_downgrade_run` = `SST_Credit_DG_standalone`; no binary optimisation, no optimisation pass, CoD floor at zero | 63 |
| Q4-4 | Correct = runs without errors and writes the expected artefacts; the user then reviews the figures | 64 |
| Q4-5 | Stale handoff (latest Stage 3 not `Succeeded`, or not for the current run) stops Stage 4 | 65 |
| Q4-6 | Everything stays in the stress folder; a Stage 4 rerun replaces earlier Stage 4 results | 66 |
| Q4-7 | No `_opt` suffix guard for now; "Stage 5" in older documents means Stage 4 | 67 |
| Button | Stage 4 button, audit panel and `--stage 4` in the same fashion as Stages 1-3 | 76-77 |

## Verified Discovery (Read-Only, 4 November 2026)

### The model

- Folder: `C:\Users\Joao Branco\OneDrive - Just Group\Financial Resilience Team - General\00 Ad Hoc\MACalc_Batch`. Relative to the Draft it is `..\..\00 Ad Hoc\MACalc_Batch`.
- Git: HEAD `3415f59` (17 March 2026, "Fixes bug when running reop in base as well as test 2 inflation"). The working copy has **uncommitted changes** to `MACalcrunmanager.py`, `MACalc.py`, `MACE.py`, `MACalcutils.py`, `test2.py`, the bundled `MACalcinPython.xlsb` and both submodules. The model that runs is the working copy, not a commit.
- Entry: `MACalcrunmanager.main(root_repository=None, name_substring='MACalc', use_binary_optimisation=False, run_optimisation=False, is_downgrade_run=False, floor_cod_at_zero=True, argv=None)`. Called with keyword arguments and `argv=None`, it uses them as given. `DEFAULT_ROOT_FOLDER` is hard-coded to an old `Bonds Lite output\20261103_161318` folder, so the root must always be passed.
- `find_workbooks(root)`: `rglob('*')`; files only; skips `~$*`; extension `.xls/.xlsb/.xlsm/.xlsx`; name contains `MACalc` (case-insensitive); stem must not contain `_opt`; sorted.
- `confirm_workbooks(workbooks)` prints the list and calls `input('Proceed with these workbooks? [y/N]: ')`; only `y`/`yes` proceeds. On cancel or when nothing is found, `main` returns `[]` and the process exits 0.
- Run: a new hidden Excel (`xw.App(visible=False, add_book=False)`). For each workbook it opens it read-write with `update_links=False`, sets `prm_Reoptimise` to FALSE, recalculates, calls `MACalc.main(wb=..., use_binary_optimisation, is_downgrade_run, floor_cod_at_zero)`, **saves in place** (raises `PermissionError` if Excel opened it read-only, for example because it is open elsewhere), then writes `<stem>_unoptimised_assets.csv` beside it from `O. PuLPoutputAssets`, excluding Mortgage and ILS rows.
- Failures: each workbook's traceback is collected, and if any failed `main` raises `SystemExit(1)`. Success returns `[(workbook, None, csv_path, None)]` when optimisation is off.

### Base-run behaviour with the agreed options

- `is_downgrade_run` is read only inside `MACE.reoptimise` (`MACE.py` line 776), which runs only in the optimisation branch (`MACalc.py` line 210: `prm_Reoptimise` or `prm_ReoptimiseInTest2Infl`). With the pass off and `prm_ReoptimiseInTest2Infl` FALSE, the flag has no effect on the results (Decision 63, finding).
- `tbl_ForcedCompAperc`, `tbl_MinCompAperc`, `lst_pre16assets` and `lst_noPRoC` are also read only in that branch. `lst_REITs` is read on every run (`calcSCR.calccreditSCR`).
- LPI workbooks are opened only if Test 2 inflation is on or the input and run inflation parameters differ. In this output both LPI names are `202606 - LPIcurves_v2.1.0.xlsb` and all six PCs are 0, so no LPI workbook is opened.
- `prm_CloudRun` is TRUE (`prm_PythonEnv` = `azurelocalprd`), but it is used only by Test 2, and `prm_RunTest2IntRate` and `prm_RunTest2Infl` are FALSE. The base run is local.
- The base run writes `O. BaseMAP`, `O. PuLPoutputAssets`, `O. AllocatedDefAdjAssetCFs`, `O. UnallocatedGrossAssetCFs` and `O. SIIMAP`; all exist in the Stage 3 output. It asserts that Mortgages + Illiquids + Corporates + Govvies + Inflation Swaps = all assets (the Decision 46 guard covers this).

### Python environment

| Interpreter | Python | xlwings | numpy | pandas | scipy | pulp | smithwilson |
|---|---|---|---|---|---|---|---|
| `%LOCALAPPDATA%\anaconda3\python.exe` (controller) | 3.13.5 | 0.32.1 | 2.1.3 | 2.2.3 | 1.15.3 | missing | missing |
| `%LOCALAPPDATA%\anaconda3\envs\py311\python.exe` | 3.11.13 | 0.36.14 | 1.26.4 | 2.2.3 | 1.14.1 | 3.2.2 | 0.2.0 |

- No environment has the pinned xlwings 0.19.5. `py311` matches every other pin in `requirements.txt` (numpy 1.26.4, pandas 2.2.3, scipy 1.14.1, pulp 3.2.2).
- The run manager uses `with xw.App(...)` and `books.open(..., ignore_read_only_recommended=True)`, which need a newer xlwings than 0.19.5, and the model folder holds `cpython-311` bytecode for every module. The pin is stale; `py311` is the environment in use.
- With `-B` (no bytecode written), `import MACalcrunmanager, MACalc, calcSCR, MACE, pulp, xlwings` succeeds under `py311` from the model folder. It is the same interpreter as `BONDS_LITE_MODEL_PYTHON`.

### Decision 51 list contract (verified on `SST Output_Test\Credit\MACalc input_Test_4a0b639930e7c09c.xlsb`, opened read-only)

| Check | Model code | Live output | Result |
|---|---|---|---|
| `tbl_Assets` is read dynamically, all funded and asset rows | `names('tbl_Assets')` with `header=1`; sizes from `len(df_assets)` | `OFFSET('I. Assets'!A1,0,0,COUNTA(A:A),COUNTA(1:1))` gives `A1:ATQ1013`: header + 1,012 rows (11 Mortgage, 8 ILS, 5 funded, 988 assets), 1,213 columns, no blank A, D header `ISIN`, D unique, nothing in A or D below row 1013 | Pass |
| `tbl_ForcedCompAperc` joins D exactly | left merge on `ISIN` (= D) | 24 entries, 24 matched, 0 duplicates | Pass |
| `tbl_MinCompAperc` joins D exactly | left merge on `ISIN` | 164 entries, 106 matched, 0 duplicates | Pass |
| `lst_pre16assets` | exact membership (`in`) | 322 entries, 31 matched | Pass |
| `lst_REITs` | exact membership, every run | 16 entries, 16 matched | Pass |
| `lst_noPRoC` unchanged | exact membership | 8 entries, 8 matched (the ILS rows) | Pass |
| Duplicates and unmatched keys | merge duplicates would misalign rows; unmatched means "not forced" or "not a member" | no duplicates in any list; unmatched entries are ignored, as in Slice C (C2) | Pass |
| Missing names | bare `except` gives all-False silently | all five names exist and resolve | Pass (risk noted) |
| Rating suffix | MACalc never parses D; it compares exactly. The CSV export strips `+`/`-` from the `Rating` column (I) for its own output only | unaffected | Pass |

The counts equal the accepted Slice C summary (24/106/31/16), so Slice C is not reopened.

### Live handoff (saved Draft, read with openpyxl at 13:50)

- `SST_Stage1_Status` = `Succeeded`, `SST_Stage1_LastSuccessfulOutput` = `SST Output_Test/Credit`.
- `SST_Stage2_Status` = `Succeeded`, `SST_Stage2_InputRun` = `SST_Stage2_LastSuccessfulResults` = `SST Output_Test/Credit`.
- `SST_Stage3_Status` = `Succeeded`, `SST_Stage3_InputRun` = `SST Output_Test/Credit`, `SST_Stage3_LastSuccessfulOutput` = `SST Output_Test/Credit/MACalc input_Test_4a0b639930e7c09c.xlsb`.
- `SST_Credit_DG_standalone` = TRUE, `SST_Credit_spread_standalone` = TRUE, suffix `Test`.
- `SST Output_Test\Credit` contains one file matching MACalc Batch's rule: the audited workbook. `JRL Bonds\`, `JRL MAP\` and `.bonds-lite-*\` contain no `MACalc*` workbook.

## Proposals Awaiting Approval

The user approved these design choices on 4 November 2026.

- **P4-1. Two modules, two interpreters (Decision 68).** `src/sst_stage_four.py` runs in the controller's Python 3.13. It attaches to the open Draft, checks the audit, reads the flag, resolves the model settings, runs a preflight, launches the child and reports the result. `src/macalc_batch_entry.py` runs in the model's Python 3.11: it imports the run manager, guards the confirmation, runs and checks the artefacts. The controller cannot import MACalc (no pulp, numpy 2.1). This mirrors `bonds_lite_model_entry.py`.
- **P4-2. Reuse per-user Python setting (Decision 69).** Resolve `BONDS_LITE_MODEL_PYTHON` using the existing process-then-HKCU registry logic; do not add `SST_MACALC_PYTHON`. Each direct-source user can configure their own interpreter path, so no user's machine-specific path is embedded in the code. `SST_MACALC_DIR` remains optional, defaulting to `<Draft folder>\..\..\00 Ad Hoc\MACalc_Batch`; it must contain `MACalcrunmanager.py`. If MACalc Batch later needs a different environment from Stage 2, add a separate setting then.
- **P4-3. Confirmation guard (Decision 70).** In the child process only, replace `MACalcrunmanager.confirm_workbooks` at run time with a function that receives the list `main` is about to run. It prints the list and returns True only when the list is exactly `[audited workbook]`, compared with `os.path.normcase` on resolved paths. No model file changes. The entry also calls `find_workbooks` before `main`, to stop early with a clear message.
- **P4-4. Keep the model folder untouched and record what ran (Decision 71).** Run the child with `-B`, `PYTHONDONTWRITEBYTECODE=1` and the model folder as working directory. Log the model's git HEAD, whether the working copy is modified, and the SHA-256 of `MACalcrunmanager.py`, `MACalc.py`, `MACE.py`, `calcSCR.py`, `calcmp.py` and `MACalcutils.py`.
- **P4-5. Preflight before any model run (Decision 72).** Stop if a `~$<workbook name>` lock file exists beside the audited workbook. Probe the model interpreter: Python 3.11 and `import MACalcrunmanager` succeed from the model folder.
- **P4-6. Terminal command (Decision 73, superseded by Decision 76).** Stage 4 is started like Stages 1-3: the **Run MACalc** button calls `LaunchBondsLite 4`, which runs the controller with `--caller <Draft> --excel-hwnd <hwnd> --stage 4`. Maintainers in source mode use the same `python -m bonds_lite_input ... --stage 4` command that was used for Stage 3. The Draft must be open in desktop Excel, so the live audit is read. Stage 4 writes only its own `SST_Stage4_*` audit cells, and saves the Draft as Stage 3 does. The separate `python -m sst_stage_four --draft` command is dropped.
- **P4-7. Replace one automatic pre-run backup (Decision 74).** Before every run, copy the audited workbook to `SST Output_<suffix>\Stage 4 backup\<same workbook name>`. Replace that same backup on each run; do not add timestamped versions. Keep the backup outside `Credit`, so MACalc Batch's recursive scan cannot mistake it for another input. If the backup cannot be created or replaced, stop before launching MACalc Batch.
- **P4-8. No timeout; stream output (Decision 75).** The run duration is unknown, so stream the child's output to the console as Stage 2 does and add no deadline. Ctrl+C stops the child. MACalc's hidden Excel is a separate COM process, so an interrupted run may leave it holding the workbook (see Risks).

## Stage 4 Button And Audit Panel (Decisions 76-77, approved)

Mirrors Stage 3 (`sst_stage_three.py`: `STAGE_THREE_*` constants, `_install_stage_three_panel`, `_stage_three_button_row`, `install_stage_three_controls`, `_install_button`, `write_stage_three_audit`, `run_stage_three`).

| Item | Stage 3 (existing) | Stage 4 (planned) |
|---|---|---|
| Button name / caption | `btnPrepareMACalc` / Prepare MACalc | `btnRunMACalc` / Run MACalc |
| VBA macro | `PrepareMACalc` → `LaunchBondsLite 3` | `RunMACalc` → `LaunchBondsLite 4` |
| Panel name / heading | `SST_Stage3_Panel` / Stage 3 - Prepare MACalc | `SST_Stage4_Panel` / Stage 4 - Run MACalc |
| Audit fields | `SST_Stage3_{LastRun, Start, Finish, Elapsed, User, Status, InputRun, LastSuccessfulOutput}` | `SST_Stage4_` + the same eight |
| `InputRun` | Stage 1 run folder | portable path of the audited Stage 3 workbook |
| `LastSuccessfulOutput` | MACalc input workbook | portable path of `<stem>_unoptimised_assets.csv` |
| Placement | below the Stage 2 panel, button beside or below it | below the Stage 3 panel, button beside or below it, with the same occupancy and overlap checks |
| Installer | `bonds_lite_input --install-stage-three <Draft>` | `bonds_lite_input --install-stage-four <Draft>`; Draft closed; backup to `references\backups`; requires the Stage 3 panel |
| Dispatch | `--stage 3` → `run_stage_three` | `--stage 4` → `run_stage_four` |
| Packaged runtime | controller only | controller + new `MACalc\MACalcModel.exe` (Decision 77) |

Audit sequence (same as Stage 3):
1. Take a run guard on the Draft.
2. Resolve the handoff (Decision 65). A stale or ambiguous handoff stops here, before any audit write, as Stage 3 does when Stage 2 is not ready.
3. Write `Running`, with `InputRun` set and the previous `LastSuccessfulOutput` kept.
4. Run the preflight, the backup (Decision 74) and MACalc.
5. Write `Succeeded` with the CSV path, or `Failed`/`Interrupted` with the previous output kept.

The Draft is saved after each write; an unsaved audit is an error.

Until a release containing Stage 4 is promoted, the button in the live Draft calls release `20261104-cache-v9`, which rejects `--stage 4` with a console error. Stage 3 had the same interim state.

## Steps

1. **Approval (complete).** P4-1 to P4-8 and the Stage 4 button details were approved on 4 November 2026 (Decisions 68-77). Implementation can proceed from the test-first step.
2. **Tests first (Supir Test)** (depends on 1). Create `tests/test_sst_stage_four.py` with the tests in "Test Strategy", using fake Draft books and a fake `MACalcrunmanager` module written to `tmp_path`. Extend `tests/test_bonds_lite_input.py` with the `--stage 4` dispatch, `--install-stage-four` and VBA launcher cases, and the packaged-builder tests with the MACalc app. Pass: the new tests fail only because the code is missing; existing tests are unchanged. [✓ Implemented — see addendum]
3. **Handoff resolution** in `src/sst_stage_four.py` (depends on 2): `resolve_stage_four_input(caller) -> StageFourInput(run_folder, workbook, is_downgrade_run)`. [✓ Implemented — see addendum]
   - Reuse `bonds_lite_input.resolve_stage_two_results` (Stage 2 `Succeeded` for the current Stage 1 run).
   - Read the Stage 3 audit through the existing `SST_Stage3_*` names in `sst_stage_three.py`. Require status `Succeeded` and `SST_Stage3_InputRun` equal to the current run's portable path; otherwise stop with "rerun Stage 3" (Decision 65).
   - Resolve `SST_Stage3_LastSuccessfulOutput` relative to the Draft folder. It must be an existing, non-symlinked file directly inside the run folder, and its name must match `MACalc input_<suffix>_<16 lowercase hex>.xlsb`, with `<suffix>` equal to the parent `SST Output_<suffix>` suffix. A path in `Stage 3 staging` is rejected.
   - Read `SST_Credit_DG_standalone` with the Stage 1 reader rules (Excel boolean only).
   - Pass: the handoff and flag tests pass.
4. **Settings and preflight** (parallel with 3): reuse the existing `BONDS_LITE_MODEL_PYTHON` resolver and add `load_macalc_config(caller_path)` for the optional model directory (P4-2); implement `preflight_macalc_environment(config)` (P4-5), the lock-file check and model provenance (P4-4). Pass: settings and preflight tests pass for environment and registry configuration, including a different per-user Python path. [✓ Implemented — see addendum]
5. **Model-side entry** `src/macalc_batch_entry.py` (parallel with 3): arguments `--model-dir`, `--root`, `--expected`, `--downgrade-run`/`--no-downgrade-run`. Insert the model folder into `sys.path`, import the run manager, pre-scan, install the P4-3 guard, and call `main(root_repository=root, use_binary_optimisation=False, run_optimisation=False, is_downgrade_run=flag, floor_cod_at_zero=True)`. Exit 0 only if `main` returns exactly one entry for the expected workbook and its CSV exists and is not empty. Exit 2 for a guard or validation stop; propagate `SystemExit(1)` for model failures. Pass: entry tests pass. [✓ Implemented — see addendum]
6. **Backup, launch and report** in `sst_stage_four.py` (depends on 3-5): before launching, copy the audited workbook to the fixed backup path in `Stage 4 backup`, replacing the prior backup without creating versioned copies. Do not start the child if backup creation/replacement fails. Then build the command (`[python, "-B", entry, ...]`, `shell=False`, model folder as working directory, `PYTHONDONTWRITEBYTECODE=1`), stream the output, and turn a non-zero exit into a `StageFourError` that gives the code. Log input, backup path, options (downgrade ON/OFF, noting that it is inactive while optimisation is off), provenance, elapsed time and output paths. Pass: backup creation, replacement, failure-stop, launch and result tests pass. [✓ Implemented — see addendum]
7. **Audit, dispatch and launcher** (depends on 6):
   - In `sst_stage_four.py`, add the `SST_Stage4_*` constants (Decision 77), `require_stage_four_panel`, `write_stage_four_audit` and `run_stage_four(caller) -> Path | None`, all following the Stage 3 equivalents. The audit sequence is described in "Stage 4 Button And Audit Panel". `run_stage_four` returns the CSV path.
   - In `bonds_lite_input.main`, add `"4"` to the `--stage` choices and dispatch to `sst_stage_four.run_stage_four`, logging "Stage 4 succeeded; MACalc results saved". Add `--install-stage-four`.
   - Add `Public Sub RunMACalc()` / `LaunchBondsLite 4` to `VBA_LAUNCHER` and `PACKAGED_VBA_LAUNCHER`.
   - Pass: the audit, dispatch and launcher-text tests pass; Stage 1-3 tests are unchanged. [✓ Implemented — see addendum]
8. **Installer** (parallel with 7): add `install_stage_four_controls(caller) -> Path` and `_install_stage_four_panel` to `sst_stage_four.py`, mirroring `install_stage_three_controls`. [✓ Implemented — see addendum]
   - Refuse to run while the Draft is open, and back it up to `references\backups` first.
   - Require the Stage 3 panel, and place the panel below it with the same occupancy and overlap checks.
   - Place `btnRunMACalc` beside or below the panel, using the Stage 3 row rule.
   - Insert `RunMACalc` into `modBondsLiteLauncher` only if it is missing.
   - Re-running is safe: complete names are validated, and incomplete names stop.
   - Pass: installer tests pass on fakes, and the opt-in Excel test passes on a disposable Draft copy.
9. **Docs** (parallel with 7-8): in `references/bonds-lite-setup.md`, add a "Stage 4 - Run MACalc" section. It should cover: [✓ Implemented — see addendum]
   - the button and panel;
   - the installer command;
   - the source-mode command (`--stage 4`), the per-user `BONDS_LITE_MODEL_PYTHON` setting and the optional `SST_MACALC_DIR`;
   - the single-workbook rule, in-place behaviour and the automatic single-file backup;
   - how to clear earlier `MACalc input_*` files;
   - troubleshooting.

   Also change "Stage 5" to "Stage 4". Add an implementation addendum to this plan and update the handoff.
10. **Automated verification** (depends on 2-9): targeted `python -m pytest -q -p no:cacheprovider tests/test_sst_stage_four.py tests/test_bonds_lite_input.py`, then the full suite. Expected: the baseline (1022 passed, 88 skipped, the same 2 known failures) plus the new tests passing. [✓ Implemented — see addendum]
11. **Install the panel and button** (depends on 10 and the user's go-ahead). With the Draft closed, run `python -m bonds_lite_input --install-stage-four "<Draft>"` from `src`. Pass: a backup exists in `references\backups`; the Stage 4 panel shows `Not Run`; the button sits below the Stage 3 controls without overlap; the Stage 1-3 panels and buttons are unchanged. [✓ Implemented — see addendum]
12. **Live run, source mode** (depends on 11 and the user's go-ahead), on `SST Output_Test\Credit`, with the Draft open in the add-in-enabled Excel session. Use source mode because the live Draft's button calls release v9, which has no Stage 4. Set `--caller` to Excel's exact `Workbook.FullName`; for a OneDrive-synced Draft this may be a SharePoint URL. Pass the local synced path separately with `--caller-local`. [✓ Implemented — see addendum]
    ```powershell
    python -u -m bonds_lite_input --caller "<Excel Workbook.FullName>" --caller-local "<local Draft path>" --excel-hwnd <hwnd> --stage 4
    ```
    - Before: re-read the live audit. Confirm there is one `MACalc*` workbook in the folder and no `~$` lock.
    - Expect on the console:
      - the backup is replaced in `SST Output_Test\Stage 4 backup`;
      - one workbook is found and confirmed automatically;
      - binary optimisation OFF, downgrade ON (inactive), optimisation pass OFF, CoD floor ON;
      - `Processed 1 workbook(s).`;
      - exit 0.
    - Expect on disk:
      - the workbook is saved later than 13:43:05 (its Stage 3 time);
      - `MACalc input_Test_4a0b639930e7c09c_unoptimised_assets.csv` is beside it, with columns `Index, ISIN, Name, Zspread, Type, Sector, Rating, Index Linked?, PV, 1..100` and 993 rows (634 Bond + 182 Infrastructure + 177 PP);
      - there is no `~$` file and no hidden `EXCEL.EXE` left over.
    - Expect in the Draft:
      - `SST_Stage4_Status` = `Succeeded`;
      - `InputRun` = the workbook, and `LastSuccessfulOutput` = the CSV;
      - the Stage 1-3 audit cells are unchanged.
    - Rerun Stage 4 once: it succeeds again, the folder has the same files with the CSV replaced (Decision 66), and the backup is replaced, not duplicated.
    - Hand the figures to the user for the reasonableness review (Decision 64).
13. **Packaged release** (depends on 12, the user's acceptance and separate approval, as for Stage 3). In `tools/build_packaged_stages.py`:
    - Add `--hidden-import sst_stage_four` to the controller build.
    - Build `MACalcModel` from `src/macalc_batch_entry.py` with the model Python. Add the MACalc Batch folder to the search path and collect `pulp` data, which includes the solver.
    - Add a `--self-test` that imports the run manager.
    - Record the MACalc source hashes in `release.json` and check that they did not change during the build.
    - Make `load_macalc_config` resolve `MACalc\MACalcModel.exe` when the controller is frozen.
    - Build under a new release ID, then promote only with approval. Then click **Run MACalc** in the Draft and repeat the step 12 checks.
    - Pass: the builder tests pass; the self-tests in the built release pass; the packaged button run succeeds.
    - Gate status: the user requested the start of step 13 on 4 November 2026. Packaging code and automated tests are implemented; commit, build, promotion and the packaged button run still require separate approval. Step 13 is not complete until the built release and packaged run pass.

## Test Strategy

Write all tests before the code (testing-strategy skill). No Excel or real model is needed: fake caller books expose names and values, and a fake `MACalcrunmanager.py` in `tmp_path` records its calls.

1. **Handoff** (`resolve_stage_four_input`):
   - success with the live shape (`SST Output_Test/Credit/MACalc input_Test_<16 hex>.xlsb`);
   - Stage 3 status `Failed`, `Running`, `Interrupted`, `Not Run` or blank: stop, and the message says to rerun Stage 3;
   - `SST_Stage3_InputRun` differs from the current Stage 1 run: stop;
   - Stage 2 not `Succeeded` for the current run: stop;
   - last output blank, missing, a folder, a symlink, in a subfolder, in `Stage 3 staging`, or in another stress folder: stop;
   - bad names: `MACalc input.xlsb`, a 15- or 17-character token, uppercase hex, suffix mismatch (`MACalc input_Other_<hex>.xlsb` under `SST Output_Test`): stop;
   - a `~$` lock beside the workbook: stop.
2. **Flag:** TRUE gives True; FALSE gives False; `"TRUE"`, `1`, blank or a missing name stop.
3. **Settings:** missing/invalid `BONDS_LITE_MODEL_PYTHON` fails through the existing resolver; verify the configured per-user interpreter is used rather than a hard-coded local path. `SST_MACALC_DIR` unset gives `<Draft>\..\..\00 Ad Hoc\MACalc_Batch`; a folder without `MACalcrunmanager.py` stops.
4. **Command and environment:** exact argv for TRUE and FALSE flags; `-B` present; `PYTHONDONTWRITEBYTECODE=1`; working directory = model folder; `shell=False`.
5. **Entry, using the fake run manager:**
   - the list `[expected]` proceeds without calling `input()`;
   - `[]`, `[expected, other]` and `[other]` stop with exit 2, list the workbooks and never call `MACalc.main`;
   - a case-only path difference is treated as equal;
   - `main` receives `run_optimisation=False`, `use_binary_optimisation=False`, `floor_cod_at_zero=True`, the flag and the root;
   - `main` returning `[]` exits 2; `SystemExit(1)` exits 1 and the traceback is printed;
   - success but a missing or empty CSV exits 2;
   - provenance is printed: HEAD, the modified state, and six hashes (using a fake git repository or a stubbed git call).
6. **Orchestration:** child exit 0 gives success; child exit 1 or 2 gives a `StageFourError` with the code; Stage 4 writes only `SST_Stage4_*` names in the fake Draft, never another stage's.
7. **Backup:** the source workbook is copied to the fixed path outside `Credit`; a later run replaces the same backup; no timestamped backup accumulates; copy/replace failure prevents child launch.
8. **Audit:**
   - A stale handoff stops before any audit write.
   - After that, the sequence is `Running`, then `Succeeded`, with `InputRun` = the portable workbook path and `LastSuccessfulOutput` = the portable CSV path.
   - A child failure gives `Failed`, and Ctrl+C gives `Interrupted`. Both keep the previous `LastSuccessfulOutput`.
   - A missing or incomplete `SST_Stage4_*` panel stops the run with "install the Stage 4 controls".
   - An unsaved Draft after a write is an error.
   - Unknown audit fields are rejected.
9. **Dispatch and VBA:** `--stage 4` calls `run_stage_four` with the matched caller; `--stage 5` is rejected; both launcher texts contain `Public Sub RunMACalc()` and `LaunchBondsLite 4`; Stage 1-3 subs are unchanged; `--install-stage-four` calls the installer and logs the backup.
10. **Installer:**
    - On fakes: the Draft is open → stop; the Stage 3 panel is missing → stop; the area is occupied or merged → stop; it overlaps a shape → stop; incomplete names → stop; a complete existing install → validated, not duplicated; `RunMACalc` is inserted once.
    - Opt-in Excel test on a disposable Draft copy: the panel, names, button and macro exist after a save and reopen, and the Stage 1-3 controls are unchanged.
11. **Packaged builder:** the controller command includes `sst_stage_four`; the MACalc app command uses the model Python, the MACalc folder path and `pulp` data; frozen `load_macalc_config` finds `MACalc\MACalcModel.exe` and stops with a clear message when it is missing; a manifest hash mismatch stops publication.
12. **Live runs** (steps 12-13): the real data check. Fixtures copy no production values.

## Affected Files

- `src/sst_stage_four.py`: new controller module (handoff, flag, settings, preflight, backup, launch, audit, `run_stage_four`, panel/button installer).
- `src/macalc_batch_entry.py`: new model-side entry for Python 3.11 (also the packaged `MACalcModel` entry, step 13).
- `src/bonds_lite_input.py`: `--stage 4` choice and dispatch, `--install-stage-four`, and `RunMACalc` in `VBA_LAUNCHER` and `PACKAGED_VBA_LAUNCHER`. Stage 1-3 behaviour is unchanged.
- `tests/test_sst_stage_four.py` (new) and `tests/test_bonds_lite_input.py` (dispatch, launcher text and builder cases).
- `Draft Input spreadsheet for SST.xlsm`: Stage 4 panel, button and `RunMACalc` macro, at step 11 only, after a backup and with the user's go-ahead.
- `tools/build_packaged_stages.py`: the MACalc app and the controller hidden import, at step 13 only, with separate approval.
- `references/bonds-lite-setup.md`, `references/decisions.md`, this plan and the Stage 4 handoff.
- **Not changed:** `MACalc_Batch` (code, workbook, environment), `src/sst_stage_three.py` (read-only reuse of its helpers), the Stage 1-3 panels and buttons, and the active release v9 until a new release is approved.

## Decisions

- Decisions 60-75: the user's answers to Q4-1 to Q4-7, the confirmation rule, proposals P4-1 to P4-8, and the status note on Decision 51.
- Decision 76: the Stage 4 button, in the same fashion as the other stages. It supersedes parts of Decisions 60 and 73.
- Decision 77: the button, panel and packaging details, approved by the user.

## Risks And Unknowns

- **Stage 3 output changed in place (Decision 61).** After Stage 4, the file named in `SST_Stage3_LastSuccessfulOutput` holds MACalc's results and `prm_Reoptimise` = FALSE (the template had TRUE). The one-file backup (Decision 74) holds the workbook as it was immediately before the most recent Stage 4 run and is overwritten on every run. Thus, if Stage 4 is rerun without Stage 3, the backup contains the prior Stage 4 result, not a clean Stage 3 input. Rerun Stage 3 to get a clean input again.
- **Earlier MACalc inputs block Stage 4.** Stage 3 keeps earlier outputs (Decision 54), and MACalc Batch would run all of them, so Stage 4 stops (Decision 62) until the user removes or archives them. If this becomes a nuisance, a later change could make Stage 3 remove earlier outputs, or move them out of the scan.
- **Uncommitted model code.** MACalc Batch runs from a modified working copy on top of `3415f59`. Results depend on its state on the day; P4-4 records it.
- **Stale xlwings pin.** `requirements.txt` pins 0.19.5, but the code needs a newer xlwings and `py311` has 0.36.14. Confirm in the live run.
- **Fresh Excel instance.** MACalc Batch opens a new hidden Excel, without the user's add-ins (such as the JRC XLL), with `update_links=False`, and recalculates. If any input that MACalc reads depends on an XLL or external link, it could evaluate to an error. Not yet verified; the live run will show it.
- **Interrupted runs.** MACalc's hidden Excel is not a child process of the launcher. After Ctrl+C or a crash it can keep the workbook open, so later runs fail with `PermissionError` until that Excel is ended (by its process ID).
- **Button before release.** After step 11, clicking **Run MACalc** in the live Draft calls release v9, which rejects `--stage 4` with a console error until a Stage 4 release is promoted (step 13). Stage 3 went through the same interim state. Tell colleagues, or install the button only when the release is ready.
- **Packaged MACalc snapshot.** The packaged `MACalcModel.exe` freezes the MACalc Batch sources at build time. Today they are an uncommitted working copy (see above), so the release manifest's hashes are the only record of what colleagues run. PyInstaller may also miss `pulp`'s solver binaries or `smithwilson` data; the build self-test must import them, and the first packaged run must be checked.
- **Stale Stage 4 audit after reruns.** A Stage 1-3 rerun does not reset the Stage 4 panel. `SST_Stage4_InputRun` shows which Stage 3 workbook the last Stage 4 result belongs to; any later stage must check it, as Stage 4 checks Stage 3.
- **Silent list fallback.** If a list name ever goes missing, MACalc treats it as empty without warning. Stage 3's read-back currently guarantees that the names exist.
- **Downgrade flag inactive today, and read late.** The flag changes nothing until optimisation is on (Decision 63). It is read from the Draft at Stage 4 time, so it may differ from the value Stage 1 used if the Draft was edited in between. Revisit when the optimisation flag is added.
- **Cloud and Test 2.** `prm_CloudRun` is TRUE in the template. If a future template turns Test 2 on, MACalc would try the Azure run path. Out of scope; the launcher logs both Test 2 parameters.
- **`_opt` suffix** (Decision 67): no guard. MACalc finds nothing and Stage 4 stops with "no workbook found".
- **Run time** is unknown (no timeout, P4-8). Record it in the live run.

## Addendum: Implementation Progress

### Completed

- Steps 1-9: added the test-first Stage 4 controller and model-side entry, validated Stage 2/3 handoffs and settings, protected the child run with a single-workbook guard, implemented backup/audit/CLI/VBA/installer behaviour, and documented source-mode usage.
- The Stage 4-focused suite passes: **66 passed, 1 skipped**.
- The full suite reports **1,092 passed, 90 skipped, 2 failed**. The two failures are the known baseline failures recorded in the handoff: `test_frozen_main_validates_cached_bundle_before_stage_dispatch` and `test_model_results_transaction_rollback_failure_retains_recoverable_backup`.
- The packaged-builder test was initially skipped pending step 13; it is now enabled and passes with the additional packaging tests below.

### Step 13: packaging implementation

- Completed the test-first implementation after the user's instruction to start
  step 13. The initial focused run reported 23 failures and 3 passes; the
  implemented packaging selection now reports **52 passed**.
- The controller includes `sst_stage_four`. The builder requires `--macalc-dir`,
  builds `MACalcModel` with the model Python, includes the LPI import paths and
  collects PuLP (including CBC) and Smith-Wilson resources. The Bonds model's
  build options remain separate.
- The manifest records the six MACalc source hashes and executable hash;
  source drift prevents publication. Cached bundle validation also checks the
  MACalc executable when its manifest fields are present, while accepting
  older Stage 1-3 manifests.
- Frozen MACalc execution uses the bundled import directory, does not require
  loose model sources or runtime git, and restores the working directory.
  The self-test imports the model and checks CBC without opening Excel.
- Added pinned PuLP and Smith-Wilson requirements for the isolated build
  environment. Both are currently absent there; no dependencies were installed.
  The external model environment and model sources were not changed.
- Verified the accepted six source hashes are unchanged. The source self-test
  prints `MACALC_SELF_TEST_OK`; this is not a frozen-release self-test.
- Full suite: **1,114 passed, 89 skipped, 2 known baseline failures** (the same
  two named above). Black passed for the new modules/builder and changed ranges
  of legacy files; `git diff --check` passed.
- No commit, build, promotion, rollback or live workbook run was performed.
  The active release remains `20261104-cache-v9`. Next gate: ask whether to
  commit Stage 4 source, tests and docs without pushing.

### Deferred gates

- Step 10 automated verification is complete against the existing baseline. The opt-in Excel installer test was not performed; the live model run is recorded under step 12 below.
- Step 11 was completed on 4 November 2026 after the user's go-ahead. The installer created `references\backups\Draft Input spreadsheet for SST_20261104_151920_099479.xlsm`; the Stage 4 panel reads `Not Run`. Read-back verification confirmed the button is below Stage 3 and the Stage 1-3 audit names, values, button captions, targets and cell anchors are unchanged.
- Step 12 completed in source mode after the user's go-ahead. MACalc processed the single audited workbook twice, each run taking about 51 seconds, and returned exit code 0. The output CSV has the expected 109 columns and 993 rows (634 Bond, 182 Infrastructure and 177 PP); no ISIN or PV values are blank. The Stage 4 audit is `Succeeded`, and the Stage 1-3 audits remain unchanged.
- The final fixed backup contains the workbook as it stood before the second run (the first Stage 4 result), not the original Stage 3 input. There is one backup and one output CSV; no lock file or extra Excel process remained after the run.
- Model provenance for both runs: HEAD `3415f599b21ebb2cb2aeebd190b7c3ff15ae730f`; working copy modified. SHA-256: `MACalcrunmanager.py` `81b5a8319b2c788d0d9dc1604dca10b2db083667406d287ab76dcb53aea095fb`; `MACalc.py` `61f5390b187b95a102010c27fe1b432594d3647c38126772a3703202b9603180`; `MACE.py` `afffd967be348874226796f3686e199f41534a6dda972abd9ba00fdbf450c7de`; `calcSCR.py` `4dc150e7f747dab9a1003a2dccbe46132964a2e6d6578870d410be0d852a1859`; `calcmp.py` `25f42b064593e4a0f0054ea0edaf9b2828490181b6a14f68257ebb0140b133e2`; `MACalcutils.py` `2a54454fc004279bf1ae5bb468d69a5284d96a19d4387b7a616d1fa92ce22026`.
- The first launch attempt used the local Draft path as `--caller` and stopped before audit or model execution because Excel exposed `Workbook.FullName` as a SharePoint URL. The successful runs used that exact URL with `--caller-local` set to the local synced path.
- Step 12 results were **accepted** by the user on 4 November 2026 (about 15:43 UTC: "It looks to be running okay!"), which satisfies the Decision 64 reasonableness review. MACalc Batch was run but not changed.
- Step 13's builder and frozen-runtime changes are implemented and automatically verified, as recorded above. The package build, release promotion and packaged acceptance run remain pending. Each gate (commit, build, promote, packaged run) needs separate explicit approval. The detailed brief is in the handoff under "Next Session: Promote Stage 4".

### Deviations

- No implementation deviations. The packaged-builder test was skipped until step 13, as required by the live acceptance gate; it is now enabled. Build and promotion have not been treated as authorised by the instruction to start implementation.