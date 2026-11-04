# Session Handoff: Stage 4 MACalc Batch

Version 2.4 | Prepared 3 November 2026, updated 4 November 2026 | Step 13 code verified; next gate: commit approval, then build

> **Start here in the next session:** read "Next Session: Promote Stage 4"
> below. Steps 1-12 are done and the user has accepted the live source run.
> Step 13's packaging code and automated checks are now done; the release
> has not been built or promoted. Everything before that section is history.

## Where We Stopped (4 November 2026)

1. Stage 4 discovery of MACalc Batch was done (read-only). The findings are in
   "MACalc Batch Discovery" below.
2. Five planning questions were asked. The user then paused Stage 4 to change
   Stage 1 (credit stress flags and a new output layout). That change is
   implemented, committed, pushed, released as `20261104-cache-v9`, and the
   user's packaged Stage 1-3 rerun worked as expected.
3. Q4-1 to Q4-7 were answered and recorded as Decisions 60-67. The Decision 51
   checks and the environment check (`py311`) passed read-only.
4. **Historical update (4 November 2026, 14:22 UTC):** The user approved P4-1 to
   P4-8; Decisions 68-75 record them. Implementation had not started at that
   point. Nothing was committed.
5. **Update (4 November 2026, 14:30 UTC):** The user asked for a Stage 4 button
   in the same fashion as Stages 1-3 (Decision 76), then approved the button,
   audit and packaging details (Decision 77). Plan v0.4 adds:
   - the `SST_Stage4_*` panel, the **Run MACalc** button and the `RunMACalc` macro;
   - `--stage 4` dispatch and an installer;
   - a source-mode live run, then a packaged release that needs separate approval.

   At that point, the Draft had not been changed; installing the button was
   assigned to plan step 11.
6. **Implementation update (4 November 2026):** The Stage 4 source controller,
   guarded MACalc Batch entry point, CLI dispatch, VBA launcher text and installer
   are implemented. Stage 4-focused tests pass (**66 passed, 1 skipped**); the full
   suite reports **1,092 passed, 90 skipped, 2 known baseline failures**. At that
   time, the Draft controls, live MACalc run and packaged builder/release were
   still untouched pending their respective plan gates. See the implementation
   addendum at the end.
7. **Draft installation update (4 November 2026, 15:19 UTC):** After the user's
   go-ahead, the Stage 4 panel and button were installed in the Draft. The
   installer backup is
   `references\backups\Draft Input spreadsheet for SST_20261104_151920_099479.xlsm`.
   Read-back verification confirmed the Stage 4 status is `Not Run`, the
   `RunMACalc` button is below Stage 3, and the Stage 1-3 audit names/values,
   button captions, macro targets and cell anchors are unchanged. The live
   MACalc run has not been authorised or performed. The active packaged release
   still rejects `--stage 4`, so do not use the installed button until a
   Stage 4-capable release is promoted.
8. **Live source-run update (4 November 2026):** After the user's instruction to
   continue, Stage 4 ran twice against the single audited workbook. Both runs
   exited 0 in about 51 seconds. The CSV passed the plan's schema and count
   checks: 109 columns, 993 rows (634 Bond, 182 Infrastructure and 177 PP), with
   no blank ISIN or PV values. The saved Stage 4 audit is `Succeeded`; Stage 1-3
   audit values remain unchanged. The fixed backup now contains the first
   Stage 4 result from immediately before the second run, not the clean Stage 3
   input. One backup and one replaceable CSV remain; no workbook lock or extra
   Excel process remained after execution.

   Both runs used MACalc Batch HEAD
   `3415f599b21ebb2cb2aeebd190b7c3ff15ae730f` with a modified working copy.
   Source hashes: `MACalcrunmanager.py`
   `81b5a8319b2c788d0d9dc1604dca10b2db083667406d287ab76dcb53aea095fb`;
   `MACalc.py` `61f5390b187b95a102010c27fe1b432594d3647c38126772a3703202b9603180`;
   `MACE.py` `afffd967be348874226796f3686e199f41534a6dda972abd9ba00fdbf450c7de`;
   `calcSCR.py` `4dc150e7f747dab9a1003a2dccbe46132964a2e6d6578870d410be0d852a1859`;
   `calcmp.py` `25f42b064593e4a0f0054ea0edaf9b2828490181b6a14f68257ebb0140b133e2`;
   `MACalcutils.py` `2a54454fc004279bf1ae5bb468d69a5284d96a19d4387b7a616d1fa92ce22026`.

   The first launch attempt used the local Draft path for `--caller` and stopped
   before any audit or model changes: Excel exposes `Workbook.FullName` as a
   SharePoint URL for this synced file. The successful commands used that exact
   URL with `--caller-local` set to the local path. The setup guide now documents
   this.
9. **User acceptance (4 November 2026, about 15:43 UTC):** The user reviewed the
   live source run and said "It looks to be running okay!". This is the
   step 12 reasonableness review required by Decision 64. Step 12 is accepted.
   The user asked for this handoff so that Stage 4 can be packaged and promoted
   (plan step 13) in a new session. Step 13 itself has **not** been approved
   yet; see the next section.
10. **Step 13 implementation update (4 November 2026):** The user asked to
    read this handoff and start step 13. Test-first packaging implementation
    is complete. The builder now produces the MACalc app, records/checks its
    provenance and includes Stage 4 in the controller. Frozen execution and
    cached MACalc executable integrity are covered by tests. The packaging
    selection passes **52 tests**; the full suite reports **1,114 passed,
    89 skipped and the same 2 baseline failures**. No commit, build,
    promotion, rollback or live workbook run was performed. Ask about the
    commit gate next; do not infer build or promotion approval.

## Next Session: Promote Stage 4

### Status at handover

- Plan steps 1-12 in [plan-stage-four.md](plan-stage-four.md) are complete. The
  user accepted the step 12 live run (update item 9).
- Nothing from Stage 4 is committed. Source, tests and docs are uncommitted
  changes in the worktree (see "Uncommitted Stage 4 changes" below).
- **Active central release:** `20261104-cache-v9`. **Rollback:**
  `20261103-cache-v8`. v9 rejects `--stage 4`, so the **Run MACalc** button
  in the Draft fails with a console error until a Stage 4 release is promoted.
- The Draft has the Stage 4 panel (`H112:O122`) and button (`btnRunMACalc`,
  anchored at `D112:G115`, macro `RunMACalc`). Its Stage 4 audit says
  `Succeeded`, from the source run.
- Step 13 packaging code and automated tests are implemented. The accepted
  six model source hashes were rechecked and are unchanged. The source
  `--self-test` prints `MACALC_SELF_TEST_OK` and confirms CBC is available
  without opening Excel; the three frozen-release self-tests still await
  an approved build.
- The isolated `model-py311` build environment lacks `pulp` and
  `smithwilson`. Their accepted pins (`3.2.2` and `0.2.0`) are now in
  `tools\requirements-model-build.txt`; refresh only that isolated
  environment before the approved build. No installation was performed.

### Approvals required (ask in chat, one gate at a time)

1. **Commit first?** Recommend committing the Stage 4 source, tests and docs
   (Supir Commit, conventional message, Co-authored-by trailer) before
   building. That way the release can name a commit, as v9 named `84e2dae`.
   Leave the pre-existing dirty items alone (see "Working Tree And
   Permissions"). Do not push without approval.
2. **Build** (no `--promote`) under a new release ID, for example
   `20261105-cache-v10`. Use the next date if the build day changes. The ID
   must not already exist under `releases\`.
3. **Promote** (`current.txt` -> the new ID) only after both the build checks
   and the self-tests pass, **and** the user explicitly approves.
4. **Packaged button run** by the user (or with their go-ahead), then the checks
   below.
5. **Rollback** to v9 only with approval if the packaged run fails.

### Step 13 tasks (from the plan)

Test -> Implement is complete: tasks 1-5 below are implemented and verified.
Tasks 6-8 (build, promotion, packaged acceptance and release-status updates)
remain pending their gates. Retain the detailed checklist for the release.

1. **Tests.** Remove the `@pytest.mark.skip` from
   `test_packaged_builder_adds_macalc_model_application_and_pulp_runtime`
   in `tests/test_bonds_lite_input.py` (around line 1550). Add tests for:
   - the controller command including `--hidden-import sst_stage_four`;
   - the MACalc source hashes recorded in `release.json`;
   - the build being rejected if those hashes change during the build;
   - the frozen `load_macalc_config` resolving `MACalc\MACalcModel.exe`.
2. **Builder, `tools/build_packaged_stages.py`:**
   - Controller build: add `--hidden-import sst_stage_four` beside
     `sst_stage_three` (around line 61).
   - `pyinstaller_command` with `model=` currently adds
     `pyBondsHMutils.py` as data. That is Bonds-specific and does not exist
     in the MACalc folder. Separate the MACalc build (for example, a model
     "kind" or a separate options branch). The MACalc build adds
     `--paths <MACalc dir>`, `--collect-all pulp` (or `--collect-data pulp`
     plus binaries) and the MACalc modules as hidden imports. The skipped
     test fixes the expected contract: name `MACalcModel`, entry
     `src/macalc_batch_entry.py`, model dir in the command, `pulp`
     collected, and no `sst_stage_four` hidden import.
   - Add a `--macalc-dir` argument. Default it to the resolved MACalc Batch
     folder, or make it required, to match `--model-dir`. Build
     `MACalcModel` with `--model-python` (py311), move it into
     `Bonds Lite App\MACalc\`, and run
     `MACalc\MACalcModel.exe --self-test`.
   - Hash the six MACalc sources before and after the build. Record them in
     `release.json` (for example, `macalc_source_sha256`), along with
     `macalc_executable_sha256` and the updated `layout`.
   - Add `pulp` (and any other new model dependencies) to the
     `collect_licences` name list.
3. **Model entry, `src/macalc_batch_entry.py`:** add a `--self-test` flag. It
   must import `MACalcrunmanager` (with cwd/`sys.path` set to the bundled
   snapshot), import `pulp`, and print e.g. `MACALC_SELF_TEST_OK`. It must
   not open Excel or any workbook.
4. **Frozen config:** `load_macalc_config` in `src/sst_stage_four.py`
   (around line 84) already has a frozen branch for
   `MACalc\MACalcModel.exe`. Confirm it agrees with the bundle layout.
   Confirm no runtime git or hash check runs in frozen mode (Decision 77:
   `release.json` replaces it).
5. **Verify:** run the Stage 4 tests and the full suite (expect at least the
   baseline below), then Black on the changed files and `git diff --check`.
6. **Build** without `--promote` using the same environments as v9
   (`controller-py313`, `model-py311`). The model folder for Bonds is
   unchanged from v9. Confirm `MODEL_SELF_TEST_OK`,
   `CONTROLLER_SELF_TEST_OK` and the new MACalc self-test pass. Compare
   with v9: the Controller and Model file sets should be unchanged apart
   from the controller's new `sst_stage_four` module, and `MACalc\` is
   new. Check that the `release.json` MACalc hashes equal the provenance
   hashes below.
7. **Promote** with approval, then run the packaged-run checks below.
8. **Docs:** in `references/bonds-lite-setup.md`, remove the "Until a Stage 4
   controller release…" caveat and record the new release ID. Mark step 13
   ✓ in the plan and its addendum. Update this handoff. Record the release
   in `decisions.md` only if a new decision was made.

### Model provenance to compare against

MACalc Batch folder (read-only; do not modify):
`...\Financial Resilience Team - General\00 Ad Hoc\MACalc_Batch`, HEAD
`3415f599b21ebb2cb2aeebd190b7c3ff15ae730f` with a **modified working copy**.
The packaged build freezes that working copy. The `release.json` hashes are
then the only record of what runs.

| File | SHA-256 at the accepted live run |
|---|---|
| `MACalcrunmanager.py` | `81b5a8319b2c788d0d9dc1604dca10b2db083667406d287ab76dcb53aea095fb` |
| `MACalc.py` | `61f5390b187b95a102010c27fe1b432594d3647c38126772a3703202b9603180` |
| `MACE.py` | `afffd967be348874226796f3686e199f41534a6dda972abd9ba00fdbf450c7de` |
| `calcSCR.py` | `4dc150e7f747dab9a1003a2dccbe46132964a2e6d6578870d410be0d852a1859` |
| `calcmp.py` | `25f42b064593e4a0f0054ea0edaf9b2828490181b6a14f68257ebb0140b133e2` |
| `MACalcutils.py` | `2a54454fc004279bf1ae5bb468d69a5284d96a19d4387b7a616d1fa92ce22026` |

If any hash differs at build time, stop and tell the user before building.
The model has changed since the accepted run.

### Packaged-run acceptance checks

1. **Get a clean input first.** Each Stage 4 run saves MACalc results into the
   audited Stage 3 workbook in place (Decision 61). The single
   `SST Output_<suffix>\Stage 4 backup\` copy is replaced with the pre-run
   workbook on every run (Decision 74). The current backup holds the
   **first Stage 4 result**, not the clean Stage 3 output. For a clean
   comparison, ask the user to rerun Stage 3 (packaged button) before
   clicking **Run MACalc**. Stage 4 then uses the new
   `SST_Stage3_LastSuccessfulOutput` (a new `_<16 hex>` token).
   - Note: the Credit folder will then hold two `MACalc input_*` workbooks.
     `macalc_batch_entry.py` pre-scans with `find_workbooks` and stops
     (exit 2, before MACalc runs or the prompt is answered) unless it finds
     exactly the audited workbook (Decisions 62 and 70). Before the run,
     ask the user to archive the older workbook outside the stress folder.
     Do not delete it.
2. The user clicks **Run MACalc** in the Draft (packaged mode, no CLI).
3. Check:
   - the console exits cleanly;
   - the panel says `Succeeded` with the new `InputRun` and
     `LastSuccessfulOutput`;
   - Stage 1-3 audits are unchanged;
   - no `~$` lock and no orphan Excel process remain.
4. Compare the new `_unoptimised_assets.csv` with the accepted source run:
   109 columns (`Index, ISIN, Name, Zspread, Type, Sector, Rating, Index
   Linked?, PV, 1..100`); 993 rows (634 Bond, 182 Infrastructure,
   177 PP); no blank ISIN or PV. PV sums (units unspecified):
   - Bond 13,135,651,535.34
   - Infrastructure 4,625,370,518.85
   - PP 2,755,166,247.38
   - **Total 20,516,188,301.57**

   With unchanged Stage 1-3 inputs and model hashes, expect these to
   match. If the user changed inputs, ask for a reasonableness review
   instead.
5. Show the results to the user and get acceptance before closing the step.

### Gotchas learned in this session

- **OneDrive caller URL (source mode only):** Excel's `Workbook.FullName` for the
  synced Draft is a SharePoint URL. `book.fullname` (xlwings) gives the local
  path. Source-mode CLI runs need `--caller "<SharePoint URL>"
  --caller-local "<local path>"`. The packaged VBA launcher already passes
  `LocalDraftFullName`, so the packaged button should not hit this. If it
  does, it fails with "The exact calling workbook is no longer open" before
  touching anything.
- **Excel HWND** changes between reads. Read it immediately before any
  source-mode launch.
- **Installer verification:** compare cell anchors, not pixel `Top`/`Width`,
  when two workbooks are open in one Excel instance.
- **MACalc Batch** opens its own hidden Excel (no add-ins,
  `update_links=False`). Importing it needs cwd set to the model folder; the
  frozen entry must do the same with its bundled snapshot.
- **Packaging risks:**
  - PyInstaller may miss pulp's CBC solver binaries or `smithwilson` data.
    Check `pulp.listSolvers(onlyAvailable=True)` in the self-test if feasible.
  - MACalc `requirements.txt` pins xlwings 0.19.5, while py311 has 0.36.14.
    0.36.14 worked in the live run.
- **Run options** are fixed: binary optimisation OFF, optimisation pass OFF, CoD
  floor ON. The downgrade flag (`SST_Credit_DG_standalone` = TRUE) is passed
  but has no effect while optimisation is OFF.

### Current verification after step 13 implementation

- Packaging/frozen-runtime selection: **52 passed**.
- Changed test modules: **888 passed, 85 skipped, 2 baseline failures**.
- Full suite: **1,114 passed, 89 skipped, 2 baseline failures**.
- Source MACalc self-test: `MACALC_SELF_TEST_OK`, no Excel opened.
- Black passed for new files and changed ranges in legacy files;
  `git diff --check` passed.

### Historical test baseline before step 13

- `python -m pytest -q -p no:cacheprovider tests/test_sst_stage_four.py`:
  **66 passed, 1 skipped**. The skip is a Windows/symlink-dependent case.
  The skipped packaged-builder test lives in `tests/test_bonds_lite_input.py`.
- Full suite: **1,092 passed, 90 skipped, 2 known baseline failures**:
  `test_frozen_main_validates_cached_bundle_before_stage_dispatch` and
  `test_model_results_transaction_rollback_failure_retains_recoverable_backup`.
  These predate Stage 4.

### Uncommitted Stage 4 changes

- Untracked (new): `src/sst_stage_four.py`, `src/macalc_batch_entry.py`,
  `tests/test_sst_stage_four.py`, `references/plan-stage-four.md`.
- Modified: this handoff, `src/bonds_lite_input.py` (`--stage 4`,
  `--install-stage-four`, `RunMACalc` in both launcher VBA strings, the
  model-Python resolver), `tests/test_bonds_lite_input.py`,
  `references/bonds-lite-setup.md` (Stage 4 section) and
  `references/decisions.md` (Decisions 60-77).
- Step 13 additionally changes `tools/build_packaged_stages.py`,
  `tools/requirements-model-build.txt`, the MACalc entry's frozen/runtime
  handling, cached-bundle MACalc hash validation and the packaging tests/docs.
- Live artefacts (not for commit):
  - the Draft (Stage 4 controls installed, audit `Succeeded`);
  - its installer backup
    `references\backups\Draft Input spreadsheet for SST_20261104_151920_099479.xlsm`;
  - `SST Output_Test\Credit\MACalc input_Test_4a0b639930e7c09c.xlsb`
    (now contains MACalc results) and its `_unoptimised_assets.csv`;
  - `SST Output_Test\Stage 4 backup\`.

Check `git status` at the start. Some of the files above may already be
tracked, and the pre-existing dirty items must stay untouched.

## Goal

The user's instruction at the end of today's session:

> Stage 4 will be just exploring the ma_calc_batch model and make it run with
> the output handed off by Stage 3.

Start by understanding the consuming model, then connect its existing execution
path to the validated Stage 3 output. Do not redesign its financial calculations
or add unrelated functionality.

**Stage numbering and permission:** older Stage 3 documents and Decision 51 call
the consuming-model run "Stage 5". The user now explicitly calls this work
**Stage 4**. Decision 51 kept the model outside Stage 3's scope; tomorrow's request
authorises exploring it and implementing its Stage 3 handoff as Stage 4. The
model has been inspected read-only; it has not been changed or executed.
Reconcile the older numbering in the plan and decision log during Stage 4
planning; do not treat the Stage 3 restriction as a prohibition on tomorrow's work.

## Read First

- [Stage 3 plan](plan-stage-three.md), version 0.15, particularly Slices C and D
  and their implementation addenda.
- [Decisions](decisions.md), especially 49-54: the list contract, consuming-model
  boundary, CSV rating suffix and unique filename.
- [Stage 3 handoff](session-handoff-2026-11-02-stage-three.md). Its opening status
  is current; later sections describe older snapshots and are not current release
  instructions.
- [Setup guide](bonds-lite-setup.md), including deployment, Stage 3 handoff and
  ISIN list alignment.
- [Stage 1 flags and output plan](plan-stage-one-output-flags.md), Revision 0.2
  and "Commit and release", and Decisions 55-59: the credit stress flags, the
  `SST Output_<suffix>\<stress>` layout, the overwrite rule and the Stage 3
  filename.

## Current State (4 November 2026, 13:45 UTC)

- Stage 3 Slices A-D are implemented and accepted (`0fdf81a`).
- Stage 1 flags/output change, all on `origin/main`:
  - `2bd2f0e docs: Record the v8 release and add the Stage 4 handoff`
  - `84e2dae feat: Add credit stress flag checks and the SST Output folder layout`
  - `019764f docs: Record the 20261104-cache-v9 release`
- **Active central release:** `20261104-cache-v9` (built from `84e2dae`, both
  self-tests passed, same files and modules as v8). Promoted with the user's
  approval. **Rollback:** `20261103-cache-v8`. Do not roll back without approval.
- The user's packaged Stage 1-3 rerun on v9 worked as expected (suffix `Test`).
- Tests for the changed modules (`tests/test_bonds_lite_input.py`,
  `tests/test_sst_stage_three.py`): **1022 passed, 88 skipped, 2 known
  failures**, which predate this work:
  - `test_frozen_main_validates_cached_bundle_before_stage_dispatch`
  - `test_model_results_transaction_rollback_failure_retains_recoverable_backup`
- A packaged Stage 1 MAP import once failed with a `HeaderRowRange` error
  (3 November). It has not recurred; investigate only if it does.

### Stage 1 behaviour that Stage 4 must respect (Decisions 55-59)

- Flags in `I. Control`: `SST_Credit_DG_standalone` (D16) and
  `SST_Credit_spread_standalone` (D17). Interest (D14), Inflation (D15),
  Property Value (D18) and FX (D19) are not named yet; the user will name them
  later and each will get its own stress folder.
- Flag/stress inconsistencies give a warning and a `[y/N]` prompt; both flags
  FALSE stops.
- Output root: `<Draft folder>\SST Output_<SST_output_folder_suffix>\`, with one
  folder per selected stress group. Today only `Credit` exists
  (`STRESS_OUTPUT_FOLDERS` in `src/bonds_lite_input.py`).
- Rerunning Stage 1 with the same suffix **replaces** the stress folder,
  including earlier Stage 2 and Stage 3 outputs (and, later, Stage 4 outputs if
  they live there). A new suffix or archiving keeps earlier output.
- Legacy `Bonds Lite output\<timestamp>` handoffs are rejected.

Central package root:

```text
\\server090\internal model$\29 Financial Resilience\Sandbox\SST spreadsheet executable file
```

## Stage 3 Handoff Contract

The input path must come from the Draft's **`SST_Stage3_LastSuccessfulOutput`**.
Never reconstruct the filename, choose by modification time or assume that
`MACalc input.xlsb` is the current input.

- Stage 3 names each output `MACalc input_<suffix>_<16 hex>.xlsb` (Decision 57).
  Every Stage 3 rerun gets a fresh token and keeps earlier successful outputs in
  the same stress folder, until Stage 1 replaces that folder.
- Relevant audit names are `SST_Stage3_Status`, `SST_Stage3_InputRun` and
  `SST_Stage3_LastSuccessfulOutput`. Determine tomorrow how Stage 4 should handle
  a later unsuccessful Stage 3 attempt that has retained an older successful
  path; do not silently invent a stale-input policy.
- Stage 3 saves and closes the workbook, reopens it read-only for verification,
  and publishes it only after verification passes.
- Both the output path and staged path are limited to 218 characters.
- Staging is under `SST Output_<suffix>\Stage 3 staging\<token>\`, outside the
  stress folder, and is removed afterwards. Do not use staged or incomplete files.

The user's v9 rerun produced (checked on disk at 13:45; the Draft was open in
Excel, so its audit cells were not reread):

```text
SST Output_Test\Credit\MACalc input_Test_4a0b639930e7c09c.xlsb
```

The same folder holds `Bonds Lite input_Bonds.xlsx`, `Bonds Lite input_MAP.xlsx`,
`bonds_batch_summary_*.csv`, `JRL Bonds\` and `JRL MAP\` (Stage 2 CSVs and
input workbooks) and a dot-prefixed `.bonds-lite-*\` temporary workbook folder.
This is a **candidate for the first live test**, not a hard-coded handoff:
resolve `SST_Stage3_LastSuccessfulOutput` at the start of the session.

**Folder-scanning risk (confirmed by discovery):** MACalc Batch scans the folder
it is given *recursively* and runs every workbook whose name contains `MACalc`.
Today only the Stage 3 output matches in the `Credit` folder, but a Stage 3
rerun leaves several `MACalc input_*` files side by side, and all would run. Do
not delete earlier outputs to work around this (see Q4-2).

## MACalc Batch Discovery (read-only, 4 November 2026)

Model folder (confirmed):

```text
C:\Users\Joao Branco\OneDrive - Just Group\Financial Resilience Team - General\00 Ad Hoc\MACalc_Batch
```

- **Entry point:** `MACalcrunmanager.py`. `ma_calc_batch` is only the user's name
  for it. The old `MACalc.py` in `v_3_3` is not the consuming model.
- **Workbook selection (`find_workbooks`):** `rglob` over the folder; filename
  contains `MACalc` (case-insensitive); extension `.xls`, `.xlsb`, `.xlsm` or
  `.xlsx`; skips names starting `~$` and names containing `_opt`. So a Draft
  suffix containing `_opt` would hide the Stage 3 output.
- **Interaction:** asks a `y/N` confirmation through `input()` before running,
  which a launcher must answer or bypass.
- **Execution:** calls `MACalc.main` through xlwings **0.19.5** (Bonds Lite uses
  0.36.14). The Python environment for MACalc Batch is not yet identified.
- **Side effects:** the base run **saves over the input workbook** and exports
  `<stem>_unoptimised_assets.csv` beside it.
- **Default options:** no binary optimisation, no optimisation pass, not a
  downgrade run, cost-of-default floor at zero.
- **Inputs read by `MACalc.py` / `calcSCR.py`:** `tbl_Assets`, `lst_pre16assets`,
  `lst_noPRoC`, `tbl_ForcedCompAperc` and `tbl_MinCompAperc` (left merge on
  ISIN), and `lst_REITs`. They use exact keys, but bare `except` blocks fall back
  to all-False when a name is missing, so a missing list fails silently.

Before the first run, verify the remaining assumptions deferred by Decision 51:

- `tbl_Assets` is read dynamically, including all funded and asset rows.
- The following lists join to `I. Assets` column D by exact key:
  `tbl_ForcedCompAperc`, `tbl_MinCompAperc`, `lst_pre16assets` and `lst_REITs`.
- List lookup, duplicate-key and unmatched-key behaviour agree with the Stage 3
  contract. `lst_noPRoC` is unchanged by Stage 3.
- Rating-suffixed keys use the row's CSV `Final_Rating` (column H), not
  `Final_Notched_Rating` (column I). Stage 3 still writes the notched rating into
  `I. Assets` column I; do not strip `+`/`-` or infer the suffix.

Any mismatch needs a reported finding and an agreed correction, not a silent
change to the accepted Stage 3 output.

## Open Stage 4 Questions (for the user to answer)

Each question gives the context and the options found so far. Answer in chat;
"other" is always allowed.

**Q4-1. Execution surface.** The controller CLI accepts `--stage 1|2|3` only.
- (a) Source-mode proof first: a Python launcher run from the terminal on the
  audited Stage 3 output; button/packaging later. *Recommended*: smallest step,
  proves the model works with Stage 3 output before touching the Draft or VBA.
- (b) Full integration now: `--stage 4`, a Draft button, Stage 4 audit cells and
  a new release.

**Q4-2. Protecting the Stage 3 output.** MACalc Batch saves over its input
workbook and scans its folder recursively for every `MACalc*` workbook.
- (a) Copy the audited Stage 3 file into its own Stage 4 folder (for example
  `SST Output_<suffix>\Credit\Stage 4\`) and point MACalc Batch at that folder,
  so only one workbook is found and the Stage 3 original stays intact.
  *Recommended.*
- (b) Run in place on the audited file in the stress folder. Simpler, but it
  changes the Stage 3 output and would also pick up any earlier `MACalc input_*`
  files in that folder.

**Q4-3. Run options.** MACalc Batch defaults: no binary optimisation, no
optimisation pass, not a downgrade run, CoD floor at zero.
- (a) Use the defaults.
- (b) Other settings. Notably: should `SST_Credit_DG_standalone = TRUE` switch on
  MACalc's downgrade-run option, or is the downgrade already fully carried by the
  Stage 2 results in the MACalc input?

**Q4-4. Proof of correctness.** What reference shows that Stage 4 output is
right? For example: a manual MACalc run on the same input, a previous SST result,
specific SCR/MA figures, or only that it completes and produces the expected
files. Please name the file or figures and an acceptable tolerance.

**Q4-5. Stale handoff.** If the latest Stage 3 attempt failed but
`SST_Stage3_LastSuccessfulOutput` still holds an older successful path:
- (a) Stop and ask the user to rerun Stage 3. *Recommended* (consistent with how
  Stage 3 treats Stage 2).
- (b) Run on the older successful output with a warning.

**Q4-6. Stage 4 output location and overwrite.** Where should MACalc results go,
and may a Stage 4 rerun replace earlier Stage 4 results? Suggested: inside the
stress folder (so a Stage 1 rerun with the same suffix clears them, as Decision
59 does for Stages 2 and 3), with a rerun replacing the previous Stage 4 output.

**Q4-7. `_opt` in the suffix.** MACalc Batch skips files whose names contain
`_opt`. Should Stage 1 reject such a suffix up front, or should Stage 4 simply
stop with a clear message?

Also to confirm during planning: which Python environment runs MACalc Batch
(it needs xlwings 0.19.5), and whether the `y/N` prompt may be answered
automatically by the launcher.

## Remaining Checklist

1. Done: inspected the worktree and releases.
2. Done: explored the MACalc Batch entry point, selection rules and side
   effects (above). Still open: its dependency manifest/environment and any
   accompanying specification or reference results.
3. Next: get the user's answers to Q4-1 to Q4-7 and record them as decisions.
4. Verify the deferred list contract and single-input selection rules above.
5. Write `references/plan-stage-four.md`: the smallest integration plan. Keep
   financial model changes out of scope unless a specific incompatibility is
   found and the user approves its repair.
6. Write failing tests for the agreed handoff and launcher behaviour before
   implementation. Cover the exact audited filename, invalid/missing inputs,
   stale-run policy, extra workbooks/earlier outputs in the same folder, explicit
   model failure reporting and the expected output artefacts.
7. Run targeted tests, then a controlled live run on the audited Stage 3 output.
   Check expected financial results with the user's reference, not merely a zero
   exit code. Confirm that Stage 3's input and earlier outputs remain intact.
8. Only expand into packaging, Draft/VBA changes or release work if agreed.
   Preserve the controller/Bonds Lite model separation if packaging is required.

Use base coding and testing standards; no team-specific skills are active.
Suggested workflow: Supir Explore for bounded model discovery, then
Plan -> Test -> Implement -> Review -> Commit. The planning agent must hand off
production implementation rather than writing source code itself.

## Working Tree And Permissions

All work up to `019764f` is committed and pushed. Leave these uncommitted items
alone (do not revert, delete or stage them):

- the Draft workbook (modified by the user; agents must not edit it);
- tracked deletions of old `Bonds Lite output\2026102*` run folders;
- untracked `Bonds Lite output\20261103_183144\`, `Bonds Lite output\Stage 3
  staging\` and the user's new `SST Output_Test\`;
- modified bytecode, `.pytest_cache`, Draft backups in `references\backups\` and
  the `MACalcinPython_v4.0f.xlsb` / `REM_v1.10.xlsb` templates.

The MACalc Batch folder is read-only for this work unless the user approves a
change. Do not commit, push, promote or roll back without explicit approval.
Use Supir Commit when a commit is approved, with the required Co-authored-by
trailer. Do not change the Draft's controls/names or VBA without discussing the
Stage 4 integration scope.

**Question tool:** `vscode/askQuestions` is not available in the Copilot CLI/SDK
agent host. In that host, ask questions in chat and wait for answers rather than
logging agent-made decisions.

For the existing source test baseline:

```powershell
python -m pytest -q -p no:cacheprovider
```

## Implementation Status (4 November 2026)

- Implemented in source: `src/sst_stage_four.py`, `src/macalc_batch_entry.py`,
  Stage 4 CLI dispatch and launcher text in `src/bonds_lite_input.py`, and the
  Stage 4 setup-guide section.
- Automated Stage 4 tests: 66 passed, 1 skipped. Full suite: 1,092 passed,
  90 skipped, with the two baseline failures listed above.
- Step 13's package-builder test is now enabled. The builder, frozen entry,
  cached executable integrity check and automated tests are implemented;
  the release itself has not been built.
- Stage 4 controls are installed in the Draft; the panel status is `Succeeded`.
  The installer backup is `references\backups\Draft Input spreadsheet for
  SST_20261104_151920_099479.xlsm`. The live output is
  `SST Output_Test\Credit\MACalc input_Test_4a0b639930e7c09c_unoptimised_assets.csv`;
  PV sums to 20,516,188,301.57 in the CSV's unspecified units (Infrastructure:
  4,625,370,518.85; Bond: 13,135,651,535.34; PP: 2,755,166,247.38).
- Step 12 is complete and **accepted by the user** (4 November 2026, about
  15:43 UTC). The active packaged release remains `20261104-cache-v9`, which
  does not accept `--stage 4`. Do not use the installed button until a
  Stage 4-capable release is promoted.
- **Next:** ask whether to commit the Stage 4 source/tests/docs without
  pushing, then obtain build approval. Promotion and the packaged run need
  separate explicit approval. Follow "Next Session: Promote Stage 4" above.
