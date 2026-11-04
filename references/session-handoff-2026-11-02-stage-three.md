# Session Handoff: Stage 3 MACalc Input (2 November 2026, updated 3 November 2026)

Read this file together with `references/plan-stage-three.md` (v0.13), which is
the authoritative Stage 3 plan. Decisions are in `references/decisions.md`
(Decisions 32-54). The earlier handoff `references/session-handoff-2026-11-02.md`
covers Stages 1 and 2.

## Next Session: Stage 4 (4 November 2026)

Use the [Stage 4 handoff](session-handoff-2026-11-04-stage-four.md).
The user now asks to explore MACalc Batch and run it with Stage 3's handed-off
output as **Stage 4**. Older references call this Stage 5; their no-model-work
restriction applies to Stage 3, not to this newly requested Stage 4 work.

## Slice D Status (3 November 2026, 18:54 UTC): v8 promoted

Slice D is committed as `0fdf81a` (`feat: Give each MACalc input a unique
filename`). The 218-character output/staging path limit is included in
release `20261103-cache-v8`. The user approved promotion; the central pointer
now targets v8. The user's LocalAppData cache is expected to update when Excel
next checks the pointer. Decision 54 is approved with changes.

What Stage 3 now does:
- **Output name:** `MACalc input_<run ID>_<16 hex>.xlsb`, with a fresh token on
  every run. Stage 4 must read the path from `SST_Stage3_LastSuccessfulOutput`.
- **Reruns (Q-D2):** earlier outputs are kept. There is no overwrite prompt,
  backup or deletion.
- **Staging (Q-D3):** the workbook is staged in
  `<output root>\Stage 3 staging\<token>\`, never in the run folder. The folder
  is removed afterwards, including when OneDrive has marked it read-only.
  Staging in local `%TEMP%` does not work, because Excel rewrites the
  template's SharePoint links.
- **Path limit:** the run stops if the output or staged path exceeds 218
  characters (Excel's limit). The staged path is 204 characters on this
  Draft. This limit is in the active v8 release.
- **Legacy files (Q-D4):** `MACalc input.xlsb` files are ignored.

Verification so far:
- The full suite gives 885 passed, 85 skipped and the 2 known failures.
- All 4 opt-in Excel tests pass. The three Stage 3 write tests need
  `--basetemp` inside the synced folder; the installer test needs the default
  `%TEMP%`.
- Three live source-mode runs on `20261103_161318` each printed the accepted
  Slice C summary.
- v8 build: both self-tests passed. Compared with v6, the controller has the
  same 82 files and 469 frozen modules, and the model has the same 869 files
  and 1,770 modules; neither bundle has added or removed files or modules.
  The central pointer targets `20261103-cache-v8`. Rollback is to v7.

Status and remaining steps:
1. Done: the user confirmed that `MACalc_Batch` scans only the folder holding
   its input (the run folder), so the sibling staging folder is not seen.
2. Done: the user approved promotion and the central pointer now targets
   `20261103-cache-v7`. The user's LocalAppData cache is expected to update on
   the next button click. Rollback is to `20261103-cache-v6`.
3. Done: the user's packaged Stage 1-3 test on v7 passed. One Stage 1
   attempt first failed once with `'NoneType' object has no attribute
   'HeaderRowRange'` in the MAP import. This could not be reproduced (six
   repeat MAP imports succeeded) and the retry passed; it is unrelated to
   Slice D.
4. Done: with approval, committed `0fdf81a feat: Give each MACalc input a
   unique filename`, including the 218-character path limit.
5. Done: built and verified `20261103-cache-v8` without `--promote`; the user
   then approved promotion. The central pointer targets v8; rollback is to
   `20261103-cache-v7`.

During Stage 3, do not inspect, change or run `MACalc_Batch` (Decision 51).
Tomorrow's separately requested Stage 4 work is covered by the handoff above.

## Current Status (3 November 2026, 16:45)

- **Slice A:** implemented, run on live data and accepted by the user (Decision 40).
- **Slice B:** implemented (steps B1-B7), run live in source mode on handoff `20261102_111001` and **accepted by the user on 3 November 2026**. The run took 112 s and produced 5 funded rows (21-25) and 988 asset rows (26-1013), with A:C filled from MACE `Control` and nothing below row 1013. Decision 47 (implementation departures, including the standard-library CSV reader without pandas) and Decision 48 (the K:O side check covers column P only, because MACE has other panels from Q onwards) are approved. Tests: 797 passed, 84 skipped and the 2 known failures.
- **Slice C implemented on 3 November** (steps C0-C7, plan v0.11, "Addendum: Slice C Implementation Progress"; Decisions 49-53 approved; **accepted by the user on 3 November 2026**). In the same single run, Stage 3 now realigns `tbl_ForcedCompAperc`, `tbl_MinCompAperc`, `lst_pre16assets` and `lst_REITs` with the rating-suffixed column D, in the same staged book, with one save, one read-back and one audit. `I. Assets` and `lst_noPRoC` are untouched. Decision 53 (a clarification, approved by the user): the row's "own CSV rating" in the approved exact rule is the CSV `Final_Rating` (column H), which is the `Unique_ISIN` suffix on all 988 rows of 111001; `I. Assets` column I holds the notched rating and is not used for the suffix. Guard 7 accepts D = E or D = E + "_" + `Final_Rating` exactly, and nothing else. **Live source-mode run on `20261102_111001` succeeded on 3 November (109 s)** with the expected plan C6 summary (Forced 19/5/0/0, Min 0/106/58/0, Pre16 0/31/291/0, REITs 0/16/0/0 kept/remapped/unmatched/added; no duplicates or fan-outs). A one-line logger fix in `bonds_lite_input.py` was needed so that the summary shows in source mode (plan addendum deviation 5). The user spot-checked the output and accepted Slice C. All 4 opt-in Excel tests pass.
- **Committed and released.** Slices A-C were committed in `0f6ef09` (after
  `85cbe07`, the Stage 1/2 handoff). Release `20261103-cache-v6` is **live**
  (central `current.txt`); the user's local cache is also on v6.
  - The user ran Stages 1-3 with the packaged release on handoff
    `20261103_161318` and reported that everything works.
  - `20261103-cache-v5` is on the share but was never promoted, because its
    model bundled the Stage 3 modules by mistake.
  - The fix makes the hidden import controller-only. It is in
    `tools/build_packaged_stages.py`, with the regression test
    `test_pyinstaller_command_hidden_imports_stage_three_in_controller_only`
    in `tests/test_bonds_lite_input.py`, and is **not yet committed**.
  - Details are in the plan's "Commit and release" note.
- **Slice D planned** (unique MACalc input filename for Stage 4); see "Next Session" above.

Key points from Slice B (full detail in the plan):

- Slices A and B are one process: one "Prepare MACalc" run, one output, one read-back and one audit result. No new button, panel or VBA changes.
- **Source CSV:** the file containing `py_bonds_Base_Scenario_nozeros` (held as the module constant `MAP_BASE_OUTPUT_PATTERN`). Look only in the entity's MAP output folder from Stage 2 (today `JRL MAP\`). Stop if zero or more than one file matches.
- **CSV to MACalc mapping:** B→D (`Unique_ISIN`), A→E (ISIN), D×10000→F, F→H, I→I, E→K, J:ATN→M:ATQ.
- **From MACE `I. Bonds`:** G, J and L, keyed on column D (ISIN).
- **From MACE `Control` K:O:** A, B and C (L→A, M→B, O→C), keyed on column K, for every post-ILS row, **including the funded rows**.
- **Stop when:**
  - a CSV asset is missing from MACE, or there are duplicates, blanks or errors (list the ISINs);
  - the table is not contained in K:O (exact row-9 headers, column P empty, no second block within 50 rows; Decision 48);
  - the valuation dates differ;
  - column A is blank anywhere in rows 2..last;
  - a Type is outside the MACalc set.
- **Warn only (do not stop):**
  - the CSV asset set differs from the Stage 1 MAP set;
  - MACE has extra rows.
- Do not compare PV.
- **Expected live result on `20261102_111001`:** 5 funded rows (21-25) and 988 asset rows (26-1013), A:C populated from row 21 down, nothing below row 1013.
- **MACalc reference:** `MACalc.py` in `\\server090\internal model$\7 ALM\2. Model Management\Model Repository\Release Preparation\MACalc\v_3_3` reads `tbl_Assets` dynamically (line 431).

## User Goal

Stage 3 prepares a MACalc input workbook from the MACalc template, the Stage 1
MAP input and the central MACE workbook. Delivery is in slices. **Slice A**
(template checks and funded-leg transfer) is implemented, run on live data and
**accepted by the user on 3 November 2026** (Decision 40). The next session
will **plan** the remaining Stage 3 slices, which the user describes as
"similar data transfers between spreadsheets". The user will supply their
requirements at the start of that session.

## Slice A Requirements (confirmed by the user)

- Template path: named cell `SST_MACalc_Input_Template` (I. Control!D44, value
  `MACalcinPython_v4.0f.xlsb`). MACE path: `SST_MACE_Central_Template`
  (P. Parameter!C8). `SST_MACalc_Central_Template` is P. Parameter!C9.
- Template `I. Assets` tab: a non-empty Mortgage block at the top, followed by a
  non-empty ILS block (identified by the Type column). Otherwise show a message
  asking the user to update the template for LTM/Swap assets and stop.
- Funded assets: those whose MACE funding-leg cashflow vector (MACE
  `C. Assets - Funding`, AJ:AUM) is all negative (zeros allowed in the tail).
  MAP cashflows are **not** copied.
- Copy from MAP `I. Bonds Data` to template `I. Assets` for each funded asset:
  C->D (ISIN), C->E (Name), E->K (Currency), F->H (Sector), H->I (Rating),
  J->J (Index Linked?, as values).
- Copy from MACE `C. Assets - Funding`: C->F (Zspread), D->G (Type),
  I->L (FSAO), AI x -1 -> M (PV), funding-leg cashflows -> N:ATQ.
- Columns A:C are left empty (a later slice fills them from another MACE tab).
- Rows below the funded block are cleared.
- Zero funded assets: warn and continue (not an error).
- Anomalous MACE vectors or MAP/MACE set differences: stop and report.
- Output `MACalc input.xlsb` is saved in the timestamped Stage 1 run folder
  (Slice D will rename it to `MACalc input_<run ID>_<token>.xlsb`).
- Stage 3 lives in its own module, `src/sst_stage_three.py`, and requires a
  validated Stage 2 run on the **same** Stage 1 handoff.

## Implemented (committed in `0f6ef09`)

- `src/sst_stage_three.py` (new): validation, reads, reconciliation, writer
  with read-back verification, external-link restoration, audit, guard,
  rollback, `run_stage_three`, panel/button installer
  (`install_stage_three_controls`, `_stage_three_button_row`).
- `src/bonds_lite_input.py`: `--stage 3` lazy dispatch; `--install-stage-three
  <Draft>` CLI; `Public Sub PrepareMACalc()` (calls `LaunchBondsLite 3`) in
  both `VBA_LAUNCHER` and `PACKAGED_VBA_LAUNCHER`; shared-helper aliases.
- `tests/test_sst_stage_three.py` (new, 60 tests incl. opt-in Excel tests via
  `BONDS_LITE_RUN_STAGE_THREE_EXCEL_TESTS=1`); `tests/test_bonds_lite_input.py`
  invalid-stage test updated.
- `tools/build_packaged_stages.py`: `--hidden-import sst_stage_three`.
- Docs: `references/plan-stage-three.md` v0.6, `references/decisions.md`
  (32-40), `references/bonds-lite-setup.md` (new Stage 3 section).
- **Live Draft installed** (`Draft Input spreadsheet for SST.xlsm`):
  `SST_Stage3_Panel` = `'I. Control'!$H$100:$O$110` (blue, fields L101-L109,
  Status "Succeeded" for the current handoff); button `btnPrepareMACalc` at
  D100 -> `PrepareMACalc`.
  Backup: `references/backups/Draft Input spreadsheet for SST_20261102_171859_041698.xlsm`.
  The Draft is also modified by the user's own D44 addition.

## Verification

- Slice C (3 November, after the exact `Final_Rating` rule and the logger fix):
  `python -m pytest -q` gives 865 passed, 85 skipped and the same 2
  pre-existing failures. The opt-in Slice C Excel fan-out test and the 111001
  real-data regression results are in the plan's Slice C addendum.
- Slice C live run (3 November, 13:27-13:29): Draft audit `Succeeded`, output
  `Bonds Lite output/20261102_111001/MACalc input.xlsb`, read-back verified,
  console summary equal to plan C6. Before the run, the accepted Slice B
  output was moved (not deleted) to the agent session folder
  (`files\slice-b-output-backup\`), together with the two diagnostic run
  outputs, so that no overwrite prompt was needed.
- An orphan staged file from an earlier run,
  `Bonds Lite output/20261102_111001/.MACalc-stage3-kdnte13n.xlsb`
  (3 November 11:54), was left untouched; the user may delete it.

- `python -m pytest -q`: **721 passed, 83 skipped, 2 failed**. The 2 failures
  are pre-existing and unrelated:
  `test_frozen_main_validates_cached_bundle_before_stage_dispatch` and
  `test_model_results_transaction_rollback_failure_retains_recoverable_backup`.
- Disposable real-data round-trip (091803 MAP + central MACE + template copy):
  5 funded ISINs, exact mappings, template hash unchanged, 9 external links
  preserved.
- Live Draft install compared with its backup: no existing names, Stage 1/2
  buttons, D44 or Stage 1/2 audit values changed; all prior VBA retained.
- Stage 2 reran successfully on the current Stage 1 handoff `20261102_111001`
  on 3 November 2026. Summary:
  `Bonds Lite output/20261102_111001/bonds_batch_summary_20261103_093358.csv`.
- Stage 3 ran successfully against that handoff on 3 November 2026. The saved
  workbook passed the runner's read-back verification, and the Draft audit
  records `Succeeded` with output
  `Bonds Lite output/20261102_111001/MACalc input.xlsb` (about 26.8 MB,
  86 s end to end). The user compared the five funded rows and accepted them.

## Outstanding

1. **Commit the v6 build fix** before starting Slice D (done: `ce799db`, with the notes in `d4e94c9`):
   - the changes to `tools/build_packaged_stages.py` and
     `tests/test_bonds_lite_input.py`;
   - the release notes in `references/plan-stage-three.md`,
     `references/bonds-lite-setup.md` (deployment status) and this handoff.

   Use Supir Commit, with a message such as
   `fix: Hidden-import Stage 3 in the controller build only` and the
   Co-authored-by trailer. The full suite gives 866 passed, 85 skipped and the
   2 expected failures:
   - `test_frozen_main_validates_cached_bundle_before_stage_dispatch`;
   - `test_model_results_transaction_rollback_failure_retains_recoverable_backup`.

   Do not commit `__pycache__`, `.pytest_cache`, `Bonds Lite output/`, the
   Draft backups or the `.xlsb` templates.
2. **Slice D**: see "Next Session" above.
3. **Optional repository hygiene** (proposed Commit 3; not yet approved):
   - stop tracking the two `.pyc` files;
   - ignore `.pytest_cache/` and `Bonds Lite output/`;
   - decide what to do with the templates and backups.

## Running Stage 3 In Source Mode (maintainers; the packaged release v6 includes Stage 3)

The 3 November run used the source controller against the open Draft. Excel
had the Draft open by its SharePoint URL, so `--caller` must be that URL and
`--caller-local` the synced local path (as the packaged launcher does). Check
first that `SST_Stage2_Status` is `Succeeded`, `SST_Stage2_InputRun` matches
the current Stage 1 handoff, and the Draft is saved. Then run from `src`:

```powershell
python -m bonds_lite_input --stage 3 --caller "<Draft FullName from Excel>" `
    --caller-local "<local synced Draft path>" --excel-hwnd <Excel Hwnd>
```

Use `python -u` to see the Stage 3 summary as it is printed. Since Slice D,
each run writes a new tokenised file and never prompts, so it can run
unattended. Older v6 builds still prompt before overwriting `MACalc input.xlsb`.

Look up `FullName` and `Hwnd` through `xlwings.apps`. Without
`--caller-local`, a URL-opened Draft is rejected with "Save and open Draft
from its local synchronised folder".

## Working Notes

- Run tests with `python -m pytest -q` (bare `pytest` fails to import `tools`).
- Close the Draft in Excel before any installer; installers back up first.
- Excel teardown may print `0x800706ba`/`0x800706be`/`0x80010108` after
  successful runs; these are harmless.
- PowerShell 5: no `&&`; pipe here-strings to `python -`.
- Use British English. Preserve the user's unrelated dirty/untracked files.