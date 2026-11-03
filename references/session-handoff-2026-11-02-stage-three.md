# Session Handoff: Stage 3 MACalc Input (2 November 2026, updated 3 November 2026)

Read this file together with `references/plan-stage-three.md` (v0.12), which is
the authoritative Stage 3 plan. Decisions are in `references/decisions.md`
(Decisions 32-54; 54 is proposed). The earlier handoff `references/session-handoff-2026-11-02.md`
covers Stages 1 and 2.

## Next Session: Implement Slice D (Unique MACalc Input Filename)

The user's request (3 November 2026): `MACalc_Batch` (Stage 4) needs the MACalc
input workbook to be uniquely named across synced SharePoint folders, so
Stage 3 should add a unique suffix to the name. The plan is in
`references/plan-stage-three.md`, section "Slice D: Unique MACalc Input
Filename" (steps D0-D6). The proposal (Decision 54, pending) is
`MACalc input_<Stage 1 run ID>_<token>.xlsb`, for example
`MACalc input_20261103_161318_3f9c2a7be104d65e.xlsb`, with a fresh
16-character random token on every run. The run ID alone is not enough: it is
unique only within one output root, so two Drafts starting Stage 1 in the same
second, or a copied run folder, would clash. Because the name changes on every
run, earlier outputs are kept until the new one is verified and audited, then
deleted, and Stage 4 must read the path from `SST_Stage3_LastSuccessfulOutput`.

Start the session by:
1. Committing the pending v6 build fix and release notes (D0; see Outstanding).
2. Asking the user Q-D1 to Q-D4 from the plan:
   - Q-D1: any naming rule in `MACalc_Batch` (spaces, pattern, length), and is
     a 16-character token acceptable or is a full UUID preferred?
   - Q-D2: delete earlier tokenised outputs after a successful rerun?
   - Q-D3: does `MACalc_Batch` ignore the dot-prefixed staged file?
   - Q-D4: leave legacy `MACalc input.xlsb` files in place?

   Then record Decision 54 as approved or changed.
3. Following D1-D6 with tests first. Only `src/sst_stage_three.py` and its
   tests change. Then build and promote v7 with approval, have the user test
   Stages 1-3, and commit.

Do not inspect, change or run `MACalc_Batch` (Decision 51).

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

1. **Commit the v6 build fix** before starting Slice D:
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

Use `python -u` to see the Stage 3 summary as it is printed. If the output
already exists, the run shows a system-modal overwrite prompt; move the old
output aside first when running unattended.

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