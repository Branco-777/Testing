# Plan: Property Stage 1 - Prepare LTM Flash

Version 0.2 (approved) | 6 November 2026 | All questions answered; Decisions 87-95 approved by the user (Decision 92 changed: the Property Value flag gates the workstream).

## Overview

Property Stage 1 is the first stage of the Property stress workstream. It copies the LTM Flash template named in the Draft (`SST_LTMFlash_input_template`, `I. Control!D46`, "already set up"), replaces the contents of its `I. ScenFlipper` sheet with the Draft's `I. ScenFlipper`, sets its valuation date (`ValDate`) to the Draft's `I_Valuation_Date`, and saves the result in `SST Output_<suffix>\Property\`, beside `Credit`. It is started from a new button on `I. Control`, placed to the right of the Credit Stage 1 audit panel, with its own named audit panel. All Property code lives in a new module, `src/sst_property_stage_one.py`; it reuses the generic Excel I/O, folder, guard and audit helpers in `bonds_lite_input.py`.

## Verified Discovery (Read-Only, 6 November 2026)

Read from temporary copies of the Draft and `LTM_FlashModel_v3.2a.xlsb`; the originals were not changed.

### Draft (`Draft Input spreadsheet for SST.xlsm`)
- `SST_LTMFlash_input_template` -> `I. Control!D46` = `LTM_FlashModel_v3.2a.xlsb` (relative to the Draft). `SST_LTMNote_input_template` -> D47 (blank; a later stage).
- `I_Valuation_Date` -> `I. Control!D11` = 31/12/2025 (datetime). `SST_output_folder_suffix` -> D63 = `Test`.
- `D18` "Property Value" (label in C18, row id `3f` in B18) = TRUE in the current Draft; it has **no** workbook name. The Credit flags are named `SST_Credit_DG_standalone` -> D16 and `SST_Credit_spread_standalone` -> D17.
- `I. ScenFlipper`: used range `A1:AG21`, **values only** (no formulae). B1 `detailed` (list validation `detailed,GrossandNNEG,total`), B2 = 16 scenarios, B3 = 16 (current scenario), headers in row 5 (`Scenario` ... `Combined_Vol`), scenarios in rows 6-21. Draft names `prm_ScenFlipperOutputType`, `prm_NoOfScens`, `prm_CurrentScen` point at B1:B3 of the Draft's own sheet.
- Layout as moved by the user: Credit buttons occupy columns J:N; panels occupy O:V (`SST_Stage1_Panel` = `O9:V19`, Stage 2 `O22:V32`, Stage 3 `O35:V45`, Stage 4 `O48:V58`). `btnGenerateBondsLiteInput` spans J13:N15. Columns W onwards are empty (no values, fills or merges) in rows 1-80.

### LTM Flash template (`LTM_FlashModel_v3.2a.xlsb`, 4.4 MB, has a VBA project)
- `I. ScenFlipper` exists; used range `A1:AG27`; same row-5 headers as the Draft; column A uses formulae (`=A8+1`...); B3 = 1. Rows 14-21 property stresses are -0.1 here versus -0.1275/-0.1059 in the Draft.
- Names `prm_CurrentScen`/`prm_NoOfScens`/`prm_ScenFlipperOutputType` -> `'I. ScenFlipper'!B3/B2/B1`.
- **Other sheets reference the sheet directly**, e.g. `I. Inputs!C1 = OFFSET('I. ScenFlipper'!$B$5,prm_CurrentScen,0)`. Deleting the sheet and copying the Draft's in would turn these into `#REF!`. The replacement must therefore overwrite the existing sheet's cells, not swap the sheet.
- `ValDate` -> `I. Control!D12` = 31/12/2025, a typed constant (format `dd/mm/yyyy`). `P. Curves` uses `JRC_SpotAnnual(...ValDate...)` (the XLL add-in).
- Ten external links to `\\server090\internal model$\...`; calculation is automatic in the saved file.

### Code (Supir Explore brief)
- Reusable: `resolve_caller_path`, `_control_cell`, `validate_output_suffix`, `allocate_stress_folder` / `discard_replaced_stress_folder` / `remove_incomplete_stress_folder` (Decision 59 safe-replace), `workbook_run_guard`, `owned_book`, `excel_state`, `_named_audit_region`, `_install_stage_button`, Stage 3/4 `_portable_file_path` and installer pattern.
- Credit-specific (not reused): `STRESS_OUTPUT_FOLDERS` flag registry, `read_stress_flags`, `portable_output_path`, `validate_inputs`, `synchronise_controls`, `resolve_stage_input`.
- `allocate_stress_folder` rejects names not in `STRESS_OUTPUT_FOLDERS`; adding `Property` there would make Credit Stage 1 require Property flag names.
- VBA `LaunchBondsLite(ByVal stage As Long)` passes `--stage <n>`; CLI choices are `1`-`4`. No existing installer places a panel to the right of another.
- Packaging needs `--hidden-import` for each lazily imported stage module (`tools/build_packaged_stages.py`), and `tools/controller_entry.py` mirrors the CLI choices.

## Answered Questions (6 November 2026)

| Ref | Question | Answer |
|---|---|---|
| QP-1 | Valuation date: **set** LTM `ValDate` to the Draft date, or **check** and stop if different? | Set it, log old -> new, read back (Decision 89) |
| QP-2 | Replace by overwriting the existing sheet's cells (keeps `I. Inputs` formulae and `prm_*` names) rather than deleting/copying the tab? | Overwrite in place (Decision 88) |
| QP-3 | Draft column A values replace the template's `=A8+1` formulae, and B3 (current scenario 16 vs 1) is copied too. Is that intended? | Yes, copy the Draft sheet exactly |
| QP-4 | Recalculate LTM Flash in Stage 1? A recalculation needs the JRC add-in and `server090` links. | No: manual calculation, no recalculation, links not updated; the LTM run stage recalculates (Decision 90) |
| QP-5 | Output filename in `Property\`? | `<template stem>_<suffix>.xlsb`, e.g. `LTM_FlashModel_v3.2a_Test.xlsb` (Decision 91) |
| QP-6 | Should `D18` "Property Value" gate the stage? | **Yes** (user changed the proposal): name D18 `SST_Property_standalone`; a FALSE flag stops the Property workstream, as the Credit flags do for Credit (Decision 92) |
| QP-7 | Rerun with the same suffix overwrites `Property\` (Decision 59 behaviour, independent of `Credit`)? | Yes |
| QP-8 | Stage identity, button, panel and CLI names | As proposed (Decision 93) |
| QP-9 | Share the Draft run guard with Credit stages (only one stage of any stress runs at a time)? | Yes, same Draft mutex (Decision 94) |
| QP-10 | Packaged release containing Property Stage 1: build and promote only after the source-mode live run passes, with separate approval? | Yes, as for Stages 3-4 |

## Design

### Property flag gate (Decision 92)
- Constant `PROPERTY_FLAG = "SST_Property_standalone"` in `sst_property_stage_one.py` (Property stage modules share it). Not added to `STRESS_OUTPUT_FOLDERS`, so Credit Stage 1's "at least one Credit flag TRUE" check and Credit Stage 4's downgrade read are unchanged.
- `read_property_flag(caller) -> bool`: reads the name via `_control_cell`; a missing name, or a value that is not a real Excel TRUE/FALSE, raises `BondsLiteError` (same wording as `read_stress_flags`).
- `require_property_selected(flag)`: FALSE raises `BondsLiteError("Property stress is not selected: set SST_Property_standalone to TRUE")`.
- The installer names D18: if `SST_Property_standalone` already exists, keep it (it may have been moved); otherwise require `I. Control!C18` = `Property Value` (stripped) and add `='I. Control'!$D$18`. A different label stops the installer without changes.

### Stage identity (Decision 93)
- Module `src/sst_property_stage_one.py`; public `run_property_stage_one(caller) -> Path | None` and `install_property_stage_one_controls(caller) -> Path`.
- CLI token `P1`: `--stage P1` and `--install-property-stage-one`. VBA: `Public Sub PrepareLTMFlash()` -> `LaunchBondsLite "P1"`; `LaunchBondsLite` changes from `ByVal stage As Long` to `ByVal stage As String` (existing `LaunchBondsLite 4` calls still work through VBA coercion). Later Property stages use `P2`, `P3`...
- Button `btnPrepareLTMFlash`, caption `Prepare LTM Flash`. Panel `SST_PropStage1_Panel`, heading `Property Stage 1 - Prepare LTM Flash`.
- Audit fields `SST_PropStage1_{LastRun, Start, Finish, Elapsed, User, Status, InputTemplate, LastSuccessfulOutput}`. `InputTemplate` = portable path of the template used; `LastSuccessfulOutput` = portable path of the saved LTM Flash (the handoff for Property Stage 2).

### Placement (Decision 95)
- Read `SST_Stage1_Panel` (currently O9:V19) and the `btnGenerateBondsLiteInput` shape at install time.
- Leave one empty gap column after the panel, then mirror the Credit arrangement: button (same width as `btnGenerateBondsLiteInput`, same vertical offset relative to its panel) followed by the panel (same 11 x 8 size, same top row). With today's layout: button X13:AB15, panel AC9:AJ19.
- Same occupancy (`CountA`), merge and shape-overlap checks as Stage 3/4; stop with a clear message if the area is not free. Placement is install-time only; runtime finds everything by name, so the user can move it afterwards.
- Property Stage 2 will later go below Property Stage 1 (button + panel), using the Stage 3 "below" rule relative to `SST_PropStage1_Panel`.

### Run sequence (`run_property_stage_one`)
1. `resolve_caller_path`; take `workbook_run_guard` on the Draft (Decision 94); `require_property_stage_one_panel`; read the previous `LastSuccessfulOutput`.
2. Read and validate, in this order: the Property flag (gate above); `SST_LTMFlash_input_template` (absolute or Draft-relative via the Stage 3 `_resolve_control_path` rules; must exist; extension `.xlsb`/`.xlsm`/`.xlsx`); `I_Valuation_Date` (Excel date only); `SST_output_folder_suffix` (`validate_output_suffix`).
3. Validate the Draft `I. ScenFlipper` (Decision 88 guards, below).
   - On any failure in steps 2-3, including a FALSE flag, do as Credit Stage 1 does: write `Failed` (or `Interrupted`) with Start/Finish/Elapsed/User, keep the previous `LastSuccessfulOutput`, save the Draft, re-raise. No folder is created or replaced.
4. Write `Running` to the panel (previous `LastSuccessfulOutput` kept); save the Draft.
5. `allocate_stress_folder(..., "Property")` (safe-replace, Decision 59).
6. Copy the template into a private staging folder inside `Property`; open it in the caller's Excel with `owned_book` (`update_links=False`) under `excel_state` (events, alerts, screen updating and `CalculateBeforeSave` off; calculation manual).
7. Template guards: `I. ScenFlipper` and `ValDate` exist; row-5 headers equal the Draft's exactly.
8. Replace `I. ScenFlipper` in place: `Clear` the union of the template's and Draft's used ranges, then `Draft UsedRange.Copy(Destination=same address)` and copy column widths. Record `LinkSources` before and after.
9. Set `ValDate` to the Draft date (Decision 89).
10. Verify in the open copy: every cell formula/value in the union range equals the Draft's (exact); `prm_*` names unchanged; `I. Inputs!C1` formula still references `'I. ScenFlipper'` and none of the formulas that reference the sheet contain `#REF!`; `LinkSources` unchanged (no link back to the Draft); `ValDate` reads back equal.
11. Save; check `Saved`; close; rename to the final filename (Decision 91); `discard_replaced_stress_folder`.
12. Write `Succeeded` with `LastSuccessfulOutput`; on any error/interrupt: `remove_incomplete_stress_folder` (restores the previous `Property`), write `Failed`/`Interrupted`, keep the previous output.

### ScenFlipper guards (Decision 88)
- The sheet exists in both workbooks; the Draft used range starts at A1.
- Row-5 headers: non-blank, unique, and equal between the Draft and the template (the template's formulae use fixed column offsets).
- `B2` is a positive whole number equal to the number of scenario rows (contiguous non-blank A from row 6); `B3` is a whole number in `1..B2`; `B1` is one of `detailed`, `GrossandNNEG`, `total`.
- No error values; no formula that refers to another sheet or workbook (`!` or `[`), so the copy cannot create a link to the Draft.

### Shared-code changes (kept small)
- `bonds_lite_input.py`: `allocate_stress_folder` (and the matching folder-name check near line 1596) accepts `STRESS_OUTPUT_FOLDERS` keys **or** a new `PROPERTY_OUTPUT_FOLDER = "Property"`; the Property flag is **not** added to the Credit registry (Decision 92); `--stage` choices gain `P1` with a lazy-import dispatch; `--install-property-stage-one`; `PrepareLTMFlash` and the `String` signature in `VBA_LAUNCHER` and `PACKAGED_VBA_LAUNCHER`; the Decision 86 banner reads `Starting Property Stage 1...`.
- `tools/build_packaged_stages.py`: `--hidden-import sst_property_stage_one`. `tools/controller_entry.py`: mirror the new choice and flag.

## Steps

1. **Approval.** Done 6 November 2026: QP-1 to QP-10 answered; Decisions 87-95 approved (Decision 92 revised to gate on the Property flag).
2. **Tests first (Supir Test)** (depends on 1). Create `tests/test_sst_property_stage_one.py` (see Test Strategy) and extend `tests/test_bonds_lite_input.py` for the `P1` dispatch, install flag, VBA text (String signature, `PrepareLTMFlash`) and `allocate_stress_folder("Property")`. Pass: new tests fail only because code is missing; existing tests unchanged. [✓ Implemented — initial tests failed because the module was missing; see addendum]
3. **Flag gate, control reading and ScenFlipper validation** (depends on 2): `read_property_flag`, `require_property_selected`, `read_property_stage_one_inputs(caller)`, `validate_scenflipper(values, formulas)`. Pass: flag, input and guard tests pass. [✓ Implemented — see addendum]
4. **Folder allocation** (parallel with 3): the `allocate_stress_folder` change. Pass: Property allocate/replace/restore tests pass; Credit tests unchanged. [✓ Implemented — see addendum]
5. **Template build** (depends on 3-4): staging copy, `replace_scenflipper(source_sheet, target_sheet)`, `set_valuation_date(book, date)`, `verify_ltm_flash(book, expected)`, save/rename. Pass: fake-book tests for clear range, copy destination, verification failures and clean-up pass. [✓ Implemented — temporary-copy Excel acceptance passed; live P1 output reopened successfully]
6. **Audit and run** (depends on 5): `SST_PropStage1_*` constants, `require_property_stage_one_panel`, `write_property_stage_one_audit`, `run_property_stage_one`. Pass: audit sequence tests pass (Running -> Succeeded/Failed/Interrupted; previous output kept on failure; Draft saved after each write). [✓ Implemented — failure audit path verified; see addendum]
7. **Installer** (parallel with 6): `install_property_stage_one_controls` with the `SST_Property_standalone` naming (Decision 92), right-of-panel placement, VBA insertion (adds `PrepareLTMFlash`, upgrades `LaunchBondsLite` to `String` only if it is the `Long` version), Draft-closed and backup rules. Pass: placement tests (today's layout gives X13:AB15 and AC9:AJ19; an occupied area stops; reinstall is idempotent); flag-name tests (adds D18 when C18 = `Property Value`; keeps an existing name; wrong label stops with no changes). [✓ Implemented — live installer and runtime acceptance passed]
8. **CLI, launcher and packaging** (depends on 6-7). Pass: dispatch, launcher and builder tests pass; full `pytest` green. [Implemented — 1,213 passed and 90 skipped; two unrelated existing failures remain, see addendum]
9. **Opt-in Excel test** (depends on 8): on temp copies of the real Draft and `LTM_FlashModel_v3.2a.xlsb`, run the build and check: ScenFlipper equals the Draft's, `I. Inputs!C1` has no `#REF!`, `prm_*` names intact, `ValDate` = Draft date, `LinkSources` unchanged, file reopens. Pass: test passes; originals untouched (hashes). [✓ Implemented — passed on temporary copies; original hashes unchanged]
10. **Live install and source-mode run** (depends on 9, user approval): back up the Draft, install the controls, click the button in source mode with suffix `Test`; check `SST Output_Test\Property\` beside `Credit`. Then set D18 to FALSE, click again, and confirm the panel shows `Failed` with the previous output kept and `Property\` unchanged; restore D18 to TRUE. Pass: TRUE run shows `Succeeded` and the output opens and matches step 9's checks; FALSE run stops as described. [Partial — live installer and TRUE/FALSE acceptance passed through direct P1 dispatch; the installed VBA button was not clicked]
11. **Docs** (parallel with 8-10): add a "Property Stage 1" section to `references/bonds-lite-setup.md`. [✓ Implemented — see addendum]
12. **Packaged release** (depends on 10, separate approval). [Pending separate approval]

## Test Strategy

Unit tests use the existing `SimpleNamespace`/fake-book patterns (`tests/test_sst_stage_four.py` installer fakes; `tests/test_sst_stage_three.py` `_GuardCaller`/`_NamedRange`).
- **Property flag:** TRUE runs; FALSE stops with the "not selected" message, writes `Failed`, keeps the previous output and never calls `allocate_stress_folder`; a missing name and non-boolean values (`"TRUE"`, `1`, `None`) stop with the TRUE/FALSE message. Regression: Credit `read_stress_flags` and `validate_stress_selection` ignore the Property flag (Credit FALSE/FALSE still stops when Property is TRUE).
- **Inputs:** template path absolute/relative/missing/URL/wrong extension; missing name; non-date valuation date; invalid suffix.
- **ScenFlipper guards:** a synthetic 3-scenario sheet with distinctive values; failures for header mismatch, duplicate header, B2 != row count, B3 = 0 or > B2, invalid B1, error value, cross-sheet formula, used range not at A1, missing sheet in either book.
- **Replace in place:** asserts `Clear` on the union range (template bigger than Draft, and smaller), `Copy` destination address, the sheet object is never deleted/added, and verification catches a changed `prm_*` name, a `#REF!` dependent formula and a new link source.
- **Valuation date:** written and read back; different-date log message; missing `ValDate` stops.
- **Folder:** first run; rerun replaces `Property` and leaves `Credit` untouched; failure restores the previous `Property`.
- **Audit:** status sequence, timings, portable paths, previous output kept on failure, unsaved Draft is an error.
- **Installer:** geometry from `SST_Stage1_Panel` at O9:V19 and at a moved position; occupied/merged/overlapping area stops; Draft open stops; backup before changes; idempotent rerun; partial names stop.
- **Opt-in Excel (`excel` marker):** step 9, using real files copied to `tmp_path`.

## Affected Files
- `src/sst_property_stage_one.py` - new; all Property Stage 1 logic.
- `tests/test_sst_property_stage_one.py` - new.
- `src/bonds_lite_input.py` - folder allow-list, CLI choice/flag/dispatch, VBA launcher text.
- `tests/test_bonds_lite_input.py` - dispatch, VBA and folder tests.
- `tools/build_packaged_stages.py`, `tools/controller_entry.py` - hidden import and CLI mirror.
- `references/bonds-lite-setup.md`, `references/decisions.md` - docs and decisions.
- Draft `I. Control` - button, panel, names and VBA, only via the installer at step 10.

## Decisions
See `references/decisions.md`, Decisions 87-95 (approved 6 November 2026).

## Risks And Unknowns
- **Flag name not yet installed:** until the installer runs, `SST_Property_standalone` does not exist and the stage stops with a missing-name error. The installer adds it.
- **Stale calculated values:** with no recalculation, a changed `ValDate` leaves `P. Curves` and outputs stale until the LTM run recalculates. If the dates are equal (as today), nothing is stale.
- **External links:** the template's `server090` links are kept as they are, not updated; a later run stage must handle them.
- **ScenFlipper growth:** more scenarios in the Draft than rows the template's formulae expect would pass Stage 1 but may not be used by LTM; the guards check only shape and headers.
- **Template already open:** `owned_book` stops if a book with the same name is open in the caller's Excel.
- **Interim button state:** until a packaged release contains `P1`, the packaged launcher rejects `--stage P1` (same as Stages 3-4 before release).
- **VBA signature change:** changing `LaunchBondsLite` to `String` touches the shared launcher; tests must prove `RunMACalc`/`LaunchBondsLite 4` still work.

## Addendum: Implementation Progress

- Implemented the Property Stage 1 module, safe `Property` folder replacement, audit/run flow, installer, `P1` CLI/VBA dispatch, controller startup message, packaged hidden import, and setup documentation.
- Added unit coverage for flag validation, ScenFlipper guards, date and template path validation, in-place range clearing/copying, folder rollback, CLI dispatch, launcher text, and package builder configuration. The opt-in Excel test uses temporary copies and checks the original workbook hashes.
- Targeted validation: Property Stage 1 tests pass (39 passed, 1 skipped). Black check and `git diff --check` passed.
- The opt-in Excel test passed on temporary copies (1 passed); the original Draft and LTM Flash template hashes were unchanged.
- Live installation used the installer backup `references/backups/Draft Input spreadsheet for SST_20261107_021453_090119.xlsm`. Source-mode TRUE runs succeeded and saved `SST Output_Test\Property\LTM_FlashModel_v3.2a_Test.xlsb`; the FALSE run recorded `Failed`, retained the last successful output, and left the Property folder contents and file hash unchanged. The flag was restored to TRUE. The P1 dispatch was invoked directly through `main()` rather than by clicking the installed VBA button.
- One live run initially left a staging copy because OneDrive denied removal of its read-only folder. Its workbook hash matched the published output; the staging folder was safely removed. Cleanup now uses the existing read-only retry helper, and a subsequent live TRUE run confirmed the staging folder is removed.
- Full validation (`PYTHONPATH=.;src pytest -q`): 1,213 passed, 90 skipped, 2 failed. The remaining failures are outside this implementation: `test_frozen_main_validates_cached_bundle_before_stage_dispatch` calls `main(argv)` although `main` accepts no argument, and `test_model_results_transaction_rollback_failure_retains_recoverable_backup` expects a recovery path that the current rollback code does not report. No unrelated code was changed.
- No implementation decisions deviated from Decisions 87-95.
