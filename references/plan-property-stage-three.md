# Plan: Property Stage 3 - Prepare NoteVal From The Stage 2 LTM Flash Run (And Entity Prefix For Stage 1)

Version 1.0 (approved for implementation) | 11 November 2026 | Decisions 113-120 approved by the user.

## Overview

Property Stage 3 copies the LTM Note Valuation Tool ("NoteVal") named in the Draft (`SST_LTMNote_input_template`, `'I. Control'!D47`) and repoints its links from the old LTM Flash workbook to the one that Property Stage 2 just ran. It changes only the external workbook reference in `I_LTM_IFRS_Value` (`'I. Control'!D12`) and in every formula in `'I. Inputs'` columns C, E and G and cell I10. Each formula keeps its sheet, cell addresses and arithmetic. It also writes the Draft `I_Valuation_Date` into NoteVal `I_ValuationDate`. The copy is saved beside the Stage 2 Flash workbook as `<Entity>_<template stem>_<suffix><ext>`. In the same change, Property Stage 1 gains the same entity prefix for its Flash output. Running the NoteVal `Run` macro is **out of scope** for this stage (Decision 113, user-approved). Property Stage 4 will run it, following the Property Stage 2 pattern, so that the NoteVal setup is checked after Stage 3 and before the run. Stage 4 will take its input from the Stage 3 audit.

## Verified Discovery (Read-Only, 11 November 2026)

Read from temporary copies; the originals were not changed. Excel cannot open `LTM Note Valuation Tool v1.40.xlsb` in place: the file is a OneDrive `ReparsePoint` and Excel reported "cannot access the file". A local copy opened normally (see Risk R1).

### Draft (`Draft Input spreadsheet for SST.xlsm`)
- Entity: `'I. Control'!D9` = `JRL` (label C9 "Select the life entity:"). **No defined name** exists for it.
- `I_Valuation_Date` -> `'I. Control'!D11` = 31/12/2025. `SST_output_folder_suffix` -> D63 = `Test`.
- `SST_LTMNote_input_template` -> `'I. Control'!D47` = `LTM Note Valuation Tool v1.40.xlsb` (relative to the Draft, as D46 is for Flash).
- Property Stage 1 panel `SST_PropStage1_Panel` = `AB9:AI19`; Stage 2 `SST_PropStage2_Panel` = `AB22:AI32`. `AB35:AI45` is free, matching Credit Stage 3 at `O35:V45`.
- Stage 2 handoff currently recorded: `SST_PropStage2_Status` = `Succeeded`, `SST_PropStage2_InputWorkbook` = `SST Output_Test/Property/LTM_FlashModel_v3.2a_Test.xlsb` (equal to `SST_PropStage1_LastSuccessfulOutput`), `SST_PropStage2_LastSuccessfulRun` = 11/11/2026 17:44:27. Stage 2 runs Flash **in place**, so the Flash workbook is the Stage 1 output file.

### NoteVal template (`LTM Note Valuation Tool v1.40.xlsb`, 6.4 MB, has a VBA project)
- `I_ValuationDate` -> `'I. Control'!D10`, typed constant 30/06/2026 (label: "Enter the valution date (last working day format)").
- `I_LTM_IFRS_Value` -> `'I. Control'!D12` = `=SUM('<url>/01 Flash/[202606_JRL_LTM_FlashModel_v3.2a (Base and Property Value).xlsb]O. PVOutput'!$CJ$3:$CL$3)-SUM('...'!$CM$3:$CN$3)`.
- `'I. Inputs'` used range `B2:FE729`. Linked formulas: C10:C729, E10:E729, G10:G729 (rows 9 are text headers `Gross`, plus E7/G7 notes) and I10. Every link reads `O. CFOutput` of the same Flash workbook, for example:
  - `C10 = SUM('...[Flash]O. CFOutput'!CJ3:CL3,-'...[Flash]O. CFOutput'!CM3)` (Redemptions+ISMInt+ERC-Expenses)
  - `E10 = -'...[Flash]O. CFOutput'!CN3` (NNEG_100%, "Only for MI")
  - `G10 = -'...[Flash]O. CFOutput'!CO3` (IntNNEG_100%, "For Waterfall IFRS")
  - `I10 = -'...[Flash]O. CFOutput'!$CT$3:$DM$722` (NNEG attachment points)
- `I10` is the anchor of an Excel dynamic-array spill over `I10:AB729` (720 rows by 20 columns); the spill cells are results of that one formula, not separate formulas.
- **Workbook-wide scan:** the Flash workbook is referenced in exactly 2,162 formulas: `'I. Control'` (1 cell, D12) and `'I. Inputs'` (2,161 cells in columns C, E, G, I). No defined names refer to it. The user's scope therefore matches every reference to that source exactly.
- `LinkSources` lists six workbooks. Only one is Flash; the other five (EIOPA Adj Processor, ALM Output, NNEG Hedge Allocator, LTM Note Schedules, MAP Cash & Securitisation AoC) must be left untouched.
- VBA module `modControl` has `Public Sub Run()`, assigned to a button on `'I. Control'` ("Click on the button to run the model"). Its last recorded run took about 9 minutes.

### Stage 2 Flash output (`SST Output_Test/Property/LTM_FlashModel_v3.2a_Test.xlsb`)
- `O. CFOutput` used range `A1:FQ734`; row 2 headers at CJ:CO = `Redemptions, ISMInt, ERC, Expenses, NNEG_100%, IntNNEG_100%`; CJ data ends at row 722, matching the NoteVal ranges (rows 3-722). `CT2` = `NNEG_100%`.
- `O. PVOutput` row 3 holds the CJ:CN values used by `I_LTM_IFRS_Value`.
- So the existing NoteVal formulas already point at the right sheets and cells; only the workbook path must change (as the user stated).

### Code
- `src/sst_property_stage_one.py` (has **uncommitted user changes**: Interest/Property block selection, Decision 112). Output name is built in `run_property_stage_one` as `f"{template.stem}_{suffix}{ext}"`; Stage 2 picks the path up from `SST_PropStage1_LastSuccessfulOutput`, so a renamed Stage 1 output flows through to Stage 2 with no Stage 2 change.
- `src/sst_property_stage_two.py`: `resolve_property_stage_two_inputs`, `_resolve_handoff_path`, `_ensure_workbook_not_open`, `_excel_local_path` (SharePoint URL mapping), panel/button/macro installer pattern (`install_property_stage_two_controls`, macro `RunLTMFlash` -> `LaunchBondsLite "P2"`).
- `src/bonds_lite_input.py`: shared helpers (`resolve_caller_path`, `_control_cell`, `read_output_suffix`, `workbook_run_guard`, `owned_book`, `excel_state`), CLI `--stage` choices `1-4, P1, P2` and `--install-property-stage-two`.
- `tools/controller_entry.py`, `tools/run_stage_from_source.py` and `tools/build_packaged_stages.py` (hidden imports) list the stages and need `P3`.

## Design

### Entity name and prefix (Decision 116, 117)
- The installer adds a workbook name `SST_Entity` -> `'I. Control'!$D$9`. It is idempotent: if the name already exists it must point at that cell, otherwise the installer fails with a clear message.
- New shared reader `read_entity(caller) -> str` in `bonds_lite_input.py`. It strips whitespace, requires a non-empty string and rejects characters that are invalid in Windows filenames, using the same rules as `validate_output_suffix`. A missing name raises an actionable error ("run --install-property-stage-three ...").
- New helper `entity_output_name(entity, template_path, suffix) -> str` returning `f"{entity}_{stem}_{suffix}{ext}"`. Both Property stages use it:
  - Stage 1: `JRL_LTM_FlashModel_v3.2a_Test.xlsb`
  - Stage 3: `JRL_LTM Note Valuation Tool v1.40_Test.xlsb`
- The prefix is not doubled when the template stem already starts with `<entity>_`. The template names are user-entered, and the old SharePoint Flash file was itself named `202606_JRL_...`.
- Stage 1 reads the entity during input validation, so a blank entity fails before the Property folder is replaced.

### Stage 3 handoff (Decision 115)
`resolve_property_stage_three_inputs(caller)`:
1. Require `SST_Property_standalone` TRUE (same gate as Stages 1 and 2).
2. Require `SST_PropStage2_Status == "Succeeded"` and a valid `SST_PropStage2_LastSuccessfulRun`.
3. Resolve `SST_PropStage2_InputWorkbook` with Stage 2's `_resolve_handoff_path`; require it to equal the resolved `SST_PropStage1_LastSuccessfulOutput` and to exist. Also require Stage 2's last successful run to be no earlier than Stage 1's finish, so a Stage 1 rerun cannot be paired with an older Stage 2.
4. Resolve `SST_LTMNote_input_template` relative to the Draft (reuse Stage 1 `_resolve_template_path`); it must exist, be `.xlsb`/`.xlsm`/`.xlsx`, and must not be the output path itself.
5. Read `I_Valuation_Date` (Stage 1 `_date_value`), the entity and the suffix.
6. Output path = `flash_path.parent / entity_output_name(...)`. Check that the folder is the current `SST Output_<suffix>\Property` folder beside the Draft (reuse `_property_output_path`).

### Run sequence (`run_property_stage_three`)
Follows the Stage 1 structure (guard -> Running audit -> work in temp -> verify -> commit -> Succeeded audit; Failed/Interrupted audit on any error).
1. `workbook_run_guard`; check that neither the Flash workbook nor the target NoteVal file is open (reuse `_ensure_workbook_not_open`).
2. Copy the template to a `~property-stage3-*` temp folder **inside the Flash folder**. Keeping it in the same folder lets Excel store the Flash link as a relative path and lets the final rename be atomic.
3. In the caller's Excel instance, with `excel_state` + manual calculation, open the copy with `update_links=False`, `ignore_read_only_recommended=True` and AutomationSecurity "force disable" (no NoteVal macros may run).
4. **Snapshot and scope check** (`snapshot_flash_links`): from `I_LTM_IFRS_Value`'s formula, take the single external workbook it references (the "old Flash source"). Find its full path in `LinkSources()`. Use `SpecialCells(xlCellTypeFormulas)` across every sheet to find each formula containing that source, and fail unless the set equals exactly {`'I. Control'!D12`} plus `'I. Inputs'` C/E/G formulas plus I10. Also fail if any C/E/G formula (rows 10+) or I10 references another workbook, and if any defined name refers to the old source. Record every in-scope formula text and the other link sources.
4a. **Stressed-scenario check** (`check_stressed_scenario`, Decision 120). Open the Flash workbook read-only in the same Excel instance, with links not updated and macros disabled. It stays open for the value checks in step 8. Take the first column that the NoteVal formulas reference (`CJ` in v1.40) and require all of these:
   - Row 1 of that column holds the same scenario number, `9`, in both `O. CFOutput` and `O. PVOutput`.
   - In the Flash `I. ScenFlipper`, the row for scenario 9 has a column J `scen desc` of `Interest`, `Property` or `Interest_Property`, not `Base`.
   - Scenario 8 is `Base`, so scenario 9 is the first scenario of the stress block.
   On failure, stop before any change, with a message naming the scenario number and description found. This guards against a Base block that is not 8 rows, which would make NoteVal value the wrong scenario.
5. **Repoint** with `Workbook.ChangeLink(old_source, flash_path, xlLinkTypeExcelLinks)` (Decision 114). One call updates every reference to that source. Excel reads the closed Flash workbook's values, and formula structure is untouched.
6. **Write** `I_ValuationDate` = Draft valuation date as a date value; keep the cell's number format.
7. Calculate the workbook (`CalculateFull`). Calculation only, so the `Run` macro is not executed.
8. **Verify** (`verify_noteval`):
   - `LinkSources()` no longer contains the old Flash source, contains `flash_path`, and the other five sources are unchanged.
   - Each in-scope formula equals `rewrite_external_reference(original, old_source, flash_path)`, a pure function. Comparison is on the workbook/sheet/cell structure: Excel may write the path as relative or absolute, so path text is normalised before comparing.
   - Values: open the Flash workbook read-only in the same instance and check `I_LTM_IFRS_Value == SUM(PV CJ3:CL3) - SUM(PV CM3:CN3)`. Require the I10 spill to match the dimensions and element-by-element negation of its Flash range (`I10:AB729 == -CF CT3:DM722` for this template; Decision 124). Also check every C/E/G row against the CF row it references, within 1e-6 relative tolerance. No `#REF!`/`#VALUE!` errors.
   - `I_ValuationDate` reads back equal to the Draft date.
9. Save; require `book.api.Saved`. Close; `os.replace(temp, output)`. A rerun overwrites the previous Stage 3 output, as Stage 1 does for LTM Flash (user, Q5), but only after the new copy is verified (Decision 117). If the run fails, the previous output is left as it was. Unlike Stage 1, Stage 3 replaces only its own file and not the whole `Property` folder, which also holds the Flash workbook. A Stage 1 rerun still replaces the whole folder, NoteVal output included, which is correct because that NoteVal would point at a stale Flash run. Remove the temp folder.
10. Write the Succeeded audit with the portable Flash and output paths.

### Audit panel, button and CLI (Decision 118)
- Panel `SST_PropStage3_Panel` = `'I. Control'!AB35:AI45`, with fields `LastRun, Start, Finish, Elapsed, User, Status, InputWorkbook (Flash), LastSuccessfulOutput`. Field offsets follow the Stage 2 panel.
- Button `btnPrepareNoteVal` beside the panel (Stage 2 layout); macro `Public Sub PrepareNoteVal()` -> `LaunchBondsLite "P3"`, installed idempotently with the Stage 2 source check.
- CLI: `--stage P3` and `--install-property-stage-three <draft>` in `bonds_lite_input.main`. The installer also adds `SST_Entity` and takes the Draft backup to `references/backups/`, as the Stage 2 installer does.

### Valuation date (Decision 119)
The Draft date is written unchanged. NoteVal labels the cell "last working day format". Stage 3 logs a warning if the date falls on a Saturday or Sunday, but does not fail, because Stage 1 has already fed the same date to Flash. UK bank holidays are not checked.

## Steps

1. **Tests for the entity reader and name helper** (`tests/test_bonds_lite_input.py`): see the Test Strategy, tests T1-T4. All must fail first. [✓ Implemented — see addendum]
2. **Implement `read_entity` and `entity_output_name`** in `src/bonds_lite_input.py`. Verify: T1-T4 pass and the existing suite stays green. *Depends on 1.* [✓ Implemented — see addendum]
3. **Stage 1 prefix tests** (`tests/test_sst_property_stage_one.py`): T5-T7; update every existing assertion that expects `LTM_FlashModel_v3.2a_<suffix>.xlsb`. These edits build on the **uncommitted Decision 112 changes**, which must be preserved, so commit them first or work on top of them. *Depends on 2.* [✓ Implemented — see addendum]
4. **Stage 1 prefix implementation**: `read_property_stage_one_inputs` adds `entity` to `PropertyStageOneInputs`; `run_property_stage_one` uses `entity_output_name`. The Stage 1 installer also ensures `SST_Entity`. Verify: T5-T7 and the full Stage 1 and Stage 2 test files pass. Stage 2 needs no change because it reads the path from the audit. *Depends on 3.* [✓ Implemented — see addendum]
5. **Pure Stage 3 tests** (`tests/test_sst_property_stage_three.py`): T8-T16 and T24-T26 (formula rewrite, scope check, handoff, naming, date, stressed-scenario check). *Parallel with 3-4 after 2.* [✓ Implemented — see addendum]
6. **Create `src/sst_property_stage_three.py`** with constants, `PropertyStageThreeError(stage_two.PropertyStageTwoError)`, `rewrite_external_reference`, `snapshot_flash_links`, `resolve_property_stage_three_inputs`, `set_noteval_valuation_date`, `verify_noteval`, audit read/write, `run_property_stage_three`, and the installer. Verify: T8-T16 pass. *Depends on 5.* [✓ Implemented — see addendum]
7. **Run-sequence tests with fakes**: T17-T21 (success path, failure leaves the old output and records Failed, interruption, no macro executed, guard conflict). Implement until they pass. *Depends on 6.* [✓ Implemented — see addendum]
8. **Wire the CLI and packaging**: add `P3` to the choices in `bonds_lite_input.py`, `tools/controller_entry.py` (plus a "Starting Property Stage 3..." message) and `tools/run_stage_from_source.py`; add `sst_property_stage_three` to the hidden imports in `tools/build_packaged_stages.py`; add `--install-property-stage-three`. Verify: T22 and the existing CLI tests. *Depends on 6.* [✓ Implemented — see addendum]
9. **Opt-in Excel integration test** (`@pytest.mark.excel`): T23 on temp copies of the real NoteVal template, the `SST Output_Test` Flash workbook and a Draft copy. *Depends on 7-8.* [Implemented — post-change real-workbook rerun pending; see addendum]
10. **Install on a Draft backup copy first**, then on the live Draft with user approval: run `--install-property-stage-three`, check the panel, button, `SST_Entity`, and that `RunLTMFlash` is untouched. *Depends on 8.* [✓ Implemented — see addendum]
11. **Manual acceptance** (user): run P1 -> P2 -> P3 on the Test suffix; open the NoteVal output in Excel, confirm Edit Links shows the local Flash and the five unchanged sources, then run the NoteVal `Run` macro by hand and compare with the last manual NoteVal result. *Depends on 10.*
12. **Docs**: update `references/bonds-lite-setup.md` (Property Stage 3 button, entity name requirement) and add an Implementation Progress addendum to this plan. *Depends on 11.*

## Test Strategy

Write each test first and confirm it fails before implementing. Fakes follow `tests/test_sst_property_stage_two.py` (fake caller/book/range objects). Fixtures are synthetic formula strings copied from the verified NoteVal formulas above; there are no binary fixtures in unit tests.

| ID | Test | Input -> expected |
|----|------|-------------------|
| T1 | `test_read_entity_returns_trimmed_value` | `" JRL "` -> `"JRL"` |
| T2 | `test_read_entity_rejects_blank_or_non_text` | `None`, `""`, `5` -> error naming `SST_Entity` |
| T3 | `test_read_entity_rejects_invalid_filename_characters` | `"JR/L"`, `"A:B"` -> error |
| T4 | `test_entity_output_name_prefixes_and_does_not_double` | (`JRL`, `LTM Note Valuation Tool v1.40.xlsb`, `Test`) -> `JRL_LTM Note Valuation Tool v1.40_Test.xlsb`; stem `JRL_X` -> `JRL_X_Test.xlsb` |
| T5 | `test_stage_one_output_name_has_entity_prefix` | fake run -> `.../Property/JRL_LTM_FlashModel_v3.2a_Test.xlsb`; audit `last_output` matches |
| T6 | `test_stage_one_blank_entity_fails_before_folder_replacement` | no folder allocation call; Failed audit |
| T7 | `test_stage_two_accepts_prefixed_stage_one_handoff` | Stage 2 resolves the prefixed path |
| T8 | `test_rewrite_external_reference_cf_formulas` | C10/E10/G10/I10 strings -> same formula with only `<path>[<file>]` changed; `$` and ranges preserved |
| T9 | `test_rewrite_external_reference_pv_formula` | D12 -> both `SUM` terms repointed |
| T10 | `test_rewrite_handles_apostrophes_and_spaces_in_paths` | old/new paths containing `'` (escaped `''`) and spaces |
| T11 | `test_scope_check_accepts_exact_reference_set` | fake workbook with refs only in D12 + I. Inputs C/E/G/I10 |
| T12 | `test_scope_check_rejects_reference_outside_scope` | extra ref in another sheet -> error listing the cell |
| T13 | `test_scope_check_rejects_second_source_in_scope_cells` | a C cell referencing another workbook -> error |
| T14 | `test_resolve_stage_three_requires_successful_stage_two` | status `Failed` / missing run time -> error |
| T15 | `test_resolve_stage_three_requires_matching_handoffs` | Stage 2 input != Stage 1 output, or Stage 2 earlier than Stage 1 finish -> error |
| T16 | `test_resolve_stage_three_output_beside_flash` | output = Flash folder / `JRL_..._Test.xlsb`; template equal to output path -> error |
| T17 | `test_run_repoints_sets_date_verifies_and_saves` | calls in order: open (links off) -> snapshot -> ChangeLink -> date -> calculate -> verify -> save -> replace; Succeeded audit |
| T18 | `test_run_failure_keeps_previous_output_and_records_failed` | verify raises -> previous output byte-identical, temp removed, Failed audit, last successful output unchanged |
| T19 | `test_run_records_interrupted_on_keyboard_interrupt` | |
| T20 | `test_run_never_invokes_noteval_macro` | fake `app.api.Run` must not be called; AutomationSecurity forced off during open |
| T21 | `test_run_guard_rejects_concurrent_property_stage` | |
| T22 | `test_cli_accepts_stage_p3_and_installer` | argument parsing and dispatch |
| T23 | `test_stage_three_on_real_noteval_copy` (excel, opt-in) | 2,162 repointed formulas; old source gone; 5 other links unchanged; all 14,400 I10 spill values and the other checked values match Flash `O. CFOutput`/`O. PVOutput`; `I_ValuationDate` = 31/12/2025; file named `JRL_LTM Note Valuation Tool v1.40_Test.xlsb` beside the Flash workbook; scenario check passes (scenario 9 = first `Interest_Property`) |
| T24 | `test_stressed_scenario_check_accepts_first_stress_scenario` | CF/PV row 1 at `CJ` = 9; ScenFlipper scenarios 1-8 `Base`, 9 `Property` -> passes |
| T25 | `test_stressed_scenario_check_rejects_base_or_shifted_block` | scenario 9 `Base` (Base block of 10 rows), or scenario 8 not `Base` (Base block of 6 rows) -> error naming the scenario and description; no `ChangeLink` call |
| T26 | `test_stressed_scenario_check_rejects_mismatched_header` | CF row 1 at `CJ` = 9 but PV = 10, or blank -> error |

Regression tolerance for T23 values: 1e-6 relative. They are identical cell reads, so any difference indicates a wrong reference.

## Affected Files

- `src/bonds_lite_input.py` - `read_entity`, `entity_output_name`, CLI `P3` and `--install-property-stage-three`.
- `src/sst_property_stage_one.py` - entity in inputs, prefixed output name, `SST_Entity` in installer (on top of the uncommitted Decision 112 work).
- `src/sst_property_stage_three.py` - **new** module.
- `tests/test_bonds_lite_input.py`, `tests/test_sst_property_stage_one.py`, `tests/test_sst_property_stage_two.py` (T7), `tests/test_sst_property_stage_three.py` (**new**).
- `tools/controller_entry.py`, `tools/run_stage_from_source.py`, `tools/build_packaged_stages.py` - `P3` wiring.
- `Draft Input spreadsheet for SST.xlsm` - via the installer only (name, panel, button, macro), with a backup first.
- `references/bonds-lite-setup.md`, `references/decisions.md`, this plan.

## Decisions

Full entries are in `decisions.md`. Decisions 113-120 were approved by the user on 11 November 2026.
- **113** Stage 3 prepares and saves NoteVal only; Property Stage 4 runs `Run`.
- **114** Repoint with `Workbook.ChangeLink` on the single Flash source, guarded by an exact-scope check and post-verification. User-approved. Formulas are not replaced with values.
- **115** The handoff comes from the Stage 2 audit (Succeeded, matching the Stage 1 output, not older than Stage 1). User-approved.
- **116** Add the `SST_Entity` name for `'I. Control'!D9`; validate it as filename-safe. User-approved.
- **117** Name outputs `<Entity>_<template stem>_<suffix><ext>` in Stages 1 and 3; Stage 3 replaces its output only after verification.
- **118** Panel `AB35:AI45`, button `btnPrepareNoteVal`, macro `PrepareNoteVal`, CLI `P3`. User-approved.
- **119** Write the valuation date unchanged; warn on weekends.
- **120** Stop unless the Flash columns NoteVal reads are scenario 9 and that is the first stressed scenario (not Base). User-approved.

## Risks And Unknowns

- **R1 OneDrive placeholder.** Excel could not open the NoteVal template in place, though a copy worked. Stage 3 opens a **copy** anyway (step 2), so this is avoided; `shutil.copy2` triggers the OneDrive download. If the copy size is wrong or zero, fail with "make the file available offline".
- **R2 Which Flash scenario NoteVal reads (resolved, Decision 120).** NoteVal values the stressed IFRS scenario, which the fixed columns `CJ:CO` give as scenario 9 (user, 11 November 2026). This holds only while the Base block has 8 rows, so Stage 3 checks it at step 4a.
- **R3 ChangeLink side effects.** Excel may prompt or recalculate on `ChangeLink`. `display_alerts=False` and manual calculation are set; T23 must prove that no other link is refreshed.
- **R4 Uncommitted Stage 1 work.** The Decision 112 changes in `src/sst_property_stage_one.py` and its tests are not committed. Implementation must build on them, not revert them; committing them first is recommended.
- **R5 Existing Stage 1 outputs.** Outputs made before this change keep the unprefixed name. Stage 3 will still work because it follows the audit path, but its own output is always prefixed.
- **R6 NoteVal `Run` duration (about 9 min) and its own links** (five server090/SharePoint sources) matter only if Stage 4 later runs the macro.

## Questions

Answered 11 November 2026:
1. The `Run` macro goes in Property Stage 4, so that the NoteVal setup is checked after Stage 3 first.
3. Keep the suffix: `<Entity>_<template stem>_<suffix>`.
4. Weekend dates are not expected; if one occurs, log a warning and do not stop.
5. Overwrite the previous output on rerun, as Stage 1 does.
2. NoteVal values the stressed IFRS scenario, which is scenario 9 in the Flash run (a). Stage 3 checks this and stops if it does not hold (b). Recorded as Decision 120.

## Addendum: Implementation Progress

- Steps 1–8 completed. Step 9's real-workbook test originally passed using temporary copies; its post-change rerun remains pending. Focused Property Stage 1–3/CLI tests passed except for two unrelated existing `bonds_lite_input` failures; the full suite reported 1,398 passed, 93 skipped and the same two failures.
- The Excel integration run emitted `0x80010108` COM diagnostics while the three installers were exercised, although pytest completed the test successfully. No live Draft was modified.
- Steps 10–12 remain pending: install on a Draft backup and obtain approval before any live installation; complete manual P1 → P2 → P3 acceptance; then finish the setup guide. The setup guide was not changed as part of this implementation.
- Decision 121's protected-sheet fallback was approved by the user (Decision 123). Decision 122's single-cell I10 verification was rejected in favour of the complete spill range (Decision 124). Stage 3 preserves the existing formula and verifies its spill dimensions and every value.
- A previously prepared real-workbook copy had a 720 × 20 spill with no mismatches against the Flash range, including after recalculation with Flash closed. Follow-up Property tests reported 159 passed and 4 skipped. The updated verifier also passed an opt-in Excel test on disposable workbooks, including detection of a mismatch in `AB729`; Excel emitted a `0x80010001` COM diagnostic despite the test passing. A post-change run of T23 remains pending because the OneDrive source workbook was inaccessible.
- Step 10 completed. The installer first succeeded on a temporary Draft copy and then on the live Draft, with the user's approval. It created the backup `references/backups/Draft Input spreadsheet for SST_20261111_194509_512755.xlsm`. A read-only check confirmed:
  - `SST_PropStage3_Panel` is `'I. Control'!$AB$35:$AI$45`; Status is `Not Run`, and all eight audit field names are present.
  - `btnPrepareNoteVal` is at `W35:AA38` and calls `PrepareNoteVal`, which runs `LaunchBondsLite "P3"`.
  - `SST_Entity` refers to `$D$9`.
  - The existing Stage 1/2 panels and buttons are unchanged.
- The packaged controller release is still `20261109-cache-v25`, which predates P3. The button cannot run until a P3-capable release is built and promoted. Until then, P3 can run from source with `tools/run_stage_from_source.py P3`.