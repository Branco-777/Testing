# Plan: Stage 3 MACalc Input (Slice A: Template Checks And Funded-Leg Transfer)

Version 0.15 | 3 November 2026 | Slices A-C committed in `0f6ef09` and released as `20261103-cache-v6`. **Slice D (unique MACalc input filename: run ID plus a fresh random token) implemented on 3 November with the user's answers to Q-D1 to Q-D4 (Decision 54 approved with changes); `20261103-cache-v7` promoted to the central pointer on 3 November at 18:18 UTC and the user's packaged Stage 1-3 test passed. The staged-path limit was lowered from 259 to 218 characters and included in `20261103-cache-v8`, built and verified without promotion, then promoted with the user's approval at 18:54 UTC. Slice D was committed as `0fdf81a`.** Stage 5 runs the consuming Python MACCalc (`MACalc_Batch`) on the validated MACalc input; Stage 3 does not touch it (Decision 51)

## Current Release Status (6 November 2026)

Stage 3 funded-set validation uses the negative-cashflow ISINs in the Stage 1
Bonds input, not MAP (Decision 82). Release `20261106-cache-v15` is promoted;
v14 added auto-saving of an unsaved Draft before Stages 2-4 (Decision 83) and
v15 accepts templates whose asset rows hold the only links to a file
(Decision 84).
On v13, the packaged Draft-button Stage 3 run passed with the local cache against
`SST Output_Test/Credit`; the Draft audit is `Succeeded` and points to
`SST Output_Test/Credit/MACalc input_Test_d58742fd77236ce3.xlsb`. The
orchestrator verified the generated workbook before publishing it. Packaged
Stage 2 and Stage 4 acceptance remain pending.

## Overview

Stage 3 prepares a populated MACalc input workbook from the MACalc template, the Stage 1 Bonds Lite MAP input and the central MACE workbook. This plan covers Slice A only: validate the template's Mortgage/ILS blocks, identify funded assets from MACE's funding-leg cashflows, and write their funded-leg rows using the user-specified MAP and MACE column mappings. The cashflows come from MACE, not MAP. All Stage 3 logic lives in a new module, `src/sst_stage_three.py`. It reuses the existing Stage 1/2 handoff, audit and Excel-lifecycle helpers and does not add MACalc business logic to `src/bonds_lite_input.py`. Slice A was accepted by the user on 3 November 2026 (Decision 40). Later Stage 3 transfers will be planned as further slices; see "Planning The Next Slices" at the end of this plan.

Items marked **PENDING** are provisional defaults that the user can still change. None of them block implementation.

## Confirmed Requirements (User, 2 November 2026)

1. The MACalc template path comes from the Draft name `SST_MACalc_Input_Template`; the MACE path comes from `SST_MACE_Central_Template`.
2. In the template's `I. Assets` sheet, identified by the `Type` column, there must be a non-empty block of Mortgage rows at the top, immediately followed by a non-empty block of ILS rows. Otherwise show a pop-up asking the user to update the template to account for LTM/Swap assets, then stop.
3. The next block after ILS holds the funded leg of funded assets. An asset is funded when its MACE funding-leg cashflow vector (`C. Assets - Funding` AJ:AUM, months 1-1200) is negative. Months after the last cashflow are zero, so the operational rule is: at least one negative month and no positive month. (Corrected 2 November; v0.1 used the MAP vector.)
4. MAP `I. Bonds Data` to MACalc `I. Assets`:

   | MAP column (header) | MACalc column (header) | Note |
   |---|---|---|
   | C (ISIN) | D (ISIN) | |
   | C (ISIN) | E (Name) | ISIN deliberately used as Name |
   | E (Currency) | K (Currency) | |
   | F (Sector) | H (Sector) | |
   | H (NotchedRating) | I (Rating) | |
   | J (Index Linked?) | J (Index Linked?) | Values only, never the formula |

   MAP Q:ATT cashflows are **not** copied (user correction, 2 November).

5. MACE `C. Assets - Funding` to MACalc `I. Assets`, matched on ISIN:

   | MACE column (header) | MACalc column (header) | Transform |
   |---|---|---|
   | C (Z-Spread) | F (Zspread) | none |
   | D (Type) | G (Type) | none |
   | I (FS Add On) | L (FSAO) | none |
   | AI (0) | M (PV) | multiply by -1 |
   | AJ:AUM (months 1-1200) | N:ATQ (months 1-1200) | none; funding-leg cashflows |

6. Stage 3 depends on a validated Stage 2 output, is implemented in a separate Python file, and must have full guardrails.
7. Rows below the funded block are cleared; later slices fill them.
8. Columns A:C (Aallocations, Ballocations, PartialRecog) on funded rows are left empty. A later slice fills them from another MACE tab.
9. Interim outputs for all stages live in the timestamped Stage 1 run folder.
10. Zero funded assets is valid, particularly for other entities. Stage 3 warns the user and continues; it does not stop with an error.

## Verified Discovery (read-only, 2 November 2026)

All files were opened read-only in a separate hidden Excel instance. Nothing was saved.

- **Draft:**
  - `SST_MACalc_Input_Template` refers to `I. Control!D44`. Its value is the bare file name `MACalcinPython_v4.0f.xlsb`, which exists in the repo root.
  - `SST_MACE_Central_Template` refers to `P. Parameter!C8` and points to the reachable `\\server090\...\202606 - JRL - MACE v5.20 c2.xlsb`.
  - A separate `SST_MACalc_Central_Template` (`P. Parameter!C9`) points to a different MACalc file and is not used by this slice.
  - `I. Control!H44` holds a check formula that warns when central templates do not match the valuation date in D9.
- **MACalc template, `I. Assets`:**
  - Headers are on row 1: A Aallocations, B Ballocations, C PartialRecog, D ISIN, E Name, F Zspread, G Type, H Sector, I Rating, J Index Linked?, K Currency, L FSAO, M PV, then N:ATQ for months 1-1200.
  - Rows 2-12 are Mortgage (11 rows) and rows 13-20 are ILS (8 rows), so the template passes requirement 2.
  - Rows 21-1008 hold 988 existing Bond/PP/Infrastructure rows, matching the MAP Bonds Data assets in the same order.
  - Columns A:B on every row are formulas looking up an external `202603 - JRL - MAP_Allocations_v3.22.xlsb`. All other cells are values.
  - The sheet has an AutoFilter. Defined name `tbl_Assets` sizes itself by COUNTA over column A and row 1.
- **MAP input** (`Bonds Lite output/20261102_091803/Bonds Lite input_MAP.xlsx`, the last Stage 2-validated run):
  - Headers are on row 8 and data on rows 9-996 inside table `I_tbl_Monthly_Bonds`.
  - All referenced columns match the headers above, and Q:ATT are numeric.
  - Column J is a formula with cached values.
  - ISINs are unique.
- **Funded assets: 5.** `GR0000000013_JRINFRA_Aviva`, `IL0000000122_JRINFRA_Aviva`, `IL0000000167_JRINFRA_BlackRock`, `IL0000000167A_JRINFRA_BlackRock`, `JRCM60_JRINFRA_Rothschild`. Each appears exactly once in both MAP and MACE.
- **MAP sign pattern:** no MAP asset has an all-negative vector. The 5 funded assets have mixed signs (1-26 negative months, 31-201 positive months), so the MAP vector cannot identify funded assets under the corrected rule.
- **MACE, `C. Assets - Funding`:**
  - Headers are on row 9 and data on rows 10-1018. The ISIN is in column B and is unique.
  - AI9 = 0, and AJ9:AUM9 = 1..1200, which align one-to-one with MACalc N:ATQ.
  - Of 1,009 rows, exactly 5 have negative-only AJ:AUM vectors, and these are the 5 assets above. The other 1,004 are all zero. No row has mixed signs or positive-only values, and the 5 funded vectors are fully numeric.
  - C, D and I are formulas linked to `I. Bonds`.
  - AI is `=-Q`, where Q is the funding-leg value total, so MACalc PV = AI * -1 = Q.
  - For the funding legs Q is negative, so PV is negative. Example: GR...13 gives PV of about -139.39m.
- **Observation (resolved):** MACE's funding-leg cashflows equal only the negative months of the MAP vector; the full MAP vector differs by up to 19.8m. The user confirmed that MACE AJ:AUM is the correct source.
- **Current Draft state:** the latest Stage 1 handoff is `20261102_111001`, but Stage 2 last succeeded on `20261102_091803`. Stage 3 would correctly refuse to run until Stage 2 is rerun on the current handoff.

## Defaults And Confirmations

Rows marked **Confirmed** were approved by the user on 2 November. The rest remain **PENDING** and can still be changed.

| ID | Topic | Default |
|---|---|---|
| P1 | Placement of funded block | **Confirmed.** Write from the first row after the ILS block (row 21). Clear A:ATQ for the rows below the funded block down to the last used row, so later slices fill them. Never shift or rewrite Mortgage/ILS rows. |
| P2 | Columns A:C on funded rows | **Confirmed for Slice A; superseded by Slice B Q-B1.** Slice A left A:C empty. In the combined run they are filled from MACE `Control`. |
| P3 | Funded row order | Keep MACE row order. In the current data this is also MAP order. |
| P4 | Zero funded assets | **Confirmed.** Print a prominent console warning and show a non-blocking pop-up at the end. Continue: write the output with an empty funded block, clear the rows from the first post-ILS row, and record Succeeded with a "0 funded assets" note in the console. |
| P5 | Type matching | Trimmed, case-insensitive match on `Mortgage` and `ILS`. Any blank Type or out-of-order row inside or between the blocks is invalid. |
| P6 | Pop-up text | "The MACalc template must start with Mortgage assets followed by ILS assets. Please update the template to account for LTM/Swap assets and rerun Stage 3." The console also lists the first offending rows. |
| P7 | Path resolution | A relative `SST_MACalc_Input_Template` resolves against the Draft's folder, as Stage 1 does. A relative MACE path is rejected. |
| P8 | MAP source file | `Bonds Lite input_MAP.xlsx` in the Stage 2-validated Stage 1 run folder, not the model's copy under `JRL MAP/`. Use cached values only. |
| P9 | MACE reading | Open read-only, `update_links=False`, no recalculation, use saved values. |
| P10 | MACE data guardrails | Stop on blank or duplicate ISINs in MACE, or if a funded row's C, I, AI or any AJ:AUM value is blank, an error or non-numeric. D (Type) must be non-blank text. Every funded ISIN must exist exactly once in MAP, which supplies the descriptive columns. |
| P11 | Valuation-date guard | Stop if `I. Control!H44` shows its mismatch warning. No other MACE date check in Slice A. |
| P12 | Output | **Confirmed folder.** `<Stage 1 run folder>/MACalc input.xlsb`, saved directly in the timestamped run folder as a copy of the template. The template is never modified. This is safe for Stage 2 because its runner only picks up direct-child `*.xlsx` files. Slice D (Decision 54) changes the filename to `MACalc input_<run ID>_<token>.xlsb`, with a fresh 16-character random token on every run. |
| P13 | Button and panel | New "Prepare MACalc" button with a movable `SST_Stage3_Panel` and named audit fields (`SST_Stage3_LastRun`, `_Start`, `_Finish`, `_Elapsed`, `_User`, `_Status`, `_InputRun`, `_LastSuccessfulOutput`), mirroring Stage 2. |
| P14 | Reruns | Ask before replacing an existing MACalc output. Back it up, restore it on failure and delete the backup only after verified success and saved audit. **Superseded by Slice D (Decision 54):** every run writes a new tokenised file, earlier outputs are kept untouched, and there is no prompt or backup. |
| P15 | Packaging | Include in the packaged controller after tests pass, as release `cache-v5`, with explicit user approval before promotion. |
| P16 | Anomalous MACE vectors | Stop with an error if any MACE funding row has mixed signs (both positive and negative months) or any positive month. Neither occurs in current data, and either would mean the funding leg is not purely a liability. All-zero rows are not funded. |
| P17 | MAP/MACE reconciliation | Stop if the MACE-funded set differs from the set of MAP assets with any negative month (min(Q:ATT) < 0). Both sets are the same 5 assets today. This catches MACE and Stage 1 data from different valuation dates or portfolios. |

## Steps

### 1. Write The Test Contract First

Affected file: new `tests/test_sst_stage_three.py`.

Write failing unit tests for every pure function in steps 3-6 using small synthetic in-memory tables:
- a few Mortgage/ILS rows;
- MACE funding vectors that are all zero, negative-only, negative then zero tail, mixed-sign, positive-only, and contain a blank or error;
- MAP rows with all-positive and mixed vectors, for the P17 reconciliation;
- 1200-month vectors built with NumPy;
- a MACE table with matching, missing and duplicate ISINs, plus a zero-funded table.

Separate opt-in Excel tests (same marker convention as Stage 1/2) use a disposable copy of the real template.

Pass criteria: tests collect and fail only for missing implementation. No production data is copied into `tests/fixtures/`. Small synthetic `.xlsx` fixtures are generated by test code.

[✓ Implemented — see Addendum: Implementation Progress]

### 2. Add Stage 3 Dispatch And Module Skeleton (depends on step 1)

Affected files: new `src/sst_stage_three.py`; `src/bonds_lite_input.py`, limited to `main()` and the `--stage` choices.

- Extend `--stage` choices to include `"3"` and lazily import `sst_stage_three.run_stage_three`, which avoids circular imports.
- Stage 3 imports only shared helpers from `bonds_lite_input`: path resolution, named-cell reads, audit writing, the run guard, the Excel app lifecycle and pop-ups.
- If any helper is private and needed, expose it with a thin public alias rather than duplicating logic.

Pass criteria:
- `--stage 3` dispatches.
- Existing Stage 1/2 tests remain green except for the two known pre-existing failures.
- `bonds_lite_input.py` contains no MACalc/MACE logic.

[✓ Implemented — see Addendum: Implementation Progress]

### 3. Resolve And Validate Stage 3 Inputs (depends on step 2)

Function `resolve_stage_three_inputs(caller) -> StageThreeInputs` (frozen dataclass).

- Read the Stage 2 audit and require `Status == Succeeded` and `SST_Stage2_InputRun == current committed Stage 1 handoff`. A stale or failed Stage 2 stops with a message naming both runs.
- Resolve the MAP path (P8), template path (P7) and MACE path. Each must exist, and they must be distinct files.
- Apply the valuation-date guard (P11).

Pass criteria: tests cover stale handoff, failed Stage 2, missing file, relative MACE path, the H44 warning, and the happy path.

[✓ Implemented — see Addendum: Implementation Progress]

### 4. Validate The Template Asset Blocks (parallel with step 5)

Function `validate_template_blocks(types: Sequence[object]) -> TemplateLayout` returns the Mortgage range, the ILS range and `funded_start_row`.

- Locate columns by header text on row 1 (ISIN, Name, Zspread, Type, ...) and assert that each sits at the expected letter (D, E, F, G, ... M, N, ATQ).
- If a header drifts, stop. Do not silently remap.
- If the block rule fails, raise `TemplateLayoutError`. The orchestrator turns this into the pop-up (P6) and stops before any file is written.

Pass criteria:
- Tests cover no Mortgage, no ILS, ILS before Mortgage, a blank row, Mortgage after ILS, header drift and month-header mismatch (N must be month 1 and ATQ month 1200).
- The real template yields Mortgage 2-12, ILS 13-20 and a start at row 21.

[✓ Implemented — synthetic tests pass; real template inspection not repeated]

### 5. Read MAP Descriptive Data (parallel with step 4)

Function `read_map_bonds(path) -> MapBonds`.

- Open in a dedicated hidden Excel app, read-only. Locate the header row (8) and table `I_tbl_Monthly_Bonds`.
- Verify headers C, E, F, H, J and the month headers Q:ATT (1-1200). Read values in one bulk array, not cell by cell.
- Require unique, non-blank ISINs. Keep C, E, F, H and J (cached value) per ISIN.
- Compute `map_negative_isins` (min(Q:ATT) < 0) for the P17 reconciliation only. MAP cashflows are never written.
  - **Superseded (Decision 82, 6 November 2026):** `read_map_bonds` now returns descriptions only. `read_bonds_negative_isins` reads the negative-month ISINs from `Bonds Lite input_Bonds.xlsx`, because Stage 1 removes negative cashflows from the MAP input.

Pass criteria:
- Synthetic tests cover each guard.
- On the 091803 MAP, `map_negative_isins` is exactly the 5 ISINs listed above.

[✓ Implemented — synthetic reader tests pass; production MAP not read]

### 6. Read MACE Funding Rows And Identify Funded Assets (parallel with step 4)

Functions `read_mace_funding(path) -> MaceFundingTable` and `identify_funded(table) -> list[MaceFunding]`.

- Open read-only with links not updated and no calculation (P9).
- Verify the row-9 headers: B ISIN, C Z-Spread, D Type, I FS Add On, AI = 0, and AJ:AUM = 1..1200.
- Read B:AUM down to the last ISIN in one bulk array.
- Classify each vector:
  - **funded:** at least one value < 0 and none > 0;
  - **not funded:** all zero;
  - **anomalous:** any value > 0 combined with a negative, or any positive month. Anomalous rows stop the run (P16).
- Apply the guards from P10 to funded rows. Compute `pv = -1 * AI` exactly.

Pass criteria:
- Tests cover each vector class, the boundaries -0.0 (not negative) and -1e-12 (negative), a missing or duplicate ISIN, an error value, a blank value, text in a numeric column, and correct negation (a positive AI gives a negative PV and the reverse).
- The real MACE yields exactly the 5 ISINs, 1,004 not-funded rows and 0 anomalous rows. The real GR...13 PV is negative.

[✓ Implemented — synthetic tests pass; production MACE not read]

### 6a. Reconcile MAP And MACE (depends on steps 5 and 6)

Function `reconcile_funded(map_bonds, funded) -> bool`, returning `True` only when both funded sets are empty.

- Every funded ISIN must exist in MAP (P10).
- The funded set must equal `map_negative_isins` (P17). On mismatch, stop and list the ISINs found only in MAP or only in MACE.
- If the funded set is empty (and so `map_negative_isins` is empty too), return `True` for the P4 warning; otherwise return `False`.

Pass criteria: tests cover equal sets, MAP-only, MACE-only, a missing MAP ISIN, and both sets empty (warning, not error).

**Change (Decision 82, 6 November 2026):** the signature is now `reconcile_funded(map_bonds, funded, bonds_negative_isins)`. The funded set must equal the negative-month ISINs of the Stage 1 Bonds input, not MAP. Mismatches are reported as "Bonds input only" or "MACE only". Every funded ISIN must still exist in MAP. Stage 3 stops if the Bonds input is missing from the run folder.

[✓ Implemented — see Addendum: Implementation Progress]

### 7. Build The Funded Block And Write The Output (depends on steps 3-6a)

Functions `build_funded_rows(...) -> list[list[object]]` (pure) and `write_macalc_output(...)`.

- Build each row in template column order D:ATQ. D, E, H, I, J and K come from MAP; F, G, L, M and N:ATQ come from MACE (Confirmed Requirements 4-5).
- Leave A:C empty (P2).
- Write values only, so J is never a formula.
- Copy the template to the staged output path (P12, P14) and open the copy in a hidden dedicated Excel app.
- Clear A:ATQ from the first post-ILS row down to the last used row.
- Write the block as one bulk range assignment from `funded_start_row`, and preserve the AutoFilter and `tbl_Assets`.
- With zero funded assets, write nothing after the clear (P4).
- Save as `.xlsb` and close. Never touch the user's Draft Excel instance or the template.

Pass criteria:
- A pure test asserts exact per-cell mapping for one synthetic asset, including N:ATQ = MACE AJ:AUM, not MAP Q:ATT.
- An opt-in Excel test confirms that the template file hash is unchanged.

[✓ Implemented — pure tests, opt-in Excel test and disposable real-input write/read-back passed; template hash unchanged]

### 8. Verify By Reopening Before Success (depends on step 7)

Function `verify_macalc_output(path, expected) -> None`.

- Reopen the saved workbook read-only. For each funded row, compare D:ATQ to the expected values with exact equality (no tolerance), check that J contains no formula, and check that A:C are empty.
- Confirm that Mortgage/ILS rows 2-20 are identical to the template.
- Confirm that every row below the funded block is empty in A:ATQ.

Pass criteria: any mismatch or missing write raises an error and triggers rollback. Success is not recorded until this passes.

[✓ Implemented — exact read-back passed on disposable output; 9 external-link targets and protected formulas preserved]

### 9. Orchestrate With Audit, Guard, Rerun And Rollback (depends on steps 3-8)

Function `run_stage_three(caller) -> Path | None`, mirroring `run_stage_two`:
- single-run guard;
- Running audit saved before any work;
- overwrite confirmation (P14);
- template pop-up stop (P6);
- zero-funded warning that does not stop the run (P4);
- staged write, verify, then atomic move into place;
- Succeeded audit with the output link, recording `SST_Stage3_InputRun` so later stages can detect a stale MACalc;
- Failed/Interrupted audit on error, with prior output restored.

Console-only reporting; no new log files (Decision 14).

Pass criteria: tests cover a declined overwrite, a failure after a partial write, an audit-save failure, a stale Stage 2 and the zero-funded warning path. None of these leaves a half-written output or a false Succeeded. Zero funded ends in Succeeded with the warning shown.

[✓ Implemented — synthetic orchestration tests pass; no financial run performed]

### 10. Install Panel And Button, Then Verify On Real Data (depends on step 9)

Affected files: `src/sst_stage_three.py` (installer), the Draft (backed up first), `references/bonds-lite-setup.md`.

Install `SST_Stage3_Panel` and the button below the Stage 2 controls with occupancy checks, preserving moved Stage 1/2 layouts.

Stage 3 ran on 3 November 2026 after Stage 2 succeeded on handoff `20261102_111001`. The generated workbook passed the orchestrator's read-back verification and the Draft audit records `Succeeded`. The user compared the five funded rows against their expectation and accepted them (Decision 40).

Packaging (P15) happens only after this check and requires separate approval.

[✓ Panel and button installed in the live Draft on 2 November 2026 (backup `references/backups/Draft Input spreadsheet for SST_20261102_171859_041698.xlsm`); Stage 3 run on current handoff and accepted by the user on 3 November 2026; packaging not performed]

## Test Strategy

1. **Pure unit tests (default suite):** block validation, header contract, MACE vector classification and boundaries, MAP/MACE reconciliation, guards, row construction (cashflows from MACE), PV negation, the zero-funded path and Stage 2 readiness. Synthetic data only.
2. **Synthetic workbook tests:** small `.xlsx` files generated at test time exercise the read paths through Excel.
3. **Opt-in Excel tests:**
   - On a disposable copy of `MACalcinPython_v4.0f.xlsb`: write, reopen-verify, confirm the template is unchanged, and check the AutoFilter and `tbl_Assets` remain intact.
   - Panel installation on a disposable Draft copy.
4. **Real-data regression (opt-in):** 091803 MAP plus the central MACE produce exactly the 5 funded ISINs with the values recorded at verification time, kept as a golden fixture of IDs and a few checksum values, not a production-data copy. Exact equality for copied values, per the regression-testing skill's exact tier for pure transfers.
5. **Reference output:** if you provide a hand-built MACalc, its funded rows become the acceptance fixture.

## Affected Files

- `src/sst_stage_three.py` (new): all Stage 3 Slice A logic, orchestration and installer.
- `src/bonds_lite_input.py`: `--stage 3` dispatch and, if needed, public aliases for shared helpers only.
- `tests/test_sst_stage_three.py` (new): unit, synthetic and opt-in Excel tests.
- `tools/build_packaged_stages.py`: include the new module in the controller build (step 10, after approval).
- `Draft Input spreadsheet for SST.xlsm`: Stage 3 button and panel (backed up first).
- `references/bonds-lite-setup.md`, `references/decisions.md`: usage, recovery and decisions.

## Decisions

See `references/decisions.md`: Decision 36 records the user's requirements; Decision 37 records the user's approval of the fail-closed guards; Decisions 38-39 cover Excel handling and the controls; Decision 40 records the user's acceptance of Slice A. Implementation-level Excel handling is summarised in the progress addendum.

## Risks And Unknowns

- **Stale MACalc after a Stage 2 rerun:** Stage 3 records `SST_Stage3_InputRun`. A later Stage 1/2 rerun in a new folder leaves the old MACalc in its own folder, and future slices must check the run identity, as Stage 3 does for Stage 2.
- **Interim `tbl_Assets` size:** `tbl_Assets` depends on COUNTA over column A. With A left empty on funded rows and the rows below cleared, the interim workbook's `tbl_Assets` covers only rows 1-20 until a later slice fills column A. This is expected for an interim output, but MACalc must not be run on it before then.
- **Saved MACE values:** MACE C, D and I are formulas. Reading saved values assumes the central MACE was last calculated and saved correctly for the valuation date.
- **Funded rule depends on zero tails:** the rule treats zeros as compatible with "all negative". If MACE ever stores blanks instead of zeros after maturity, P10 stops the run. Changing that behaviour needs a confirmed rule.
- **Exact equality:** `.xlsb` round-trip of doubles is exact in practice; any failure here is investigated, not tolerated.
- **Stage 2 state:** Stage 2 succeeded on handoff `20261102_111001` on 3 November 2026. A leftover temp folder `.bonds-lite-t4uk4bfv` exists in `111001` and was not touched.
- **Excel lifecycle:** dedicated hidden instances must be closed in `finally`, and the user's Excel instance is never touched (as in Stage 1/2).

## Resolved Questions (2 November 2026)

1. **Cashflows:** use MACE AJ:AUM, not MAP Q:ATT. A funded asset has an all-negative funding-leg vector.
2. **Placement:** clear the rows below the funded block.
3. **Columns A:C:** leave empty; a later slice fills them from another MACE tab.
4. **Output:** save in the timestamped Stage 1 run folder.
5. **Zero funded:** warn the user and continue.

6. **Anomalous MACE vectors and MAP/MACE set differences:** stop and report the issue; both empty means warn and continue (approved).
7. **Output filename/format:** `MACalc input.xlsb` (approved). Changed by Slice D (Decision 54) to `MACalc input_<run ID>_<token>.xlsb`.

## Historical Operational Status (3 November 2026; superseded by Current Release Status)

- Stage 2 succeeded for the current Stage 1 handoff `20261102_111001` on 3 November 2026. Its summary is `bonds_batch_summary_20261103_093358.csv`.
- Stage 3 completed successfully for that handoff. The Draft audit records `Succeeded` and `SST_Stage3_LastSuccessfulOutput` points to `Bonds Lite output/20261102_111001/MACalc input.xlsb`. The user accepted the five funded rows on 3 November 2026.
- The Stage 3 panel and button are installed in the live Draft. The button needs a packaged controller rebuilt with `sst_stage_three` (release v5). Release v4 rejects `--stage 3`. Until then, maintainers can run Stage 3 in source mode (see the session handoff for the command used on 3 November).
- The Stage 3 packaged release has not been built or promoted. Consider building v5 once the next slices are implemented, so colleagues receive one release rather than several.

## Addendum: Implementation Progress

- Implemented Slice A in `src/sst_stage_three.py` and the separate `--stage 3` dispatch. Stage 1/2 behaviour was not intentionally changed. The Stage 3 panel/button installer source is included.
- Added 55 Stage 3 tests, including a disposable Excel regression. The disposable real-input regression used the 091803 MAP workbook, the central MACE workbook and a temporary copy of the MACalc template; it identified the expected five funded assets, wrote and reopened the output, verified exact cell mappings, retained the Mortgage/ILS rows and all 9 external-link targets, and confirmed the template hash was unchanged. Temporary output was removed.
- Excel-specific fixes: start the owned Excel app with a workbook open before setting application properties; validate MAP month ordinals as text; restore and verify external-link paths that Excel rewrites when saving the staged copy. These fixes retain exact formula and output verification.
- Final `python -m pytest -q`: **721 passed, 83 skipped, 2 failed**. The two failures are the pre-existing frozen-`main()` argument mismatch and Stage 2 recovery-location assertion. The `pytest` console entry point has an additional pre-existing `tools` namespace import issue; the focused packaged-builder test passes with `python -m pytest`.
- A disposable real-input round-trip had been verified before the first live run. On 3 November, after Stage 2 succeeded for the current handoff, the Stage 3 orchestrator completed and recorded `Succeeded`. It wrote `Bonds Lite output/20261102_111001/MACalc input.xlsb`; the orchestrator reopened and verified the saved output before success. The user accepted the five funded rows on 3 November 2026 (Decision 40). No package was built or promoted, and no MACalc financial model was run.
- **Stage 3 controls (2 November 2026):**
  - `PrepareMACalc` (calls `LaunchBondsLite 3`) was added to both launcher texts, so Stage 1/2 reinstalls keep it.
  - Button placement now mirrors Stage 2: it sits beside the panel when it fits, otherwise below.
  - Added the `--install-stage-three <Draft>` CLI and an explicit `sst_stage_three` hidden import in the packaging script.
  - An opt-in Excel test installed twice on a disposable Draft copy. It checked placement, `OnAction`, single macro, unchanged Stage 1/2 names and buttons, and idempotency.
  - Live installation result:
    - `SST_Stage3_Panel` is `'I. Control'!$H$100:$O$110`, with fields in L101–L109 and Status set to "Not Run".
    - `btnPrepareMACalc` sits at D100 with the same size as the other buttons.
  - Comparison against the backup:
    - no existing names changed;
    - Stage 1/2 button geometry is unchanged;
    - D44 and the Stage 1/2 audit values are unchanged;
    - all previous VBA lines are retained.
  - `python -m pytest -q`: **721 passed, 83 skipped, 2 failed**. These are the same two pre-existing failures.

## Planning The Next Slices

Slice A is closed. Plan each further Stage 3 transfer as a new lettered slice (Slice B, C, ...) in this document, with its own confirmed requirements, defaults, steps and pass criteria, before any code is written. The user will supply the requirements at the start of the planning session.

### Known Inputs For Later Slices

- **Columns A:C on funded rows** (Aallocations, Ballocations, PartialRecog) are left empty by Slice A. The user said they will be filled from another MACE tab; the tab and mapping are not yet specified.
- **Rows below the funded block** are cleared by Slice A. In the template they held 988 Bond/PP/Infrastructure rows in MAP order, with A:B as formulas linking to an external `202603 - JRL - MAP_Allocations_v3.22.xlsb`. A later slice is expected to repopulate them.
- **`tbl_Assets`** sizes itself by COUNTA over column A, so the interim output covers rows 1-20 only until column A is filled. MACalc must not be run on the interim output.
- **Other workbooks referenced by the Draft:** `SST_MACalc_Central_Template` (P. Parameter!C9) points to a different MACalc file and is not yet used. Confirm whether any later slice needs it.
- **Mortgage/ILS rows 2-20** are never rewritten by Slice A and are verified unchanged against the template.

### Reusable Building Blocks In `src/sst_stage_three.py`

- Readiness and inputs: `resolve_stage_three_inputs` (Stage 2 identity check, path resolution, H44 date guard).
- Template contract: `validate_template_blocks` and the row-1 header checks.
- Readers: `read_map_bonds` (MAP `I. Bonds Data`) and `read_mace_funding` (MACE `C. Assets - Funding`), both read-only bulk reads in a dedicated hidden Excel app.
- Writer and verifier: `write_macalc_output` and `verify_macalc_output` (exact read-back, protected-row comparison, external-link restoration).
- Orchestration: `run_stage_three` (guard, audit, overwrite confirmation, staged write, backup and rollback).

### Questions To Settle For Each New Slice

1. Which source workbook, sheet and header row; which source and destination columns, and any transform (sign, units, values versus formulas).
2. Which rows are in scope and how they are matched (ISIN, row order or another key), and what happens to unmatched or duplicate keys.
3. Whether the slice extends the single Stage 3 button run or needs its own button, audit fields and run-identity check.
4. Which guards stop the run and which only warn.
5. How the slice is verified: exact read-back, reconciliation counts, and any hand-checked reference output.

## Slice B: Asset Block And Allocations

### Overview

Slice B completes `I. Assets`. Directly below the Slice A funded block it writes one row per asset from the Stage 2 MAP output CSV (`*py_bonds_Base_Scenario_nozeros*`). It fills Type, Index Linked? and FSAO from MACE `I. Bonds`, and fills columns A:C on every post-ILS row from the MACE `Control` allocation table (K:O). Slices A and B are **one process**: a single "Prepare MACalc" run produces one `MACalc input.xlsb`, verifies it once and records one audit result. Commit and the v5 release wait until Slice C is accepted (Decision 41).

### Confirmed Requirements (User, 3 November 2026)

1. **Source file:** the Stage 2 MAP output whose file name contains `py_bonds_Base_Scenario_nozeros`. The substring may change with a future Bonds Lite model update, so it must be easy to change in one place.
2. **Placement:** the new block starts on the row immediately after the last funded row (row 21 if there are no funded assets).
3. **CSV to MACalc `I. Assets`:**

   | CSV column (header) | MACalc column (header) | Transform |
   |---|---|---|
   | B (Unique_ISIN) | D (ISIN) | none |
   | A (ISIN) | E (Name) | none |
   | D (IM_Zspreads) | F (Zspread) | × 10000 |
   | F (Sector) | H (Sector) | none |
   | I (Final_Notched_Rating) | I (Rating) | none |
   | E (Currency) | K (Currency) | none |
   | J:ATN (Base_PV, months 0-1199) | M:ATQ (PV, months 1-1200) | none |

4. **MACE `I. Bonds` to MACalc**, matched on MACalc Name (column E): Type → G, Index Linked? → J, FSAO → L.
5. **MACE `Control` table** (columns K:O, variable length) to MACalc, matched on MACalc Name (column E), for every row after the ILS block: Comp. A % → A (Aallocations), Comp. B % → B (Ballocations), Partial CF Flag → C (PartialRecog). Comp. C % (N) is ignored. If the table is not contained within K:O, stop with an error saying so.
6. **Guards:** stop on unmatched or duplicate ISINs, blanks or errors, and list the offending ISINs so the user can investigate.
7. **Optional sanity check:** MACE `Control` valuation date equals the Draft valuation date.

### Verified Discovery (read-only, 3 November 2026, handoff `20261102_111001`)

- **CSV:** `JRL MAP\py_bonds_Base_Scenario_nozeros_BWS_CoeffRefresh.csv`, 988 rows × 1,211 columns. A copy with the same name also sits in `JRL Bonds\`.
  - Headers: A ISIN, B Unique_ISIN, C Proxy_Grouping, D IM_Zspreads, E Currency, F Sector, G Start_Rating, H Final_Rating, I Final_Notched_Rating, J Base_PV, K:ATN months `0`…`1199`, ATO Last_CF.
  - `Unique_ISIN` is `ISIN` plus `_<rating>`. ISIN and Unique_ISIN are both unique and never blank.
  - J:ATN has no blank or non-numeric cells, and the mapped descriptive columns have no blanks.
  - CSV month `0` aligns with MACalc month 1 (N), so J:ATN → M:ATQ is a one-to-one copy and `Last_CF` is excluded.
  - `IM_Zspreads` ranges from -0.0553 to 0.2797 (decimal), so × 10000 gives basis points (-553 to 2,797 bps).
- **MACE `I. Bonds`:**
  - **Column D is the ISIN.** It is unique and contains all 988 CSV ISINs.
  - Column E is the company name, which is not unique and is not the matching key.
  - PV, Currency, Sector and Rating agree exactly with the CSV.
  - On 246 rows the cashflow timing and Z-spread differ from the CSV, so `I. Bonds` cannot be a parity source for those columns.
  - Extra rows: only Inf_swap rows, which are not in the CSV.
- **MACE `Control` table:**
  - Headers are on row 9: **K ISIN**, L Comp. A %, M Comp. B %, N Comp. C %, O Partial CF Flag.
  - It has 996 data rows from row 10: the 988 CSV ISINs (including the 5 funded ones) plus 8 Inf_swap rows. There are no duplicates.
  - Six Partial CF Flags are tiny floats (about 1e-16) rather than exactly 0.
  - A neighbouring LTM-notes table uses column J, so "contained in K:O" cannot require J to be empty.
- **Valuation date:** MACE `Control!C9` = Draft `I_Valuation_Date` (`I. Control!D11`) = 30/06/2026.
- **Overlap with Slice A:** all 5 funded assets also appear in the CSV. Each therefore gets a funded-leg row (D = ISIN) from Slice A and an asset row (D = Unique_ISIN) from Slice B. Column D stays unique.

### Interpretation Of The Requirements (confirmed by the user, 3 November 2026, Decision 44)

- **I1 – `I. Bonds` key:** MACalc E holds the CSV ISIN, so the lookup matches it against `I. Bonds` column **D (ISIN)**. Column E there holds company names and would not match.
- **I2 – `Control` key:** the requirement names "ISIN (Column L)", but the live header row has **ISIN in K**, with Comp. A % in L, Comp. B % in M and Partial CF Flag in O. The plan uses the K key and the L → A, M → B, O → C mapping. Columns are located by header text on row 9, and any drift stops the run.
- **I3 – Containment rule for K:O:**
  - Row 9 must hold exactly the five headers above.
  - Data runs from row 10 down to the first blank in K.
  - The run stops if any cell in column P between row 9 and the end of the table is non-blank, or if any non-blank K appears in the 50 rows below the gap. Columns from Q onwards hold other MACE panels and are not checked (Decision 48; originally P:Z).
  - The error message says the table is not contained within K:O.
- **I4 – Unmatched or duplicate:** this means any CSV ISIN that is missing from `I. Bonds` D or `Control` K, or appears more than once in the CSV, `I. Bonds` D or `Control` K. The run stops and lists every offending ISIN by category. Matching is one-way: MACE rows that are not in the CSV are allowed (Q-B3).
- **I5 – Blanks/errors:**
  - The run stops on any blank, error or non-numeric value in:
    - CSV B, A, D, F, I, E and J:ATN;
    - `I. Bonds` Type, FSAO and Index Linked? for matched rows;
    - `Control` L, M and O for matched rows.
  - Text columns must be non-blank text, and numeric columns must be finite numbers.
  - A legitimate 0 is allowed.

### Confirmed Answers (User, 3 November 2026)

| ID | Topic | Confirmed rule |
|---|---|---|
| Q-B1 | Columns A:C on the funded rows | **Confirmed.** After the run, A:C are filled on every post-ILS row, funded and asset rows alike. Slice A left them empty only so they could be populated together. The funded rows take their A:C from the `Control` table, matched on Name (E). This replaces P2. |
| Q-B2 | Locating the CSV | **Confirmed.** Search only the entity's MAP output folder from Stage 2. That is the `output_folder_path` validated by `bonds_lite_input.read_model_inputs` for the MAP dataset, `<run folder>\<entity> MAP\` (today `JRL MAP\`). Do not search recursively or in the Bonds folder. Stop if zero or more than one file matches. The substring is a single module constant, `MAP_BASE_OUTPUT_PATTERN = "py_bonds_Base_Scenario_nozeros"`, in `src/sst_stage_three.py`. |
| Q-B3 | MACE rows not in the CSV | **Confirmed.** Allowed. Print their count to the console and do not stop. Any CSV asset missing from MACE `I. Bonds` or `Control` stops the run (I4). |
| Q-B4 | Extra reconciliation | **Confirmed.** If the CSV ISIN set differs from the Stage 1 MAP `I. Bonds Data` set, **warn** and continue, listing the differences in the console and a non-blocking pop-up at the end, like P4. **Stop** if MACE `Control!C9` differs from Draft `I_Valuation_Date`. Do **not** compare CSV PV with MACE PV. |
| Q-B5 | Partial CF Flag float noise | **Confirmed.** Copy exactly, with no rounding. |
| Q-B6 | Row order and clearing | **Confirmed.** Use CSV row order. Slice C does not touch `I. Assets`, so row order has no downstream dependency. After the funded and asset blocks are written, everything in A:ATQ below the last asset row must be empty, and this is verified on read-back. |
| Q-B7 | Columns D and E | **Confirmed.** Write `Unique_ISIN` to D and the CSV `ISIN` to E, as specified. The funded rows keep their plain ISIN in D, so D stays unique across both blocks. Updating `tbl_ForcedCompAperc`, `tbl_MinCompAperc` and `lst_pre16assets` to match is Slice C. |

### Steps

#### B1. Extend The Test Contract First

Affected file: `tests/test_sst_stage_three.py`.

Write failing tests with synthetic data only. Use small CSVs and a 1,201-value vector generated with NumPy, plus small `.xlsx` workbooks generated in code. The tests cover:
- CSV discovery: none, one or several matches; a match in the Bonds folder only is not found; the folder is taken from `read_model_inputs`;
- CSV header contract: drift in A, B, D, E, F, I or J, and month headers that are not `0`…`1199`;
- the ×10000 Zspread transform and the exact J:ATN → M:ATQ copy;
- each I4 and I5 guard, with the offending ISINs named in the message;
- a MACE-only row is allowed, but a CSV-only asset stops the run;
- the I3 containment cases: data in P, a second block below a gap, and header drift;
- a CSV/MAP set difference warns and does not stop; a valuation-date mismatch stops;
- a funded asset appearing in both blocks;
- zero funded assets, where the Slice B block starts at row 21.

Pass criteria: the new tests fail only because the implementation is missing, and the existing 60 Stage 3 tests stay green.

[✓ Implemented — see addendum]

#### B2. Locate And Read The Base Scenario CSV (depends on B1)

Function `find_map_base_output(map_output_folder: Path) -> Path` and `read_map_base_output(path: Path) -> MapBaseOutput` (frozen dataclass).

- Get `map_output_folder` from `bonds_lite_input.read_model_inputs(run_folder)["MAP"]`, which is the same validated entity folder Stage 2 used. Match `*{MAP_BASE_OUTPUT_PATTERN}*.csv` directly in that folder, without recursing.
- Read the CSV with pandas. No Excel is needed, which keeps this step fast.
- Verify headers by name at the expected letters, and require months `0`…`1199` in K:ATN.
- Apply the I5 guards to the CSV columns. Require unique ISIN and Unique_ISIN.
- Keep the rows in file order.

Pass criteria: the synthetic tests pass. On `20261102_111001` the reader returns 988 rows with no guard failures.

[✓ Implemented — see addendum]

#### B3. Read MACE `I. Bonds` And The `Control` Allocation Table (parallel with B2)

Functions `read_mace_bond_attributes(book) -> dict[str, BondAttributes]` and `read_mace_allocations(book) -> dict[str, Allocation]`.

- Reuse the read-only MACE session that `read_mace_funding` opens: links not updated, no recalculation and saved values (P9). MACE is opened once per run.
- `I. Bonds`: locate the ISIN (D), Type, Index Linked? and FSAO columns by header text. Confirm the header row and labels during implementation against the live file, and record them in the addendum.
- `Control`: apply I2 and I3. Read K:O in one bulk read and also return `valuation_date` (C9).
- Detect duplicate keys at read time (I4).

Pass criteria: the synthetic workbook tests pass. On the live MACE, the `Control` table has 996 rows, the `I. Bonds` keys include all 988 CSV ISINs, and there are no duplicates.

[✓ Implemented — see addendum]

#### B4. Reconcile And Build The Slice B Block (depends on B2, B3)

Functions `reconcile_slice_b(...) -> list[str]` (returns warnings) and `build_asset_rows(...) -> list[list[object]]`, both pure.

- Apply the I4 and I5 stops, plus a stop on a valuation-date mismatch. Collect every problem before raising, so a single message lists all offending ISINs grouped by category.
- Return the warnings rather than raising them: the CSV/MAP set difference (Q-B4) and the count of MACE-only rows (Q-B3).
- Build each row in A:ATQ order:
  - A:C from `Control`;
  - D, E, F (× 10000), H, I, K and M:ATQ from the CSV;
  - G, J and L from `I. Bonds`.
- Write values only.

Pass criteria: a per-cell test for one synthetic asset covers every column, including F = IM_Zspreads × 10000 exactly as a double and M = Base_PV.

[✓ Implemented — see addendum]

#### B5. Fill A:C On Funded Rows (depends on B3)

Change `build_funded_rows` so that A:C come from the `Control` table, matched on the funded ISIN in column E.

The Slice A guards still apply. A funded ISIN with no `Control` row stops the run.

Pass criteria: the updated Slice A tests assert populated A:C, and all other funded-row cells are unchanged.

[✓ Implemented — see addendum]

#### B6. Write And Verify As One Output (depends on B4, B5)

Extend `write_macalc_output` and `verify_macalc_output`. Do not add a second writer.

- After the existing clear, write the funded block and then the asset block as two bulk range assignments into the same staged copy, and save once.
- Exact read-back covers A:ATQ for every funded and asset row. Mortgage/ILS rows 2-20 must be unchanged, and every row below the asset block must be empty.
- `tbl_Assets` (COUNTA over column A) should now cover rows 1 to the last asset row. Verify its resolved size equals 20 + funded + assets.
- Keep the external-link restoration from Decision 38.
- Before saving, check two MACalc conditions (see the Slice B risks):
  - every row from 2 to the last asset row has a non-blank A;
  - every G value is in the MACalc Type set (`Mortgage`, `Bond`, `Infrastructure`, `PP`, `Swap`, `ILS`).

  If either fails, stop and list the offending rows.

Pass criteria:
- The opt-in Excel test on a disposable template copy passes, and the template hash is unchanged.
- On live data the output has 5 funded rows (21-25) and 988 asset rows (26-1013). A:C are populated on every post-ILS row, and `tbl_Assets` covers rows 1-1013.

[✓ Implemented — see addendum]

#### B7. Orchestrate As One Run (depends on B6)

Extend `run_stage_three`. The existing guard, Running audit, overwrite confirmation, backup/rollback and Succeeded/Failed audit cover both slices.

- Resolve the CSV path in `resolve_stage_three_inputs`, so a missing file stops the run before any write.
- Print the CSV path, row counts and the MACE-only count to the console. Show the CSV/MAP difference warning in the console and in a non-blocking pop-up at the end. No new audit fields are added (Decision 14).
- Leave the panel, button and VBA unchanged, so no reinstall is needed.

Pass criteria:
- The orchestration tests cover a Slice B guard failure after the Running audit (no output, Failed audit, previous output restored).
- One live source-mode run on handoff `20261102_111001` records Succeeded. The user then checks a sample of asset rows and A:C values.

[Code and orchestration tests implemented — live run pending the containment decision; see addendum]

### Test Strategy (Slice B)

1. **Pure tests:** CSV header contract, transforms, guards, reconciliation, row construction and K:O containment.
2. **Synthetic workbook tests:** `I. Bonds` and `Control` readers on generated `.xlsx` files.
3. **Opt-in Excel and real-data regression:** a disposable template copy with the 111001 CSV and the central MACE gives 5 + 988 rows with exact read-back.
   - Keep a golden fixture of IDs and checksums only: row count, sum of M, sum of F, the first and last Unique_ISIN, and the sums of A and B. Do not copy production data.
   - Use exact equality for copied cells and for F (×10000 is deterministic on the same double).
4. **Parity check (informative):** compare the old template rows 21-1008 with the new asset block on ISIN/Name, A:C and G/J/L. The template predates this quarter, so differences are reported, not asserted.

### Affected Files

- `src/sst_stage_three.py`: the new readers, reconciliation, row builder, the `MAP_BASE_OUTPUT_PATTERN` constant, and extended write/verify/orchestration.
- `tests/test_sst_stage_three.py`: the Slice B tests and updated Slice A A:C expectations.
- `references/bonds-lite-setup.md`: the Stage 3 section, covering the CSV dependency, the changeable substring and the new error messages.
- `references/decisions.md`: Decisions 41-43.
- `src/bonds_lite_input.py`, the Draft and the packaging: **no change**.

### Risks And Unknowns

- **CSV name change:** a Bonds Lite model update may rename the CSV. The single constant plus the "zero or several matches" stop fails safely, rather than reading the wrong file.
- **Two copies of the CSV:** the copies in `JRL MAP\` and `JRL Bonds\` share a name. Searching only the entity MAP folder that Stage 2 validated avoids reading the Bonds copy.
- **Z-spread, cashflow and PV differences between the CSV and MACE `I. Bonds`:** the user specified the CSV as the source and confirmed that these differences are ignored (Q-B4).
- **CSV and Stage 1 MAP sets differ:** this only warns, so a CSV from a mismatched run could get through if MACE still contains all its assets. The warning lists the differences so the user can judge.
- **`tbl_Assets` growth (checked 3 November 2026, safe):**
  - MACalc reads assets with `wb.names('tbl_Assets').refers_to_range.options(pd.DataFrame, header=1)` (`MACalc.py` line 431, `Release Preparation\MACalc\v_3_3`). This is the closest source to the frozen `MACalcrunmanager_v4.0f` that the template calls.
  - Every later use is sized by `len(df_assets)`, and no worksheet formula refers to `tbl_Assets` or `I. Assets`. So 1,013 rows is fine; there is no fixed limit of 1,008.
  - The `_FilterDatabase` range already reaches row 1056.
  - Two conditions follow, and both become guards in step B6:
    - **Column A must be non-blank on every row 2..last.** `tbl_Assets` is `COUNTA(A:A)` rows deep, so a blank A silently drops the last rows. Any stray value below the block adds rows. A legitimate 0 is still counted.
    - **Type (G) must be one of `Mortgage`, `Bond` (any case), `Infrastructure`, `PP`, `Swap` or `ILS`, spelled exactly** (MACalc lines 576 and 918-949). Otherwise MACalc's assertion that the asset groups sum to the total fails. The funded rows are `Infrastructure`, and the template's 988 rows were 634 Bond, 177 PP and 177 Infrastructure.
- **ISIN format in column D (confirmed, Q-B7, Decision 45):**
  - In the template, D held the CSV `ISIN` and E the company name.
  - MACalc joins D to `tbl_ForcedCompAperc`, `tbl_MinCompAperc` and `lst_pre16assets` (MACalc lines 1264-1289). Today these lists hold plain ISINs.
  - The user chose to keep **`Unique_ISIN` → D** and **CSV `ISIN` → E**, as specified.
    - The funded rows (D = ISIN) are meant to differ from their rating-suffixed counterparts in the asset block.
    - E (the CSV ISIN) is also a matching key.
  - Bringing those three lists in line with the new D values is the scope of **Slice C**. Until Slice C is done, the interim `MACalc input.xlsb` must not be used for optimisation runs, because the forced and minimum Comp A constraints and the pre-2016 flags would not match.
- **Runtime:** the extra MACE reads use the already-open session. The extra write is roughly 988 × 1,213 cells in one bulk assignment. Expect the 86 s run time to rise, and measure it before optimising (Decision 13).
- **Slice C does not touch `I. Assets`:** the run ensures that everything below the last asset row is empty.

## Skills Consulted

Testing-strategy, regression-testing, coding-standards, document-extraction, excel-analysis, decision-logging and git-workflow. No team-specific conventions are active (no `.github/supir-agents.lock`).

## Addendum: Slice B Implementation Progress

Session of 3 November 2026. Steps B1-B7 are implemented, and the live source-mode run on handoff `20261102_111001` was accepted by the user (Decisions 47 and 48 approved). Nothing is committed, built or promoted.

- **Tests:** `python -m pytest -q` gives 797 passed, 84 skipped and the 2 known pre-existing `test_bonds_lite_input.py` failures. The opt-in empty-block Excel test passes, and a new opt-in real-data regression (`test_excel_slice_b_real_data_regression_on_handoff_111001`) holds IDs and checksums only.
- **Live facts recorded during implementation:**
  - MACE `I. Bonds`: headers on row 7 (D ISIN, G Type, J Index Linked?, L FSAO); data rows 8-1003; 996 unique keys; Index Linked? is text such as `FXD` or `RPI(-99.0,99.0)`.
  - MACE `Control`: C9 = 30/06/2026; K9:O9 headers (`Comp. A % ` has a trailing space, so headers are trimmed); 996 rows (10-1005); no duplicates; 8 MACE-only rows in each tab.
  - CSV: `JRL MAP\py_bonds_Base_Scenario_nozeros_BWS_CoeffRefresh.csv`, 988 rows, read in 0.23 s.
- **Disposable end-to-end run** (real template copy, 111001 CSV and MAP, central MACE, side check limited to column P): 5 funded rows (21-25) and 988 asset rows (26-1013); exact A:ATQ read-back passed; `tbl_Assets` = rows 1-1013; template hash unchanged; no CSV/MAP difference; 136 s.
- **Deviations** (Decision 47):
  - The CSV is read with the standard-library `csv` module, not pandas, because pandas is excluded from the packaged controller.
  - The `I. Bonds` and `Control` reader tests use fake sheets and pure parsers rather than generated `.xlsx` files.
  - The readers return dataclasses that carry the duplicate keys, so reconciliation can report them all in one message.
  - `reconcile_slice_b` returns a `SliceBReconciliation` (CSV/MAP differences and MACE-only counts) instead of a list of warning strings.
  - Pre-existing Slice A bug fixed: a failure before publishing deleted the previous output. Now only an output published by the failing run is removed.
- The disposable run above is a rehearsal only. It used a relaxed side check and is **not** B7 evidence. B7 still needs the live source-mode run under the rule the user approves.
- **Resolved — containment rule I3:** the user chose option (b), so the side check now covers column P only (Decision 48; `CONTROL_SIDE_COLUMNS = ("P", "P")`, with new tests).
- **B7 live source-mode run (3 November, 11:53-11:54, 112 s):** handoff `20261102_111001`, central MACE v5.20 c2. Audit `Succeeded`; output `Bonds Lite output/20261102_111001/MACalc input.xlsb`; rows 2-20 template Mortgage/ILS, 5 funded rows (21-25), 988 asset rows (26-1013), rows 1014+ empty; Types: 634 Bond, 182 Infrastructure, 177 PP, 11 Mortgage, 8 ILS. Accepted by the user on 3 November (row 285's C = 2.19e-16 copied exactly from MACE, left as is). The original blocker was: on 3 November, `read_mace_inputs` on the central MACE (`202606 - JRL - MACE v5.20 c2.xlsb`) stops with "not contained within K:O". It found data in T9, U9, T10, U10, Q12, R12, T12, U12, Q13, R13, Q14, R14, T14, Q15, R15, Q16, R16, Q18, T19-T22, U22, Q23, T23, Q24, Q25, T25, Q26, R26, T26, Q27, T27, T28, T29, U29, Q31 and T31. Column P is empty. A live Stage 3 run would fail at this step, before any output is written. Changing the rule needs the user's approval (it would update Decisions 42/45).

## Slice C: Align MACalc ISIN Lists With Column D

**Status: plan approved by the user on 3 November; implemented on 3 November (see "Addendum: Slice C Implementation Progress"). The live C6 run succeeded on 3 November with the expected summary. Decision 53 approved and Slice C accepted by the user on 3 November.** Decisions 49-52 record the confirmed mapping, contents and duplicate handling.

### Overview

Slice B writes `Unique_ISIN` (ISIN plus `_<rating>`) to `I. Assets` column D for every asset row (Decision 45). MACalc joins D to three hard-coded ISIN lists in the template. Today none of the 106 matching `tbl_MinCompAperc` entries and none of the 31 matching `lst_pre16assets` entries match the output. Only 19 of the 24 `tbl_ForcedCompAperc` entries match (the Mortgage and ILS rows).

Slice C rewrites the key column of each list in the staged MACalc copy, so that every entry that matched an asset in the template matches the same asset's new D value. It does not touch `I. Assets`. It runs inside the existing single Stage 3 run: one staged write, one save, one read-back and one audit result.

### Requirements (user brief, 3 November 2026)

- Align `tbl_ForcedCompAperc`, `tbl_MinCompAperc` and `lst_pre16assets` in the MACalc output with the rating-suffixed ISINs in `I. Assets` column D.
- Do not touch `I. Assets`.
- Run in the same single Stage 3 run as Slices A and B.
- The user left four points open: the layouts, the mapping rule, unmatched and multiple matches, and other contents. They are answered from discovery below, with defaults marked PENDING.

### Verified Discovery (read-only, 3 November 2026, template `MACalcinPython_v4.0f.xlsb` and output `20261102_111001\MACalc input.xlsb`)

**Where the lists live.**

All five ISIN lists are workbook-scoped dynamic names of the form `OFFSET(<first cell>,0,0,COUNTA(<column>)-1,<width>)`. Each has a title in row 5 (the `-1` in COUNTA discounts it), row 6 blank, and data from row 7. Every entry is a hard-coded value, and no formula anywhere refers to the names. The output is cell-for-cell identical to the template here: Slices A and B do not touch these sheets.

| Name | Sheet | Range today | Entries | Columns | Neighbouring columns |
|---|---|---|---|---|---|
| `tbl_ForcedCompAperc` | `I. Inputs` | `H7:I30` | 24 | H key, I Comp A % (0-1) | G and J empty |
| `tbl_MinCompAperc` | `I. Inputs` | `K7:L170` | 164 | K key, L min Comp A % (0-1) | J and M empty |
| `lst_pre16assets` | `P. Parameters` | `K7:K328` | 322 | K key | J and L empty |
| `lst_REITs` | `P. Parameters` | `M7:M22` | 16 | M key | L empty |
| `lst_noPRoC` | `I. Inputs` | `N7:N14` | 8 | N key | M empty |

**How MACalc uses them** (`Release Preparation\MACalc\v_3_3`). `lst_ISINs` and the joins are on `df_assets['ISIN']`, which is `I. Assets` column D.

- `tbl_ForcedCompAperc` (`MACalc.py` 1275-1283): left-merged with `df_assets['ISIN']`. A match forces Comp A to the given %; no match means unforced.
- `tbl_MinCompAperc` (1286-1292): same left-merge; sets the minimum Comp A % in the optimiser.
- **A duplicated key in either table would duplicate rows in the merge.** The resulting arrays would no longer align with `df_assets`. Duplicate keys are therefore a hard stop.
- `lst_pre16assets` (1262-1269, 1296-1297, 1347): membership test; pre-2016 assets cannot be sold and are forced to 100% Comp A in TMTP optimisation.
- **`lst_REITs` (`calcSCR.py` 32-40, 56, 130): membership test on the same ISIN; it reclassifies REITs as Financial for the SCR.** It affects every MACalc run, not just optimisation. It is not one of the three lists in the brief (Q-C1).
- `lst_noPRoC`: not read by any `.py` in `v_3_3`. Its 8 entries are the ILS D values, which Stage 3 never changes.

**Matching against the 111001 output.**

- Rows 2-12 are Mortgage (D = E = `Note 1`…`Note 11`). Rows 13-20 are ILS (D = `Inf_swap_…`, E = `Inf_swap`, `Inf_swap_collateral_fixed` or `…_floating`, so E is **not** unique there). Rows 21-25 are funded (D = E = plain ISIN). Rows 26-1013 are assets (D = `Unique_ISIN`, E = CSV ISIN).
- Across rows 26-1013, `D == E + "_" + rating` on every row, with ratings `GOV`, `AAA`, `AA`, `A`, `BBB`, `BB` and `B`. E is unique within the asset block, and D is unique across rows 2-1013.

| List | Exact match on D today | Match via asset E (to remap) | Unmatched (stale) | Matched in template D | Matched after remap |
|---|---|---|---|---|---|
| `tbl_ForcedCompAperc` | 19 (11 Notes, 8 Inf_swap) | 5 (`IL0000000111_JRABACUS_Just Pvt`, 4 × `GB00…_JRUSPP_Just`) | 0 | 24 | 24 |
| `tbl_MinCompAperc` | 0 | 106 | 58 | 106 | 106 |
| `lst_pre16assets` | 0 | 31 | 291 | 31 | 31 |
| `lst_REITs` | 0 | 16 | 0 | 16 | 16 |
| `lst_noPRoC` | 8 | 0 | 0 | 8 | 8 |

- The unmatched entries were **already unmatched in the template**. They are stale entries for matured or sold assets, for example `Note1_XS0138717441` and plain `XS…` codes without the `_JR…` suffix. MACalc ignores them today.
- One `lst_pre16assets` near miss: `XS0306647792` is unmatched, while the CSV holds `XS0306647792_JRABACUS_Just`. It did not match the template either.
- No list entry is a funded asset's ISIN, so today the "one plain ISIN matches a funded row and an asset row" case does not arise. It can arise in a future quarter (Q-C3).
- No list has duplicate keys. All values in the two tables are floats in [0, 1].

### Interpretation And Defaults (PENDING, Decision 49)

- **C1 – Mapping rule (one rule for every case).** For each entry key `k`, in list order:
  - build the **matching row set** `M(k)` = {rows 2..last whose D equals `k`} ∪ {**asset-block** rows whose E equals `k`}. Because D is unique (Slice B), the first part has at most one row;
  - emit the D value of **every** row in `M(k)`, in `I. Assets` row order (top to bottom), each with `k`'s Comp A value, at `k`'s position in the list;
  - so |M| = 0 → unmatched (C2); |M| = 1 → kept (D match) or remapped (E match); |M| ≥ 2 → fan-out (C3).
  - Matching is exact (case- and space-sensitive, as MACalc's join is).
  - E is not used for Mortgage/ILS rows, because ILS E values are not unique.
  - **Only rating-suffixed rows are reached through E (user, 3 Nov).** An asset row joins `M(k)` through its E only when its D is exactly `E + "_" + <that row's own CSV rating>` (`MapBaseRow.rating`). A row whose D equals its E (not suffixed) is already matched on D, so its entry is kept unchanged. If no asset row is suffixed, Slice C changes nothing: every list is written back exactly as in the template. Any other shape (D differs from E but is not `E_<rating>`) stops the run (Guard 7).
  - The output therefore restores precisely the template's matches (table above). With today's data every entry has |M| ≤ 1.
- **C2 – Unmatched entries:** keep them unchanged, in place. MACalc ignores them, as it does today, and the template content is preserved. Print a per-list count to the console, plus the list of entries when there are 50 or fewer. Do not stop and do not show a pop-up.
- **C3 – One key matching several rows (|M| ≥ 2): fan out deterministically.** Emit one entry per matching row, contiguously at `k`'s original position, ordered by `I. Assets` row number, all with `k`'s value. The restriction then applies to every leg that shared the plain ISIN, as it did in the template. Two shapes can produce this:
  - (a) `k` is a funded row's plain ISIN that is also an asset E. Funded rows (21-25) precede the asset block, so this gives `k` first, then the `Unique_ISIN`(s);
  - (b) several asset rows share the same E, for example the same ISIN at two ratings in a future CSV (`X_A` on row 40, `X_BBB` on row 31). This gives `X_BBB` then `X_A` (row order, not alphabetical or rating order).
  - Both together give `k` followed by every asset leg in row order.
  - "Added" for the entry is |M| − 1. Each fan-out is listed on the console (entry → emitted keys). It does not stop the run.
  - Zero cases today: E is unique within the asset block on 111001 and no entry is a funded ISIN, so Added = 0 for every list.
  - If a fan-out produces a key that another entry also produces (for example the source list holds both `X` and `X_BBB`), Guard 5 applies: the duplicate is silently removed in a membership list, and a table follows the Q-C5 table answer.
- **C4 – Order and size:** keep the original order. The list grows only by C3 additions, so its new size = old size + Σ(|M| − 1) over entries with |M| ≥ 2. Write the whole key/value block back from row 7 and clear any old rows below it, so the range stays contiguous.
- **C5 – Values:** copy Comp A values exactly, without rounding, so the float32 artefacts such as `0.10000000149011612` are preserved, as in Q-B5.
- **C6 – No near-miss fuzzy matching.** `XS0306647792` stays unmatched; it is reported in the C2 console list.
- **C7 – No new coverage rule:** assets that appear in no list remain unrestricted, as in the template.

### Guards

**Stop** (raise before saving; list offending entries, grouped as in Slice B):

1. A list's title cell (row 5) or blank row 6 has drifted. The titles are `Forced Component A allocation percentages`, `Minimum Component A allocation percentages`, `Pre-16 asset ISINs` and (if in scope) `REITs (classified as Fin for SCR)`.
2. The name no longer resolves to `OFFSET(<sheet>!$<col>$7,0,0,COUNTA(<sheet>!$<col>:$<col>)-1,<width>)` on the expected sheet and column.
3. A blank or error key inside the list range, or a non-blank key below the first blank (COUNTA would misalign the range).
4. A blank, error, non-numeric or out-of-[0,1] value in a table's value column for any entry.
5. **Tables** (`tbl_ForcedCompAperc`, `tbl_MinCompAperc`): remove duplicate keys silently, keeping the first occurrence when their values are exactly equal; stop and report all conflicting values when duplicate keys have different values. This applies to source duplicates and duplicates created by remapping/fan-out. **Membership lists** (`lst_pre16assets`, `lst_REITs`): remove duplicate keys silently, keeping the first occurrence. See Decisions 49 and 52.
6. Data in the empty neighbouring columns listed in the discovery table, on rows 5 to the list end. This is the same "not contained" idea as Decision 48, and it catches a widened table.
7. An asset-block row whose D is neither its E nor `E_<its rating>`. Slice B should never produce this, so it means the suffix rule has changed and the remap could hit the wrong row.

**Warn (console only):** unmatched counts (C2) and C3 fan-outs (entry → emitted keys, in row order).

### Open Questions For The User

| ID | Question | Default (PENDING) |
|---|---|---|
| Q-C1 | Include `lst_REITs`? It is keyed on D and drives the SCR Financial reclassification in `calcSCR.py`, so with today's output 16 REITs would be treated as non-financial in every MACalc run. | **Confirmed 3 Nov: include** (Decision 50) |
| Q-C2 | `lst_noPRoC` (8 Inf_swap entries, not read by `v_3_3`): leave it, with read-back only? | **Confirmed 3 Nov: leave; verify unchanged** |
| Q-C3 | One entry matching several rows: a funded plain ISIN that is also an asset's E, **or** several asset rows sharing one E (same ISIN at different ratings). Fan out to one entry per matching row in `I. Assets` row order (C3), or stop the run? (A third option, keeping only the `Unique_ISIN`(s), is rejected for the funded overlap because it would silently drop the funded leg.) | **Confirmed 3 Nov: fan out** every list entry to all matching rating-suffixed D values, in row order, same value; only when the asset D values are actually rating-suffixed |
| Q-C4 | Stale unmatched entries (58 + 291): keep in place (C2), or remove from the output? | **Confirmed 3 Nov: keep**; console count |
| Q-C5 | How should duplicate keys be handled in the membership lists and Comp A tables, including duplicates created by fan-out? | **Confirmed 3 Nov: membership-list duplicates are silently removed, keeping the first. Table duplicates with identical values are silently removed, keeping the first; duplicates with conflicting values stop the run.** |
| Q-C6 | Is `MACalc.py` in `v_3_3` the right reference for the `MACalcrunmanager_v4.0f` the template calls (as assumed in Slice B)? | **Answered 3 Nov: no.** Stage 5 is the **run** of the Python MACalc in `00 Ad Hoc\MACalc_Batch` on the MACalc input once Stage 3 has validated it. Stage 3 does not read, change or run that model. Slice C is specified against the template's contract (lists keyed on `I. Assets` D, exact match); Stage 5 planning confirms that contract before the first run (Decision 51) |

### Steps

#### C0. Hand-Over To Stage 5 (Decision 51) [✓ Implemented — see addendum]

Stage 3 does not open, change or run the Python MACalc in `C:\Users\Joao Branco\OneDrive - Just Group\Financial Resilience Team - General\00 Ad Hoc\MACalc_Batch`. Its job ends with a validated `MACalc input.xlsb`.

**Stage 5 is the run of that model** on the validated input. It is a later stage, planned separately. Before its first run, Stage 5 planning must also confirm, for each list:
- that the model reads it from the same sheet and name;
- that it joins the list to `I. Assets` D by exact match;
- how it treats duplicate table keys.

It must also check that no other D-keyed list exists. Any gap found there reopens Slice C. This step does not block C1.

#### C1. Extend The Test Contract First [✓ Implemented — see addendum]

Affected file: `tests/test_sst_stage_three.py`. Use synthetic data only.

Write failing tests for a pure `align_isin_lists(lists, asset_keys) -> IsinListAlignment`. Here `asset_keys` holds the `(D, E, block)` triples for rows 2..last, built from the template rows 2..ILS end plus the in-memory funded and asset rows. The tests cover:

- an exact D match is kept: a Note, an Inf_swap and a funded plain ISIN;
- an asset E match is remapped to its `Unique_ISIN` with the value unchanged;
- an ILS E value such as `Inf_swap` is **not** used as a key;
- an unmatched entry is kept in place and counted;
- C3 fan-out (a): a funded plain ISIN that is also an asset E gives `k` then the `Unique_ISIN`, both with the entry's value, Added = 1;
- C3 fan-out (b): two asset rows with the same E, supplied with `X_BBB` on an earlier row than `X_A`, give exactly `[X_BBB, X_A]` (row order, not sorted), both with the entry's value, at the entry's original position between its neighbours, Added = 1;
- C3 (a)+(b) together: funded `X` plus two asset legs give three entries in row order, Added = 2;
- determinism: the same inputs with the list entries or asset rows shuffled give output that depends only on list order and row order (run twice, identical result);
- same-value duplicates in either Comp A table (in the source or after fan-out) are removed silently, keeping the first; duplicates with conflicting values stop under Guard 5 and name the key and all values;
- in a membership list a source duplicate or a fan-out collision is removed silently, keeping the first occurrence;
- before duplicate removal, list size after mapping equals size before + total Added; final size equals that count minus duplicates removed;
- order is preserved;
- exact float copy (`0.10000000149011612`);
- each stop in Guards 3, 4 and 7, with the offending keys or rows named in the message; Guard 5 stops only for conflicting table values and names the key and values;
- an empty list is allowed and stays empty;
- no suffixed asset rows (D = E everywhere): every list is returned exactly as given (no remap, no fan-out);
- mixed: a suffixed row is remapped, an unsuffixed row's entry is kept;
- Guard 7: an asset row with D = `X_FOO` where E = `X` and the rating is `A` stops the run, naming the row.

Also write failing parser tests, on fake sheets in the Decision 47 style, for Guards 1, 2 and 6 (title drift, name formula drift, a stray value beside or below the list).

Pass criteria: the new tests fail only because the code is missing, and the existing Stage 3 tests stay green.

#### C2. Read The Lists From The Template (depends on C1) [✓ Implemented — see addendum]

Add `read_isin_lists(book) -> TemplateIsinLists`, with the list specification held as one module constant, `MACALC_ISIN_LISTS`: name, sheet, key column, value column (if any), title, and neighbouring columns.

- Call it from `_read_template_snapshot`, so it reuses the already-open read-only template; there is no new Excel open.
- Read each column from row 5 to the last used row in one bulk read, and apply Guards 1-4 and 6.
- Check each name's `RefersTo` text against the expected OFFSET form (Guard 2).
- Add the lists to `TemplateSnapshot`.

Pass criteria: the parser tests pass. On the real template it returns 24, 164, 322 and 16 entries (and 8 for `lst_noPRoC` if it is read for verification).

#### C3. Align (depends on C2) [✓ Implemented — see addendum]

Implement `align_isin_lists` (pure). It builds `M(k)` per C1 and emits D values in row order. It returns the new key/value columns per list and a summary: kept, remapped, unmatched and added counts, the unmatched keys, fan-outs (entry → emitted keys), and de-duplicated counts. It silently keeps the first same-value table duplicate, silently keeps the first membership-list duplicate, and raises on Guard 5 only for conflicting Comp A values.

Pass criteria: the C1 pure tests pass, including all the C3 fan-out, ordering and duplicate-handling tests. Before de-duplication, every emitted key is the D of a row in its entry's `M(k)`, and every row in `M(k)` is emitted exactly once; afterward, duplicate keys follow Decision 52.

#### C4. Write In The Same Staged Book (depends on C3) [✓ Implemented — see addendum]

Extend `write_macalc_output` with an `isin_lists` argument. **Do not** add a second writer or a second save.

- After the `I. Assets` blocks and `check_macalc_asset_table`, and before `book.save()`, write each list's key column, and value column for the tables, from row 7 as one bulk assignment.
- Clear any old rows between the new end and the template end.
- Write values only; keys are written as text.
- Leave the title (row 5) and every other cell on `I. Inputs` and `P. Parameters` untouched.
- Keep the external-link restoration (Decision 38).

Pass criteria: an opt-in Excel test on a disposable template copy passes, the names resolve to the new sizes, and the template hash is unchanged.

#### C5. Verify On Read-Back (depends on C4) [✓ Implemented — see addendum]

Extend `verify_macalc_output`:

- Exact equality of each list block against the expected keys and values.
- Each name's resolved `refers_to_range` equals rows 7..6+n on the expected column(s).
- No formulas in the list blocks.
- Every other cell in the used ranges of `I. Inputs` and `P. Parameters` equals the template (formula comparison, as for protected rows). The only exception is the list blocks, including cleared tails.
- Re-run the alignment against the **read-back** `I. Assets` D/E. Every non-unmatched key must now be an exact D value. For every source entry, the set of emitted keys must equal {D of rows in `M(k)`}. The per-list matched count must equal the template's matched-on-D count (24/106/31/16 on 111001), plus the Added count when there are fan-outs.
- Existing `I. Assets` checks are unchanged. They already prove that Slice C did not alter `I. Assets`.

Pass criteria: the Excel test also proves that a deliberately corrupted list cell makes verification fail. A synthetic fan-out case (two asset rows sharing one E) writes, resizes the name and verifies correctly.

#### C6. Orchestrate In The Single Run (depends on C5) [✓ Implemented — see addendum]

In `run_stage_three`, pass the template lists from the snapshot and the funded and asset rows to `align_isin_lists` **before** the staged write, so a Guard 5 failure stops before any output. Print a per-list summary to the console: kept, remapped, unmatched and added.

There are no new audit fields, panel, button or VBA changes, and no reinstall (Decision 14).

Pass criteria:

- An orchestration test covers a Slice C guard failure after the Running audit: no new output, Failed audit, previous output kept.
- One live source-mode run on handoff `20261102_111001` records Succeeded, with the expected summary:

  | List | Kept | Remapped | Unmatched | Added |
  |---|---|---|---|---|
  | Forced | 19 | 5 | 0 | 0 |
  | Min | 0 | 106 | 58 | 0 |
  | Pre16 | 0 | 31 | 291 | 0 |
  | REITs (if in scope) | 0 | 16 | 0 | 0 |

- The console shows no fan-outs on 111001 (Added = 0 everywhere), consistent with E being unique in the asset block.
- The user spot-checks a few remapped entries against `I. Assets` D.

#### C7. Documentation (parallel with C6) [✓ Implemented — see addendum]

- Update `references/bonds-lite-setup.md` (the Stage 3 section: the list alignment, the new error messages and the console summary).
- Update the handoff, and add a Slice C addendum here.

### Test Strategy (Slice C)

1. **Pure:** `align_isin_lists` and the list parsers, on synthetic keys that mimic the real shapes: `Note 1`, `Inf_swap_FXD` / `Inf_swap`, a funded plain ISIN, `X_JRUSPP_Metlife` / `X_JRUSPP_Metlife_BBB`, the same E at two ratings (fan-out), stale `XS…`, and float32 values.
2. **Opt-in Excel:** a disposable template copy, covering write, verify, name resizing, the cleared tail, untouched neighbours and the corrupted-cell negative test.
3. **Real-data regression (opt-in):** extend `test_excel_slice_b_real_data_regression_on_handoff_111001`. Keep IDs and counts only, as in the C6 summary table: list sizes and matched counts per list. No production values are copied into fixtures.
4. **Informative parity:** after the run, the set of (`I. Assets` row, Comp A %) pairs picked up by MACalc's merge on D equals the template's set of (asset by ISIN, Comp A %) pairs.

### Affected Files

- `src/sst_stage_three.py`: the `MACALC_ISIN_LISTS` constant, `read_isin_lists`, `align_isin_lists`, the `TemplateSnapshot` field, and extended write, verify and orchestration.
- `tests/test_sst_stage_three.py`: the Slice C tests and the extended real-data regression.
- `references/bonds-lite-setup.md`, `references/decisions.md` (49-53) and this plan.
- `src/bonds_lite_input.py`: one line only, the controller logger name (addendum deviation 5).
- **No change:** the Draft, the VBA, the packaging and `I. Assets`.

### Risks And Unknowns

- **Template lists are stale.** 58 Min and 291 pre-16 entries match nothing. Slice C preserves this; curating the lists is a business task outside Stage 3.
- **The near miss `XS0306647792`.** It may be a real pre-16 asset that has lost its match because of the `_JRABACUS_Just` suffix. It was already unmatched in the template, so Slice C reports it but does not fix it (Q-C4 / C6).
- **Future multi-matches (C3)**, from a funded overlap or the same ISIN at several ratings, make a list longer. This is safe because the names are dynamic, and it is covered by the name-size check in C5.
- **Consuming model not checked in Stage 3 (Q-C6).** The output will feed the Python MACalc in `MACalc_Batch`, not `v_3_3`. The list usage, the duplicate-key risk and the REIT/noPRoC conclusions come from `v_3_3` and the template, and they are confirmed only when Stage 5 (the model run on the validated input) is planned (C0). If that model joins differently, Slice C is reopened.
- **Runtime:** about 530 cells are written to an already open book, which is negligible against the 112 s run.

## Addendum: Slice C Implementation Progress

Implemented on 3 November 2026 (steps C0-C7). Nothing has been committed, built or promoted. `MACalc_Batch` was not opened, changed or run (C0, Decision 51).

**Tests first (C1).** Slice C tests were added to `tests/test_sst_stage_three.py` and were seen failing only on the missing code. They cover:
- pure alignment: keep, remap, ILS E not used, unmatched, fan-outs (a), (b) and (a)+(b) in row order, determinism, the Decision 52 duplicate rules, sizes, Guards 3, 4 and 7, empty list, no-suffix and mixed;
- parser tests on fake sheets for Guards 1, 2 and 6;
- two orchestration tests: a guard failure keeps the previous output with a Failed audit; the aligned lists go to the single write and verify;
- an opt-in Excel fan-out test on a disposable template copy;
- the extended opt-in real-data regression (`SLICE_C_GOLDEN`, informative parity).

**Code (C2-C6)** is in `src/sst_stage_three.py`:
- `MACALC_ISIN_LISTS`, `parse_isin_list` and `read_isin_lists` read the lists through the already-open template in `_read_template_snapshot`. There is no new Excel open.
- `build_asset_keys`, `align_isin_lists` and `log_isin_alignment` are pure.
- `write_macalc_output(isin_lists=...)` writes in the same staged book before the single save.
- `verify_macalc_output(isin_lists=..., template_isin_lists=..., funded_count=...)` reuses the existing template open and read-back.
- `run_stage_three` aligns before the staged write.
- `MapBaseRow.final_rating` (CSV column H, `Final_Rating`) is read and passed as `asset_ratings` to `build_asset_keys`, both in the alignment and in verify (deviation 1).
- There are no audit, panel, button, VBA or packaging changes. `bonds_lite_input.py` has a one-line logger change (deviation 5).

**Results:**
- `tests/test_sst_stage_three.py`: 193 passed, 4 skipped (the opt-in Excel tests).
- Opt-in Excel run (`BONDS_LITE_RUN_STAGE_THREE_EXCEL_TESTS=1`): the Slice C fan-out test and the 111001 real-data regression both passed (2 passed, 257 s). The fan-out test covered write, name resize, positive verify and the corrupted-cell negative check. The regression showed Forced 24/19/5/0/0, Min 164/0/106/58/0, Pre16 322/0/31/291/0 and REITs 16/0/16/0/0 (size/kept/remapped/unmatched/added), with matched counts 24/106/31/16. Parity held, and no fan-outs occurred.
- Full suite (`python -m pytest -q`): 856 passed, 85 skipped, 2 failed. The failures are the 2 known pre-existing `test_bonds_lite_input.py` failures (`test_frozen_main_validates_cached_bundle_before_stage_dispatch`, `test_model_results_transaction_rollback_failure_retains_recoverable_backup`); the Slice B baseline was 797/84/2.

**Deviations and interpretations:**
1. **Guard 7's "own CSV rating" is the CSV `Final_Rating` (column H) (Decision 53, a clarification approved by the user on 3 November).** `I. Assets` column I holds `Final_Notched_Rating` (CSV column I, e.g. `BBB+`). On 111001, `Unique_ISIN` = ISIN + `_` + `Final_Rating` holds on all 988 rows but on only 430 with the notched rating. Guard 7 therefore compares exactly: D must be E or E + `_` + `Final_Rating`; anything else stops the run. The CSV header H is checked, and the ratings travel in CSV order to the alignment and to verify, because `Final_Rating` is not written to `I. Assets`. An earlier draft used `rating.rstrip("+-")`; that heuristic is withdrawn.
2. **The C5 matched-count check is internal, not a fixed number.** Template assets change each quarter, so the runtime verify checks that the number of written keys found in the read-back D equals the alignment's own matched count. It also checks that realigning against the read-back `I. Assets` reproduces the written keys and values exactly. The 111001 numbers (24/106/31/16) are asserted in the real-data regression.
3. **Duplicate removal (Decision 52) runs even when no asset row is suffixed.** Mapping is then a no-op, so a list without duplicates is returned unchanged, as C1 requires.
4. **The opt-in Excel test opens its disposable copy with `ignore_read_only_recommended=True`.** The template is saved as read-only recommended; without the flag, the corrupted cell could not be saved for the negative check. Production code already opens it this way (`shared.owned_book`).
5. **One-line logger change in `bonds_lite_input.py` (the plan said "No change").** In source mode (`python -m bonds_lite_input`) the controller's module logger was named `__main__`, so the console handler never saw Stage 3 messages and the C6 summary was invisible. The controller logger is now named `bonds_lite_input` whichever way it is launched, and the Stage 3 logger is its child. Two tests cover this (logger name; summary reaches the controller console).

**Live run (C6), 3 November 2026.** Source mode against the open Draft on handoff `20261102_111001` (Stage 2 `Succeeded`). Result: `Succeeded` in 109 s (Draft audit 13:27:46-13:29:25), output `Bonds Lite output/20261102_111001/MACalc input.xlsb`, with read-back verification passed. Console summary:

| List | Kept | Remapped | Unmatched | Added | Duplicates removed |
|---|---|---|---|---|---|
| Forced | 19 | 5 | 0 | 0 | 0 |
| Min | 0 | 106 | 58 | 0 | 0 |
| Pre16 | 0 | 31 | 291 | 0 | 0 |
| REITs | 0 | 16 | 0 | 0 | 0 |

There were 5 funded and 988 asset rows; 8 MACE-only rows were ignored in each of `I. Bonds` and `Control`, and there were no fan-outs. This matches the expected C6 table. Two earlier attempts on the same day were diagnostic: the first stopped before any output because `--caller` used the local path while Excel held the SharePoint URL; the second succeeded but showed no summary (deviation 5). The accepted Slice B output and both earlier outputs are kept outside the repository in the agent session folder.

**Final opt-in Excel run (after the `Final_Rating` change).** All 4 opt-in Excel tests passed (Slice C fan-out, 111001 real-data regression, Slice B write and Draft-copy installer). The installer test's baseline comparisons were adjusted so that a Draft that already carries the Stage 3 names, button and a `Succeeded` audit is compared like for like; the test still proves the original Draft is unchanged. Excel printed the known harmless `0x80010108` teardown message after the pass.

**Acceptance.** On 3 November 2026 the user approved Decision 53 and accepted Slice C after the spot-check. Slices A, B and C are all accepted.

**Commit and release (3 November 2026).** Slices A-C were committed in `0f6ef09` (after `85cbe07`, the Stage 1/2 handoff). The first build, `20261103-cache-v5`, was never promoted: the shared PyInstaller options added `--hidden-import sst_stage_three` to the model build as well, which bundled `bonds_lite_input`, `sst_stage_three` and `win32job.pyd` into `BondsLiteModel.exe`. The hidden import now applies to the controller only (`tools/build_packaged_stages.py`, regression test `test_pyinstaller_command_hidden_imports_stage_three_in_controller_only`). `20261103-cache-v6` was built from `0f6ef09` plus that fix. Its model bundles exactly v4's modules and v4's 869 files; its controller bundles v5's modules, including `sst_stage_three`. Both self-tests passed. The central `current.txt` was switched from v4 to v6 at about 16:09, and the user then ran Stages 1-3 with the packaged release (handoff `20261103_161318`; Stage 2 summary `bonds_batch_summary_20261103_161723.csv`; output `MACalc input.xlsb` at 16:24) and reported that everything works. The build fix and these release notes are not yet committed. Rollback: write `20261102-cache-v4` to the central `current.txt` (v4 rejects Stage 3).

## Slice D: Unique MACalc Input Filename

Planned on 3 November 2026 for implementation in a new session; revised the same day so that the suffix is collision-resistant across synced folders, not just within one output root. **Answered and implemented on 3 November 2026:** the user's answers to Q-D1 to Q-D4 changed D4 and D6 to D8 (see the design table and "Addendum: Slice D Implementation Progress"). Decision 54 is approved with changes.

### Overview

The Python MACalc used in Stage 4 (`MACalc_Batch`) needs each MACalc input workbook to have a name that is unique across the synced SharePoint folders. Today every Stage 1 run folder holds the same name, `MACalc input.xlsb`. Slice D changes the filename to `MACalc input_<Stage 1 run ID>_<token>.xlsb`, for example `MACalc input_20261103_161318_3f9c2a7be104d65e.xlsb`, where `<token>` is 16 random lowercase hexadecimal characters generated fresh on every Stage 3 run. The run ID keeps the file traceable to its handoff; the token makes the name unique across Drafts, output roots, users and copied run folders. The workbook contents, checks, read-back, audit fields, button and VBA are unchanged. Because the name can no longer be derived, Stage 4 must locate its input through `SST_Stage3_LastSuccessfulOutput`. Colleagues only get the new name through a new packaged release (v7), because v6's controller writes the old name.

### Requirement (user, 3 November 2026)

"The MACalc_batch model that will be used for Stage 4 requires the MACalc input spreadsheet to be uniquely named across synced SharePoint folders. So when creating the MACalc input at the start of Stage 3 can we add a unique suffix after it?"

### Verified Discovery (read-only, 3 November 2026, commit `0f6ef09`)

- The name is the module constant `OUTPUT_FILENAME = "MACalc input.xlsb"` (`src/sst_stage_three.py`, line 36). It is used once, in `run_stage_three` (about line 2177): `output_path = current_run / OUTPUT_FILENAME`, followed by `_confirm_overwrite(caller, output_path)`.
- Publishing today (about lines 2276-2319): if `output_path` exists it is renamed to `_backup_path(output_path)` (`.<stem>.<timestamp>.backup.xlsb`), the staged file is renamed onto `output_path`, the audit is written, then the backup is deleted. On failure the published file is removed and the backup is renamed back. This replace-in-place logic exists only because the name is fixed.
- The audit `SST_Stage3_LastSuccessfulOutput` stores the output path relative to the Draft folder, and the console logs the full path, so both follow whatever `output_path` is.
- The staged file is `.MACalc-stage3-<random>.xlsb` (from `tempfile.mkstemp`) in the run folder. It is renamed onto the output, or deleted on failure.
- Run folder names come from `allocate_run_folder` (`src/bonds_lite_input.py`, about line 1270): `YYYYMMDD_HHMMSS`, plus `_001`, `_002` and so on on a same-second clash in the same output root. **They are unique only within one output root.** Two Drafts (or two users) whose Stage 1 starts in the same second would produce the same run ID, and a copied run folder repeats it. The run ID alone therefore cannot meet the requirement.
- Tests that use the name or the replace logic (`tests/test_sst_stage_three.py`, at `0f6ef09`):
  - lines 149, 264 and 2938 use `OUTPUT_FILENAME` only as an arbitrary output path;
  - line 987 (orchestrator fixture), `test_run_stage_three_zero_funded_succeeds_and_warns_without_overwriting_template` (about line 1085, asserts no `*.backup.xlsb`), `test_run_stage_three_declined_overwrite_does_not_change_audit_or_output` (about 1107), `test_run_stage_three_rolls_back_output_when_success_audit_fails` (about 1124), `test_run_stage_three_slice_b_guard_failure_keeps_previous_output` (about 2124) and `test_run_stage_three_slice_c_guard_failure_keeps_previous_output` (about 2802).
- Docs that use the name: P12 and Resolved Question 7 in this plan; `references/bonds-lite-setup.md` (2 places); the Stage 3 handoff.
- Path length on this Draft: today's output path is 154 characters; the new name adds 33 (`_20261103_161318_` plus 16 token characters), giving about 187. That is under Windows' 260 limit and Excel's 218-character limit for opening a workbook path, but colleagues with longer OneDrive paths could exceed 218, hence D7.
- `MACalc_Batch` was not inspected (Decision 51). Its naming rules are not known; see Q-D1.

### Design (Decision 54, approved with changes on 3 November 2026)

| ID | Topic | Design |
|---|---|---|
| D1 | Name | `MACalc input_<run ID>_<token>.xlsb`. `<run ID>` is the run folder name; `<token>` is `secrets.token_hex(8)` (16 lowercase hex characters, 64 random bits). Q-D1: `MACalc_Batch` has no naming conventions. |
| D2 | Uniqueness | A fresh token on every run, so no two Stage 3 runs share a name, whatever the Draft, user, output root or copied folder. With 64 random bits, the chance of any clash among 10,000 outputs is about 3 in a trillion (birthday bound n²/2⁶⁵). Stage 3 cannot see every synced folder, so this is collision resistance rather than a registry check; it additionally regenerates the token if the name already exists in the run folder. |
| D3 | Run ID check | The run folder name must fully match `\d{8}_\d{6}(_\d{3})?`. Otherwise Stage 3 raises `StageThreeError` naming the folder, before any audit or file write. |
| D4 | Reruns on the same handoff | **Changed (Q-D2):** earlier outputs are kept. Every rerun writes a new tokenised file alongside them, with no prompt and no deletion. `SST_Stage3_LastSuccessfulOutput` points to the newest successful output. |
| D5 | Failure and rollback | Because the new name never equals an earlier one, nothing has to be renamed aside or restored: on failure Stage 3 deletes only the file this run published (if the audit was not committed) and its staged file, and earlier outputs remain exactly as they were. `_backup_path`, `_confirm_overwrite` and the backup/restore branch are removed. The P14 guarantee (the previous output survives any failure) is kept by construction. |
| D6 | Legacy outputs | **Changed (Q-D4):** a plain `MACalc input.xlsb` is simply ignored; no prompt, log line, rename or deletion. Those run folders will be deleted. |
| D7 | Path length | **Changed (Q-D1):** before any write, Stage 3 stops with `StageThreeError` if the output path **or the staged path** exceeds 218 characters (Excel's limit for opening or saving a workbook; first 259, lowered by the user on 3 November), telling the user to shorten the output folder path. |
| D8 | Staged file | **Changed (Q-D3):** `MACalc_Batch` does not ignore dot-prefixed files, so nothing is staged in the run folder. The workbook is staged as `<output root>\Stage 3 staging\<token>\<final name>`, a sibling of the run folders, and the token folder is removed after every run (the `Stage 3 staging` folder too, once empty). Local `%TEMP%` was tried and rejected: Excel rewrites the template's SharePoint link paths when saving outside the synced folder, and the read-back link check stops the run. OneDrive marks the folders read-only, so removal clears that flag and retries; a folder that still cannot be removed is logged as a warning. The user confirmed on 3 November that `MACalc_Batch` scans only the folder holding the MACalc input it is given (the run folder), so the sibling staging folder is not seen. |
| D9 | Stage 4 contract | Stage 4 must read `SST_Stage3_LastSuccessfulOutput` (and check `SST_Stage3_InputRun` matches the handoff) to find its input. It must not build or glob the filename. |
| D10 | Release | Build a new release (for example `<date>-cache-v7`) without `--promote`; check that the model bundles exactly v4/v6's modules and the controller v6's; run both self-tests; promote with the user's approval; the user tests Stages 1-3. Rollback is to `20261103-cache-v6`. |

Alternatives for Decision 54:
- **Run ID only** (the first draft of this plan). Readable and stable across reruns, but unique only within one output root: two Drafts starting Stage 1 in the same second, or a copied run folder, give the same name. **Rejected:** does not meet the requirement.
- **Run ID plus Windows user name.** Still clashes for the same user running two Drafts in the same second, and puts the user name in the file name. Rejected.
- **Run ID plus a token reused across reruns** (found from the existing file). Keeps one stable name per handoff, but a copied run folder would carry the token and produce duplicates on rerun. Rejected in favour of a fresh token per run.
- **Full UUID4 (32 hex characters).** 122 random bits, but adds 16 more characters to a path that is already near Excel's 218 limit on some machines. 64 bits is ample for the volumes involved. Kept as a fallback if the user prefers it (Q-D1).
- **Stage 3 timestamp instead of a token.** Not unique across Drafts started in the same second. Rejected.

### Questions For The User (start of the implementation session)

- **Q-D1 Naming rules in `MACalc_Batch`.** Does it require a pattern, for example a prefix, separator, no spaces (today's name has one), case or a maximum length? Is a 16-character random token acceptable, or is a full UUID preferred?
- **Q-D2 Earlier outputs on rerun.** After a successful rerun on the same handoff, delete the earlier tokenised outputs (proposed, D4) or keep them?
- **Q-D3 Hidden files.** Does `MACalc_Batch` ignore dot-prefixed files in synced folders (the short-lived staged file)? If not, the staged file may need a different location.
- **Q-D4 Legacy outputs.** Leave the existing `MACalc input.xlsb` files in older run folders (proposed, D6), rename them, or delete them? `20261102_111001` and `20261103_161318` have one each.

### Steps

**D0. Prerequisites.** [? Implemented ? see addendum]
- Commit the v6 build fix (`tools/build_packaged_stages.py` and its regression test in `tests/test_bonds_lite_input.py`) with the v6 release notes and the Slice D plan.
- Get the answers to Q-D1 to Q-D4 and record Decision 54 as approved or changed.
- Verify: `git status` shows no tracked source changes apart from the two `.pyc` files.

**D1. Write the tests first** [? Implemented ? see addendum] in `tests/test_sst_stage_three.py` (depends on D0):
- `test_output_filename_joins_run_id_and_token`: `stage_three.output_filename(Path("Bonds Lite output") / "20261103_161318", "3f9c2a7be104d65e") == "MACalc input_20261103_161318_3f9c2a7be104d65e.xlsb"`.
- `test_output_filename_keeps_collision_suffix`: run folder `20261103_161318_001` gives `MACalc input_20261103_161318_001_<token>.xlsb`.
- `test_output_filename_rejects_unexpected_run_folder`, parametrised over `""`, `"Bonds Lite output"`, `"20261103"` and `"20261103_161318 copy"`: raises `StageThreeError` matching `"run folder"`.
- `test_new_output_token_is_16_lowercase_hex_and_varies`: matches `[0-9a-f]{16}`; 1,000 calls give 1,000 distinct values.
- `test_new_output_path_regenerates_token_when_name_exists`: with a token source that first returns an existing file's token, the returned path is the second token's.
- `test_previous_run_outputs_matches_only_this_run`: in a folder holding `MACalc input.xlsb`, `MACalc input_<run>_<tok>.xlsb`, `MACalc input_<other run>_<tok>.xlsb`, `MACalc input_<run>_XYZ.xlsb` and `.MACalc-stage3-abc.xlsb`, only the second is returned.
- `test_run_stage_three_rejects_unexpected_run_folder_before_any_write` and `test_run_stage_three_rejects_output_path_over_218_characters_before_any_write`: no prompt, no audit write, no new file.
- Orchestrator fixture (about line 987): stop asserting a fixed name; assert the single published file matches `MACalc input_<run>_[0-9a-f]{16}\.xlsb` and that `SST_Stage3_LastSuccessfulOutput` points to it.
- `test_run_stage_three_rerun_replaces_earlier_output_after_success`: an earlier `MACalc input_<run>_<tok>.xlsb` with sentinel bytes; the prompt receives it; after success it is gone, exactly one new tokenised file exists, and the audit points to the new one.
- `test_run_stage_three_rerun_keeps_earlier_output_if_deletion_fails`: deletion raises `OSError`; the run still succeeds and logs a warning naming the file.
- Adapt `test_run_stage_three_declined_overwrite_does_not_change_audit_or_output`: the prompt is driven by an earlier tokenised file.
- Replace `test_run_stage_three_rolls_back_output_when_success_audit_fails`: the earlier tokenised file is byte-for-byte unchanged, the new file is removed, and no `*.backup.xlsb` exists.
- Adapt the Slice B and Slice C guard-failure tests (about lines 2124 and 2802): the earlier tokenised output is unchanged and no new file exists.
- `test_run_stage_three_leaves_legacy_output_untouched`: a `MACalc input.xlsb` with sentinel bytes is unchanged, is not passed to the prompt, and one information line mentions it.
- Lines 149, 264 and 2938: replace `OUTPUT_FILENAME` with a literal test filename. These tests only need some output path.
- Use an injectable token source (for example a module-level `_token = secrets.token_hex` that tests monkeypatch) so names are deterministic in tests.
- Verify: the new and updated tests fail for the expected reason (missing functions, the old fixed name, or the old backup behaviour), and nothing else fails.

**D2. Implement** [? Implemented ? see addendum] in `src/sst_stage_three.py` (depends on D1):
- Replace `OUTPUT_FILENAME` with `OUTPUT_STEM = "MACalc input"`, `OUTPUT_EXTENSION = ".xlsb"`, `LEGACY_OUTPUT_FILENAME = "MACalc input.xlsb"`, `MAX_OUTPUT_PATH = 218` and compiled run-ID and token patterns.
- Add small pure helpers:

  ```python
  def output_filename(run_folder: Path, token: str) -> str:
      """Return the MACalc input name for a Stage 1 run folder and token."""

  def new_output_path(run_folder: Path) -> Path:
      """Return an unused, token-suffixed output path in the run folder."""

  def previous_run_outputs(run_folder: Path) -> list[Path]:
      """Return earlier tokenised outputs of this run, sorted by name."""
  ```

- In `run_stage_three`: compute `output_path = new_output_path(current_run)` and `previous_outputs = previous_run_outputs(current_run)`; apply the D7 length check; ask `_confirm_overwrite` only when `previous_outputs` is not empty (pass the list); log the D6 legacy line.
- Publishing: `staged_path.replace(output_path)`, write the `Succeeded` audit, then delete each of `previous_outputs`, logging a warning for any that cannot be deleted. Remove `_backup_path` and the backup/restore branch; on failure, delete only `output_path` if this run published it and the audit was not committed.
- Verify: `python -m pytest -q -p no:cacheprovider` gives 866 plus the new tests passed, 85 skipped, and only the 2 known failures (`test_frozen_main_validates_cached_bundle_before_stage_dispatch` and `test_model_results_transaction_rollback_failure_retains_recoverable_backup`).

**D3. Opt-in Excel tests** [? Implemented ? see addendum] (depends on D2): set `BONDS_LITE_RUN_STAGE_THREE_EXCEL_TESTS=1` and run with `-m excel`, with the Draft closed.
- Verify: 4 passed. The `0x80010108` teardown message is harmless.

**D4. Update the docs** [? Implemented ? see addendum] (parallel with D3):
- P12, P14 and Resolved Question 7 in this plan, plus an implementation addendum.
- `references/bonds-lite-setup.md`, around lines 110 and 130: the new name, that it changes on every run, and that Stage 4 reads it from `SST_Stage3_LastSuccessfulOutput`.
- The handoff.
- Decision 54 marked as approved.

**D5. Live source-mode run** [? Implemented ? see addendum] (depends on D2): on the current handoff (today `20261103_161318`, if `SST_Stage2_InputRun` matches it). Run Stage 3 twice.
- Verify after the first run:
  - exactly one `MACalc input_20261103_161318_<token>.xlsb` exists and the audit shows `Succeeded`;
  - `SST_Stage3_LastSuccessfulOutput` points to it;
  - the legacy `MACalc input.xlsb` is byte-for-byte unchanged and the legacy line was logged;
  - the Slice C console summary equals that of the v6 run on the same handoff.
- Verify after the second run (accept the prompt): a different token, the first tokenised file deleted, the audit pointing to the new file, and no `.backup` or `.MACalc-stage3-*` files left.

**D6. Release and commit** (depends on D3-D5):
- Build the next release without `--promote` and check its bundled modules as in D10.
- Promote with the user's approval and have the user run Stages 1-3 with the packaged release.
- Then commit with `feat: Give each MACalc input a unique filename` and the Co-authored-by trailer.

### Test Strategy (Slice D)

- Unit tests for the pure helpers: name format, collision-suffixed run IDs, rejected folder names, token format and variety, regeneration on an existing name, and the earlier-output matcher.
- Orchestrator tests, using the existing fakes and an injected token: the prompt only for earlier tokenised outputs, the audit path, deletion of earlier outputs only after a committed audit, a failed deletion as a warning, the earlier output left byte-for-byte unchanged on every failure path, the legacy file left alone, and the invalid-folder and path-length guards with no side effects.
- Real data: the two live D5 runs on `20261103_161318`, compared with the v6 output of the same handoff. Contents should be equal apart from the file name, so compare the Slice C summary and the read-back result rather than the file hash.

### Affected Files (Slice D)

- `src/sst_stage_three.py`: the name constants, `output_filename`, `new_output_path`, the staging folder in the output root and its read-only-aware removal, and `run_stage_three` publishing and failure handling; `_backup_path` and `_confirm_overwrite` removed. (`previous_run_outputs` was not needed once Q-D2 kept earlier outputs.)
- `tests/test_sst_stage_three.py`: new and updated tests.
- `references/plan-stage-three.md`, `references/bonds-lite-setup.md`, `references/decisions.md` (Decision 54) and the Stage 3 handoff.
- No changes to `src/bonds_lite_input.py`, the VBA, the Draft or the build script.

### Risks And Unknowns (Slice D)

- `MACalc_Batch` has no naming rules (Q-D1, answered by the user); its code was not inspected (Decision 51).
- Uniqueness is probabilistic (64 random bits per run), not checked against every synced folder, which Stage 3 cannot see. The clash risk is negligible at the expected volumes.
- The name changes on every rerun, so anything that remembered the old path (a user's shortcut, a Stage 4 configuration) must use `SST_Stage3_LastSuccessfulOutput` instead (D9).
- Earlier outputs accumulate in the run folder (Q-D2). If `MACalc_Batch` is pointed at a run folder rather than a single file, it will see every rerun's output; it must take the path from the audit (D9).
- The short-lived staged workbook lives in `<output root>\Stage 3 staging\`. `MACalc_Batch` scans only the folder holding its input (user, 3 November), so it does not see it.
- Because `MACalc_Batch` scans the run folder, it will also see earlier tokenised outputs kept there (Q-D2) and any other workbooks in that folder.
- v7 is live in the central pointer, but this user's LocalAppData cache remains on v6 until Excel checks the pointer on its next button click. Stage 4 must not start until the packaged Stage 1-3 test has passed.
- Path length: the output path is 187 characters on this Draft and the staged path 204. D7 stops cleanly above 218 characters, Excel's limit for opening or saving a workbook path. The staged path therefore has 14 characters of headroom on this Draft; a colleague with a longer OneDrive path gets D7's message before any write.

## Addendum: Slice D Implementation Progress

**3 November 2026.**

- **D0.** The full suite gave 866 passed, 85 skipped and the 2 known failures. Committed `ce799db` (`fix: Hidden-import Stage 3 in the controller build only`) and `d4e94c9` (`docs: Plan Slice D and record the v6 release`, including the Draft, at the user's request).
- **User answers.** Q-D1: no naming conventions; stay within the Windows path limit. Q-D2: keep earlier outputs. Q-D3: `MACalc_Batch` does not ignore dot-prefixed files, so stage elsewhere. Q-D4: leave legacy files; those run folders will be deleted. Decision 54 is approved with these changes; the design table above is updated.
- **D1-D2.** Tests were written first and failed for the expected reasons; the change is only in `src/sst_stage_three.py`.
- **Final suite:** 885 passed, 85 skipped, only the 2 known failures.

**Deviations from the plan (all from the user's answers, Decision 54):**
1. No overwrite prompt, no deletion of earlier outputs and no `previous_run_outputs` helper (Q-D2). The tests planned for prompting and deletion became `test_run_stage_three_rerun_keeps_earlier_output_without_prompting`.
2. The path limit applies to the staged path as well as the output path. It was first 259 (Windows `MAX_PATH`, Q-D1). After the packaged test the user lowered it to the planned 218, so that a long path stops before any write instead of failing inside Excel. The 218 limit is active in `20261103-cache-v8`; rollback to v7 restores the 259 limit.
3. Staging moved from the run folder to `<output root>\Stage 3 staging\<token>\` (Q-D3). Local `%TEMP%` was tried first and rejected: Excel rewrites the template's SharePoint link paths when the copy is saved outside the synced folder, so the read-back link check stops the run, and restoring the links reorders them and raised a COM error.
4. The first live run showed that OneDrive marks the staging folders read-only, so `shutil.rmtree(..., ignore_errors=True)` silently left empty folders behind. Removal now clears the read-only flag and retries, and logs a warning if a folder still cannot be removed. Tests: `test_run_stage_three_removes_staging_folders_marked_read_only` and `test_run_stage_three_warns_when_staging_folder_cannot_be_removed`.
5. No legacy log line (Q-D4); legacy files are simply ignored.

**D3, opt-in Excel tests.** These were run with `BONDS_LITE_RUN_STAGE_THREE_EXCEL_TESTS=1` and `-m excel`, with the Draft closed. All 4 pass, each in a different location:
- The three Stage 3 write tests pass with `--basetemp` set to a throwaway folder inside the synced repository, deleted afterwards. In the default local `%TEMP%` they fail at the link check for the reason given in deviation 3. This is a test-location effect and not a Slice D regression.
- The Draft-copy installer test passes in the default `%TEMP%`, but fails with a COM error on `Names` when it is run in the OneDrive basetemp.

**D5, live source-mode runs.** There were three runs on handoff `20261103_161318` (Draft opened in Excel by the agent; Excel reports its SharePoint URL, so `--caller-local` was used; `SST_Stage2_InputRun` matched).
- The legacy `MACalc input.xlsb` had already been renamed by the user to `MACalc input_test.xlsb` before the session. That file and `MACalc input_test_unoptimised_assets.csv` were byte-for-byte unchanged afterwards.
- **Run 1** (112.5 s) wrote `MACalc input_20261103_161318_c72f95ed916b8491.xlsb`.
- **Run 2** wrote `..._64bb8099ead76e39.xlsb`. The first file was kept and the audit pointed to the new one (`Succeeded`).
- **Run 3** (98.2 s, after the read-only fix) wrote `..._3889e5420fb7e83a.xlsb` and left no `Stage 3 staging` folder. The two empty token folders left by runs 1 and 2 were removed by hand.
- Every run printed the accepted Slice C summary: Forced 19/5/0/0, Min 0/106/58/0, Pre16 0/31/291/0 and REITs 0/16/0/0, with no duplicates removed. Each run had 5 funded and 988 asset rows, with 8 MACE-only rows ignored in each of `I. Bonds` and `Control`.
- There were no `*.backup.xlsb` or `.MACalc-stage3-*` files.

**Open:**
- (Resolved) The user confirmed that `MACalc_Batch` scans only the folder holding its input, i.e. the run folder, not the output root where staging lives.
- **D6 build (3 November, about 18:15).** `20261103-cache-v7` was built from the uncommitted Slice D working tree on top of `d4e94c9`, with the same build venvs (`controller-py313`, `model-py311`) and model folder as v6, **without `--promote`**. Both self-tests passed (`MODEL_SELF_TEST_OK`, `CONTROLLER_SELF_TEST_OK`). Compared with v6: the model has the same 869 files and 1,770 frozen modules, with no Stage 3 modules; the controller has the same 82 files and 469 modules (including `bonds_lite_input` and `sst_stage_three`), and nothing was added or removed. The v7 controller's `sst_stage_three` contains `STAGING_FOLDER` and `_remove_staging` and no longer contains `_confirm_overwrite`; v6's is the reverse. The central `current.txt` still points to v6. **Awaiting the user's approval to promote.**
- **D6 promotion (3 November, 18:18 UTC).** The user approved promotion. The central `current.txt` now points to `20261103-cache-v7`; the LocalAppData cache is expected to update on the next button click. Rollback is to `20261103-cache-v6`.
- **218-character limit release (3 November, 18:45 UTC).** Built `20261103-cache-v8` without `--promote`, using the established controller-py313 and model-py311 environments and the same model directory. Both self-tests passed (`MODEL_SELF_TEST_OK`, `CONTROLLER_SELF_TEST_OK`). Compared with v6, both bundles have identical files and frozen module sets: model 869 files / 1,770 modules; controller 82 files / 469 modules.
- **D6 promotion (3 November, 18:54 UTC).** The user approved promotion; the central `current.txt` now points to `20261103-cache-v8`. The user's LocalAppData cache is expected to update when Excel next checks the pointer. Rollback is to `20261103-cache-v7`.
- D6, still to do: the user's packaged Stage 1-3 test and the final Slice D commit.
