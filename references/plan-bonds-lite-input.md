# Plan: Generate Bonds Lite Input

Version 1.1 | 28 October 2026 | Implementation present; final live verification outstanding

## Overview

Add a Generate Bonds Lite Input button to the Draft workbook's I. Control sheet. A minimal VBA launcher starts Python with xlwings, without an Excel add-in; Python imports bond inputs, copies credit stresses, recalculates through desktop Excel and publishes a timestamped output beside the Draft workbook.

Use the existing Anaconda environment, with a configurable interpreter for each colleague. This is a Windows desktop Excel workflow, not an Excel Online solution. Stop at meaningful implementation decisions for user review.

## Evidence From The Workbooks

- Draft Input spreadsheet for SST.xlsm has I. Control and I. Credit Stresses. Although macro-enabled, its package currently has no VBA project.
- The destination sheet is I. Sensitivity, singular, not I. Sensitivities. Do not rename it.
- Both supplied bond-data sheets contain I_tbl_Monthly_Bonds at B8:AUN1006: 1,235 columns and 998 body rows. Worksheet dimensions extend beyond the actual table and must not determine import size.
- The destination also contains a separate Deleted assets table, I_tbl_Monthly_Bonds41. It is not an import target.
- Credit stress row groups are 9:15, 20:26, 31:37 and 42:48, respectively GBP financial, GBP nonfinancial, USD financial and USD nonfinancial.
- For every group, C:J contains transition probabilities, M:DI contains 101 transitioned-rating spread inputs, and DL:HH contains 101 original-rating spread inputs. Source and destination labels and positions match.
- Draft stress input cells are currently blank. The destination contains example transitions and zero spread inputs; these must never silently substitute for missing source inputs.
- Bond formula columns include J:K and ATU:AUN. Some depend on a network escalation workbook. Other existing external links and names also exist in the templates.
- There are no supporting Word specifications, expected-output files, architecture reference or existing plan documents in the supplied folder. Inspection was read-only using openpyxl and workbook package metadata; no Excel calculation or link refresh was performed.

## Confirmed Behaviour

1. Read the Data workbook path from Draft I. Control!D39 and the Bonds Lite input template path from D41. Accept absolute paths or filenames relative to the Draft workbook's local folder.
2. Read bond records from the selected workbook's I. Bonds Data table. Support more or fewer rows; do not add a bond-data table to Draft or silently accommodate schema changes.
3. Copy raw input values, retaining and recalculating the destination's formula columns. Preserve destination formatting, table identity and other sheets.
4. Copy all 12 numeric credit-stress blocks to matching cells in I. Sensitivity, as values, with no scaling. Spreads are decimals. Do not interpret scenario toggles or implement a scenario engine.
5. Reject missing stress inputs, empty bond data, changed or missing headers, inaccessible required dependencies and failed required calculations. Report missing-cell locations rather than retaining examples.
6. Copy Draft I. Control!D11 to destination I. Control!D10 and Draft D9 to destination D13. Leave other destination control settings unchanged.
7. Create Bonds Lite output under the Draft workbook's folder if absent. Keep every successful run as a timestamped .xlsx copy; never overwrite either source workbook or an earlier output.
8. Close the generated workbook after saving and show its full saved path.

## Steps

### 1. Establish The Test Contract Before Implementation

[✓ Implemented — see addendum]

Affected files: tests/test_bonds_lite_input.py and tests/fixtures/bonds_lite_cases.json.

Write initially failing tests for path resolution, mapping, numeric validation, schema validation, output naming and the orchestration contract. Store distinctive expected cell values in the small JSON fixture; do not duplicate the large workbooks unnecessarily. Tests involving Excel use temporary copies of the supplied workbooks.

Verification: every required stress block has an independent expected source/destination pair; blank and zero are distinguished; tests fail because the implementation is absent, not because fixture setup is broken. No real data is uploaded or placed in a new repository.

### 2. Configure Python And The Add-In-Free Launcher

Depends on step 1. Affected files: pyproject.toml, src/bonds_lite_input.py, and Draft Input spreadsheet for SST.xlsm, VBA module modBondsLiteLauncher and I. Control.

Use a small setuptools src-layout project, runtime xlwings, and development pytest/openpyxl dependencies. Check the actual Anaconda interpreter and choose a supported Python range before declaring package metadata. Install the project in that environment; no xlwings add-in or RunPython dependency is permitted.

At an implementation checkpoint, confirm a per-user interpreter setting, recommended as the BONDS_LITE_PYTHON environment variable containing the absolute python.exe path. Each colleague configures their own existing Anaconda environment rather than inheriting the author's personal path.

Add a Form Control button labelled Generate Bonds Lite Input on I. Control in the existing run area near row 73 without covering content. Assign GenerateBondsLiteInput in modBondsLiteLauncher. The launcher validates the interpreter and starts the installed Python module with the caller workbook's full name and exact Excel instance identifier. Quote Windows command arguments correctly, including spaces, apostrophes and non-ASCII usernames.

Launch asynchronously so VBA does not block Excel's COM message handling. Python must attach to the same running Excel instance, resolve the exact caller workbook, and read its live values, including unsaved edits. Do not reopen Draft from disk, select an arbitrary active workbook or start a replacement Excel application. Add a per-caller run guard and release it on both success and failure.

Verification: a button click reaches Python with no add-in installed, uses the intended interpreter, reads unsaved control edits, does not freeze Excel, and rejects duplicate concurrent clicks. Test two running Excel instances and both supported Office bitness configurations if both are deployed.

### 3. Validate Inputs Before Modifying A Working Copy

Depends on step 2. Affected file: src/bonds_lite_input.py, resolve_paths, validate_inputs and read_credit_stresses.

Require a saved caller with a local or UNC filesystem folder. If Excel exposes a SharePoint URL instead, stop with an instruction to use the locally synchronised folder; do not treat an HTTPS URL as a filesystem path. Resolve D39/D41 against that folder and reject source/caller/output aliasing.

Read all 5,880 required stress cells: four groups of 7 x 8 transitions and two 7 x 101 spread grids. Reject blanks, booleans, nonnumeric values, non-finite numbers and Excel error values, while allowing genuine zero values. Validate headers and rating/year ordering against the expected destination layout. Do not introduce unstated transition normalisation, probability-sum tolerances or spread transformations.

Inspect the source table body, reject an empty dataset, blank record identifiers and schema mismatches, and obtain its actual row count from the table, not UsedRange. Open data/template files read-only, with their own macros disabled and no automatic link refresh. Initially reject files already open for editing rather than taking ownership of a colleague's unsaved source state. Restore Excel security settings after opening files.

Verification: missing files, bad paths, each missing stress block, schema drift and empty data fail with contextual messages and no published output. Data and template modification times/content remain unchanged.

### 4. Resize And Populate Bond Data Safely

Depends on step 3. Affected file: src/bonds_lite_input.py, populate_bond_data.

Create an isolated working copy of the input template. Identify I_tbl_Monthly_Bonds specifically on I. Bonds Data. Capture its formula definitions and calculated-column metadata before resizing. Derive writable input columns from the verified template schema, checking all relevant rows for mixed formulas/constants rather than assuming every first-row formula represents an entire column.

Resize the table to exactly the source body-row count while retaining its header and 1,235-column schema. Clear obsolete body rows on shrink without damaging surrounding content; check for collisions with content or array/spill ranges on growth. Populate raw inputs in bounded two-dimensional blocks rather than per-cell COM calls.

Reapply internal and external destination formula columns to the correct resized extent. Handle the array formula in investment expense rate and modern Excel formulas explicitly; do not overwrite spill ranges or assume Formula2 and FormulaArray are interchangeable. Preserve table names and structured references. Do not copy source external-link formulas, replace the whole sheet or change Deleted assets.

Verification: smaller and larger test cases produce exactly the intended number of bond rows, no stale records, unchanged headers/table identity and valid formula coverage for every new row. Growth/shrink tests include formula and array-range behaviour.

### 5. Populate Sensitivities, Refresh Dependencies And Recalculate

Depends on step 4. Affected file: src/bonds_lite_input.py, populate_credit_stresses, synchronise_controls and refresh_and_validate.

Write the 12 approved blocks as numeric values to I. Sensitivity. Preserve labels and all cells outside those blocks. Synchronise only valuation date and entity.

Resolve the external workbook dependencies actually required by the retained bond formula columns. Refresh those dependencies explicitly, not every unrelated legacy link in every workbook. Stop if a required dependency is unavailable, update fails or required input/calculated cells contain Excel errors after calculation. A successful Calculate call alone does not prove a link refreshed: verify link availability/status and dependency values. Never silently accept cached values after a failed refresh.

Calculate within the owned working copy and check Excel calculation completion with a bounded timeout. Avoid recalculating unrelated user workbooks. Do not redefine the template's network paths, risk-factor directory, internal output-folder setting, suffix or model assumptions without a further user decision.

Verification: all 5,880 exported stress values equal the source exactly, date/entity match, formula coverage is intact and required calculated cells have no errors. An unavailable escalation workbook or deliberately broken required formula prevents publication.

### 6. Publish And Restore Excel State

Depends on step 5. Affected file: src/bonds_lite_input.py, publish_output and generate_bonds_lite_input.

Use the Draft folder, not the current directory or destination template's output-folder setting, for Bonds Lite output. Proposed filename: Bonds Lite input_YYYYMMDD_HHMMSS.xlsx; if a collision occurs, add a numeric suffix without overwriting. Verify that final names are unused immediately before publishing.

Save and validate an intermediate .xlsx working file, close it, then publish to the final name on the same filesystem. Clean up only files created by the current run on failure. Do not announce success until the published workbook exists and a reopen check confirms the expected sheets, table bounds, formulas and stress/control values.

In finally cleanup, close only owned workbooks and restore any changed calculation, event, screen-update, alert and security settings. Never quit the user's Excel instance. Release the run guard. Display one actionable error on failure, or the full saved output path on success.

Verification: denied folder access, locked destinations, naming collisions, calculation timeout and save failure leave no successful-looking partial output. Earlier outputs and source files survive unchanged. The generated workbook is closed and Excel remains usable.

### 7. End-To-End Verification And Colleague Setup

Depends on step 6. Affected files: tests/test_bonds_lite_input.py and references/bonds-lite-setup.md.

Run the unit tests using the selected interpreter's python -m pytest. Run explicitly marked Windows/Excel integration tests using temporary copies, then exercise the actual button in a test Draft workbook. Record interpreter, Excel version, input dimensions, link availability and measured runtime.

Document Anaconda package installation, interpreter configuration, macro/trusted-location requirements, button use, expected paths, and recovery from validation/link errors. Do not instruct users to enable all macros or weaken organisation-wide security. Check any organisational macro-signing requirement before distributing the changed Draft workbook.

Verification: a second colleague can follow setup without an xlwings add-in or editing personal paths inside VBA. Measure the supplied typical-size dataset and agree a performance threshold after the baseline; no unapproved speed guarantee is assumed.

## Test Strategy

1. Unit tests first: all mapping coordinates, unchanged decimal units, valid zeros, aggregated missing-cell errors, path resolution and collision-safe names. Use exact equality for copied values rather than loose numerical tolerances.
2. Excel integration: import existing real data using temporary copies, replace blank stresses with distinctive approved synthetic values, compare every populated block and all copied raw values. Preserve sources and identify each fixture's workbook origin.
3. Dynamic rows: one bond, fewer than the template, the supplied typical count and a larger count. Reject zero bonds, missing identifiers and schema drift. Test shrinking after a larger import to expose stale rows.
4. Formula regression: compare recalculated cells against an independently prepared manual Excel copy for identical inputs and dependencies. Establish and approve a tolerance only if calculated numeric differences occur; never relax copy equality to hide a defect.
5. Failure and lifecycle: missing interpreter, disabled macros, wrong Excel instance, duplicate click, missing files/stresses/links, formula errors, unwritable folder, save collision, read-only sources and failed cleanup. Assert no original file changes or orphaned working output.
6. Performance: measure table transfer, link refresh, calculation and save separately at representative size. Use bulk COM transfers; profile only if measurements show a bottleneck. Integration/performance tests must be opt-in and must never refresh production links unexpectedly in ordinary unit tests.

## Affected Files

- Draft Input spreadsheet for SST.xlsm: new launcher module and Form Control button; D39/D41 input cells; preserve existing sheets and formulas.
- src/bonds_lite_input.py: installed Python entry point and the narrowly scoped workflow/helpers named above.
- pyproject.toml: setuptools src layout, supported interpreter range, xlwings runtime and test/development dependencies.
- tests/test_bonds_lite_input.py: unit and opt-in Excel integration checks.
- tests/fixtures/bonds_lite_cases.json: compact synthetic stress expectations and schema cases, not a second copy of granular production bond data.
- references/bonds-lite-setup.md: per-user deployment and troubleshooting notes, to be written during implementation.
- references/decisions.md: confirmed requirements and unresolved implementation checkpoints.
- Bonds Lite output/: runtime-created generated workbooks; do not create an empty output folder during planning.
- Bonds Lite data template.xlsx and Bonds Lite input template.xlsx: read-only sources, not implementation edit targets.

## Decisions

See references/decisions.md. The approved technology is Python with xlwings and an add-in-free VBA launcher, not a VBA-only solution.

## Risks And Implementation Checkpoints

- The original planning session created no production code. Implementation progress and remaining verification gates are recorded in the addendum below.
- Verify xlwings availability/licensing policy, the selected Anaconda environment and supported Python/Excel versions before installation. Package compatibility is not yet certified.
- Confirm the per-user interpreter configuration and run-guard mechanism before implementing the launcher. Keep it asynchronous to avoid Excel COM re-entrancy deadlocks.
- Required network dependencies must be accessible to every colleague. Recalculation cannot reconstruct unavailable source data.
- Array/spill formulas, worksheet protection and Excel's table-growth behaviour need live-Excel verification; package inspection does not establish that edits are permitted.
- Template controls include existing network paths. Only date and entity are approved for synchronisation; unrelated control clean-up is out of scope.
- Output timestamps preserve runs but are not an audit system. No new audit-log or source-data retention policy is implied.
- All-blank stresses currently prevent a successful real run until completed. No independently completed expected output was supplied.
- Shared OneDrive workbooks can be locked or exposed as URLs. Use local synchronised paths and do not claim atomic cloud synchronisation from a local rename.

## Skills Consulted

Questioning-patterns, document-extraction, excel-analysis, testing-strategy, regression-testing, performance-optimisation, decision-logging, coding-standards, developer-guides and learn-more. No team-specific conventions were applied, as requested.

## Addendum: Implementation Progress

### 28 October 2026

- Resumed from an existing partial Python module, package configuration and test suite rather than replacing those files. Reviewed the code and established a baseline of 154 passing Excel-independent tests.
- Step 1 complete: added nine initially failing publication regressions, implemented the helper, and verified all nine passing. Subsequent ordinary suite result: 163 passed. Additional Excel regressions cover literal raw text, dynamic rows, installer backup, isolated publication and missing dependencies.
- Step 2 implementation installed: local package installed in Anaconda Python 3.13.5, xlwings 0.32.1; per-user BONDS_LITE_PYTHON configured. The actual Draft workbook contains the Generate Bonds Lite Input Form Control button and VBA launcher. Original backup: references/backups/Draft Input spreadsheet for SST_20261028_091248_427529.xlsm. Installation regression passed, but Excel emitted a native RPC warning during disposal of the installer-owned instance. An actual button-click subprocess test remains outstanding.
- Step 3 implementation present: saved local/UNC path resolution, distinct-file/schema/identifier checks, live stress validation and safe workbook opening. The pure validation suite passes. Missing-dependency integration passed with no published workbook.
- Step 4 implementation present: bounded raw input transfers, table resize, formula preservation and exact-value read-back. Live one-bond and typical-size 998-row cases passed. Corrected raw text being interpreted as a formula; both equals-prefixed and literal-apostrophe text regressions subsequently passed. Final five-row and 1001-row results remain unverified.
- Step 5 implementation corrected: diagnostic isolated the UpdateLink failure to opening the source first. A closed source refreshed successfully with status 0; the same source open failed with status 9. The implementation now leaves dependencies closed, rejects already-open dependencies and requires status 0. Both network and local failures before this correction were therefore implementation evidence, not proof of a defective network workbook. The post-correction complete-workflow result remains unverified because terminal execution/output access stopped yielding reliable results.
- Step 6 implementation present: collision-safe publish_output moves only a validated working file and does not delete earlier outputs on failed candidate moves. Its nine regressions pass. Complete publication after the refresh repair still requires a successful end-to-end run.
- Step 7 partially complete: colleague guide saved as references/bonds-lite-setup.md. Normal Excel tests use a temporary synthetic escalation workbook; real-network refresh requires an additional explicit environment flag. This fixture deviation was approved by the user; see Decision 12 in decisions.md. Full final suite, real-network refresh, actual button click, two-instance/Office-bitness checks, second-colleague setup and a complete runtime baseline remain open.
- Verification constraint: repeated terminal launches returned no process output or final JUnit report, including after the user opened a fresh terminal. Do not treat those launches as passes or suppress failing tests. Resume the focused isolated workflow test first when command execution is reliable.

### Next Verification Commands

Use the configured interpreter from the project folder. First run the ordinary suite, then set BONDS_LITE_RUN_EXCEL_TESTS=1 and run test_excel_full_workflow_publishes_validated_copy. Keep BONDS_LITE_RUN_NETWORK_TESTS unset until intentionally testing the real dependency. Exact commands and interpretation are in the setup guide. Do not publish production results until the complete workflow has passed its refresh, calculation, save and reopen checks.

## Amendment: Console Progress Before Optimisation

28 October 2026 | Planned, not implemented | Decisions 13 and 14

### Overview

Keep the existing Excel processing, recalculation, required-link refresh and verification unchanged. First add console-only progress reporting so users can identify the current stage, see elapsed time and measure a complete representative run. The user has no measured runtime because a previous run was interrupted without visible progress; do not claim Excel opening is the bottleneck or promise a speed improvement.

### Confirmed Scope

- The button opens a visible console asynchronously, leaving Excel free to service COM calls.
- Display stage entry, successful completion or failure, stage duration and total elapsed processing time.
- Keep the console open until Enter after success or failure. Do not require Enter in ordinary programmatic calls to generate_bonds_lite_input.
- Console output only: no log files, separate error files or bond-record dumps.
- Do not replace Excel automation with file-only processing, change formula/dependency behaviour, or introduce a hidden Excel instance in this amendment.
- Preserve native FullName caller matching and the separate local path used for filesystem operations; do not hard-code a user's OneDrive directory.

### Steps And Verification

1. Write failing tests first in tests/test_bonds_lite_input.py. Revise test_launcher_uses_configured_python_without_an_addin to expect a visible, asynchronous launch instead of shell.Run command, 0, False. Extend test_main_matches_native_caller_identity to capture console results and stub any requested Enter prompt without weakening caller rejection or local-path assertions. Add fake-clock timing tests, success/failure/EOF pause tests, and failure-cleanup tests. Pass criteria: tests expose absent progress/visibility/pause behaviour, and normal unit tests neither start Excel nor block on input.

[✓ Implemented — see console progress addendum]

2. Add narrowly scoped console reporting in src/bonds_lite_input.py. Use a module logger and configure its console StreamHandler only from main, with immediate flushing and no FileHandler. Proposed local helpers are configure_console_logging and a timed_stage context manager using time.perf_counter. Include stage name and total elapsed time at entry; include stage duration on completion or failure. Re-raise original stage errors so existing cleanup and exit-code handling remain authoritative. Pass criteria: fake-clock tests assert exact durations and meaningful failed-stage messages; repeated main calls do not accumulate console handlers or modify unrelated logging handlers.

[✓ Implemented — see console progress addendum]

3. Instrument the controlling workflow, depending on step 2. In main, report startup and caller lookup before touching the workbook. In generate_bonds_lite_input, report input validation, run-guard acquisition, Excel-state setup/restoration, template copy, bond import, stress export, control synchronisation, save, persisted verification, publication and cleanup. In owned_book, report each open and close separately, distinguishing source data, source template, working copy and verification copy. In refresh_and_validate, report dependency discovery, each targeted link update, worksheet calculation, calculation waiting and formula validation. Include row counts but not row contents. Pass criteria: a known failure identifies the owning stage; each expensive COM call is preceded by a flushed message; changing instrumentation does not change call order, workbook writes or verification rules.

[✓ Implemented — see console progress addendum; live output-parity gate remains open]

4. Provide waiting visibility for long operations. Proposed default: every five seconds, print the current stage and elapsed stage/total time. Keep all Excel/COM access on the original thread; a reporting thread may use only synchronised plain Python timing/state values and an Event for termination. Start/stop it in a bounded context, and stop/join it on all exit paths. Label this output as waiting, not percentage completion, progress confirmation or proof that Excel is responsive. Pass criteria: fake event/clock tests establish reporting and termination without sleeping; no Excel object crosses into the reporting thread. Review the five-second interval before implementation under the user's decision-checkpoint preference.

[✓ Implemented — see console progress addendum; interval approved in Decision 15]

5. Add console completion behaviour, depending on steps 2-4. Print the saved path or an actionable failed-stage error and total processing time before waiting for Enter. Pause only for a button-launched interactive run, using an explicit CLI flag such as --pause; preserve noninteractive invocation and --install-button behaviour. Handle EOF without changing the existing success/failure code. Do not retain both a modal result message box and an Enter prompt. Perform workbook cleanup, Excel-state restoration and run-guard release before any pause. Pass criteria: success exits 0 and failure exits 1 after the expected pause; no ownership locks or temporary workbooks remain held while a user reads the summary.

[✓ Implemented — see console progress addendum]

6. Update VBA_LAUNCHER and the actual Draft workbook, depending on step 5. Change window style to visible, retain asynchronous execution, request unbuffered Python output and pass the pause flag. Preserve the configured per-user interpreter, native caller FullName, Excel HWND and quoted arguments. Update the embedded modBondsLiteLauncher only after creating a backup; use install_button with Draft closed and do not change unrelated workbook contents. A change to the Python VBA string alone does not modify the installed macro. Pass criteria: the saved embedded module contains the visible launch, and a real button click opens one console immediately without blocking Excel or requiring a user-specific VBA path.

[✓ Implemented — user confirmed installer completion; visible-window acceptance remains open]

7. Verify and measure, depending on step 6. Run the default suite, then explicitly enabled isolated Excel tests with the synthetic local dependency and no real-network flag. Check the installed button on a temporary Draft copy first. Complete one representative approximately 998-bond, 1,235-column run with the intended saved dependencies. Record stage durations and total runtime, excluding the Enter pause; report any Python startup delay separately rather than pretending the Python timer starts at the physical button click. Compare output raw values, formulas, stresses, controls and workbook ownership with the pre-instrumentation behaviour. Update references/bonds-lite-setup.md to describe console stages, waiting messages, Enter completion and the absence of log files. Pass criteria: stages are visible, output parity and cleanup checks pass, and a reproducible timing baseline exists before considering I/O optimisations.

[Partially verified — default suite and documentation complete; live checks and baseline blocked]

### Test Strategy

- Reuse tests/test_bonds_lite_input.py and existing temporary-workbook fixtures; no new test framework or large workbook fixtures.
- Unit tests first: deterministic stage duration, correct ordering, error propagation, handler de-duplication, console-only configuration, waiting-reporter shutdown, pause gating and EOF handling.
- Retain the eight native caller identity cases for SharePoint/local paths and different usernames.
- Integration checks: successful isolated export, missing dependency, invalid stresses, duplicate run and failed save; all must identify the stage and restore existing Excel state before prompting.
- Launcher test: inspect the actual installed macro on a disposable Draft copy for visible asynchronous execution, unbuffered output and pause flag.
- Manual Windows acceptance: a real button click shows output during the run, including long waits; success and failure remain readable until Enter; the console shows the final total before the prompt.
- Do not run real-network tests by default, fake a completion percentage or interpret a heartbeat as successful Excel calculation.

### Affected Files

- src/bonds_lite_input.py: reporting helpers, instrumented workflow/ownership boundaries, CLI completion and launcher string.
- tests/test_bonds_lite_input.py: revise the hidden-launch assertion and notification assertions; add focused timing, waiting and prompt tests.
- Draft Input spreadsheet for SST.xlsm: backed-up update to embedded modBondsLiteLauncher; retain the existing button and inputs.
- references/bonds-lite-setup.md: console behaviour, elapsed-time meaning and troubleshooting instructions, updated during implementation.
- references/decisions.md: approved scope and console-only preferences.

### Risks And Deferred Work

- The current refresh/calculation timeout cannot interrupt a COM call that blocks internally. A waiting message improves visibility but is not a watchdog or cancellation guarantee.
- Closing the console window is not a guaranteed graceful cancellation mechanism. Ctrl+C/error cleanup needs explicit testing; do not promise safe forced termination.
- Exact physical button-to-Python startup time is not included in a timer started by Python. The console should state its timing basis.
- No runtime target or speedup has been agreed. Review measured opening, transfer, refresh, calculation, saving and verification costs after the baseline.
- Hybrid file reading or wholly Excel-free manipulation remains a separate decision. File-only libraries do not provide Excel's recalculation/required-link-refresh guarantees by themselves.

### Addendum: Console Progress Implementation

28 October 2026

- Implemented configure_console_logging, run_progress, timed_stage and the plain-state waiting reporter in the existing Python module. There is one reporter per CLI run; Event termination and join occur before the Enter prompt, including error/interruption paths.
- Added timings around existing workbook open/close, validation, transfer, refresh, calculation, save, persisted verification, publication and cleanup calls. No file-only data processing or I/O reduction was introduced.
- The CLI reports through stdout instead of invoking modal run-result message boxes. --pause is explicit; default library/CLI/installer calls remain noninteractive. Total time is printed before the prompt and excludes startup/Enter waiting.
- Updated VBA_LAUNCHER to python -u, --pause and shell.Run window style 1 with asynchronous execution. The user saved/closed Draft, ran the backed-up installer manually after the terminal tool stopped executing commands, and confirmed completion. Backup directory now also contains Draft Input spreadsheet for SST_20261028_101609_855714.xlsm.
- Step 1 red phase: 27 failed, 163 passed, 10 skipped. Final default suite: 190 passed, 10 skipped. Twelve dedicated helper tests cover clocks, nesting, failures and reporter lifecycle; fifteen CLI cases cover caller identity and interactive completion. Tests use fake events/clocks, not real five-second waits.
- The temporary test logger needed a current-stdout capture handler to follow pytest's capture lifecycle; this correction is confined to the existing test fixture. It does not change production console behaviour.
- Attempted the isolated export, missing-dependency and installer integration tests. No final result or new console-excel.xml was produced through the terminal tool; do not count these launches as passes. Even a simple terminal output command returned no output. Manual installation was approved and completed, but that does not verify output parity or full-run duration.
- Updated references/bonds-lite-setup.md with stages, waiting semantics, console-only output, Enter completion, timing basis and remaining gates. No persistent run log files were added.
- Remaining: real visible-console acceptance, completed isolated live export/output parity, and a representative intended-data timing baseline. No performance gain or production certification is claimed. Review measured stage durations before any later I/O optimisation.

## Amendment: Native Excel Copy And Two-Open Workflow

28 October 2026 | Implemented; final isolated-export rerun pending | Decisions 16 and 17

### Scope And Implementation

1. Reused the existing module, package, launcher and console reporting. Added native_copy_values to issue Excel Range.Copy and values-only PasteSpecial for contiguous raw-input blocks. Source errors and exact copied values/types are checked inside Excel with scalar SUMPRODUCT results; Python no longer reads, escapes, writes and reads back the million-cell input arrays. The helper requires matching dimensions and the same Excel instance and clears CutCopyMode in finally.

[✓ Implemented — see measured verification below]

2. Preserved destination formatting, formula columns and dynamic-row safeguards. Formula inspection remains timed. Avoided resizing an unchanged table and avoided rewriting all existing formula cells; only newly added rows receive formula fill. Legacy-array/mixed-formula guards remain. Native error-text tests distinguish literal text from an actual Excel error.

[✓ Implemented — row-count and copy-parity checks passed]

3. Reduced orchestration from four Excel opens to two: data and working template copy. The working copy receives schema/identifier checks through populate_bond_data; the original-template inspection open was redundant. Save is followed by Excel Saved-flag and in-memory table/formula/stress/control checks, closure and a nonempty-file check before publication. No separate persisted-file reopen occurs in production; integration tests retain their output reopen/parity assertions.

[✓ Implemented — fast two-open regression passed; final live-export acceptance pending]

4. Corrected a measured three-minute calculation wait: xlPending in manual mode was incorrectly treated as active calculation. Wait only for xlCalculating, keeping explicit worksheet calculations and error validation. Check link statuses after calculation. Per Decision 17, warn on old status 3 after a successful update; reject all other bad statuses and formula errors. Source-workbook freshness remains an upstream responsibility.

[✓ Implemented — 24 refresh/wait policy cases passed]

5. Kept the installed visible asynchronous console launcher unchanged. No workbook/VBA reinstall or new runtime dependency is needed. Native copying uses the Excel clipboard; copy mode is cleared, but prior clipboard contents are not restored. Documented that users should not copy/paste during transfer.

[✓ Implemented — setup guide updated]

### Measured Verification

- Excel-independent suite: 215 passed in 1.04 seconds, with 22 opt-in Excel/performance cases deselected.
- Ten native range-copy cases passed in 18.708 seconds across isolated Excel fixtures. They cover literal text, empty strings, blanks, values/types, formats, source bytes, names/links, real Excel errors, paste failures, scalar-only transfer and copy-mode cleanup.
- Four real-template population cases passed for 1, 5, 998 and 1001 rows. Their 69.33-second test duration includes separate fixture creation and workbook I/O, not just copying.
- Latest same-dataset benchmark: 998 rows x 1,235 columns; old raw Python loop 9.231456 seconds; native raw transfer 0.765760 seconds; complete population 1.131788 seconds. The 24.97-second benchmark test also constructs workbooks, runs both algorithms and checks all raw values. Do not present its test duration or the 12x raw-transfer ratio as a production end-to-end guarantee.
- Earlier benchmark: old raw transfer 11.974216 seconds, native raw transfer 1.096302 seconds, but population 25.251822 seconds due to redundant formula filling/table resizing. The subsequent formula-fill change was followed by the four row-count checks and the latest benchmark above.
- Full isolated export initially failed status 3 at about three seconds; moving status checks after calculation exposed the old 180-second manual-pending wait defect. Two deterministic regressions reproduced/fixed that wait. A separately calculated synthetic source still returned status 3; user-approved warning policy was then implemented and 22 status/formula combinations were added.
- Full export after Decision 17 has not returned a verified final result through the terminal tool. Earlier status-3 failures are not successful publication tests. Final saved-output acceptance, real-network refresh and full production runtime remain open; the validated copy-stage optimisation is not a claim that those gates passed.

### Tests And Safeguards

- Small native-copy tests now use blank workbooks instead of loading the 1.2-million-cell template for tiny grids. Full-template tests and the benchmark retain the real workbook structure.
- The source verification expression deliberately avoids IFERROR inside Evaluate: Excel added an _xlfn.IFERROR compatibility name during the first regression. Native comparison now introduces no such source-name change.
- Test-only old Python loop is retained inside the opt-in benchmark for fair comparison, never as a production fallback.
- Real network tests remain separately opt-in. No original data/template files were edited and no dependencies are refreshed just by running ordinary unit tests.
- There is no promise that changing the copy mechanism accelerates Excel's external link updates or model calculation. Console stages continue to expose those costs separately.

## Amendment: Raw ATT Boundary And Destination Formula Ownership

28 October 2026 | Implemented | Decision 18

### Workbook Rescan

- Rescanned both actual files read-only using ZIP table metadata and openpyxl formula seeds; no Excel opening, calculation or saving was needed for discovery.
- Data table I_tbl_Monthly_Bonds is B8:ATT1066, with 1,215 headers and 1,058 records. Input table remains B8:AUN1006, with 1,235 headers and 998 seed records. The source headers exactly match the destination's first 1,215 headers.
- Both files retain formula seeds in J:K. The destination alone retains the 20 calculated columns ATU:AUN. Destination ATV:AUD now look up values on P. Bond Escalation file rather than in the previous external escalation workbook.
- P. Bond Escalation file spans A1:O61561 and contains zero formulas in the scan. Its network-path text is metadata, not a required external formula reference. No required external formula dependencies were found in the revised bond calculation chain.

### Implementation And Checks

1. Added validate_bond_headers with an exact B:ATT source width and strict destination-prefix comparison, leaving generic validate_headers unchanged. Require both tables to begin in column B; reject the former full-width AUN source format and malformed, missing, reordered or additional raw columns.

[✓ Implemented — 27 schema cases passed]

2. Bound all native-copy intervals to ATT and exclude every destination formula column. The supplied layout copies B:I and L:ATT, preserving target J:K and ATU:AUN. Require any destination columns after ATT to be formula-owned; no source access beyond ATT occurs. Existing dynamic table resizing and new-row-only formula filling continue to operate on the full destination width.

[✓ Implemented — default suite and isolated export passed]

3. Keep the data source read-only with link updates disabled and manual calculation while it is open. Retain formula inspection/filling and refresh_and_validate only against the working template copy. Source J:K formulas are ignored rather than recalculated or copied. Original input template remains a filesystem copy source, never an Excel calculation target.

[✓ Implemented — isolated source UpdateLink/Calculate guards passed]

4. Adapted existing tests and fixtures to raw-width sources and full-width targets. The isolated escalation fixture now writes known values only to the temporary template's local reference sheet without rewriting its lookup formulas. Separate injected-local-dependency tests retain external update/missing-file coverage without reaching production network paths. Revised row-count cases cover 1, 5, 998, 1001 and 1058 records; the benchmark explicitly distinguishes source and destination widths.

[✓ Implemented — full revised live matrix not rerun]

### Verification And Remaining Risk

- Read-only supplied-template layout check: 1 passed in 1.35 seconds.
- Header-contract tests: 27 passed in 0.09 seconds; initially failed on the missing helper before implementation.
- Complete Excel-independent suite: 243 passed in 0.94 seconds, with 24 opt-in cases deselected.
- User-run isolated full export: 1 passed, 266 deselected, 11.82 seconds including setup/teardown. This checks raw-only source handling, destination lookup values, source formula/cache snapshots, original-file immutability and guards against source refresh/calculation.
- Windows native RPC diagnostics 0x800706be and 0x800706ba appeared after the passing test assertion result during cleanup. Do not claim a clean Excel-lifecycle run; the fixture shutdown issue remains separate from the verified export behaviour.
- No new runtime dependency, workbook/VBA update or button reinstall is required. No new performance claim is made for the revised workbook layout. Earlier speed measurements remain historical.

## Amendment: Stage 1 Audit And Stage 2 Handoff

28 October 2026 | Implemented | Decision 19

### Approved Scope

Latest-run panel only beside the existing button, within rows 73-83; no history sheet. Automatically save Draft at run start and terminal audit commit, including other pending edits. Record Windows login, full start/finish timestamps and monotonic duration. Invalidate the active Stage 2 input before a new accepted attempt; preserve a separate last-successful-output link and all earlier files. Store relative hyperlinks for colleague portability. Stage 2 execution is deferred.

### Steps And Progress

1. Reused tests/test_bonds_lite_input.py and existing fixtures. Added failing audit, guard, save-recovery and portable-path contracts; adapted main's stubs to the new run_stage_one boundary without weakening caller identity tests. Kept public generator immutability tests separate from caller-audit persistence.

[✓ Implemented — initially missing APIs produced red tests]

2. Read-only preflight verified H73:O83 was empty, unprotected and did not overlap the D74:F76 button or existing picture. Added an idempotent initialiser that rejects occupied cells/formulas/merges/protection/shapes and wrong/incomplete names. Installed the coral/pink panel with wrapped path rows and date/time/duration formats; did not widen existing columns or add a sheet.

[✓ Implemented — blank-workbook layout/persistence and actual saved-panel checks passed]

3. Added STAGE_ONE_FIELDS, write_stage_audit and named fields SST_Stage1_Panel, SST_Stage1_Status, SST_Stage1_LastSuccessfulOutput and SST_Stage2_Input. Audit writes are restricted to mapped cells/hyperlinks, suppress and restore events, save Draft and check Saved. Dates remain full timestamps even when displayed as date/time; duration supports midnight and more than 24 hours.

[✓ Implemented — field, event restoration, save-error and duration tests passed]

4. Added run_stage_one as the CLI's audited entry point. One workbook guard spans initial audit mutation, generation and final audit save. Split the generation body into a guard-free private backend so the wrapper does not take a nested mutex; direct public generation remains guarded and does not save audit data. Start commits Running with a cleared handoff; terminal failures/interruption preserve last success and re-raise the original error. Final audit-save failure retains the published output, attempts to clear/recover the handoff, and reports failure instead of success.

[✓ Implemented — lifecycle, original-error, duplicate and output-retention cases passed]

5. Added portable_output_path and resolve_stage_input. Require a saved successful audit, matching active/last-successful paths and an existing nonempty .xlsx directly in Bonds Lite output. Reject absolute paths, traversal, missing files and temporary working files. The resolver does not invoke a model. Future Stage 2 should take the same stage guard across its complete run; this helper does not provide a cross-user/coauthoring lock.

[✓ Implemented — 18 handoff-state cases and cross-user path cases passed]

6. Extended the existing backed-up installer and installed the panel into the actual Draft workbook. Backup: references/backups/Draft Input spreadsheet for SST_20261028_124921_052826.xlsm. Existing button, input values and model-generation logic remain. The embedded command is unchanged; main now routes to the audit wrapper. Updated setup instructions and Decision 19.

[✓ Implemented — actual installer completed; saved layout verified read-only]

### Verification

- Complete default suite: 287 passed, 1 skipped, 26 deselected in 1.14 seconds.
- Audit lifecycle/direct-generation/CLI slice: 29 passed before the additional saved-state and duration cases.
- Portable-path/handoff initial slice: 27 passed; final handoff-state slice: 18 passed, including unsaved audit rejection.
- Midnight and more-than-one-day duration cases: 2 passed.
- Blank-workbook panel integration: 1 passed in 8.32 seconds; temporary Draft installer: 1 passed in 5.90 seconds.
- Actual workbook preflight: 1 passed in 0.74 seconds before install. Actual installed named ranges, colours, number formats and status verified read-only: 1 passed in 0.76 seconds.
- User-run full audited export: 1 passed in 11.33 seconds including fixture setup/teardown. It generated an output, saved/reopened the audit, resolved the handoff, then failed a new attempt and confirmed the handoff was blank while the previous successful link/file and data/template sources remained unchanged.
- Native RPC diagnostics 0x800706be/0x800706ba appeared in live test shutdown, including after the passing audited-export assertions. These remain an unresolved Excel teardown issue, not a clean lifecycle certification.
- No Stage 2 subprocess, Stage 2 button or all-stage runner was added. The supplied model command and portable-directory proposal are documented for the next change, not executed now.

### Operational Notes

- The installed panel starts Not Run and does not adopt an earlier file merely because it is newest in Bonds Lite output. The next button run populates its audit and handoff.
- Stage work duration excludes Python startup, final audit-save time and the Enter wait; console total is the broader metric.
- Latest-run metadata is editable and not an immutable history. The user rejected a Run History sheet.
- Draft auto-save was explicitly approved; it saves unrelated pending user edits too. If audit recovery cannot be saved, the workbook on disk may retain an earlier record; the console reports this instead of claiming success.
- Existing runtime-model inputs and source-only versus destination-formula ownership are unchanged.

## Amendment: Movable And Resizable Audit Panel

28 October 2026 | Implemented | Decision 20

### Implementation

1. Added STAGE_ONE_FIELD_NAMES for all eight audit fields and expanded STAGE_ONE_NAMES to include those plus the panel boundary. STAGE_ONE_FIELDS and STAGE_ONE_PANEL remain initial-install defaults only. Runtime audit reads/writes resolve current named ranges through _stage_field_cells and use the top-left cell of each current merged value area.

[✓ Implemented — eight named-routing unit cases passed]

2. Replaced fixed-coordinate validation with ownership and geometry checks: writable/unprotected caller, names in this workbook on I. Control, one rectangular panel, logical single-cell/merged field regions entirely within the panel, and no overlapping fields. Invalid/missing references fail before writes; there is no fallback to the initial cells.

[✓ Implemented — live invalid-layout cases passed]

3. Made initialisation idempotent for a fully named moved/resized panel. Added guarded migration for the existing four-name layout at its current location, including a whole-panel move before upgrade. Infer five missing field anchors only for the original 11 x 8 relative layout with coherent retained field pointers and matching labels. Add names only; preserve values, styles, hyperlinks and current position. Roll back newly added names on migration failure.

[✓ Implemented — six legacy migration cases passed]

4. Updated the actual Draft via the backed-up installer after the user saved and closed it. Backup timestamp: 20261028_140917_887382. The user confirmed the read-only saved-name verification passed. Did not move the real panel or reset prior audit values.

[✓ Implemented — actual workbook upgrade verified]

5. Updated usage instructions for selecting SST_Stage1_Panel, Cut/Paste to an empty area on I. Control, resizing rows/columns and maintaining names during structural redesign. Stage 1 generation, relative-path handoff, console and auto-save policies remain unchanged.

[✓ Implemented — setup guide updated]

### Verification And Caveats

- Default suite: 295 passed, one skipped, 48 deselected in 1.58 seconds.
- Movement/resizing tests cover Cut/Paste, writes that leave the old area unchanged, row/column insertion, merged-value resizing, repeated initialisation, save/reopen and handoff reading.
- Invalid-layout checks cover missing/#REF names, external ownership, multi-area/unmerged ranges, overlap and out-of-panel fields. The first combined live run passed 19 cases and then stopped on an overly narrow error-message expectation, not unsafe acceptance. The error now identifies the audit panel; all six migration cases subsequently passed in the user's focused rerun.
- Corrected a test snapshot to compare Style.Name instead of COM Style proxy identity; this does not change production formatting.
- Actual saved-workbook verification was supplied by the user after integrated terminal execution became unreliable. Do not treat the earlier output-free command as installation success; the later saved-name check establishes the upgrade.
- Existing native Excel RPC teardown diagnostics remain unresolved. The change does not promise arbitrary copying, moving to another sheet, deleted names, or resizing during a running stage. Copy/Paste duplicates do not transfer the named panel identity; use Cut/Paste.

## Plan: Paired Bonds And MAP Stage 1

28 October 2026 | Implemented — see paired-run addendum | Decisions 21 and 22

### Overview

Stage 1 will require the Bonds and MAP datasets, apply the same input template and captured controls to each, and return a completed timestamped run folder. Preserve Python orchestration, native Excel transfer, destination-only calculation, console timings, audit autosave and the movable named panel. Stage 2 execution remains outside this change.

### Read-Only Workbook Evidence

- Draft I. Control!D39 contains Bonds Lite data template.xlsx and C39 labels the Bonds asset source.
- C40 labels the MAP bonds asset source, so D40 is the correct new input cell. It is currently blank; the user chose to enter its path manually.
- Inserting the new input row moved the shared template source to D42, containing Bonds Lite input template.xlsx. D41 is blank. The implementation and installer must not continue using or populating D41.
- Both supplied datasets contain I_tbl_Monthly_Bonds at B8:ATT1066: 1,215 columns. The input template retains its B8:AUN1006 table and destination-owned formulas. Runtime prefix validation remains mandatory for each source.
- The input template defines output_folder_path as 'I. Control'!$D$19. Resolve this name when writing; do not hard-code D19.
- Entity remains D9 (currently JRL), valuation date D11, and credit stresses are shared. The audit panel is now named at H72:O82, not its original H73:O83; all updates must continue using names rather than initial coordinates.
- No architecture reference file is present. Discovery used actual workbook table XML and read-only openpyxl control/name extraction; neither templates nor Draft were saved or recalculated.

### Agreed Output Contract

For entity JRL, one successful attempt produces:

```text
Bonds Lite output/
  20261028_153118/
	 Bonds Lite input_Bonds.xlsx
	 Bonds Lite input_MAP.xlsx
	 JRL Bonds/
	 JRL MAP/
```

- Timestamp format: YYYYMMDD_HHMMSS, with a collision suffix such as _001 when needed. Never overwrite a preceding run.
- Each run uses its own pair of empty model-output folders. Do not execute the model or create Stage 2 results in this change.
- The Bonds input's output_folder_path points to the final current-user absolute path of <run>/<entity> Bonds; the MAP input points to <run>/<entity> MAP.
- The audit's Last Successful Output and active Stage 2 Input store the relative run-folder path, for example Bonds Lite output/20261028_153118, not a single workbook path. Their hyperlinks open the run folder.
- A successful Stage 1 means both inputs have been generated, refreshed, saved and verified. Failure or interruption of either dataset removes only this attempt's newly created directory and blocks the handoff; previous runs and last-successful audit links remain intact.
- Preserve the existing separate recovery policy when both outputs are valid but final Draft audit persistence fails: retain the complete pair, report its folder, and keep Stage 2 unauthorised. Do not confuse a retained complete pair with an accepted partial generation.

### Steps

1. Write failing tests first in tests/test_bonds_lite_input.py. Extend the existing filesystem, generator, audit and isolated Excel fixtures for two distinct datasets and a run-folder return. Use different identifiers, PVs and row counts so accidental reuse of the Bonds dataset for MAP is detectable. Add cases for missing D40, template D42 versus blank D41, aliases between caller/data/template files, unsafe entity folder characters, timestamp collisions, first/second-copy failure and named output-folder assignment.

	Verification: failures identify the absent two-dataset contract, not invalid fixture setup. Default unit tests do not start Excel, refresh network data or modify real workbooks.

2. Refactor validate_inputs/resolve_paths in src/bonds_lite_input.py to require D39, D40 and D42. Keep filename-relative-to-Draft and absolute local/UNC path support. Capture entity, valuation date and all credit stresses once per run; pass that same snapshot to both generations, synchronisation and verification. Reject missing inputs and path aliases before creating output artifacts. Validate entity text for Windows filename-invalid characters without silently changing the entity.

	Verification: both sources are required; D41 is ignored; relative paths resolve for different users; caller, both datasets and template are distinct files; subsequent edits to live controls cannot produce different entities/dates/stresses between the pair.

3. Add a collision-safe run-directory allocator under Bonds Lite output. Atomically reserve the final timestamped directory with mkdir(exist_ok=False), retrying only name collisions; record ownership of this attempt. Create <entity> Bonds and <entity> MAP inside it. Use private temporary working files inside the owned run directory, not staging names that would later invalidate embedded absolute paths.

	Verification: earlier directories and their bytes remain unchanged; both model-output folders have exact requested names; missing/invalid inputs create no completed-looking run. The audit remains Running with no active handoff while files are being prepared; do not claim atomic directory publication merely because the final directory is reserved early.

4. Extract a reusable per-dataset helper from _generate_bonds_lite_input. It receives the Excel application, selected source path, shared template path, captured settings, temporary working path, intended input filename and final model-output folder. Process Bonds and MAP sequentially: open only the current source, make/open one template copy, populate raw data, preserve/fill destination formula columns, populate shared stresses, set controls and the named model-output path, refresh/calculate only that working copy, save, verify and close. Close the first source/copy before opening MAP. Use private working filenames that do not collide with previous fixed-name output workbooks a user may have open; filesystem rename after closure preserves the existing low-I/O pattern.

	Verification: each workbook contains the correct dataset; source UpdateLink/Calculate guards never fire; raw import still ends at ATT; destination J:K and ATU:AUN remain formulas; no original input/template saves; no unnecessary template inspection or output verification reopen in production.

5. Add a small set_model_output_folder helper and extend saved-copy verification. Resolve working.names['output_folder_path'].refers_to_range and require a local single value cell/merged area on the working copy's I. Control. Write the absolute final dataset-specific output-folder path before refresh/save. Check the named value after saving, along with row count, formula coverage, date/entity and all stress values. Use the snapshot rather than rereading caller controls.

	Verification: Bonds points to <run>/<entity> Bonds and MAP to <run>/<entity> MAP. Paths include the final reserved run folder, never a private temporary directory or another user's hard-coded profile. Missing, foreign, broken or multi-area output_folder_path names fail the attempt without editing the original template.

6. Make the private Stage 1 backend complete and validate both fixed input files before returning the run-folder Path. On any generation failure or KeyboardInterrupt, close owned workbooks, restore Excel settings, remove private working files and the entire newly owned attempt directory, then preserve the original error. Never remove the output root, a prior run or any directory not created by this attempt. Keep the existing workbook guard across both datasets, cleanup and audit commit; do not acquire a nested guard per dataset.

	Verification: inject failures before creation, during Bonds, during MAP, refresh, save and verification. Every incomplete attempt is removed, prior outputs survive and the Stage 2 handoff is empty. A final audit-save failure after successful backend return retains the complete pair under the existing recovery rule.

7. Update portable_output_path, run_stage_one and the handoff reader from single-file to run-folder semantics. Audit success requires both fixed .xlsx files to be nonempty within the same direct-child run directory. Store one relative folder link through the existing named audit cells; do not expand, reposition or reset the movable panel. Add resolve_stage_inputs returning an explicitly labelled Bonds/MAP pair, optionally with its run folder. Keep resolve_stage_input only with an explicit run-folder contract or a clear deprecation/ambiguity error; it must never silently select Bonds. Require a saved Succeeded audit, matching active/last-successful links and a complete pair. Reject legacy flat-file handoffs with an instruction to rerun Stage 1.

	Verification: the two paths resolve under Joao and Sarah folders; a missing/empty member, traversal, temporary directory, unsaved/failed/running audit or mismatched run folder is rejected. No Stage 2 subprocess is launched. If a different user later runs Stage 2 against another user's prepared inputs, that future controller must rebase or validate embedded absolute model-output paths before execution; the relative audit alone does not rewrite workbook contents.

8. Update install_button and references/bonds-lite-setup.md. Preserve entered D39/D40/D42 values; keep D40 blank if blank, and place the shared-template default only in D42. Remove assumptions that D41 is the template cell. Any workbook changes go through a saved/closed, backed-up installation; keep the relocated audit names and existing user controls. Document both required datasets, run-folder layout, named output paths, all-or-nothing generation and the separate complete-pair audit-save recovery rule. Append implementation progress and verification results to this plan after code changes.

	Verification: installer does not fill D40 or D41, create a second template control, reset audit data or move the panel. Setup examples use portable filenames and distinguish Stage 1 inputs from the two model-output directories.

### Test Strategy

- Reuse tests/test_bonds_lite_input.py, existing temporary-workbook fixtures, fake clocks and audit boundary fakes. No new test framework or large production-data copies.
- Unit-first: two-source path validation; D42 control mapping; one shared snapshot; Windows-safe entity folders; collision allocation; complete-pair folder validation; rollback ownership; run-folder audit; persisted success/failed-rerun/initial-save/final-save cases.
- Isolated Excel: distinct Bonds and MAP source rows with different counts, both using the same actual template copy and known local escalation values. Reopen both saved inputs in the test and compare raw values, formulas, stresses, controls and each named output folder. Compare original data/template bytes before/after.
- Failure tests: invalid MAP schema or formula/save failure after a successful Bonds copy must remove the attempt directory and leave prior runs untouched. Keep bad-source guards and copied-text/type tests.
- Moved/resized audit: retain current named-range tests and verify the run-folder handoff after save/reopen without depending on initial cell positions.
- Performance: retain console timing per dataset and measure a representative paired run separately from fixture creation. Two inputs necessarily add real transfer/save work; do not reuse the previous single-dataset speed numbers as a paired-run guarantee.
- Ordinary test execution must not access the real network or run the Bonds Lite model. Record native RPC teardown diagnostics separately from assertion results; absent terminal output is not a pass.

### Affected Files

- src/bonds_lite_input.py: control snapshot, run allocation, dataset helper, named model-output assignment, paired generation/rollback, folder audit and paired handoff reader.
- tests/test_bonds_lite_input.py: update single-source/file assumptions and add paired-success/failure/output-path cases.
- references/bonds-lite-setup.md: D39/D40/D42, new tree, folder-link semantics and execution-path portability.
- references/decisions.md: paired-run and audit choices, retaining prior decision history.
- Draft Input spreadsheet for SST.xlsm: existing D40 control is already correctly labelled; no automatic MAP value by user choice. Preserve moved panel. No second template control.
- Data/template workbooks: read-only sources, never implementation edit targets.

### Boundaries And Open Verification

- Implemented: current code reads D39/D40/D42, generates a pair and exposes the completed run folder. Verification evidence and remaining limits are recorded below.
- D40 is currently blank, so a future two-dataset run must refuse to start generation until the user supplies it. The available filename is Bonds Lite data template_MAP.xlsx.
- Source headers were scanned for widths/table names; compare their full prefixes at runtime and in the paired fixtures. MAP row count need not match Bonds.
- Stage 2 execution, a run-all-stages command, immutable run history and cross-machine coauthoring locks remain excluded.

### Addendum: Paired Run Implementation

28 October 2026

- Step 1: implemented paired test fixtures and red contracts in the existing test module. Initial default run exposed 146 expected failures against the old single-output implementation; no production sources were modified by the test agent.
- Step 2: added frozen StageOneInputs, four-file resolution and validate_entity. D39/D40/D42 are read once, D41 is ignored, all source/template identities are distinct. Source credit header layouts and numeric stresses are captured/validated once; destination header checks remain per copy. Date/entity synchronisation and verification consume the captured snapshot.
- Step 3: added allocate_run_folder with atomic mkdir reservation, timestamp collisions and numeric suffixes. Final run/model-output paths are reserved before writing; no staging path leaks into workbook output_folder_path.
- Steps 4-5: extracted _generate_dataset_input, using one source and unique private template copy at a time. Native raw transfer and destination-only formula refresh remain. Both files receive the same captured controls/stresses, their own input data and their matching absolute model-output directory via the template's defined name. Verification checks the saved named path too.
- Step 6: the backend validates both files before returning the run folder and rolls back only its owned attempt on failure/interruption. Filesystem cleanup errors are logged without concealing the original generation error; an incomplete leftover folder cannot pass paired handoff validation. Complete-pair final-audit-save failure retains both files under the prior recovery policy.
- Step 7: portable_output_path and resolve_stage_input now validate a complete timestamped folder; resolve_stage_inputs returns a labelled Bonds/MAP dictionary. Legacy flat-file handoffs, missing/empty members, foreign member symlinks, traversal and temporary paths reject. Existing audit names, moved panel geometry, autosave and guard rules remain unchanged.
- Step 8: installer defaults use D42 and never fill D40 or D41. No actual workbook edit/reinstallation was required because the user already inserted/labelled D40 and moved the shared template path to D42. Setup documentation now describes the paired tree and folder handoff. D40 remains blank until the user supplies it.

### Paired Verification Results

- Final default suite: 421 passed, 8 skipped, 63 opt-in Excel/performance cases deselected in 4.13 seconds. Skips are environment/layout-dependent; they are not counted as passes.
- Four-file/snapshot/entity slice: 77 passed, 6 skipped. Paired folder/audit slice: 87 passed, 1 skipped. Dataset-helper/rollback/missing-control slice: 24 passed. Final cleanup/CLI slice: 23 passed.
- Ten isolated output_folder_path-name cases passed, covering correct/merged values, broken/missing/foreign/multi-area names and invalid output folders.
- First paired export reached saved lookup assertions and failed only because xlwings converted date serials into datetime objects. The fixture now compares raw Excel Value2 values, preserving actual workbook date formatting.
- The user confirmed the corrected isolated full paired export and paired audited-run/failed-rerun tests passed. These use different Bonds/MAP records and row counts, verify both output directories and named paths, reopen saved copies, and retain source bytes and previous completed runs.
- Integrated terminal execution remained intermittent; manual user-provided pytest results were used where automated launches returned no result. Those output-free launches were not counted as passes.
- No Stage 2 model was executed, no real network refresh was requested by the isolated tests, no original data/template files were saved, and no new runtime dependency was added.
- Historical single-dataset benchmark figures are not a paired-run speed guarantee. A representative full production pair timing and native Excel RPC teardown investigation remain separate follow-ups.

## Next Stage: Bonds Lite Batch Execution

Stage 2's user-confirmed implementation plan is saved in references/plan-stage-two.md. It uses the committed Stage 1 run folder with one external run_bonds_batch.py invocation, a separate py311 interpreter, a Run Bonds Lite button, named audit panel and transactional output recovery. See Decisions 23-25. This is planned, not implemented; no model run was performed during planning.