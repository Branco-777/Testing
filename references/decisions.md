# Bonds Lite Planning Decisions

Requirements confirmed on 27 October 2026. Implementation decisions added on 28 October 2026; see the plan addendum for verified progress and open gates.

### Decision 1: Planning And Collaboration (2026-10-27)
**Question:** What planning detail and review cadence should apply?
**Decided by:** User
**Choice:** One step-by-step plan, inline verification, prose only; stop at every meaningful decision; allow up to 40 further questions; apply no team conventions.
**Alternatives considered:** High-level or separate phase documents; autonomous work between milestones; team-specific conventions.
**Rationale:** These are the user's explicit session preferences.

### Decision 2: File Paths And Button (2026-10-27)
**Question:** How should inputs be chosen and the workflow started?
**Decided by:** User
**Choice:** Generate Bonds Lite Input button on Draft I. Control; data path in D39 and template path in D41, accepting absolute paths or filenames beside Draft.
**Alternatives considered:** File picker on each run; always using a fixed data filename; different control cells.
**Rationale:** User wants persistent cell-based file selection and approved the existing source rows.

### Decision 3: Stress Mapping And Units (2026-10-27)
**Question:** Which stress blocks are exported and are units transformed?
**Decided by:** User
**Choice:** Copy all four transition matrices and both spread grids for each currency/sector group to matching positions in the actual I. Sensitivity sheet. Use values unchanged; spreads are decimals.
**Alternatives considered:** Basis-point inputs; transformations or alternative mappings.
**Rationale:** Workbook layouts match, and the user confirmed unchanged decimal inputs.

### Decision 4: Missing Inputs And Validation (2026-10-27)
**Question:** What should happen with blank stress inputs or invalid data?
**Decided by:** User
**Choice:** Stop and identify missing cells. Reject empty bond data, changed/missing headers and failed required dependencies/calculations before publishing output. Preserve original data and template files.
**Alternatives considered:** Retain destination example values; substitute identity transitions and zero spreads; permissive schema handling.
**Rationale:** Silent substitutions or schema drift could generate misleading inputs.

### Decision 5: Bond Import And Dynamic Sizes (2026-10-27)
**Question:** Should the import copy formulas or values, and where should row counts vary?
**Decided by:** User
**Choice:** Copy raw inputs and preserve/recalculate destination formula columns. Resize only the destination bond table for varying row counts; keep Draft as controls and stresses.
**Alternatives considered:** Copy evaluated values for the whole table; replace the whole sheet; add a bond-data table to Draft; support changing column counts.
**Rationale:** User requires template formula behaviour and variable-sized datasets without duplicating bond data into Draft.

### Decision 6: Dependencies And Control Synchronisation (2026-10-27)
**Question:** Should external dependencies refresh and which controls should change?
**Decided by:** User
**Choice:** Refresh required links and stop on unavailable dependencies/calculation failures. Copy Draft date/entity D11/D9 to destination D10/D13; leave other controls unchanged.
**Alternatives considered:** Cached external values without refresh; leave date/entity unchanged; update additional controls.
**Rationale:** User requires recalculated formulas and consistent date/entity without expanding the scope to other template assumptions.

### Decision 7: Single Output And Retention (2026-10-27)
**Question:** How many outputs are generated and how are repeat runs handled?
**Decided by:** User
**Choice:** One input workbook per click using all credit blocks, ignoring individual/combined-stress toggles. Create Bonds Lite output beside Draft, timestamp each saved copy and retain every run. Close the generated workbook and show the saved path.
**Alternatives considered:** Scenario/toggle-driven exports; fixed-name overwrite prompts; leave generated workbook open.
**Rationale:** User confirmed a single-input generation workflow and retention of repeat runs.

### Decision 8: Verification Data And Runtime (2026-10-27)
**Question:** What fixtures and performance target should guide verification?
**Decided by:** User
**Choice:** Supplied data is typical but row counts vary. Use temporary test copies with distinctive synthetic stress values and exact copy checks. Measure typical-size runtime before setting a target; no completed expected output is available.
**Alternatives considered:** A supplied completed golden workbook; an immediate fixed runtime requirement.
**Rationale:** Tests must cover variable size, and the user approved synthetic test stresses and baseline measurement.

### Decision 9: Technology And Deployment (2026-10-27)
**Question:** Should implementation use native VBA or Python, and what deployment is acceptable?
**Decided by:** User
**Choice:** Python with xlwings plus a minimal VBA button launcher that starts a configured Python executable without an Excel add-in. Target existing Anaconda environments and multiple colleagues on Windows desktop Excel.
**Alternatives considered:** Native VBA only, recommended initially for simpler deployment but not selected; standard xlwings add-in/RunPython, explicitly rejected; project-local virtual environment, not selected; standalone openpyxl, unsuitable for required Excel recalculation.
**Rationale:** User chose Python-backed implementation and an add-in-free launcher after reviewing the deployment trade-offs.

### Decision 10: Persistent Planning Documents (2026-10-27)
**Question:** Should the plan be saved or only presented in chat?
**Decided by:** User
**Choice:** Save the detailed plan and confirmed decisions under references without modifying workbooks.
**Alternatives considered:** Chat-only plan; editing the workbook during this planning session.
**Rationale:** User approved persistent documents; current mode permits planning, not production implementation.

## Implementation Proposals Pending Review

- Per-user interpreter configuration via BONDS_LITE_PYTHON, exact timestamp filename format and collision suffix, and the run-guard mechanism remain technical proposals for the implementation checkpoint, not additional user-approved requirements.
- Keep launch asynchronous and attach to the exact calling Excel instance; verify the supported mechanism with the installed xlwings version before implementation.
- Determine dependency-specific refresh verification, array/spill formula handling, supported interpreter range and live-Excel test prerequisites before certifying the workflow.

### Decision 11: Interpreter Configuration And Concurrent Runs (2026-10-28)
**Question:** Approve the per-user interpreter setting, asynchronous launcher, named mutex and collision suffixes?
**Decided by:** User
**Choice:** Configure BONDS_LITE_PYTHON for the user's existing Anaconda python.exe. Use the asynchronous, add-in-free launcher with a per-caller Windows named mutex and timestamped filenames with numeric collision suffixes.
**Alternatives considered:** Other interpreter configuration or concurrency mechanisms; no alternative was requested after the implementation checkpoint.
**Rationale:** Colleagues can use their own interpreter without personal paths embedded in VBA; concurrent clicks and repeat filenames must not damage outputs. This approves the corresponding proposals listed above.

### Decision 12: Isolated Dependency Verification (2026-10-28)
**Question:** Continue verification using a temporary local escalation workbook when the initial real-dependency workflow failed at UpdateLink?
**Decided by:** User
**Choice:** Continue with a synthetic local dependency and remap links only in temporary test template copies. Keep the real-network test separately opt-in. Production retains required-link refresh and stops on unavailable dependencies or calculation errors.
**Alternatives considered:** Pause until a corrected real dependency was supplied; verify only against the real dependency.
**Rationale:** Isolated tests distinguish implementation faults from network/data issues without weakening production checks or changing the supplied source/template files.
**Subsequent evidence:** The same error occurred locally and was traced to opening a source before calling UpdateLink. Closed-source refresh succeeded in the diagnostic; the implementation was repaired accordingly. Do not infer that the real network workbook is defective. Full post-repair workflow verification remains outstanding due to terminal execution/output failures.

### Decision 13: Measure Before Changing Excel Processing (2026-10-28)
**Question:** Replace the Excel manipulation with a hybrid/file-only approach, or keep the current processing and add timings first?
**Decided by:** User
**Choice:** Keep the existing processing and add timings first. No measured current runtime or acceptable runtime target is available; a prior run was interrupted because the user could not tell whether it was still running.
**Alternatives considered:** Read data directly and open only a working copy in hidden Excel; avoid Excel for generated files while preserving formulas without recalculating.
**Rationale:** Visibility and a measured baseline are needed before attributing cost to workbook opening or changing formula calculation and external-link behaviour. Background visibility alone is not evidence of faster processing.

### Decision 14: Console-Only Stages And Completion Prompt (2026-10-28)
**Question:** Where should progress be shown, and should it remain visible after completion?
**Decided by:** User
**Choice:** Console only, with run stages and elapsed-time information. Keep the console open until Enter after either success or failure. Do not add log files.
**Alternatives considered:** Console plus persistent logs; optional DEBUG detail; automatically closing after success.
**Rationale:** The user wants to see whether a run is active and where time is spent without producing extra files. The implementation plan proposes stage boundaries and waiting messages; heartbeat interval and exact interactive CLI flag remain implementation details for review.

### Decision 15: Waiting Interval And Launcher Installation (2026-10-28)
**Question:** Approve a five-second waiting message and a backed-up embedded launcher update with Draft saved and closed?
**Decided by:** User
**Choice:** Five-second waiting messages with stage/total elapsed time and no estimated percentage. Save and close Draft before updating its launcher. When the terminal tool stopped executing commands, the user ran the supplied installer in their own PowerShell terminal and confirmed completion.
**Alternatives considered:** A different waiting interval; leaving Draft open and providing the launcher update separately; claiming installation from the changed Python string alone.
**Rationale:** Five seconds provides visibility during long Excel calls without continuous output. The workbook must contain the updated visible-console VBA, and a backup must precede that change. Manual installation resolved the tool-execution blocker without changing Excel security policy or workbook processing.

### Decision 16: Excel-Native Transfer With Python Orchestration (2026-10-28)
**Question:** Implement native Excel copying while retaining Python for future orchestration stages?
**Decided by:** User
**Choice:** Replace Python-mediated raw-cell transfers with native Excel values-only copy/paste. Keep formula refresh, dynamic rows and console timings; remove the redundant original-template open and saved-output reopen. Permit rewriting the existing implementation where needed.
**Alternatives considered:** Retain bulk arrays through Python; create a separate VBA copy engine; use file-only processing without Excel recalculation.
**Rationale:** The user wants Python orchestration without unnecessary data-marshalling or workbook-opening overhead. Excel can copy ranges within its own process. Scalar Excel-side comparisons retain full value/type checks without bringing every copied value into Python.
**Measured evidence:** On the same prepared 998 x 1,235 test dataset, the user measured old raw transfer 9.231456 seconds, native raw transfer 0.765760 seconds and complete optimised population 1.131788 seconds. This is roughly 12 times faster for raw transfer, not a measured whole-run improvement. Initial redundant formula filling made population take 25.251822 seconds; avoiding unchanged-row resize/formula rewrites removed that cost.

### Decision 17: Warn On Excel Old-Link Status After Successful Refresh (2026-10-28)
**Question:** What should happen when a successful UpdateLink and calculation still return status 3 for a freshly calculated test dependency?
**Decided by:** User
**Choice:** Print a visible warning for status 3 and continue only if the explicit update succeeded and calculated formulas are error-free. Missing files, update exceptions, other bad link statuses and formula errors remain fatal.
**Alternatives considered:** Keep status 0 mandatory and leave export blocked for further diagnosis; silently ignore all status failures.
**Rationale:** A fully calculated synthetic source still received status 3, so that indicator alone did not discriminate freshness reliably in this workflow. The user explicitly accepted the limitation that source formulas must be maintained current upstream. This supersedes the status-0-only implementation gate, not the requirement to perform link refresh and formula validation.

### Decision 18: Raw Source Ends At ATT; Formulas Belong To Output (2026-10-28)
**Question:** Adapt to a data source containing only columns through ATT and perform all formula refresh in the generated Bonds Lite input?
**Decided by:** User
**Choice:** Require the raw B:ATT source schema and compare it with the destination header prefix. Preserve and calculate destination-owned formula columns only in the working input copy after raw import. Never explicitly update links, calculate worksheets or fill formulas in the source workbook; do not save either original template.
**Alternatives considered:** Continue requiring identical full-width source/destination tables; copy or recalculate source formula columns; import calculated tail columns from the source.
**Rationale:** The source no longer contains ATU:AUN, and the input template now owns those formulas and its local escalation reference data. Retaining full-width validation or sourcing calculated values from the data workbook would violate the new ownership boundary.
**Observed layout:** Source B8:ATT1066, 1,215 columns; destination B8:AUN1006, 1,235 columns. Raw headers match exactly. J:K remain formulas in both files and are excluded from source copying. Destination ATV:AUD reference P. Bond Escalation file; its scanned contents are values, not formulas, and no required external formula dependencies were found for the revised calculation chain.

### Decision 19: Stage 1 Audit Panel And Portable Stage 2 Handoff (2026-10-28)
**Question:** How should development-stage runs expose audit metadata and hand off Stage 1's output before adding the model runner?
**Decided by:** User
**Choice:** Add a latest-run panel beside Generate Bonds Lite Input within rows 73-83. Use the screenshot's coral/pale-pink visual style; record date, start, finish, duration, Windows login and status. Retain a separate last-successful-output link; clear the active Stage 2 input at the start of an accepted attempt and leave it blank after failure/interruption. Automatically save Draft, including unrelated pending edits. Use relative clickable paths for different users' OneDrive folders. Do not add a Run History sheet or implement Stage 2 execution yet.
**Alternatives considered:** Right-side panel near row 9; latest panel plus append-only history; manual audit saves; retaining an active prior handoff after a failed rerun; absolute per-user paths; Excel's editable displayed username; implementing the model runner in this change.
**Rationale:** Users need to inspect stages independently during development and later chain them reliably. A committed, portable successful output is the handoff, not an inferred newest file. Failed runs must not silently reuse stale inputs. The user explicitly accepted whole-Draft autosave and chose a compact latest-run panel rather than full audit history.
**Implementation evidence:** H73:O83 was verified free before installation; its saved colours, names and formats were verified read-only afterward. Backup retained at references/backups/Draft Input spreadsheet for SST_20261028_124921_052826.xlsm.
**Future model information:** The user supplied a model directory ending in 02 Sensitivities/2025/05 ORSA/Models/Bonds Lite Batch/bonds-hm-cst and command python run_bonds_batch.py <input spreadsheet>. No model code was inspected or executed in this change; future configuration must not hard-code Joao's user folder.

### Decision 20: Movable And Resizable Named Audit Fields (2026-10-28)
**Question:** Allow the audit panel to be moved or resized within its worksheet without changing Python addresses?
**Decided by:** User
**Choice:** Give every audit field a stable workbook name and use named references for all runtime audit reads/writes. Permit whole-panel Cut/Paste within I. Control and layout resizing that retains valid field names. Keep the original coordinates only for first installation and coherent legacy migration, not runtime routing.
**Alternatives considered:** Continue fixed addresses; require code edits for each move; infer runtime fields from labels or silently fall back to the original location.
**Rationale:** Users need to adjust the workbook layout without breaking stage recording or the Stage 2 handoff. Names track Excel moves; explicit geometry/ownership checks reject ambiguous or broken references instead of writing to unintended cells. Existing audit values and successful output links must survive the upgrade.
**Implementation evidence:** Added five missing names to the existing four-name panel. Backup retained at references/backups/Draft Input spreadsheet for SST_20261028_140917_887382.xlsm; the user confirmed the saved-upgrade read-only check passed. The actual panel was not moved by the implementation.

### Decision 21: Bonds And MAP Pair Per Timestamped Stage 1 Run (2026-10-28)
**Question:** Extend Stage 1 to two asset datasets with one shared input template and separate model-output folders?
**Decided by:** User
**Choice:** Read Bonds from D39 and MAP from D40. Use the single shared template, scanned at D42 after row insertion, plus the same captured entity, valuation date and credit stresses. Create two input workbooks per timestamped run folder under Bonds Lite output, with <Entity> Bonds and <Entity> MAP subfolders. Set each generated workbook's output_folder_path name to its matching final model-output folder. Leave blank D40 blank in the installer; the user will enter it.
**Alternatives considered:** A second template-path cell; different MAP controls/stresses; a default MAP path written automatically; flat single-file outputs.
**Rationale:** Dataset and model-output destination are the intended differences between the pair. One run folder keeps inputs and later model results together without embedding a particular user's OneDrive path in the audit.
**Read-only evidence:** D40 has the MAP label and no value; D41 is blank; D42 contains Bonds Lite input template.xlsx. Both source tables are B8:ATT1066 with 1,215 columns. output_folder_path resolves to the input template's I. Control!D19. No production code or workbook changes were made while planning.

### Decision 22: Folder Audit Handoff And All-Or-Nothing Pair Generation (2026-10-28)
**Question:** How should the existing single-output panel and partial failures behave with two Stage 1 inputs?
**Decided by:** User
**Choice:** Keep the existing compact movable audit panel and link to the successful timestamped run folder. Stage 2 resolves the fixed Bonds/MAP input filenames inside that folder. Both datasets must succeed; a generation failure removes that attempt's incomplete folder/files and blocks Stage 2 while preserving prior runs and their last-successful link.
**Alternatives considered:** Expand the panel with separate workbook/handoff links; retain a partial generated pair for debugging.
**Rationale:** Stage 1 produces one coherent run, not two independently authorised partial handoffs. Retaining named audit-field ownership avoids disturbing the user's moved/resized panel.
**Recovery boundary:** A failure generating either dataset is different from final Draft audit persistence failing after a complete validated pair exists. The existing retain-complete-output-on-audit-save-failure rule remains in place; no handoff is committed in that case.

### Decision 23: Folder-Based Batch Stage 2 And Separate Interpreter (2026-10-28)
**Question:** How should Stage 2 run the Bonds Lite model for both committed inputs?
**Decided by:** User
**Choice:** Invoke python run_bonds_batch.py <Stage 1 run folder> once, using a separately configured model environment. Add a Run Bonds Lite button and a named/movable Stage 2 audit panel below Stage 1, subject to safe layout preflight.
**Alternatives considered:** One single-file model invocation per input; the controller's Python 3.13 interpreter; another panel location.
**Rationale:** The existing batch runner accepts the folder and processes the pair as supplied. Model dependencies belong in their own environment. The supplied py311 environment contains python.exe; the full executable path is C:/Users/Joao Branco/AppData/Local/anaconda3/envs/py311/python.exe, to be configured per user rather than embedded in VBA.

### Decision 24: Confirmed Reruns With Transactional Model-Output Recovery (2026-10-28)
**Question:** How should Stage 2 handle existing outputs and a partial model failure?
**Decided by:** User
**Choice:** Ask before replacing results. Back up previous Bonds/MAP results before a confirmed rerun, restore them if the attempt fails, and remove the backup only after verified success. Remove new outputs from failed attempts. Show only the run-folder link in the audit, not a summary-specific link.
**Alternatives considered:** Block every rerun; retain partial failed outputs for diagnosis; link the batch summary directly in the panel.
**Rationale:** Development needs deliberate reruns without destroying earlier successful results. The model's summary remains a result-validation artifact; console diagnostics explain failures before failed-attempt artifacts are removed.

### Decision 25: Stage 2 Timeout And Real Verification Permission (2026-10-28)
**Question:** What run limit and real-model verification are authorised?
**Decided by:** User
**Choice:** A 60-minute timeout for the complete two-input batch. After implementation and isolated tests pass, perform one real batch verification on the latest eligible audited Stage 1 run, respecting overwrite confirmation. Do not start it during planning or automatically retry failures.
**Alternatives considered:** Thirty-minute/no timeout; leave all real-model execution to the user.
**Rationale:** The numerical model can take several minutes per input. A bounded run with visible progress is needed, and the user explicitly authorised a real integration check after prerequisite gates.
**Implementation evidence:** The initial authorised run on `Bonds Lite output/20261028_154621` ended Failed and rolled back cleanly. The user later explicitly authorised a retry after a model fix. The current audited run was `Bonds Lite output/20261028_170055`; its four pre-existing summaries reported `Permission denied` reading the Bonds input, with a model input-copy file in the Bonds output folder. After the user approved replacement and the inputs were confirmed closed with no lock files, the retry ended Failed (17:23:48-17:23:54). Transaction rollback restored the four summaries and prior Bonds file unchanged. The terminal did not retain the retry's exception text. No further real run should occur until that exception is reviewed and explicitly authorised.

### Decision 26: Preserve XLL-Calculated Base Curves And Fail Closed (2026-10-28)
**Question:** How should Stage 1 handle model base-curve formulas that depend on the JRC Excel add-in?
**Decided by:** User
**Choice:** Keep worksheet recalculation and reject Stage 1 publication unless all `P. Base Curves[GBP SONIA]` cells are numeric afterward. Run Stage 1 from the installed button in the normal Draft Excel session with the JRC add-in loaded; do not silently skip recalculation or trust potentially stale cached curves.
**Alternatives considered:** Skip recalculating the base-curve sheet and reuse cached values; launch another hidden Excel instance and accept `#NAME?` caches.
**Rationale:** The custom `JRC_SpotAnnual` formula depends on valuation date. Reusing an old numeric cache when the date changes can silently produce the wrong model input; explicit failure is safer than publishing nonnumeric or stale curves.
**Implementation evidence:** The source template contained 1,201 numeric GBP SONIA cached values. A Stage 1 output produced through a newly created hidden Excel instance contained `#NAME?` for all 1,201 after calculation. The external model then raised `TypeError` at `1 + nparray_swaprate`. The regression guard and full non-Excel suite pass. Stage 1 must be rerun through the user's add-in-enabled Excel session before Stage 2.
**Stage 2 safeguard:** `read_model_inputs` independently checks the saved GBP SONIA cached column and refuses a handoff containing Excel errors, blanks or nonnumeric values, before result backup or model launch.

### Decision 27: Freeze XLL Curve Caches Only In Stage 2 Copies (2026-10-28)
**Question:** How should Stage 2 run when the hidden model Excel instance recalculates JRC XLL formulas to `None`?
**Decided by:** User
**Choice:** Stage 2 may create byte-verified local copies of the Bonds/MAP inputs and convert the three validated XLL-backed base-curve columns to their cached numeric values in those temporary copies only. Preserve original Stage 1 workbooks, all other assumptions/formulas, and embedded output paths; write model results to the original JRL output folders.
**Alternatives considered:** Change the external model; globally configure hidden Excel/add-ins; skip Stage 1 recalculation and trust stale caches; run directly against OneDrive inputs.
**Rationale:** Excel opens the OneDrive input in a hidden model process but recalculates `_xll.JRC_SpotAnnual` to `None`; direct read-only inspection showed the saved Stage 1 cache is numeric. The temporary copy isolates the required compatibility adjustment without modifying the audited inputs or changing model outputs.
**Implementation evidence:** Controlled tests verify byte-identical input staging, local runner path, summary handback, curve-cache freezing and source immutability. Full non-Excel suite: 600 passed, 8 skipped, 71 Excel tests deselected. A live add-in-enabled Excel run is still required to validate the freeze helper before model execution.
**Successful verification:** Stage 1 and Stage 2 succeeded on `Bonds Lite output/20261028_181415` after running from the add-in-enabled Draft Excel session. The saved Stage 2 audit matches that run; the 18:53:29-18:56:12 attempt validated summary `bonds_batch_summary_20261028_185611.csv`, both dataset result sets and unchanged original input hashes.

### Decision 28: Local Sidecar Runtime With Excel Context Handoff (2026-10-30)
**Question:** Which deployment model best supports fast startup, a portable Draft and per-user SharePoint sync paths without requiring user Python setup?
**Decided by:** User
**Choice:** Plan for a complete local `Bonds Lite App` folder beside Draft, installed/copied and verified by a maintainer. Use a REM-style workbook/Excel-instance handoff, but do not use REM's fixed UNC executable path. Target desktop Excel opening a locally synced SharePoint workbook; Excel for the web is out of scope. Prioritise speed and reliability over a hard package-size cap, while avoiding a full Anaconda runtime. No build is authorised by this decision; request separate approval before packaging.
**Alternatives considered:** Run the frozen application directly from UNC (one centrally updated copy, but the user reports slow launch and a Windows Run prompt); per-user LocalAppData cache (rejected by the user); a single embedded Python runtime (adds compatibility/parity work between controller Python 3.13 and model Python 3.11, with no established startup advantage).
**Rationale:** Local executable loading is the leading speed candidate, while REM's HWND/workbook handoff decouples workbook identity from executable location. A 332 MB isolated bundle passed preflight in staging but failed SciPy import from the OneDrive-synced sibling, so a no-build A/B location gate is required before implementation/build approval.
**Implementation evidence:** Planning only. REM was inspected read-only with macros disabled. No code, workbook or external model was modified; no build or financial batch was started.
**Superseded by:** Decision 29 — use the central release share only as a copy source and execute from LocalAppData.

### Decision 29: Local Runtime Cache From Central Release Source (2026-10-30)
**Question:** How can each user run a portable synced SharePoint Draft quickly while avoiding both UNC execution and SciPy's failure from a OneDrive-synced app folder?
**Decided by:** User
**Choice:** Read the central release pointer once on the first button click in each Excel session. Copy a changed version into a versioned per-user `%LOCALAPPDATA%` cache and launch controller/model only from that cache. Draft remains in the user's locally synced SharePoint/OneDrive folder; resolve its local path dynamically and attach to the exact Excel instance. Subsequent clicks in the same session use the cached local bundle without checking the share. If the share/new release is unavailable or invalid, use the previous complete cache with a visible StatusBar warning; stop only when no complete cache exists. No separate refresh macro/script is required.
**Alternatives considered:** Run frozen applications from the central UNC share (slow and prompts); execute the sidecar directly from the OneDrive sync tree (the compact package failed SciPy loading there); require users to relocate the Draft outside OneDrive (rejected portability); per-user Python installation/configuration (rejected).
**Rationale:** LocalAppData execution passed the compact bundle's model preflight, while executing the byte-identical folder from OneDrive failed. A local cache preserves a user-independent Draft path while avoiding repeated network execution. One central-pointer check per Excel session discovers updates without making every button click network-dependent.
**Implementation evidence:** Cache/pointer logic, frozen model selection, local SharePoint path handoff, packaged Stage 2 command and model entrypoint are implemented. Deployment-only regressions passed (16 tests); the Stage 1/2 safety subset passed (133 tests); both disposable Excel cache/fallback assertions passed. The combined focused run had one intermittent rollback-location assertion that passed in isolation. Excel emitted the known `0x800706ba` teardown diagnostic after passing assertions. Release `20261030-cache-v1` is 332,229,903 bytes (984 files); controller/model self-tests and clean-environment model preflight passed. The central copy completed with zero Robocopy mismatches/failures, matching source/copy file totals and both manifest executable hashes; `current.txt` was verified on the new release. The packaged launcher was installed into Draft with a timestamped backup. No real financial batch was run. The full suite and target-user warning/startup behavior remain unverified.

### Decision 30: Match Normal Excel Calculation For JRC Curves (2026-10-30)
**Question:** Should Stage 1 match normal F9 calculation or force fresh JRC curve retrieval on every run, including unchanged valuation dates?
**Decided by:** User
**Choice:** Match normal F9. Preserve the caller's calculation mode while opening workbooks, use manual mode only for bulk input edits, and restore the caller's mode before one application-level calculation. Do not rewrite valuation-date or entity controls that already match. Retain external-link, numeric-curve and bond-formula validation. Changed controls must still be applied and their dependencies calculated; do not substitute a prior valuation date's curve values. Application-level calculation can recalculate other open workbooks in the calling Excel session, but source workbooks are not saved.
**Alternatives considered:** Keep the sheet-by-sheet calculation sequence (reproduced a stall at P. Base Curves); change only the calculation method (not sufficient in the full production settings wrapper); force fresh JRC retrieval every time (the user also reported native Ctrl+Alt+F9 remaining busy and selected normal F9 instead); bypass curve validation or reuse another valuation date's results (rejected for correctness).
**Rationale:** A byte-identical disposable template worked when opened normally. Preserving automatic mode during opening completed the production open/refresh path in 1.79 seconds. The corrected production settings/open/manual-edit/control/refresh sequence completed in 2.11 seconds, validated all 1,201 SONIA rows and restored Excel settings. Focused calculation, mode-boundary, ownership and control tests passed (61 tests). The wider non-Excel run had 637 passes and the same two previously observed failures: the frozen-main test-signature mismatch and rollback-location reporting assertion. The user approved one replacement build and publication/installation only after the relevant checks pass; rollout evidence is recorded in the implementation plan. No real financial batch is authorised by this decision.

### Decision 31: Name Stage 1 Path Controls (2026-11-02)
**Question:** Can the Bonds, MAP and shared-template controls move without changing fixed cell references in Stage 1?
**Decided by:** User
**Choice:** Replace runtime reads of D39, D40 and D42 with workbook names `SST_Stage1_BondsDataPath`, `SST_Stage1_MAPDataPath` and `SST_Stage1_InputTemplatePath`. On installation, create all three names only when the legacy adjacent labels match; preserve existing complete names and reject partial/invalid sets. Do not require the user to name cells manually.
**Alternatives considered:** Keep fixed cell references (moving controls would break Stage 1); ask the user to create names manually (more error-prone and inconsistent across colleagues).
**Rationale:** Excel carries defined names with cut/moved cells, allowing the layout to change without embedding workbook coordinates in runtime logic. Label-checked migration avoids guessing on structurally changed Drafts. Runtime path resolution still validates local ownership and scalar-cell shape.
**Implementation evidence:** Stage 1 now resolves all three paths exclusively through the named controls; D39/D40/D42 remain only as label-checked one-time installer migration anchors. Migration preserves a complete existing moved set and rejects partial names, mismatched labels and ranges outside the local I. Control single-cell contract. Focused named-control tests passed (22); disposable installer Excel tests passed (2) and verified entered/default values and name placement. The non-Excel suite passed 664 tests, skipped 8 and retained two pre-existing failures (frozen-main test API mismatch and rollback-location capture). Release `20261102-cache-v4` passed model/controller self-tests, embedded-function parity checks and central/local copy/hash checks; both central and local pointers now select v4. The backed-up Draft contains all three names mapped to their original cells, with path values preserved. The packaged runtime can now resolve moved path controls. No Stage 1/2 financial run was performed for this change.
## Agent-made decisions pending user approval

### Decision 32: Stage 3 In A Separate Module, Delivered In Slices (2026-11-02)
**Question:** Where should Stage 3 logic live, and how much of it should be planned and built first?
**Decided by:** Agent (user required a separate file and asked to stop after the funded-leg transfer)
**Choice:** Put all MACalc/MACE logic in a new `src/sst_stage_three.py`. `bonds_lite_input.py` only gains `--stage 3` dispatch, through a lazy import, and public aliases for shared helpers. Slice A covers template block validation, funded-asset detection and the MAP/MACE transfer into the funded block. Later transfers are planned as further slices.
**Alternatives considered:** Extend `bonds_lite_input.py` (rejected by the user and adds to a 2,600-line module); plan all of Stage 3 at once (rejected because the remaining requirements are not yet supplied and the user wants Slice A verified first).
**Rationale:** Isolates new data-transfer risk, keeps Stage 1/2 regression surface unchanged and lets each slice be verified on real data before the next.

### Decision 33: Stage 3 Readiness Requires Matching Stage 2 Run Identity (2026-11-02)
**Question:** What makes Stage 2 output valid for Stage 3?
**Decided by:** Agent
**Choice:** Require `SST_Stage2_Status == Succeeded` and `SST_Stage2_InputRun` to equal the current committed Stage 1 handoff. Read `Bonds Lite input_MAP.xlsx` from that run folder using cached values.
**Alternatives considered:** Status alone (a new Stage 1 run would be paired with old Stage 2 results); the model's `JRL MAP/` input copy (it is a model artifact, not the Stage 1 handoff contract).
**Rationale:** Follows the readiness rule already set in plan-stage-two Step 8.

### Decision 34: Exact Header Contract And Exact-Equality Read-Back (2026-11-02)
**Question:** How should Stage 3 guard against data-transfer errors?
**Decided by:** Agent
**Choice:**
- Verify every source and destination header at its expected column letter and stop on any drift.
- Reject blank, error and non-numeric values in numeric fields, and missing or duplicate ISINs.
- Write values only.
- Reopen the saved output and require exact equality for every copied cell, plus exact `-1` negation for PV, before recording success.
**Alternatives considered:** Remap columns by header (silently hides template changes); tolerance-based comparison (not appropriate for pure copies); trust the write without read-back (cannot detect partial or misaligned writes).
**Rationale:** The user flagged data transfer as the main risk. A fail-closed design with verification is consistent with Decisions 18 and 26.

### Decision 35: Provisional Defaults For Unconfirmed Slice A Details (2026-11-02)
**Question:** How should the plan proceed while some layout and output details are unconfirmed?
**Decided by:** Agent
**Choice:** Record provisional defaults P1-P15 in `references/plan-stage-three.md` and mark them PENDING. Five of them (cashflow vector, placement below ILS, columns A:C, output location/format, zero-funded case) block their implementation step until the user confirms.
**Alternatives considered:** Withhold the plan until every question is answered (delays review of the overall design); implement on assumptions (rejected by the 2 November handoff).
**Rationale:** Lets the user review and amend the whole design in one pass while preventing unconfirmed choices from reaching code.
### Decision 36: Stage 3 Slice A Answers (2026-11-02)
**Question:** Which cashflows go into the funded block, how are funded assets identified, and how should placement, columns A:C, output location and the zero-funded case work?
**Decided by:** User
**Choice:**
- Funded-leg cashflows (MACalc N:ATQ) come from MACE `C. Assets - Funding` AJ:AUM, not MAP Q:ATT.
- A funded asset has an all-negative MACE funding-leg vector. With zero tails after maturity, this is implemented as at least one negative month and no positive month.
- Rows below the funded block are cleared for later slices.
- A:C on funded rows are left empty; they will be filled later from another MACE tab.
- Interim outputs for all stages are saved in the timestamped Stage 1 run folder.
- Zero funded assets produces a warning and the run continues; it is not an error.
**Alternatives considered:** Full MAP vector (the user's original specification, withdrawn because it includes the asset leg's positive months and differs from MACE by up to 19.8m); copied allocation formulas in A:B (rejected by the user); stop on zero funded (rejected because other entities may legitimately have none).
**Rationale:** MACE holds the funding leg directly. Read-only checks on 2 November found exactly 5 negative-only rows of 1,009, with the rest all zero, matching the 5 MAP assets with negative months.

### Decision 37: Fail-Closed Funding-Vector And Reconciliation Guards (2026-11-02)
**Question:** What should happen when MACE funding vectors are unusual, or MACE and MAP disagree on which assets are funded?
**Decided by:** User
**Choice:**
- Stop on any MACE funding vector with positive months (mixed sign or positive-only).
- Stop if the MACE-funded set differs from the MAP set of assets with any negative month, listing the differences.
- Both empty is the zero-funded warning case, not an error.
**Alternatives considered:** Ignore anomalous rows (could silently drop or misclassify a funding leg); trust MACE alone (would not detect MACE and Stage 1 data from different dates or portfolios).
**Rationale:** Both checks pass on current data and protect the main data-transfer risk the user raised. The user approved both checks on 2 November 2026.
**Superseded in part by:** Decision 82. The funded-set reconciliation now uses the Stage 1 Bonds input, not MAP.

### Decision 38: Preserve Excel State And External Links In MACalc Output (2026-11-02)
**Question:** How should the Stage 3 writer handle Excel application startup and workbook links when saving a staged MACalc copy?
**Decided by:** Agent (implementation detail within the user's approved Stage 3 implementation)
**Choice:**
- Start the owned hidden Excel application with a workbook open before setting application properties; an empty instance returned invalid values and rejected the `CalculateBeforeSave` setter.
- After Excel saves the staged copy, restore and verify the template's external-link targets before read-back validation. Excel rewrites SharePoint URL targets to UNC paths even when protected formula bodies and cached values are unchanged.
- Keep exact protected-formula checks; do not weaken validation to ignore link-path changes.
**Alternatives considered:** Set properties on an empty Excel instance (reproduced COM failure); accept Excel's rewritten link targets (changes workbook state); weaken protected-row comparisons (masks unintended edits).
**Rationale:** These safeguards were derived from failures reproduced in disposable Excel runs. The final disposable real-input output passed exact read-back, retained all 9 external-link targets and left the source template hash unchanged. No live Draft or Stage 3 orchestration was involved.

### Decision 39: Stage 3 Controls Mirror Stage 2 And Survive Launcher Reinstalls (2026-11-02)
**Question:** How should the requested Stage 3 button and audit panel be installed?
**Decided by:** User (requested the same style as previous stages); placement details by Agent
**Choice:**
- Use a blue 11-row panel `SST_Stage3_Panel`, with the same fields and formats as Stage 2, placed below the Stage 2 controls.
- Use a `btnPrepareMACalc` button placed by the Stage 2 rule: beside the panel when a 205pt button fits, otherwise below.
- Add `PrepareMACalc` to both the source and packaged launcher texts, so a Stage 1/2 reinstall cannot remove it.
- Install with `--install-stage-three`, after an automatic backup.
**Alternatives considered:**
- Append the macro only during Stage 3 installation. Rejected: Stage 1/2 reinstalls rewrite the module and would drop it.
- Always place the button below the panel. Rejected: inconsistent with Stage 2.
**Rationale:** The panel and button stay consistent with Stages 1 and 2, and the install is idempotent. Live installation was verified against the backup.

### Decision 40: Stage 3 Slice A Accepted (2026-11-03)
**Question:** Does the first live Stage 3 output match the user's expectation, so that Slice A can be closed and the remaining slices planned?
**Decided by:** User
**Choice:** Accept Slice A. The user compared the five funded rows in `Bonds Lite output/20261102_111001/MACalc input.xlsb` (Stage 2 summary `bonds_batch_summary_20261103_093358.csv`) and confirmed they look correct. The remaining Stage 3 transfers will be planned as new slices in a new session.
**Alternatives considered:** Hold Slice A open for a golden-fixture acceptance test against a hand-built MACalc (not requested; can still be added later).
**Rationale:** The orchestrator's exact read-back passed and the user's own check agrees, so later slices can build on the Slice A output layout.

### Decision 41: Stage 3 Slice B Requirements And Single-Run Delivery (2026-11-03)
**Question:** How is the rest of MACalc `I. Assets` populated after the funded block, and how does Slice B relate to Slice A?
**Decided by:** User
**Choice:**
- Copy the Stage 2 MAP output whose name contains `py_bonds_Base_Scenario_nozeros` into the rows directly below the funded block. The mapping is B→D, A→E, D×10000→F, F→H, I→I, E→K and J:ATN→M:ATQ. The substring must be easy to change.
- Fill Type (G), Index Linked? (J) and FSAO (L) from MACE `I. Bonds`, and A:C from the MACE `Control` K:O table (Comp. A %→A, Comp. B %→B, Partial CF Flag→C; Comp. C % ignored), both matched on MACalc Name (E).
- Stop and list the ISINs on unmatched or duplicate keys, blanks or errors. Stop if the `Control` table is not contained in K:O.
- Slices A and B run as one process. Commit and the v5 release wait until Slice C is accepted.
**Alternatives considered:** A separate Slice B button or run (rejected by the user: the split exists only for development).
**Rationale:** These are the user's stated requirements. The details are in `plan-stage-three.md`, section "Slice B".

### Decision 42: Slice B Interpretations And Provisional Defaults (2026-11-03)
**Question:** How should details that the requirements leave open, or that conflict with the live files, be handled?
**Decided by:** Agent (pending user approval)
**Choice:**
- Match on `I. Bonds` column D (ISIN), not E (company name).
- In `Control`, use K as the ISIN key, because the live headers put ISIN in K and Comp. A % in L.
- Define containment within K:O as the header contract plus no data in P:Z and no second block below the gap.
- Provisional defaults Q-B1 to Q-B6:
  - Q-B1: fill A:C on the funded rows too.
  - Q-B2: look for the CSV only in the MAP subfolder, with the substring held as a code constant.
  - Q-B3: ignore MACE rows that are not in the CSV.
  - Q-B4: stop if the CSV and MAP sets differ or the valuation dates differ; warn only on a PV difference.
  - Q-B5: copy the Partial CF Flag exactly.
  - Q-B6: keep CSV order and clear the rows below.
**Alternatives considered:**
- `I. Bonds` E as the key: company names are not unique and do not match the ISINs.
- `Control` L as the key: L holds Comp. A %, not the ISIN.
- A Draft named cell for the substring: avoids a code release, but adds Draft surface and audit risk.
- Stopping on any MACE extra: today the only extras are Inf_swap rows.
**Rationale:** Read-only checks on handoff `20261102_111001` show D and K are the unique ISIN columns, covering all 988 CSV ISINs. The defaults mirror the fail-closed guards approved for Slice A (Decision 37).

### Decision 43: Slice B Open Points Confirmed (2026-11-03)
**Question:** Which provisional Slice B defaults (Q-B1 to Q-B6) does the user accept?
**Decided by:** User
**Choice:**
- **Q-B1:** A:C are filled on every post-ILS row, funded rows included, from the MACE `Control` table. This supersedes P2.
- **Q-B2:** Look for the CSV only in the entity's MAP output folder from Stage 2 (the validated `output_folder_path`). The substring stays a code constant.
- **Q-B3:** MACE rows that are not in the CSV are allowed. A CSV asset that is not in MACE stops the run.
- **Q-B4:**
  - Warn, without stopping, if the CSV assets differ from the Stage 1 MAP assets.
  - Stop if the valuation dates of MACE and the Draft differ.
  - Do not compare CSV PV with MACE PV.
- **Q-B5:** Copy the Partial CF Flags exactly.
- **Q-B6:** Write in CSV order and verify that nothing remains below the last asset row. Slice C does not touch `I. Assets`.
- The user's reply did not address the Decision 42 interpretations (`I. Bonds` key D, `Control` key K, K:O containment rule). They stay **pending explicit confirmation** and are not accepted by silence.
**Alternatives considered:** The agent's defaults in Decision 42. The user changed Q-B4 from stop to warn for the CSV/MAP set difference and dropped the PV cross-check.
**Rationale:** These are the user's answers. They are recorded in `plan-stage-three.md`, under "Slice B, Confirmed Answers".

### Decision 44: Slice B Matching Keys And K:O Containment Confirmed (2026-11-03)
**Question:** Are the Decision 42 interpretations (I1-I3) correct?
**Decided by:** User
**Choice:**
- I1: match MACalc Name (E) against MACE `I. Bonds` column D (ISIN).
- I2: the MACE `Control` key is column K (ISIN), mapping L → A, M → B and O → C.
- I3: "contained in K:O" means the exact row-9 headers, no data in P:Z beside the table, and no second K block within 50 rows below the first blank. *(Amended by Decision 48: the side check now covers column P only.)*
**Alternatives considered:** `I. Bonds` E and `Control` L as keys (they hold company names and Comp. A % respectively).
**Rationale:** The user confirmed all three on 3 November 2026. The user also asked for a check that MACalc copes with `tbl_Assets` growing beyond the template's 1,008 rows; its result is recorded in the plan's Slice B risks.

### Decision 45: Rating-Suffixed ISIN In Column D; Slice C Aligns MACalc ISIN Lists (2026-11-03)
**Question:** The template held the plain ISIN in D, and MACalc joins D to `tbl_ForcedCompAperc`, `tbl_MinCompAperc` and `lst_pre16assets`. Should Slice B write `Unique_ISIN` or the plain ISIN to D, and what goes in E?
**Decided by:** User
**Choice:** Option (b): `Unique_ISIN` → D and the CSV `ISIN` → E, as originally specified. Funded rows keep the plain ISIN in D, so they are deliberately distinct from their rating-suffixed counterparts. E is also usable as a matching key. Aligning the three MACalc ISIN lists is the scope of Slice C.
**Alternatives considered:** (a) plain ISIN in D, as in the template. Rejected: funded and non-funded rows would share a D value, and the user intends rating-suffixed identifiers.
**Rationale:** This is the user's intended identifier scheme. Until Slice C is done, the interim output must not be used for optimisation runs.

### Decision 46: Guard MACalc's `tbl_Assets` Assumptions (2026-11-03)
**Question:** Will MACalc cope with `tbl_Assets` growing from 1,008 to 1,013 rows, and what must hold for it to read the assets correctly?
**Decided by:** Agent (from the user-requested check of `MACalc.py`)
**Choice:**
- Row growth is safe, because MACalc sizes everything from `len(df_assets)`.
- Before saving, stop if any column A cell in rows 2 to the last asset row is blank: `tbl_Assets` is sized by COUNTA over A.
- Stop if any G value is outside `Mortgage`, `Bond`, `Infrastructure`, `PP`, `Swap` or `ILS`: MACalc asserts that its asset groups sum to the total.
**Alternatives considered:** Rely on MACalc's own assertion. Rejected: it fails only at MACalc run time and does not catch a blank A, which silently truncates the table.
**Rationale:** Both conditions silently or late-failingly break MACalc, and both are cheap to check in Stage 3.

### Decision 47: Slice B Implementation Deviations (2026-11-03)
**Question:** Where did the Slice B implementation have to depart from the plan's wording?
**Decided by:** Agent; approved by the user on 3 November 2026. On point 1, the user asked whether pandas should be added to the packaged controller for speed. It was kept out: the CSV read takes 0.23 s of the 112 s live run, which is mostly Excel work (opening MACE over the network alone took about 27 s). Importing pandas takes about 0.7 s, and the build excludes pandas and numpy from the controller behind every stage button (`tools/build_packaged_stages.py`).
**Choice:**
- Read the base-scenario CSV with the standard-library `csv` module.
- Test the MACE readers with fake sheets and pure parsers.
- Return dataclasses that carry duplicate keys, and a `SliceBReconciliation` result object.
- Fix the Slice A rollback so a failure never deletes the previous output.
**Alternatives considered:** pandas (rejected: excluded from the packaged controller, so the frozen build would fail); generated `.xlsx` fixtures (rejected: slower and needs Excel or openpyxl formatting fidelity for no extra coverage of the parsing logic); leaving the rollback as is (rejected: the new Slice B guard test requires the previous output to survive).
**Rationale:** Keeps Stage 3 runnable in the packaged build, keeps the tests fast, and makes the "keep previous output" guarantee true.

### Decision 48: K:O Containment Checks Column P Only (2026-11-03)
**Question:** The live MACE `Control` sheet has other panels in Q, R, T and U (rows 9-31), so the confirmed I3 rule ("no data in P:Z") stops every live run. What should the side check cover?
**Decided by:** User
**Choice:** Option (b): only column P, from row 9 to the last table row, must be empty (`CONTROL_SIDE_COLUMNS = ("P", "P")`). The exact row-9 header check and the "no second K block within 50 rows below the gap" check are unchanged. This amends I3 in Decisions 42 and 44.
**Alternatives considered:** (a) keep P:Z and move the MACE panels: rejected, because the panels are a normal part of the MACE workbook. (c) a different range: not needed, because P is the only column between the table and the other panels.
**Rationale:** P still catches a table that has grown sideways (for example a new "Comp. D %" column), while the user confirmed that the columns from Q onwards are expected to hold data.

### Decision 49: Slice C Mapping Rule And Defaults (2026-11-03)
**Question:** How should each entry of `tbl_ForcedCompAperc`, `tbl_MinCompAperc` and `lst_pre16assets` map to the new `I. Assets` D values, and what happens to unmatched entries, multiple matches and duplicates?
**Decided by:** Agent (pending user approval)
**Choice:**
- C1: for each entry `k`, the matching row set `M(k)` is {rows 2..last with D = `k`} ∪ {asset-block rows with E = `k`}. Emit the D of every row in `M(k)`, in `I. Assets` row order, at `k`'s position, with `k`'s value unchanged. Matching is exact. E is not used for Mortgage/ILS rows.
- An asset row is reached through E only when its D is exactly `E_<its own rating>`. If no row is suffixed, the lists are written back unchanged. Any other D/E shape stops the run (Guard 7). **User-confirmed 3 Nov (Q-C3).**
- C2: |M| = 0 means the entry stays in place, unchanged, with a console count. **User-confirmed 3 Nov (Q-C4).**
- C3: |M| ≥ 2 (a funded plain ISIN that is also an asset E, and/or several asset rows sharing one E, such as the same ISIN at two ratings) fans out to one entry per matching row, in row order, contiguously and with the same value. The fan-out is reported on the console and does not stop the run. Added = |M| − 1. If a fan-out collides with another entry's key, the run stops for a table, and the duplicate is silently removed for a membership list. There are zero cases today.
- C4/C5: original order, values copied exactly, and the block is rewritten from row 7 with the old tail cleared.
- C6: no fuzzy matching.
- Stops: title, name-formula or containment drift; blank or error keys; non-numeric or out-of-[0,1] values; an asset row whose D is neither E nor `E_<rating>` (Guard 7); duplicate source or remapped keys **in the two tables** (the default, pending the user's choice between a, b and c in Q-C5).
- Membership lists (`lst_pre16assets`, `lst_REITs`): duplicates (in the source or from a fan-out) are removed silently, keeping the first occurrence. **User-decided 3 Nov (Q-C5).**
- Slice C runs inside the existing single write, save, verify and audit.
**Alternatives considered:**
- Match on the ISIN prefix of D (fuzzy; could hit the wrong row, and it is not how MACalc joins).
- Drop stale entries (changes template content for no model effect).
- Emit only the `Unique_ISIN` for funded overlaps (silently drops the restriction from the funded leg).
- Stop the run on any multi-match. This is safe but would block a valid quarter in which one ISIN is held at two ratings, a case the rating suffix exists to support.
- Pick one row for a multi-match (first, or highest rating). This is arbitrary, and it silently leaves the other legs unrestricted.
- Write formulas linked to `I. Assets` (MACalc reads values; adds recalculation dependence).
**Rationale:** Discovery on handoff `20261102_111001` shows that this rule restores exactly the template's match counts (Forced 24, Min 106, Pre16 31, REITs 16). MACalc's left-merge makes duplicate keys a silent misalignment risk, so they stop the run.
**Superseded by:** Decision 52 — same-value duplicates in the Comp A tables are removed; only conflicting values stop the run.

### Decision 50: Include `lst_REITs` In Slice C (2026-11-03)
**Question:** The brief names three lists, but `calcSCR.py` (`v_3_3`, lines 32-40) also matches `lst_REITs` against `I. Assets` D, to reclassify REITs as Financial for the SCR. Should Slice C align it?
**Decided by:** User, 3 November 2026 (Q-C1 include; Q-C2 leave `lst_noPRoC`, verify unchanged). Subject to re-check in step C0 (Decision 51)
**Choice:** Include `lst_REITs` (`P. Parameters` M7:M22, 16 entries, all remapped via E) under the same rule as Decision 49. Leave `lst_noPRoC` unchanged: it holds the 8 Inf_swap D values, which Stage 3 never changes, and `v_3_3` does not read it.
**Alternatives considered:** Restrict Slice C to the three named lists (16 REITs would be treated as non-financial in every MACalc run, which misstates the SCR).
**Rationale:** The impact goes beyond optimisation runs, and the change uses the same mechanism at negligible cost.

### Decision 51: Slice C Reference Model Is The New Python MACalc (2026-11-03)
**Question:** Q-C6. Is `MACalc.py` in `v_3_3` the right reference for how the MACalc output's ISIN lists are used?
**Decided by:** User
**Choice:** No. The output will feed the Python MACalc in `C:\Users\Joao Branco\OneDrive - Just Group\Financial Resilience Team - General\00 Ad Hoc\MACalc_Batch`, Stage 5 is the run of that model on the MACalc input once Stage 3 has validated it. That model is not read, changed or run in Stage 3. Slice C is specified against the template's contract: the lists keyed on `I. Assets` D, with an exact match. Before Stage 5's first run, its planning also confirms the join, the list set and the duplicate-key behaviour, and any gap reopens Slice C.
**Alternatives considered:** (a) Read the new model now and block Slice C on it. Rejected by the user: no work on that model before Stage 5. (b) Treat `v_3_3` as authoritative. Rejected, because it is not the consuming model.
**Rationale:** This keeps Stage 3 independent of Stage 5, and the template contract is the interface that both must honour.
**Status (4 November 2026):** "Stage 5" here is the work the user now calls **Stage 4** (Decision 67). The deferred checks were done read-only on 4 November against `MACalc_Batch` and the live Stage 3 output, and every one passed; Slice C is not reopened. The evidence is in `references/plan-stage-four.md`, "Verified Discovery".

### Decision 52: Slice C Duplicate Handling And Plan Approval (2026-11-03)
**Question:** How should Slice C handle duplicate keys in the Comp A tables, and is the Slice C plan approved?
**Decided by:** User
**Choice:** Approve the Slice C plan. In `tbl_ForcedCompAperc` and `tbl_MinCompAperc`, silently remove duplicate keys with identical values, keeping the first occurrence; stop and report the key and all values if a duplicate key has conflicting values. In `lst_pre16assets` and `lst_REITs`, silently remove duplicates, keeping the first occurrence. This applies to duplicates already in the source and collisions created by fan-out.
**Alternatives considered:**
- Stop on every table duplicate — rejected because duplicates with identical values do not create an allocation ambiguity.
- Silently keep the first table entry regardless of value — rejected because conflicting percentages would make the intended allocation ambiguous.
**Rationale:** Preserve an unambiguous single constraint for each key, while failing explicitly when duplicate entries disagree. The user approved the completed Slice C plan on 3 November 2026.
### Decision 53: Guard 7 Compares D Exactly With The Row's CSV `Final_Rating` (2026-11-03, revised)
**Question:** Plan C1 and Decision 49 define a rating-suffixed asset row as one whose D is exactly `E + "_" + <that row's own CSV rating>`, and name the field as `MapBaseRow.rating`. That field is the CSV `Final_Notched_Rating` (column I, e.g. `BBB+`), which is not the suffix Bonds Lite writes: on handoff `20261102_111001` it would stop 558 of 988 asset rows. Which of the row's CSV ratings does the approved exact rule refer to?
**Decided by:** Agent, approved by the user on 3 November 2026 ("the suffix comes from Final_Rating"). This is a clarification of the approved rule, not a new rule.
**Choice:** Guard 7 and the E-match compare D **exactly** with `E + "_" + Final_Rating`, where `Final_Rating` is the CSV column H value of the same row. Stage 3 now reads and header-checks CSV column H (`MapBaseRow.final_rating`) and passes it to the asset-block keys before the write and on read-back. There is no string manipulation: any other D stops the run, including the notched form (`X_BBB+`), another grade, and stray whitespace. `I. Assets` column I still receives `Final_Notched_Rating`, unchanged.
**Evidence:** In the 111001 base-scenario CSV (988 rows), `Unique_ISIN == ISIN + "_" + Final_Rating` on all 988 rows; with `Final_Notched_Rating` it holds on only 430. The plan's own discovery ("`D == E + "_" + rating` on every row, with ratings GOV, AAA, AA, A, BBB, BB and B") describes `Final_Rating`, not the notched field.
**Alternatives considered:**
- Literal `Final_Notched_Rating` (the field named in the plan). Rejected: it contradicts the plan's verified discovery and would stop every live run.
- Strip `+`/`-` from the notched rating (the first version of this decision). Withdrawn: it derives the suffix heuristically rather than comparing exactly with a CSV field, which is what the approved rule asks for.
- Accept any grade from a fixed set. Rejected: it no longer ties the suffix to the row's own rating.
**Rationale:** Exact equality with the row's own CSV rating field is the approved rule; `Final_Rating` is the only CSV field for which the plan's discovery holds. Reverse it by changing the field read into `MapBaseRow.final_rating`.

### Decision 54: Unique MACalc Input Filename (Slice D) (2026-11-03, approved with changes)
**Question:** `MACalc_Batch` (Stage 4) needs each MACalc input workbook to have a name that is unique across synced SharePoint folders. What unique suffix should Stage 3 add to `MACalc input.xlsb`, and how should reruns, the staged file and legacy outputs behave?
**Decided by:** User, 3 November 2026 (answers to plan Q-D1 to Q-D4), on the agent's proposal. The staging location (below) is the agent's choice under Q-D3; the user then confirmed that `MACalc_Batch` scans only the folder holding the MACalc input it is given (the run folder), so the sibling staging folder is not seen.
**Choice:**
- Name the output `MACalc input_<Stage 1 run ID>_<token>.xlsb`, where the run ID is the run folder name (`YYYYMMDD_HHMMSS`, with `_NNN` on a same-second collision) and `<token>` is `secrets.token_hex(8)` (16 lowercase hex characters), generated fresh on every Stage 3 run and regenerated if the name already exists in the run folder. Q-D1: `MACalc_Batch` has no naming conventions, so the 16-character token is kept.
- Stop before any write if the folder name does not match the run-ID pattern, or if the output path or the staged path exceeds 218 characters, the longest workbook path Excel can open or save. Under Q-D1 this was first set to 259 (Windows `MAX_PATH`). On 3 November the user lowered it to 218, because a staged path between 219 and 259 characters would pass the check and then fail inside Excel. The 218 limit was added after `20261103-cache-v7` was built, so it takes effect in the next release.
- Q-D2: earlier outputs are **kept**. A rerun writes a new tokenised file alongside them, with no overwrite prompt and no deletion. The rename-to-backup and restore logic is removed, because a new name never overwrites an old one.
- Q-D3: `MACalc_Batch` does not ignore dot-prefixed files, so nothing is staged in the run folder. The workbook is staged in `<output root>\Stage 3 staging\<token>\`, a sibling of the run folders, and the token folder is removed after every run (the `Stage 3 staging` folder too, once no other run is using it). Staging stays inside the synced output root because saving the template copy in local `%TEMP%` makes Excel rewrite the template's SharePoint link paths, which the exact link check then rejects. OneDrive marks these folders read-only, so removal clears that flag and retries; a folder that still cannot be removed is a warning, not a failure.
- Q-D4: legacy `MACalc input.xlsb` files are left where they are, with no prompt, log line or other handling. Their run folders will be deleted.
- Stage 4 must find its input through `SST_Stage3_LastSuccessfulOutput`.
**Alternatives considered:**
- Stage 3 run timestamp: not unique across Drafts started in the same second.
- Run ID only (the first proposal): unique only within one output root. Two Drafts starting Stage 1 in the same second, or a copied run folder, give the same name, so it does not meet the requirement.
- Run ID plus the Windows user name: still clashes for one user running two Drafts in the same second, and puts the user name in the file name.
- Run ID plus a token reused across reruns: one stable name per handoff, but a copied run folder carries the token and duplicates it on rerun.
- Full UUID4 (32 hex characters): 122 random bits, 16 more path characters, no benefit at these volumes.
- Ask, then delete earlier outputs after a successful rerun (the proposal): rejected by the user (Q-D2).
- Staging in the run folder under a dot-prefixed name (as in v6): rejected, because `MACalc_Batch` would see it (Q-D3).
- Staging in local `%TEMP%`: rejected after testing; Excel rewrites the SharePoint link paths, the read-back link check stops the run, and restoring the links reorders them and raised a COM error.
**Rationale:** The requirement is uniqueness across synced folders, which a run ID cannot give because it is unique only within one output root. A fresh 64-bit random token per run makes a clash negligible (about 3 in a trillion among 10,000 outputs) without a central registry, which Stage 3 cannot consult. Keeping the run ID preserves traceability. Because new and old names never coincide, earlier outputs survive any failure by construction, which keeps the P14 guarantee with simpler code. `MACalc_Batch` was not inspected (Decision 51); its rules come from the user's answers.

## Agent-made decisions pending user approval

### Decision 55: Credit Stress Flags And Selection Checks (2026-11-04)
**Question:** How should Stage 1 read and enforce `SST_Credit_DG_standalone` and `SST_Credit_spread_standalone`?
**Decided by:** Agent (pending user approval); the identity and all-zero rules are the user's.
**Choice:** Both names must exist and hold Excel booleans. If both are FALSE, stop, because no stress run is selected. If both are TRUE, produce one combined run, as today. Downgrade TRUE: stop if any of `C9:J15`, `C20:J26`, `C31:J37` or `C42:J48` is exactly the identity (1 on the seven rating diagonals, 0 everywhere else, including D). Spread TRUE: stop if all eight `M:DI`/`DL:HH` grids are exactly zero. Comparisons are exact. A FALSE flag is not checked, and its grids are still copied unchanged.
**Alternatives considered:** a numerical tolerance (rejected: grids are typed inputs and exact matching is predictable); zeroing or identity-replacing the grids of a FALSE flag in the generated inputs (not requested and would change financial inputs; left open); treating non-boolean text such as "TRUE" as true (rejected: the input is ambiguous).
**Rationale:** Implements the requested checks without changing which stresses reach the model.
**Status (4 November 2026):** Partly superseded by Decision 58. The user approved stopping when both flags are FALSE, and approved exact comparisons. Stopping on an inconsistent selection is replaced by a warning and a prompt, and FALSE flags are now checked.

### Decision 56: SST Output Folder Layout (2026-11-04)
**Question:** Where does each run write, and what happens when the folder already exists?
**Decided by:** Agent (pending user approval); the layout is the user's.
**Choice:** Write to `<Draft folder>\SST Output_<suffix>\<stress folder>`, using the registry `{"Credit": (DG flag, spread flag)}`; later stresses are added to the registry. The suffix must be non-blank and filename-safe; integer-valued numbers are written without `.0`. The root may already exist. An existing stress folder stops Stage 1 before any write and asks for a different suffix. A failed attempt removes the stress folder, and also the root if this attempt created it and it is empty. The audit hands off `SST Output_<suffix>/Credit`, and legacy `Bonds Lite output\<timestamp>` handoffs are rejected.
**Alternatives considered:** replacing the existing `Credit` folder (rejected: it destroys earlier Stage 2/3 results without consent); adding a numeric collision suffix (rejected: it contradicts the requested exact folder name); accepting legacy handoffs as well (rejected: the request replaces that layout).
**Rationale:** Keeps the old guarantee that a run never overwrites earlier outputs.
**Status (4 November 2026):** Partly superseded by Decision 59. The user approved the suffix rules and the rejection of legacy handoffs; they will rerun all stages at the end. An existing stress folder is now overwritten instead of stopping Stage 1.

### Decision 57: Stage 3 Filename Under The New Layout (2026-11-04)
**Question:** Decision 54 named the MACalc input after the timestamped run ID. What replaces it?
**Decided by:** Agent (pending user approval).
**Choice:** `MACalc input_<suffix>_<16 hex>.xlsb`, with the suffix read from the parent `SST Output_<suffix>` folder. Staging moves to `SST Output_<suffix>\Stage 3 staging\<token>`, outside `Credit`. The fresh token, the 218-character limit and the kept earlier outputs are unchanged.
**Alternatives considered:** `MACalc input_Credit_<token>` (rejected: carries no run identity); token only (rejected: loses the traceability Decision 54 asked for).
**Rationale:** The suffix is now the run's identity, as the run ID was.
**Status (4 November 2026):** Not contested by the user, so treated as accepted. The user did not explicitly approve it.

### Decision 58: Flag/Stress Inconsistencies Warn And Ask Instead Of Stopping (2026-11-04)
**Question:** What happens when the credit flags and the `I. Credit Stresses` grids disagree?
**Decided by:** User, 4 November 2026. The prompt mechanism and the inconsistency wording are the agent's choices (pending approval).
**Choice:** Both flags are checked, whether TRUE or FALSE. Comparisons are exact, and a matrix counts as the identity when it has 1 on the seven rating diagonals and 0 everywhere else, including D. These are inconsistencies:
- downgrade TRUE: each transition matrix that is the identity;
- downgrade FALSE: each transition matrix that is not the identity;
- spread TRUE: all eight spread grids are zero;
- spread FALSE: each spread grid that holds any non-zero value.

If there are any, Stage 1 logs one warning that lists each inconsistency with its flag and block addresses. It then asks in the Stage 1 console: `Proceed with these inconsistencies? [y/N]`. `y`/`yes` continues with the stresses exactly as validated. Any other answer, a blank or no console input cancels the run. On cancel, Stage 1 creates and replaces no folder, leaves the Stage 1 audit exactly as it was, and says to fix the flags or the stresses and rerun. These still stop the run as before: both flags FALSE, a missing flag name, or a flag that is not an Excel boolean. The grids are always copied unchanged.
**Alternatives considered:**
- Stopping the run (Decision 55): replaced at the user's request.
- A Windows message box: rejected because every stage already confirms in the console (Stage 2's overwrite prompt, Decision 14).
- Neutralising the grids of a FALSE flag: rejected because the user wants consistency to be the user's responsibility.
**Rationale:** The user remains responsible for consistency, but is always told about any mismatch before anything is written.

### Decision 59: Reusing A Suffix Overwrites The Stress Folder (2026-11-04)
**Question:** What happens when `SST Output_<suffix>\<stress folder>` already exists?
**Decided by:** User, 4 November 2026. The safe-replace mechanism is the agent's choice (pending approval).
**Choice:** The existing stress folder is overwritten without a further prompt. To keep an earlier output, change `SST_output_folder_suffix` or archive the folder before running. Safe-replace:
- After validation and any inconsistency prompt, Stage 1 renames the existing folder to a backup beside it, `SST Output_<suffix>\~replacing Credit <YYYYMMDD_HHMMSS>`. It then builds a fresh `Credit`.
- On success, it deletes the backup. If that deletion fails, it warns and names the backup, and the run still succeeds.
- On failure or interrupt, it removes the incomplete `Credit` and renames the backup back.
- If the first rename fails (a file is open in Excel, or OneDrive has it locked), Stage 1 stops before writing and asks the user to close files from that folder.

Earlier Stage 2 results and Stage 3 MACalc inputs inside the replaced folder are deleted with it, so Stages 2 and 3 must be rerun.
**Alternatives considered:**
- Stop and require a new suffix (Decision 56): rejected by the user.
- Delete the old folder before generating: rejected because a failed run would then lose the earlier output too.
- Ask before overwriting: rejected because the user asked for a plain overwrite.
**Rationale:** Gives the requested overwrite while keeping the previous output if the replacement run fails.

## Stage 4 Decisions (MACalc Batch On The Stage 3 Output)

### Decision 60: Stage 4 Starts As A Source-Mode Launcher (2026-11-04)
**Question:** Q4-1. Should Stage 4 start as a terminal launcher, or be fully integrated now (`--stage 4`, a Draft button, audit cells and a release)?
**Decided by:** User, 4 November 2026.
**Choice:** Option (a). First build a source-mode Python launcher, run from the terminal, that runs MACalc Batch on the audited Stage 3 output. The Draft button, audit cells, VBA, packaging and release come later, under a separate agreement.
**Alternatives considered:** (b) Full integration now. Deferred: it changes the Draft and needs a release before it is known that the model runs on Stage 3 output.
**Rationale:** The smallest step that proves the handoff works.
**Superseded in part by:** Decision 76. The user asked on 4 November 2026 for a Stage 4 button, audit panel and `--stage 4` dispatch in the same fashion as Stages 1-3. Building and promoting a packaged release still needs separate approval.

### Decision 61: MACalc Batch Runs In Place In The Stress Folder (2026-11-04)
**Question:** Q4-2. MACalc Batch saves over its input workbook and scans its folder recursively for `MACalc*` workbooks. Should Stage 4 copy the audited file into its own folder, or run in place?
**Decided by:** User, 4 November 2026.
**Choice:** Option (b). Run MACalc Batch on `SST Output_<suffix>\<stress folder>` (today `Credit`), in place on the audited Stage 3 workbook. MACalc's results are written into that workbook, and the CSV export is written beside it. The audited file is therefore changed by Stage 4.
**Alternatives considered:** (a) Copy into a `Stage 4` subfolder. Rejected by the user in favour of keeping everything in one folder (Decision 65).
**Rationale:** Keeps one self-contained output folder for each stress, ready for a future single run of all stages.
**Consequence:** The scan finds every `MACalc*` workbook under the folder, so Decision 62 makes the launcher stop when it finds anything other than the audited file. Stage 3 keeps earlier outputs (Decision 54, Q-D2), so after a Stage 3 rerun the user must remove or archive the earlier `MACalc input_*` files (or rerun Stage 1-3 with the same suffix) before Stage 4 runs. The launcher never deletes them.

### Decision 62: Automatic Confirmation Only For The Single Audited Workbook (2026-11-04)
**Question:** MACalc Batch asks `Proceed with these workbooks? [y/N]` through `input()`. May the launcher answer it?
**Decided by:** User, 4 November 2026.
**Choice:** The launcher answers `y` if, and only if, MACalc Batch's own workbook list contains exactly one workbook and that workbook is the path in `SST_Stage3_LastSuccessfulOutput`. Otherwise it answers no, lists what was found and stops with an error; nothing is run.
**Alternatives considered:** Always answer `y` (rejected: could run stale or unrelated workbooks); always leave the prompt to the user (rejected: unnecessary when the input is proven).
**Rationale:** Automation is safe only when the model will run exactly the audited input.

### Decision 63: Stage 4 Run Options (2026-11-04)
**Question:** Q4-3. Which MACalc Batch options should Stage 4 use, and should `SST_Credit_DG_standalone` drive the downgrade-run option?
**Decided by:** User, 4 November 2026.
**Choice:** `is_downgrade_run` = the Draft's `SST_Credit_DG_standalone` (an Excel TRUE/FALSE, read like Stage 1 reads it). Everything else stays at the MACalc Batch defaults during development: no binary optimisation, no optimisation pass, CoD floor at zero. A future Draft flag will say whether MACalc should optimise; it does not exist yet.
**Alternatives considered:** Use all defaults, including no downgrade run (rejected: the user wants the flag to drive it).
**Rationale:** Ties the model option to the Draft's stress selection now, so the future optimisation flag only needs to switch the optimisation pass on.
**Finding (agent, read-only, 4 November):** In the current MACalc Batch code, `is_downgrade_run` is read only by the optimiser (`MACE.reoptimise`, which groups downgraded infrastructure/PP rows under one binary variable). With the optimisation pass off it has no effect on the base-run results. The launcher still passes it and logs that it is inactive while optimisation is off.

### Decision 64: Stage 4 Proof Of Correctness (2026-11-04)
**Question:** Q4-4. What shows that a Stage 4 run is correct?
**Decided by:** User, 4 November 2026.
**Choice:** At this stage, a run that completes without errors and produces the expected artefacts: the workbook saved in place with MACalc's results, and `<stem>_unoptimised_assets.csv` beside it. The user then checks whether the numbers are reasonable. There is no numerical reference or tolerance yet.
**Alternatives considered:** A manual MACalc run, a previous SST result, or named SCR/MA figures. Deferred by the user until the run itself is established.
**Rationale:** Establish a clean run before judging the figures.

### Decision 65: Stale Stage 3 Handoff Stops Stage 4 (2026-11-04)
**Question:** Q4-5. If the latest Stage 3 attempt failed but `SST_Stage3_LastSuccessfulOutput` still names an older successful output, what should Stage 4 do?
**Decided by:** User, 4 November 2026.
**Choice:** Option (a). Stop and ask the user to rerun Stage 3. Stage 4 runs only when `SST_Stage3_Status` is `Succeeded`, `SST_Stage3_InputRun` is the current Stage 1 run, Stage 2 has succeeded for that run, and the output file exists.
**Alternatives considered:** (b) Run on the older output with a warning. Rejected.
**Rationale:** Consistent with how Stage 3 treats Stage 2.

### Decision 66: Stage 4 Output Stays In The Stress Folder (2026-11-04)
**Question:** Q4-6. Where should MACalc results go, and may a Stage 4 rerun replace earlier Stage 4 results?
**Decided by:** User, 4 November 2026.
**Choice:** Everything stays in the same stress folder, as centralised as possible, because all stages will later run in one go. A Stage 4 rerun replaces the previous Stage 4 results: MACalc rewrites its output sheets in the same workbook and overwrites the CSV. A Stage 1 rerun with the same suffix clears them with the rest of the folder (Decision 59).
**Alternatives considered:** A separate `Stage 4` subfolder, or keeping each rerun's results. Rejected.
**Rationale:** One self-contained folder per stress.

### Decision 67: `_opt` Suffix Not Guarded; Stage Numbering (2026-11-04)
**Question:** Q4-7. MACalc Batch skips files whose names contain `_opt`. Should Stage 1 reject such a suffix, or Stage 4 stop? Also, older documents call this work "Stage 5".
**Decided by:** User for `_opt`, 4 November 2026. The numbering note is the agent's record of the user's instruction in the Stage 4 handoff.
**Choice:** No `_opt` guard for now; the user considers that suffix very unlikely. If it happens, MACalc Batch finds no workbook and Decision 62 stops the launcher with a message that lists nothing found. "Stage 5" in Decision 51, the Stage 3 plan and the setup guide means **Stage 4**, the run of MACalc Batch on the Stage 3 output.
**Alternatives considered:** Reject `_opt` suffixes in Stage 1, or add a dedicated Stage 4 check. Deferred.
**Rationale:** Not worth the extra rule today.

### Decision 68: Separate Stage 4 Controller And Model Entry (2026-11-04)
**Question:** P4-1. Should Stage 4 use two modules for the controller and the MACalc Batch child, or one module with two modes?
**Decided by:** User, 4 November 2026.
**Choice:** Use two modules: `src/sst_stage_four.py` under the controller's Python 3.13 and `src/macalc_batch_entry.py` under the model's Python 3.11.
**Alternatives considered:** One module with controller and child modes. Rejected because it would need to support two Python environments and combine distinct responsibilities.
**Rationale:** The controller and MACalc Batch require different dependency environments; separating their entry points keeps each process aligned with its Python runtime.

### Decision 69: Reuse The Existing Model Python Setting (2026-11-04)
**Question:** P4-2. How should the MACalc Python configuration work for another user, and should Stage 4 add a separate interpreter setting?
**Decided by:** User, 4 November 2026.
**Choice:** Reuse the per-user `BONDS_LITE_MODEL_PYTHON` setting and its existing environment/registry resolution for the MACalc Batch child. Keep `SST_MACALC_DIR` optional, defaulting to the model folder relative to the Draft.
**Alternatives considered:** Add a required `SST_MACALC_PYTHON` setting. Rejected to avoid duplicate per-user configuration while Stage 2 and MACalc Batch use the same Python environment.
**Rationale:** The existing setting is user-specific, so each direct-source user can point it to their own Python 3.11 interpreter; no machine-specific path is embedded in Stage 4. A separate setting can be added if the models later require different interpreters.

### Decision 70: Guard MACalc Batch Confirmation Against The Audited Input (2026-11-04)
**Question:** P4-3. How should Stage 4 safely answer MACalc Batch's workbook confirmation prompt?
**Decided by:** User, 4 November 2026.
**Choice:** In the child process, answer yes only when the exact list MACalc Batch is about to run contains exactly the audited Stage 3 workbook; otherwise stop. Pre-scan the folder for a clear early error.
**Alternatives considered:** Pipe `y` to standard input after a pre-scan. Rejected because that would not bind confirmation to the actual list passed to MACalc Batch.
**Rationale:** Prevents an unreviewed workbook from being run and preserves the user's single-input confirmation rule without editing MACalc Batch.

### Decision 71: Record Model Provenance Without Modifying It (2026-11-04)
**Question:** P4-4. How should Stage 4 identify the exact MACalc Batch code used while leaving its working copy untouched?
**Decided by:** User, 4 November 2026.
**Choice:** Run with bytecode writing disabled and log the model Git HEAD, modified/clean state, and SHA-256 hashes of the six specified model files.
**Alternatives considered:** Modify or clean the model repository, or record only its commit. Rejected because the model must not be changed and its current working tree has uncommitted changes.
**Rationale:** The recorded hashes identify the actual working-copy code while the no-bytecode settings avoid incidental files in the model folder.

### Decision 72: Preflight MACalc Batch Before Launch (2026-11-04)
**Question:** P4-5. What must be checked before Stage 4 modifies the workbook?
**Decided by:** User, 4 November 2026.
**Choice:** Stop if a matching Excel lock file exists beside the audited workbook; verify that the configured interpreter is Python 3.11 and can import `MACalcrunmanager` from the model folder.
**Alternatives considered:** Launch MACalc Batch without these checks. Rejected because an open workbook or unusable environment would make the in-place run fail or behave unpredictably.
**Rationale:** Catch the known workbook and environment blockers before invoking the model.

### Decision 73: Run Stage 4 From A Source-Mode CLI (2026-11-04)
**Question:** P4-6. How should the user invoke Stage 4 in this planning scope?
**Decided by:** User, 4 November 2026.
**Choice:** Provide `python -m sst_stage_four --draft "<local Draft path>"`. Require the Draft to be open in desktop Excel, resolve the exact open workbook, read its live audit and do not write to the Draft.
**Alternatives considered:** Add a Draft button or integrate `--stage 4` now. Deferred; the user selected a source-mode launcher first (Decision 60).
**Rationale:** Reuses the live Draft audit while keeping this stage self-contained and out of the release and VBA surfaces.
**Superseded by:** Decision 76. Stage 4 now starts from a Draft button through `bonds_lite_input --stage 4`, like Stages 1-3, and writes its own `SST_Stage4_*` audit cells. The separate `python -m sst_stage_four` command is dropped.

### Decision 74: Replace One Automatic Pre-Run Backup (2026-11-04)
**Question:** P4-7. Because MACalc runs in place, should the user make a manual backup, should the launcher manage one, or should there be no backup?
**Decided by:** User, 4 November 2026.
**Choice:** Before every Stage 4 run, automatically copy the audited workbook to `SST Output_<suffix>\Stage 4 backup\<same workbook name>`. Replace that same backup file on each run; do not accumulate versioned backups. The backup remains outside the `Credit` scan folder.
**Alternatives considered:** Ask the user to make a one-time manual copy (rejected in favour of consistent automatic protection); make no backup and rely on rerunning Stage 3 or OneDrive history (rejected).
**Rationale:** Preserves the pre-run workbook for recovery or comparison while keeping only one backup per input and ensuring MACalc Batch cannot discover it as another input.

### Decision 75: Stream MACalc Output Without A Timeout (2026-11-04)
**Question:** P4-8. Should the child process have a timeout, and how should its output be presented?
**Decided by:** User, 4 November 2026.
**Choice:** Stream child output to the console and impose no timeout. Ctrl+C is the user-controlled interruption.
**Alternatives considered:** Set a fixed timeout. Rejected because the live run duration is unknown and a valid calculation could exceed an arbitrary limit.
**Rationale:** Keeps the user informed without prematurely terminating a long calculation.

### Decision 76: Stage 4 Button And Audit Panel Like Stages 1-3 (2026-11-04)
**Question:** Should Stage 4 also have a Draft button, in the same fashion as the other stages?
**Decided by:** User, 4 November 2026 (the request). Decision 77's implementation details were initially proposed by the agent and approved by the user at 14:30.
**Choice:** Stage 4 gets the same integration as Stage 3:
- a **Run MACalc** button on `I. Control`;
- an `SST_Stage4_*` audit panel below the Stage 3 panel;
- a `RunMACalc` VBA macro that calls `LaunchBondsLite 4`;
- `--stage 4` dispatch in `bonds_lite_input.main`;
- an `--install-stage-four` installer, run with the Draft closed after a backup.

The packaged controller and a new packaged MACalc model app are built, and the release promoted, only after the source-mode live run passes and with separate approval, as for Stage 3. Supersedes the deferral in Decision 60 and the CLI-only entry and no-Draft-writes rule in Decision 73.
**Alternatives considered:**
- Keep the source-mode CLI and add the button later: rejected by the user's request.
- Add the button without an audit panel: rejected because it would differ from Stages 1-3, and the user asked for the same fashion.
- Install the button in the Draft now, before Stage 4 code exists: rejected because the active release `20261104-cache-v9` rejects `--stage 4`, and this session's instruction forbids changing the Draft.
**Rationale:** One consistent way to start every stage, with an audit trail in the Draft, while the order (tests, code, source live run, then release) protects colleagues from a broken button.

## Agent-made decisions pending user approval

### Decision 77: Stage 4 Button, Panel And Packaging Details (2026-11-04)
**Question:** Which names, audit fields and packaging approach should the Stage 4 button use?
**Decided by:** User, 4 November 2026 at 14:30.
**Choice:**
- **Names:** button `btnRunMACalc`, caption `Run MACalc`, macro `RunMACalc`, panel heading `Stage 4 - Run MACalc`, panel name `SST_Stage4_Panel`.
- **Audit fields:** the same eight as Stage 3: `SST_Stage4_LastRun`, `Start`, `Finish`, `Elapsed`, `User`, `Status`, `InputRun` and `LastSuccessfulOutput`.
  - `InputRun` holds the portable path of the audited Stage 3 workbook.
  - `LastSuccessfulOutput` holds the portable path of `<stem>_unoptimised_assets.csv`.
  - Status goes `Running`, then `Succeeded`, `Failed` or `Interrupted`. The Draft is saved after each change, as in Stage 3.
- **Placement and installer:** use the Stage 3 code, `_install_stage_three_panel` and `_stage_three_button_row`, with the same occupancy and overlap checks. Placement is relative to the Stage 3 panel and button. The installer inserts `RunMACalc` into `modBondsLiteLauncher` only if it is missing. `VBA_LAUNCHER` and `PACKAGED_VBA_LAUNCHER` also gain `RunMACalc`.
- **Packaged runtime:** `tools/build_packaged_stages.py` builds a third app, `MACalc/MACalcModel.exe`, from `src/macalc_batch_entry.py` under the model Python. It bundles a snapshot of the MACalc Batch sources taken at build time and pulp's solver data. Their SHA-256 hashes go in `release.json`, which replaces P4-4's run-time git and hash check in packaged mode. In frozen mode, `load_macalc_config` resolves `MACalc/MACalcModel.exe` beside `Controller`.
**Alternatives considered:**
- Panel fields named after Stage 2 (`LastSuccessfulResults`): rejected; Stage 3 is the closer pattern.
- `LastSuccessfulOutput` = the workbook: rejected because it already equals `InputRun` (in place, Decision 61); the CSV is the new artefact.
- Running MACalc from the shared OneDrive folder in packaged mode: rejected because the release design never runs code from the sync tree, and colleagues have no model Python.
**Rationale:** Mirrors Stage 3, so the code, tests and user experience match, and keeps the packaged release self-contained.

### Decision 78: Line-Buffer The Packaged Bonds Model Console (2026-11-04)
**Question:** How should packaged Bonds Lite model progress be made visible while the batch is running?
**Decided by:** User, 4 November 2026.
**Choice:** Reconfigure `sys.stdout` and `sys.stderr` for line buffering at the start of the model entry point. Do not change the pinned Bonds Lite model sources or rely on `PYTHONUNBUFFERED`, which PyInstaller's frozen interpreter ignores.
**Alternatives considered:** Modify the model's `print()` calls, or configure buffering through an environment variable. Rejected because the model sources are hash-pinned and the frozen interpreter ignores Python environment variables.
**Rationale:** The packaged model entry point is ours to change and can flush each progress line without changing the model snapshot.

### Decision 79: Show Model Heartbeats Only After Quiet Periods (2026-11-04)
**Question:** How often should the controller show `Waiting:` while a model is producing progress?
**Decided by:** User, 4 November 2026.
**Choice:** During a model batch, suppress `Waiting:` while model output has arrived within the last 30 seconds; report a heartbeat after 30 seconds of silence. Outside a model batch, preserve the existing 5-second heartbeat.
**Alternatives considered:** Keep the 5-second heartbeat regardless of model output. Rejected because it obscures the model's own live progress.
**Rationale:** The user's request was to see model progress as it happens while retaining a useful signal if the model becomes quiet.

### Decision 80: Line-Buffer The Packaged MACalc Console (2026-11-04)
**Question:** Should the same packaged console fix apply to MACalc Batch?
**Decided by:** User, 4 November 2026.
**Choice:** Reconfigure `sys.stdout` and `sys.stderr` for line buffering at the start of the MACalc entry point, using a local helper rather than importing across entry points.
**Alternatives considered:** Leave MACalc's packaged output buffered, or add a cross-entry import. Rejected to provide live MACalc progress without introducing an unnecessary PyInstaller dependency.
**Rationale:** MACalc Batch has the same piped-stdout buffering behaviour as the Bonds Lite model.

### Decision 81: Confirm The Live Model Output Changes (2026-11-04)
**Question:** Do Decisions 78-80 reflect the requested implementation?
**Decided by:** User, 4 November 2026.
**Choice:** Confirm Decisions 78-80: line-buffer the two packaged model consoles and show Stage 2 waiting heartbeats only after 30 seconds without model output.
**Alternatives considered:** None.
**Rationale:** The user confirmed the implementation choices recorded in the approved version 0.2 live-output plan.

### Decision 82: Reconcile MACE Funded Assets With The Bonds Input, Not MAP (2026-11-06)
**Question:** Which Stage 1 workbook should define the set of funded assets that Stage 3 compares with MACE?
**Decided by:** User, 6 November 2026.
**Choice:**
- Stage 3 compares the MACE-funded set with the ISINs that have any negative month in `Bonds Lite input_Bonds.xlsx` (`I_tbl_Monthly_Bonds`, Q:ATT).
- MAP still supplies the funded rows' descriptions, and every funded ISIN must exist in MAP.
- Stage 3 stops if `Bonds Lite input_Bonds.xlsx` is missing from the run folder.
- Stage 1 is unchanged.
**Alternatives considered:** Keep the MAP comparison (Decision 37). Rejected because Stage 1 removes negative cashflows from the MAP input and adjusts its PV to MACE, so a correct MAP input has no negative months and the check would always fail.
**Rationale:** The Bonds input keeps the original negative cashflows, so it is the Stage 1 workbook that identifies the funded assets.

### Decision 83: Auto-Save An Unsaved Draft Before Stages 2-4 (2026-11-06)
**Question:** What should Stages 2-4 do when the Draft has unsaved edits, such as a changed MACalc template path?
**Decided by:** User, 6 November 2026.
**Choice:**
- `resolve_stage_input` saves the Draft (with Excel events suppressed) before it reads the Stage 1 handoff, instead of failing with "Save Draft's successful Stage 1 audit before running Stage 2".
- A read-only Draft, or a save that leaves the Draft unsaved, still stops the stage with an explicit message.
- The handoff checks are unchanged: status Succeeded, matching handoff and last-output paths, and a complete run folder.
**Alternatives considered:**
- Keep the block and only reword the message. Rejected because the user preferred not to be interrupted by ordinary edits.
- Drop the check without saving. Rejected because the handoff would then not match the Draft on disk.
**Rationale:**
- Every stage already saves the Draft when it writes its own audit, so the edits were saved anyway once the stage ran.
- Saving first keeps the guarantee that downstream stages use the handoff from the saved Draft, without the manual Ctrl+S step.

### Decision 84: Allow Stage 3 To Drop Template Links Used Only By Replaced Asset Rows (2026-11-06)
**Question:** What should Stage 3 do when the MACalc template has external links that are used only by the asset rows Stage 3 replaces?
**Decided by:** Agent (Supir Debug), 6 November 2026, after the user reported the failure; pending user approval of the release.
**Choice:**
- Before saving, Stage 3 records which template links are still in use, matching them in order by file name. These are the links Excel keeps after the asset rows below the ILS block are replaced.
- Only these links are restored to the template's original paths and verified after saving.
- Each dropped link is logged in the console.
- An added or reordered link still fails with the existing messages.
**Alternatives considered:**
- Keep the strict count check and require templates with no links in the asset block. Rejected because the formulas in those rows are overwritten with values anyway, so the links cannot affect the output.
- Remove the link check entirely. Rejected because it guards against Excel silently re-pathing the remaining links.
**Rationale:**
- The `202512 - JRL - MACalcinPython_v4.0f_100bpsSpread_updREIT.xlsb` template links rows 21+ of `I. Assets` to an old model CSV and to another MACalc workbook. Excel removes those links as soon as the rows are cleared, so the count check always failed.

### Decision 85: Retain Only The Latest Stage 4 Backup (2026-11-06)
**Question:** Should the Stage 4 backup folder keep backups for every Stage 3 workbook, or only the most recent pre-run input?
**Decided by:** User, 6 November 2026.
**Choice:**
- Keep only the latest verified Stage 4 pre-run workbook in `SST Output_<suffix>/Stage 4 backup`.
- Remove older matching MACalc backup workbooks after the new backup has been copied and verified.
- Serialise Stage 4 runs targeting the same output folder within a Windows session, so active recovery backups are not removed by another local run.
- Remove the identified stale Stage 3 staging folder; do not remove staging files while a Stage 3 run is active.
**Alternatives considered:**
- Keep one backup for each distinct MACalc input. Rejected because the user preferred only the latest backup.
- Disable Stage 4 backups. Rejected because a pre-run recovery copy remains useful if MACalc fails.
**Rationale:** Limit output-folder clutter while preserving a verified recovery copy for the most recent Stage 4 attempt.

### Decision 86: Identify The Stage In Console Messages And Hide The MACalc Command (2026-11-06)
**Question:** How should the shared controller identify the stage being run, and should the raw MACalc child command appear in normal console output?
**Decided by:** User, 6 November 2026.
**Choice:**
- Stage runs open with `Starting Stage <N>...` before the workflow is imported; the timing, failure and interruption messages also identify the selected stage.
- Self-test, installation and help commands keep the generic controller banner.
- The raw `MACalc child command:` message is DEBUG-only. Readable MACalc options, progress and errors remain visible, without adding log files or verbosity settings.
**Alternatives considered:**
- Keep the generic opening message and the raw command at INFO. Rejected because the user wanted to identify the stage and avoid technical argument text in normal output.
- Remove the command diagnostic entirely. Rejected because it remains useful when DEBUG logging is enabled during troubleshooting.
**Rationale:** Make the console clearer without losing early startup feedback or useful diagnostic information.

## Property Stress Decisions (Property Stage 1 - Prepare LTM Flash)

### Decision 87: Property Workstream In Its Own Module (2026-11-06)
**Question:** Where should Property stress code live?
**Decided by:** User, 6 November 2026.
**Choice:** All Property stages go in new modules (`src/sst_property_stage_one.py`, later stages likewise); they reuse the generic Excel I/O, folder, guard and audit helpers in `bonds_lite_input.py`. Output goes to `SST Output_<suffix>\Property`, beside `Credit` (one folder per stress). The button and audit panel follow the Credit stages' pattern, Property Stage 1 to the right of the Credit Stage 1 panel and later Property stages below it.
**Alternatives considered:** Extending `bonds_lite_input.py` or the Credit stage modules. Rejected by the user to keep each stress separate.
**Rationale:** The user's request.

### Decision 88: Replace ScenFlipper Contents In Place, With Guards (2026-11-06)
**Question:** How should the template's `I. ScenFlipper` be "replaced" by the Draft's?
**Decided by:** Agent; approved by the user, 6 November 2026 (QP-2, QP-3).
**Choice:** Keep the template's sheet object. Clear the union of both used ranges, copy the Draft's used range to the same address (contents, formats, validation, column widths), and verify exact equality. Guards: sheet in both books; identical non-blank unique row-5 headers; `B2` = number of scenario rows; `B3` in `1..B2`; `B1` in `detailed/GrossandNNEG/total`; no error values; no formula referring to another sheet or workbook. Verify that `prm_*` names, formulae that reference the sheet (for example `I. Inputs!C1`) and `LinkSources` are unchanged.
**Alternatives considered:** Delete the sheet and `Copy After` the Draft's (rejected: `I. Inputs` formulae such as `OFFSET('I. ScenFlipper'!$B$5,...)` become `#REF!`); copy values only (rejected: loses the B1 list validation and formats).
**Rationale:** Gives the Draft's tab content exactly, without breaking the model's references.

### Decision 89: Set The LTM Flash Valuation Date From The Draft (2026-11-06)
**Question:** How should LTM Flash's valuation date be kept in line with the Draft?
**Decided by:** Agent; approved by the user, 6 November 2026 (QP-1).
**Choice:** Write the Draft's `I_Valuation_Date` (`I. Control!D11`) into the template's `ValDate` (`I. Control!D12`), logging the old and new dates, and read it back. Both must be Excel dates.
**Alternatives considered:** Check only and stop when different (safer if the template's date must be set up with other date-dependent inputs; not chosen because Credit Stage 1 also writes the Draft date into its template, Decision 6).
**Rationale:** Consistent with Credit Stage 1.

### Decision 90: No Recalculation Or Link Update In Property Stage 1 (2026-11-06)
**Question:** Should Stage 1 recalculate LTM Flash after the changes?
**Decided by:** Agent; approved by the user, 6 November 2026 (QP-4).
**Choice:** Open with `update_links=False`, manual calculation and `CalculateBeforeSave` off; save without recalculation. The LTM run stage recalculates.
**Alternatives considered:** Full recalculation (needs the JRC add-in and `\\server090` links; could save `#NAME?` or stale-link errors into the output).
**Rationale:** Stage 1 only prepares inputs; recalculation belongs with the model run.

### Decision 91: Property Stage 1 Output Filename (2026-11-06)
**Question:** What should the saved LTM Flash be called?
**Decided by:** Agent; approved by the user, 6 November 2026 (QP-5, QP-7).
**Choice:** `<template stem>_<suffix><template extension>`, e.g. `LTM_FlashModel_v3.2a_Test.xlsb`, built in a staging folder inside `Property` and renamed on success. A rerun with the same suffix replaces `Property` (Decision 59 safe-replace), leaving `Credit` untouched.
**Alternatives considered:** The template's own name (no run identity); `LTM Flash_<suffix>.xlsb` (loses the model version).
**Rationale:** Keeps both the model version and the run identity.

### Decision 92: Property Value Flag Gates The Property Workstream (2026-11-06)
**Question:** Should `I. Control!D18` ("Property Value") control Property Stage 1?
**Decided by:** User, 6 November 2026 (QP-6; the agent had proposed no gate).
**Choice:** Yes, as the two Credit flags gate the Credit workstream (Decisions 55 and 58). The installer names D18 `SST_Property_standalone` after checking that C18 reads `Property Value`, the same way `initialise_stage_path_names` names labelled controls. If the name already exists it is kept, so the user can move the control. Each Property stage reads the flag at the start and requires an Excel TRUE/FALSE. If the flag is FALSE or not TRUE/FALSE, the stage stops with a clear message before creating or replacing `Property`. As in Credit Stage 1, the run is recorded as `Failed` and the previous output is kept. Property gets its own flag constant and check in `sst_property_stage_one.py` and is not added to `STRESS_OUTPUT_FOLDERS`; `allocate_stress_folder` gets a separate allow-list entry for `Property`.
**Alternatives considered:** No gate (the agent's proposal; rejected by the user). Adding `Property` to `STRESS_OUTPUT_FOLDERS` (rejected: Credit Stage 1's "at least one flag TRUE" check would then pass with only Property TRUE, and Credit Stages 1 and 4 would need the Property name).
**Rationale:** The user wants one Draft flag per stress to decide whether that stress runs, without the Credit and Property checks affecting each other.

### Decision 93: Property Stage 1 Names And Launcher (2026-11-06)
**Question:** Which identifiers should the Property Stage 1 button use?
**Decided by:** Agent; approved by the user, 6 November 2026 (QP-8).
**Choice:** CLI `--stage P1` and `--install-property-stage-one`; VBA `PrepareLTMFlash` -> `LaunchBondsLite "P1"`, with `LaunchBondsLite` changed to `ByVal stage As String`; button `btnPrepareLTMFlash` ("Prepare LTM Flash"); panel `SST_PropStage1_Panel` ("Property Stage 1 - Prepare LTM Flash"); fields `SST_PropStage1_{LastRun, Start, Finish, Elapsed, User, Status, InputTemplate, LastSuccessfulOutput}`.
**Alternatives considered:** Numeric stage 5 (implies it follows Credit Stage 4); a second VBA launcher sub (duplicates launcher code).
**Rationale:** Workstream-prefixed IDs scale to later stresses (P2, ...).

### Decision 94: Share The Draft Run Guard Across Stresses (2026-11-06)
**Question:** Can a Property stage run while a Credit stage is running?
**Decided by:** Agent; approved by the user, 6 November 2026 (QP-9).
**Choice:** No; reuse `workbook_run_guard` on the Draft.
**Alternatives considered:** A Property-only mutex (both stages save the Draft's audit cells, which risks conflicting saves).
**Rationale:** One writer to the Draft at a time.

### Decision 95: Place Property Stage 1 To The Right Of Credit Stage 1 (2026-11-06)
**Question:** Where exactly should the button and panel go?
**Decided by:** User (position); agent (geometry), approved by the user, 6 November 2026.
**Choice:** At install, from `SST_Stage1_Panel` (O9:V19): one gap column, then a button the size of `btnGenerateBondsLiteInput` at the same relative row, then an 11 x 8 panel at the same top row; today X13:AB15 and AC9:AJ19. The same occupancy, merge and overlap checks as Stage 3/4 apply. Later Property stages go below.
**Alternatives considered:** Panel directly right of the Credit panel with the button below it (breaks the button-then-panel arrangement the user set up).
**Rationale:** Mirrors the user's Credit column, so each stress is one column of buttons and panels.

## Agent-made decisions pending user approval

### Decision 96: Proceed Without Unavailable Team-Specific Skills (2026-11-06)
**Question:** Which team-specific conventions should apply when the repository team memory and `.github/supir-agents.lock` are both missing, and the question tool is unavailable?
**Decided by:** Agent (pending user approval).
**Choice:** Proceed with no team-specific skills for this implementation, record `none` in the repository team memory, and follow the approved plan and shared Just skills.
**Alternatives considered:** Infer a team from the business domain (unsupported by repository evidence); block implementation until the user can be asked (unnecessary for executing the already-approved plan).
**Rationale:** No declared team information is available, so applying team-specific rules would be speculative. The choice is recorded for the user to confirm or override.

### Decision 97: Property Stage 2 Runs LTM Flash In Place (2026-11-07)
**Question:** What should Property Stage 2 execute, how should it manage the LTM workbook, and what should count as success?
**Decided by:** User, 7 November 2026.
**Choice:** Stage 2 reads the latest successful Stage 1 workbook handoff and requires the Property flag to be TRUE and Stage 1 status to be Succeeded. It opens that workbook in a separate visible Excel instance, recalculates without updating external links, and calls the existing LTM macro `CallRunFrozenPython "flipper"` directly, without clicking the multi-run button or recreating its option prompt. The model updates the same workbook; Stage 2 creates no output workbook or folder. It saves in place only after fresh run timestamps, all `Run N - End of run!` markers through `prm_NoOfScens`, the terminal `End of multirun!` message, and non-empty `O. CFOutput`, `O. PVOutput`, and `O. ConcRisk` are verified. It also verifies `Workbook.Saved` and a changed file modification time. Row-1 run-number validation is deferred. There is no timeout. The macro is allowed to clear its usual `O. NoteVal` and `O. NoteValRebase` areas. Add a Stage 2 button and panel below Property Stage 1 using the Credit Stage 2 placement pattern. Use CLI `P2`, Draft macro `RunLTMFlash`, button `btnRunLTMFlash`, panel `SST_PropStage2_Panel`, and fields `SST_PropStage2_{LastRun, Start, Finish, Elapsed, User, Status, InputWorkbook, LastSuccessfulRun}`. `LastSuccessfulRun` is a timestamp retained after failures. No team-specific conventions apply.
**Alternatives considered:**
- Recreate the button's 0/1/2 prompt or click its event procedure — rejected because the user asked to run the macro path for option 2 directly.
- Run in the Draft's Excel instance — rejected; a separate visible Excel instance keeps the Draft usable and gives the LTM process its own workbook context.
- Update external links or run without recalculation — rejected; recalculate with links disabled as confirmed by the user.
- Save after every macro return or on partial results — rejected; save only after the agreed completion and output checks.
- Add a timeout or process kill — rejected; user chose no time limit, and the macro can display an interactive Excel error dialog.
- Preserve or restore the NoteVal ranges/table — rejected; the user explicitly allows the macro to clear them for now.
- Add row-1 run-number checks — deferred by the user.
**Rationale:** The existing macro is the supported LTM execution path and writes results into the Stage 1 workbook. The checks compensate for its lack of a reliable success return, while failure-without-save preserves the committed Stage 1 file.
**Superseded by:** Decision 98 removes the unused ConcRisk check; Decision 99 adds the agreed long-run safeguards.

## User-approved decisions

### Decision 98: Stop Validating O. ConcRisk After Property Stage 2 (2026-11-09)
**Question:** A live Property Stage 2 run failed with "LTM output sheet O. ConcRisk is empty after the run". Should that sheet remain a success check?
**Decided by:** User, 9 November 2026.
**Choice:** Remove `O. ConcRisk` from the Stage 2 output checks. `O. CFOutput` and `O. PVOutput` must still be non-blank. `O. ConcRisk` may be empty or absent. This supersedes that part of Decision 97. Released in `20261109-cache-v22`.
**Alternatives considered:** Keep the check (rejected; the tab is not used, so an empty sheet is not a failure); warn instead of failing (rejected; the user does not need the tab).
**Rationale:** The check rejected otherwise valid runs on a sheet that has no downstream use.

### Decision 99: Property Stage 2 Long-Run Safeguards (2026-11-09)
**Question:** What should Property Stage 2 do when an LTM run takes too long or stops progressing?
**Decided by:** User, 9 November 2026.
**Choice:** Monitor the run without changing the LTM macro. Print a heartbeat every 5 minutes with elapsed time and scenario progress (`Scenario N of M complete`). Warn after 10 minutes without runner activity. Set the time limit to 10 minutes per scenario (`prm_NoOfScens`), with a 30-minute minimum. At the limit, ask the user in the console whether to keep waiting or abort; never abort without an answer. Aborting closes only the LTM Excel instance owned by Stage 2, never saves, restores the Stage 1 workbook if it changed, and records `Timed Out`. Scope is Property Stage 2 only. This supersedes "There is no timeout" in Decision 97.
**Alternatives considered:**
- Fixed 2-hour limit — rejected; a per-scenario limit fits both short and 16-scenario runs (1-2 hours typical, so 160 minutes).
- Warn only at the limit — rejected; the user wants the option to stop.
- Abort automatically at the limit — rejected; the macro can show an error dialog that needs the user's attention.
- Apply the same safeguards to Stage 4 MACalc — deferred by the user.
**Rationale:** A 16-scenario run normally takes 1-2 hours, so a 160-minute limit does not interrupt normal runs. A prompt keeps the user in control when Excel may be showing a dialog.

## Agent-made decisions pending user approval

### Decision 100: Watch The LTM Runner From Outside Excel (2026-11-09)
**Question:** How can Stage 2 observe a run while its main thread is blocked in the synchronous `CallRunFrozenPython` COM call?
**Decided by:** Agent (pending user approval).
**Choice:** A watcher thread in the controller. It identifies the runner through the process tree under the Excel instance owned by Stage 2: `EXCEL.EXE`, then `cmd.exe /C`, then the `LTMflashArunmanager` executable. Discovery uses the Windows Toolhelp snapshot API, and runner CPU time comes from `pywin32`; no new dependency is added. Activity means the total CPU time of the runner and its descendants increased. If the runner has exited but the macro has not returned after 2 minutes, warn that Excel is probably showing an error dialog. Scenario progress comes from a read-only, separate COM connection on the watcher thread to the owned Excel instance, looked up by PID. A rejected or busy call skips that update and never fails the run. After the user chooses to keep waiting, prompt again 30 minutes later. Abort ends the runner tree first, then the owned Excel process.
**Alternatives considered:**
- `psutil` for process data — rejected; it adds a packaged dependency for functions that `pywin32` and the Windows API already provide.
- Read progress on the main thread — impossible while the macro call is blocking.
- Modify the LTM macro or the xlwings VBA to report progress — rejected; the model is not ours to change.
- Kill only Excel on abort — rejected; the runner would be left trying to write to a dead Excel instance.
- Prompt on stall or dialog detection — rejected; the user asked only for warnings there.
**Rationale:** The LTM macro uses `WScript.Shell.Run` with `WaitOnReturn=True`, so the runner is an observable child process. The runner writes progress into the workbook through COM while the macro waits, which suggests Excel will also service a read-only progress query. A live test must confirm this before the progress line is relied on.
**Superseded by:** Decision 102 bounds the duration of a progress COM query so it cannot block the watchdog.

### Decision 102: Bound Property Stage 2 Progress Reads (2026-11-09)
**Question:** Can a progress COM read itself block the watchdog while the LTM macro is busy?
**Decided by:** Agent (pending user approval).
**Choice:** Run each progress read on a dedicated daemon thread and wait no more than 5 seconds. If Excel does not respond, disable progress reads for the rest of that run and continue elapsed-time, process-activity and timeout monitoring.
**Alternatives considered:**
- Read on the watchdog thread without a deadline — rejected; a busy Excel COM call could block heartbeats, stall detection and the time-limit prompt.
- Keep retrying with a new worker on every heartbeat after a hung read — rejected; it could accumulate blocked COM threads.
- Remove scenario progress — rejected; the user requested it, and responsive COM reads can still provide it.
**Rationale:** The watchdog must remain responsive even if Excel queues a read request. A single bounded worker preserves the other safeguards and prevents a repeated-query thread leak. This is a conservative fallback because the progress read has not yet been live-tested.
**Superseded by:** Decision 103 approves this refinement and authorises build and promotion.

## User-approved decisions

### Decision 101: Approve Property Stage 2 Watchdog Implementation (2026-11-09)
**Question:** Should the approved safeguard plan and its technical approach be implemented?
**Decided by:** User, 9 November 2026.
**Choice:** Proceed with implementation of Decisions 99 and 100 and `plan-property-stage-two-safeguards.md`. The live acceptance run and release promotion remain separate approval gates.
**Alternatives considered:** Defer implementation or revise the watchdog design — not chosen.
**Rationale:** The user approved implementation after reviewing the time-limit behaviour, warnings, process monitoring and abort handling.

### Decision 103: Approve Property Stage 2 Safeguards Release (2026-11-09)
**Question:** Should the bounded progress-read refinement in Decision 102 be approved and the safeguards be built and promoted?
**Decided by:** User, 9 November 2026.
**Choice:** Approve Decision 102 and build and promote the Property Stage 2 safeguards after the focused tests, packaged self-tests and published executable hash checks pass. The user requested "Sure, build and promote" after the refinement was explained. Do not initiate a live LTM run; live acceptance remains outstanding and is no longer a prerequisite for this requested release. This supersedes the release approval gate in Decision 101 and the dependency on live acceptance in step 7 of `plan-property-stage-two-safeguards.md`.
**Alternatives considered:**
- Leave Decision 102 pending and defer packaging - not chosen; the user accepted the refinement and requested build and promotion.
- Require a live LTM acceptance run before promotion - not chosen for this release; the user requested promotion now, and live acceptance still needs separate approval.
**Rationale:** The user wants the safeguards available in the packaged workflow. Automated and packaged checks remain mandatory, while live scenario-progress and abort behaviour must still be confirmed in a separately approved run.

### Decision 104: Load JRC In Property Stage 2 And Reject Invalid Curves Or PVs (2026-11-09)
**Question:** A Property Stage 2 run produced an all-zero `O. PVOutput`. How should Stage 2 prevent and detect this?
**Decided by:** User, 9 November 2026 (Option 1 plus both checks).
**Choice:** Before opening the LTM workbook, register in Stage 2's own Excel instance the XLL add-ins that the user's normal Excel opens at start-up. These are read from the `OPEN*` values under `HKCU\Software\Microsoft\Office\16.0\Excel\Options`; on this machine the JUST DataMart loader provides `JRC_SpotAnnual`. Keep the full recalculation. Then stop before the macro if any `tbl_Curves` (`P. Curves`) data value is not numeric. After the run, fail without saving if any scenario in `O. PVOutput` is missing or has no non-zero PV.
**Alternatives considered:**
- Replace the full recalculation with a normal one — rejected by the user; it would rely on curves cached by an earlier Excel session.
- Hard-code the DataMart add-in path — rejected; the path is per-user and could change when the add-in is updated, while the registry entry reflects what normal Excel loads.
- Checks only, without loading the add-in — rejected; the run would fail every time instead of working.
**Rationale:** An Excel instance started by automation does not load start-up add-ins. `CalculateFullRebuild` therefore turned every `P. Curves` value into `#NAME?`. LTM discounted with NaN curves, and its pandas total turned NaN PVs into exact zeros, which passed the non-blank output check. In a disposable-copy test, registering the DataMart loader in an automation instance produced numeric curves equal to the cached values after a full rebuild.

### Decision 105: Reuse Unchanged Model And MACalc Builds (2026-11-09)
**Question:** How can releases that change only controller code be built faster?
**Decided by:** User requested the builder change, 9 November 2026; the design below was chosen by the agent.
**Choice:** After all self-tests pass, the builder caches the `Model` and `MACalc` applications locally, with a SHA-256 fingerprint of the PyInstaller options, entry script, every `.py` file in the model source folder (excluding `__pycache__`, `.git`, `venv`, `build` and `dist`), and the model interpreter version and installed packages. A later build reuses a component only if the fingerprint matches and the cached executable hash is intact. It rechecks the fingerprints before publishing. Reuse is automatic; `--rebuild-all` or `--clean` disables it. `release.json` records `component_fingerprints` and `reused_components`. Self-tests, source checks and executable hashes are unchanged.
**Alternatives considered:**
- A manual `--controller-only` flag — rejected; it relies on the person building to judge correctly that the models have not changed.
- Reference shared component folders across releases instead of copying them — rejected; the Excel launcher installs a whole release folder, so it would need workbook VBA changes.
- Copy the model applications from the previous central release — rejected; that reads about 600 MB from the network, and older manifests have no fingerprint to confirm they match.
**Rationale:** The model and MACalc builds take most of the PyInstaller time and rarely change. A content fingerprint makes reuse safe without manual judgement, and any input change triggers a rebuild.

### Decision 106: Publish Releases With Robocopy And Verify The Copy (2026-11-09)
**Question:** Can the release copy to the central share be made faster?
**Decided by:** User requested the change, 9 November 2026.
**Choice:** Publish with `robocopy /E /MT:16 /R:2 /W:5`, treating exit codes of 8 or above as failure. Then require every staged file to exist in the release with the same size. On any failure, delete the partial release folder before `current.txt` is touched.
**Alternatives considered:** Keep `shutil.copytree` — rejected; it neither verifies the copy nor removes a partial release, and it was slightly slower.
**Rationale:** Measured on the 284 MB, 869-file Model application, `shutil.copytree` took 306 s and robocopy plus verification took 288 s. Both ran at about 1 MB/s, so the network bandwidth, not per-file overhead, limits the copy. The speed gain is small, but the verification and clean-up make failed publications safer. Larger savings would need less data to transfer, for example a smaller bundle or releases that share unchanged components. That would require changes to the Excel launcher.

## Agent-made decisions pending user approval

### Decision 107: Limit The Packaged CBC Solver To Supported Windows Architectures (2026-11-09)
**Question:** Which bundled CBC solver binaries are needed in the Windows MACalc executable?
**Decided by:** Agent (pending user approval).
**Choice:** For a packaging experiment, keep Pulp and its Windows 64-bit CBC solver, and omit Linux, macOS and Windows 32-bit solver binaries. Treat 64-bit Windows as the target platform unless the user confirms that 32-bit Windows must remain supported. Do not remove CBC or its licence files.
**Alternatives considered:**
- Keep all six platform variants — rejected for the experiment; the application is a Windows desktop executable and the self-test checks the Windows solver.
- Remove all solver data — rejected; the packaged MACalc self-test requires CBC availability.
- Keep Windows 32-bit CBC as well — deferred; the build environment and current packaged target are 64-bit, but deployment compatibility must be confirmed before excluding it.
**Rationale:** The v25 MACalc bundle contains about 34.4 MB of solver binaries; Windows 64-bit CBC is about 4.2 MB. This could save about 30 MB if all supported PCs are 64-bit Windows. Keep this as a gated packaging change until the target architecture is confirmed and the packaged smoke test passes.
**Superseded by:** Decision 111 confirms 64-bit Windows as the only supported target.

### Decision 108: Evaluate Numba Removal From MACalc Only (2026-11-09)
**Question:** Can the MACalc bundle omit Numba and llvmlite to reduce package size?
**Decided by:** Agent (pending user approval).
**Choice:** Test excluding Numba and llvmlite from the MACalc executable only. Retain them in the Bonds Lite Model executable, which directly uses Numba for a numerical kernel. Do not accept the MACalc exclusion unless its packaged self-test and a representative MACalc regression run pass.
**Alternatives considered:**
- Remove Numba from both model executables — rejected; the Bonds Lite Model uses Numba directly.
- Keep Numba in MACalc without testing exclusion — deferred; the analysed MACalc source has no direct Numba import, while PyInstaller includes it through dependency analysis.
- Remove it from MACalc based on the self-test alone — rejected; the self-test does not exercise every pandas or model code path.
**Rationale:** The MACalc bundle contains about 84.5 MB of llvmlite files, and none of its packaged Python source files directly imports Numba. This is a material potential saving, but indirect runtime use must be ruled out with an actual regression run before release.
**Superseded by:** Decision 111 retains Numba and llvmlite in both model applications.

### Decision 109: Publish Immutable Shared Runtime Components (2026-11-09)
**Question:** How should unchanged Model and MACalc applications be shared between releases and user caches?
**Decided by:** User approved implementation on 9 November 2026.
**Choice:** Implement an immutable, content-addressed component store on the central share and in each user's LocalAppData cache. Keep the controller with each release; have the release manifest reference the Model and MACalc component fingerprints and executable hashes. Copy only missing or changed components into a staging path, verify them, then atomically publish the component and release manifest before updating `current.txt`. Preserve compatibility with existing full-layout releases and the last-known-good local fallback.
**Alternatives considered:**
- Continue copying the complete bundle for every release — rejected; a controller-only change still transfers about 636 MB.
- Use junctions or hard links between the Model and MACalc `_internal` folders — deferred; the two applications have different `base_library.zip` files and the current launcher copies folders, so this needs a separate runtime-layout proof.
- Reuse components based only on the release ID or executable filename — rejected; component reuse needs explicit fingerprint and hash validation.
**Rationale:** The current launcher copies the entire release to a versioned local folder, while both frozen entrypoints find sibling applications relative to the controller. A manifest-backed component cache can reduce a controller-only release from about 636 MB to about 33 MB, but requires coordinated changes to the builder, VBA launcher and Python path resolvers.

### Decision 110: Keep Real Model Runs And Promotion Outside This Implementation Scope (2026-11-09)
**Question:** Which release and live-run gates apply while reducing package size and changing the component cache?
**Decided by:** Agent (pending user approval).
**Choice:** Validate with unit tests, packaged self-tests, disposable release trees and hash/transfer measurements. Do not run a real financial batch or promote a release as part of implementation unless the user separately approves those actions. Keep v25 as the active release until a new release is approved.
**Alternatives considered:**
- Promote immediately after local tests — rejected; the release changes both packaging content and the launcher/cache contract.
- Run a real MACalc financial batch without separate approval — rejected; it has operational and business-data effects beyond packaging validation.
- Rely only on unit tests for the cache changes — rejected; copy, fallback and VBA behaviour need opt-in Excel integration checks on disposable workbooks.
**Rationale:** A failed packaging trim can affect numerical execution, and a shared-component release changes user installation and rollback behaviour. Build and live acceptance gates should remain explicit and separate.

## User-approved decisions

### Decision 111: Retain Numba And Target 64-bit Windows (2026-11-09)
**Question:** Should the package optimisation remove Numba from MACalc, and must 32-bit Windows remain supported?
**Decided by:** User, 9 November 2026.
**Choice:** Keep Numba and llvmlite in both the MACalc and Bonds Lite Model bundles; do not pursue the MACalc Numba exclusion. Support 64-bit Windows only. The packaging work may omit CBC solver binaries for other platforms while retaining the Windows 64-bit solver, Pulp and their required licence files. Verify the retained solver in the packaged runtime. This supersedes Decisions 107 and 108.
**Alternatives considered:**
- Exclude Numba and llvmlite from MACalc — not chosen by the user.
- Retain CBC binaries for 32-bit Windows, Linux and macOS — not needed for the confirmed 64-bit Windows target.
**Rationale:** The user chose to leave Numba in place and confirmed that 64-bit Windows can be the only supported target. The remaining solver trim avoids unused platform binaries without changing the model's Numba dependencies.
