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

## Agent-made decisions pending user approval

### Decision 54: Unique MACalc Input Filename (Slice D) (2026-11-03, proposed)
**Question:** `MACalc_Batch` (Stage 4) needs each MACalc input workbook to have a name that is unique across synced SharePoint folders. What unique suffix should Stage 3 add to `MACalc input.xlsb`?
**Decided by:** Agent (proposed for the user's approval at the start of the Slice D session, together with plan questions Q-D1 to Q-D4).
**Choice:** Name the output `MACalc input_<Stage 1 run ID>_<token>.xlsb`, where the run ID is the run folder name (`YYYYMMDD_HHMMSS`, with `_NNN` on a same-second collision) and `<token>` is `secrets.token_hex(8)` (16 lowercase hex characters), generated fresh on every Stage 3 run and regenerated if the name already exists in the run folder. Stop before any write if the folder name does not match the run-ID pattern or the full path exceeds Excel's 218-character limit. On a rerun, ask before replacing earlier tokenised outputs of the same run; leave them untouched until the new file is verified, published and the `Succeeded` audit is saved, then delete them (a failed deletion is a warning). The rename-to-backup and restore logic is removed, because a new name never overwrites an old one. Existing `MACalc input.xlsb` files are left untouched. Stage 4 must find its input through `SST_Stage3_LastSuccessfulOutput`.
**Alternatives considered:**
- Stage 3 run timestamp: rejected, because every rerun would add a new file and bypass the overwrite prompt and backup.
- Run ID only (the first proposal): unique only within one output root. Two Drafts starting Stage 1 in the same second, or a copied run folder, give the same name, so it does not meet the requirement.
- Run ID plus the Windows user name: still clashes for one user running two Drafts in the same second, and puts the user name in the file name.
- Run ID plus a token reused across reruns: one stable name per handoff, but a copied run folder carries the token and duplicates it on rerun.
- Full UUID4 (32 hex characters): 122 random bits, but 16 more path characters near Excel's 218 limit. Fallback if the user prefers it (plan Q-D1).
- Stage 3 timestamp: not unique across Drafts started in the same second.
**Rationale:** The requirement is uniqueness across synced folders, which a run ID cannot give because it is unique only within one output root. A fresh 64-bit random token per run makes a clash negligible (about 3 in a trillion among 10,000 outputs) without a central registry, which Stage 3 cannot consult. Keeping the run ID preserves traceability. Because new and old names never coincide, earlier outputs survive any failure by construction, which keeps the P14 guarantee with simpler code. `MACalc_Batch` was not inspected (Decision 51), so its naming rules are still to be confirmed by the user (plan Q-D1).
