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