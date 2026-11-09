# Bonds Lite Input Setup And Use

Version 1.3 | 9 November 2026 | Local-cache v24 active; Property Stage 2 safeguards

## Overview

Generate Bonds Lite Input on the Draft workbook's I. Control sheet starts the workflow through a small VBA launcher. Stage 1 prepares two validated inputs, Bonds and MAP, using one template and one captured set of date/entity/stresses. Both must succeed before the run folder becomes eligible for Stage 2. Stage 2 runs the unchanged folder-based Bonds Lite batch against that exact pair.

The colleague deployment uses frozen controller/model executables cached under the current Windows user's LocalAppData. Draft stays in the user's locally synced SharePoint/OneDrive folder; the launcher resolves that user's local Draft path and attaches to the exact open Excel instance. The central release share is checked once at the first button click in each Excel session. Later clicks use the local cache without network access. Release v24 is active and Stage 1 reads its three path controls through workbook names.

## Prerequisites

- Windows desktop Excel; Excel in a browser is not supported.
- Draft must be locally synced from SharePoint/OneDrive before opening in desktop Excel. The workflow does not support Excel's browser URL as a filesystem path.
- Colleagues need central release-share access on the first click in each Excel session, write access to their LocalAppData cache, and normal permission to run the approved VBA/executables. They do not install Python or configure interpreter paths after the packaged release is deployed.
- Maintainers using source mode need Python 3.13, the model's separate Python 3.11 environment, and the existing Bonds Lite Batch project. Model files remain unchanged.
- The JRC Excel add-in that provides `_xll.JRC_SpotAnnual` must be available in the Excel session used to run Stage 1. Use the installed Draft button from that session; a separate automation Excel instance may not have the add-in.
- The project folder, including pyproject.toml and src, available at a stable local or synchronised location.
- Access to the selected data/template files and all saved external workbooks required by the destination formulas.
- Permission to run the known Draft macro under your organisation's Excel policy. No xlwings Excel add-in is required. Do not enable all macros or weaken global security settings.

## Installation

### Colleagues: Packaged Runtime

Once the packaged release is built and approved, open the Draft from its local
OneDrive-synced folder in desktop Excel. The first button click in each Excel
session checks the central release pointer and copies a changed release into
LocalAppData, with an Excel status message during copying. The workbook then
launches the cached controller/model locally. Later clicks in that Excel
session do not contact the release share. No Python installation or
environment-variable setup is required.

The packaged Draft launcher strips Windows CR/LF line endings from both the
central and LocalAppData `current.txt` pointers before validating their release
IDs. If the central pointer or release is unavailable or incomplete, it can use
the last complete local cache.

The central share is a package source only; the workbook never launches an
executable from UNC or from inside the OneDrive sync tree. Keep the Draft's
existing scenario inputs, templates and output paths as configured; those data
paths are separate from the application cache. Do not copy application files
between releases or delete LocalAppData cache folders while a run is active.

**Deployment status:** release `20261109-cache-v25` is active in the central
pointer since 9 November 2026 (built from `f9071b9`). Property Stage 2 now
loads the JRC add-in in its own Excel, stops before the macro if `P. Curves`
is not numeric, and rejects a run whose `O. PVOutput` has no non-zero PVs for
any scenario. Other stages are unchanged. The controller, model and MACalc
packaged self-tests passed, and the published executable hashes match
`release.json`.

Live LTM acceptance remains outstanding; scenario-progress reads and abort
behaviour have not yet been verified in a live Excel run. Once any active
run has finished, fully close and reopen Excel so the next button click
checks the promoted release. To roll back, write `20261109-cache-v24` to the
central `current.txt` as plain ASCII without a trailing line break.

Release `20261109-cache-v24` was active earlier on 9 November 2026 (built
from `f1d486a`). It changed the generic Property Stage 2 waiting message from
every five seconds to every minute. Its previous rollback target was
`20261109-cache-v23`.

Release `20261109-cache-v23` was active earlier on 9 November 2026. It added
Property Stage 2 heartbeats, scenario progress when readable, activity
warnings and the user-controlled time-limit prompt. Its three packaged
self-tests passed. Its previous rollback target was `20261109-cache-v22`.

Release `20261109-cache-v22` was active earlier on 9 November 2026. It removed
the requirement for the unused `O. ConcRisk` sheet to exist or contain data
after a Property Stage 2 run. Its controller, model and MACalc packaged
self-tests passed. The previous rollback target was `20261107-cache-v20`.

Release `20261107-cache-v19` (built from `02c2fa0`) was active earlier on
7 November 2026. It fixes Property Stage 2
failing with "Property Stage 1 output workbook is read-only": the LTM template
is saved as *read-only recommended*, so Stage 2 now opens the Stage 1 copy with
that recommendation ignored. Its self-tests passed and its executable hashes
match `release.json`. To roll back, write `20261107-cache-v18` to the central
`current.txt`.

Release `20261107-cache-v18` (built from `2fa2cab`) was active earlier on
7 November 2026. It fixes Property Stage 2
failing with the xlwings "SharePoint configuration" error when OneDrive holds
another file with the LTM workbook's name (such as a restored Stage 1 staging
copy). Stage 2 now maps the workbook's SharePoint URL through the Draft's own
folder. Its self-tests passed and its executable hashes match `release.json`.
To roll back, write `20261107-cache-v17` to the central `current.txt`.

Release `20261107-cache-v17` (Credit Stages 1-4 and
Property Stages P1/P2, built from `4db7698`) was active from 7 November 2026. Its controller, model and MACalc self-tests passed, its
executable hashes match `release.json`, and its controller accepts `--stage P1`
and `--stage P2`. The model and MACalc sources are byte-identical to v16. Rollback from v17 was to
`20261106-cache-v16` in the central `current.txt` (plain ASCII, no
trailing newline).

Release `20261106-cache-v16` (Stages 1-4, built from
`b03090a`) was active from 6 November 2026. Its
controller, model and MACalc packaged self-tests passed. It retains only the
latest verified Stage 4 pre-run backup (Decision 85), identifies the selected
stage in console startup and error messages, and hides the raw MACalc child
command from normal output (Decision 86). Full end-to-end acceptance of v16 is
pending. The user reported a successful Stage 1-4 comparison against a live
example before this release.

Release `20261106-cache-v15` (built from `e7f2d15`) lets Stage 3 accept MACalc
templates whose asset rows hold the only links to a file (Decision 84).
Release `20261106-cache-v14` (built from
`2f7326f`) made
Stages 2-4 save the Draft automatically before starting, so an unsaved edit
such as a changed MACalc template path no longer blocks a run (Decision 83).
Release `20261106-cache-v13`
(built from `eb098ea`) added the
Stage 3 fix that reconciles MACE funded assets with the Stage 1 Bonds input
instead of MAP (Decision 82); its Model and MACalc executables are identical to
v12. Release `20261105-cache-v12` added line-buffered live model output in
Stage 2 with a 30-second quiet-period heartbeat, and retries deletion of
OneDrive read-only `~replacing Credit ...` backups and failed Stage 1 folders.
Release `20261104-cache-v11` was built from stale source (legacy `Bonds Lite
output` folders, no Stages 3/4) and must never be promoted. On v13, the
packaged Stage 3 Draft-button run succeeded on 6 November against
`SST Output_Test/Credit`. The audited output was
`MACalc input_Test_d58742fd77236ce3.xlsb`. Packaged Stage 2 and Stage 4
acceptance remain pending. Each user's LocalAppData cache updates when Excel
next checks the pointer on a button click; restart Excel if it is already open.
To roll back, write `20261106-cache-v14` to the central `current.txt` as plain
ASCII with no trailing newline (the Excel launcher rejects a pointer that ends
in a line break). The
backed-up Draft has the three Stage 1 path names installed. No path-control
movement is required; moving a control is optional.

### Maintainers: Source Mode

Run these commands in PowerShell from the project folder, adjusting the Anaconda path if needed:

```powershell
$python = "$env:LOCALAPPDATA\anaconda3\python.exe"
& $python -m pip install -e ".[dev]"
[Environment]::SetEnvironmentVariable("BONDS_LITE_PYTHON", $python, "User")
[Environment]::SetEnvironmentVariable("BONDS_LITE_MODEL_PYTHON", "$env:LOCALAPPDATA\anaconda3\envs\py311\python.exe", "User")
```

After changing either model environment variable, fully close and restart Excel before clicking Run Bonds Lite. A running Excel process retains the environment it inherited when it started, so the launched controller may otherwise report the model interpreter or directory as missing.

If the model project is elsewhere, set the optional per-user `BONDS_LITE_MODEL_DIR` to the folder containing `run_bonds_batch.py`. Otherwise it resolves relative to Draft as `../../02 Sensitivities/2025/05 ORSA/Models/Bonds Lite Batch/bonds-hm-cst`. There is no fallback to the controller's Python 3.13 environment.

The editable installation points to this project folder. Keep the folder in place; repeat installation if it moves. Each source-mode maintainer configures their own interpreter paths.

The supplied Draft workbook already has the button and audit panel installed. If installation or reinstallation is required, close Draft in Excel and run:

```powershell
& $python -m bonds_lite_input --install-button "Draft Input spreadsheet for SST.xlsm"
```

Installation requires Excel to allow programmatic access to the VBA project under your organisation's policy. If access is blocked, ask your authorised administrator or workbook maintainer; the installer does not change security policy. Original Draft copies are retained under references/backups before installation changes.

The Stage 1 panel was initially installed at I. Control!H73:O83. Stage 2 adds its own named panel below the current Stage 1 panel and controls. Installation checks occupied cells, merges and shape overlap and preserves the existing Stage 1 button geometry on reinstall. The installer creates the Stage 1 path names on legacy Drafts only after validating all three adjacent labels. Keep Draft saved and closed; a timestamped backup is made before installation.

## Configuration

The following Python environment variables are for maintainer source mode only.
The packaged colleague runtime does not read them.

| Setting | Type | Value And Effect |
| --- | --- | --- |
| BONDS_LITE_PYTHON | Per-user environment variable | Absolute python.exe path; required, no built-in default |
| BONDS_LITE_MODEL_PYTHON | Per-user environment variable | Separate model environment's Python 3.11 python.exe; required for Stage 2 |
| BONDS_LITE_MODEL_DIR | Per-user environment variable | Optional model project folder containing run_bonds_batch.py |
| Stage 1 Bonds path | Workbook name `SST_Stage1_BondsDataPath` (initial cell D39) | Data workbook path; installer fills the supplied data filename only when blank |
| Stage 1 MAP path | Workbook name `SST_Stage1_MAPDataPath` (initial cell D40) | Required MAP data workbook path; deliberately left blank by the installer if blank |
| Stage 1 template path | Workbook name `SST_Stage1_InputTemplatePath` (initial cell D42) | Shared input template path for both datasets; installer supplies its default only when blank |
| Draft I. Control!D11 | Excel date | Valuation date copied to output I. Control!D10 |
| Draft I. Control!D9 | Text | Entity copied to output I. Control!D13 |
| Draft I. Credit Stresses | Numeric cells | All 12 blocks required; decimals, copied without scaling |
| Credit downgrade flag | Workbook name `SST_Credit_DG_standalone` (I. Control!D16) | Excel TRUE/FALSE (text "TRUE", 1 or blank stop Stage 1). Checked against the transition matrices `C9:J15`, `C20:J26`, `C31:J37` and `C42:J48` with exact comparisons, where the identity means 1 on the diagonal and 0 elsewhere, including D. TRUE with an identity matrix (no downgrade), or FALSE with a non-identity matrix, is an inconsistency: Stage 1 lists it and asks whether to proceed (see Use) |
| Credit spread flag | Workbook name `SST_Credit_spread_standalone` (I. Control!D17) | Excel TRUE/FALSE. Checked against the eight spread grids (`M:DI` and `DL:HH` for each row group). TRUE with every grid exactly zero, or FALSE with any grid holding a non-zero value, is an inconsistency that Stage 1 lists and asks about. Both flags FALSE always stops Stage 1; both TRUE gives one combined run. The grids are always copied unchanged, whatever the flags say |
| Output folder suffix | Workbook name `SST_output_folder_suffix` (I. Control!D66) | Non-blank text or a whole number (7.0 becomes `7`), without `<>:"/\|?*` or control characters and not ending in a dot or space. Names the run's `SST Output_<suffix>` folder. Reusing a suffix overwrites that folder's `Credit` output, including its Stage 2 results and Stage 3 MACalc inputs. To keep earlier outputs, change the suffix or archive the folder first |
| Output location | Derived, not an input cell | `SST Output_<suffix>\Credit` beside the saved Draft workbook |

Paths may be absolute or filenames relative to Draft. Use the locally synchronised OneDrive path, not an HTTPS link. Draft, Bonds data, MAP data and the shared template must all be distinct files. D41 is unused. On first installation or migration, the installer checks the labels beside D39, D40 and D42 and creates the three workbook names there. After migration, move a path cell with Excel Cut/Paste so its defined name follows it; copying a cell does not move the named control. Do not delete or rename these names. Partial/missing names or labels that do not match the legacy layout are rejected instead of guessed.

Each data source must contain I. Bonds Data and its I_tbl_Monthly_Bonds table, beginning at B8 and ending in column ATT: 1,215 columns. Its headers must match the destination's B:ATT prefix in order. The input template retains its calculated tail through AUN. Bonds and MAP row counts may differ; missing, reordered or additional raw columns are rejected. The old full-width AUN data-source format is no longer accepted.

The supplied source was rescanned as B8:ATT1066, or 1,058 bond rows; the supplied input table was B8:AUN1006, or 998 seed rows. These row counts are not limits: the output table resizes to the source records.

## Use

1. Save Draft in its local or locally synced OneDrive folder and open it in desktop Excel. Do not open it in a browser.
2. Set the named Bonds, MAP and shared-template path cells. They may be moved with Cut/Paste after names are installed. Close the files and required external dependency workbooks in this Excel instance before running. Dependencies must contain the saved data you intend to use. A blank MAP path blocks generation; no MAP filename is assumed.
3. Confirm date/entity and populate every required numeric stress cell. Genuine zero is valid; a blank is not zero.
4. On I. Credit Stresses, use C:J, M:DI and DL:HH for each row group 9:15, 20:26, 31:37 and 42:48. These are the four transition matrices and the two spread grids per currency/sector.
5. Click Generate Bonds Lite Input on I. Control near row 73. Keep Draft open while the asynchronous Python process runs. Do not edit the inputs during a run.
6. Follow the stages in the visible console. If the credit flags and I. Credit Stresses are inconsistent, Stage 1 first lists every inconsistency and asks `Proceed with these inconsistencies? [y/N]`. Answer `y` or `yes` to continue with the stresses exactly as shown. Any other answer, or a blank, cancels the run: no folder is created or replaced and the Stage 1 audit is left as it was, so fix the flags or the stresses and rerun. Wait for the saved-path or failure summary and total processing time. Audit persistence, cleanup and guard release finish before the console asks you to press Enter to close it. A second simultaneous run for the same Draft is rejected without changing the active run's audit. A failure message is not a successful output.
7. After Stage 1 succeeds and its saved handoff is current, click **Run Bonds Lite**. Stage 2 passes that run folder to the batch script once. Close both generated inputs in Excel first. If old results exist, answer `y` to the console overwrite question; blank or any other answer defaults to No. Stage 2 allows up to 60 minutes for the complete batch.

The controller opens with `Starting Stage <N>...`, followed by a stage-labelled timing line. Failure and interruption summaries also identify the stage. Stage 4 hides the raw `MACalc child command:` line from normal console output; it is available only at DEBUG level. The readable MACalc options, progress and errors remain visible.

The workflow generates two workbooks using all credit blocks, regardless of scenario toggles. Date, entity and stresses are captured once, so both outputs use the same settings even if controls are subsequently edited. Stage 2 executes the Bonds Lite model only after a saved Stage 1 handoff is validated. Entity must be nonempty Windows-safe text with no surrounding whitespace because it is used unchanged in directory names.

## Stage 3 Audit And MACalc Input

Click **Prepare MACalc** on I. Control (around row 100). It runs only when Stage 2 has succeeded on the current Stage 1 handoff. It then writes `MACalc input_<suffix>_<token>.xlsb` into that same `SST Output_<suffix>\Credit` folder, where `<suffix>` is taken from the parent `SST Output_<suffix>` folder name; for example `MACalc input_Q2 2026_3f9c2a7be104d65e.xlsb`. The 16-character token is new on every run, so each name is unique across synced folders. Earlier outputs in the folder are kept, and a rerun simply adds a new file (a Stage 1 rerun with the same suffix replaces the whole `Credit` folder, including these files). **Last Successful Output** always names the newest one; Stage 4 reads its input from there and never guesses the name.

While it runs, Stage 3 builds the workbook in `SST Output_<suffix>\Stage 3 staging\` (outside `Credit`) and removes its staging folder when it finishes. An interruption or a OneDrive cleanup failure can leave staging files behind; remove them only when no Stage 3 run is active. The run stops before writing anything if the run folder is not inside an `SST Output_<suffix>` folder, or if the output path is longer than 218 characters, the longest path Excel can open (shorten the output folder path).

The blue panel `SST_Stage3_Panel` beside the button records the following:
- Last Run, Start, Finish, Time and User;
- Status (it starts as Not Run);
- Input Run;
- Last Successful Output.

Like the Stage 1/2 panels, it can be moved with Cut/Paste. To reinstall the controls, close Draft and run `python -m bonds_lite_input --install-stage-three "Draft Input spreadsheet for SST.xlsm"` from `src`. The command backs up the Draft first.

### What Stage 3 writes

One run writes one `I. Assets` sheet. The template's Mortgage/ILS rows (2-20) are kept. After them come two blocks:
- the **funded block**: one row per funded ISIN, from MAP and the MACE `C. Assets - Funding` tab;
- the **asset block**: one row per row of the Stage 2 base-scenario CSV, in CSV order.

Every post-ILS row takes A:C (Comp. A %, Comp. B %, Partial CF Flag) from MACE `Control` K:O. Asset rows take Type, Index Linked? and FSAO from MACE `I. Bonds`, and everything else from the CSV. Zspread is ×10000. Nothing is left below the last asset row.

The CSV is the single file in `<run folder>\<entity> MAP\` whose name contains `py_bonds_Base_Scenario_nozeros`. If a Bonds Lite model update renames it, change `MAP_BASE_OUTPUT_PATTERN` in `src/sst_stage_three.py`.

The run stops before writing anything, and earlier MACalc inputs are left unchanged, when any of the following occurs:
- no CSV or several matching CSVs are found, or the CSV headers or values are invalid;
- a CSV or funded ISIN is missing or duplicated in MACE `I. Bonds` (column D) or `Control` (column K), or has an invalid Type, Index Linked?, FSAO, Comp. A %, Comp. B % or Partial CF Flag (all offending ISINs are listed in one message);
- MACE `Control!C9` differs from the Draft `I_Valuation_Date`;
- the `Control` table is not contained within K:O;
- a post-ILS row would have a blank column A, or a Type MACalc does not recognise;
- an ISIN list check fails (see below).

If the CSV and MAP ISIN sets differ, Stage 3 still succeeds and shows a warning that lists the differences. MACE rows that are not in the CSV are ignored and counted in the console.

### ISIN List Alignment

MACalc matches four template lists against `I. Assets` column D, which now holds the rating-suffixed `Unique_ISIN`. In the same run, Stage 3 rewrites the key column of each list so that it matches column D:

| Name | Sheet | Columns | Content |
| --- | --- | --- | --- |
| `tbl_ForcedCompAperc` | `I. Inputs` | H:I | key, forced Comp A % |
| `tbl_MinCompAperc` | `I. Inputs` | K:L | key, minimum Comp A % |
| `lst_pre16assets` | `P. Parameters` | K | key |
| `lst_REITs` | `P. Parameters` | M | key |

`lst_noPRoC` and `I. Assets` are not changed. The read-back proves that every other cell on both sheets equals the template.

For each list entry:
- An entry that already equals a D value (a Mortgage `Note`, an `Inf_swap_…` row or a funded plain ISIN) is **kept**.
- An entry that equals the plain ISIN (E) of a rating-suffixed asset row is **remapped** to that row's D. The value is copied exactly. An asset row is rating-suffixed when D is exactly E + `_` + the row's CSV `Final_Rating` (column H of the base-scenario CSV), for example `X_BBB` for a row whose `Final_Rating` is `BBB` (its notched rating in column I may be `BBB+`).
- An entry that matches several rows (a funded ISIN that is also an asset, or one ISIN at two ratings) is **fanned out**: one entry per row, in `I. Assets` row order, at the entry's position, with the same value.
- An entry that matches nothing stays in place **unmatched**, as it was already unmatched in the template.
- When no asset row is suffixed, the lists are left as they are.
- Duplicate keys, whether in the template or created by a fan-out, keep their first occurrence. In the two Comp A tables, a duplicate key with a different value stops the run.

The console prints one line per list, for example `tbl_MinCompAperc: kept 0, remapped 106, unmatched 58, added 0, duplicates removed 0`. It also prints the unmatched keys when there are 50 or fewer, and each fan-out.

An ISIN list check stops the run before anything is written, and keeps the previous output, when any of the following occurs:
- `The MACalc template ISIN lists failed validation`: a title (row 5) has changed, row 6 is not blank, a cell above the title is used, a key sits below a blank key, data appears beside or below a list, or a name is no longer the expected `OFFSET(…COUNTA…)` formula;
- `Stage 3 could not align the MACalc ISIN lists`: a key is blank, an error or not text; a Comp A value is blank, an error, non-numeric or outside 0 to 1; a Comp A duplicate has conflicting values; or an asset row's D is neither its E nor exactly E + `_` + its CSV `Final_Rating`;
- `MACalc ISIN list read-back failed`: the saved lists, their names or the other cells on `I. Inputs`/`P. Parameters` differ from what was expected.

Stage 4 runs MACalc Batch only after the current Stage 2 and Stage 3 audits pass. The Stage 3 ISIN-list alignment was checked against MACalc's exact-match column-D contract during Stage 4 discovery.

The packaged button needs a controller release that includes Stage 3 (v5 or later). Release v4 rejects `--stage 3`.

## Stage 2 Audit And Results

Stage 2 has its own movable panel and does not change `SST_Stage2_Input`, the active Stage 1 handoff. Its Input Run records the captured relative run folder; Last Successful Results is set only after both batch rows, required files, unchanged input hashes and the saved success audit pass. A later Stage 1 run does not make an older Stage 2 result current; readiness compares both run identities.

The batch summary must contain exactly one successful row for each expected input, with matching output destinations and suffixes. Required mapped ISINs, Grouped, BaseScenario, BaseScenario nozeros, model input copy and InvExp_Base files must be present in both result folders. InvExp_Base may be empty when investment expenses are disabled. Conditional diagnostic files are not universal requirements.

On an approved rerun, previous result-folder contents and old batch summaries are backed up outside the selected run folder on the same volume. A failed, interrupted, timed-out, invalid or unauditable attempt removes new artifacts and restores the old files. If recovery is blocked, the console identifies the retained backup. Inputs and unrelated run-folder files are not moved. Do not retry a real model failure automatically; resolve its reported prerequisite first.

Earlier verification attempts failed because hidden Excel could not access the OneDrive inputs and recalculated JRC XLL curves to `None`. The current workflow stages byte-verified local copies and freezes only the validated GBP SONIA, GBP GOV and GBP RPI formula caches in those copies; the original Stage 1 workbooks and embedded output paths stay unchanged. The successful end-to-end rerun used `Bonds Lite output/20261028_181415`; its saved audit and two-row summary both report success, and required artifacts were validated in both output folders. The summary is `bonds_batch_summary_20261028_185611.csv` in the run folder.

## Stage 1 Audit And Stage 2 Input

The coral/pale-pink panel beside the button starts as Not Run and is populated by the next audited run:

| Field | Initial Value Cell | Meaning |
| --- | --- | --- |
| Last Run | L74 | Date of the latest accepted attempt |
| Start | L75 | Start time; the stored value includes its full date |
| Finish | L76 | Finish time after generation cleanup; full date retained |
| Time | L77 | Monotonic elapsed duration, displayed as [h]:mm:ss |
| User | L78 | Windows login running Python |
| Status | L79 | Not Run, Running, Succeeded, Failed or Interrupted |
| Last Successful Output | L80 | Clickable relative path to the last successfully committed run folder |
| Stage 2 Input | L82 | Active relative run-folder handoff, set only after both inputs and the audit save succeed |

The two paths use wrapped merged cells through rows 81 and 83. For example, SST Output_Q2 2026/Credit is resolved relative to the current user's local Draft folder, not the user who created the workbook. The panel remains compact: there are no separate Bonds/MAP audit fields.

### Moving And Resizing The Panel

Runtime reads, writes and validation now follow Excel names, not the initial addresses listed above.

1. Select SST_Stage1_Panel in Excel's Name Box to select the complete panel, including both output-path rows.
2. Use Cut, select an empty destination on I. Control, then Paste. Do not copy/paste a duplicate and expect the code to choose the new copy. Keep clear of existing inputs, buttons and other shapes.
3. Adjust column widths and row heights as needed. Inserting rows or columns within the panel is supported when Excel updates the names and keeps the value fields intact.
4. Save Draft before running the next stage.

For a structural redesign, use Formulas > Name Manager to ensure SST_Stage1_Panel encloses the complete panel and each field name below identifies its correct value cell. A field may name one cell, a cell inside one merged value area, or the complete merged value area. The code writes to that merged area's top-left cell.

| Field | Workbook Name |
| --- | --- |
| Last Run | SST_Stage1_LastRun |
| Start | SST_Stage1_Start |
| Finish | SST_Stage1_Finish |
| Time | SST_Stage1_Elapsed |
| User | SST_Stage1_User |
| Status | SST_Stage1_Status |
| Last Successful Output | SST_Stage1_LastSuccessfulOutput |
| Stage 2 Input | SST_Stage2_Input |

Keep the panel and all fields on I. Control. Do not delete or rename these names. Deleted/#REF references, references to another workbook/sheet, multi-area references, overlapping fields and fields outside the named panel are rejected before audit writes or saves. The code never falls back to the old L74/L82 cells when names are broken. Do not move or resize the panel while a run is active.

The existing four-name panel was upgraded without resetting its values, hyperlinks, colours or position. A legacy panel that has already been structurally redesigned cannot always be inferred safely; installation will refuse it rather than guess which cells contain the audit fields.

**Draft is saved automatically at run start and at its terminal audit update. This also saves all other pending edits in Draft.** The source data and original Bonds Lite input template are never saved by this process.

Before generating, Stage 1 records Running and clears Stage 2 Input. If that first save fails, generation does not begin. On success, the complete run folder is committed to both path fields. If either dataset fails or is interrupted, only the new incomplete attempt folder is removed; Stage 2 Input stays blank and previous runs/last-successful links are retained. If Windows refuses cleanup, the console identifies the leftover incomplete folder and preserves the original error; it is not an authorised handoff. This is a latest-run panel, not append-only history.

If both valid input files were created but the final audit save fails, the complete run folder is retained. The console reports its folder and audit failure rather than announcing success. The active handoff is cleared in memory and recovery is saved where possible. This differs from a dataset-generation failure, which removes the incomplete attempt. If recovery cannot be saved, do not assume the workbook on disk contains the latest failure record. Resolve save access and rerun Stage 1.

Panel duration includes the initial audit-save and generation work but excludes Python startup, final audit-save time and the Enter wait. Console total processing time is the broader measure. Midnight crossings and durations greater than 24 hours are supported.

Workbook names SST_Stage1_Status, SST_Stage1_LastSuccessfulOutput and SST_Stage2_Input identify the handoff contract. Stage 2 requires a saved caller, successful status, matching folder links and both fixed nonempty input files inside a real (non-symlinked) stress folder, currently `Credit`, whose parent is `SST Output_<suffix>` directly beside Draft. Legacy `Bonds Lite output\<timestamp>` handoffs, missing members, legacy flat-file links, temporary folders, absolute handoffs and traversal are rejected; rerun Stage 1 once to replace a legacy handoff. There is no fallback to the newest folder or to Bonds alone. `resolve_stage_input` returns the run folder; `resolve_stage_inputs` returns labelled Bonds and MAP file paths.

Stage 2 is installed as an independent **Run Bonds Lite** button. It invokes `python run_bonds_batch.py <local staged input folder>` once from the configured model directory using the separate Python 3.11 environment. The two workbooks are copied byte-for-byte to local temporary storage for Excel COM access; only the three JRC-backed curve formula columns are frozen to their already validated cached numeric values in those temporary copies. Their embedded JRL output paths are unchanged, so result files still go to the corresponding folders inside the audited run. The controller checks saved output paths and scenario files as-is; if those absolute paths do not match the current run/user, regenerate Stage 1 rather than silently rewriting inputs. The temporary batch summary is validated and moved back to the audited run folder only for the existing validation/audit transaction. Original Stage 1 workbooks remain unchanged.

## Console Progress And Timing

The button starts a visible console asynchronously, with unbuffered Python output. It processes Bonds then MAP, opening only the current data file and a unique private template copy for each. It retains formula recalculation and required-link refresh, but does not open the original template for duplicate checks or reopen saved outputs in Excel. No log files are created. Console stages identify which dataset is being processed.

- Start identifies the operation about to run, including each workbook open/close, data import, dependency refresh, worksheet calculation, save and verification.
- Done shows the completed stage's duration and total elapsed processing time.
- Waiting appears every five seconds during longer controller-side operations and shows the active stage and elapsed stage/total time. During a Stage 2 model batch, model lines stream live and Waiting appears only after 30 seconds without model output. It means Python is waiting; it is not proof Excel is responsive or making progress.
- Failed identifies the stage that raised an error. The final message explains the failure; it is not a successful export.
- Total processing time excludes Python startup and the time spent waiting for Enter. It therefore is not an exact measurement from the physical button click.

Timing values in this example are illustrative, not a measured benchmark:

```text
Start: Open workbook: working.xlsx | total 2.0s
Waiting: Open workbook: working.xlsx | stage 5.0s | total 7.0s
Done: Open workbook: working.xlsx | stage 6.0s | total 8.0s
```

The console stays open after success or failure until Enter. Closing its window forcibly is not guaranteed to be graceful cancellation. Ordinary command-line calls do not pause unless --pause is passed; the workbook launcher adds that flag automatically. All Excel/COM work stays on the main Python thread; the waiting reporter uses timing information only.

## Native Copy And Refresh

Excel copies contiguous raw-input blocks through ATT with values-only PasteSpecial. Python does not receive the large raw-data arrays. Excel-side scalar checks verify exact values and distinguish text, numbers and blanks; real Excel errors are rejected. Destination formula columns are excluded from every source copy. In the supplied templates this means copying B:I and L:ATT, retaining destination formulas in J:K and ATU:AUN. Table row counts remain dynamic, and only newly added rows need formula filling.

The data file still contains formulas in J:K; they are neither imported nor explicitly refreshed. The source is opened read-only, with macros/link updates disabled and manual calculation in effect, and is closed without saving. All formula filling, required-link updates and explicit worksheet calculation target only the generated working copy, never the source data workbook or original input template.

The revised input template's escalation lookups reference its own P. Bond Escalation file sheet. The scan found values, not formulas, on that sheet and no required external formula dependencies for the current bond calculation chain. Literal network-path metadata on that sheet is not treated as a formula dependency. Future templates with genuine external formula dependencies still use the existing destination-only refresh checks.

This operation uses Excel's clipboard. Do not copy/paste elsewhere while the transfer is running. Copy mode is cleared on success or failure, but previous clipboard contents are not restored.

The workflow explicitly updates required links and calculates the worksheets. It waits only while Excel is actively calculating; pending work in manual calculation mode is not treated as an active calculation to wait for indefinitely. Stage 1 also verifies that every saved GBP SONIA value in P. Base Curves is numeric after calculation. If `_xll.JRC_SpotAnnual` is unresolved, Stage 1 fails before publishing the pair rather than handing invalid curve caches to Stage 2.

If Excel still reports link status 3 (old) after successful refresh and calculation, the console prints a warning and formula-error validation continues, as approved by the user. Other bad statuses, missing files, update failures and calculated formula errors stop publication. This does not independently prove that source-workbook formulas were recalculated recently: keep those sources current and saved upstream.

Before each save, the generated workbook's output_folder_path defined name is assigned the final absolute path of its matching model-output directory. After saving, the workflow checks Saved, table/formula/stress/control state and that named path, then closes and validates the file before moving it to its fixed final filename. The pair is authorised only when both inputs pass. Persisted-output comparisons remain integration-test responsibilities.

## Outputs

Each attempt creates `SST Output_<suffix>\Credit` beside Draft, where `<suffix>` is `SST_output_folder_suffix`. `SST Output_<suffix>` may already exist (for example from another stress or Stage 3 staging). If `Credit` already exists, Stage 1 overwrites it without asking: after validation (and any inconsistency prompt) it renames the old folder to `SST Output_<suffix>\~replacing Credit <YYYYMMDD_HHMMSS>` and builds a fresh `Credit`. On success the `~replacing` folder is deleted, together with the earlier Stage 2 results and Stage 3 MACalc inputs it held, so rerun Stages 2 and 3. If the attempt fails or is interrupted, the incomplete `Credit` is removed and the earlier one is renamed back. To keep an earlier output, change the suffix or archive the folder before running. The `Credit` folder is created when either credit flag is TRUE; later stresses (Interest, Inflation, Property, FX) will add sibling folders. If an attempt fails, `SST Output_<suffix>` is also removed when this attempt created it and it is now empty. For entity JRL:

```text
SST Output_Q2 2026/
	Credit/
		Bonds Lite input_Bonds.xlsx
		Bonds Lite input_MAP.xlsx
		JRL Bonds/
		JRL MAP/
```

The two input filenames are fixed within each run. Both workbooks are closed after saving and verification. The two entity-labelled directories are empty model-output destinations; Stage 1 does not run the model or write its results.

In the Bonds input, output_folder_path points to the final current-user absolute path of JRL Bonds; the MAP input points to JRL MAP. The template's name currently resolves to I. Control!D19, but the code uses the name rather than that address. Both folders are created before either generated workbook is saved, so embedded paths never point at a temporary staging location.

Audit links are relative and portable. The paths embedded inside generated workbooks are intentionally absolute execution paths for the user who created the run. Stage 2 validates that each saved path points to the correct dataset folder in this run. If another user reuses the inputs, moving the audit link does not rewrite those paths; regenerate Stage 1 under that user's local folder.

The destination tab is I. Sensitivity, singular. Raw inputs replace the resized table body, formula columns are retained/recalculated, and shared date/entity/stresses are applied. Other template controls remain unchanged except output_folder_path. Original datasets and template are not saved. Private working files are removed after each input; prior completed runs survive failures.

## Example

Set D39 to Bonds Lite data template.xlsx, D40 to Bonds Lite data template_MAP.xlsx and D42 to Bonds Lite input template.xlsx. Enter the intended entity/date, the credit flags and a complete approved stress scenario. D40 is deliberately not filled automatically. Missing data paths, unsafe entity text, blank required stresses, non-boolean flags, both flags FALSE or an unsafe suffix report validation errors before creating any output folder. Flag/stress inconsistencies are listed and need a `y` to proceed.

After validation, with `SST_output_folder_suffix` set to `Q2 2026`, a run would create the paired tree shown above and audit SST Output_Q2 2026/Credit. Do not reuse example stresses merely to bypass validation unless those are the intended scenario.

## Troubleshooting

| Message Or Symptom | Action |
| --- | --- |
| Set BONDS_LITE_PYTHON to your Anaconda python.exe before running | Configure the per-user variable and ensure that executable exists |
| No module named bonds_lite_input | Install this project into the exact configured interpreter |
| Bonds/MAP/template path required | Fill D39, D40 and D42. D41 is unused; D40 is not defaulted by the installer |
| Stage output must be an SST Output_<suffix>\<stress> folder | Rerun Stage 1 to create the paired output; legacy `Bonds Lite output\<timestamp>` and single-workbook audit links are not accepted |
| Cannot replace the existing stress output folder | A file in `SST Output_<suffix>\Credit` is open (for example in Excel) or locked by OneDrive. Close it, let OneDrive finish syncing and rerun Stage 1. Nothing was written and the earlier output is untouched |
| Stress output folder is not a real folder and cannot be replaced | `Credit` is a file or a link. Move it away and rerun Stage 1 |
| The replaced output could not be deleted: ...\~replacing Credit ... | Stage 1 succeeded. Read-only (OneDrive) attributes are cleared automatically, so this means a file is open or locked. Delete the named `~replacing` folder by hand when no files are open from it |
| The earlier Credit output could not be restored ... it is kept at ... | Stage 1 failed and could not put the earlier output back. Remove any incomplete `Credit` folder, then rename the named `~replacing` folder back to `Credit` |
| No Stage 2 base-scenario CSV ... rerun Stage 2 | Stage 1 has replaced the `Credit` folder since Stage 2 ran, so its results were removed. Rerun Stage 2, then Stage 3 |
| No stress run is selected | Set `SST_Credit_DG_standalone` and/or `SST_Credit_spread_standalone` to TRUE |
| Proceed with these inconsistencies? [y/N] | Review the listed flag/stress mismatches. Answer `y` only if the stresses are intended; otherwise press Enter, fix the flags or the stresses and rerun Stage 1 |
| Stress flag ... must be TRUE or FALSE | Enter a real Excel TRUE/FALSE, not text or a number |
| Stage run folder has no complete Bonds/MAP input | Do not run Stage 2 on a partial pair. Recreate both inputs with Stage 1 |
| Incomplete run folder could not be removed | Close any file opened from that incomplete attempt and remove it after resolving access. Stage 1 remains failed; prior runs are not deleted |
| Invalid credit stresses | Complete every reported cell with a real finite number; text numbers and booleans are rejected |
| Bond column headers do not match the template schema | Use the correct data table layout; do not silently drop or reorder columns |
| Close the input workbook before running | Close the specified source in the calling Excel instance |
| Close the required dependency before running | Save intended dependency changes yourself, then close that source before retrying |
| Required formula dependency is unavailable | Check the named saved file and your network permissions |
| Required formula dependency could not be refreshed | Check the named file and referenced sheet/ranges in Excel; no stale-value fallback is used |
| Warning: Excel reports a dependency as old (status 3) | The update succeeded but Excel retains its old-status flag. Keep source calculations current upstream; formula-error validation still runs before publication |
| Native copy verification failed | Avoid changing the clipboard during copying and check source/destination values and cell types; publication is stopped |
| Excel formula errors after recalculation | Repair the listed calculated columns or their source dependencies before retrying |
| P. Base Curves[GBP SONIA] is nonnumeric after calculation | Open Draft in the normal Excel session with the JRC add-in loaded, then run Stage 1 using the installed button; do not generate inputs from a separate hidden Excel instance |
| Stage 2 reports nonnumeric GBP SONIA cached values | Do not retry Stage 2 on that handoff. Run Stage 1 from the installed button in Draft's add-in-enabled Excel session, then verify the new audit and input pair |
| Use a local synchronised path, not a workbook URL | Open Draft from the locally synchronised OneDrive folder |
| Unable to read a complete central release and no complete local cache exists | The packaged launcher now accepts CR/LF-terminated central and local `current.txt` files. If this persists, check that the central pointer names a complete release and the LocalAppData pointer names a complete cached release |
| A Bonds Lite run is already in progress | Let the first run finish; do not keep clicking |
| Install the Stage 1 audit panel with Draft closed before running | Close Draft and run the installer; the Python button entry point now requires the installed panel and names |
| Draft is read-only / I. Control is protected | Resolve workbook access/protection through the authorised owner; no stage is started without a writable audit |
| Generated output retained ... audit/handoff could not be saved | The generated file is retained, but Stage 2 is not authorised to use it. Fix Draft save access and rerun Stage 1 |
| Saved MACalc workbook has a different number of external links from the template | Releases before v15 fail when the MACalc template's asset rows (below the ILS block on `I. Assets`) contain the only formulas linking to a file, such as an old model CSV, because Excel removes the link once Stage 3 replaces those rows. From v15 Stage 3 accepts this and lists each removed link in the console (Decision 84). On older releases, paste values over the template's asset rows (row 21 onwards on `I. Assets`) or break those links, save the template, then rerun Stage 3 |
| Draft has unsaved changes but is open read-only / could not be saved (releases before v14: "Save Draft's successful Stage 1 audit before running Stage 2") | From v14, Stages 2, 3 and 4 save the Draft automatically before they start (Decision 83), so ordinary edits such as a changed template or MACE path no longer block a run. This message appears only when that save is impossible: open an editable copy of the Draft, save it with Ctrl+S, then click the stage button again. Stage 1 and Stage 2 do not need rerunning |
| No progress console appears | Check the configured interpreter/package and reinstall the updated workbook launcher with Draft closed; changing Python's VBA string alone does not update the embedded macro |
| Repeated Waiting messages for one stage | Read the stage name and elapsed duration. A slow or blocked Excel call is possible; a waiting message is not a completion percentage or a watchdog |
| Installer emits an RPC/OLE warning on shutdown | Check the saved button/module and retained backup; an installer regression saved them successfully despite a native shutdown warning |

## Verification

Ordinary tests must not start Excel or refresh network data:

```powershell
& $python -m pytest -q -m "not excel and not performance"
```

Temporary-copy Excel tests use a synthetic local dependency by default:

```powershell
$env:BONDS_LITE_RUN_EXCEL_TESTS = "1"
& $python -m pytest -q tests/test_bonds_lite_input.py::test_excel_full_workflow_publishes_validated_copy --tb=short
& $python -m pytest -q -m excel --tb=short
Remove-Item Env:BONDS_LITE_RUN_EXCEL_TESTS
```

Real-network verification additionally requires BONDS_LITE_RUN_NETWORK_TESTS=1. Enable it only intentionally, with permission to refresh the production dependency. Do not confuse an opt-in skip, an unresponsive terminal or a process launch with a passing test.

Current paired-run verification: 421 default tests passed, eight environment/layout cases skipped and 63 opt-in Excel/performance cases deselected in 4.13 seconds. Ten isolated named-output-folder cases passed. The user confirmed the corrected full paired export and paired audited-run/failed-rerun checks passed, using distinct dataset records and row counts. Reopened inputs were checked for their raw data, shared controls/stresses, retained formulas and matching final output_folder_path. Source bytes and earlier runs remained unchanged. No production model execution or whole-run timing guarantee is implied.

Historical raw-only verification: 243 Excel-independent tests passed in 0.94 seconds, including 27 prefix-schema cases; 24 opt-in cases were deselected. The read-only scan of both supplied templates passed. The isolated full export passed in 11.82 seconds including fixture setup and teardown, with source-link-update/source-calculation guards, unchanged source snapshots and calculated destination lookup checks. Native Windows RPC diagnostics appeared during test cleanup after the passing assertion result; the cleanup warning remains a residual risk.

Stage-audit verification: 287 default tests passed in 1.14 seconds, one read-only layout test skipped and 26 opt-in Excel cases deselected. The blank-workbook panel test, temporary Draft installer, actual installed-panel read-only check, and user-run full audited export all passed. The full audited export test saved/reopened the audit, verified the portable handoff, then checked that a failed rerun blocks Stage 2 and retains the successful output without changing data/template sources. The live test took 11.33 seconds including fixtures; native RPC cleanup diagnostics remain unresolved.

Movable-panel verification: 295 default tests passed in 1.58 seconds, with one read-only layout case skipped and 48 opt-in cases deselected. The live movement/resizing and invalid-layout cases passed; all six legacy migration cases passed after clarifying the invalid-name error. The real workbook's saved named-field upgrade passed its read-only check. Test coverage includes writes leaving the old panel area untouched, row/column insertion, merged-field resizing and save/reopen handoff. Native Excel RPC teardown diagnostics remain a separate unresolved issue.

The run guard is local to the Windows session, not a cross-user/coauthoring lock. Do not have colleagues run stages simultaneously against coauthored copies of the same workbook and assume the Last Run panel is a distributed audit system.

The earlier native-copy benchmark and row-count results are retained in the plan as historical evidence. Revised raw-only row-count/performance fixtures cover the ATT source and wider destination, but were not all rerun this session. A complete intended-data runtime and clean multi-user Excel lifecycle remain separate gates; no whole-run speedup is claimed. This Python-only change does not require reinstalling the button.

## Stage 4 - Run MACalc

Stage 4 runs the existing MACalc Batch model on the exact workbook recorded in `SST_Stage3_LastSuccessfulOutput`. Keep the Draft open in desktop Excel so the launcher can read the saved audit. Stage 4 stops if Stage 2 or Stage 3 is unsuccessful or stale, the workbook is missing or locked, or MACalc Batch would find more than the audited workbook.

The Stage 4 panel records Last Run, Start, Finish, Time, User, Status, Input Run and Last Successful Output. Input Run names the Stage 3 workbook; Last Successful Output names the results CSV. Stage 4 changes only its own `SST_Stage4_*` audit fields.

### Prerequisites and configuration

- Stage 2 and Stage 3 must have succeeded for the current Stage 1 run.
- Close the audited MACalc input workbook in every Excel window. Leave the Draft open.
- Set the per-user `BONDS_LITE_MODEL_PYTHON` variable to the MACalc environment's Python 3.11 `python.exe`. Stage 4 uses the same setting as Stage 2; it does not use the controller's Python for model imports.
- Optionally set `SST_MACALC_DIR` to the MACalc Batch folder. By default, Stage 4 looks for `MACalcrunmanager.py` at `<Draft folder>\..\..\00 Ad Hoc\MACalc_Batch`.

### Install the controls

With the Draft closed, run this from the project's `src` folder:

```powershell
python -m bonds_lite_input --install-stage-four "Draft Input spreadsheet for SST.xlsm"
```

The installer backs up the Draft under `references\backups` before adding the Stage 4 panel, **Run MACalc** button and `RunMACalc` macro. It requires the Stage 3 panel and stops rather than overwriting occupied cells. The button requires a controller release that supports `--stage 4`; release `20261104-cache-v10` and later do.

### Run from source

Maintainers can run Stage 4 from source, without the packaged release, from the project's `src` folder with the Draft open in desktop Excel:

```powershell
python -u -m bonds_lite_input --caller "<Excel Workbook.FullName>" --caller-local "C:\path\to\Draft Input spreadsheet for SST.xlsm" --excel-hwnd <Excel-window-handle> --stage 4
```

Use the Python 3.13 controller interpreter for `python`. Set `--caller` to the exact `Workbook.FullName` reported by Excel; for a OneDrive-synced Draft this can be a SharePoint URL. Set `--caller-local` to the Draft's local synced filesystem path. The window handle is the `Application.Hwnd` of the Excel instance containing the Draft.

### Maintainers: package Stage 4

The Stage 4 bundle adds `MACalc\MACalcModel.exe` alongside the existing
`Controller` and `Model` folders. The builder requires `--macalc-dir` pointing
to the accepted MACalc Batch working copy, as well as `--model-dir` for Bonds
Lite. It uses the isolated Python 3.11 model build environment for both models.
The model build requirements include `pulp==3.2.2` and `smithwilson==0.2.0`;
refresh that isolated environment before an approved build. Do not install
these dependencies into the external MACalc working copy's environment.

The MACalc self-test imports the bundled model and checks CBC availability
without opening Excel. The builder records the six MACalc source hashes and
the MACalc executable hash in `release.json`, and refuses publication if
the sources change during the build. Frozen runs use the bundled modules;
they do not require the external model folder, Python installation or git.
The controller validates the MACalc executable against the manifest before
dispatching a packaged stage.

The builder reuses the previous verified `Model` and `MACalc` builds when
their inputs are unchanged, and rebuilds only the controller. After a build
passes all three self-tests, it caches each model application under
`%LOCALAPPDATA%\BondsLiteBuild\components`, with a fingerprint of everything
that determines it:

- the PyInstaller options
- the entry script in `src`
- every `.py` file in `--model-dir` or `--macalc-dir`
- the model Python version and installed package versions

On the next build, a component with a matching fingerprint and an intact
executable is copied from the cache. Otherwise it is rebuilt and the cache is
replaced. The console prints `Reusing unchanged BondsLiteModel build ...`
when it reuses a component. `release.json` lists the fingerprints under
`component_fingerprints` and the reused applications under
`reused_components`. All three self-tests and the source checks still run on
every build. Pass `--rebuild-all` (or `--clean`) to force a full rebuild.
The release copy to the central folder is unchanged and still includes every
application, because the Excel launcher installs whole releases.

The builder publishes with `robocopy /MT:16` and then checks that every staged
file exists in the release with the same size. If the copy fails or is
incomplete, it deletes the partial release folder and leaves `current.txt`
unchanged. The copy is limited by network bandwidth (about 1 MB/s measured on
9 November 2026), so multi-threading saves only about 5%. Expect about
10 minutes to publish a full 636 MB release.

Commit, build, promotion and the packaged button run are separate approval
gates. Build under a new release ID without `--promote`, verify all three
self-tests, and obtain approval before changing `current.txt`. Release
`20261104-cache-v10` was built and promoted this way on 4 November 2026. For the
packaged acceptance run, first rerun Stage 3 for a clean input and archive
older MACalc inputs outside the stress folder; follow the checks in the
[Stage 4 handoff](session-handoff-2026-11-04-stage-four.md).

### Input, backup and outputs

MACalc Batch runs in place on the audited workbook in `SST Output_<suffix>\Credit`. It does not create a separate Stage 4 input copy. Stage 4 verifies the pre-run backup at `SST Output_<suffix>\Stage 4 backup\<workbook name>`, then removes older matching MACalc backups so the folder retains only the latest verified copy (Decision 85). Runs targeting the same output folder are serialised within the Windows session. If MACalc fails, its pre-run backup is retained; "latest" means the most recent attempt, not the most recent successful run. If you rerun Stage 4 without rerunning Stage 3 first, the backup contains the prior Stage 4 result; rerun Stage 3 when you need a clean input.

The expected result is `<workbook stem>_unoptimised_assets.csv` beside the workbook. MACalc runs with binary optimisation OFF, the optimisation pass OFF and the cost-of-default floor ON. The Draft's `SST_Credit_DG_standalone` value is passed as the downgrade flag; that option is currently inactive while optimisation is OFF. A successful process and a non-empty CSV confirm execution, not the reasonableness of the figures. Review the results before sign-off.

Stage 3 keeps earlier token-named `MACalc input_*` workbooks in `Credit`. MACalc Batch scans recursively, so Stage 4 requires exactly one matching workbook: the audited Stage 3 output. Before running, archive or remove older MACalc inputs from `Credit`; moving them to a subfolder inside `Credit` does not help. Keep any archive outside the stress folder, for example in `SST Output_<suffix>\Stage 3 archive`.

### Troubleshooting

| Message or symptom | Action |
| --- | --- |
| `Stage 3 is stale or unsuccessful; rerun Stage 3` | Rerun Stage 3 after Stage 2 succeeds for the current Stage 1 run. |
| `Stage 3 workbook is locked or open in Excel` | Close that MACalc input workbook in every Excel window, then retry. |
| `MACALC Batch must find exactly the audited Stage 3 workbook` | Archive or remove earlier `MACalc input_*` files from the stress folder; do not move them into a subfolder of `Credit`. |
| `Set BONDS_LITE_MODEL_PYTHON to the model environment's python.exe` | Set the per-user variable to the existing Python 3.11 interpreter and open a new console. |
| `MACalc environment preflight failed` | Check that `BONDS_LITE_MODEL_PYTHON` is valid and that the configured MACalc Batch directory contains importable model modules. Set `SST_MACALC_DIR` if the default folder is not correct. |
| `MACalc Batch exited with code 1` | Read the traceback in the console. The automatic backup contains the workbook as it was immediately before this attempt. |

## Property Stage 1 - Prepare LTM Flash

Property Stage 1 prepares the LTM Flash workbook for the Property stress
workstream. It uses the template named by `SST_LTMFlash_input_template`, copies
the Draft's `I. ScenFlipper` values into the existing template sheet, and sets
the template's `ValDate` to the Draft's `I_Valuation_Date`. The sheet is
overwritten in place so formulas and names in the LTM Flash workbook continue
to refer to the same sheet.

The named Draft flag `SST_Property_standalone` controls this workflow. The
installer names `I. Control!D18` only when `I. Control!C18` is labelled
`Property Value`; an existing valid name is preserved. If the flag is FALSE,
the stage records a failed attempt and retains the previous successful output.

The output is saved in `SST Output_<suffix>\Property` beside the Draft as
`<template stem>_<suffix>.<extension>`. A rerun with the same suffix replaces
only the `Property` folder; the existing `Credit` folder is preserved and the
previous Property output is restored if the attempt fails. The stage does not
recalculate LTM Flash or update its external links. The later LTM run is
responsible for calculation.

### Source-mode installation and run

Close the Draft in Excel and install the new audit panel and button from the
project folder:

```powershell
python -m bonds_lite_input --install-property-stage-one "Draft Input spreadsheet for SST.xlsm"
```

The installer creates a timestamped backup under `references\backups`. As
with other workbook installers, Excel must permit programmatic access to the
VBA project. The installed **Prepare LTM Flash** button invokes Property
Stage 1 through the source-mode controller. It shares the Draft run guard with
the Credit stages, so only one stress stage can run at a time.

The `P1` command token is available to the source controller and, from
release `20261107-cache-v17`, to the packaged controller.

## Property Stage 2 - Run LTM Flash

Property Stage 2 requires the Property flag to be Excel TRUE and the Stage 1
audit to say `Succeeded`. It uses only the workbook recorded in
`SST_PropStage1_LastSuccessfulOutput`; it does not select a newer file from the
Property output folder. The Stage 1 workbook must be a non-empty `.xlsb` or
`.xlsm` file, and the LTM model executable referenced by its macro must be
available.

Close the Draft in Excel and install the Stage 2 audit panel and button from
the project folder:

```powershell
python -m bonds_lite_input --install-property-stage-two "Draft Input spreadsheet for SST.xlsm"
```

The installer creates a timestamped Draft backup under `references\backups`.
It places the `Run LTM Flash` button and audit panel below the current Property
Stage 1 controls, preserves controls that have been moved or resized, and
rejects partial named installations or occupied cells/shapes. Excel must allow
programmatic access to the VBA project.

The Stage 2 button opens the handed-off LTM workbook in a second, visible Excel
instance. It does not update external links. Excel does not load start-up
add-ins in an instance started this way, so Stage 2 first loads the XLL
add-ins that your normal Excel opens at start-up (the `OPEN` entries under
Excel's options in the registry). This includes the JUST DataMart add-in that
provides `JRC_SpotAnnual`. After a full recalculation, Stage 2 checks that
every `P. Curves` value is numeric and stops before the macro if any is not.
It then calls the existing `CallRunFrozenPython("flipper")` macro directly, without
opening the macro button's option prompt. The model updates the same workbook;
Stage 2 creates no additional output workbook or folder. Keep the second Excel
instance visible and attend to any Excel or VBA error dialog.

Stage 2 saves in place only after confirming fresh `RunStartTime` and
`RunEndTime` values, completion markers for every scenario, the final
`End of multirun!` progress message, and non-blank output on `O. CFOutput`,
`O. PVOutput`. Every scenario must also appear in `O. PVOutput` with at least
one non-zero PV. The unused `O. ConcRisk` tab is not validated and may be empty
or absent. It then confirms Excel reports the workbook saved and its
modification time advanced. If a run or check fails before that commit, the
Stage 1 workbook is closed without saving. If post-save verification fails,
the original workbook is restored from a temporary pre-run copy. A
successfully verified LTM workbook is retained if only the final Draft audit
save fails; the audit reports the distinct workbook and Draft states and
retains the previous `Last Successful Run`.

From release `20261109-cache-v23`, while the LTM macro runs, Stage 2 reports
elapsed time and, when Excel allows a read-only progress query, the number of
completed scenarios every five minutes. The query has a five-second limit;
if Excel does not respond, scenario updates are disabled for the rest of that
run while the other safeguards continue. Stage 2 warns if the runner has had
no CPU activity for 10 minutes.
If the runner exits but the macro is still waiting after two minutes, check the
LTM Excel window for an error dialog. These warnings do not stop the run.

From release `20261109-cache-v24`, the controller prints generic
`Waiting: Processing` messages every minute for Property Stage 2 only,
leaving the five-minute LTM updates unchanged. Other stages retain their
five-second generic waiting interval.

From release `20261109-cache-v25`, Stage 2 loads the JRC add-in and adds the
curve and PV checks above. Earlier releases could save an `O. PVOutput` of
all zeros (with blank ERC columns in `O. CFOutput`), because `P. Curves`
recalculated to `#NAME?`; re-run Stage 2 for any workbook produced that way.
The checks report these errors:

| Error | Cause | Action |
|---|---|---|
| `LTM P. Curves contains errors or blanks after recalculation; the JRC add-in (JRC_SpotAnnual) did not load in the Stage 2 Excel` | The DataMart/JRC add-in is not registered for your normal Excel, or failed to load | Confirm `JRC_SpotAnnual` calculates in normal Excel, then re-run. Nothing was run or saved. |
| `LTM O. PVOutput has no non-zero PVs for scenario(s) ...` | The run completed but those scenarios produced no PVs | The workbook was not saved. Check `P. Curves` and the LTM progress log, then re-run. |

The time limit is 10 minutes per scenario, with a 30-minute minimum (160
minutes for 16 scenarios). At the limit, the console asks whether to wait 30
more minutes or abort. A blank response or no response keeps the run going.
Choosing abort ends the LTM runner and only the separate Excel instance owned
by Stage 2; it never closes the Excel instance containing the Draft. Stage 2
does not save the LTM workbook and restores the Stage 1 copy if it changed.
The audit status is `Timed Out`. In source mode only, `SST_P2_TIME_LIMIT_MINUTES`
can temporarily override the limit for acceptance testing; the packaged
runtime ignores this override.

The LTM macro normally clears `O. NoteVal` and the `O. NoteValRebase` table.
Stage 2 does not snapshot, restore, validate or populate these areas. The
`P2` token is available to the source controller and, from release
`20261107-cache-v17`, to the packaged controller. A real LTM run and live Draft
installation of the Stage 2 button require separate user approval.

## Glossary

- Caller: the open Draft workbook whose button was clicked.
- Dependency: a saved external workbook referenced by a retained formula column.
- Table body: bond records excluding the header row; unlike UsedRange, it determines the actual import size.
- xlwings: the Python library controlling the installed desktop Excel application.