# Handoff: Property Stage 2 Planning and Combined Promotion

Version 1.1 | 7 November 2026

## Purpose

Use this handoff with
[`plan-property-stage-two.md`](./plan-property-stage-two.md) to implement
Property Stage 2 and, after Stage 2 is implemented and accepted, promote
Property Stages 1 and 2 together. Stage 1 is implemented and exercised in
source mode. Stage 2 implementation and automated verification are complete;
the real Stage 2 run, live Draft installation, and combined package promotion
remain separate approval gates. No packaged release has been built or promoted.

The approved Stage 1 plan and its detailed progress record are in
[`plan-property-stage-one.md`](./plan-property-stage-one.md). This handoff is a
status and context summary; it does not replace that plan or approve a release.

## Current state

- The Property Stage 1 module, installer, audit flow, source-controller `P1`
  dispatch, package-builder hidden import, tests and setup documentation are in
  the repository.
- The live Draft has the Property button, audit panel, named flag and VBA
  launcher installed. The installer created
  `references/backups/Draft Input spreadsheet for SST_20261107_021453_090119.xlsm`.
- The live named flag `SST_Property_standalone` is restored to TRUE.
- Property Stage 2 is implemented in `src/sst_property_stage_two.py`, with
  source/package dispatch, installer controls, unit tests and setup guidance.
  The Stage 2 installer has not been run against the live Draft, and the LTM
  model has not been launched by Stage 2.
- The last live run was the intentional FALSE-flag negative test, so the audit
  panel currently says `Failed`. `SST_PropStage1_LastSuccessfulOutput` still
  points to the prior successful output:
  `SST Output_Test/Property/LTM_FlashModel_v3.2a_Test.xlsb`.
- The Property output folder contains that workbook only. Its contents and file
  hash were unchanged by the FALSE-flag test.
- The source run was dispatched directly through `main()` in the source
  controller. The installed VBA button itself has not yet been clicked in a
  live acceptance test.

## Stage 1 behaviour and approved decisions

Decisions 87-95 are recorded in [`decisions.md`](./decisions.md). Key
constraints for Stage 2 and combined promotion:

- Keep Property stages in separate modules; reuse generic Excel, folder,
  audit and run-guard helpers from `src/bonds_lite_input.py`.
- Property output is under `SST Output_<suffix>/Property`, independently of
  `Credit`. A same-suffix rerun replaces only the Property folder and restores
  the prior folder on failure.
- `SST_Property_standalone` names the Draft's `I. Control!D18` flag. Each
  Property stage must require an Excel TRUE/FALSE and stop before replacing
  output when the flag is FALSE. Do not add this Property flag to the Credit
  stress registry.
- Stage 1 copies the Draft's `I. ScenFlipper` contents into the existing
  template sheet in place, preserving the sheet object and its references;
  validates scenario structure; sets `ValDate` from `I_Valuation_Date`; and
  does not recalculate or update external links.
- Stage 1 writes its handoff path to
  `SST_PropStage1_LastSuccessfulOutput`. Its CLI token is `P1`; the VBA entry
  point is `PrepareLTMFlash`.
- Property stages share the Draft run guard with Credit stages.
- Stage 1's installed layout is to the right of Credit Stage 1. Decision 95
  says later Property stages go below it. Stage 2's button and panel should
  follow the Credit Stage 2 placement and collision-check pattern.

## Implementation and reference files

| File | Relevance |
|---|---|
| `src/sst_property_stage_one.py` | Stage 1 validation, template preparation, audit, run, installer and Excel verification |
| `references/plan-property-stage-two.md` | Confirmed Stage 2 scope, implementation steps, checks and acceptance gates |
| `src/bonds_lite_input.py` | Shared helpers, `P1` CLI dispatch, VBA launchers and Property output-folder support |
| `tools/controller_entry.py` | Packaged controller stage and installer dispatch |
| `tools/build_packaged_stages.py` | Includes the Stage 1 module as a hidden import |
| `tests/test_sst_property_stage_one.py` | Stage 1 unit and opt-in real-Excel tests |
| `tests/test_bonds_lite_input.py` | Shared workflow, CLI, installer and package-builder tests |
| `references/bonds-lite-setup.md` | Source-mode setup and current packaged-release warning |
| `references/plan-property-stage-one.md` | Approved Stage 1 plan, test strategy and progress |
| `references/decisions.md` | Property decisions 87-97; Decision 96 records the earlier team-context choice |

## Verification completed

- Stage 1 tests: **39 passed, 1 skipped**.
- Property Stage 2 fake-Excel tests: **30 passed, 1 skipped**.
- Focused Property Stage 1/2 and launcher suites: **934 passed, 86 skipped,
  2 failed**. The full suite completed with **1,247 passed, 91 skipped, 2
  failed**. The only failures are the two known unrelated failures listed
  below; no new failures were observed.
- `black --check`, Python compilation for the changed modules, and
  `git diff --check` passed.
- The opt-in Stage 2 Excel installer test is implemented but could not be
  completed in this session: Excel automation terminated Python with COM
  error `0x80010108` while entering the installer. It used a temporary Draft
  copy; no LTM model was launched and the live Draft was not touched. Rerun it
  in a clean, responsive Excel session before live installation.
- Opt-in Excel acceptance test on temporary copies: **1 passed**. It checked the
  workbook output and left the original Draft and LTM Flash template hashes
  unchanged.
- Live source-mode TRUE run succeeded and saved the expected LTM Flash output.
- Live FALSE-flag run recorded `Failed`, retained the last successful output,
  and left the Property output folder and file hash unchanged. The flag was
  then restored to TRUE.
- A OneDrive read-only staging-folder cleanup issue was found. The staged and
  published workbook hashes were confirmed identical, the generated staging
  folder was safely removed, and cleanup was updated to use the existing
  read-only retry helper. A subsequent live TRUE run left no staging folder.
- Black and `git diff --check` passed.
- Full suite: **1,213 passed, 90 skipped, 2 failed**. The two failures are
  outside the Property changes and were also recorded during earlier full
  runs:
  - `test_frozen_main_validates_cached_bundle_before_stage_dispatch` calls
    `main(argv)` although `main` accepts no argument.
  - `test_model_results_transaction_rollback_failure_retains_recoverable_backup`
    expects a recovery path that the current rollback code does not report.

Do not silently fix these unrelated failures as part of Stage 2; decide
separately if they block the combined release.

## Stage 2 agreed scope

The user confirmed the Stage 2 requirements in Decision 97. The detailed
implementation plan and progress record is
[`plan-property-stage-two.md`](./plan-property-stage-two.md). In summary,
Stage 2 opens the latest successful Stage 1 workbook in a separate visible
Excel instance, recalculates without updating links, and invokes
`CallRunFrozenPython "flipper"` directly. It updates that workbook in place,
creates no output folder or workbook, and saves only after the agreed
timestamp, progress-log, output-sheet, and save checks pass. There is no
timeout; the LTM macro may clear `O. NoteVal` and `O. NoteValRebase`. The
Property button and audit panel go below Stage 1 using the Credit Stage 2
placement pattern. Implementation, the real model run, and packaged promotion
remain separate approval gates.

## Combined promotion gates

The Stage 1 package changes are present in source/build configuration, but the
currently deployed packaged release does not include `P1`. Do not install the
Property button into a packaged Draft until the combined release is built,
tested and approved.

Before promoting Stage 1 and Stage 2 together:

1. Complete and approve the Stage 2 plan and decisions.
2. Implement and test Stage 2 without dropping Stage 1 coverage.
3. Test the actual installed Stage 1 and Stage 2 buttons in source mode,
   including the appropriate TRUE/FALSE and rollback paths.
4. Verify packaged controller choices and hidden imports include both stages;
   build and test the combined package from a clean, identified source state.
5. Validate the installed package against the Draft and representative
   workbooks, and obtain explicit promotion approval.

The source-mode P1 path has been tested; the actual installed VBA button path
remains an acceptance item. The packaged release has not been built or
promoted.

## Worktree and team-context cautions

- The repository was already dirty before this work. It still contains
  unrelated workbook and Credit-output modifications/deletions, as well as
  Property implementation and live-test artefacts. Do not run `git reset`,
  broad cleanup, or restore unrelated files. Review `git status --short` before
  staging or committing.
- The live Draft was intentionally modified by the installer and live tests;
  use the timestamped backup above if a controlled restore is required.
- Decision 96 records that no team-specific skills were applied because team
  configuration was unavailable during Stage 1. The repository team memory
  says `none`, and the user reconfirmed that no team-specific conventions
  apply to Stage 2 (Decision 97).
