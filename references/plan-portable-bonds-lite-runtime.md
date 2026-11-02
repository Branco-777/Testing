# Plan: Portable Bonds Lite Runtime

30 October 2026 | Release published and Draft launcher installed; no financial batch run

## Overview

Keep Draft in the user's locally synced SharePoint/OneDrive folder, but do not
run binaries from that sync tree. The VBA launcher reads a versioned release
pointer on the central share, copies and verifies the bundle into a versioned
per-user `%LOCALAPPDATA%` cache, and launches both controller and model only
from that local cache. Resolve each user's local Draft path dynamically and
pass the running Excel workbook identity and HWND using the useful part of
REM's `RunFrozenPython` pattern. Keep all Stage 1/Stage 2 validation, audit,
rollback and model semantics; do not change REM or external model sources.

This avoids both repeated UNC execution and the observed SciPy failure from
executing the same bundle inside a OneDrive-synced tree. The previous isolated
332 MB bundle passed model preflight from LocalAppData and failed from the
OneDrive-synced sibling path despite byte-identical files. A central UNC share
is only a release-download source; no EXE is launched there. One isolated
staging build was separately approved; central promotion and Draft installation
remain unapproved.

## User Requirements

- Run from desktop Excel with Draft locally synced from SharePoint/OneDrive;
  Excel in a browser is not supported because the workflow requires VBA and
  desktop Excel COM.
- Colleagues need no Python installation, environment variables or interpreter
  configuration.
- Keep Draft in the user's locally synced SharePoint/OneDrive folder. On the
  first button click in each Excel session, compare the central release ID with
  the user's LocalAppData pointer and copy/verify only if it changed. Cache the
  selected bundle in a versioned LocalAppData folder; later clicks in that
  Excel session use the cached path without network access. Never execute
  binaries from UNC or the OneDrive sync tree.
- If a changed pointer or central package is unavailable/invalid, retain and
  run the last complete local cache with a visible Excel StatusBar warning. If
  there is no complete cache, stop with an actionable error.
- Use a maintainer-approved central share as a versioned copy source only; do
  not add a separate refresh script, macro or workbook button.
- Show immediate progress during a first-use or changed-release copy. Preserve
  older local versions for rollback.
- Prioritise speed and reliability over a hard package-size ceiling; avoid
  packaging a full Anaconda environment unnecessarily.
- Do not start a package build until the user separately approves it. Do not
  run a real financial batch without separate explicit authorisation.
- Preserve the user's no-warning and no-policy-change requirements. Do not
  bypass Windows security prompts or use `Unblock-File`; if the local package
  still warns, stop rollout and report the trust-policy blocker.

## REM Findings And Limits

- `modRunPython.CallRunPython` has active `RUNTYPE = "EXE"` and calls a
  hard-coded central frozen executable. The workbook can move because code
  location is independent of workbook location; this does not remove its
  network dependency.
- xlwings 0.19.5 `RunFrozenPython` changes to the executable's parent folder,
  uses `pushd` for a UNC executable, then invokes `ExecuteWindows`.
- `ExecuteWindows` passes the workbook name, Excel executable path and Excel
  HWND. It waits synchronously (`WaitOnReturn = True`), hides the console by
  default (`SHOW CONSOLE` defaults to False), and redirects stderr to
  `%APPDATA%\xlwings.log`. REM calculates before and after the call and records
  start/end timestamps; it does not provide a visible percentage indicator in
  the inspected VBA.
- The REM's `LOCAL` branch is commented out and requires a Python interpreter
  configuration. Its active frozen branch does not make the many fixed UNC and
  SharePoint workbook links portable. Reuse its Excel-context handshake, not its
  fixed code path or its claim of general SharePoint portability.
- The REM has 22 sheets, embedded VBA modules `modRunPython` and `xlwings`, and
  a Form Control shape on `I. Control` assigned to `CallRunPython`. The workbook
  was opened read-only with macros/events disabled; no REM macro was executed.

## Options

| Option | Speed and portability | Decision |
| --- | --- | --- |
| Fixed frozen executable on UNC, as REM uses | Central updates are simple and no Python setup is needed, but each launch depends on the network share and can incur network/security scanning. It does not adapt external data links. | Do not use for the Bonds Lite runtime; the reported central version was slower and prompted. |
| Full local sidecar beside Draft | Simple layout, but the tested compact model failed SciPy DLL loading when run from the OneDrive-synced folder. | Do not execute from this location unless separately fixed and retested. |
| Versioned per-user LocalAppData cache, downloaded from central release share | Keeps binaries local and outside OneDrive while allowing central release selection; no user Python setup or hand-configured paths. | **Selected.** First copy costs time/storage; show status and verify integrity. The share is a copy source only. |

The existing isolated bundle measured 332,228,371 bytes, while the old full
Anaconda bundle was about 1.98 GB. Treat those as observations, not a package
size promise; capture the size of any approved final build.

## Steps

1. **Establish the no-build local-runtime baseline.** The 332 MB staged
  artifact is already available in LocalAppData. Run only model `--preflight`
  and controller `--self-test` with Python settings removed and `PATH` limited
  to Windows. Repeat cold/warm timing from staging and from a test LocalAppData
  cache. Do not invoke the financial batch. The earlier OneDrive-sibling SciPy
  failure is established; do not spend time repeating it without a specific
  file-state hypothesis.
2. **Test package copying without a build.** Use a disposable temporary release
  tree with a pointer and sentinel files to test path validation, safe release
  IDs, staging-copy/atomic-rename, interrupted/incomplete copy cleanup, hash
  checks and cache reuse. Do not copy the existing 1.8 GB central bundle into
  a user's cache. No build or real release publication is included.
3. **Write launcher/cache tests first.** In `tests/test_bonds_lite_input.py`,
   require that the packaged launcher has no Python environment settings, uses
   a local release pointer for warm starts, reads the central release ID only
   on first use or explicit refresh, copies into a versioned LocalAppData
   directory, reuses a complete cache, rejects invalid/incomplete releases,
   and never invokes an UNC executable. Test source-mode fallback only when the
   process is not frozen. Keep UNC data paths valid independently of runtime
   deployment.
4. **Add the REM-style Excel attachment contract.** Adapt the current Python
   entrypoint to locate the exact Excel instance by HWND and the calling
   workbook by native name/full name, while accepting a separate local Draft
   path for filesystem operations. Do not match an HTTPS `FullName` against a
   local path. Preserve the active Excel process; never open a replacement
   Excel session for the caller. Add tests for two Excel HWNDs and ambiguous
   workbook names. Keep the existing asynchronous VBA launch so Excel remains
   serviceable; REM itself waits synchronously, which is not the selected UX.
5. **Replace only the Python launcher surface.** Add a packaged cache launcher,
  while retaining editable source mode for maintainers. In packaged mode,
  resolve Stage 2 to the cached `Model/BondsLiteModel.exe` and never read
  per-user Python settings. Add an early flushed controller message before
  importing the main workflow. Show Excel `StatusBar` progress during first
  release copy. Keep the existing console timing, validation, run guard,
  Stage 1 audit, curve protections, Stage 2 transaction and result checks.
6. **Repair/reconstitute reproducible build support.** The current tracked
   `tools/build_packaged_stages.py` references
   `src/bonds_lite_model_entry.py`, which is absent in this Git snapshot, and
   there are no pinned isolated-build requirements or venv preparation script
   in `tools/`. Add the minimum model entrypoint, pinned controller/model build
   requirements and isolated build preparation. Keep the external
   `run_bonds_batch.py`, `BondsModelHandler.py` and `pyBondsHMutils.py`
   unchanged. Use PyInstaller `--onedir`; exclude unrelated controller
   dependencies only with tests/evidence. Build output goes to a new staging
   path and refuses to overwrite an existing folder.
7. **Define release publication separately.** After a
  user-approved build, publish a versioned `Bonds Lite App` under the central
  release share and atomically update `current.txt` only after all package
  tests pass. Do not install or run the executable from the share. Keep prior
  versions for rollback. The workbook copies the selected release to
  `%LOCALAPPDATA%` on first use. Warm clicks read only the local release
  pointer. At the first button click after Excel starts, compare the central
  pointer with the user's local pointer; copy only when the release ID changes.
  Keep prior cache versions for rollback.
8. **Run acceptance gates.** Test the generated VBA on disposable workbooks,
   inspect the closed production Draft backup before installation, and verify
  its controls/audit names and VBA hash. Require under 5 seconds from warm
  button click to visible workflow output for both stages and no Open File
  warning on the target user's machine. Measure first-copy time separately;
  show an explicit status during it. Measure warm starts and Stage 1/Stage 2
  work separately from executable startup. Do not run a real model batch until
  the user grants separate permission.
9. **Update documentation after acceptance.** Revise the setup guide to state
   desktop Excel/local OneDrive prerequisites, copy/update procedure, path
   behavior, supported SharePoint access, rollback and the no-browser limit.
   Remove user Python setup steps from colleague instructions while retaining
   maintainer-only source/build steps.

## Test Strategy

1. **Fast per-step tests:** use targeted `pytest -k` filters for launcher,
   resolver, installer, entrypoint and builder cases. Do not run the entire
   roughly 700-case module after each edit. Run the full non-Excel suite once
   after the focused gates, capturing a final exit code and duration; the last
   full-suite terminal attempt did not provide a reliable summary.
2. **OneDrive path tests:** cover local path, HTTPS SharePoint URL, multiple
   `SyncEngines` mappings, missing/offline file, mismatched workbook identity,
   and a path containing spaces/non-ASCII characters. Use disposable files and
   mocked registry mappings in ordinary unit tests.
3. **UNC boundary tests:** normal packaged execution must use the local cache
  and must not read the central share. Only first-use and explicit refresh
  paths may access it; no path may launch an executable from UNC. Keep
  absolute UNC input/template/result behavior supported where the current
  workflow permits it.
4. **Excel integration:** explicitly opt in and use disposable workbook copies
   to compile the generated VBA, resolve the local caller and sibling
   controller, and confirm the exact application/HWND/workbook attachment. The
   user must manually observe any Windows security prompt on the target machine;
   never dismiss or bypass it programmatically.
5. **Packaged smoke tests:** controller `--self-test`; model `--preflight` and
  `--self-test` from staging and the LocalAppData cache with Python settings
  removed. `--self-test` must not run `run_bonds_batch.py`. Compare package
  hashes before/after copy and hashes of the unchanged external model modules.
6. **Stage 1/2 regression:** retain controlled runner tests for paired inputs,
   numeric JRC curves, audit handoff, exact one-batch invocation, required
   outputs, rollback and source immutability. Do not use the real model in
   ordinary tests.
7. **Performance gates:** capture button click to first console line, first
   useful progress, and model preflight, cold and warm. Goal: first useful
   output within 5 seconds. Report actual package bytes and startup times; do not
   infer a speedup from build time or total financial-run duration.

## Affected Files

- `src/bonds_lite_input.py` — packaged LocalAppData cache resolution, release
  pointer validation/copying, local Draft path override, exact Excel context
  matching and frozen model configuration. Preserve Stage 1/2 transactions.
- `src/bonds_lite_model_entry.py` — new bounded preflight/self-test and
  folder-batch entrypoint for Python 3.11 frozen model; ensure normal preflight
  does not JIT-compile or start a batch.
- `tools/controller_entry.py` — early visible startup message.
- `tools/build_packaged_stages.py` — versioned release output/manifest/source
  hashes and self-tests; central share is a copy source, never an execution
  path.
- `tools/prepare_packaged_envs.py` and
  `tools/requirements-controller-build.txt` /
  `tools/requirements-model-build.txt` — isolated pinned build environments.
- `tests/test_bonds_lite_input.py` — launcher/path, builder, packaging and
  unchanged Stage 1/2 contract tests.
- `references/bonds-lite-setup.md` — colleague-sidecar copy and SharePoint
  instructions after the target path passes.
- `references/decisions.md` — record approved local deployment and A/B gate.
- `Draft Input spreadsheet for SST.xlsm` — back up before installing the new
  launcher; do not alter its input controls, panels or other formulas.
- `REM_v1.10.xlsb` and external model source modules — inspect only; never edit.

## Decisions

- User selected a versioned central package source with an automatic per-user
  LocalAppData execution cache. The central pointer is checked once on the
  first click in each Excel session; later clicks use the module-cached local
  bundle without network I/O. Never launch binaries from UNC or the OneDrive
  sync tree.
- Draft remains in the user's locally synced SharePoint/OneDrive folder and
  runs in desktop Excel. Excel for the web is out of scope.
- The launcher passes native `ThisWorkbook.FullName` and Excel HWND for exact
  attachment, plus a separately resolved local Draft path for file operations.
- No separate refresh script, macro or button is required. A new Excel session
  observes a changed central release pointer and copies that version locally.
- Speed and reliability take priority over a hard package-size ceiling; avoid
  full Anaconda packaging and measure the final approved bundle.
- Do not build, install into Draft or run a real financial batch until the
  corresponding explicit user approval.

## Risks And Unknowns

- The 332 MB compact package passed SciPy preflight from LocalAppData but failed
  from the OneDrive-synced sibling with `_matfuncs_expm: Invalid access to
  memory location`. All 869 model-runtime files had identical hashes; path
  lengths were below 260 characters; tested files had no `Zone.Identifier`.
  The LocalAppData copy is the selected execution path, but the updated cache
  VBA has not yet been validated in Excel.
- The central UNC share must be reachable at first use of a release in each
  Excel session. Once the local bundle is selected, subsequent clicks should
  use only that local cache. If both central share and a complete local cache
  are unavailable, the launcher must fail clearly.
- The REM EXE's fixed central UNC path proves that its workbook can be moved
  independently of code, not that code or external links are portable. It has
  many fixed UNC/SharePoint dependencies and does not resolve them per user.
- SharePoint URL-to-OneDrive mapping must be verified for different users and
  library mount names. Browser Excel cannot run this VBA/COM workflow.
- The tests used mocked SyncEngines/copy-source files. A multi-user mapping
  check and real first-click cache copy on the target machine are still needed.
- A local path does not guarantee Windows will suppress an Open File warning.
  Do not bypass the warning; if it appears without a policy-approved trust
  solution, stop rollout.
- Release `20261030-cache-v1` was published after Robocopy reported 984 files,
  zero mismatches, and zero failures; total source/copy bytes matched and both
  executable SHA-256 values matched the manifest. `current.txt` was verified on
  the new release. The packaged launcher was installed into Draft with a
  timestamped backup at `references/backups/Draft Input spreadsheet for
  SST_20261030_135007_737437.xlsm`. Target-user warning/startup behavior remains
  unverified; no real financial batch ran.

## Addendum: Implementation Progress (30 October 2026)

- Added central manifest/hash validation, atomic per-user cache staging,
  frozen sibling-model configuration, packaged preflight, and a packaged model
  entrypoint. The controller validates the cached bundle before normal frozen
  Stage 1/2 dispatch.
- Added packaged VBA for SharePoint URL-to-local path mapping, central pointer
  selection once per Excel session, LocalAppData caching, and separate native
  workbook identity/local filesystem path arguments. Added an explicit
  `--packaged` installer option; source-mode launcher remains the default.
- Added packaged Stage 2 execution with a clean environment and direct cached
  model EXE invocation. Existing source-mode model command remains unchanged.
- Builder publishes versioned copy-source releases and promotes only with an
  explicit flag. Pinned controller/model build requirements and isolated venv
  preparation are present. Approved staging release `20261030-cache-v1` was
  built under LocalAppData: 332,229,903 bytes / 984 files. Both model numerical
  self-test and controller self-test passed; clean-environment preflight passed
  with Python-related PATH entries removed. The release was copied to the central
  share with zero Robocopy mismatches/failures; file totals and both executable
  hashes matched. The central pointer was verified on this release. The packaged
  launcher was installed in Draft and a timestamped backup was confirmed.
- Added focused cache, launcher, SharePoint path, installer, model entrypoint,
  Stage 2 command and bundle-publication tests. The deployment-only suite passed
  16 tests in 0.24 seconds. The broader focused suite passed 148 tests with one
  intermittent existing rollback-location assertion; that assertion passed
  alone. The fast Stage 1/2 safety subset passed 133 tests in 6.95 seconds.
- The two opt-in Excel cache tests passed (2 passed, 697 deselected, 5.94
  seconds): local release copy/resolution and fallback to a previous complete
  cache when the central release is incomplete. Excel emitted the known
  `0x800706ba` fixture teardown diagnostic after the passing assertions.
- Static diagnostics report no errors for the modified Python and tests. The
  latest non-Excel run had 648 passed, 8 skipped and 73 deselected, with the
  same two previously observed failures: the frozen-main test-signature
  mismatch and rollback-location reporting assertion. This is not a full pass.
- Per Decision 30, corrected Stage 1 to preserve calculation mode while opening
  workbooks, scope manual calculation to bulk edits, avoid unchanged control
  writes and use one F9-style calculation. All 61 focused tests passed. A live
  disposable-copy check of the production settings/open/edit/refresh sequence
  completed in 2.11 seconds and validated all 1,201 SONIA rows; it did not run
  Stage 1/2 or a financial batch. The user also confirmed that forced native
  Ctrl+Alt+F9 remains busy and selected normal F9 semantics instead.
- The user approved one replacement build and verified rollout for this fix.
  Release `20261030-cache-v2` was built and verified: 984 files, 332,230,035
  bytes, matching executable hashes, unchanged model sources and all four
  corrected embedded controller functions matching the tested source. Model
  and controller self-tests passed both during build and from the local cache
  with the clean packaged environment. The central copy completed with zero
  mismatches/failures; a subsequent comparison returned zero differences and
  relative file sizes matched. Central and local pointers were verified on v2.
  Draft's packaged launcher was installed and read back with a backup at
  `references/backups/Draft Input spreadsheet for SST_20261030_152208_786602.xlsm`.
  The prior v1 release remains available. The live curve-path check used the
  normal automatic calculation mode; changed valuation dates can still require
  fresh JRC retrieval, whose runtime has not been resolved by this correction.
- The user confirmed both stages succeeded with v2 in run `20261030_153601`.
  The batch summary reports 101.8 seconds for Bonds and 111.3 seconds for MAP;
  the later cleanup error did not cause those calculation times. The retained
  rollback directory contained only empty read-only directories (attributes
  `0x31`). A disposable test reproduced WinError 5 and confirmed that clearing
  only the read-only bit permits deletion. Added a deletion-error handler that
  retries that case once, keeps genuine permission failures visible without a
  success-run traceback, and leaves failed-rollback recovery protection intact.
  All 25 focused cleanup/transaction tests passed. The approved empty leftover
  was rechecked and removed without touching the successful results. One v3
  replacement build/rollout was approved and completed: 984 files, 332,230,482
  bytes, zero copy mismatches/failures/extras and matching central/local file
  sizes and executable hashes. All six relevant embedded controller functions
  match tested source; the extracted compiled cleanup handler passed a real
  read-only deletion check. Model-source hashes are unchanged and both local
  clean-environment self-tests passed. Central and local pointers were verified
  on `20261030-cache-v3`. The VBA launcher is byte-identical to v2, so no Draft
  edit was necessary; reopening Draft selects the new cached runtime. Existing
  v1/v2 releases were preserved. No financial model batch was rerun for this fix.
- No REM edits, security-policy changes or agent-run real financial batch were
  performed. The user's successful run is not an independent financial-result
  validation; target-user security-warning acceptance remains a separate gate.