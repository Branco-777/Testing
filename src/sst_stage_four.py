"""Run MACalc Batch on the exact workbook published by Stage 3."""

from __future__ import annotations

import getpass
import logging
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any

import bonds_lite_input as shared
import sst_stage_three

logger = shared.logger.getChild("sst_stage_four")

STAGE_FOUR_FIELD_NAMES = {
    "last_run": "SST_Stage4_LastRun",
    "start": "SST_Stage4_Start",
    "finish": "SST_Stage4_Finish",
    "elapsed": "SST_Stage4_Elapsed",
    "user": "SST_Stage4_User",
    "status": "SST_Stage4_Status",
    "input_run": "SST_Stage4_InputRun",
    "last_output": "SST_Stage4_LastSuccessfulOutput",
}
STAGE_FOUR_PANEL_NAME = "SST_Stage4_Panel"
STAGE_FOUR_BUTTON_NAME = "btnRunMACalc"
STAGE_FOUR_BUTTON_MACRO = "RunMACalc"
STAGE_FOUR_LABELS = {
    "last_run": "Last Run",
    "start": "Start:",
    "finish": "Finish:",
    "elapsed": "Time",
    "user": "User",
    "status": "Status",
    "input_run": "Input Run",
    "last_output": "Last Successful Output",
}
STAGE_FOUR_OFFSETS = {
    "last_run": 1,
    "start": 2,
    "finish": 3,
    "elapsed": 4,
    "user": 5,
    "status": 6,
    "input_run": 7,
    "last_output": 9,
}
MACALC_ENTRY = Path(__file__).with_name("macalc_batch_entry.py")
STAGE_FOUR_BACKUP_FOLDER = "Stage 4 backup"
MACALC_NAME_PATTERN = re.compile(r"MACalc input_(.+)_([0-9a-f]{16})\.xlsb")


class StageFourError(shared.BondsLiteError):
    """An actionable Stage 4 handoff, model or audit failure."""


@dataclass(frozen=True)
class StageFourInput:
    """The validated Stage 3 workbook and the MACalc downgrade setting."""

    run_folder: Path
    workbook: Path
    is_downgrade_run: bool


@dataclass(frozen=True)
class MacalcConfig:
    """The separate MACalc Python environment and source folder."""

    python_executable: Path
    model_dir: Path
    packaged: bool = False


def load_macalc_config(caller_path: Path) -> MacalcConfig:
    """Resolve the per-user model Python and MACalc Batch source folder."""
    if getattr(sys, "frozen", False):
        controller = Path(sys.executable).resolve()
        executable = controller.parent.parent / "MACalc" / "MACalcModel.exe"
        if not executable.is_file():
            raise StageFourError(
                f"Packaged MACalc model missing from the local cache: {executable}"
            )
        return MacalcConfig(executable, executable.parent, packaged=True)

    try:
        python = shared.resolve_model_python()
    except shared.BondsLiteError as error:
        raise StageFourError(str(error)) from error
    configured_dir = os.environ.get("SST_MACALC_DIR", "").strip()
    if configured_dir:
        model_dir = Path(os.path.expandvars(configured_dir)).expanduser().resolve()
    else:
        model_dir = (
            Path(caller_path).resolve().parent
            / ".."
            / ".."
            / "00 Ad Hoc"
            / "MACalc_Batch"
        ).resolve()
    if not model_dir.is_dir() or not (model_dir / "MACalcrunmanager.py").is_file():
        raise StageFourError(
            f"MACalc Batch directory has no MACalcrunmanager.py: {model_dir}"
        )
    return MacalcConfig(python, model_dir)


def _stage_three_fields(caller: Any) -> dict[str, object]:
    try:
        return sst_stage_three._read_stage_three_fields(caller)
    except Exception as error:
        raise StageFourError(
            f"Stage 3 audit is unavailable; rerun Stage 3: {error}"
        ) from error


def _resolve_stage_three_workbook(
    caller_path: Path, value: object, run_folder: Path
) -> Path:
    if not isinstance(value, str) or not value.strip():
        raise StageFourError("Stage 3 has no successful output; rerun Stage 3")
    relative = Path(value)
    if relative.is_absolute() or relative.drive or ".." in relative.parts:
        raise StageFourError("Stage 3 output is not a portable path; rerun Stage 3")

    candidate = caller_path.parent / relative
    for path in (candidate, *candidate.parents):
        if path.is_symlink():
            raise StageFourError(f"Stage 3 output must not use a symbolic link: {path}")
        if path == caller_path.parent:
            break
    try:
        workbook = candidate.resolve(strict=True)
    except OSError as error:
        raise StageFourError(
            f"Stage 3 output is missing; rerun Stage 3: {candidate}"
        ) from error

    suffix = shared.output_root_suffix(run_folder.parent.name)
    match = MACALC_NAME_PATTERN.fullmatch(candidate.name)
    if (
        suffix is None
        or match is None
        or match.group(1) != suffix
        or candidate.parent != run_folder
        or workbook.parent != run_folder
        or workbook != candidate
        or not workbook.is_file()
    ):
        raise StageFourError(
            "Stage 3 output is outside the current stress folder or has an invalid name; "
            "rerun Stage 3"
        )
    if (run_folder / f"~${workbook.name}").exists():
        raise StageFourError(
            f"Stage 3 workbook is locked or open in Excel; close it before Stage 4: "
            f"{workbook.name}"
        )
    return workbook


def resolve_stage_four_input(caller: Any) -> StageFourInput:
    """Require a current Stage 2/3 handoff and read the downgrade flag."""
    caller_path = shared.resolve_caller_path(caller)
    try:
        run_folder = shared.resolve_stage_two_results(caller)
    except shared.BondsLiteError as error:
        raise StageFourError(
            f"Stage 2 has not succeeded for the current run; rerun Stage 2: {error}"
        ) from error

    stage_three = _stage_three_fields(caller)
    portable_run = shared.portable_output_path(caller_path, run_folder)
    if stage_three["status"] != "Succeeded" or stage_three["input_run"] != portable_run:
        raise StageFourError("Stage 3 is stale or unsuccessful; rerun Stage 3")
    workbook = _resolve_stage_three_workbook(
        caller_path, stage_three["last_output"], run_folder
    )

    try:
        downgrade = shared.read_stress_flags(caller)[shared.CREDIT_DG_FLAG]
    except shared.BondsLiteError as error:
        raise StageFourError(str(error)) from error
    return StageFourInput(run_folder, workbook, downgrade)


def preflight_macalc_environment(config: MacalcConfig) -> None:
    """Check that the configured Python can import MACalc Batch before mutation."""
    if config.packaged:
        if not config.python_executable.is_file():
            raise StageFourError(
                f"Packaged MACalc model is missing: {config.python_executable}"
            )
        return
    probe = (
        "import sys; assert sys.version_info[:2] == (3, 11), "
        "'Model Python 3.11 is required'; import MACalcrunmanager; "
        "print('MACALC_PREFLIGHT_OK')"
    )
    environment = os.environ.copy()
    environment["PYTHONDONTWRITEBYTECODE"] = "1"
    try:
        result = subprocess.run(
            [str(config.python_executable), "-B", "-c", probe],
            cwd=str(config.model_dir),
            capture_output=True,
            text=True,
            errors="replace",
            timeout=30,
            shell=False,
            env=environment,
        )
    except (OSError, subprocess.SubprocessError) as error:
        detail = getattr(error, "stderr", "") or str(error)
        raise StageFourError(
            f"MACalc environment preflight failed: {detail}"
        ) from error
    if result.returncode:
        raise StageFourError(
            f"MACalc environment preflight failed: {result.stderr.strip()}"
        )
    logger.info("%s", result.stdout.strip())


def backup_stage_four_input(workbook: Path) -> Path:
    """Atomically replace the single pre-run backup outside the stress folder."""
    workbook = Path(workbook).resolve(strict=True)
    output_root = workbook.parent.parent
    backup_folder = output_root / STAGE_FOUR_BACKUP_FOLDER
    if backup_folder.is_symlink():
        raise StageFourError(
            f"Stage 4 backup folder must not be a symbolic link: {backup_folder}"
        )
    backup_folder.mkdir(parents=True, exist_ok=True)
    destination = backup_folder / workbook.name
    temporary: Path | None = None
    try:
        with tempfile.NamedTemporaryFile(
            prefix=f".{workbook.name}.", suffix=".tmp", dir=backup_folder, delete=False
        ) as stream:
            temporary = Path(stream.name)
        shutil.copy2(workbook, temporary)
        if shared._file_hash(temporary) != shared._file_hash(workbook):
            raise StageFourError(f"Stage 4 backup verification failed: {temporary}")
        os.replace(temporary, destination)
    except Exception as error:
        if temporary is not None:
            try:
                temporary.unlink(missing_ok=True)
            except OSError:
                logger.exception(
                    "Could not remove incomplete Stage 4 backup: %s", temporary
                )
        if isinstance(error, StageFourError):
            raise
        raise StageFourError(
            f"Cannot create or replace Stage 4 backup: {error}"
        ) from error
    return destination


def macalc_command(config: MacalcConfig, inputs: StageFourInput) -> list[str]:
    """Build the child command for the source or packaged model runtime."""
    command = [str(config.python_executable)]
    if not config.packaged:
        command.extend(["-B", str(MACALC_ENTRY)])
    command.extend(
        [
            "--model-dir",
            str(config.model_dir),
            "--root",
            str(inputs.run_folder),
            "--expected",
            str(inputs.workbook),
            "--downgrade-run" if inputs.is_downgrade_run else "--no-downgrade-run",
        ]
    )
    return command


def run_macalc_batch(config: MacalcConfig, inputs: StageFourInput) -> Path:
    """Stream MACalc output and return its verified unoptimised-assets CSV."""
    environment = os.environ.copy()
    environment["PYTHONDONTWRITEBYTECODE"] = "1"
    command = macalc_command(config, inputs)
    logger.info(
        "MACalc options: binary optimisation OFF; optimisation pass OFF; "
        "downgrade %s (inactive while optimisation is OFF); CoD floor at zero",
        "ON" if inputs.is_downgrade_run else "OFF",
    )
    logger.info("MACalc child command: %s", subprocess.list2cmdline(command))
    child = subprocess.Popen(
        command,
        cwd=str(config.model_dir),
        shell=False,
        env=environment,
    )
    try:
        result = child.wait()
    except KeyboardInterrupt:
        if child.poll() is None:
            child.terminate()
        try:
            child.wait(timeout=5)
        except TypeError:
            pass
        except subprocess.TimeoutExpired:
            child.kill()
            child.wait()
        except KeyboardInterrupt:
            child.kill()
            child.wait()
        raise
    if result:
        raise StageFourError(f"MACalc Batch exited with code {result}")

    output = inputs.workbook.with_name(f"{inputs.workbook.stem}_unoptimised_assets.csv")
    if not output.is_file() or output.stat().st_size == 0:
        raise StageFourError(
            f"MACalc Batch did not produce a non-empty results CSV: {output}"
        )
    return output


def _stage_four_cells(caller: Any) -> dict[str, Any]:
    cells = {}
    for field, name in STAGE_FOUR_FIELD_NAMES.items():
        region = caller.names[name].refers_to_range
        merged = region.api.Cells(1, 1).MergeArea
        cells[field] = region.sheet.range((merged.Row, merged.Column))
    return cells


def require_stage_four_panel(caller: Any) -> None:
    """Validate Stage 4's named audit panel before reading or writing it."""
    try:
        if caller.api.ReadOnly:
            raise StageFourError(
                "Draft is read-only; its Stage 4 audit cannot be saved"
            )
        sheet = caller.sheets["I. Control"]
        if sheet.api.ProtectContents:
            raise StageFourError(
                "I. Control is protected; its Stage 4 audit cannot be updated"
            )
        panel = caller.names[STAGE_FOUR_PANEL_NAME].refers_to_range
        if panel.sheet.name != "I. Control" or panel.api.Areas.Count != 1:
            raise StageFourError("Stage 4 audit panel name is invalid")
        bounds = (
            panel.row,
            panel.column,
            panel.row + panel.rows.count - 1,
            panel.column + panel.columns.count - 1,
        )
        occupied: list[tuple[int, int, int, int]] = []
        for field, name in STAGE_FOUR_FIELD_NAMES.items():
            region = caller.names[name].refers_to_range
            if region.sheet.name != "I. Control":
                raise StageFourError("Stage 4 audit fields must remain on I. Control")
            merge = region.api.Cells(1, 1).MergeArea
            field_bounds = (
                merge.Row,
                merge.Column,
                merge.Row + merge.Rows.Count - 1,
                merge.Column + merge.Columns.Count - 1,
            )
            if not (
                bounds[0] <= field_bounds[0] <= field_bounds[2] <= bounds[2]
                and bounds[1] <= field_bounds[1] <= field_bounds[3] <= bounds[3]
            ):
                raise StageFourError(
                    f"Stage 4 audit field is outside its panel: {name}"
                )
            if any(
                field_bounds[0] <= old[2]
                and field_bounds[2] >= old[0]
                and field_bounds[1] <= old[3]
                and field_bounds[3] >= old[1]
                for old in occupied
            ):
                raise StageFourError(f"Stage 4 audit fields overlap: {field}")
            occupied.append(field_bounds)
    except StageFourError:
        raise
    except Exception as error:
        raise StageFourError(
            "Stage 4 audit panel names are missing or invalid; install the Stage 4 controls"
        ) from error


def _read_stage_four_fields(caller: Any) -> dict[str, object]:
    require_stage_four_panel(caller)
    return {field: cell.value for field, cell in _stage_four_cells(caller).items()}


def write_stage_four_audit(caller: Any, fields: dict[str, object]) -> None:
    """Write only the owned Stage 4 audit values and save the Draft."""
    require_stage_four_panel(caller)
    if set(fields) - STAGE_FOUR_FIELD_NAMES.keys():
        raise StageFourError("Unknown Stage 4 audit field")
    cells = _stage_four_cells(caller)
    events = caller.app.api.EnableEvents
    try:
        caller.app.api.EnableEvents = False
        for field, value in fields.items():
            cell = cells[field]
            if field in ("input_run", "last_output"):
                if cell.api.Hyperlinks.Count:
                    cell.api.Hyperlinks.Delete()
                if value:
                    cell.add_hyperlink(str(value), text_to_display=str(value))
                else:
                    cell.value = None
            else:
                cell.value = value
        caller.save()
        if not caller.api.Saved:
            raise StageFourError("Draft's Stage 4 audit was not saved")
    finally:
        caller.app.api.EnableEvents = events


def _portable_file_path(caller_path: Path, output: Path) -> str:
    try:
        return (
            output.resolve(strict=True)
            .relative_to(caller_path.parent.resolve())
            .as_posix()
        )
    except (OSError, ValueError) as error:
        raise StageFourError(
            f"Stage 4 result is outside Draft's output tree: {output}"
        ) from error


def run_stage_four(caller: Any) -> Path:
    """Run MACalc on the latest verified Stage 3 output and record its audit."""
    caller_path = shared.resolve_caller_path(caller)
    with shared.workbook_run_guard(caller_path):
        require_stage_four_panel(caller)
        inputs = resolve_stage_four_input(caller)
        previous = _read_stage_four_fields(caller)["last_output"]
        input_path = _portable_file_path(caller_path, inputs.workbook)
        started = datetime.now()
        clock = time.perf_counter()
        write_stage_four_audit(
            caller,
            {
                "last_run": started.replace(hour=0, minute=0, second=0, microsecond=0),
                "start": started,
                "finish": None,
                "elapsed": None,
                "user": getpass.getuser(),
                "status": "Running",
                "input_run": input_path,
                "last_output": previous,
            },
        )
        try:
            config = load_macalc_config(caller_path)
            preflight_macalc_environment(config)
            backup = backup_stage_four_input(inputs.workbook)
            logger.info("Stage 4 input: %s", inputs.workbook)
            logger.info("Stage 4 backup: %s", backup)
            output = run_macalc_batch(config, inputs)
            portable_output = _portable_file_path(caller_path, output)
            write_stage_four_audit(
                caller,
                {
                    "finish": datetime.now(),
                    "elapsed": (time.perf_counter() - clock) / 86400,
                    "status": "Succeeded",
                    "last_output": portable_output,
                },
            )
            logger.info("Stage 4 results: %s", output)
            return output
        except BaseException as error:
            try:
                write_stage_four_audit(
                    caller,
                    {
                        "finish": datetime.now(),
                        "elapsed": (time.perf_counter() - clock) / 86400,
                        "status": (
                            "Interrupted"
                            if isinstance(error, KeyboardInterrupt)
                            else "Failed"
                        ),
                        "last_output": previous,
                    },
                )
            except Exception as audit_error:
                logger.error(
                    "Stage 4 failed and its audit could not be saved: %s", audit_error
                )
            raise


def _install_stage_four_panel(book: Any) -> None:
    """Install Stage 4's named audit beneath the existing Stage 3 panel."""
    try:
        sst_stage_three.require_stage_three_panel(book)
    except Exception as error:
        raise StageFourError(
            f"Install the Stage 3 controls before Stage 4: {error}"
        ) from error
    expected_names = {STAGE_FOUR_PANEL_NAME, *STAGE_FOUR_FIELD_NAMES.values()}
    existing = {name.name.split("!")[-1] for name in book.names}
    if existing.intersection(expected_names):
        if not expected_names.issubset(existing):
            raise StageFourError(
                "Stage 4 audit names are incomplete; installation was not changed"
            )
        require_stage_four_panel(book)
        return

    sheet = book.sheets["I. Control"]
    previous = book.names[sst_stage_three.STAGE_THREE_PANEL_NAME].refers_to_range
    row = previous.row + previous.rows.count + 2
    column = previous.column
    area = sheet.range((row, column), (row + 10, column + 7))
    if (
        book.app.api.WorksheetFunction.CountA(area.api)
        or area.api.MergeCells is not False
    ):
        raise StageFourError(
            "Proposed Stage 4 audit area is occupied; nothing was overwritten"
        )
    for shape in sheet.api.Shapes:
        if (
            shape.Left < area.left + area.width
            and shape.Left + shape.Width > area.left
            and shape.Top < area.top + area.height
            and shape.Top + shape.Height > area.top
        ):
            raise StageFourError(
                "Proposed Stage 4 audit area overlaps a shape or button"
            )

    events = book.app.api.EnableEvents
    added: list[str] = []
    try:
        book.app.api.EnableEvents = False
        area.color, area.font.name, area.font.size = (230, 245, 230), "Calibri", 10
        area.api.VerticalAlignment = -4108
        heading = sheet.range((row, column), (row, column + 7))
        heading.merge()
        heading.value, heading.color = "Stage 4 - Run MACalc", (0, 153, 0)
        heading.font.color, heading.font.bold = (255, 255, 255), True
        book.names.add(STAGE_FOUR_PANEL_NAME, f"='I. Control'!{area.get_address()}")
        added.append(STAGE_FOUR_PANEL_NAME)
        for field, offset in STAGE_FOUR_OFFSETS.items():
            end = row + offset + (1 if field in ("input_run", "last_output") else 0)
            label = sheet.range((row + offset, column), (end, column + 3))
            value = sheet.range((row + offset, column + 4), (end, column + 7))
            label.merge()
            value.merge()
            label.value, label.color, label.font.color = (
                STAGE_FOUR_LABELS[field],
                (0, 153, 0),
                (255, 255, 255),
            )
            label.api.WrapText = value.api.WrapText = True
            value.font.color, value.api.HorizontalAlignment = (0, 0, 0), -4131
            value.number_format = (
                "dd/mm/yyyy"
                if field == "last_run"
                else (
                    "hh:mm:ss"
                    if field in ("start", "finish")
                    else "[h]:mm:ss" if field == "elapsed" else "General"
                )
            )
            if field == "status":
                value.value = "Not Run"
            name = STAGE_FOUR_FIELD_NAMES[field]
            book.names.add(name, f"='I. Control'!{value[0, 0].get_address()}")
            added.append(name)
    except Exception:
        for name in reversed(added):
            book.names[name].delete()
        raise
    finally:
        book.app.api.EnableEvents = events
    require_stage_four_panel(book)


def _stage_four_button_row(panel: Any, button_left: float) -> int:
    """Place the button beside the panel when it fits, otherwise below it."""
    if button_left + 205 <= panel.left:
        return panel.row
    return panel.row + panel.rows.count + 2


def install_stage_four_controls(caller: Path) -> Path:
    """Back up and install Stage 4 controls into a closed Draft workbook."""
    import xlwings as xw

    caller = Path(caller).resolve(strict=True)
    for application in xw.apps:
        for book in application.books:
            if book.name.casefold() == caller.name.casefold():
                raise StageFourError(
                    f"Close {caller.name} in Excel before installation"
                )
    backup_folder = caller.parent / "references" / "backups"
    backup_folder.mkdir(parents=True, exist_ok=True)
    backup = backup_folder / f"{caller.stem}_{datetime.now():%Y%m%d_%H%M%S_%f}.xlsm"
    shutil.copy2(caller, backup)
    application = xw.App(visible=False, add_book=True)
    try:
        application.api.AutomationSecurity = 3
        application.api.EnableEvents = False
        application.api.Calculation = -4135
        application.display_alerts = False
        book = application.books.open(str(caller), update_links=False)
        try:
            if book.api.ReadOnly:
                raise StageFourError("Draft workbook is read-only or locked")
            _install_stage_four_panel(book)
            sheet = book.sheets["I. Control"]
            panel = book.names[STAGE_FOUR_PANEL_NAME].refers_to_range
            row = _stage_four_button_row(panel, sheet.range((panel.row, 4)).left)
            anchor = sheet.range((row, 4), (row + 2, 6))
            _install_button(book, anchor)
            components = book.api.VBProject.VBComponents
            module = components.Item("modBondsLiteLauncher")
            code = module.CodeModule
            text = code.Lines(1, code.CountOfLines) if code.CountOfLines else ""
            if "Public Sub RunMACalc()" not in text:
                code.InsertLines(
                    code.CountOfLines + 1,
                    "\nPublic Sub RunMACalc()\n    LaunchBondsLite 4\nEnd Sub\n",
                )
            book.save()
        finally:
            book.close()
    finally:
        application.api.DisplayAlerts = False
        application.api.Quit()
    return backup


def _install_button(book: Any, anchor: Any) -> None:
    """Create or validate the Stage 4 Forms button without covering workbook cells."""
    sheet = anchor.sheet
    names = [shape.Name for shape in sheet.api.Shapes]
    if STAGE_FOUR_BUTTON_NAME in names:
        button = sheet.api.Buttons().Item(STAGE_FOUR_BUTTON_NAME)
    else:
        width = max(anchor.width, 205)
        end_column = anchor.column
        while sheet.range((anchor.row, end_column)).left < anchor.left + width:
            end_column += 1
        covered = sheet.range(
            (anchor.row, anchor.column),
            (anchor.row + anchor.rows.count - 1, end_column - 1),
        )
        if (
            book.app.api.WorksheetFunction.CountA(covered.api)
            or covered.api.MergeCells is not False
        ):
            raise StageFourError(
                "Proposed Run MACalc button covers occupied cells or merges"
            )
        for shape in sheet.api.Shapes:
            if (
                shape.Left < anchor.left + width
                and shape.Left + shape.Width > anchor.left
                and shape.Top < anchor.top + anchor.height
                and shape.Top + shape.Height > anchor.top
            ):
                raise StageFourError(
                    "Proposed Run MACalc button overlaps an existing shape"
                )
        button = sheet.api.Buttons().Add(anchor.left, anchor.top, width, anchor.height)
        button.Name = STAGE_FOUR_BUTTON_NAME
    button.Caption = "Run MACalc"
    button.OnAction = (
        "'" + book.name.replace("'", "''") + "'!" + STAGE_FOUR_BUTTON_MACRO
    )
