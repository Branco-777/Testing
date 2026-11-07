"""Run LTM Flash against the successful Property Stage 1 workbook."""

from __future__ import annotations

import getpass
import logging
import os
import re
import shutil
import tempfile
import time
from contextlib import contextmanager
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from numbers import Real
from pathlib import Path
from typing import Any, Iterator
from urllib.parse import unquote

import bonds_lite_input as shared
import sst_property_stage_one as stage_one

logger = shared.logger.getChild("sst_property_stage_two")

STAGE_TWO_PANEL_NAME = "SST_PropStage2_Panel"
STAGE_TWO_BUTTON_NAME = "btnRunLTMFlash"
STAGE_TWO_BUTTON_MACRO = "RunLTMFlash"
STAGE_TWO_FIELD_NAMES = {
    "last_run": "SST_PropStage2_LastRun",
    "start": "SST_PropStage2_Start",
    "finish": "SST_PropStage2_Finish",
    "elapsed": "SST_PropStage2_Elapsed",
    "user": "SST_PropStage2_User",
    "status": "SST_PropStage2_Status",
    "input_workbook": "SST_PropStage2_InputWorkbook",
    "last_successful_run": "SST_PropStage2_LastSuccessfulRun",
}
STAGE_TWO_LABELS = {
    "last_run": "Last Run",
    "start": "Start:",
    "finish": "Finish:",
    "elapsed": "Time",
    "user": "User",
    "status": "Status",
    "input_workbook": "Input Workbook",
    "last_successful_run": "Last Successful Run",
}
STAGE_TWO_OFFSETS = {
    "last_run": 1,
    "start": 2,
    "finish": 3,
    "elapsed": 4,
    "user": 5,
    "status": 6,
    "input_workbook": 7,
    "last_successful_run": 9,
}
LTM_RUNNER_PATH = Path(
    r"\\server090\internal model$\7 ALM\2. Model Management\Model Repository"
    r"\LTMflashArunmanager_v3.2a(d9bb4fbe).exe"
)
LTM_LOCAL_RUN_NAME = "prm_LocalRun"
LTM_DATA_SOURCE_NAME = "prm_DataSource"
LTM_RUN_START_NAME = "RunStartTime"
LTM_RUN_END_NAME = "RunEndTime"
LTM_SCENARIO_COUNT_NAME = "prm_NoOfScens"
LTM_PROGRESS_FIRST_NAME = "ProgressFirstCell"
LTM_OUTPUT_SHEETS = ("O. CFOutput", "O. PVOutput", "O. ConcRisk")
XL_UP = -4162
XL_AUTOMATION_SECURITY_LOW = 1


class PropertyStageTwoError(stage_one.PropertyStageOneError):
    """An actionable Property Stage 2 input, Excel run or audit failure."""


@dataclass(frozen=True)
class PropertyStageTwoInputs:
    """Validated Stage 1 handoff and caller location for an in-place LTM run."""

    caller_path: Path
    workbook_path: Path


def _property_stage_two_field_cells(caller: Any) -> dict[str, Any]:
    """Resolve Stage 2 audit fields to the first cell of their merged ranges."""
    cells: dict[str, Any] = {}
    for field, name in STAGE_TWO_FIELD_NAMES.items():
        region = stage_one._audit_region(caller, name)
        merged = region.api.Cells(1, 1).MergeArea
        cells[field] = region.sheet.range((merged.Row, merged.Column))
    return cells


def require_property_stage_two_panel(caller: Any) -> None:
    """Validate Stage 2's named panel and ensure it does not overlap Stage 1."""
    try:
        shared._require_named_stage_panel(
            caller, STAGE_TWO_PANEL_NAME, STAGE_TWO_FIELD_NAMES
        )
        stage_one.require_property_stage_one_panel(caller)
        first = shared._audit_bounds(
            shared._named_audit_region(caller, stage_one.PROPERTY_PANEL_NAME)
        )
        second = shared._audit_bounds(
            shared._named_audit_region(caller, STAGE_TWO_PANEL_NAME)
        )
    except shared.BondsLiteError as error:
        raise PropertyStageTwoError(str(error)) from error
    if (
        first[0] <= second[2]
        and first[2] >= second[0]
        and first[1] <= second[3]
        and first[3] >= second[1]
    ):
        raise PropertyStageTwoError("Property Stage 1 and Stage 2 audit panels overlap")


def _read_property_stage_two_fields(caller: Any) -> dict[str, object]:
    """Read current Stage 2 audit values after validating the named panel."""
    require_property_stage_two_panel(caller)
    return {
        field: cell.value
        for field, cell in _property_stage_two_field_cells(caller).items()
    }


def write_property_stage_two_audit(caller: Any, fields: dict[str, object]) -> None:
    """Write and save Stage 2 audit values through the named cells."""
    require_property_stage_two_panel(caller)
    if set(fields) - STAGE_TWO_FIELD_NAMES.keys():
        raise PropertyStageTwoError("Unknown Property Stage 2 audit field")
    cells = _property_stage_two_field_cells(caller)
    events = caller.app.api.EnableEvents
    try:
        caller.app.api.EnableEvents = False
        for field, value in fields.items():
            cell = cells[field]
            if field == "input_workbook":
                if cell.api.Hyperlinks.Count:
                    cell.api.Hyperlinks.Delete()
                if value:
                    cell.add_hyperlink(str(value), text_to_display=str(value))
                else:
                    cell.value = None
            else:
                cell.value = f"'{value}" if isinstance(value, str) else value
        caller.save()
        if not caller.api.Saved:
            raise PropertyStageTwoError("Draft's Property Stage 2 audit was not saved")
    finally:
        caller.app.api.EnableEvents = events


def _resolve_handoff_path(value: object, caller_path: Path) -> Path:
    """Resolve a saved Stage 1 path relative to the Draft when necessary."""
    if not isinstance(value, str) or not value.strip():
        raise PropertyStageTwoError(
            "Property Stage 1 has no successful output workbook handoff"
        )
    path = Path(os.path.expandvars(value.strip())).expanduser()
    if not path.is_absolute():
        path = caller_path.parent / Path(str(path).replace("/", os.sep))
    try:
        resolved = path.resolve(strict=True)
    except OSError as error:
        raise PropertyStageTwoError(
            f"Property Stage 1 output workbook does not exist: {path}"
        ) from error
    if not resolved.is_file():
        raise PropertyStageTwoError(
            f"Property Stage 1 output path is not a file: {resolved}"
        )
    try:
        size = resolved.stat().st_size
        with resolved.open("rb") as workbook_file:
            workbook_file.read(1)
    except OSError as error:
        raise PropertyStageTwoError(
            f"Property Stage 1 output workbook is unreadable: {resolved}"
        ) from error
    if size == 0:
        raise PropertyStageTwoError(
            f"Property Stage 1 output workbook is empty: {resolved}"
        )
    if resolved.suffix.casefold() not in {".xlsb", ".xlsm"}:
        raise PropertyStageTwoError(
            "Property Stage 1 output must be a macro-enabled .xlsb or .xlsm workbook"
        )
    return resolved


def resolve_property_stage_two_inputs(caller: Any) -> PropertyStageTwoInputs:
    """Validate the Property flag and resolve the successful Stage 1 handoff."""
    caller_path = shared.resolve_caller_path(caller)
    try:
        flag = stage_one.read_property_flag(caller)
        stage_one.require_property_selected(flag)
        stage_one.require_property_stage_one_panel(caller)
        fields = stage_one._read_property_stage_one_fields(caller)
    except shared.BondsLiteError as error:
        raise PropertyStageTwoError(str(error)) from error
    if fields["status"] != "Succeeded":
        raise PropertyStageTwoError(
            "Property Stage 2 requires Property Stage 1 status Succeeded"
        )
    workbook_path = _resolve_handoff_path(fields["last_output"], caller_path)
    return PropertyStageTwoInputs(caller_path, workbook_path)


def _excel_datetime(value: object, name: str) -> datetime:
    """Convert Excel timestamps, including serial dates, to Python datetimes."""
    if isinstance(value, datetime):
        return value.replace(tzinfo=None)
    if isinstance(value, date):
        return datetime.combine(value, datetime.min.time())
    if isinstance(value, Real) and not isinstance(value, bool):
        try:
            return datetime(1899, 12, 30) + timedelta(days=float(value))
        except (OverflowError, ValueError) as error:
            raise PropertyStageTwoError(f"{name} is not a valid timestamp") from error
    if isinstance(value, str):
        try:
            return datetime.fromisoformat(value.strip())
        except ValueError as error:
            raise PropertyStageTwoError(f"{name} is not a valid timestamp") from error
    raise PropertyStageTwoError(f"{name} is missing or is not a valid timestamp")


def _excel_now(application: Any) -> datetime:
    """Capture the invocation instant using the owned Excel instance's clock."""
    try:
        # WorksheetFunction has no Now method; Evaluate returns NOW() as a serial date.
        value = application.api.Evaluate("NOW()")
    except Exception as error:
        raise PropertyStageTwoError(
            f"Cannot read the LTM Excel invocation time: {error}"
        ) from error
    return _excel_datetime(value, "Excel invocation time")


def _named_range_value(workbook: Any, name: str) -> object:
    """Read a scalar from an LTM workbook name with actionable errors."""
    try:
        region = workbook.names[name].refers_to_range
        if region.api.Areas.Count != 1 or region.size != 1:
            raise PropertyStageTwoError(
                f"LTM workbook name must refer to one cell: {name}"
            )
        return region.value
    except PropertyStageTwoError:
        raise
    except Exception as error:
        raise PropertyStageTwoError(
            f"LTM workbook is missing a usable {name} name: {error}"
        ) from error


def _grid_values(value: object) -> list[object]:
    """Flatten Excel range values to inspect content without retaining blanks."""
    if value is None:
        return []
    if isinstance(value, (tuple, list)):
        return [
            item
            for row in value
            for item in (row if isinstance(row, (tuple, list)) else (row,))
        ]
    return [value]


def _output_has_content(sheet: Any) -> bool:
    """Return whether a result sheet's used range contains any non-blank value."""
    try:
        used = sheet.api.UsedRange
        return any(
            value is not None and not (isinstance(value, str) and not value.strip())
            for value in _grid_values(used.Value)
        )
    except Exception as error:
        raise PropertyStageTwoError(
            f"Cannot inspect LTM output sheet {sheet.name}: {error}"
        ) from error


def _progress_messages(workbook: Any) -> list[str]:
    """Read non-blank progress messages below the named first progress cell."""
    try:
        first = workbook.names[LTM_PROGRESS_FIRST_NAME].refers_to_range
        row, column = int(first.row), int(first.column)
        sheet = first.sheet
        last = int(sheet.api.Cells(sheet.api.Rows.Count, column).End(XL_UP).Row)
        if last < row:
            return []
        value = sheet.range((row, column), (last, column)).value
    except Exception as error:
        raise PropertyStageTwoError(
            f"Cannot read the LTM progress log from {LTM_PROGRESS_FIRST_NAME}: {error}"
        ) from error
    values = _grid_values(value)
    return [
        str(item).strip() for item in values if item is not None and str(item).strip()
    ]


def verify_ltm_run(
    workbook: Any, invocation_time: datetime, expected_scenarios: int
) -> None:
    """Validate timestamps, per-scenario progress, terminal message and outputs."""
    run_start = _excel_datetime(
        _named_range_value(workbook, LTM_RUN_START_NAME), LTM_RUN_START_NAME
    )
    run_end = _excel_datetime(
        _named_range_value(workbook, LTM_RUN_END_NAME), LTM_RUN_END_NAME
    )
    if run_start < invocation_time - timedelta(seconds=2):
        raise PropertyStageTwoError("LTM RunStartTime is stale for this invocation")
    if run_start > run_end:
        raise PropertyStageTwoError("LTM RunEndTime is before RunStartTime")
    if expected_scenarios < 1:
        raise PropertyStageTwoError("prm_NoOfScens must be a positive whole number")
    messages = _progress_messages(workbook)
    for scenario in range(1, expected_scenarios + 1):
        marker = f"Run {scenario} - End of run!"
        if marker not in messages:
            raise PropertyStageTwoError(f"LTM progress log is missing {marker}")
    if not messages or messages[-1] != "End of multirun!":
        raise PropertyStageTwoError(
            "LTM progress log does not end with End of multirun!"
        )
    for sheet_name in LTM_OUTPUT_SHEETS:
        try:
            sheet = workbook.sheets[sheet_name]
        except Exception as error:
            raise PropertyStageTwoError(
                f"LTM workbook is missing output sheet {sheet_name}"
            ) from error
        if not _output_has_content(sheet):
            raise PropertyStageTwoError(
                f"LTM output sheet {sheet_name} is empty after the run"
            )


def _scenario_count(workbook: Any) -> int:
    """Read and validate the run's expected scenario count."""
    value = _named_range_value(workbook, LTM_SCENARIO_COUNT_NAME)
    if (
        not isinstance(value, Real)
        or isinstance(value, bool)
        or not float(value).is_integer()
        or value < 1
    ):
        raise PropertyStageTwoError("prm_NoOfScens must be a positive whole number")
    return int(value)


def _check_runner() -> None:
    """Require the exact LTM executable referenced by the workbook macro."""
    if not LTM_RUNNER_PATH.is_file():
        raise PropertyStageTwoError(
            f"LTM model executable is unavailable: {LTM_RUNNER_PATH}"
        )


def _validate_local_run_settings(workbook: Any) -> None:
    """Reject workbook settings that make the public LTM macro refuse the run."""
    local_run = _named_range_value(workbook, LTM_LOCAL_RUN_NAME)
    data_source = _named_range_value(workbook, LTM_DATA_SOURCE_NAME)
    if local_run is not True:
        raise PropertyStageTwoError(
            "LTM workbook prm_LocalRun must be TRUE before the macro can run"
        )
    if (
        not isinstance(data_source, str)
        or data_source.strip().casefold() != "sqlserver"
    ):
        raise PropertyStageTwoError(
            "LTM workbook prm_DataSource must be SQLserver before the macro can run"
        )


def _excel_local_path(workbook: Any, caller: Any, caller_path: Path) -> Path:
    """Map Excel's FullName to a local path through the Draft's own location.

    xlwings maps synced SharePoint URLs by searching for the file name, which
    fails when OneDrive holds another file with the same name.
    """
    raw = str(workbook.api.FullName)
    if not re.match(r"^[a-z]+://", raw, re.IGNORECASE):
        return Path(raw).resolve()
    caller_raw = unquote(str(caller.api.FullName))
    value = unquote(raw)
    if re.match(r"^[a-z]+://", caller_raw, re.IGNORECASE):
        folder = caller_raw.rsplit("/", 1)[0] + "/"
        if value.casefold().startswith(folder.casefold()):
            return caller_path.parent.joinpath(*value[len(folder):].split("/")).resolve()
    raise PropertyStageTwoError(
        f"Cannot map the LTM workbook location to a local path: {raw}"
    )


def _ensure_workbook_not_open(path: Path, caller: Any, caller_path: Path) -> None:
    """Reject a target already open in any Excel instance before writing."""
    try:
        import xlwings as xw

        wanted = os.path.normcase(str(path.resolve()))
        for application in xw.apps:
            for workbook in application.books:
                try:
                    opened = _excel_local_path(workbook, caller, caller_path)
                except PropertyStageTwoError:
                    continue
                if os.path.normcase(str(opened)) == wanted:
                    raise PropertyStageTwoError(
                        f"Close {path.name} in Excel before running Property Stage 2"
                    )
    except PropertyStageTwoError:
        raise
    except Exception as error:
        raise PropertyStageTwoError(
            f"Cannot verify that the LTM workbook is closed: {error}"
        ) from error


def _require_expected_workbook(
    workbook: Any, expected_path: Path, caller: Any, caller_path: Path
) -> None:
    """Confirm Excel opened the exact Stage 1 handoff in the owned instance."""
    try:
        actual_path = _excel_local_path(workbook, caller, caller_path)
    except PropertyStageTwoError:
        raise
    except Exception as error:
        raise PropertyStageTwoError(
            f"Cannot verify the opened LTM workbook path: {error}"
        ) from error
    if os.path.normcase(str(actual_path)) != os.path.normcase(
        str(expected_path.resolve(strict=True))
    ):
        raise PropertyStageTwoError(
            f"Excel opened an unexpected LTM workbook: {actual_path}"
        )


def _portable_workbook_path(caller_path: Path, workbook_path: Path) -> str:
    """Return the Stage 2 audit link as a Draft-relative portable path."""
    try:
        return (
            workbook_path.resolve().relative_to(caller_path.parent.resolve()).as_posix()
        )
    except ValueError:
        return os.path.relpath(
            workbook_path.resolve(), caller_path.parent.resolve()
        ).replace(os.sep, "/")


def _audit_values(
    started: datetime,
    clock: float,
    status: str,
    workbook_link: str | None,
    last_successful_run: object,
    *,
    finish: datetime | None = None,
) -> dict[str, object]:
    """Create common Stage 2 audit values while retaining the last success."""
    return {
        "last_run": started.replace(hour=0, minute=0, second=0, microsecond=0),
        "start": started,
        "finish": finish,
        "elapsed": (time.perf_counter() - clock) / 86400 if finish else None,
        "user": getpass.getuser(),
        "status": status,
        "input_workbook": workbook_link,
        "last_successful_run": last_successful_run,
    }


def _record_failure(
    caller: Any,
    started: datetime,
    clock: float,
    workbook_link: str | None,
    previous_success: object,
    status: str,
) -> None:
    """Persist a terminal failure while retaining the previous successful run."""
    write_property_stage_two_audit(
        caller,
        _audit_values(
            started,
            clock,
            status,
            workbook_link,
            previous_success,
            finish=datetime.now(),
        ),
    )


def _invoke_ltm(
    workbook: Any, application: Any, invocation_time: datetime, scenarios: int
) -> None:
    """Call the supported public LTM macro once, then verify its result."""
    try:
        workbook.macro("Module1.CallRunFrozenPython")("flipper")
    except Exception as error:
        raise PropertyStageTwoError(
            f"LTM CallRunFrozenPython failed: {error}"
        ) from error
    verify_ltm_run(workbook, invocation_time, scenarios)


def run_property_stage_two(caller: Any) -> Path:
    """Run and verify LTM Flash in a separate visible Excel instance, in place."""
    caller_path = shared.resolve_caller_path(caller)
    with shared.workbook_run_guard(
        caller_path,
        conflict_message="A stress run is already in progress for this workbook",
    ):
        previous_success: object = None
        try:
            previous_success = _read_property_stage_two_fields(caller)[
                "last_successful_run"
            ]
        except Exception as error:
            raise PropertyStageTwoError(
                f"Property Stage 2 audit panel is unavailable: {error}"
            ) from error
        started = datetime.now()
        clock = time.perf_counter()
        workbook_link: str | None = None
        application = workbook = None
        model_saved = False
        save_attempted = False
        backup_directory: tempfile.TemporaryDirectory[str] | None = None
        backup_path: Path | None = None
        try:
            inputs = resolve_property_stage_two_inputs(caller)
            workbook_link = _portable_workbook_path(caller_path, inputs.workbook_path)
            _check_runner()
            _ensure_workbook_not_open(inputs.workbook_path, caller, caller_path)
            write_property_stage_two_audit(
                caller,
                _audit_values(
                    started,
                    clock,
                    "Running",
                    workbook_link,
                    previous_success,
                ),
            )

            import xlwings as xw

            before_mtime = inputs.workbook_path.stat().st_mtime_ns
            backup_directory = tempfile.TemporaryDirectory(prefix="property-stage-two-")
            backup_path = Path(backup_directory.name) / inputs.workbook_path.name
            shutil.copy2(inputs.workbook_path, backup_path)
            application = xw.App(visible=True, add_book=False)
            application.api.AutomationSecurity = XL_AUTOMATION_SECURITY_LOW
            # The LTM template is saved "read-only recommended"; Stage 1 copies inherit it.
            workbook = application.books.open(
                str(inputs.workbook_path),
                update_links=False,
                read_only=False,
                ignore_read_only_recommended=True,
            )
            _require_expected_workbook(
                workbook, inputs.workbook_path, caller, caller_path
            )
            if workbook.api.ReadOnly:
                raise PropertyStageTwoError(
                    "Property Stage 1 output workbook is read-only"
                )
            _validate_local_run_settings(workbook)
            application.api.CalculateFullRebuild()
            invocation_time = _excel_now(application)
            scenarios = _scenario_count(workbook)
            _invoke_ltm(workbook, application, invocation_time, scenarios)

            save_attempted = True
            workbook.save()
            if not workbook.api.Saved:
                raise PropertyStageTwoError(
                    "LTM Flash workbook did not report that it was saved"
                )
            if inputs.workbook_path.stat().st_mtime_ns <= before_mtime:
                raise PropertyStageTwoError(
                    "LTM Flash workbook modification time did not advance after saving"
                )
            model_saved = True
            finished = datetime.now()
            try:
                write_property_stage_two_audit(
                    caller,
                    _audit_values(
                        started,
                        clock,
                        "Succeeded",
                        workbook_link,
                        finished,
                        finish=finished,
                    ),
                )
            except Exception as audit_error:
                failure_message = (
                    "LTM Flash workbook was verified and saved, but the final "
                    "Property Stage 2 success audit could not be saved; the "
                    "previous Last Successful Run was retained"
                )
                raise PropertyStageTwoError(
                    f"{failure_message}: {audit_error}"
                ) from audit_error
            logger.info(
                "Property Stage 2 succeeded; updated LTM Flash in place: %s",
                inputs.workbook_path,
            )
            return inputs.workbook_path
        except BaseException as error:
            status = "Interrupted" if isinstance(error, KeyboardInterrupt) else "Failed"
            try:
                _record_failure(
                    caller,
                    started,
                    clock,
                    workbook_link,
                    previous_success,
                    status,
                )
            except Exception as audit_error:
                if model_saved:
                    logger.error(
                        "LTM Flash was saved, but its final audit could not "
                        "be persisted: %s",
                        audit_error,
                    )
                else:
                    logger.error(
                        "Property Stage 2 failed and its audit could not be "
                        "saved: %s",
                        audit_error,
                    )
            raise
        finally:
            if workbook is not None:
                try:
                    workbook.api.Close(SaveChanges=False)
                except Exception as cleanup_error:
                    logger.error(
                        "Could not close the owned LTM workbook: %s",
                        cleanup_error,
                    )
            if application is not None:
                try:
                    application.api.DisplayAlerts = False
                    application.quit()
                except Exception as cleanup_error:
                    logger.error(
                        "Could not close the owned LTM Excel instance: %s",
                        cleanup_error,
                    )
            restore_error: Exception | None = None
            if save_attempted and not model_saved and backup_path is not None:
                try:
                    shutil.copy2(backup_path, inputs.workbook_path)
                    if shared._file_hash(backup_path) != shared._file_hash(
                        inputs.workbook_path
                    ):
                        raise OSError("the restored workbook hash does not match")
                except Exception as error:
                    restore_error = error
            if backup_directory is not None:
                try:
                    backup_directory.cleanup()
                except Exception as cleanup_error:
                    logger.error(
                        "Could not remove the temporary Stage 2 backup: %s",
                        cleanup_error,
                    )
            if restore_error is not None:
                raise PropertyStageTwoError(
                    "Property Stage 2 could not restore the Stage 1 workbook "
                    f"after failed save verification: {restore_error}"
                ) from restore_error


def _name_keys(book: Any) -> set[str]:
    """Return workbook-level Excel name keys without sheet-qualified prefixes."""
    return {name.name.split("!")[-1] for name in book.names}


def _panel_overlap(
    first: tuple[float, float, float, float], second: tuple[float, float, float, float]
) -> bool:
    """Return whether two Excel shape/range rectangles intersect."""
    return (
        first[0] < second[0] + second[2]
        and first[0] + first[2] > second[0]
        and first[1] < second[1] + second[3]
        and first[1] + first[3] > second[1]
    )


def _initialise_property_stage_two_panel(book: Any) -> None:
    """Install a move-safe named audit panel below the current Stage 1 panel."""
    stage_one.require_property_stage_one_panel(book)
    expected = {STAGE_TWO_PANEL_NAME, *STAGE_TWO_FIELD_NAMES.values()}
    existing = _name_keys(book)
    if existing.intersection(expected):
        if not expected.issubset(existing):
            raise PropertyStageTwoError(
                "Property Stage 2 audit names are incomplete; installation was not changed"
            )
        require_property_stage_two_panel(book)
        return

    sheet = book.sheets["I. Control"]
    previous = shared._named_audit_region(book, stage_one.PROPERTY_PANEL_NAME)
    try:
        stage_one_button = sheet.api.Shapes.Item(stage_one.PROPERTY_BUTTON_NAME)
    except Exception as error:
        raise PropertyStageTwoError(
            "Cannot locate the Property Stage 1 button; no Stage 2 panel was installed"
        ) from error
    row = previous.row + previous.rows.count + 2
    while (
        sheet.range((row, previous.column)).top
        < max(
            previous.top + previous.height,
            float(stage_one_button.Top) + float(stage_one_button.Height),
        )
        + 12
    ):
        row += 1
    column = previous.column
    area = sheet.range((row, column), (row + 10, column + 7))
    if (
        book.app.api.WorksheetFunction.CountA(area.api)
        or area.api.MergeCells is not False
    ):
        raise PropertyStageTwoError(
            "Proposed Property Stage 2 audit area contains cells or merges; nothing was overwritten"
        )
    panel_box = (area.left, area.top, area.width, area.height)
    for shape in sheet.api.Shapes:
        shape_box = (
            float(shape.Left),
            float(shape.Top),
            float(shape.Width),
            float(shape.Height),
        )
        if _panel_overlap(shape_box, panel_box):
            raise PropertyStageTwoError(
                "Proposed Property Stage 2 audit area overlaps a shape or button"
            )

    events = book.app.api.EnableEvents
    added: list[str] = []
    try:
        book.app.api.EnableEvents = False
        area.color = (230, 230, 255)
        area.font.name = "Calibri"
        area.font.size = 10
        area.api.VerticalAlignment = -4108
        heading = sheet.range((row, column), (row, column + 7))
        heading.merge()
        heading.value = "Property Stage 2 - Run LTM Flash"
        heading.color = (127, 127, 255)
        heading.font.color = (255, 255, 255)
        heading.font.bold = True
        book.names.add(STAGE_TWO_PANEL_NAME, f"='I. Control'!{area.get_address()}")
        added.append(STAGE_TWO_PANEL_NAME)
        for field, offset in STAGE_TWO_OFFSETS.items():
            end_row = row + offset + (1 if field == "input_workbook" else 0)
            label = sheet.range((row + offset, column), (end_row, column + 3))
            value = sheet.range((row + offset, column + 4), (end_row, column + 7))
            label.merge()
            value.merge()
            label.value = STAGE_TWO_LABELS[field]
            label.color = (127, 127, 255)
            label.font.color = (255, 255, 255)
            label.api.WrapText = True
            value.api.WrapText = True
            value.font.color = (0, 0, 0)
            value.api.HorizontalAlignment = -4131
            if field == "last_run":
                value.number_format = "dd/mm/yyyy"
            elif field in ("start", "finish", "last_successful_run"):
                value.number_format = "dd/mm/yyyy hh:mm:ss"
            elif field == "elapsed":
                value.number_format = "[h]:mm:ss"
            else:
                value.number_format = "General"
            if field == "status":
                value.value = "Not Run"
            name = STAGE_TWO_FIELD_NAMES[field]
            book.names.add(
                name,
                f"='I. Control'!{value[0, 0].get_address()}",
            )
            added.append(name)
    except Exception:
        for name in reversed(added):
            book.names[name].delete()
        raise
    finally:
        book.app.api.EnableEvents = events
    require_property_stage_two_panel(book)


def _install_property_stage_two_button(book: Any) -> None:
    """Install or refresh the Stage 2 button without moving existing controls."""
    sheet = book.sheets["I. Control"]
    if STAGE_TWO_BUTTON_NAME in [shape.Name for shape in sheet.api.Shapes]:
        button = sheet.api.Buttons().Item(STAGE_TWO_BUTTON_NAME)
    else:
        stage_one_button = sheet.api.Shapes.Item(stage_one.PROPERTY_BUTTON_NAME)
        panel = book.names[STAGE_TWO_PANEL_NAME].refers_to_range
        left = float(stage_one_button.Left)
        top = float(panel.top)
        width = float(stage_one_button.Width)
        height = float(stage_one_button.Height)
        row = panel.row
        column = int(stage_one_button.TopLeftCell.Column)
        end_column = column
        while sheet.range((row, end_column)).left < left + width:
            end_column += 1
        cells = sheet.range((row, column), (row + 2, end_column - 1))
        if (
            book.app.api.WorksheetFunction.CountA(cells.api)
            or cells.api.MergeCells is not False
        ):
            raise PropertyStageTwoError(
                "Proposed Run LTM Flash button covers occupied cells or merges"
            )
        proposed = (left, top, width, height)
        for shape in sheet.api.Shapes:
            shape_box = (
                float(shape.Left),
                float(shape.Top),
                float(shape.Width),
                float(shape.Height),
            )
            if _panel_overlap(shape_box, proposed):
                raise PropertyStageTwoError(
                    "Proposed Run LTM Flash button overlaps an existing shape"
                )
        button = sheet.api.Buttons().Add(left, top, width, height)
        button.Name = STAGE_TWO_BUTTON_NAME
    button.Caption = "Run LTM Flash"
    button.OnAction = "'" + book.name.replace("'", "''") + "'!" + STAGE_TWO_BUTTON_MACRO


def _install_property_stage_two_macro(book: Any) -> None:
    """Add the P2 VBA entry point to the existing shared launcher module."""
    try:
        module = book.api.VBProject.VBComponents.Item("modBondsLiteLauncher")
        code = module.CodeModule
        source = code.Lines(1, code.CountOfLines) if code.CountOfLines else ""
    except Exception as error:
        raise PropertyStageTwoError(
            "Cannot access modBondsLiteLauncher in the Draft VBA project"
        ) from error
    if "Public Sub RunLTMFlash()".casefold() in source.casefold():
        if 'LaunchBondsLite "P2"'.casefold() not in source.casefold():
            raise PropertyStageTwoError(
                "Draft RunLTMFlash macro has an unexpected implementation"
            )
        return
    if (
        "Private Sub LaunchBondsLite(ByVal stage As String)".casefold()
        not in source.casefold()
    ):
        raise PropertyStageTwoError(
            "Draft launcher does not contain a supported LaunchBondsLite signature"
        )
    code.InsertLines(
        code.CountOfLines + 1,
        '\nPublic Sub RunLTMFlash()\n    LaunchBondsLite "P2"\nEnd Sub\n',
    )


def install_property_stage_two_controls(caller: Path) -> Path:
    """Install Stage 2's named panel, button and macro into a closed Draft."""
    import xlwings as xw

    caller = Path(caller).resolve(strict=True)
    for application in xw.apps:
        for book in application.books:
            if book.name.casefold() == caller.name.casefold():
                raise PropertyStageTwoError(
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
                raise PropertyStageTwoError("Draft workbook is read-only or locked")
            stage_one.require_property_stage_one_panel(book)
            _initialise_property_stage_two_panel(book)
            _install_property_stage_two_button(book)
            _install_property_stage_two_macro(book)
            book.save()
        finally:
            book.close()
    finally:
        application.api.DisplayAlerts = False
        application.api.Quit()
    return backup
