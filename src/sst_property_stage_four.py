"""Run NoteVal's Run macro on the successful Property Stage 3 workbook."""

from __future__ import annotations

import gc
import getpass
import math
import os
import shutil
import sys
import tempfile
import threading
import time
from dataclasses import dataclass
from datetime import datetime, timedelta
from numbers import Real
from pathlib import Path
from typing import Any, Callable

import bonds_lite_input as shared
import sst_property_stage_one as stage_one
import sst_property_stage_three as stage_three
import sst_property_stage_two as stage_two
import sst_property_stage_two_watchdog as stage_two_watchdog

logger = shared.logger.getChild("sst_property_stage_four")

STAGE_FOUR_PANEL_NAME = "SST_PropStage4_Panel"
STAGE_FOUR_PANEL_RANGE = "AB48:AI58"
STAGE_FOUR_PANEL_HEADING = "Property Stage 4 - Run NoteVal"
STAGE_FOUR_BUTTON_NAME = "btnRunNoteVal"
STAGE_FOUR_BUTTON_MACRO = "RunNoteVal"
STAGE_FOUR_BUTTON_CAPTION = "Run NoteVal"
STAGE_FOUR_FIELD_NAMES = {
    "last_run": "SST_PropStage4_LastRun",
    "start": "SST_PropStage4_Start",
    "finish": "SST_PropStage4_Finish",
    "elapsed": "SST_PropStage4_Elapsed",
    "user": "SST_PropStage4_User",
    "status": "SST_PropStage4_Status",
    "input_workbook": "SST_PropStage4_InputWorkbook",
    "last_successful_run": "SST_PropStage4_LastSuccessfulRun",
}
STAGE_FOUR_LABELS = dict(stage_two.STAGE_TWO_LABELS)
STAGE_FOUR_OFFSETS = dict(stage_two.STAGE_TWO_OFFSETS)

NOTEVAL_MACRO = "modControl.Run"
NOTEVAL_START_NAME = "I_StartTime_1"
NOTEVAL_END_NAME = "I_EndTime_1"
START_TOLERANCE = timedelta(seconds=2)
# GoalSeek targets set by the five NoteVal solvers (Decision 129).
NOTEVAL_RESIDUAL_CELLS = (
    ("O. MAC MI", ("D12",)),
    (
        "C. Note Valuation ex Hedge",
        ("F12", "H12", "AB31:AM31", "AF45", "AQ33:AQ36", "AB57:AM57"),
    ),
    ("C. Note Valuation", ("T23", "T50", "T56")),
)

TIME_LIMIT_ENVIRONMENT = "SST_P4_TIME_LIMIT_MINUTES"
DEFAULT_TIME_LIMIT = timedelta(minutes=30)
CONTINUE_INTERVAL_SECONDS = 30 * 60
HEARTBEAT_SECONDS = 5 * 60
POLL_SECONDS = 30
XL_AUTOMATION_SECURITY_LOW = stage_two.XL_AUTOMATION_SECURITY_LOW
XL_AUTOMATION_SECURITY_FORCE_DISABLE = 3


class PropertyStageFourError(stage_three.PropertyStageThreeError):
    """An actionable Property Stage 4 input, NoteVal run or audit failure."""


class PropertyStageFourTimeoutError(PropertyStageFourError):
    """The user chose to abort a Property Stage 4 run at its time limit."""


@dataclass(frozen=True)
class PropertyStageFourInputs:
    """Validated Stage 3 handoff for an in-place NoteVal run.

    Attributes:
        caller_path: Local path of the calling Draft workbook.
        workbook_path: The prepared NoteVal workbook that Stage 4 runs in place.
        flash_link: Draft-relative link to the Flash workbook NoteVal reads.
    """

    caller_path: Path
    workbook_path: Path
    flash_link: str


# ---------------------------------------------------------------------------
# Audit panel
# ---------------------------------------------------------------------------


def _property_stage_four_field_cells(caller: Any) -> dict[str, Any]:
    """Resolve the Stage 4 audit fields to the first cells of merged ranges."""
    cells: dict[str, Any] = {}
    for field, name in STAGE_FOUR_FIELD_NAMES.items():
        region = stage_one._audit_region(caller, name)
        merged = region.api.Cells(1, 1).MergeArea
        cells[field] = region.sheet.range((merged.Row, merged.Column))
    return cells


def require_property_stage_four_panel(caller: Any) -> None:
    """Validate Stage 4's named panel and ensure it does not overlap Stage 3.

    Args:
        caller: The Draft workbook.

    Raises:
        PropertyStageFourError: If the panel is missing, malformed or overlaps
            the Stage 3 panel.
    """
    try:
        shared._require_named_stage_panel(
            caller, STAGE_FOUR_PANEL_NAME, STAGE_FOUR_FIELD_NAMES
        )
        stage_three.require_property_stage_three_panel(caller)
        third = shared._audit_bounds(
            shared._named_audit_region(caller, stage_three.STAGE_THREE_PANEL_NAME)
        )
        fourth = shared._audit_bounds(
            shared._named_audit_region(caller, STAGE_FOUR_PANEL_NAME)
        )
    except shared.BondsLiteError as error:
        raise PropertyStageFourError(str(error)) from error
    if (
        third[0] <= fourth[2]
        and third[2] >= fourth[0]
        and third[1] <= fourth[3]
        and third[3] >= fourth[1]
    ):
        raise PropertyStageFourError(
            "Property Stage 3 and Stage 4 audit panels overlap"
        )


def _read_property_stage_four_fields(caller: Any) -> dict[str, object]:
    """Read current Stage 4 audit values after validating the named panel."""
    require_property_stage_four_panel(caller)
    return {
        field: cell.value
        for field, cell in _property_stage_four_field_cells(caller).items()
    }


def write_property_stage_four_audit(caller: Any, fields: dict[str, object]) -> None:
    """Write and save Stage 4 audit values through the named cells.

    Args:
        caller: The Draft workbook.
        fields: Audit values keyed by ``STAGE_FOUR_FIELD_NAMES`` keys.

    Raises:
        PropertyStageFourError: If a field is unknown or the Draft did not save.
    """
    require_property_stage_four_panel(caller)
    if set(fields) - STAGE_FOUR_FIELD_NAMES.keys():
        raise PropertyStageFourError("Unknown Property Stage 4 audit field")
    cells = _property_stage_four_field_cells(caller)
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
            raise PropertyStageFourError("Draft's Property Stage 4 audit was not saved")
    finally:
        caller.app.api.EnableEvents = events


def _audit_values(
    started: datetime,
    clock: float,
    status: str,
    workbook_link: str | None,
    last_successful_run: object,
    *,
    finish: datetime | None = None,
) -> dict[str, object]:
    """Build the Stage 4 audit fields while retaining the last successful run."""
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
    write_property_stage_four_audit(
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


# ---------------------------------------------------------------------------
# Input resolution (Decision 126)
# ---------------------------------------------------------------------------


def resolve_property_stage_four_inputs(caller: Any) -> PropertyStageFourInputs:
    """Resolve the NoteVal workbook from a current, successful Stage 3 audit.

    Args:
        caller: The Draft workbook.

    Returns:
        The validated Stage 4 inputs.

    Raises:
        PropertyStageFourError: If Property is not selected, Stage 3 did not
            succeed, its output is missing, its Flash input differs from Stage
            2's, or Flash was rerun after NoteVal was prepared.
    """
    caller_path = shared.resolve_caller_path(caller)
    try:
        stage_one.require_property_selected(stage_one.read_property_flag(caller))
        stage_two.require_property_stage_two_panel(caller)
        stage_three.require_property_stage_three_panel(caller)
        stage_two_fields = stage_two._read_property_stage_two_fields(caller)
        stage_three_fields = stage_three._read_property_stage_three_fields(caller)
    except shared.BondsLiteError as error:
        raise PropertyStageFourError(str(error)) from error
    if stage_three_fields["status"] != "Succeeded":
        raise PropertyStageFourError(
            "Property Stage 4 requires Property Stage 3 status Succeeded"
        )
    try:
        workbook_path = stage_two._resolve_handoff_path(
            stage_three_fields["last_successful_output"],
            caller_path,
            "Property Stage 3 NoteVal output workbook",
        )
    except stage_two.PropertyStageTwoError as error:
        raise PropertyStageFourError(
            f"Property Stage 3 NoteVal output cannot be used: {error}"
        ) from error
    try:
        stage_three_flash = stage_two._resolve_handoff_path(
            stage_three_fields["input_workbook"],
            caller_path,
            "Property Stage 3 Flash input workbook",
        )
        stage_two_flash = stage_two._resolve_handoff_path(
            stage_two_fields["input_workbook"],
            caller_path,
            "Property Stage 2 Flash input workbook",
        )
    except stage_two.PropertyStageTwoError as error:
        raise PropertyStageFourError(
            f"Property Stage 2 or Stage 3 Flash input cannot be resolved: {error}"
        ) from error
    if os.path.normcase(str(stage_three_flash)) != os.path.normcase(
        str(stage_two_flash)
    ):
        raise PropertyStageFourError(
            "Property Stage 3 input workbook does not match the Property Stage 2 "
            "input workbook; rerun Property Stage 3"
        )
    try:
        stage_two_run = stage_two._excel_datetime(
            stage_two_fields["last_successful_run"],
            "Property Stage 2 last successful run",
        )
        stage_three_finish = stage_two._excel_datetime(
            stage_three_fields["finish"], "Property Stage 3 finish"
        )
    except stage_two.PropertyStageTwoError as error:
        raise PropertyStageFourError(str(error)) from error
    if stage_two_run > stage_three_finish:
        raise PropertyStageFourError(
            "Property Stage 2 ran after NoteVal was prepared; rerun Property Stage 3"
        )
    return PropertyStageFourInputs(
        caller_path=caller_path,
        workbook_path=workbook_path,
        flash_link=stage_two._portable_workbook_path(caller_path, stage_three_flash),
    )


# ---------------------------------------------------------------------------
# Completion checks (Decision 129)
# ---------------------------------------------------------------------------


def _noteval_time(workbook: Any, name: str) -> datetime:
    """Read one NoteVal run timestamp, naming the macro as the missing writer."""
    try:
        return stage_two._excel_datetime(
            stage_two._named_range_value(workbook, name), name
        )
    except stage_two.PropertyStageTwoError as error:
        raise PropertyStageFourError(
            f"NoteVal Run did not record a valid {name}: {error}"
        ) from error


def _flatten(value: object) -> list[object]:
    """Flatten Excel scalar, row and two-dimensional range values."""
    if isinstance(value, (tuple, list)):
        return [item for child in value for item in _flatten(child)]
    return [value]


def _is_finite_number(value: object) -> bool:
    return (
        isinstance(value, Real)
        and not isinstance(value, bool)
        and math.isfinite(float(value))
    )


def verify_noteval_run(workbook: Any, invocation_time: datetime) -> dict[str, float]:
    """Confirm NoteVal's Run macro completed for this invocation.

    Args:
        workbook: The NoteVal workbook after ``modControl.Run`` returned.
        invocation_time: Excel ``Now()`` captured just before the macro call.

    Returns:
        The largest absolute GoalSeek residual on each checked sheet. They are
        logged for review; no tolerance is applied (Decision 129).

    Raises:
        PropertyStageFourError: If a run timestamp is missing, stale or out of
            order, or a GoalSeek target is not a finite number.
    """
    start = _noteval_time(workbook, NOTEVAL_START_NAME)
    end = _noteval_time(workbook, NOTEVAL_END_NAME)
    if start < invocation_time - START_TOLERANCE:
        raise PropertyStageFourError(
            f"NoteVal {NOTEVAL_START_NAME} is stale; the Run macro did not start "
            "in this invocation"
        )
    if end < start:
        raise PropertyStageFourError(
            f"NoteVal {NOTEVAL_END_NAME} is before {NOTEVAL_START_NAME}; "
            "the Run macro did not finish"
        )
    residuals: dict[str, float] = {}
    for sheet_name, addresses in NOTEVAL_RESIDUAL_CELLS:
        largest = 0.0
        for address in addresses:
            location = f"{sheet_name}!{address}"
            try:
                values = _flatten(workbook.sheets[sheet_name].range(address).value)
            except Exception as error:
                raise PropertyStageFourError(
                    f"Cannot read NoteVal {location}: {error}"
                ) from error
            for value in values:
                if not _is_finite_number(value):
                    raise PropertyStageFourError(
                        f"NoteVal {location} is not a finite number after the "
                        f"Run macro: {value!r}"
                    )
                largest = max(largest, abs(float(value)))
        residuals[sheet_name] = largest
        logger.info(
            "NoteVal residuals on %s: largest absolute value %.6g",
            sheet_name,
            largest,
        )
    return residuals


# ---------------------------------------------------------------------------
# Watchdog (Decision 130)
# ---------------------------------------------------------------------------


def configured_noteval_time_limit(override: str | None, frozen: bool) -> timedelta:
    """Return the NoteVal time limit, honouring a source-only override.

    Args:
        override: The ``SST_P4_TIME_LIMIT_MINUTES`` value, if set.
        frozen: Whether this is the packaged controller.

    Returns:
        30 minutes, or the override in minutes when running from source.

    Raises:
        ValueError: If a source-mode override is not a positive finite number.
    """
    if override is None or not override.strip():
        return DEFAULT_TIME_LIMIT
    if frozen:
        logger.warning("Ignoring %s in the packaged controller", TIME_LIMIT_ENVIRONMENT)
        return DEFAULT_TIME_LIMIT
    message = f"{TIME_LIMIT_ENVIRONMENT} must be a positive finite number"
    try:
        minutes = float(override)
    except ValueError as error:
        raise ValueError(message) from error
    if not math.isfinite(minutes) or minutes <= 0:
        raise ValueError(message)
    return timedelta(minutes=minutes)


class NoteValRunWatchdog(threading.Thread):
    """Log heartbeats and offer a user-controlled time limit for a NoteVal run.

    Unlike Stage 2's watchdog it has no runner process or progress log to
    inspect: NoteVal runs entirely inside the owned Excel process.
    """

    def __init__(
        self,
        *,
        excel_pid: int,
        time_limit: timedelta = DEFAULT_TIME_LIMIT,
        prompt: Callable[[str], str] | None = None,
        logger: Callable[[str], None] = print,
        clock: Any = time,
        terminate: Callable[[int], None] | None = None,
    ) -> None:
        """Prepare the watchdog for one owned Excel process.

        Args:
            excel_pid: The PID of the Excel instance owned by Stage 4.
            time_limit: Elapsed time before the first wait-or-abort prompt.
            prompt: Optional replacement for the console prompt (tests).
            logger: Receives heartbeat, prompt and warning messages.
            clock: Provides ``monotonic`` (and ``sleep`` when not ``time``).
            terminate: Terminates one exact PID; defaults to Stage 2's helper.
        """
        super().__init__(name="PropertyStageFourNoteValWatchdog", daemon=True)
        self.excel_pid = excel_pid
        self._time_limit = time_limit
        self._prompt = prompt
        self._logger = logger
        self._clock = clock
        self._sleep = getattr(clock, "sleep", time.sleep)
        self._terminate = terminate
        self._stop_event = threading.Event()
        self._aborted_event = threading.Event()
        self._termination_completed = threading.Event()

    @property
    def aborted(self) -> bool:
        """Whether the user selected abort at the time limit."""
        return self._aborted_event.is_set()

    @property
    def process_terminated(self) -> bool:
        """Whether the owned Excel process was terminated."""
        return self._termination_completed.is_set()

    def stop(self) -> None:
        """Stop the watcher and release any cancellable console prompt."""
        self._stop_event.set()

    def run(self) -> None:
        """Emit heartbeats and prompt at the time limit until stopped."""
        started = self._clock.monotonic()
        next_heartbeat = started + HEARTBEAT_SECONDS
        next_limit = started + self._time_limit.total_seconds()
        termination_retry_at: float | None = None
        while not self._stop_event.is_set():
            now = self._clock.monotonic()
            if now >= next_heartbeat:
                self._logger(
                    "NoteVal running | elapsed "
                    f"{stage_two_watchdog._format_elapsed(now - started)}"
                )
                while next_heartbeat <= now:
                    next_heartbeat += HEARTBEAT_SECONDS
            if termination_retry_at is not None and now >= termination_retry_at:
                termination_retry_at = (
                    None if self._attempt_termination() else now + POLL_SECONDS
                )
            if now >= next_limit and not self._aborted_event.is_set():
                if self._ask_to_continue() == "A":
                    self._aborted_event.set()
                    if not self._attempt_termination():
                        termination_retry_at = now + POLL_SECONDS
                else:
                    next_limit = self._clock.monotonic() + CONTINUE_INTERVAL_SECONDS
                    self._logger(
                        "NoteVal run will continue; the next time-limit prompt "
                        "is in 30 minutes"
                    )
            if self._clock is time:
                self._stop_event.wait(POLL_SECONDS)
            else:
                self._sleep(POLL_SECONDS)

    def _ask_to_continue(self) -> str:
        """Ask whether to wait another interval or abort the NoteVal run."""
        message = (
            "The Property Stage 4 run has reached its time limit. "
            "[W]ait 30 more minutes or [A]bort?"
        )
        while not self._stop_event.is_set():
            try:
                if self._prompt is None:
                    response = self._console_prompt(message)
                else:
                    response = self._prompt(message)
            except EOFError:
                response = ""
            if self._stop_event.is_set():
                return "W"
            answer = response.strip().casefold()
            if answer in {"", "w", "wait"}:
                return "W"
            if answer in {"a", "abort"}:
                return "A"
            self._logger("Please enter W to wait or A to abort")
        return "W"

    def _console_prompt(self, message: str) -> str:
        """Read a cancellable W/A response using Stage 2's console prompt."""
        # Stage 2's prompt needs only ``_logger`` and ``_stop_event``.
        return stage_two_watchdog.LtmRunWatchdog._console_prompt(self, message)

    def _attempt_termination(self) -> bool:
        """Terminate exactly the owned Excel PID, reporting failures."""
        terminate = self._terminate or stage_two_watchdog._terminate_process
        try:
            terminate(self.excel_pid)
        except Exception as error:
            self._logger(
                "Abort requested but the NoteVal Excel process could not be "
                f"terminated; retrying: {error}"
            )
            return False
        self._termination_completed.set()
        self._stop_event.set()
        return True


def _log_watchdog_message(message: str) -> None:
    """Keep termination failures visible while heartbeats stay informational."""
    if message.startswith("Abort requested"):
        logger.warning(message)
    else:
        logger.info(message)


# ---------------------------------------------------------------------------
# Run (Decisions 127 and 128)
# ---------------------------------------------------------------------------


def _invoke_noteval(workbook: Any) -> None:
    """Call NoteVal's public Run macro once."""
    try:
        workbook.macro(NOTEVAL_MACRO)()
    except Exception as error:
        raise PropertyStageFourError(f"NoteVal Run macro failed: {error}") from error


def _close_owned_excel(application: Any) -> None:
    """Quit the owned Excel instance and terminate it if COM leaves it running."""
    process_id = int(application.pid)
    try:
        application.api.DisplayAlerts = False
        application.quit()
    except Exception as error:
        logger.error("Could not quit the owned NoteVal Excel instance: %s", error)
    gc.collect()
    try:
        is_running = any(
            process.get("pid") == process_id
            for process in stage_two_watchdog.ProcessProbe().process_snapshot()
        )
    except Exception as error:
        logger.error(
            "Could not verify closure of the owned NoteVal Excel process %s: %s",
            process_id,
            error,
        )
        return
    if is_running:
        try:
            stage_two_watchdog._terminate_process(process_id)
        except Exception as error:
            logger.error(
                "Could not terminate the owned NoteVal Excel process %s: %s",
                process_id,
                error,
            )


def run_property_stage_four(caller: Any) -> Path:
    """Run and verify NoteVal in a separate visible Excel instance, in place.

    Args:
        caller: The Draft workbook that launched the stage.

    Returns:
        The NoteVal workbook path, saved after verification.

    Raises:
        PropertyStageFourTimeoutError: If the user aborted at the time limit.
        PropertyStageFourError: If inputs, the macro, verification, saving or
            the restore of the prepared workbook failed.
    """
    caller_path = shared.resolve_caller_path(caller)
    with shared.workbook_run_guard(
        caller_path,
        conflict_message="A stress run is already in progress for this workbook",
    ):
        try:
            previous_success = _read_property_stage_four_fields(caller)[
                "last_successful_run"
            ]
        except Exception as error:
            raise PropertyStageFourError(
                f"Property Stage 4 audit panel is unavailable: {error}"
            ) from error
        started = datetime.now()
        clock = time.perf_counter()
        workbook_link: str | None = None
        application = workbook = None
        watchdog: NoteValRunWatchdog | None = None
        model_saved = False
        save_attempted = False
        backup_directory: tempfile.TemporaryDirectory[str] | None = None
        backup_path: Path | None = None
        try:
            inputs = resolve_property_stage_four_inputs(caller)
            workbook_link = stage_two._portable_workbook_path(
                caller_path, inputs.workbook_path
            )
            stage_two._ensure_workbook_not_open(
                inputs.workbook_path,
                caller,
                caller_path,
                "Property Stage 4",
                "NoteVal workbook",
            )
            write_property_stage_four_audit(
                caller,
                _audit_values(
                    started, clock, "Running", workbook_link, previous_success
                ),
            )
            try:
                time_limit = configured_noteval_time_limit(
                    os.environ.get(TIME_LIMIT_ENVIRONMENT),
                    bool(getattr(sys, "frozen", False)),
                )
            except ValueError as error:
                raise PropertyStageFourError(str(error)) from error

            import xlwings as xw

            before_mtime = inputs.workbook_path.stat().st_mtime_ns
            backup_directory = tempfile.TemporaryDirectory(
                prefix="property-stage-four-"
            )
            backup_path = Path(backup_directory.name) / inputs.workbook_path.name
            shutil.copy2(inputs.workbook_path, backup_path)
            application = xw.App(visible=True, add_book=False)
            application.api.AutomationSecurity = XL_AUTOMATION_SECURITY_LOW
            # Decision 128: use the Flash values Stage 3 cached and verified.
            workbook = application.books.open(
                str(inputs.workbook_path),
                update_links=False,
                read_only=False,
                ignore_read_only_recommended=True,
            )
            stage_two._disable_autosave(workbook)
            stage_two._require_expected_workbook(
                workbook,
                inputs.workbook_path,
                caller,
                caller_path,
                "NoteVal workbook",
            )
            if workbook.api.ReadOnly:
                raise PropertyStageFourError(
                    "Property Stage 3 NoteVal workbook is read-only"
                )
            invocation_time = stage_two._excel_now(application, "NoteVal")
            watchdog = NoteValRunWatchdog(
                excel_pid=int(application.pid),
                time_limit=time_limit,
                logger=_log_watchdog_message,
            )
            watchdog.start()
            macro_error: BaseException | None = None
            try:
                _invoke_noteval(workbook)
            except BaseException as error:
                macro_error = error
            finally:
                watchdog.stop()
                watchdog.join()
            if watchdog.aborted:
                raise PropertyStageFourTimeoutError(
                    "Property Stage 4 run timed out and was aborted by the user"
                ) from macro_error
            if macro_error is not None:
                raise macro_error
            verify_noteval_run(workbook, invocation_time)

            save_attempted = True
            workbook.save()
            if not workbook.api.Saved:
                raise PropertyStageFourError(
                    "NoteVal workbook did not report that it was saved"
                )
            if inputs.workbook_path.stat().st_mtime_ns <= before_mtime:
                raise PropertyStageFourError(
                    "NoteVal workbook modification time did not advance after saving"
                )
            model_saved = True
            finished = datetime.now()
            try:
                write_property_stage_four_audit(
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
                raise PropertyStageFourError(
                    "NoteVal workbook was verified and saved, but the final "
                    "Property Stage 4 success audit could not be saved; the "
                    f"previous Last Successful Run was retained: {audit_error}"
                ) from audit_error
            logger.info(
                "Property Stage 4 succeeded; ran NoteVal in place: %s",
                inputs.workbook_path,
            )
            return inputs.workbook_path
        except BaseException as error:
            if isinstance(error, PropertyStageFourTimeoutError):
                status = "Timed Out"
            elif isinstance(error, KeyboardInterrupt):
                status = "Interrupted"
            else:
                status = "Failed"
            try:
                _record_failure(
                    caller, started, clock, workbook_link, previous_success, status
                )
            except Exception as audit_error:
                if model_saved:
                    logger.error(
                        "NoteVal was saved, but its final audit could not be "
                        "persisted: %s",
                        audit_error,
                    )
                else:
                    logger.error(
                        "Property Stage 4 failed and its audit could not be "
                        "saved: %s",
                        audit_error,
                    )
            raise
        finally:
            process_terminated = watchdog is not None and getattr(
                watchdog, "process_terminated", watchdog.aborted
            )
            if workbook is not None and not process_terminated:
                try:
                    workbook.api.Close(SaveChanges=False)
                except Exception as cleanup_error:
                    logger.error(
                        "Could not close the owned NoteVal workbook: %s",
                        cleanup_error,
                    )
            workbook = None
            if application is not None and not process_terminated:
                _close_owned_excel(application)
            application = None
            restore_error: Exception | None = None
            # Restore on any failure that left the Stage 3 file changed on disk,
            # not only after our own save, so stray writes never persist.
            if (
                not model_saved
                and backup_path is not None
                and (
                    save_attempted
                    or stage_two._file_changed_since(inputs.workbook_path, before_mtime)
                )
            ):
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
                        "Could not remove the temporary Stage 4 backup: %s",
                        cleanup_error,
                    )
            if restore_error is not None:
                raise PropertyStageFourError(
                    "Property Stage 4 could not restore the Stage 3 NoteVal "
                    f"workbook after the run failed: {restore_error}"
                ) from restore_error


# ---------------------------------------------------------------------------
# Installer (Decision 131)
# ---------------------------------------------------------------------------


def _require_stage_three_installed(book: Any) -> None:
    """Require the Stage 3 panel, on which Stage 4's layout depends."""
    try:
        stage_three.require_property_stage_three_panel(book)
    except shared.BondsLiteError as error:
        raise PropertyStageFourError(
            f"Install the Property Stage 3 controls before Stage 4: {error}"
        ) from error


def _initialise_property_stage_four_panel(book: Any) -> None:
    """Install the fixed Property Stage 4 audit panel without overwriting cells."""
    _require_stage_three_installed(book)
    existing = stage_three._name_keys(book)
    expected = {STAGE_FOUR_PANEL_NAME, *STAGE_FOUR_FIELD_NAMES.values()}
    if existing.intersection(expected):
        if not expected.issubset(existing):
            raise PropertyStageFourError(
                "Property Stage 4 audit names are incomplete; installation was not changed"
            )
        require_property_stage_four_panel(book)
        return
    sheet = book.sheets["I. Control"]
    area = sheet.range(STAGE_FOUR_PANEL_RANGE)
    if (
        book.app.api.WorksheetFunction.CountA(area.api)
        or area.api.MergeCells is not False
    ):
        raise PropertyStageFourError(
            "Proposed Property Stage 4 audit area contains cells or merges"
        )
    proposed = (area.left, area.top, area.width, area.height)
    for shape in sheet.api.Shapes:
        shape_box = (
            float(shape.Left),
            float(shape.Top),
            float(shape.Width),
            float(shape.Height),
        )
        if stage_two._panel_overlap(shape_box, proposed):
            raise PropertyStageFourError(
                "Proposed Property Stage 4 audit area overlaps a shape or button"
            )
    row, column = area.row, area.column
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
        heading.value = STAGE_FOUR_PANEL_HEADING
        heading.color = (127, 127, 255)
        heading.font.color = (255, 255, 255)
        heading.font.bold = True
        book.names.add(STAGE_FOUR_PANEL_NAME, f"='I. Control'!{area.get_address()}")
        added.append(STAGE_FOUR_PANEL_NAME)
        for field, offset in STAGE_FOUR_OFFSETS.items():
            extra_row = 1 if field == "input_workbook" else 0
            label = sheet.range(
                (row + offset, column),
                (row + offset + extra_row, column + 3),
            )
            value = sheet.range(
                (row + offset, column + 4),
                (row + offset + extra_row, column + 7),
            )
            label.merge()
            value.merge()
            label.value = STAGE_FOUR_LABELS[field]
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
            name = STAGE_FOUR_FIELD_NAMES[field]
            book.names.add(name, f"='I. Control'!{value[0, 0].get_address()}")
            added.append(name)
    except Exception:
        for name in reversed(added):
            book.names[name].delete()
        raise
    finally:
        book.app.api.EnableEvents = events
    require_property_stage_four_panel(book)


def _install_property_stage_four_button(book: Any) -> None:
    """Install the Stage 4 button beside its panel, in the Stage 3 button column."""
    sheet = book.sheets["I. Control"]
    if STAGE_FOUR_BUTTON_NAME in [shape.Name for shape in sheet.api.Shapes]:
        button = sheet.api.Buttons().Item(STAGE_FOUR_BUTTON_NAME)
    else:
        try:
            stage_three_button = sheet.api.Shapes.Item(
                stage_three.STAGE_THREE_BUTTON_NAME
            )
        except Exception as error:
            raise PropertyStageFourError(
                "Cannot locate the Property Stage 3 button; install the Property "
                "Stage 3 controls before Stage 4"
            ) from error
        panel = book.names[STAGE_FOUR_PANEL_NAME].refers_to_range
        left = float(stage_three_button.Left)
        top = float(panel.top)
        width = float(stage_three_button.Width)
        height = float(stage_three_button.Height)
        row = panel.row
        column = int(stage_three_button.TopLeftCell.Column)
        end_column = column
        while sheet.range((row, end_column)).left < left + width:
            end_column += 1
        cells = sheet.range((row, column), (row + 2, end_column - 1))
        if (
            book.app.api.WorksheetFunction.CountA(cells.api)
            or cells.api.MergeCells is not False
        ):
            raise PropertyStageFourError(
                "Proposed Run NoteVal button covers occupied cells or merges"
            )
        proposed = (left, top, width, height)
        for shape in sheet.api.Shapes:
            shape_box = (
                float(shape.Left),
                float(shape.Top),
                float(shape.Width),
                float(shape.Height),
            )
            if stage_two._panel_overlap(shape_box, proposed):
                raise PropertyStageFourError(
                    "Proposed Run NoteVal button overlaps an existing shape"
                )
        button = sheet.api.Buttons().Add(left, top, width, height)
        button.Name = STAGE_FOUR_BUTTON_NAME
    button.Caption = STAGE_FOUR_BUTTON_CAPTION
    button.OnAction = (
        "'" + book.name.replace("'", "''") + "'!" + STAGE_FOUR_BUTTON_MACRO
    )


def _install_property_stage_four_macro(book: Any) -> None:
    """Add the P4 VBA entry point without altering the existing stage macros."""
    try:
        module = book.api.VBProject.VBComponents.Item("modBondsLiteLauncher")
        code = module.CodeModule
        source = code.Lines(1, code.CountOfLines) if code.CountOfLines else ""
    except Exception as error:
        raise PropertyStageFourError(
            "Cannot access modBondsLiteLauncher in the Draft VBA project"
        ) from error
    folded = source.casefold()
    if (
        'LaunchBondsLite "P4"'.casefold() in folded
        and "Public Sub RunNoteVal()".casefold() in folded
    ):
        return
    if "Private Sub LaunchBondsLite(ByVal stage As String)".casefold() not in folded:
        raise PropertyStageFourError(
            "Draft launcher does not contain a supported LaunchBondsLite signature"
        )
    if "Public Sub RunNoteVal()".casefold() in folded:
        raise PropertyStageFourError(
            "Draft RunNoteVal macro has an unexpected implementation"
        )
    code.InsertLines(
        code.CountOfLines + 1,
        '\nPublic Sub RunNoteVal()\n    LaunchBondsLite "P4"\nEnd Sub\n',
    )


def install_property_stage_four_controls(caller: Path) -> Path:
    """Install Stage 4's panel, button and macro into a closed Draft.

    Args:
        caller: Path to the closed Draft workbook.

    Returns:
        The backup copy taken before the Draft was changed.

    Raises:
        PropertyStageFourError: If the Draft is open, read-only, lacks the
            Stage 3 controls or the proposed layout is occupied.
    """
    import xlwings as xw

    caller = Path(caller).resolve(strict=True)
    for application in xw.apps:
        for book in application.books:
            if book.name.casefold() == caller.name.casefold():
                raise PropertyStageFourError(
                    f"Close {caller.name} in Excel before installation"
                )
    backup_folder = caller.parent / "references" / "backups"
    backup_folder.mkdir(parents=True, exist_ok=True)
    backup = backup_folder / f"{caller.stem}_{datetime.now():%Y%m%d_%H%M%S_%f}.xlsm"
    shutil.copy2(caller, backup)
    application = xw.App(visible=False, add_book=True)
    try:
        application.api.AutomationSecurity = XL_AUTOMATION_SECURITY_FORCE_DISABLE
        application.api.EnableEvents = False
        application.api.Calculation = -4135
        application.display_alerts = False
        book = application.books.open(str(caller), update_links=False)
        try:
            if book.api.ReadOnly:
                raise PropertyStageFourError("Draft workbook is read-only or locked")
            _require_stage_three_installed(book)
            _initialise_property_stage_four_panel(book)
            _install_property_stage_four_button(book)
            _install_property_stage_four_macro(book)
            book.save()
        finally:
            book.close()
    finally:
        application.api.DisplayAlerts = False
        application.api.Quit()
    return backup
