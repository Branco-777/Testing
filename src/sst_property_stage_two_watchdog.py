"""Monitor a Property Stage 2 LTM run without calling through its busy Excel."""

from __future__ import annotations

import ctypes
import logging
import math
import os
import queue
import re
import threading
import time
from ctypes import wintypes
from datetime import timedelta
from pathlib import PureWindowsPath
from typing import Any, Callable

logger = logging.getLogger(__name__)

LTM_PROGRESS_FIRST_NAME = "ProgressFirstCell"
XL_UP = -4162
HEARTBEAT_SECONDS = 5 * 60
STALL_SECONDS = 10 * 60
RUNNER_EXIT_WARNING_SECONDS = 2 * 60
POLL_SECONDS = 30
TIME_LIMIT_PER_SCENARIO_SECONDS = 10 * 60
MINIMUM_TIME_LIMIT_SECONDS = 30 * 60
CONTINUE_INTERVAL_SECONDS = 30 * 60
PROGRESS_READ_TIMEOUT_SECONDS = 5
RUNNER_EXE_PREFIX = "ltmflasharunmanager"
SCENARIO_MARKER = re.compile(r"Run (\d+) - End of run!")


def ltm_time_limit(scenarios: int) -> timedelta:
    """Return the normal run limit, with a 30-minute minimum."""
    if isinstance(scenarios, bool) or not isinstance(scenarios, int) or scenarios < 1:
        raise ValueError("Scenario count must be a positive whole number")
    return timedelta(
        seconds=max(
            MINIMUM_TIME_LIMIT_SECONDS,
            TIME_LIMIT_PER_SCENARIO_SECONDS * scenarios,
        )
    )


def configured_time_limit(
    scenarios: int, *, override: str | None, frozen: bool
) -> timedelta:
    """Apply a source-mode-only time-limit override for live acceptance tests."""
    limit = ltm_time_limit(scenarios)
    if frozen:
        if override is not None and override.strip():
            logger.warning(
                "Ignoring SST_P2_TIME_LIMIT_MINUTES in the packaged controller"
            )
        return limit
    if override is None or not override.strip():
        return limit
    try:
        minutes = float(override)
    except ValueError as error:
        raise ValueError(
            "SST_P2_TIME_LIMIT_MINUTES must be a positive finite number"
        ) from error
    if not math.isfinite(minutes) or minutes <= 0:
        raise ValueError("SST_P2_TIME_LIMIT_MINUTES must be a positive finite number")
    return timedelta(minutes=minutes)


def count_completed_scenarios(messages: list[str]) -> int:
    """Count distinct completed-scenario markers in LTM's progress log."""
    return len(
        {
            int(match.group(1))
            for message in messages
            if (match := SCENARIO_MARKER.fullmatch(message.strip())) is not None
        }
    )


def _descendants(
    snapshot: list[dict[str, object]], root_pid: int
) -> list[tuple[int, int]]:
    """Return descendant PIDs with their depth below ``root_pid``."""
    children: dict[int, list[int]] = {}
    for process in snapshot:
        pid = process.get("pid")
        parent_pid = process.get("ppid")
        if isinstance(pid, int) and isinstance(parent_pid, int):
            children.setdefault(parent_pid, []).append(pid)
    found: list[tuple[int, int]] = []
    pending = [(root_pid, 0)]
    while pending:
        parent_pid, depth = pending.pop(0)
        for child_pid in children.get(parent_pid, []):
            found.append((child_pid, depth + 1))
            pending.append((child_pid, depth + 1))
    return found


def _find_runner_under_excel_pid(
    snapshot: list[dict[str, object]], excel_pid: int
) -> int | None:
    """Find the LTM runner only if it descends from the owned Excel process."""
    by_pid = {
        process["pid"]: process
        for process in snapshot
        if isinstance(process.get("pid"), int)
    }
    for pid, _depth in _descendants(snapshot, excel_pid):
        process = by_pid.get(pid, {})
        executable = str(process.get("exe", ""))
        if PureWindowsPath(executable).name.casefold().startswith(RUNNER_EXE_PREFIX):
            return pid
    return None


class ProcessProbe:
    """Inspect only the LTM process tree below Stage 2's owned Excel PID."""

    def process_snapshot(self) -> list[dict[str, object]]:
        """Return PID, parent PID and executable name for live Windows processes."""
        if os.name != "nt":
            raise OSError("LTM process monitoring is available only on Windows")

        class ProcessEntry32(ctypes.Structure):
            _fields_ = [
                ("dwSize", wintypes.DWORD),
                ("cntUsage", wintypes.DWORD),
                ("th32ProcessID", wintypes.DWORD),
                ("th32DefaultHeapID", ctypes.c_size_t),
                ("th32ModuleID", wintypes.DWORD),
                ("cntThreads", wintypes.DWORD),
                ("th32ParentProcessID", wintypes.DWORD),
                ("pcPriClassBase", wintypes.LONG),
                ("dwFlags", wintypes.DWORD),
                ("szExeFile", wintypes.WCHAR * 260),
            ]

        kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
        create_snapshot = kernel32.CreateToolhelp32Snapshot
        create_snapshot.argtypes = [wintypes.DWORD, wintypes.DWORD]
        create_snapshot.restype = wintypes.HANDLE
        process_first = kernel32.Process32FirstW
        process_first.argtypes = [
            wintypes.HANDLE,
            ctypes.POINTER(ProcessEntry32),
        ]
        process_first.restype = wintypes.BOOL
        process_next = kernel32.Process32NextW
        process_next.argtypes = [
            wintypes.HANDLE,
            ctypes.POINTER(ProcessEntry32),
        ]
        process_next.restype = wintypes.BOOL
        close_handle = kernel32.CloseHandle
        close_handle.argtypes = [wintypes.HANDLE]
        close_handle.restype = wintypes.BOOL

        snapshot_handle = create_snapshot(0x00000002, 0)
        if snapshot_handle == wintypes.HANDLE(-1).value:
            raise ctypes.WinError(ctypes.get_last_error())
        processes: list[dict[str, object]] = []
        try:
            entry = ProcessEntry32()
            entry.dwSize = ctypes.sizeof(ProcessEntry32)
            succeeded = process_first(snapshot_handle, ctypes.byref(entry))
            while succeeded:
                processes.append(
                    {
                        "pid": int(entry.th32ProcessID),
                        "ppid": int(entry.th32ParentProcessID),
                        "exe": str(entry.szExeFile),
                    }
                )
                succeeded = process_next(snapshot_handle, ctypes.byref(entry))
        finally:
            close_handle(snapshot_handle)
        return processes

    @staticmethod
    def find_runner_under_excel_pid(
        snapshot: list[dict[str, object]], excel_pid: int
    ) -> int | None:
        """Find the LTM runner only if it descends from the owned Excel process."""
        return _find_runner_under_excel_pid(snapshot, excel_pid)

    def find_runner_under_excel(self, excel_pid: int) -> int | None:
        """Find the LTM runner in the owned Excel process tree."""
        return self.find_runner_under_excel_pid(self.process_snapshot(), excel_pid)

    def runner_cpu_time(self, excel_pid: int) -> float | None:
        """Return cumulative CPU seconds used by the runner and its descendants."""
        snapshot = self.process_snapshot()
        runner_pid = self.find_runner_under_excel_pid(snapshot, excel_pid)
        if runner_pid is None:
            return None
        process_pids = [runner_pid] + [
            pid for pid, _depth in _descendants(snapshot, runner_pid)
        ]
        return sum(
            self._process_cpu_time(pid)
            for pid in process_pids
            if self._process_exists(pid)
        )

    @staticmethod
    def _process_exists(pid: int) -> bool:
        """Check that a process remains available for a CPU-time query."""
        import pywintypes
        import win32api
        import win32con

        try:
            handle = win32api.OpenProcess(
                win32con.PROCESS_QUERY_INFORMATION, False, pid
            )
        except pywintypes.error as error:
            if error.winerror in {6, 87}:
                return False
            raise
        win32api.CloseHandle(handle)
        return True

    @staticmethod
    def _process_cpu_time(pid: int) -> float:
        """Read cumulative process CPU time in seconds."""
        import pywintypes
        import win32api
        import win32con
        import win32process

        handle = win32api.OpenProcess(win32con.PROCESS_QUERY_INFORMATION, False, pid)
        try:
            times = win32process.GetProcessTimes(handle)
            return (times["KernelTime"] + times["UserTime"]) / 10_000_000
        except pywintypes.error as error:
            if error.winerror in {6, 87}:
                return 0.0
            raise
        finally:
            win32api.CloseHandle(handle)


class ProgressReader:
    """Read LTM scenario progress without blocking the run watchdog."""

    def __init__(
        self,
        excel_pid: int,
        workbook_name: str,
        timeout_seconds: float = PROGRESS_READ_TIMEOUT_SECONDS,
        warning_logger: Callable[[str], None] | None = None,
    ) -> None:
        if not math.isfinite(timeout_seconds) or timeout_seconds <= 0:
            raise ValueError("Progress read timeout must be a positive finite duration")
        self.excel_pid = excel_pid
        self.workbook_name = workbook_name
        self.timeout_seconds = timeout_seconds
        self._warned = False
        self._disabled = False
        self._read_lock = threading.Lock()
        self._warning_logger = warning_logger or logger.warning

    def read(self) -> list[str] | None:
        """Return progress or give up after a bounded read-only COM request."""
        with self._read_lock:
            if self._disabled:
                return None
            result_queue: queue.Queue[list[str] | None] = queue.Queue(maxsize=1)
            worker = threading.Thread(
                target=self._read_worker,
                args=(result_queue,),
                name="PropertyStageTwoProgressReader",
                daemon=True,
            )
            worker.start()
            try:
                return result_queue.get(timeout=self.timeout_seconds)
            except queue.Empty:
                self._disabled = True
                self._warn_once(
                    "LTM progress query did not return within "
                    f"{self.timeout_seconds:g} seconds; scenario updates are disabled"
                )
                return None

    def _read_worker(self, result_queue: queue.Queue[list[str] | None]) -> None:
        """Perform one COM read on a worker that cannot block the watchdog."""
        try:
            result_queue.put(self._read_excel_progress())
        except Exception as error:
            self._warn_once(f"Cannot read LTM scenario progress: {error}")
            result_queue.put(None)

    def _read_excel_progress(self) -> list[str] | None:
        """Read the progress range from the owned LTM workbook."""
        import pythoncom
        import pywintypes
        import xlwings as xw

        initialized = False
        try:
            pythoncom.CoInitialize()
            initialized = True
            application = xw.apps[self.excel_pid]
            workbook = application.books[self.workbook_name]
            first = workbook.names[LTM_PROGRESS_FIRST_NAME].refers_to_range
            row, column = int(first.row), int(first.column)
            sheet = first.sheet
            last = int(sheet.api.Cells(sheet.api.Rows.Count, column).End(XL_UP).Row)
            if last < row:
                return []
            value = sheet.range((row, column), (last, column)).value
            return [
                str(item).strip()
                for item in _grid_values(value)
                if item is not None and str(item).strip()
            ]
        except (
            pywintypes.com_error,
            RuntimeError,
            KeyError,
            AttributeError,
            OSError,
            TypeError,
            ValueError,
        ) as error:
            self._warn_once(
                f"Cannot read LTM scenario progress while the macro is running: {error}"
            )
            return None
        finally:
            if initialized:
                uninitialise = getattr(pythoncom, "CoUninitialize", None)
                if uninitialise is not None:
                    uninitialise()

    def _warn_once(self, message: str) -> None:
        """Log at most one progress-reader warning per LTM run."""
        if not self._warned:
            self._warning_logger(message)
            self._warned = True


def _grid_values(value: object) -> list[object]:
    """Flatten Excel's scalar and two-dimensional range values."""
    if isinstance(value, (tuple, list)):
        return [item for child in value for item in _grid_values(child)]
    return [value]


def _terminate_process(pid: int) -> None:
    """Terminate one exact process ID and wait briefly for it to exit."""
    import pywintypes
    import win32api
    import win32con
    import win32event

    try:
        handle = win32api.OpenProcess(
            win32con.PROCESS_TERMINATE | win32con.SYNCHRONIZE, False, pid
        )
    except pywintypes.error as error:
        if error.winerror in {6, 87}:
            return
        raise
    try:
        win32api.TerminateProcess(handle, 1)
        wait_result = win32event.WaitForSingleObject(handle, 10_000)
        if wait_result != win32event.WAIT_OBJECT_0:
            raise TimeoutError(f"Process {pid} did not exit after termination")
    finally:
        win32api.CloseHandle(handle)


def terminate_owned_run(excel_pid: int) -> None:
    """Terminate the owned LTM runner tree, then Stage 2's Excel process."""
    probe = ProcessProbe()
    snapshot = probe.process_snapshot()
    runner_pid = _find_runner_under_excel_pid(snapshot, excel_pid)
    if runner_pid is not None:
        parent_by_pid = {
            int(process["pid"]): int(process["ppid"])
            for process in snapshot
            if isinstance(process.get("pid"), int)
            and isinstance(process.get("ppid"), int)
        }
        runner_descendants = sorted(
            _descendants(snapshot, runner_pid),
            key=lambda process: process[1],
            reverse=True,
        )
        process_pids = [pid for pid, _depth in runner_descendants]
        process_pids.append(runner_pid)
        parent_pid = parent_by_pid.get(runner_pid)
        while parent_pid is not None and parent_pid != excel_pid:
            process_pids.append(parent_pid)
            parent_pid = parent_by_pid.get(parent_pid)
        if parent_pid != excel_pid:
            raise RuntimeError(
                f"LTM runner {runner_pid} is no longer below Excel {excel_pid}"
            )
        for pid in process_pids:
            try:
                _terminate_process(pid)
            except Exception as error:
                raise RuntimeError(
                    f"Could not terminate LTM process {pid}: {error}"
                ) from error
    _terminate_process(excel_pid)


def _format_elapsed(seconds: float) -> str:
    """Format elapsed time as hours and minutes."""
    total_minutes = int(seconds // 60)
    return f"{total_minutes // 60}:{total_minutes % 60:02d}"


class LtmRunWatchdog(threading.Thread):
    """Report progress and offer a user-controlled timeout for one LTM run."""

    def __init__(
        self,
        *,
        excel_pid: int,
        workbook_name: str,
        scenarios: int,
        probe: Any | None = None,
        reader: Any | None = None,
        prompt: Callable[[str], str] | None = None,
        logger: Callable[[str], None] = print,
        clock: Any = time,
        time_limit: timedelta | None = None,
    ) -> None:
        super().__init__(name="PropertyStageTwoLtmWatchdog", daemon=True)
        self.excel_pid = excel_pid
        self.workbook_name = workbook_name
        self.scenarios = scenarios
        self.probe = probe or ProcessProbe()
        self.reader = reader or ProgressReader(
            excel_pid, workbook_name, warning_logger=logger
        )
        self._prompt = prompt
        self._logger = logger
        self._clock = clock
        self._sleep = getattr(clock, "sleep", time.sleep)
        self._time_limit = time_limit or ltm_time_limit(scenarios)
        self._stop_event = threading.Event()
        self._aborted_event = threading.Event()
        self._termination_retry_at: float | None = None
        self._termination_completed = threading.Event()

    @property
    def aborted(self) -> bool:
        """Whether the user selected abort at the time limit."""
        return self._aborted_event.is_set()

    @property
    def process_terminated(self) -> bool:
        """Whether the owned runner tree and Excel process were terminated."""
        return self._termination_completed.is_set()

    def stop(self) -> None:
        """Stop the watcher and release any cancellable console prompt."""
        self._stop_event.set()

    def run(self) -> None:
        """Poll process activity and emit periodic status until stopped."""
        started = self._clock.monotonic()
        next_heartbeat = started + HEARTBEAT_SECONDS
        next_limit = started + self._time_limit.total_seconds()
        last_cpu_time: float | None = None
        last_activity = started
        runner_seen = False
        runner_missing_logged = False
        runner_exit_at: float | None = None
        dialog_warning_logged = False
        probe_warning_logged = False
        cpu_warning_logged = False
        stall_warning_logged = False
        while not self._stop_event.is_set():
            now = self._clock.monotonic()
            try:
                runner_pid = self.probe.find_runner_under_excel(self.excel_pid)
            except Exception as error:
                runner_pid = None
                if not probe_warning_logged:
                    self._logger(
                        f"Cannot inspect the LTM process tree; continuing with elapsed time only: {error}"
                    )
                    probe_warning_logged = True
            if runner_pid is None:
                if not runner_missing_logged:
                    self._logger(
                        "LTM runner was not found; continuing with elapsed time only"
                    )
                    runner_missing_logged = True
                last_cpu_time = None
                last_activity = now
                stall_warning_logged = False
                if runner_seen and runner_exit_at is None:
                    runner_exit_at = now
                if (
                    runner_exit_at is not None
                    and not dialog_warning_logged
                    and now - runner_exit_at >= RUNNER_EXIT_WARNING_SECONDS
                ):
                    self._logger(
                        "The LTM runner has exited but Excel is still waiting; "
                        "check the LTM Excel window for an error dialog"
                    )
                    dialog_warning_logged = True
            else:
                runner_seen = True
                runner_exit_at = None
                dialog_warning_logged = False
                try:
                    cpu_time = self.probe.runner_cpu_time(self.excel_pid)
                except Exception as error:
                    cpu_time = None
                    if not cpu_warning_logged:
                        self._logger(
                            f"Cannot inspect LTM runner activity; stall detection is unavailable: {error}"
                        )
                        cpu_warning_logged = True
                if cpu_time is None:
                    if not cpu_warning_logged:
                        self._logger(
                            "LTM runner activity is unavailable; stall detection is disabled"
                        )
                        cpu_warning_logged = True
                    last_cpu_time = None
                    last_activity = now
                    stall_warning_logged = False
                elif last_cpu_time is None or cpu_time > last_cpu_time:
                    last_cpu_time = cpu_time
                    last_activity = now
                    stall_warning_logged = False
                elif not stall_warning_logged and now - last_activity >= STALL_SECONDS:
                    self._logger(
                        "Potential LTM runner stall: no CPU activity for 10 minutes; "
                        "the run is continuing"
                    )
                    stall_warning_logged = True

            if now >= next_heartbeat:
                try:
                    messages = self.reader.read()
                except Exception as error:
                    messages = None
                    self._logger(
                        f"Cannot read LTM scenario progress; continuing without it: {error}"
                    )
                progress = ""
                if messages is not None:
                    completed = count_completed_scenarios(messages)
                    progress = f" | Scenario {completed} of {self.scenarios} complete"
                self._logger(
                    f"LTM running | elapsed {_format_elapsed(now - started)}{progress}"
                )
                while next_heartbeat <= now:
                    next_heartbeat += HEARTBEAT_SECONDS

            if self._aborted_event.is_set() and self._termination_retry_at is not None:
                if now >= self._termination_retry_at:
                    if self._attempt_termination():
                        self._termination_retry_at = None
                    else:
                        self._termination_retry_at = now + POLL_SECONDS

            if now >= next_limit and not self._aborted_event.is_set():
                response = self._ask_to_continue()
                if response == "A":
                    self._aborted_event.set()
                    if not self._attempt_termination():
                        self._termination_retry_at = now + POLL_SECONDS
                else:
                    next_limit = self._clock.monotonic() + CONTINUE_INTERVAL_SECONDS
                    self._logger(
                        "LTM run will continue; the next time-limit prompt is in 30 minutes"
                    )

            if self._clock is time:
                self._stop_event.wait(POLL_SECONDS)
            else:
                self._sleep(POLL_SECONDS)

    def _ask_to_continue(self) -> str:
        """Ask whether to wait another interval or abort the LTM run."""
        message = (
            "The Property Stage 2 run has reached its time limit. "
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
        """Read a cancellable W/A response from the packaged Windows console."""
        import msvcrt

        self._logger(message)
        self._logger("Enter W or A, then press Enter")
        response: list[str] = []
        while not self._stop_event.is_set():
            if msvcrt.kbhit():
                character = msvcrt.getwch()
                if character in {"\r", "\n"}:
                    self._logger("")
                    return "".join(response)
                if character == "\b":
                    if response:
                        response.pop()
                elif character.isprintable():
                    response.append(character)
            self._stop_event.wait(0.1)
        return ""

    def _attempt_termination(self) -> bool:
        """Terminate the exact runner tree and Excel PID, reporting failures."""
        try:
            terminate_owned_run(self.excel_pid)
        except Exception as error:
            self._logger(
                f"Abort requested but process termination was incomplete; "
                f"retrying: {error}"
            )
            return False
        self._termination_completed.set()
        self._stop_event.set()
        return True
