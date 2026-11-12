"""Behavioural contracts for the Property Stage 4 NoteVal run."""

from __future__ import annotations

import contextlib
import os
import re
import sys
import threading
import types
from datetime import datetime, timedelta
from pathlib import Path
from types import SimpleNamespace
from typing import Any

import pytest

from bonds_lite_input import BondsLiteError

FLASH_LINK = "SST Output_Test/Property/JRL_LTM_FlashModel_Test.xlsb"
NOTEVAL_LINK = "SST Output_Test/Property/JRL_LTM Note Valuation Tool v1.40_Test.xlsb"
RESIDUAL_CELLS = (
    ("O. MAC MI", ("D12",)),
    (
        "C. Note Valuation ex Hedge",
        ("F12", "H12", "AB31:AM31", "AF45", "AQ33:AQ36", "AB57:AM57"),
    ),
    ("C. Note Valuation", ("T23", "T50", "T56")),
)


def _stage_four() -> Any:
    """Import lazily so the contracts collect before Stage 4 is implemented."""
    import sst_property_stage_four

    return sst_property_stage_four


def _name(value: object) -> Any:
    return SimpleNamespace(
        refers_to_range=SimpleNamespace(
            value=value,
            size=1,
            api=SimpleNamespace(Areas=SimpleNamespace(Count=1)),
        )
    )


# ---------------------------------------------------------------------------
# T1-T5: input resolution from the Stage 2 and Stage 3 audits
# ---------------------------------------------------------------------------


def _resolver(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    *,
    flag: bool = True,
    stage_three_status: str = "Succeeded",
    stage_three_output: str = NOTEVAL_LINK,
    stage_three_input: str = FLASH_LINK,
    stage_two_input: str = FLASH_LINK,
    stage_two_run: object = datetime(2026, 11, 11, 11),
    stage_three_finish: object = datetime(2026, 11, 11, 12),
) -> tuple[Any, Path, Path]:
    """Patch the audit readers so resolution sees only the supplied values."""
    stage_four = _stage_four()
    draft_path = tmp_path / "Draft.xlsm"
    folder = tmp_path / "SST Output_Test" / "Property"
    folder.mkdir(parents=True)
    flash_path = folder / "JRL_LTM_FlashModel_Test.xlsb"
    flash_path.write_bytes(b"flash")
    noteval_path = folder / "JRL_LTM Note Valuation Tool v1.40_Test.xlsb"
    noteval_path.write_bytes(b"noteval")
    monkeypatch.setattr(stage_four.shared, "resolve_caller_path", lambda _: draft_path)
    monkeypatch.setattr(stage_four.stage_one, "read_property_flag", lambda _: flag)
    monkeypatch.setattr(
        stage_four.stage_two, "require_property_stage_two_panel", lambda _: None
    )
    monkeypatch.setattr(
        stage_four.stage_three, "require_property_stage_three_panel", lambda _: None
    )
    monkeypatch.setattr(
        stage_four.stage_two,
        "_read_property_stage_two_fields",
        lambda _: {
            "status": "Succeeded",
            "input_workbook": stage_two_input,
            "last_successful_run": stage_two_run,
        },
    )
    monkeypatch.setattr(
        stage_four.stage_three,
        "_read_property_stage_three_fields",
        lambda _: {
            "status": stage_three_status,
            "finish": stage_three_finish,
            "input_workbook": stage_three_input,
            "last_successful_output": stage_three_output,
        },
    )
    return stage_four, flash_path, noteval_path


def test_resolve_accepts_current_stage_three_handoff(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    stage_four, _flash, noteval = _resolver(tmp_path, monkeypatch)

    inputs = stage_four.resolve_property_stage_four_inputs(SimpleNamespace())

    assert inputs.caller_path == tmp_path / "Draft.xlsm"
    assert inputs.workbook_path == noteval.resolve()
    assert inputs.flash_link == FLASH_LINK


def test_resolve_requires_property_selected(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    stage_four, _flash, _noteval = _resolver(tmp_path, monkeypatch, flag=False)

    with pytest.raises(
        stage_four.PropertyStageFourError, match="Property stress is not selected"
    ):
        stage_four.resolve_property_stage_four_inputs(SimpleNamespace())


@pytest.mark.parametrize("status", ["Failed", "Running", "Not Run", None])
def test_resolve_rejects_stage_three_that_did_not_succeed(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, status: object
) -> None:
    stage_four, _flash, _noteval = _resolver(
        tmp_path, monkeypatch, stage_three_status=status
    )

    with pytest.raises(
        stage_four.PropertyStageFourError,
        match="Property Stage 3 status Succeeded",
    ):
        stage_four.resolve_property_stage_four_inputs(SimpleNamespace())


@pytest.mark.parametrize("output", ["SST Output_Test/Property/missing.xlsb", None, ""])
def test_resolve_rejects_missing_stage_three_output(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, output: object
) -> None:
    stage_four, _flash, _noteval = _resolver(
        tmp_path, monkeypatch, stage_three_output=output
    )

    with pytest.raises(stage_four.PropertyStageFourError, match="Property Stage 3"):
        stage_four.resolve_property_stage_four_inputs(SimpleNamespace())


def test_resolve_rejects_stage_three_input_not_matching_stage_two(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    other = "SST Output_Test/Property/other_flash.xlsb"
    stage_four, flash, _noteval = _resolver(
        tmp_path, monkeypatch, stage_three_input=other
    )
    (flash.parent / "other_flash.xlsb").write_bytes(b"other flash")

    with pytest.raises(stage_four.PropertyStageFourError, match="does not match"):
        stage_four.resolve_property_stage_four_inputs(SimpleNamespace())


def test_resolve_rejects_stage_two_rerun_after_stage_three(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    stage_four, _flash, _noteval = _resolver(
        tmp_path,
        monkeypatch,
        stage_two_run=datetime(2026, 11, 11, 13),
        stage_three_finish=datetime(2026, 11, 11, 12),
    )

    with pytest.raises(
        stage_four.PropertyStageFourError, match="rerun Property Stage 3"
    ):
        stage_four.resolve_property_stage_four_inputs(SimpleNamespace())


def test_resolve_accepts_stage_two_run_equal_to_stage_three_finish(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    moment = datetime(2026, 11, 11, 12)
    stage_four, _flash, noteval = _resolver(
        tmp_path, monkeypatch, stage_two_run=moment, stage_three_finish=moment
    )

    inputs = stage_four.resolve_property_stage_four_inputs(SimpleNamespace())

    assert inputs.workbook_path == noteval.resolve()


def test_resolve_rejects_missing_stage_three_finish(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    stage_four, _flash, _noteval = _resolver(
        tmp_path, monkeypatch, stage_three_finish=None
    )

    with pytest.raises(stage_four.PropertyStageFourError, match="Property Stage 3"):
        stage_four.resolve_property_stage_four_inputs(SimpleNamespace())


# ---------------------------------------------------------------------------
# T6-T9: completion checks after NoteVal's Run macro
# ---------------------------------------------------------------------------

INVOCATION = datetime(2026, 11, 11, 21, 0)


def _cell_count(address: str) -> int:
    """Count cells in a single-row or single-column A1 address."""
    if ":" not in address:
        return 1
    first, last = address.split(":")
    first_column, first_row = re.fullmatch(r"([A-Z]+)(\d+)", first).groups()
    last_column, last_row = re.fullmatch(r"([A-Z]+)(\d+)", last).groups()

    def number(letters: str) -> int:
        value = 0
        for letter in letters:
            value = value * 26 + ord(letter) - 64
        return value

    return (number(last_column) - number(first_column) + 1) * (
        int(last_row) - int(first_row) + 1
    )


def _noteval_book(
    *,
    start: object = INVOCATION + timedelta(seconds=1),
    end: object = INVOCATION + timedelta(minutes=9),
    residual: object = 0.25,
    overrides: dict[tuple[str, str], object] | None = None,
) -> Any:
    """Build the named timestamps and residual cells that verification reads."""
    values: dict[tuple[str, str], object] = {}
    for sheet, addresses in RESIDUAL_CELLS:
        for address in addresses:
            count = _cell_count(address)
            values[(sheet, address)] = residual if count == 1 else [residual] * count
    values.update(overrides or {})
    sheets = {
        sheet: SimpleNamespace(
            name=sheet,
            range=lambda address, sheet=sheet: SimpleNamespace(
                value=values[(sheet, address)]
            ),
        )
        for sheet, _addresses in RESIDUAL_CELLS
    }
    return SimpleNamespace(
        names={"I_StartTime_1": _name(start), "I_EndTime_1": _name(end)},
        sheets=sheets,
    )


def test_residual_cells_match_noteval_goal_seek_targets() -> None:
    assert _stage_four().NOTEVAL_RESIDUAL_CELLS == RESIDUAL_CELLS


@pytest.mark.parametrize(
    "start",
    [
        INVOCATION + timedelta(seconds=1),
        INVOCATION,
        INVOCATION - timedelta(seconds=2),
    ],
)
def test_verify_accepts_fresh_times_and_finite_residuals(start: datetime) -> None:
    stage_four = _stage_four()
    book = _noteval_book(start=start, residual=-0.5)

    residuals = stage_four.verify_noteval_run(book, INVOCATION)

    assert residuals == {
        "O. MAC MI": 0.5,
        "C. Note Valuation ex Hedge": 0.5,
        "C. Note Valuation": 0.5,
    }


def test_verify_logs_residuals_for_review(monkeypatch: pytest.MonkeyPatch) -> None:
    stage_four = _stage_four()
    messages: list[str] = []
    monkeypatch.setattr(
        stage_four.logger,
        "info",
        lambda message, *args: messages.append(message % args),
    )

    stage_four.verify_noteval_run(_noteval_book(residual=3.0), INVOCATION)

    text = "\n".join(messages)
    assert "C. Note Valuation ex Hedge" in text
    assert "residual" in text.casefold()
    assert "3" in text


@pytest.mark.parametrize(
    ("start", "message"),
    [
        (INVOCATION - timedelta(seconds=3), "stale"),
        (INVOCATION - timedelta(hours=1), "stale"),
        (None, "NoteVal Run did not record"),
        ("not a time", "NoteVal Run did not record"),
    ],
)
def test_verify_rejects_stale_or_missing_start(start: object, message: str) -> None:
    stage_four = _stage_four()

    with pytest.raises(stage_four.PropertyStageFourError, match=message):
        stage_four.verify_noteval_run(_noteval_book(start=start), INVOCATION)


@pytest.mark.parametrize("end", [INVOCATION, None, INVOCATION - timedelta(days=1)])
def test_verify_rejects_end_before_start_or_missing(end: object) -> None:
    stage_four = _stage_four()

    with pytest.raises(stage_four.PropertyStageFourError, match="I_EndTime_1"):
        stage_four.verify_noteval_run(
            _noteval_book(start=INVOCATION + timedelta(seconds=5), end=end),
            INVOCATION,
        )


@pytest.mark.parametrize(
    ("sheet", "address", "value"),
    [
        ("O. MAC MI", "D12", None),
        ("O. MAC MI", "D12", "#N/A"),
        ("C. Note Valuation", "T50", True),
        ("C. Note Valuation", "T23", float("nan")),
        ("C. Note Valuation ex Hedge", "AF45", float("inf")),
        ("C. Note Valuation ex Hedge", "AB31:AM31", [0.0] * 11 + [None]),
        ("C. Note Valuation ex Hedge", "AQ33:AQ36", [[1.0], [2.0], ["#DIV/0!"], [4.0]]),
    ],
)
def test_verify_rejects_error_or_non_numeric_residual(
    sheet: str, address: str, value: object
) -> None:
    stage_four = _stage_four()
    book = _noteval_book(overrides={(sheet, address): value})

    with pytest.raises(stage_four.PropertyStageFourError) as raised:
        stage_four.verify_noteval_run(book, INVOCATION)

    assert f"{sheet}!{address}" in str(raised.value)


# ---------------------------------------------------------------------------
# T10-T13: run sequence through fake Excel
# ---------------------------------------------------------------------------


class _RecordingWatchdog:
    """Record the watchdog lifecycle without starting a thread."""

    def __init__(
        self,
        events: list[tuple[Any, ...]],
        *,
        aborted: bool = False,
        on_start: Any = None,
        **kwargs: Any,
    ) -> None:
        self.events = events
        self.kwargs = kwargs
        self.aborted = aborted
        self.process_terminated = aborted
        self.on_start = on_start

    def start(self) -> None:
        self.events.append(("watchdog", "start"))
        if self.on_start is not None:
            self.on_start()

    def stop(self) -> None:
        self.events.append(("watchdog", "stop"))

    def join(self) -> None:
        self.events.append(("watchdog", "join"))


def _run_harness(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    *,
    macro_error: BaseException | None = None,
    complete_run: bool = True,
    saved: bool = True,
    macro_overwrites_workbook: bool = False,
    macro_release: threading.Event | None = None,
    watchdog_aborted: bool = False,
) -> tuple[Any, list[tuple[Any, ...]], list[dict[str, object]], Any, Path, Any]:
    """Run Stage 4 through fake Excel while recording every external action."""
    stage_four = _stage_four()
    target = tmp_path / "JRL_LTM Note Valuation Tool v1.40_Test.xlsb"
    target.write_bytes(b"prepared noteval")
    original_mtime = target.stat().st_mtime_ns
    events: list[tuple[Any, ...]] = []
    audits: list[dict[str, object]] = []
    workbook = _noteval_book(
        start=datetime(2026, 11, 11, 9), end=datetime(2026, 11, 11, 9, 9)
    )
    workbook.name = target.name
    workbook.fullname = str(target)
    workbook.api = SimpleNamespace(
        ReadOnly=False,
        Saved=False,
        AutoSaveOn=True,
        FullName=str(target),
        Close=lambda **kwargs: events.append(("close_workbook", kwargs)),
    )

    def macro(name: str) -> Any:
        def invoke(*arguments: Any) -> None:
            events.append(("macro", name, arguments))
            events.append(("autosave_during_macro", workbook.api.AutoSaveOn))
            if macro_overwrites_workbook:
                target.write_bytes(b"partial noteval")
                os.utime(
                    target,
                    ns=(target.stat().st_atime_ns, original_mtime + 2_000_000_000),
                )
            if macro_release is not None:
                macro_release.wait(timeout=2)
            if macro_error is not None:
                raise macro_error
            if not complete_run:
                return
            workbook.names["I_StartTime_1"].refers_to_range.value = (
                INVOCATION + timedelta(seconds=1)
            )
            workbook.names["I_EndTime_1"].refers_to_range.value = (
                INVOCATION + timedelta(minutes=9)
            )

        return invoke

    def save() -> None:
        events.append(("save",))
        workbook.api.Saved = saved
        target.write_bytes(b"noteval run output")
        os.utime(
            target,
            ns=(target.stat().st_atime_ns, target.stat().st_mtime_ns + 2_000_000_000),
        )

    workbook.macro = macro
    workbook.save = save

    class FakeApplication:
        def __init__(self) -> None:
            self.pid = 4321
            self.api = SimpleNamespace(
                AutomationSecurity=None,
                DisplayAlerts=True,
                Evaluate=lambda formula: (
                    events.append(("now",)) or INVOCATION
                    if formula == "NOW()"
                    else pytest.fail(formula)
                ),
            )
            self.books = SimpleNamespace(
                open=lambda path, **kwargs: (
                    events.append(("open", path, kwargs)),
                    workbook,
                )[1]
            )

        def quit(self) -> None:
            events.append(("quit",))

    application = FakeApplication()
    import xlwings

    monkeypatch.setattr(
        xlwings,
        "App",
        lambda **kwargs: (events.append(("app", kwargs)), application)[1],
    )
    monkeypatch.setattr(
        stage_four.shared,
        "workbook_run_guard",
        lambda *_args, **_kwargs: contextlib.nullcontext(),
    )
    monkeypatch.setattr(
        stage_four,
        "_read_property_stage_four_fields",
        lambda _caller: {"last_successful_run": datetime(2026, 10, 1)},
    )
    monkeypatch.setattr(
        stage_four,
        "resolve_property_stage_four_inputs",
        lambda _caller: stage_four.PropertyStageFourInputs(
            tmp_path / "Draft.xlsm", target, FLASH_LINK
        ),
    )
    monkeypatch.setattr(
        stage_four.stage_two, "_ensure_workbook_not_open", lambda *_args: None
    )
    monkeypatch.setattr(
        stage_four,
        "write_property_stage_four_audit",
        lambda _caller, fields: (
            audits.append(fields.copy()),
            events.append(("audit", fields["status"])),
        ),
    )
    monkeypatch.delenv("SST_P4_TIME_LIMIT_MINUTES", raising=False)
    watchdogs: list[_RecordingWatchdog] = []

    def make_watchdog(**kwargs: Any) -> _RecordingWatchdog:
        on_start = (
            (lambda: threading.Timer(0.01, macro_release.set).start())
            if macro_release is not None
            else None
        )
        watchdog = _RecordingWatchdog(
            events, aborted=watchdog_aborted, on_start=on_start, **kwargs
        )
        watchdogs.append(watchdog)
        return watchdog

    monkeypatch.setattr(stage_four, "NoteValRunWatchdog", make_watchdog)
    caller = SimpleNamespace(
        fullname=str(tmp_path / "Draft.xlsm"),
        api=SimpleNamespace(FullName=str(tmp_path / "Draft.xlsm")),
    )
    harness = SimpleNamespace(
        workbook=workbook, application=application, watchdogs=watchdogs
    )
    return stage_four, events, audits, caller, target, harness


def _index(events: list[tuple[Any, ...]], event: tuple[Any, ...]) -> int:
    return events.index(event)


def _macro_index(events: list[tuple[Any, ...]]) -> int:
    return next(index for index, event in enumerate(events) if event[0] == "macro")


def test_run_succeeds_saves_and_audits_succeeded(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    stage_four, events, audits, caller, target, harness = _run_harness(
        tmp_path, monkeypatch
    )

    result = stage_four.run_property_stage_four(caller)

    assert result == target
    assert events[0] == ("audit", "Running")
    assert ("app", {"visible": True, "add_book": False}) in events
    assert harness.application.api.AutomationSecurity == 1
    assert next(event for event in events if event[0] == "open")[1:] == (
        str(target),
        {
            "update_links": False,
            "read_only": False,
            "ignore_read_only_recommended": True,
        },
    )
    assert [event[1:] for event in events if event[0] == "macro"] == [
        ("modControl.Run", ())
    ]
    assert ("autosave_during_macro", False) in events
    macro = _macro_index(events)
    assert _index(events, ("now",)) < _index(events, ("watchdog", "start")) < macro
    assert (
        macro
        < _index(events, ("watchdog", "stop"))
        < _index(events, ("watchdog", "join"))
    )
    assert _index(events, ("watchdog", "join")) < _index(events, ("save",))
    assert [audit["status"] for audit in audits] == ["Running", "Succeeded"]
    assert audits[-1]["input_workbook"] == target.name
    assert audits[-1]["last_successful_run"] == audits[-1]["finish"]
    assert events[-2:] == [("close_workbook", {"SaveChanges": False}), ("quit",)]
    assert harness.watchdogs[0].kwargs["excel_pid"] == 4321
    assert harness.watchdogs[0].kwargs["time_limit"] == timedelta(minutes=30)
    assert target.read_bytes() == b"noteval run output"


def test_run_terminates_owned_excel_if_quit_leaves_process_running(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    stage_four, _events, _audits, caller, _target, _harness = _run_harness(
        tmp_path, monkeypatch
    )
    terminated: list[int] = []
    monkeypatch.setattr(
        stage_four.stage_two_watchdog,
        "ProcessProbe",
        lambda: SimpleNamespace(process_snapshot=lambda: [{"pid": 4321}]),
    )
    monkeypatch.setattr(
        stage_four.stage_two_watchdog, "_terminate_process", terminated.append
    )

    stage_four.run_property_stage_four(caller)

    assert terminated == [4321]


def test_run_macro_failure_restores_backup_and_audits_failed(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    stage_four, events, audits, caller, target, _harness = _run_harness(
        tmp_path,
        monkeypatch,
        macro_error=RuntimeError("VBA failed"),
        macro_overwrites_workbook=True,
    )

    with pytest.raises(
        stage_four.PropertyStageFourError, match="NoteVal Run macro failed"
    ):
        stage_four.run_property_stage_four(caller)

    assert ("save",) not in events
    assert audits[-1]["status"] == "Failed"
    assert audits[-1]["last_successful_run"] == datetime(2026, 10, 1)
    assert target.read_bytes() == b"prepared noteval"
    assert events[-2:] == [("close_workbook", {"SaveChanges": False}), ("quit",)]
    assert ("watchdog", "join") in events


def test_run_rejects_macro_return_without_fresh_completion_times(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    stage_four, events, audits, caller, target, _harness = _run_harness(
        tmp_path, monkeypatch, complete_run=False
    )

    with pytest.raises(stage_four.PropertyStageFourError, match="stale"):
        stage_four.run_property_stage_four(caller)

    assert ("save",) not in events
    assert audits[-1]["status"] == "Failed"
    assert target.read_bytes() == b"prepared noteval"


def test_run_rejects_failed_save_and_restores_backup(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    stage_four, _events, audits, caller, target, _harness = _run_harness(
        tmp_path, monkeypatch, saved=False
    )

    with pytest.raises(stage_four.PropertyStageFourError, match="saved"):
        stage_four.run_property_stage_four(caller)

    assert audits[-1]["status"] == "Failed"
    assert target.read_bytes() == b"prepared noteval"


def test_run_records_interruption_and_restores_backup(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    stage_four, events, audits, caller, target, _harness = _run_harness(
        tmp_path,
        monkeypatch,
        macro_error=KeyboardInterrupt(),
        macro_overwrites_workbook=True,
    )

    with pytest.raises(KeyboardInterrupt):
        stage_four.run_property_stage_four(caller)

    assert audits[-1]["status"] == "Interrupted"
    assert ("save",) not in events
    assert target.read_bytes() == b"prepared noteval"


def test_run_watchdog_abort_audits_timed_out_without_closing_excel(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    stage_four, events, audits, caller, target, _harness = _run_harness(
        tmp_path,
        monkeypatch,
        macro_error=RuntimeError("COM call interrupted by termination"),
        macro_overwrites_workbook=True,
        macro_release=threading.Event(),
        watchdog_aborted=True,
    )

    with pytest.raises(stage_four.PropertyStageFourTimeoutError, match="timed out"):
        stage_four.run_property_stage_four(caller)

    assert audits[-1]["status"] == "Timed Out"
    assert ("save",) not in events
    assert not any(event[0] in {"close_workbook", "quit"} for event in events)
    assert target.read_bytes() == b"prepared noteval"
    assert [event for event in events if event[0] == "watchdog"] == [
        ("watchdog", "start"),
        ("watchdog", "stop"),
        ("watchdog", "join"),
    ]


def test_run_guard_conflict_is_reported_without_side_effects(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    stage_four, events, audits, caller, _target, _harness = _run_harness(
        tmp_path, monkeypatch
    )

    @contextlib.contextmanager
    def reject_duplicate(*_args: Any, **_kwargs: Any) -> Any:
        raise BondsLiteError("A stress run is already in progress for this workbook")
        yield

    monkeypatch.setattr(stage_four.shared, "workbook_run_guard", reject_duplicate)

    with pytest.raises(BondsLiteError, match="already in progress"):
        stage_four.run_property_stage_four(caller)

    assert events == []
    assert audits == []


def test_run_rejects_invalid_source_time_limit_override(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    stage_four, events, audits, caller, _target, _harness = _run_harness(
        tmp_path, monkeypatch
    )
    monkeypatch.setenv("SST_P4_TIME_LIMIT_MINUTES", "never")

    with pytest.raises(stage_four.PropertyStageFourError, match="SST_P4"):
        stage_four.run_property_stage_four(caller)

    assert not any(event[0] in {"app", "macro"} for event in events)
    assert audits[-1]["status"] == "Failed"


# ---------------------------------------------------------------------------
# T14-T15: watchdog time limit, prompt, termination and override
# ---------------------------------------------------------------------------


class _FakeClock:
    """Accelerated monotonic clock for exercising the watchdog state machine."""

    def __init__(self) -> None:
        self.seconds = 0.0

    def monotonic(self) -> float:
        return self.seconds

    def sleep(self, seconds: float) -> None:
        self.seconds += seconds


def _watchdog(
    *,
    prompt: Any = None,
    logger: Any = None,
    terminate: Any = None,
    clock: _FakeClock | None = None,
    time_limit: timedelta = timedelta(minutes=30),
) -> tuple[Any, _FakeClock, list[str], list[int]]:
    stage_four = _stage_four()
    fake_clock = clock or _FakeClock()
    lines: list[str] = []
    terminated: list[int] = []
    watchdog = stage_four.NoteValRunWatchdog(
        excel_pid=4321,
        time_limit=time_limit,
        prompt=prompt or (lambda _message: "W"),
        logger=logger or lines.append,
        clock=fake_clock,
        terminate=terminate or terminated.append,
    )
    return watchdog, fake_clock, lines, terminated


def test_watchdog_logs_heartbeat_every_five_minutes() -> None:
    reference: list[Any] = []
    lines: list[str] = []

    def log(message: str) -> None:
        lines.append(message)
        if sum("NoteVal running" in line for line in lines) == 3:
            reference[0].stop()

    watchdog, clock, _lines, terminated = _watchdog(logger=log)
    reference.append(watchdog)

    watchdog.start()
    watchdog.join(timeout=2)

    heartbeats = [line for line in lines if "NoteVal running" in line]
    assert heartbeats == [
        "NoteVal running | elapsed 0:05",
        "NoteVal running | elapsed 0:10",
        "NoteVal running | elapsed 0:15",
    ]
    assert terminated == []
    assert not any("LTM" in line for line in lines)


def test_watchdog_prompts_at_limit_and_waits_thirty_more_minutes() -> None:
    reference: list[Any] = []
    clock = _FakeClock()
    prompts: list[tuple[float, str]] = []

    def prompt(message: str) -> str:
        prompts.append((clock.seconds, message))
        if len(prompts) == 2:
            reference[0].stop()
        return "W"

    watchdog, _clock, _lines, terminated = _watchdog(prompt=prompt, clock=clock)
    reference.append(watchdog)

    watchdog.start()
    watchdog.join(timeout=2)

    assert [seconds for seconds, _message in prompts] == [30 * 60, 60 * 60]
    assert "Property Stage 4" in prompts[0][1]
    assert "[W]ait 30 more minutes or [A]bort?" in prompts[0][1]
    assert watchdog.aborted is False
    assert terminated == []


def test_watchdog_abort_terminates_exactly_the_owned_excel_pid() -> None:
    watchdog, _clock, _lines, terminated = _watchdog(prompt=lambda _message: "A")

    watchdog.start()
    watchdog.join(timeout=2)

    assert not watchdog.is_alive()
    assert watchdog.aborted is True
    assert watchdog.process_terminated is True
    assert terminated == [4321]


def test_watchdog_retries_failed_termination() -> None:
    attempts: list[int] = []
    lines: list[str] = []

    def terminate(pid: int) -> None:
        attempts.append(pid)
        if len(attempts) == 1:
            raise OSError("access denied")

    watchdog, _clock, _lines, _terminated = _watchdog(
        prompt=lambda _message: "a", terminate=terminate, logger=lines.append
    )

    watchdog.start()
    watchdog.join(timeout=2)

    assert attempts == [4321, 4321]
    assert watchdog.process_terminated is True
    assert any(line.startswith("Abort requested") for line in lines)


def test_watchdog_reprompts_on_unrecognised_answer() -> None:
    answers = iter(["maybe", "A"])
    watchdog, _clock, lines, terminated = _watchdog(
        prompt=lambda _message: next(answers)
    )

    watchdog.start()
    watchdog.join(timeout=2)

    assert terminated == [4321]
    assert "Please enter W to wait or A to abort" in lines


def test_watchdog_stop_cancels_console_prompt(monkeypatch: pytest.MonkeyPatch) -> None:
    stage_four = _stage_four()
    monkeypatch.setitem(
        sys.modules,
        "msvcrt",
        types.SimpleNamespace(kbhit=lambda: False, getwch=lambda: ""),
    )
    prompt_started = threading.Event()
    watchdog = stage_four.NoteValRunWatchdog(
        excel_pid=4321,
        time_limit=timedelta(minutes=30),
        logger=lambda _message: prompt_started.set(),
        clock=_FakeClock(),
        terminate=lambda _pid: pytest.fail("No termination"),
    )
    result: list[str] = []
    prompt_thread = threading.Thread(
        target=lambda: result.append(watchdog._ask_to_continue())
    )

    prompt_thread.start()
    assert prompt_started.wait(timeout=1)
    watchdog.stop()
    prompt_thread.join(timeout=1)

    assert not prompt_thread.is_alive()
    assert result == ["W"]


@pytest.mark.parametrize(
    ("override", "frozen", "expected"),
    [
        (None, False, timedelta(minutes=30)),
        ("", False, timedelta(minutes=30)),
        ("  ", False, timedelta(minutes=30)),
        ("0.5", False, timedelta(seconds=30)),
        ("45", False, timedelta(minutes=45)),
        ("0.5", True, timedelta(minutes=30)),
        ("invalid", True, timedelta(minutes=30)),
    ],
)
def test_configured_time_limit_honours_source_only_override(
    override: str | None, frozen: bool, expected: timedelta
) -> None:
    assert _stage_four().configured_noteval_time_limit(override, frozen) == expected


@pytest.mark.parametrize("override", ["0", "-1", "nan", "inf", "invalid"])
def test_configured_time_limit_rejects_invalid_source_override(override: str) -> None:
    with pytest.raises(ValueError, match="SST_P4_TIME_LIMIT_MINUTES"):
        _stage_four().configured_noteval_time_limit(override, False)


def test_configured_time_limit_warns_when_frozen_override_is_ignored(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    stage_four = _stage_four()
    warnings: list[str] = []
    monkeypatch.setattr(
        stage_four.logger,
        "warning",
        lambda message, *args: warnings.append(message % args),
    )

    stage_four.configured_noteval_time_limit("5", True)

    assert any("SST_P4_TIME_LIMIT_MINUTES" in warning for warning in warnings)


# ---------------------------------------------------------------------------
# T16: installer panel, button and macro idempotence
# ---------------------------------------------------------------------------


def _column_number(letters: str) -> int:
    value = 0
    for letter in letters:
        value = value * 26 + ord(letter) - 64
    return value


def _column_letters(number: int) -> str:
    letters = ""
    while number:
        number, remainder = divmod(number - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


class _FakeSheet:
    """Record writes to the I. Control sheet through a minimal range surface."""

    def __init__(self, shapes: list[Any] | None = None) -> None:
        self.name = "I. Control"
        self.values: dict[str, object] = {}
        self.merges: list[str] = []
        self.shapes = _FakeShapes(shapes or [])
        self.api = SimpleNamespace(
            Shapes=self.shapes, Buttons=lambda: self.shapes.buttons
        )

    def range(self, first: Any, last: Any = None) -> "_FakeRange":
        if isinstance(first, str):
            parts = first.replace("$", "").split(":")
            corners = [re.fullmatch(r"([A-Z]+)(\d+)", part).groups() for part in parts]
            first = (int(corners[0][1]), _column_number(corners[0][0]))
            if len(corners) == 2:
                last = (int(corners[1][1]), _column_number(corners[1][0]))
        return _FakeRange(self, first, last or first)


class _FakeRange:
    def __init__(
        self, sheet: _FakeSheet, first: tuple[int, int], last: tuple[int, int]
    ) -> None:
        self.sheet = sheet
        self.row, self.column = first
        self.last_row, self.last_column = last
        self.left = self.column * 10.0
        self.top = self.row * 15.0
        self.width = (self.last_column - self.column + 1) * 10.0
        self.height = (self.last_row - self.row + 1) * 15.0
        self.font = SimpleNamespace()
        self.api = SimpleNamespace(MergeCells=False)
        self.color: object = None
        self.number_format: object = None

    def get_address(self) -> str:
        first = f"${_column_letters(self.column)}${self.row}"
        if (self.row, self.column) == (self.last_row, self.last_column):
            return first
        return f"{first}:${_column_letters(self.last_column)}${self.last_row}"

    def merge(self) -> None:
        self.sheet.merges.append(self.get_address())

    @property
    def value(self) -> object:
        return self.sheet.values.get(self.get_address())

    @value.setter
    def value(self, value: object) -> None:
        self.sheet.values[self.get_address()] = value

    def __getitem__(self, _index: Any) -> "_FakeRange":
        return _FakeRange(self.sheet, (self.row, self.column), (self.row, self.column))


class _FakeButtons:
    def __init__(self, shapes: "_FakeShapes") -> None:
        self.shapes = shapes
        self.added: list[Any] = []

    def Add(  # noqa: N802 - Excel COM name
        self, left: float, top: float, width: float, height: float
    ) -> Any:
        button = SimpleNamespace(
            Name="Button 1", Left=left, Top=top, Width=width, Height=height
        )
        self.added.append(button)
        self.shapes.append(button)
        return button

    def Item(self, name: str) -> Any:  # noqa: N802 - Excel COM name
        return self.shapes.Item(name)


class _FakeShapes(list):
    def __init__(self, shapes: list[Any]) -> None:
        super().__init__(shapes)
        self.buttons = _FakeButtons(self)

    def Item(self, name: str) -> Any:  # noqa: N802 - Excel COM name
        for shape in self:
            if shape.Name == name:
                return shape
        raise KeyError(name)


class _FakeNames:
    def __init__(self, ranges: dict[str, Any] | None = None) -> None:
        self.added: dict[str, str] = {}
        self.ranges = dict(ranges or {})

    def add(self, name: str, refers_to: str) -> None:
        self.added[name] = refers_to

    def __iter__(self) -> Any:
        return iter(SimpleNamespace(name=name) for name in [*self.ranges, *self.added])

    def __getitem__(self, name: str) -> Any:
        return SimpleNamespace(
            refers_to_range=self.ranges.get(name),
            delete=lambda: self.added.pop(name),
        )


def _install_book(
    sheet: _FakeSheet, names: _FakeNames | None = None, *, occupied: int = 0
) -> Any:
    return SimpleNamespace(
        name="Draft Input spreadsheet for SST.xlsm",
        names=names or _FakeNames(),
        sheets={"I. Control": sheet},
        app=SimpleNamespace(
            api=SimpleNamespace(
                EnableEvents=True,
                WorksheetFunction=SimpleNamespace(CountA=lambda _area: occupied),
            )
        ),
    )


def test_install_requires_stage_three_controls_first(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    stage_four = _stage_four()

    def missing(_book: Any) -> None:
        raise stage_four.stage_three.PropertyStageThreeError("panel is missing")

    monkeypatch.setattr(
        stage_four.stage_three, "require_property_stage_three_panel", missing
    )
    book = _install_book(_FakeSheet())

    with pytest.raises(
        stage_four.PropertyStageFourError,
        match="Install the Property Stage 3 controls before Stage 4",
    ):
        stage_four._initialise_property_stage_four_panel(book)

    assert book.names.added == {}


def test_install_panel_creates_named_fields_then_is_idempotent(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    stage_four = _stage_four()
    validated: list[Any] = []
    monkeypatch.setattr(
        stage_four.stage_three, "require_property_stage_three_panel", lambda _: None
    )
    monkeypatch.setattr(
        stage_four, "require_property_stage_four_panel", validated.append
    )
    sheet = _FakeSheet()
    book = _install_book(sheet)

    stage_four._initialise_property_stage_four_panel(book)
    first_install = dict(book.names.added)
    stage_four._initialise_property_stage_four_panel(book)

    assert first_install == {
        "SST_PropStage4_Panel": "='I. Control'!$AB$48:$AI$58",
        "SST_PropStage4_LastRun": "='I. Control'!$AF$49",
        "SST_PropStage4_Start": "='I. Control'!$AF$50",
        "SST_PropStage4_Finish": "='I. Control'!$AF$51",
        "SST_PropStage4_Elapsed": "='I. Control'!$AF$52",
        "SST_PropStage4_User": "='I. Control'!$AF$53",
        "SST_PropStage4_Status": "='I. Control'!$AF$54",
        "SST_PropStage4_InputWorkbook": "='I. Control'!$AF$55",
        "SST_PropStage4_LastSuccessfulRun": "='I. Control'!$AF$57",
    }
    assert book.names.added == first_install
    assert sheet.values["$AB$48:$AI$48"] == "Property Stage 4 - Run NoteVal"
    assert sheet.values["$AF$54:$AI$54"] == "Not Run"
    assert sheet.values["$AB$57:$AE$57"] == "Last Successful Run"
    assert "$AF$55:$AI$56" in sheet.merges
    assert book.app.api.EnableEvents is True
    assert len(validated) == 2


def test_install_panel_rejects_incomplete_names(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    stage_four = _stage_four()
    monkeypatch.setattr(
        stage_four.stage_three, "require_property_stage_three_panel", lambda _: None
    )
    names = _FakeNames({"SST_PropStage4_Panel": None})

    with pytest.raises(stage_four.PropertyStageFourError, match="incomplete"):
        stage_four._initialise_property_stage_four_panel(
            _install_book(_FakeSheet(), names)
        )

    assert names.added == {}


def test_install_panel_refuses_occupied_area(monkeypatch: pytest.MonkeyPatch) -> None:
    stage_four = _stage_four()
    monkeypatch.setattr(
        stage_four.stage_three, "require_property_stage_three_panel", lambda _: None
    )
    book = _install_book(_FakeSheet(), occupied=1)

    with pytest.raises(stage_four.PropertyStageFourError, match="contains cells"):
        stage_four._initialise_property_stage_four_panel(book)

    assert book.names.added == {}


def test_install_panel_refuses_area_overlapping_a_shape(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    stage_four = _stage_four()
    monkeypatch.setattr(
        stage_four.stage_three, "require_property_stage_three_panel", lambda _: None
    )
    shape = SimpleNamespace(
        Name="Picture 1", Left=300.0, Top=750.0, Width=20, Height=20
    )
    book = _install_book(_FakeSheet([shape]))

    with pytest.raises(stage_four.PropertyStageFourError, match="overlaps"):
        stage_four._initialise_property_stage_four_panel(book)


def test_install_button_uses_stage_three_column_and_is_idempotent() -> None:
    stage_four = _stage_four()
    stage_three_button = SimpleNamespace(
        Name="btnPrepareNoteVal",
        Left=600.0,
        Top=525.0,
        Width=100.0,
        Height=30.0,
        TopLeftCell=SimpleNamespace(Column=60),
    )
    sheet = _FakeSheet([stage_three_button])
    panel = sheet.range("AB48:AI58")
    book = _install_book(sheet, _FakeNames({"SST_PropStage4_Panel": panel}))

    stage_four._install_property_stage_four_button(book)
    stage_four._install_property_stage_four_button(book)

    assert len(sheet.shapes.buttons.added) == 1
    button = sheet.shapes.Item("btnRunNoteVal")
    assert (button.Left, button.Top, button.Width, button.Height) == (
        600.0,
        panel.top,
        100.0,
        30.0,
    )
    assert button.Caption == "Run NoteVal"
    assert button.OnAction == "'Draft Input spreadsheet for SST.xlsm'!RunNoteVal"


class _FakeCodeModule:
    def __init__(self, source: str) -> None:
        self.source = source

    @property
    def CountOfLines(self) -> int:  # noqa: N802 - Excel COM name
        return len(self.source.splitlines())

    def Lines(self, _start: int, _count: int) -> str:  # noqa: N802 - Excel COM name
        return self.source

    def InsertLines(self, _line: int, text: str) -> None:  # noqa: N802
        self.source += text


def _vba_book(code: _FakeCodeModule) -> Any:
    components = SimpleNamespace(Item=lambda _name: SimpleNamespace(CodeModule=code))
    return SimpleNamespace(
        api=SimpleNamespace(VBProject=SimpleNamespace(VBComponents=components))
    )


def test_install_macro_is_idempotent() -> None:
    stage_four = _stage_four()
    code = _FakeCodeModule(
        "Private Sub LaunchBondsLite(ByVal stage As String)\nEnd Sub\n"
        'Public Sub PrepareNoteVal()\n    LaunchBondsLite "P3"\nEnd Sub\n'
    )

    stage_four._install_property_stage_four_macro(_vba_book(code))
    stage_four._install_property_stage_four_macro(_vba_book(code))

    assert code.source.count("Public Sub RunNoteVal()") == 1
    assert 'Public Sub RunNoteVal()\n    LaunchBondsLite "P4"\nEnd Sub' in code.source
    assert code.source.count("Public Sub PrepareNoteVal()") == 1


@pytest.mark.parametrize(
    ("source", "message"),
    [
        (
            "Private Sub LaunchBondsLite(ByVal stage As String)\nEnd Sub\n"
            "Public Sub RunNoteVal()\n    LaunchBondsLite 4\nEnd Sub\n",
            "unexpected implementation",
        ),
        ("Public Sub Other()\nEnd Sub\n", "LaunchBondsLite signature"),
    ],
)
def test_install_macro_rejects_unsupported_launcher(source: str, message: str) -> None:
    stage_four = _stage_four()
    code = _FakeCodeModule(source)

    with pytest.raises(stage_four.PropertyStageFourError, match=message):
        stage_four._install_property_stage_four_macro(_vba_book(code))

    assert code.source == source


def test_stage_four_errors_share_the_property_error_hierarchy() -> None:
    stage_four = _stage_four()

    assert issubclass(
        stage_four.PropertyStageFourError,
        stage_four.stage_three.PropertyStageThreeError,
    )
    assert issubclass(
        stage_four.PropertyStageFourTimeoutError, stage_four.PropertyStageFourError
    )
    assert issubclass(stage_four.PropertyStageFourError, BondsLiteError)


@pytest.mark.excel
@pytest.mark.skipif(
    os.environ.get("BONDS_LITE_RUN_PROPERTY_STAGE_FOUR_EXCEL_TESTS") != "1",
    reason="Set BONDS_LITE_RUN_PROPERTY_STAGE_FOUR_EXCEL_TESTS=1",
)
def test_property_stage_four_runs_real_noteval_copy(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    import hashlib
    import shutil

    import xlwings as xw

    stage_four = _stage_four()
    stage_two = stage_four.stage_two
    stage_three = stage_four.stage_three
    project = Path(__file__).resolve().parents[1]
    original_draft = project / "Draft Input spreadsheet for SST.xlsm"
    original_noteval = (
        project
        / "SST Output_Test"
        / "Property"
        / "JRL_LTM Note Valuation Tool v1.40_Test.xlsb"
    )
    original_flash = (
        project / "SST Output_Test" / "Property" / "LTM_FlashModel_v3.2a_Test.xlsb"
    )
    originals = (original_draft, original_noteval, original_flash)
    hashes = {path: hashlib.sha256(path.read_bytes()).digest() for path in originals}

    draft = tmp_path / original_draft.name
    noteval = tmp_path / "SST Output_Test" / "Property" / original_noteval.name
    flash = tmp_path / "SST Output_Test" / "Property" / original_flash.name
    noteval.parent.mkdir(parents=True)
    shutil.copy2(original_draft, draft)
    shutil.copy2(original_noteval, noteval)
    shutil.copy2(original_flash, flash)
    monkeypatch.setenv(stage_four.TIME_LIMIT_ENVIRONMENT, "5")

    application = xw.App(visible=False, add_book=False)
    caller_pid = int(application.pid)
    book = None
    try:
        application.api.AutomationSecurity = 3
        application.api.EnableEvents = False
        application.display_alerts = False
        book = application.books.open(str(draft), update_links=False)
        stage_four.require_property_stage_four_panel(book)
        book.names["SST_Property_standalone"].refers_to_range.value = True
        stage_two_input = stage_two.STAGE_TWO_FIELD_NAMES
        stage_three_input = stage_three.STAGE_THREE_FIELD_NAMES
        flash_link = "SST Output_Test/Property/LTM_FlashModel_v3.2a_Test.xlsb"
        noteval_link = (
            "SST Output_Test/Property/JRL_LTM Note Valuation Tool v1.40_Test.xlsb"
        )
        book.names[stage_two_input["status"]].refers_to_range.value = "Succeeded"
        book.names[stage_two_input["input_workbook"]].refers_to_range.value = flash_link
        book.names[stage_two_input["last_successful_run"]].refers_to_range.value = (
            datetime(2026, 11, 11, 17, 44, 27)
        )
        book.names[stage_three_input["status"]].refers_to_range.value = "Succeeded"
        book.names[stage_three_input["input_workbook"]].refers_to_range.value = (
            flash_link
        )
        book.names[stage_three_input["finish"]].refers_to_range.value = datetime(
            2026, 11, 11, 21, 39, 48
        )
        book.names[
            stage_three_input["last_successful_output"]
        ].refers_to_range.value = noteval_link
        book.save()

        result = stage_four.run_property_stage_four(book)

        assert result == noteval
        audit = stage_four._read_property_stage_four_fields(book)
        assert audit["status"] == "Succeeded"
        assert isinstance(audit["last_successful_run"], datetime)
    finally:
        if book is not None:
            book.close()
        application.api.DisplayAlerts = False
        application.quit()
        if any(
            process.get("pid") == caller_pid
            for process in stage_four.stage_two_watchdog.ProcessProbe().process_snapshot()
        ):
            stage_four.stage_two_watchdog._terminate_process(caller_pid)
    assert all(
        hashlib.sha256(path.read_bytes()).digest() == expected
        for path, expected in hashes.items()
    )
