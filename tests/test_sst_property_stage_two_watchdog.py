"""Behavioural contracts for the Property Stage 2 LTM run watchdog."""

from __future__ import annotations

import sys
import threading
import types
from datetime import timedelta
from types import SimpleNamespace
from typing import Any

import pytest


def _watchdog_module() -> Any:
    """Import lazily so the contracts collect before the watchdog exists."""
    import sst_property_stage_two_watchdog

    return sst_property_stage_two_watchdog


class _FakeClock:
    """Accelerated monotonic clock for exercising the watchdog state machine."""

    def __init__(self) -> None:
        self.seconds = 0.0

    def monotonic(self) -> float:
        return self.seconds

    def sleep(self, seconds: float) -> None:
        self.seconds += seconds


class _FakeProbe:
    def __init__(
        self,
        clock: _FakeClock,
        *,
        runner: int | None = 4322,
        cpu_time: Any = None,
    ) -> None:
        self.clock = clock
        self.runner = runner
        self.cpu_time = cpu_time or (lambda: 1.0)
        self.terminate_calls: list[int] = []

    def find_runner_under_excel(self, _excel_pid: int) -> int | None:
        if callable(self.runner):
            return self.runner()
        return self.runner

    def runner_cpu_time(self, _excel_pid: int) -> float | None:
        return self.cpu_time()


class _FakeReader:
    def __init__(self, messages: list[str] | None) -> None:
        self.messages = messages

    def read(self) -> list[str] | None:
        return self.messages


def _make_watchdog(
    *,
    clock: _FakeClock | None = None,
    probe: _FakeProbe | None = None,
    reader: _FakeReader | None = None,
    prompt: Any = None,
    logger: Any = None,
    scenarios: int = 16,
) -> tuple[Any, _FakeClock, _FakeProbe, list[str]]:
    module = _watchdog_module()
    fake_clock = clock or _FakeClock()
    fake_probe = probe or _FakeProbe(fake_clock)
    lines: list[str] = []
    watchdog = module.LtmRunWatchdog(
        excel_pid=4321,
        workbook_name="LTM.xlsb",
        scenarios=scenarios,
        probe=fake_probe,
        reader=reader or _FakeReader(None),
        prompt=prompt or (lambda _message: "W"),
        logger=logger or lines.append,
        clock=fake_clock,
    )
    return watchdog, fake_clock, fake_probe, lines


@pytest.mark.parametrize(
    ("scenarios", "expected"),
    [(1, 30), (3, 30), (4, 40), (16, 160)],
)
def test_ltm_time_limit_uses_scenario_floor_and_multiplier(
    scenarios: int, expected: int
) -> None:
    assert _watchdog_module().ltm_time_limit(scenarios) == timedelta(minutes=expected)


@pytest.mark.parametrize("scenarios", [0, -1, -16])
def test_ltm_time_limit_rejects_non_positive_scenario_count(scenarios: int) -> None:
    with pytest.raises(ValueError):
        _watchdog_module().ltm_time_limit(scenarios)


def test_configured_time_limit_accepts_source_only_override() -> None:
    assert _watchdog_module().configured_time_limit(
        16, override="0.5", frozen=False
    ) == timedelta(seconds=30)


def test_configured_time_limit_ignores_override_when_frozen() -> None:
    assert _watchdog_module().configured_time_limit(
        16, override="invalid", frozen=True
    ) == timedelta(minutes=160)


@pytest.mark.parametrize("override", ["0", "-1", "nan", "invalid"])
def test_configured_time_limit_rejects_invalid_source_override(
    override: str,
) -> None:
    with pytest.raises(ValueError):
        _watchdog_module().configured_time_limit(2, override=override, frozen=False)


def test_count_completed_scenarios_counts_distinct_markers_only() -> None:
    messages = [
        "Preparing run",
        "Run 1 - End of run!",
        "Run 2 - End of run!",
        "Run 2 - End of run!",
        "Run 4 - End of run!",
        "End of multirun!",
        "Run 5 - Still running",
    ]

    assert _watchdog_module().count_completed_scenarios(messages) == 3


def test_watchdog_heartbeat_every_five_minutes_handles_unreadable_progress() -> None:
    watchdog_ref: list[Any] = []
    logged: list[str] = []

    def log(message: str) -> None:
        logged.append(message)
        if sum("LTM running" in line for line in logged) == 4:
            watchdog_ref[0].stop()

    class AlternatingReader:
        reads = 0

        def read(self) -> list[str] | None:
            self.reads += 1
            return (
                ["Run 1 - End of run!", "Run 2 - End of run!"]
                if self.reads == 1
                else None
            )

    watchdog, clock, _probe, _ = _make_watchdog(
        reader=AlternatingReader(), logger=log, scenarios=16
    )
    watchdog_ref.append(watchdog)

    watchdog.start()
    watchdog.join(timeout=2)

    heartbeats = [line for line in logged if "LTM running" in line]
    assert len(heartbeats) == 4
    assert [
        f"elapsed 0:{minutes:02d}" in line
        for minutes, line in zip((5, 10, 15, 20), heartbeats)
    ]
    assert "Scenario 2 of 16 complete" in heartbeats[0]
    assert all("Scenario" not in line for line in heartbeats[1:])
    assert clock.seconds >= 20 * 60


def test_watchdog_warns_once_per_stall_episode() -> None:
    watchdog_ref: list[Any] = []
    warnings: list[str] = []
    clock = _FakeClock()

    def cpu_time() -> float:
        return 2.0 if clock.seconds < 15 * 60 else 3.0

    probe = _FakeProbe(clock, cpu_time=cpu_time)

    def log(message: str) -> None:
        if "stall" in message.casefold():
            warnings.append(message)
            if len(warnings) == 2:
                watchdog_ref[0].stop()

    watchdog, _clock, _probe, _lines = _make_watchdog(
        clock=clock, probe=probe, logger=log
    )
    watchdog_ref.append(watchdog)

    watchdog.start()
    watchdog.join(timeout=2)

    assert len(warnings) == 2


def test_watchdog_warns_when_runner_exited_but_macro_pending() -> None:
    watchdog_ref: list[Any] = []
    warnings: list[str] = []
    clock = _FakeClock()
    probe = _FakeProbe(
        clock,
        runner=lambda: 4322 if clock.seconds < 60 else None,
    )

    def log(message: str) -> None:
        if "dialog" in message.casefold():
            warnings.append(message)
            watchdog_ref[0].stop()

    watchdog, _clock, _probe, _lines = _make_watchdog(
        clock=clock, probe=probe, logger=log
    )
    watchdog_ref.append(watchdog)

    watchdog.start()
    watchdog.join(timeout=2)

    assert len(warnings) == 1
    assert "error dialog" in warnings[0].casefold()
    assert clock.seconds >= 3 * 60


def test_watchdog_prompts_at_limit_and_waits_on_w() -> None:
    watchdog_ref: list[Any] = []
    clock = _FakeClock()
    prompt_times: list[float] = []

    def prompt(_message: str) -> str:
        prompt_times.append(clock.seconds)
        if len(prompt_times) == 2:
            watchdog_ref[0].stop()
        return "W"

    watchdog, _clock, probe, _lines = _make_watchdog(
        clock=clock, prompt=prompt, scenarios=1
    )
    watchdog_ref.append(watchdog)

    watchdog.start()
    watchdog.join(timeout=2)

    assert prompt_times == [30 * 60, 60 * 60]
    assert watchdog.aborted is False
    assert probe.terminate_calls == []


def test_watchdog_wait_interval_starts_when_user_answers() -> None:
    clock = _FakeClock()
    prompt_times: list[float] = []

    def prompt(_message: str) -> str:
        prompt_times.append(clock.seconds)
        if len(prompt_times) == 1:
            clock.seconds += 40 * 60
        else:
            watchdog.stop()
        return "W"

    watchdog, _clock, _probe, _lines = _make_watchdog(
        clock=clock, prompt=prompt, scenarios=1
    )

    watchdog.run()

    assert prompt_times == [30 * 60, 100 * 60]
    assert watchdog.aborted is False


def test_watchdog_aborts_on_a_and_terminates_owned_run_once(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    module = _watchdog_module()
    calls: list[int] = []
    monkeypatch.setattr(module, "terminate_owned_run", lambda pid: calls.append(pid))
    watchdog, _clock, _probe, _lines = _make_watchdog(
        prompt=lambda _message: "A", scenarios=1
    )

    watchdog.start()
    watchdog.join(timeout=2)

    assert watchdog.aborted is True
    assert calls == [4321]


def test_watchdog_stop_cancels_console_prompt(monkeypatch: pytest.MonkeyPatch) -> None:
    module = _watchdog_module()
    monkeypatch.setitem(
        sys.modules,
        "msvcrt",
        types.SimpleNamespace(kbhit=lambda: False, getwch=lambda: ""),
    )
    prompt_started = threading.Event()
    watchdog = module.LtmRunWatchdog(
        excel_pid=4321,
        workbook_name="LTM.xlsb",
        scenarios=16,
        logger=lambda _message: prompt_started.set(),
        clock=_FakeClock(),
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


def test_watchdog_without_runner_uses_elapsed_time_for_heartbeat_and_limit(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    module = _watchdog_module()
    messages: list[str] = []
    clock = _FakeClock()
    probe = _FakeProbe(clock, runner=None)
    prompt_times: list[float] = []
    terminated: list[int] = []
    monkeypatch.setattr(module, "terminate_owned_run", terminated.append)

    def prompt(_message: str) -> str:
        prompt_times.append(clock.seconds)
        return "A"

    watchdog, _clock, _probe, _lines = _make_watchdog(
        clock=clock,
        probe=probe,
        logger=messages.append,
        prompt=prompt,
        scenarios=1,
    )

    watchdog.start()
    watchdog.join(timeout=2)

    missing = [line for line in messages if "runner was not found" in line.casefold()]
    assert len(missing) == 1
    assert any("LTM running" in line for line in messages)
    assert prompt_times == [30 * 60]
    assert watchdog.aborted is True
    assert terminated == [4321]


def test_find_runner_under_excel_pid_ignores_another_excel_tree() -> None:
    module = _watchdog_module()
    snapshot = [
        {"pid": 100, "ppid": 1, "exe": "EXCEL.EXE"},
        {"pid": 110, "ppid": 100, "exe": "cmd.exe"},
        {"pid": 120, "ppid": 110, "exe": "LTMflashArunmanager.exe"},
        {"pid": 200, "ppid": 1, "exe": "EXCEL.EXE"},
        {"pid": 210, "ppid": 200, "exe": "cmd.exe"},
        {"pid": 220, "ppid": 210, "exe": "LTMflashArunmanager.exe"},
    ]

    assert module.ProcessProbe.find_runner_under_excel_pid(snapshot, 100) == 120


def test_terminate_owned_run_only_targets_owned_tree(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    module = _watchdog_module()
    terminated: list[int] = []
    snapshot = [
        {"pid": 100, "ppid": 1, "exe": "EXCEL.EXE"},
        {"pid": 110, "ppid": 100, "exe": "cmd.exe"},
        {"pid": 120, "ppid": 110, "exe": "LTMflashArunmanager.exe"},
        {"pid": 200, "ppid": 1, "exe": "EXCEL.EXE"},
        {"pid": 220, "ppid": 200, "exe": "LTMflashArunmanager.exe"},
    ]

    class FakeProbe:
        def process_snapshot(self) -> list[dict[str, object]]:
            return snapshot

    monkeypatch.setattr(module, "ProcessProbe", FakeProbe)
    monkeypatch.setattr(module, "_terminate_process", terminated.append)

    module.terminate_owned_run(100)

    assert terminated == [120, 110, 100]


def test_terminate_owned_run_keeps_excel_open_if_runner_cannot_be_terminated(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    module = _watchdog_module()
    terminated: list[int] = []
    snapshot = [
        {"pid": 100, "ppid": 1, "exe": "EXCEL.EXE"},
        {"pid": 110, "ppid": 100, "exe": "cmd.exe"},
        {"pid": 120, "ppid": 110, "exe": "LTMflashArunmanager.exe"},
    ]

    class FakeProbe:
        def process_snapshot(self) -> list[dict[str, object]]:
            return snapshot

    def fail_runner(pid: int) -> None:
        terminated.append(pid)
        raise OSError("access denied")

    monkeypatch.setattr(module, "ProcessProbe", FakeProbe)
    monkeypatch.setattr(module, "_terminate_process", fail_runner)

    with pytest.raises(RuntimeError, match="Could not terminate LTM process 120"):
        module.terminate_owned_run(100)

    assert terminated == [120]


def test_progress_reader_returns_none_when_com_rejects_a_call(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    module = _watchdog_module()
    rejected = RuntimeError("RPC_E_CALL_REJECTED")

    class RejectedRange:
        @property
        def value(self) -> object:
            raise rejected

    workbook = SimpleNamespace(
        names={
            "ProgressFirstCell": SimpleNamespace(
                refers_to_range=SimpleNamespace(
                    sheet=SimpleNamespace(
                        api=SimpleNamespace(
                            Cells=lambda *_args: SimpleNamespace(
                                End=lambda _direction: SimpleNamespace(Row=1)
                            )
                        ),
                        range=lambda *_args: RejectedRange(),
                    ),
                    row=1,
                    column=1,
                )
            )
        }
    )
    workbook.name = "LTM.xlsb"

    class FakeBooks:
        def __getitem__(self, name: str) -> Any:
            assert name == "LTM.xlsb"
            return workbook

    application = SimpleNamespace(books=FakeBooks())
    fake_xlwings = types.SimpleNamespace(apps={4321: application})
    fake_pythoncom = types.SimpleNamespace(CoInitialize=lambda: None)
    monkeypatch.setitem(sys.modules, "xlwings", fake_xlwings)
    monkeypatch.setitem(sys.modules, "pythoncom", fake_pythoncom)

    reader = module.ProgressReader(excel_pid=4321, workbook_name="LTM.xlsb")

    assert reader.read() is None


def test_progress_reader_disables_a_com_call_that_exceeds_its_deadline(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    module = _watchdog_module()
    release = threading.Event()
    calls: list[None] = []

    def blocked_read(_reader: Any) -> list[str]:
        calls.append(None)
        release.wait(timeout=1)
        return []

    monkeypatch.setattr(
        module.ProgressReader, "_read_excel_progress", blocked_read, raising=False
    )
    reader = module.ProgressReader(
        excel_pid=4321, workbook_name="LTM.xlsb", timeout_seconds=0.01
    )

    assert reader.read() is None
    assert reader.read() is None
    assert calls == [None]
    release.set()
