"""Behavioural contracts for the Property Stage 2 LTM Flash run."""

from __future__ import annotations

import contextlib
import os
from datetime import datetime
from pathlib import Path
from types import SimpleNamespace
from typing import Any

import pytest

from bonds_lite_input import BondsLiteError

PROPERTY_FLAG = "SST_Property_standalone"


def _stage_two() -> Any:
    """Import lazily so contracts collect before Property Stage 2 is implemented."""
    import sst_property_stage_two

    return sst_property_stage_two


def _name(value: object) -> Any:
    return SimpleNamespace(
        refers_to_range=SimpleNamespace(
            value=value,
            size=1,
            api=SimpleNamespace(Areas=SimpleNamespace(Count=1)),
        )
    )


def _draft(
    path: Path,
    handoff: object,
    *,
    flag: object = True,
    status: object = "Succeeded",
) -> Any:
    book = SimpleNamespace(
        fullname=os.path.normcase(str(path.resolve())),
        names={
            PROPERTY_FLAG: _name(flag),
            "SST_PropStage1_Status": _name(status),
            "SST_PropStage1_LastSuccessfulOutput": _name(handoff),
        },
    )
    for name in book.names.values():
        name.refers_to_range.sheet = SimpleNamespace(name="I. Control", book=book)
    return book


def _patch_stage_one_handoff(
    monkeypatch: pytest.MonkeyPatch, stage_two: Any, draft: Any
) -> None:
    """Keep these contracts focused on Stage 2's named handoff interpretation."""
    monkeypatch.setattr(
        stage_two.stage_one, "require_property_stage_one_panel", lambda _: None
    )
    monkeypatch.setattr(
        stage_two.stage_one,
        "_read_property_stage_one_fields",
        lambda _: {
            "status": draft.names["SST_PropStage1_Status"].refers_to_range.value,
            "last_output": draft.names[
                "SST_PropStage1_LastSuccessfulOutput"
            ].refers_to_range.value,
        },
    )


def test_resolve_inputs_requires_property_flag_before_handoff(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    stage_two = _stage_two()
    draft = _draft(tmp_path / "Draft.xlsm", "missing.xlsb", flag=False)
    _patch_stage_one_handoff(monkeypatch, stage_two, draft)

    with pytest.raises(BondsLiteError, match="Property stress is not selected"):
        stage_two.resolve_property_stage_two_inputs(draft)


def test_resolve_inputs_requires_successful_stage_one(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    stage_two = _stage_two()
    output = tmp_path / "prepared.xlsb"
    output.write_bytes(b"workbook")
    draft = _draft(tmp_path / "Draft.xlsm", output.name, status="Failed")
    _patch_stage_one_handoff(monkeypatch, stage_two, draft)

    with pytest.raises(BondsLiteError, match="Stage 1.*Succeeded"):
        stage_two.resolve_property_stage_two_inputs(draft)


@pytest.mark.parametrize(
    ("filename", "content", "message"),
    [
        ("missing.xlsb", None, "does not exist"),
        ("empty.xlsb", b"", "empty"),
        ("unsupported.xlsx", b"workbook", "macro-enabled"),
    ],
)
def test_resolve_inputs_rejects_invalid_stage_one_handoff(
    tmp_path: Path,
    filename: str,
    content: bytes | None,
    message: str,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    stage_two = _stage_two()
    if content is not None:
        (tmp_path / filename).write_bytes(content)
    draft = _draft(tmp_path / "Draft.xlsm", filename)
    _patch_stage_one_handoff(monkeypatch, stage_two, draft)

    with pytest.raises(BondsLiteError, match=message):
        stage_two.resolve_property_stage_two_inputs(draft)


def test_resolve_inputs_uses_named_stage_one_handoff(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    stage_two = _stage_two()
    output = tmp_path / "SST Output_Test" / "Property" / "prepared.xlsb"
    output.parent.mkdir(parents=True)
    output.write_bytes(b"workbook")
    draft = _draft(
        tmp_path / "Draft.xlsm",
        "SST Output_Test/Property/prepared.xlsb",
    )
    _patch_stage_one_handoff(monkeypatch, stage_two, draft)

    inputs = stage_two.resolve_property_stage_two_inputs(draft)

    assert inputs.workbook_path == output


def test_resolve_inputs_rejects_an_unreadable_handoff(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    stage_two = _stage_two()
    output = tmp_path / "prepared.xlsb"
    output.write_bytes(b"workbook")
    draft = _draft(tmp_path / "Draft.xlsm", output.name)
    _patch_stage_one_handoff(monkeypatch, stage_two, draft)
    original_open = Path.open

    def deny_workbook_read(path: Path, *args: Any, **kwargs: Any) -> Any:
        if path == output:
            raise PermissionError("Unreadable fixture")
        return original_open(path, *args, **kwargs)

    monkeypatch.setattr(Path, "open", deny_workbook_read)

    with pytest.raises(BondsLiteError, match="output workbook is unreadable"):
        stage_two.resolve_property_stage_two_inputs(draft)


def test_open_workbook_must_match_the_stage_one_handoff(tmp_path: Path) -> None:
    stage_two = _stage_two()
    expected = tmp_path / "prepared.xlsb"
    actual = tmp_path / "unexpected.xlsb"
    expected.touch()
    actual.touch()

    with pytest.raises(BondsLiteError, match="unexpected LTM workbook"):
        stage_two._require_expected_workbook(
            SimpleNamespace(api=SimpleNamespace(FullName=str(actual))),
            expected,
            SimpleNamespace(api=SimpleNamespace(FullName=str(tmp_path / "Draft.xlsm"))),
            tmp_path / "Draft.xlsm",
        )


_SHAREPOINT_FOLDER = (
    "https://example.sharepoint.com/sites/Team/Shared Documents/General/Testing"
)


class _AmbiguousSharePointBook:
    """Excel book whose xlwings fullname fails as for duplicate synced names."""

    def __init__(self, url: str) -> None:
        self.api = SimpleNamespace(FullName=url)

    @property
    def fullname(self) -> str:
        raise RuntimeError("Your SharePoint configuration either requires ...")


def test_excel_local_path_maps_sharepoint_url_through_draft_folder(
    tmp_path: Path,
) -> None:
    stage_two = _stage_two()
    caller = _AmbiguousSharePointBook(f"{_SHAREPOINT_FOLDER}/Draft.xlsm")
    workbook = _AmbiguousSharePointBook(
        f"{_SHAREPOINT_FOLDER}/SST%20Output_Test/Property/LTM_Test.xlsb"
    )

    result = stage_two._excel_local_path(workbook, caller, tmp_path / "Draft.xlsm")

    assert result == (tmp_path / "SST Output_Test" / "Property" / "LTM_Test.xlsb").resolve()


def test_require_expected_workbook_accepts_sharepoint_url_with_duplicate_name(
    tmp_path: Path,
) -> None:
    stage_two = _stage_two()
    expected = tmp_path / "SST Output_Test" / "Property" / "LTM_Test.xlsb"
    expected.parent.mkdir(parents=True)
    expected.touch()
    duplicate = expected.parent / "~property-stage1-x" / expected.name
    duplicate.parent.mkdir()
    duplicate.touch()

    stage_two._require_expected_workbook(
        _AmbiguousSharePointBook(
            f"{_SHAREPOINT_FOLDER}/SST Output_Test/Property/LTM_Test.xlsb"
        ),
        expected,
        _AmbiguousSharePointBook(f"{_SHAREPOINT_FOLDER}/Draft.xlsm"),
        tmp_path / "Draft.xlsm",
    )


def test_require_expected_workbook_reports_unmappable_url(tmp_path: Path) -> None:
    stage_two = _stage_two()
    expected = tmp_path / "LTM_Test.xlsb"
    expected.touch()

    with pytest.raises(BondsLiteError, match="Cannot map the LTM workbook"):
        stage_two._require_expected_workbook(
            _AmbiguousSharePointBook("https://other.sharepoint.com/x/LTM_Test.xlsb"),
            expected,
            _AmbiguousSharePointBook(f"{_SHAREPOINT_FOLDER}/Draft.xlsm"),
            tmp_path / "Draft.xlsm",
        )


def test_ensure_workbook_not_open_detects_sharepoint_url_without_fullname(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    import xlwings

    stage_two = _stage_two()
    target = tmp_path / "SST Output_Test" / "Property" / "LTM_Test.xlsb"
    caller = _AmbiguousSharePointBook(f"{_SHAREPOINT_FOLDER}/Draft.xlsm")
    other = _AmbiguousSharePointBook("https://other.sharepoint.com/x/Other.xlsb")
    opened = _AmbiguousSharePointBook(
        f"{_SHAREPOINT_FOLDER}/SST Output_Test/Property/LTM_Test.xlsb"
    )
    monkeypatch.setattr(
        xlwings, "apps", [SimpleNamespace(books=[caller, other, opened])]
    )

    with pytest.raises(BondsLiteError, match="Close LTM_Test.xlsb"):
        stage_two._ensure_workbook_not_open(target, caller, tmp_path / "Draft.xlsm")

    monkeypatch.setattr(xlwings, "apps", [SimpleNamespace(books=[caller, other])])
    stage_two._ensure_workbook_not_open(target, caller, tmp_path / "Draft.xlsm")


def test_verify_ltm_run_requires_fresh_timestamps_and_all_outputs() -> None:
    stage_two = _stage_two()
    invoked = datetime(2026, 11, 7, 2, 0)
    workbook = stage_two_test_workbook(
        start=datetime(2026, 11, 7, 2, 0),
        finish=datetime(2026, 11, 7, 2, 1),
        scenarios=2,
        progress=["Run 1 - End of run!", "Run 2 - End of run!", "End of multirun!"],
        outputs={"O. CFOutput": [[1]], "O. PVOutput": [[1]], "O. ConcRisk": [[1]]},
    )

    assert stage_two.verify_ltm_run(workbook, invoked, 2) is None


def test_verify_ltm_run_rejects_missing_scenario_marker() -> None:
    stage_two = _stage_two()
    workbook = stage_two_test_workbook(
        start=datetime(2026, 11, 7, 2, 0),
        finish=datetime(2026, 11, 7, 2, 1),
        scenarios=2,
        progress=["Run 1 - End of run!", "End of multirun!"],
        outputs={"O. CFOutput": [[1]], "O. PVOutput": [[1]], "O. ConcRisk": [[1]]},
    )

    with pytest.raises(BondsLiteError, match="Run 2 - End of run"):
        stage_two.verify_ltm_run(workbook, datetime(2026, 11, 7, 2), 2)


def test_verify_ltm_run_rejects_wrong_terminal_message() -> None:
    stage_two = _stage_two()
    workbook = stage_two_test_workbook(
        start=datetime(2026, 11, 7, 2, 0),
        finish=datetime(2026, 11, 7, 2, 1),
        scenarios=1,
        progress=["Run 1 - End of run!", "Unexpected"],
        outputs={"O. CFOutput": [[1]], "O. PVOutput": [[1]], "O. ConcRisk": [[1]]},
    )

    with pytest.raises(BondsLiteError, match="End of multirun"):
        stage_two.verify_ltm_run(workbook, datetime(2026, 11, 7, 2), 1)


def test_verify_ltm_run_rejects_stale_timestamps() -> None:
    stage_two = _stage_two()
    workbook = stage_two_test_workbook(
        start=datetime(2026, 11, 7, 1, 0),
        finish=datetime(2026, 11, 7, 1, 1),
        scenarios=1,
        progress=["Run 1 - End of run!", "End of multirun!"],
        outputs={"O. CFOutput": [[1]], "O. PVOutput": [[1]], "O. ConcRisk": [[1]]},
    )

    with pytest.raises(BondsLiteError, match="RunStartTime is stale"):
        stage_two.verify_ltm_run(workbook, datetime(2026, 11, 7, 2), 1)


def test_verify_ltm_run_requires_finish_at_or_after_start() -> None:
    stage_two = _stage_two()
    workbook = stage_two_test_workbook(
        start=datetime(2026, 11, 7, 2, 1),
        finish=datetime(2026, 11, 7, 2),
        scenarios=1,
        progress=["Run 1 - End of run!", "End of multirun!"],
        outputs={"O. CFOutput": [[1]], "O. PVOutput": [[1]], "O. ConcRisk": [[1]]},
    )

    with pytest.raises(BondsLiteError, match="RunEndTime is before"):
        stage_two.verify_ltm_run(workbook, datetime(2026, 11, 7, 2), 1)


@pytest.mark.parametrize(
    "empty_sheet",
    ["O. CFOutput", "O. PVOutput", "O. ConcRisk"],
)
def test_verify_ltm_run_rejects_each_empty_output(empty_sheet: str) -> None:
    stage_two = _stage_two()
    outputs = {
        "O. CFOutput": [[1]],
        "O. PVOutput": [[1]],
        "O. ConcRisk": [[1]],
    }
    outputs[empty_sheet] = [[None, "  "]]
    workbook = stage_two_test_workbook(
        start=datetime(2026, 11, 7, 2),
        finish=datetime(2026, 11, 7, 2, 1),
        scenarios=1,
        progress=["Run 1 - End of run!", "End of multirun!"],
        outputs=outputs,
    )

    with pytest.raises(BondsLiteError, match=f"{empty_sheet} is empty"):
        stage_two.verify_ltm_run(workbook, datetime(2026, 11, 7, 2), 1)


def _run_harness(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    *,
    macro_error: BaseException | None = None,
    complete_run: bool = True,
    saved: bool = True,
    advance_mtime: bool = True,
    local_run: bool = True,
    data_source: str = "SQLserver",
) -> tuple[Any, list[tuple[Any, ...]], list[dict[str, object]], Any, Path]:
    """Run Stage 2 through fake Excel while recording every external action."""
    stage_two = _stage_two()
    target = tmp_path / "prepared.xlsb"
    target.write_bytes(b"prepared workbook")
    original_mtime = target.stat().st_mtime_ns
    events: list[tuple[Any, ...]] = []
    audits: list[dict[str, object]] = []
    invocation_time = datetime(2026, 11, 7, 2, 0)
    workbook = stage_two_test_workbook(
        start=datetime(2026, 11, 7, 1),
        finish=datetime(2026, 11, 7, 1),
        scenarios=2,
        progress=[],
        outputs={
            "O. CFOutput": [[None]],
            "O. PVOutput": [[None]],
            "O. ConcRisk": [[None]],
        },
    )
    workbook.names.update(
        {
            "prm_LocalRun": _name(local_run),
            "prm_DataSource": _name(data_source),
        }
    )
    workbook.api = SimpleNamespace(
        ReadOnly=False,
        Saved=False,
        Close=lambda **kwargs: events.append(("close_workbook", kwargs)),
    )

    def macro(name: str) -> Any:
        def invoke(argument: str) -> None:
            events.append(("macro", name, argument))
            if macro_error is not None:
                raise macro_error
            if not complete_run:
                return
            workbook.names["RunStartTime"].refers_to_range.value = invocation_time
            workbook.names["RunEndTime"].refers_to_range.value = datetime(
                2026, 11, 7, 2, 1
            )
            progress = [
                "Run 1 - End of run!",
                "Run 2 - End of run!",
                "End of multirun!",
            ]
            progress_sheet = workbook.names["ProgressFirstCell"].refers_to_range.sheet
            progress_sheet.api.Cells = lambda _row, _column: SimpleNamespace(
                End=lambda _direction: SimpleNamespace(Row=len(progress))
            )
            progress_sheet.range = lambda *_: SimpleNamespace(
                value=[[message] for message in progress]
            )
            for sheet in workbook.sheets.values():
                sheet.api.UsedRange.Value = [[1]]

        return invoke

    workbook.macro = macro

    def save() -> None:
        events.append(("save",))
        workbook.api.Saved = saved
        target.write_bytes(b"model output")
        if advance_mtime:
            os.utime(
                target,
                ns=(
                    target.stat().st_atime_ns,
                    target.stat().st_mtime_ns + 2_000_000_000,
                ),
            )
        else:
            os.utime(
                target,
                ns=(target.stat().st_atime_ns, original_mtime),
            )

    workbook.save = save
    workbook.fullname = str(target)
    workbook.api.FullName = str(target)

    class FakeApplication:
        def __init__(self) -> None:
            self.api = SimpleNamespace(
                AutomationSecurity=None,
                DisplayAlerts=True,
                # Real Excel has no WorksheetFunction.Now; NOW() is only evaluable.
                Evaluate=lambda formula: (
                    invocation_time if formula == "NOW()" else pytest.fail(formula)
                ),
                CalculateFullRebuild=lambda: events.append(("recalculate",)),
            )
            self.books = SimpleNamespace(
                open=lambda path, **kwargs: (
                    events.append(("open", path, kwargs)),
                    workbook,
                )[1]
            )

        def quit(self) -> None:
            events.append(("quit",))

    app = FakeApplication()
    import xlwings

    monkeypatch.setattr(
        xlwings, "App", lambda **kwargs: (events.append(("app", kwargs)), app)[1]
    )
    monkeypatch.setattr(
        stage_two.shared,
        "workbook_run_guard",
        lambda *_args, **_kwargs: contextlib.nullcontext(),
    )
    monkeypatch.setattr(
        stage_two,
        "_read_property_stage_two_fields",
        lambda _caller: {"last_successful_run": datetime(2026, 10, 1)},
    )
    monkeypatch.setattr(
        stage_two,
        "resolve_property_stage_two_inputs",
        lambda _caller: stage_two.PropertyStageTwoInputs(
            tmp_path / "Draft.xlsm", target
        ),
    )
    monkeypatch.setattr(stage_two, "_check_runner", lambda: None)
    monkeypatch.setattr(stage_two, "_ensure_workbook_not_open", lambda *_args: None)
    monkeypatch.setattr(
        stage_two,
        "write_property_stage_two_audit",
        lambda _caller, fields: (
            audits.append(fields.copy()),
            events.append(("audit", fields["status"])),
        ),
    )
    caller = SimpleNamespace(
        fullname=str(tmp_path / "Draft.xlsm"),
        api=SimpleNamespace(FullName=str(tmp_path / "Draft.xlsm")),
    )
    return stage_two, events, audits, caller, target


def test_run_uses_separate_visible_excel_and_saves_after_verified_macro(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    stage_two, events, audits, caller, target = _run_harness(tmp_path, monkeypatch)

    result = stage_two.run_property_stage_two(caller)

    assert result == target
    assert events[0] == ("audit", "Running")
    assert ("app", {"visible": True, "add_book": False}) in events
    assert events.index(("recalculate",)) < next(
        index for index, event in enumerate(events) if event[0] == "macro"
    )
    assert next(event for event in events if event[0] == "open")[2] == {
        "update_links": False,
        "read_only": False,
        "ignore_read_only_recommended": True,
    }
    assert [event[1:] for event in events if event[0] == "macro"] == [
        ("Module1.CallRunFrozenPython", "flipper")
    ]
    assert events.index(("save",)) > next(
        index for index, event in enumerate(events) if event[0] == "macro"
    )
    assert [audit["status"] for audit in audits] == ["Running", "Succeeded"]
    assert audits[-1]["last_successful_run"] == audits[-1]["finish"]
    assert events[-2:] == [
        ("close_workbook", {"SaveChanges": False}),
        ("quit",),
    ]


def test_run_macro_failure_closes_without_saving_and_retains_last_success(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    stage_two, events, audits, caller, target = _run_harness(
        tmp_path, monkeypatch, macro_error=RuntimeError("VBA failed")
    )

    with pytest.raises(BondsLiteError, match="CallRunFrozenPython failed"):
        stage_two.run_property_stage_two(caller)

    assert not any(event[0] == "save" for event in events)
    assert events[-2:] == [
        ("close_workbook", {"SaveChanges": False}),
        ("quit",),
    ]
    assert audits[-1]["status"] == "Failed"
    assert audits[-1]["last_successful_run"] == datetime(2026, 10, 1)
    assert target.read_bytes() == b"prepared workbook"


def test_run_records_interruption_and_closes_without_saving(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    stage_two, events, audits, caller, target = _run_harness(
        tmp_path, monkeypatch, macro_error=KeyboardInterrupt()
    )

    with pytest.raises(KeyboardInterrupt):
        stage_two.run_property_stage_two(caller)

    assert audits[-1]["status"] == "Interrupted"
    assert audits[-1]["last_successful_run"] == datetime(2026, 10, 1)
    assert not any(event[0] == "save" for event in events)
    assert target.read_bytes() == b"prepared workbook"
    assert events[-2:] == [
        ("close_workbook", {"SaveChanges": False}),
        ("quit",),
    ]


def test_run_rejects_macro_return_without_fresh_completion_markers(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    stage_two, events, audits, caller, target = _run_harness(
        tmp_path, monkeypatch, complete_run=False
    )

    with pytest.raises(BondsLiteError, match="RunStartTime is stale"):
        stage_two.run_property_stage_two(caller)

    assert audits[-1]["status"] == "Failed"
    assert not any(event[0] == "save" for event in events)
    assert target.read_bytes() == b"prepared workbook"


@pytest.mark.parametrize(
    ("saved", "advance_mtime", "message"),
    [
        (False, True, "did not report that it was saved"),
        (True, False, "modification time did not advance"),
    ],
)
def test_run_rejects_failed_save_verification(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    saved: bool,
    advance_mtime: bool,
    message: str,
) -> None:
    stage_two, events, audits, caller, target = _run_harness(
        tmp_path, monkeypatch, saved=saved, advance_mtime=advance_mtime
    )

    with pytest.raises(BondsLiteError, match=message):
        stage_two.run_property_stage_two(caller)

    assert audits[-1]["status"] == "Failed"
    assert audits[-1]["last_successful_run"] == datetime(2026, 10, 1)
    assert target.read_bytes() == b"prepared workbook"
    assert events[-2:] == [
        ("close_workbook", {"SaveChanges": False}),
        ("quit",),
    ]


def test_run_fails_fast_when_runner_is_missing(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    stage_two, events, audits, caller, _target = _run_harness(tmp_path, monkeypatch)
    monkeypatch.setattr(
        stage_two,
        "_check_runner",
        lambda: (_ for _ in ()).throw(BondsLiteError("runner missing")),
    )

    with pytest.raises(BondsLiteError, match="runner missing"):
        stage_two.run_property_stage_two(caller)

    assert not any(event[0] == "app" for event in events)
    assert audits[-1]["status"] == "Failed"


@pytest.mark.parametrize(
    ("local_run", "data_source", "message"),
    [
        (False, "SQLserver", "prm_LocalRun must be TRUE"),
        (True, "External", "prm_DataSource must be SQLserver"),
    ],
)
def test_run_rejects_unsafe_local_run_settings_before_calling_macro(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    local_run: bool,
    data_source: str,
    message: str,
) -> None:
    stage_two, events, audits, caller, _target = _run_harness(
        tmp_path, monkeypatch, local_run=local_run, data_source=data_source
    )

    with pytest.raises(BondsLiteError, match=message):
        stage_two.run_property_stage_two(caller)

    assert not any(event[0] == "macro" for event in events)
    assert not any(event[0] == "save" for event in events)
    assert audits[-1]["status"] == "Failed"


def test_run_guard_rejects_a_concurrent_property_stage(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    stage_two, events, audits, caller, _target = _run_harness(tmp_path, monkeypatch)

    @contextlib.contextmanager
    def reject_duplicate(*_args: Any, **_kwargs: Any) -> Any:
        raise BondsLiteError("A stress run is already in progress")
        yield

    monkeypatch.setattr(stage_two.shared, "workbook_run_guard", reject_duplicate)

    with pytest.raises(BondsLiteError, match="already in progress"):
        stage_two.run_property_stage_two(caller)

    assert events == []
    assert audits == []


def test_stage_two_outputs_do_not_include_macro_cleared_noteval_areas() -> None:
    stage_two = _stage_two()

    assert stage_two.LTM_OUTPUT_SHEETS == (
        "O. CFOutput",
        "O. PVOutput",
        "O. ConcRisk",
    )
    assert not any("NoteVal" in sheet for sheet in stage_two.LTM_OUTPUT_SHEETS)


def test_property_stage_two_macro_installer_is_idempotent() -> None:
    stage_two = _stage_two()

    class FakeCodeModule:
        def __init__(self) -> None:
            self.source = "Private Sub LaunchBondsLite(ByVal stage As String)\nEnd Sub"

        @property
        def CountOfLines(self) -> int:
            return len(self.source.splitlines())

        def Lines(self, _start: int, _count: int) -> str:
            return self.source

        def InsertLines(self, _line: int, text: str) -> None:
            self.source += text

    code = FakeCodeModule()
    components = SimpleNamespace(Item=lambda _name: SimpleNamespace(CodeModule=code))
    book = SimpleNamespace(
        api=SimpleNamespace(VBProject=SimpleNamespace(VBComponents=components))
    )

    stage_two._install_property_stage_two_macro(book)
    stage_two._install_property_stage_two_macro(book)

    assert code.source.count("Public Sub RunLTMFlash()") == 1
    assert 'LaunchBondsLite "P2"' in code.source


def test_property_stage_two_macro_installer_rejects_conflicting_entry_point() -> None:
    stage_two = _stage_two()
    source = (
        "Private Sub LaunchBondsLite(ByVal stage As String)\nEnd Sub\n"
        "Public Sub RunLTMFlash()\nLaunchBondsLite 2\nEnd Sub"
    )
    code = SimpleNamespace(
        CountOfLines=len(source.splitlines()),
        Lines=lambda _start, _count: source,
    )
    book = SimpleNamespace(
        api=SimpleNamespace(
            VBProject=SimpleNamespace(
                VBComponents=SimpleNamespace(
                    Item=lambda _name: SimpleNamespace(CodeModule=code)
                )
            )
        )
    )

    with pytest.raises(BondsLiteError, match="unexpected implementation"):
        stage_two._install_property_stage_two_macro(book)


def test_run_reports_saved_workbook_when_final_audit_fails(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    stage_two, events, audits, caller, _target = _run_harness(tmp_path, monkeypatch)

    def fail_success_audit(_caller: Any, fields: dict[str, object]) -> None:
        audits.append(fields.copy())
        if fields["status"] == "Succeeded":
            raise OSError("Draft save failed")
        events.append(("audit", fields["status"]))

    monkeypatch.setattr(stage_two, "write_property_stage_two_audit", fail_success_audit)

    with pytest.raises(
        BondsLiteError,
        match="workbook was verified and saved, but the final.*audit could not be saved",
    ):
        stage_two.run_property_stage_two(caller)

    assert any(event[0] == "save" for event in events)
    assert [audit["status"] for audit in audits] == [
        "Running",
        "Succeeded",
        "Failed",
    ]
    assert audits[-1]["last_successful_run"] == datetime(2026, 10, 1)
    assert events[-2:] == [
        ("close_workbook", {"SaveChanges": False}),
        ("quit",),
    ]


@pytest.mark.excel
@pytest.mark.skipif(
    os.environ.get("BONDS_LITE_RUN_PROPERTY_STAGE_TWO_EXCEL_TESTS") != "1",
    reason="Set BONDS_LITE_RUN_PROPERTY_STAGE_TWO_EXCEL_TESTS=1",
)
def test_excel_property_stage_two_installer_is_idempotent_on_a_temporary_draft(
    tmp_path: Path,
) -> None:
    """Install and reopen controls on a disposable Draft without running LTM."""
    import shutil

    import xlwings as xw

    import sst_property_stage_one as stage_one
    import sst_property_stage_two as stage_two

    project = Path(__file__).resolve().parents[1]
    caller = tmp_path / "Draft Stage 2 installer test.xlsm"
    shutil.copy2(project / "Draft Input spreadsheet for SST.xlsm", caller)
    original_bytes = caller.read_bytes()
    first_backup = stage_two.install_property_stage_two_controls(caller)
    assert first_backup.read_bytes() == original_bytes
    second_backup = stage_two.install_property_stage_two_controls(caller)
    assert second_backup.is_file()

    with xw.App(visible=False, add_book=False) as application:
        application.api.AutomationSecurity = 3
        application.api.EnableEvents = False
        book = application.books.open(str(caller), update_links=False)
        try:
            stage_one.require_property_stage_one_panel(book)
            stage_one_panel = book.names[stage_one.PROPERTY_PANEL_NAME].refers_to_range
            stage_one_geometry = (
                stage_one_panel.row,
                stage_one_panel.column,
                stage_one_panel.rows.count,
                stage_one_panel.columns.count,
            )
            stage_one_bottom = stage_one_panel.top + stage_one_panel.height
            stage_one_button = book.sheets["I. Control"].api.Shapes.Item(
                stage_one.PROPERTY_BUTTON_NAME
            )
            stage_one_button_geometry = (
                stage_one_button.Left,
                stage_one_button.Top,
                stage_one_button.Width,
                stage_one_button.Height,
            )
            assert stage_two.require_property_stage_two_panel(book) is None
            sheet = book.sheets["I. Control"]
            panel = book.names[stage_two.STAGE_TWO_PANEL_NAME].refers_to_range
            button = sheet.api.Buttons().Item(stage_two.STAGE_TWO_BUTTON_NAME)
            stage_two_panel = (
                panel.row,
                panel.column,
                panel.rows.count,
                panel.columns.count,
            )
            stage_two_button_geometry = (
                button.Left,
                button.Top,
                button.Width,
                button.Height,
            )
            assert stage_two_panel[2:] == (11, 8)
            assert panel.top > stage_one_bottom
            assert stage_two_button_geometry[0] == stage_one_button_geometry[0]
            assert stage_two_button_geometry[2:] == stage_one_button_geometry[2:]

            panel.api.Cut(Destination=sheet.range("AB160:AI170").api)
            button.Left += 20
            button.Top += 30
            moved_panel = book.names[stage_two.STAGE_TWO_PANEL_NAME].refers_to_range
            moved_panel_geometry = (
                moved_panel.row,
                moved_panel.column,
                moved_panel.rows.count,
                moved_panel.columns.count,
            )
            moved_button_geometry = (
                button.Left,
                button.Top,
                button.Width,
                button.Height,
            )
            names = {
                name.name: name.refers_to
                for name in book.names
                if name.name.startswith("SST_PropStage2_")
            }
            stage_two._initialise_property_stage_two_panel(book)
            stage_two._install_property_stage_two_button(book)
            assert (
                book.names[stage_two.STAGE_TWO_PANEL_NAME].refers_to_range.row,
                book.names[stage_two.STAGE_TWO_PANEL_NAME].refers_to_range.column,
                book.names[stage_two.STAGE_TWO_PANEL_NAME].refers_to_range.rows.count,
                book.names[
                    stage_two.STAGE_TWO_PANEL_NAME
                ].refers_to_range.columns.count,
            ) == moved_panel_geometry
            assert (
                button.Left,
                button.Top,
                button.Width,
                button.Height,
            ) == moved_button_geometry
            book.save()

            book.close()
            book = application.books.open(
                str(caller), update_links=False, read_only=True
            )
            assert stage_two.require_property_stage_two_panel(book) is None
            moved_panel = book.names[stage_two.STAGE_TWO_PANEL_NAME].refers_to_range
            button = (
                book.sheets["I. Control"]
                .api.Buttons()
                .Item(stage_two.STAGE_TWO_BUTTON_NAME)
            )
            assert (
                panel.row,
                panel.column,
                panel.rows.count,
                panel.columns.count,
            ) == moved_panel_geometry
            assert (
                button.Left,
                button.Top,
                button.Width,
                button.Height,
            ) == moved_button_geometry
            assert stage_one_geometry == (
                book.names[stage_one.PROPERTY_PANEL_NAME].refers_to_range.row,
                book.names[stage_one.PROPERTY_PANEL_NAME].refers_to_range.column,
                book.names[stage_one.PROPERTY_PANEL_NAME].refers_to_range.rows.count,
                book.names[stage_one.PROPERTY_PANEL_NAME].refers_to_range.columns.count,
            )
            stage_one_button = book.sheets["I. Control"].api.Shapes.Item(
                stage_one.PROPERTY_BUTTON_NAME
            )
            assert (
                stage_one_button.Left,
                stage_one_button.Top,
                stage_one_button.Width,
                stage_one_button.Height,
            ) == stage_one_button_geometry
            installed_names = {
                name.name: name.refers_to
                for name in book.names
                if name.name.startswith("SST_PropStage2_")
            }
            assert installed_names == names
            module = book.api.VBProject.VBComponents.Item(
                "modBondsLiteLauncher"
            ).CodeModule
            source = module.Lines(1, module.CountOfLines)
            assert 'LaunchBondsLite "P2"' in source
            assert "Public Sub RunLTMFlash()" in source
            assert button.Caption == "Run LTM Flash"
            assert button.OnAction.endswith("!RunLTMFlash")
            assert book.api.Saved
        finally:
            book.close()


def stage_two_test_workbook(
    *,
    start: datetime,
    finish: datetime,
    scenarios: int,
    progress: list[str],
    outputs: dict[str, list[list[object]]],
) -> Any:
    """Build the small named-range/sheet surface used by run verification."""
    names = {
        "RunStartTime": _name(start),
        "RunEndTime": _name(finish),
        "prm_NoOfScens": _name(scenarios),
        "ProgressFirstCell": _name(progress[0] if progress else None),
    }
    progress_sheet = SimpleNamespace(name="I. Control")
    progress_sheet.api = SimpleNamespace(
        Rows=SimpleNamespace(Count=100),
        Cells=lambda _row, _column: SimpleNamespace(
            End=lambda _direction: SimpleNamespace(Row=len(progress))
        ),
    )
    progress_sheet.range = lambda *_: SimpleNamespace(
        value=[[value] for value in progress]
    )
    names["ProgressFirstCell"].refers_to_range.sheet = progress_sheet
    names["ProgressFirstCell"].refers_to_range.row = 1
    names["ProgressFirstCell"].refers_to_range.column = 1
    names["ProgressFirstCell"].refers_to_range.api = SimpleNamespace(
        Areas=SimpleNamespace(Count=1)
    )
    sheets = {
        name: SimpleNamespace(
            name=name,
            api=SimpleNamespace(
                UsedRange=SimpleNamespace(Value=values),
            ),
            range=lambda address, values=values: SimpleNamespace(value=values),
        )
        for name, values in outputs.items()
    }
    workbook = SimpleNamespace(names=names, sheets=sheets)
    return workbook
