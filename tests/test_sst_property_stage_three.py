"""Behavioural contracts for preparing the Property Stage 3 NoteVal copy."""

from __future__ import annotations

import os
from collections.abc import Sequence
from datetime import datetime
from pathlib import Path
from types import SimpleNamespace
from typing import Any

import pytest

OLD_SOURCE = r"C:\old folder\old.xlsb"
NEW_SOURCE = Path(r"D:\new folder\flash.xlsb")


def _stage_three() -> Any:
    """Import lazily so the contracts collect before Stage 3 is implemented."""
    import sst_property_stage_three

    return sst_property_stage_three


def _external_formula(sheet: str, expression: str = "CJ3:CL3") -> str:
    return f"=SUM('{OLD_SOURCE[:-len('old.xlsb')]}[old.xlsb]{sheet}'!{expression})"


def test_rewrite_external_reference_cf_formulas() -> None:
    stage_three = _stage_three()
    formulas = [
        _external_formula("O. CFOutput", "CJ3:CL3,-CM3"),
        f"=-'{OLD_SOURCE[:-len('old.xlsb')]}[old.xlsb]O. CFOutput'!CN3",
        f"=-'{OLD_SOURCE[:-len('old.xlsb')]}[old.xlsb]O. CFOutput'!CO3",
        f"=-'{OLD_SOURCE[:-len('old.xlsb')]}[old.xlsb]O. CFOutput'!$CT$3:$DM$722",
    ]

    rewritten = [
        stage_three.rewrite_external_reference(formula, OLD_SOURCE, NEW_SOURCE)
        for formula in formulas
    ]

    assert rewritten == [
        "=SUM('D:\\new folder\\[flash.xlsb]O. CFOutput'!CJ3:CL3,-CM3)",
        "=-'D:\\new folder\\[flash.xlsb]O. CFOutput'!CN3",
        "=-'D:\\new folder\\[flash.xlsb]O. CFOutput'!CO3",
        "=-'D:\\new folder\\[flash.xlsb]O. CFOutput'!$CT$3:$DM$722",
    ]


def test_rewrite_external_reference_pv_formula() -> None:
    stage_three = _stage_three()
    formula = (
        f"=SUM('{OLD_SOURCE[:-len('old.xlsb')]}[old.xlsb]O. PVOutput'!$CJ$3:$CL$3)"
        f"-SUM('{OLD_SOURCE[:-len('old.xlsb')]}[old.xlsb]O. PVOutput'!$CM$3:$CN$3)"
    )

    rewritten = stage_three.rewrite_external_reference(formula, OLD_SOURCE, NEW_SOURCE)

    assert rewritten == (
        "=SUM('D:\\new folder\\[flash.xlsb]O. PVOutput'!$CJ$3:$CL$3)"
        "-SUM('D:\\new folder\\[flash.xlsb]O. PVOutput'!$CM$3:$CN$3)"
    )


def test_rewrite_handles_apostrophes_and_spaces_in_paths() -> None:
    stage_three = _stage_three()
    old_source = r"C:\old user's files\old.xlsb"
    new_source = Path(r"D:\new user's files\flash model.xlsb")
    formula = "=SUM('C:\\old user''s files\\[old.xlsb]O. CFOutput'!CJ3:CL3)"

    rewritten = stage_three.rewrite_external_reference(formula, old_source, new_source)

    assert rewritten == (
        "=SUM('D:\\new user''s files\\[flash model.xlsb]O. CFOutput'!CJ3:CL3)"
    )


def test_formula_ranges_ignore_numbers_in_external_paths() -> None:
    stage_three = _stage_three()
    formula = (
        "=SUM('C:\\Folder Q2\\[Model v3.2.xlsb]O. PVOutput'!$CJ$3:$CL$3)"
        "-SUM('C:\\Folder Q2\\[Model v3.2.xlsb]O. PVOutput'!$CM$3:$CN$3)"
    )

    assert stage_three._formula_ranges(formula) == ["CJ3:CL3", "CM3:CN3"]


def _matrix_range(
    values: Sequence[Sequence[object]], row: int, column: int, rows: int, columns: int
) -> Any:
    region = SimpleNamespace(
        value=values,
        api=SimpleNamespace(
            Row=row,
            Column=column,
            Rows=SimpleNamespace(Count=rows),
            Columns=SimpleNamespace(Count=columns),
        ),
    )
    region.options = lambda **_: region
    return region


def _i10_verification_case(
    source: Sequence[Sequence[object]],
    output: Sequence[Sequence[object]],
    *,
    source_address: str = "CT3:CU4",
    has_spill: bool = True,
    spill_rows: int | None = None,
    spill_columns: int | None = None,
    spill_row: int = 10,
    spill_column: int = 9,
) -> tuple[Any, Any, Any, Any, datetime]:
    stage_three = _stage_three()
    rows, columns = len(source), len(source[0])
    source_region = _matrix_range(source, 3, 98, rows, columns)
    output_region = _matrix_range(output, 10, 9, rows, columns)
    inputs = SimpleNamespace(name="I. Inputs")
    formula = (
        f"=-'{OLD_SOURCE[:-len('old.xlsb')]}[old.xlsb]" f"O. CFOutput'!{source_address}"
    )
    anchor = SimpleNamespace(
        sheet=inputs,
        formula=stage_three.rewrite_external_reference(formula, OLD_SOURCE, NEW_SOURCE),
        value=output[0][0],
        api=SimpleNamespace(
            Row=10,
            Column=9,
            HasSpill=has_spill,
            SpillingToRange=SimpleNamespace(
                Row=spill_row,
                Column=spill_column,
                Rows=SimpleNamespace(Count=rows if spill_rows is None else spill_rows),
                Columns=SimpleNamespace(
                    Count=columns if spill_columns is None else spill_columns
                ),
            ),
        ),
    )

    def input_range(start: object, end: object = None) -> Any:
        if start == "I10":
            return anchor
        if start == (10, 9) and end == (9 + rows, 8 + columns):
            return output_region
        raise AssertionError(f"Unexpected NoteVal range: {start!r}, {end!r}")

    def flash_range(address: str) -> Any:
        if address == source_address:
            return source_region
        if address == source_address.split(":", 1)[0]:
            return SimpleNamespace(value=source[0][0])
        raise AssertionError(f"Unexpected Flash range: {address}")

    inputs.range = input_range
    flash_sheet = SimpleNamespace(name="O. CFOutput", range=flash_range)
    valuation_date = datetime(2025, 12, 31)
    date_range = SimpleNamespace(
        value=valuation_date,
        size=1,
        api=SimpleNamespace(Areas=SimpleNamespace(Count=1)),
    )
    noteval = SimpleNamespace(
        sheets={"I. Inputs": inputs},
        names={
            stage_three.NOTEVAL_DATE_NAME: SimpleNamespace(refers_to_range=date_range)
        },
        api=SimpleNamespace(LinkSources=lambda _: [str(NEW_SOURCE)]),
    )
    flash = SimpleNamespace(
        fullname=str(NEW_SOURCE), sheets={"O. CFOutput": flash_sheet}
    )
    snapshot = stage_three.FlashLinkSnapshot(
        old_source=OLD_SOURCE,
        link_sources=(OLD_SOURCE,),
        other_sources=(),
        formulas=(("I. Inputs", "I10", formula),),
    )
    return stage_three, noteval, snapshot, flash, valuation_date


def test_verify_noteval_accepts_entire_i10_spill() -> None:
    case = _i10_verification_case(
        [[2.0, 3.0], [4.0, 5.0]], [[-2.0, -3.0], [-4.0, -5.0]]
    )
    stage_three, *arguments = case

    stage_three.verify_noteval(*arguments)


def test_verify_noteval_reads_i10_source_coordinates_once() -> None:
    class CountingPosition:
        def __init__(self) -> None:
            self.Rows = SimpleNamespace(Count=2)
            self.Columns = SimpleNamespace(Count=2)
            self.row_reads = 0
            self.column_reads = 0

        @property
        def Row(self) -> int:
            self.row_reads += 1
            return 3

        @property
        def Column(self) -> int:
            self.column_reads += 1
            return 98

    case = _i10_verification_case(
        [[2.0, 3.0], [4.0, 5.0]], [[-2.0, -3.0], [-4.0, -5.0]]
    )
    stage_three, noteval, snapshot, flash, valuation_date = case
    position = CountingPosition()
    flash.sheets["O. CFOutput"].range("CT3:CU4").api = position

    stage_three.verify_noteval(noteval, snapshot, flash, valuation_date)

    assert position.row_reads == 1
    assert position.column_reads == 1


def test_verify_noteval_detects_i10_spill_mismatch_away_from_anchor() -> None:
    case = _i10_verification_case(
        [[2.0, 3.0], [4.0, 5.0]], [[-2.0, -3.0], [-4.0, -6.0]]
    )
    stage_three, *arguments = case

    with pytest.raises(stage_three.PropertyStageThreeError, match="J11"):
        stage_three.verify_noteval(*arguments)


@pytest.mark.parametrize(
    ("has_spill", "rows", "columns", "row", "column"),
    [
        (False, 2, 2, 10, 9),
        (True, 1, 2, 10, 9),
        (True, 2, 1, 10, 9),
        (True, 2, 2, 11, 9),
        (True, 2, 2, 10, 10),
    ],
)
def test_verify_noteval_rejects_missing_or_incorrect_i10_spill(
    has_spill: bool, rows: int, columns: int, row: int, column: int
) -> None:
    case = _i10_verification_case(
        [[2.0, 3.0], [4.0, 5.0]],
        [[-2.0, -3.0], [-4.0, -5.0]],
        has_spill=has_spill,
        spill_rows=rows,
        spill_columns=columns,
        spill_row=row,
        spill_column=column,
    )
    stage_three, *arguments = case

    with pytest.raises(stage_three.PropertyStageThreeError, match=r"I\. Inputs!I10"):
        stage_three.verify_noteval(*arguments)


@pytest.mark.parametrize(
    ("invalid_range", "address"), [("source", "CU4"), ("output", "J11")]
)
@pytest.mark.parametrize("invalid_value", [None, True, float("nan"), "#REF!"])
def test_verify_noteval_rejects_non_numeric_i10_range_value(
    invalid_range: str, address: str, invalid_value: object
) -> None:
    source: list[list[object]] = [[2.0, 3.0], [4.0, 5.0]]
    output: list[list[object]] = [[-2.0, -3.0], [-4.0, -5.0]]
    (source if invalid_range == "source" else output)[1][1] = invalid_value
    case = _i10_verification_case(source, output)
    stage_three, *arguments = case

    with pytest.raises(stage_three.PropertyStageThreeError, match=address):
        stage_three.verify_noteval(*arguments)


def test_verify_noteval_checks_last_cell_of_full_i10_spill() -> None:
    source: list[list[float]] = [
        [float(1 + row * 20 + column) for column in range(20)] for row in range(720)
    ]
    output: list[list[float]] = [[-value for value in row] for row in source]
    output[-1][-1] = -14399.0
    case = _i10_verification_case(source, output, source_address="CT3:DM722")
    stage_three, *arguments = case

    with pytest.raises(stage_three.PropertyStageThreeError, match="AB729"):
        stage_three.verify_noteval(*arguments)


@pytest.mark.parametrize("missing", ["row", "column"])
def test_verify_noteval_rejects_incomplete_i10_spill_values(missing: str) -> None:
    output: list[list[object]] = [[-2.0, -3.0], [-4.0, -5.0]]
    case = _i10_verification_case([[2.0, 3.0], [4.0, 5.0]], output)
    if missing == "row":
        output.pop()
    else:
        output[-1].pop()
    stage_three, *arguments = case

    with pytest.raises(stage_three.PropertyStageThreeError, match="range"):
        stage_three.verify_noteval(*arguments)


def _cell_address(row: int, column: int) -> str:
    letters = ""
    while column:
        column, remainder = divmod(column - 1, 26)
        letters = chr(65 + remainder) + letters
    return f"{letters}{row}"


def _formula_sheet(name: str, cells: dict[str, str]) -> Any:
    areas = []
    for address, formula in cells.items():
        letters = "".join(character for character in address if character.isalpha())
        digits = "".join(character for character in address if character.isdigit())
        column = 0
        for letter in letters:
            column = column * 26 + ord(letter.upper()) - 64
        areas.append(
            SimpleNamespace(
                Row=int(digits),
                Column=column,
                Formula=[[formula]],
            )
        )
    used_range = SimpleNamespace(
        HasFormula=bool(areas),
        SpecialCells=lambda _: SimpleNamespace(
            Areas=SimpleNamespace(
                Count=len(areas),
                Item=lambda index: areas[index - 1],
            )
        ),
    )
    return SimpleNamespace(
        name=name,
        api=SimpleNamespace(UsedRange=used_range),
    )


def _scope_book(inputs: dict[str, str], extra: dict[str, str] | None = None) -> Any:
    control = _formula_sheet(
        "I. Control",
        {
            "D12": (
                f"=SUM('{OLD_SOURCE[:-len('old.xlsb')]}[old.xlsb]O. PVOutput'!$CJ$3:$CL$3)"
                f"-SUM('{OLD_SOURCE[:-len('old.xlsb')]}[old.xlsb]O. PVOutput'!$CM$3:$CN$3)"
            )
        },
    )
    sheets = [
        control,
        _formula_sheet("I. Inputs", inputs),
    ]
    if extra:
        sheets.append(_formula_sheet("Other", extra))
    sources = [OLD_SOURCE, r"C:\other\another.xlsb"]
    return SimpleNamespace(
        sheets=sheets,
        names=[],
        api=SimpleNamespace(LinkSources=lambda _: sources),
    )


def test_scope_check_accepts_exact_reference_set() -> None:
    stage_three = _stage_three()
    book = _scope_book(
        {
            "C10": _external_formula("O. CFOutput", "CJ3:CL3,-CM3"),
            "E10": f"=-'{OLD_SOURCE[:-len('old.xlsb')]}[old.xlsb]O. CFOutput'!CN3",
            "G10": f"=-'{OLD_SOURCE[:-len('old.xlsb')]}[old.xlsb]O. CFOutput'!CO3",
            "I10": f"=-'{OLD_SOURCE[:-len('old.xlsb')]}[old.xlsb]O. CFOutput'!$CT$3:$DM$722",
        }
    )

    snapshot = stage_three.snapshot_flash_links(book)

    assert len(snapshot.formulas) == 5
    assert {sheet for sheet, _, _ in snapshot.formulas} == {"I. Control", "I. Inputs"}
    assert snapshot.old_source == OLD_SOURCE


def test_scope_check_ignores_flash_reference_on_unrelated_sheet() -> None:
    stage_three = _stage_three()
    book = _scope_book(
        {
            "C10": _external_formula("O. CFOutput"),
            "E10": _external_formula("O. CFOutput"),
            "G10": _external_formula("O. CFOutput"),
            "I10": _external_formula("O. CFOutput"),
        },
        {"A1": _external_formula("O. CFOutput")},
    )

    snapshot = stage_three.snapshot_flash_links(book)

    assert len(snapshot.formulas) == 5
    assert all(sheet != "Other" for sheet, _, _ in snapshot.formulas)


def test_formula_scan_does_not_access_unrelated_sheet_used_range() -> None:
    stage_three = _stage_three()
    target_sheet = _formula_sheet(
        "i. inputs", {"C10": _external_formula("O. CFOutput")}
    )

    class UnrelatedSheet:
        name = "Other"

        @property
        def api(self) -> Any:
            raise AssertionError("Unrelated sheet UsedRange must not be accessed")

    book = SimpleNamespace(sheets=[target_sheet, UnrelatedSheet()])

    assert stage_three._iter_formula_cells(book) == [
        ("i. inputs", "C10", _external_formula("O. CFOutput"))
    ]


def test_scope_check_rejects_second_source_in_scope_cells() -> None:
    stage_three = _stage_three()
    book = _scope_book(
        {
            "C10": "=SUM('C:\\other\\[another.xlsb]O. CFOutput'!CJ3:CL3)",
            "E10": _external_formula("O. CFOutput"),
            "G10": _external_formula("O. CFOutput"),
            "I10": _external_formula("O. CFOutput"),
        }
    )

    with pytest.raises(stage_three.PropertyStageThreeError, match="I. Inputs!C10"):
        stage_three.snapshot_flash_links(book)


def test_formula_scan_reads_protected_sheet_used_range() -> None:
    stage_three = _stage_three()
    formula = _external_formula("O. PVOutput")
    used_range = SimpleNamespace(
        HasFormula=True,
        Row=12,
        Column=4,
        Formula=[[formula]],
        SpecialCells=lambda _: (_ for _ in ()).throw(
            RuntimeError("You cannot use this command on a protected sheet")
        ),
    )
    book = SimpleNamespace(
        sheets=[
            SimpleNamespace(
                name="I. Control", api=SimpleNamespace(UsedRange=used_range)
            )
        ]
    )

    formulas = stage_three._iter_formula_cells(book)

    assert formulas == [("I. Control", "D12", formula)]


def _resolver_inputs(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    *,
    stage_two_status: str = "Succeeded",
    stage_two_run: object = None,
    stage_two_input: str | None = None,
    template_name: str = "LTM Note Valuation Tool v1.40.xlsb",
) -> tuple[Any, Path, Path]:
    stage_three = _stage_three()
    draft_path = tmp_path / "Draft.xlsm"
    property_folder = tmp_path / "SST Output_Test" / "Property"
    property_folder.mkdir(parents=True)
    flash_path = property_folder / "JRL_LTM_FlashModel_Test.xlsb"
    flash_path.write_bytes(b"flash")
    template_path = tmp_path / template_name
    template_path.write_bytes(b"template")
    handoff = "SST Output_Test/Property/JRL_LTM_FlashModel_Test.xlsb"
    monkeypatch.setattr(stage_three.shared, "resolve_caller_path", lambda _: draft_path)
    monkeypatch.setattr(stage_three.stage_one, "read_property_flag", lambda _: True)
    monkeypatch.setattr(
        stage_three.stage_one, "require_property_stage_one_panel", lambda _: None
    )
    monkeypatch.setattr(
        stage_three.stage_two, "require_property_stage_two_panel", lambda _: None
    )
    monkeypatch.setattr(
        stage_three.stage_one,
        "_read_property_stage_one_fields",
        lambda _: {
            "status": "Succeeded",
            "last_output": handoff,
            "finish": datetime(2026, 11, 11, 10),
        },
    )
    run_time = stage_two_run or datetime(2026, 11, 11, 11)
    monkeypatch.setattr(
        stage_three.stage_two,
        "_read_property_stage_two_fields",
        lambda _: {
            "status": stage_two_status,
            "input_workbook": stage_two_input or handoff,
            "last_successful_run": run_time,
        },
    )
    monkeypatch.setattr(
        stage_three,
        "_named_value",
        lambda _, name, kind: {
            stage_three.NOTEVAL_TEMPLATE_NAME: template_path.name,
            "I_Valuation_Date": datetime(2025, 12, 31),
        }[name],
    )
    monkeypatch.setattr(stage_three.shared, "read_entity", lambda _: "JRL")
    monkeypatch.setattr(stage_three.shared, "read_output_suffix", lambda _: "Test")
    return stage_three, flash_path, template_path


@pytest.mark.parametrize(
    ("stage_two_status", "stage_two_run", "message"),
    [
        ("Failed", None, "status Succeeded"),
        ("Succeeded", None, "last successful run"),
    ],
)
def test_resolve_stage_three_requires_successful_stage_two(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    stage_two_status: str,
    stage_two_run: object,
    message: str,
) -> None:
    stage_three, _, _ = _resolver_inputs(
        tmp_path,
        monkeypatch,
        stage_two_status=stage_two_status,
        stage_two_run=stage_two_run,
    )
    if stage_two_status == "Succeeded":
        monkeypatch.setattr(
            stage_three.stage_two,
            "_read_property_stage_two_fields",
            lambda _: {
                "status": "Succeeded",
                "input_workbook": "SST Output_Test/Property/JRL_LTM_FlashModel_Test.xlsb",
                "last_successful_run": None,
            },
        )
    caller = SimpleNamespace()

    with pytest.raises(stage_three.PropertyStageThreeError, match=message):
        stage_three.resolve_property_stage_three_inputs(caller)


@pytest.mark.parametrize("earlier", [True, False])
def test_resolve_stage_three_requires_matching_handoffs(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, earlier: bool
) -> None:
    handoff = (
        "SST Output_Test/Property/other.xlsb"
        if not earlier
        else "SST Output_Test/Property/JRL_LTM_FlashModel_Test.xlsb"
    )
    stage_three, flash_path, _ = _resolver_inputs(
        tmp_path,
        monkeypatch,
        stage_two_run=(
            datetime(2026, 11, 11, 9) if earlier else datetime(2026, 11, 11, 11)
        ),
        stage_two_input=handoff,
    )
    if not earlier:
        (tmp_path / "SST Output_Test" / "Property" / "other.xlsb").write_bytes(b"other")
    caller = SimpleNamespace()

    with pytest.raises(stage_three.PropertyStageThreeError, match="Property Stage"):
        stage_three.resolve_property_stage_three_inputs(caller)


def test_resolve_stage_three_output_beside_flash(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    stage_three, flash_path, _ = _resolver_inputs(tmp_path, monkeypatch)

    inputs = stage_three.resolve_property_stage_three_inputs(SimpleNamespace())

    assert inputs.output_path == (
        flash_path.parent / "JRL_LTM Note Valuation Tool v1.40_Test.xlsb"
    )


def test_resolve_stage_three_rejects_template_equal_to_output(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    stage_three, flash_path, _ = _resolver_inputs(
        tmp_path,
        monkeypatch,
        template_name="placeholder.xlsb",
    )
    output_template = flash_path.parent / "JRL_LTM Note Valuation Tool v1.40_Test.xlsb"
    output_template.write_bytes(b"template")
    monkeypatch.setattr(
        stage_three.shared,
        "entity_output_name",
        lambda entity, template, suffix: output_template.name,
    )
    monkeypatch.setattr(
        stage_three,
        "_named_value",
        lambda _, name, kind: {
            stage_three.NOTEVAL_TEMPLATE_NAME: str(output_template),
            "I_Valuation_Date": datetime(2025, 12, 31),
        }[name],
    )

    with pytest.raises(stage_three.PropertyStageThreeError, match="cannot be"):
        stage_three.resolve_property_stage_three_inputs(SimpleNamespace())


def _run_harness(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    *,
    verify_error: BaseException | None = None,
) -> tuple[Any, Any, list[tuple[str, Any]], Path, Path]:
    import contextlib
    import os

    stage_three = _stage_three()
    property_folder = tmp_path / "SST Output_Test" / "Property"
    property_folder.mkdir(parents=True)
    flash_path = property_folder / "JRL_LTM_FlashModel_Test.xlsb"
    flash_path.write_bytes(b"flash")
    template = tmp_path / "NoteVal.xlsb"
    template.write_bytes(b"template")
    output = property_folder / "JRL_NoteVal_Test.xlsb"
    output.write_bytes(b"previous output")
    events: list[tuple[str, Any]] = []
    application_api = SimpleNamespace(
        AutomationSecurity=1,
        EnableEvents=True,
        DisplayAlerts=True,
        ScreenUpdating=True,
        CalculateBeforeSave=True,
        Calculation=1,
        CalculateFull=lambda: events.append(("calculate", None)),
    )
    application = SimpleNamespace(api=application_api)
    caller = SimpleNamespace(
        fullname=str((tmp_path / "Draft.xlsm").resolve()), app=application
    )
    inputs = SimpleNamespace(
        caller_path=Path(caller.fullname),
        flash_path=flash_path,
        template_path=template,
        output_path=output,
        valuation_date=datetime(2025, 12, 31),
    )
    noteval_api = SimpleNamespace(
        Saved=False,
        ChangeLink=lambda *args: events.append(("change_link", args)),
        Run=lambda *args: pytest.fail("Stage 3 must never invoke NoteVal Run"),
    )

    def save() -> None:
        events.append(("save", None))
        noteval_api.Saved = True

    noteval = SimpleNamespace(api=noteval_api, save=save)
    flash = SimpleNamespace(fullname=str(flash_path))

    @contextlib.contextmanager
    def owned_book(_: Any, path: Path, read_only: bool = True) -> Any:
        if path.parent.name.startswith("~property-stage3-"):
            events.append(("open_noteval", (path, application_api.AutomationSecurity)))
            yield noteval
            events.append(("close_noteval", None))
        else:
            events.append(("open_flash", (path, read_only)))
            yield flash
            events.append(("close_flash", None))

    monkeypatch.setattr(
        stage_three.shared, "resolve_caller_path", lambda _: Path(caller.fullname)
    )
    monkeypatch.setattr(
        stage_three.shared,
        "workbook_run_guard",
        lambda *args, **kwargs: contextlib.nullcontext(),
    )
    monkeypatch.setattr(stage_three.shared, "owned_book", owned_book)
    monkeypatch.setattr(
        stage_three,
        "resolve_property_stage_three_inputs",
        lambda _: inputs,
    )
    monkeypatch.setattr(
        stage_three,
        "_read_property_stage_three_fields",
        lambda _: {"last_successful_output": "previous.xlsb"},
    )
    monkeypatch.setattr(
        stage_three.stage_two, "_ensure_workbook_not_open", lambda *args: None
    )
    monkeypatch.setattr(
        stage_three,
        "write_property_stage_three_audit",
        lambda _, fields: events.append(("audit", fields.copy())),
    )
    monkeypatch.setattr(
        stage_three,
        "snapshot_flash_links",
        lambda _: events.append(("snapshot", None))
        or SimpleNamespace(old_source=r"C:\old\old.xlsb"),
    )
    monkeypatch.setattr(
        stage_three,
        "check_stressed_scenario",
        lambda _: events.append(("scenario_check", None)),
    )
    monkeypatch.setattr(
        stage_three,
        "set_noteval_valuation_date",
        lambda *_: events.append(("set_date", None)),
    )
    if verify_error is None:
        monkeypatch.setattr(
            stage_three,
            "verify_noteval",
            lambda *_: events.append(("verify", None)),
        )
    else:

        def verify(*_: Any) -> None:
            events.append(("verify", None))
            raise verify_error

        monkeypatch.setattr(stage_three, "verify_noteval", verify)
    real_replace = os.replace

    def record_replace(
        source: str | os.PathLike[str], target: str | os.PathLike[str]
    ) -> None:
        events.append(("replace", (Path(source), Path(target))))
        real_replace(source, target)

    monkeypatch.setattr(stage_three.os, "replace", record_replace)
    return stage_three, caller, events, output, flash_path


def test_run_repoints_sets_date_verifies_and_saves(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    stage_three, caller, events, output, flash_path = _run_harness(
        tmp_path, monkeypatch
    )

    result = stage_three.run_property_stage_three(caller)

    assert result == output
    assert output.read_bytes() == b"template"
    assert [
        event
        for event, _ in events
        if event
        in {
            "open_noteval",
            "snapshot",
            "open_flash",
            "scenario_check",
            "change_link",
            "set_date",
            "calculate",
            "verify",
            "save",
            "replace",
        }
    ] == [
        "open_noteval",
        "snapshot",
        "open_flash",
        "scenario_check",
        "change_link",
        "set_date",
        "calculate",
        "verify",
        "save",
        "replace",
    ]
    opened_noteval = next(value for event, value in events if event == "open_noteval")
    assert opened_noteval[1] == 3
    link_call = next(value for event, value in events if event == "change_link")
    assert link_call == (r"C:\old\old.xlsb", str(flash_path), 1)
    assert not any(event == "run_macro" for event, _ in events)
    assert next(
        fields
        for event, fields in events
        if event == "audit" and fields["status"] == "Succeeded"
    )["last_successful_output"].endswith(output.name)


def test_run_failure_keeps_previous_output_and_records_failed(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    stage_three, caller, events, output, _ = _run_harness(
        tmp_path,
        monkeypatch,
        verify_error=_stage_three().PropertyStageThreeError("verification failed"),
    )

    with pytest.raises(
        stage_three.PropertyStageThreeError, match="verification failed"
    ):
        stage_three.run_property_stage_three(caller)

    assert output.read_bytes() == b"previous output"
    assert not list(output.parent.glob("~property-stage3-*"))
    failure = next(
        fields
        for event, fields in events
        if event == "audit" and fields["status"] == "Failed"
    )
    assert failure["last_successful_output"] == "previous.xlsb"


def test_run_restores_previous_output_if_success_audit_fails(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    stage_three, caller, events, output, _ = _run_harness(tmp_path, monkeypatch)

    def fail_success_audit(_: Any, fields: dict[str, object]) -> None:
        events.append(("audit", fields.copy()))
        if fields["status"] == "Succeeded":
            raise stage_three.PropertyStageThreeError("success audit failed")

    monkeypatch.setattr(
        stage_three, "write_property_stage_three_audit", fail_success_audit
    )

    with pytest.raises(
        stage_three.PropertyStageThreeError, match="success audit failed"
    ):
        stage_three.run_property_stage_three(caller)

    assert output.read_bytes() == b"previous output"
    failure = next(
        fields
        for event, fields in events
        if event == "audit" and fields["status"] == "Failed"
    )
    assert failure["last_successful_output"] == "previous.xlsb"


def test_run_records_interrupted_on_keyboard_interrupt(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    stage_three, caller, events, output, _ = _run_harness(
        tmp_path, monkeypatch, verify_error=KeyboardInterrupt()
    )

    with pytest.raises(KeyboardInterrupt):
        stage_three.run_property_stage_three(caller)

    assert output.read_bytes() == b"previous output"
    failure = next(
        fields
        for event, fields in events
        if event == "audit" and fields["status"] == "Interrupted"
    )
    assert failure["last_successful_output"] == "previous.xlsb"


def test_run_guard_rejects_concurrent_property_stage(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    stage_three, caller, events, _, _ = _run_harness(tmp_path, monkeypatch)

    def reject_guard(*args: Any, **kwargs: Any) -> Any:
        raise stage_three.PropertyStageThreeError("already in progress")

    monkeypatch.setattr(stage_three.shared, "workbook_run_guard", reject_guard)

    with pytest.raises(
        stage_three.PropertyStageThreeError, match="already in progress"
    ):
        stage_three.run_property_stage_three(caller)

    assert not any(event == "audit" for event, _ in events)


def test_run_never_invokes_noteval_macro(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    stage_three, caller, events, _, _ = _run_harness(tmp_path, monkeypatch)

    stage_three.run_property_stage_three(caller)

    opened_noteval = next(value for event, value in events if event == "open_noteval")
    assert opened_noteval[1] == 3
    assert not any(event == "run_macro" for event, _ in events)


def _scenario_workbook(
    cf_scenario: object,
    pv_scenario: object,
    scenario_rows: list[list[object]],
) -> Any:
    def output_sheet(name: str, value: object) -> Any:
        return SimpleNamespace(
            name=name,
            range=lambda address: SimpleNamespace(value=value),
        )

    scenario_sheet = SimpleNamespace(
        name="I. ScenFlipper",
        api=SimpleNamespace(
            UsedRange=SimpleNamespace(Value=scenario_rows),
        ),
    )
    return SimpleNamespace(
        sheets={
            "O. CFOutput": output_sheet("O. CFOutput", cf_scenario),
            "O. PVOutput": output_sheet("O. PVOutput", pv_scenario),
            "I. ScenFlipper": scenario_sheet,
        }
    )


def _scenario_rows(
    scenario_nine: str, scenario_eight: str = "Base"
) -> list[list[object]]:
    rows: list[list[object]] = []
    for scenario in range(1, 10):
        description = (
            scenario_eight
            if scenario == 8
            else scenario_nine if scenario == 9 else "Base"
        )
        rows.append([scenario] + [None] * 8 + [description])
    return rows


def test_stressed_scenario_check_accepts_first_stress_scenario() -> None:
    stage_three = _stage_three()
    workbook = _scenario_workbook(9, 9, _scenario_rows("Interest_Property"))

    stage_three.check_stressed_scenario(workbook)


@pytest.mark.parametrize(
    ("scenario_nine", "scenario_eight"),
    [("Base", "Base"), ("Property", "Property")],
)
def test_stressed_scenario_check_rejects_base_or_shifted_block(
    scenario_nine: str, scenario_eight: str
) -> None:
    stage_three = _stage_three()
    workbook = _scenario_workbook(9, 9, _scenario_rows(scenario_nine, scenario_eight))

    with pytest.raises(stage_three.PropertyStageThreeError, match="scenario 9"):
        stage_three.check_stressed_scenario(workbook)


@pytest.mark.parametrize("pv_scenario", [10, None])
def test_stressed_scenario_check_rejects_mismatched_header(
    pv_scenario: object,
) -> None:
    stage_three = _stage_three()
    workbook = _scenario_workbook(9, pv_scenario, _scenario_rows("Property"))

    with pytest.raises(stage_three.PropertyStageThreeError, match="scenario"):
        stage_three.check_stressed_scenario(workbook)


@pytest.mark.excel
@pytest.mark.skipif(
    os.environ.get("BONDS_LITE_RUN_PROPERTY_STAGE_THREE_EXCEL_TESTS") != "1",
    reason="Set BONDS_LITE_RUN_PROPERTY_STAGE_THREE_EXCEL_TESTS=1",
)
def test_verify_noteval_i10_spill_in_excel(tmp_path: Path) -> None:
    import xlwings as xw

    stage_three = _stage_three()
    flash_path = tmp_path / "Flash.xlsx"
    noteval_path = tmp_path / "NoteVal.xlsx"
    valuation_date = datetime(2025, 12, 31)
    source = [
        [float(1 + row * 20 + column) for column in range(20)] for row in range(720)
    ]
    original_formula = (
        f"=-'{OLD_SOURCE[:-len('old.xlsb')]}[old.xlsb]" "O. CFOutput'!$CT$3:$DM$722"
    )
    snapshot = stage_three.FlashLinkSnapshot(
        old_source=OLD_SOURCE,
        link_sources=(OLD_SOURCE,),
        other_sources=(),
        formulas=(("I. Inputs", "I10", original_formula),),
    )
    application = xw.App(visible=False, add_book=True)
    flash = noteval = flash_sheet = inputs = anchor = None
    try:
        application.api.AutomationSecurity = 3
        application.api.EnableEvents = False
        application.api.Calculation = -4135
        application.api.CalculateBeforeSave = False
        application.display_alerts = False
        flash = application.books.add()
        flash_sheet = flash.sheets[0]
        flash_sheet.name = "O. CFOutput"
        flash_sheet.range("CT3").value = source
        flash.save(str(flash_path))

        noteval = application.books.add()
        inputs = noteval.sheets[0]
        inputs.name = "I. Inputs"
        inputs.range("B2").value = valuation_date
        noteval.names.add(stage_three.NOTEVAL_DATE_NAME, "='I. Inputs'!$B$2")
        inputs.range("I10").api.Formula2 = stage_three.rewrite_external_reference(
            original_formula, OLD_SOURCE, flash_path
        )
        application.api.CalculateFull()
        noteval.save(str(noteval_path))
        noteval.close()
        noteval = application.books.open(
            str(noteval_path), update_links=False, read_only=True
        )
        inputs = noteval.sheets["I. Inputs"]
        anchor = inputs.range("I10").api
        assert anchor.HasSpill
        assert anchor.SpillingToRange.Address == "$I$10:$AB$729"
        actual = inputs.range("I10:AB729").options(ndim=2).value
        assert len(actual) == 720
        for actual_row, source_row in zip(actual, source):
            assert len(actual_row) == 20
            assert actual_row == pytest.approx(
                [-value for value in source_row], rel=1e-6, abs=1e-6
            )
        stage_three.verify_noteval(noteval, snapshot, flash, valuation_date)

        application.api.Calculation = -4135
        flash_sheet.range("DM722").value = source[-1][-1] + 1.0
        with pytest.raises(stage_three.PropertyStageThreeError, match="AB729"):
            stage_three.verify_noteval(noteval, snapshot, flash, valuation_date)
    finally:
        anchor = inputs = flash_sheet = None
        if noteval is not None:
            noteval.close()
            noteval = None
        if flash is not None:
            flash.close()
            flash = None
        application.api.DisplayAlerts = False
        application.api.Quit()


@pytest.mark.excel
@pytest.mark.skipif(
    os.environ.get("BONDS_LITE_RUN_PROPERTY_STAGE_THREE_EXCEL_TESTS") != "1",
    reason="Set BONDS_LITE_RUN_PROPERTY_STAGE_THREE_EXCEL_TESTS=1",
)
def test_stage_three_on_real_noteval_copy(tmp_path: Path) -> None:
    import hashlib
    import shutil

    import xlwings as xw

    stage_three = _stage_three()
    stage_one = stage_three.stage_one
    stage_two = stage_three.stage_two
    project = Path(__file__).resolve().parents[1]
    original_draft = project / "Draft Input spreadsheet for SST.xlsm"
    original_noteval = project / "LTM Note Valuation Tool v1.40.xlsb"
    original_flash = (
        project / "SST Output_Test" / "Property" / "LTM_FlashModel_v3.2a_Test.xlsb"
    )
    hashes = {
        path: hashlib.sha256(path.read_bytes()).digest()
        for path in (original_draft, original_noteval, original_flash)
    }

    draft = tmp_path / original_draft.name
    noteval_template = tmp_path / original_noteval.name
    flash_path = (
        tmp_path / "SST Output_Test" / "Property" / "JRL_LTM_FlashModel_v3.2a_Test.xlsb"
    )
    flash_path.parent.mkdir(parents=True)
    shutil.copy2(original_draft, draft)
    shutil.copy2(original_noteval, noteval_template)
    shutil.copy2(original_flash, flash_path)
    stage_one.install_property_stage_one_controls(draft)
    stage_two.install_property_stage_two_controls(draft)
    stage_three.install_property_stage_three_controls(draft)

    application = xw.App(visible=False, add_book=True)
    try:
        application.api.AutomationSecurity = 3
        application.display_alerts = False
        book = application.books.open(str(draft), update_links=False)
        try:
            book.names["SST_Property_standalone"].refers_to_range.value = True
            book.names["SST_Entity"].refers_to_range.value = "JRL"
            book.names["SST_LTMNote_input_template"].refers_to_range.value = (
                noteval_template.name
            )
            book.names["I_Valuation_Date"].refers_to_range.value = datetime(
                2025, 12, 31
            )
            book.names["SST_output_folder_suffix"].refers_to_range.value = "Test"
            book.names["SST_PropStage1_Status"].refers_to_range.value = "Succeeded"
            book.names["SST_PropStage1_LastSuccessfulOutput"].refers_to_range.value = (
                "SST Output_Test/Property/JRL_LTM_FlashModel_v3.2a_Test.xlsb"
            )
            book.names["SST_PropStage1_Finish"].refers_to_range.value = datetime(
                2026, 11, 11, 10
            )
            book.names["SST_PropStage2_Status"].refers_to_range.value = "Succeeded"
            book.names["SST_PropStage2_InputWorkbook"].refers_to_range.value = (
                "SST Output_Test/Property/JRL_LTM_FlashModel_v3.2a_Test.xlsb"
            )
            book.names["SST_PropStage2_LastSuccessfulRun"].refers_to_range.value = (
                datetime(2026, 11, 11, 11)
            )
            book.save()

            output = stage_three.run_property_stage_three(book)

            assert output == (
                flash_path.parent / "JRL_LTM Note Valuation Tool v1.40_Test.xlsb"
            )
            assert output.is_file()
            with stage_three.shared.owned_book(
                application, output, read_only=True
            ) as prepared:
                links = tuple(str(link) for link in prepared.api.LinkSources(1))
                assert len(links) == 6
                assert any(
                    Path(link.replace("\\", "/")).name == flash_path.name
                    for link in links
                )
                assert stage_three._noteval_date_range(prepared).value == datetime(
                    2025, 12, 31
                )
                inputs = prepared.sheets["I. Inputs"]
                anchor = inputs.range("I10").api
                assert anchor.HasSpill
                assert anchor.SpillingToRange.Address == "$I$10:$AB$729"
                with stage_three.shared.owned_book(
                    application, flash_path, read_only=True
                ) as flash:
                    actual = inputs.range("I10:AB729").options(ndim=2).value
                    source = (
                        flash.sheets["O. CFOutput"]
                        .range("CT3:DM722")
                        .options(ndim=2)
                        .value
                    )
                    assert len(actual) == len(source) == 720
                    for actual_row, source_row in zip(actual, source):
                        assert len(actual_row) == len(source_row) == 20
                        assert actual_row == pytest.approx(
                            [-value for value in source_row], rel=1e-6, abs=1e-6
                        )
        finally:
            book.close()
    finally:
        application.api.DisplayAlerts = False
        application.api.Quit()

    assert all(
        hashlib.sha256(path.read_bytes()).digest() == expected
        for path, expected in hashes.items()
    )
