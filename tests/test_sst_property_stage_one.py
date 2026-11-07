"""Behavioural contracts for the Property Stage 1 flag gate."""

from __future__ import annotations

import os
from datetime import date, datetime
from pathlib import Path
from types import SimpleNamespace
from typing import Any

import pytest

from bonds_lite_input import BondsLiteError

PROPERTY_FLAG = "SST_Property_standalone"


def _stage_one() -> Any:
    """Import lazily so contracts collect before Property Stage 1 is implemented."""
    import sst_property_stage_one

    return sst_property_stage_one


def _named_property_book(value: object = True, *, include_flag: bool = True) -> Any:
    """Expose the Property flag as a scalar name on this workbook's I. Control."""
    book = SimpleNamespace(fullname=os.path.normcase(str(Path("Draft.xlsm").resolve())))
    names: dict[str, Any] = {}
    if include_flag:
        names[PROPERTY_FLAG] = SimpleNamespace(
            refers_to_range=SimpleNamespace(
                value=value,
                size=1,
                sheet=SimpleNamespace(name="I. Control", book=book),
                api=SimpleNamespace(Areas=SimpleNamespace(Count=1)),
            )
        )
    book.names = names
    return book


@pytest.mark.parametrize("value", [True, False])
def test_read_property_flag_accepts_excel_boolean_values(value: bool) -> None:
    stage_one = _stage_one()

    assert stage_one.read_property_flag(_named_property_book(value)) is value


@pytest.mark.parametrize("value", ["TRUE", 1, 0, 1.0, None, ""], ids=repr)
def test_read_property_flag_rejects_non_boolean_values(value: object) -> None:
    stage_one = _stage_one()

    with pytest.raises(BondsLiteError, match=PROPERTY_FLAG):
        stage_one.read_property_flag(_named_property_book(value))


def test_read_property_flag_rejects_a_missing_workbook_name() -> None:
    stage_one = _stage_one()

    with pytest.raises(BondsLiteError, match=PROPERTY_FLAG):
        stage_one.read_property_flag(_named_property_book(include_flag=False))


def test_require_property_selected_only_accepts_true() -> None:
    stage_one = _stage_one()

    assert stage_one.require_property_selected(True) is None
    with pytest.raises(
        BondsLiteError,
        match="Property stress is not selected: set SST_Property_standalone to TRUE",
    ):
        stage_one.require_property_selected(False)


def test_false_property_flag_fails_before_folder_operations_and_preserves_audit_output(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """A disabled Property run is audited as failed without touching its output."""
    stage_one = _stage_one()
    previous_output = "SST Output_Test/Property/LTM_FlashModel_Test.xlsb"
    events: list[tuple[str, dict[str, object]]] = []

    def forbidden_allocate(*args: Any, **kwargs: Any) -> Any:
        pytest.fail("A false Property flag must be checked before folder allocation")

    def record_audit(caller: Any, fields: dict[str, object]) -> None:
        assert caller is book
        events.append(("audit", fields.copy()))

    book = _named_property_book(False)
    monkeypatch.setattr(stage_one, "allocate_stress_folder", forbidden_allocate)
    monkeypatch.setattr(stage_one, "write_property_stage_one_audit", record_audit)
    monkeypatch.setattr(
        stage_one,
        "_read_property_stage_one_fields",
        lambda caller: {"last_output": previous_output},
        raising=False,
    )

    with pytest.raises(BondsLiteError, match="Property stress is not selected"):
        stage_one.run_property_stage_one(book)

    assert events
    assert events[-1][1]["status"] == "Failed"
    assert events[-1][1]["last_output"] == previous_output


def _scenario_grid() -> tuple[list[list[object]], list[list[object]]]:
    values = [
        ["", "detailed", None],
        ["", 3, None],
        ["", 2, None],
        [None, None, None],
        ["Scenario", "Stress", "Combined_Vol"],
        [1, -0.1, 0.9],
        [2, -0.2, 0.8],
        [3, -0.3, 0.7],
    ]
    formulas = [[None for _ in row] for row in values]
    return values, formulas


def test_validate_scenflipper_accepts_consistent_scenarios() -> None:
    stage_one = _stage_one()
    values, formulas = _scenario_grid()

    headers, width = stage_one.validate_scenflipper(values, formulas)

    assert headers == ("Scenario", "Stress", "Combined_Vol")
    assert width == 3


@pytest.mark.parametrize(
    ("row", "column", "value"),
    [
        (0, 1, "TRUE"),
        (1, 1, 2),
        (2, 1, 0),
        (2, 1, 4),
        (4, 1, None),
        (4, 2, "Scenario"),
    ],
)
def test_validate_scenflipper_rejects_invalid_controls_and_headers(
    row: int, column: int, value: object
) -> None:
    stage_one = _stage_one()
    values, formulas = _scenario_grid()
    values[row][column] = value

    with pytest.raises(BondsLiteError):
        stage_one.validate_scenflipper(values, formulas)


def test_validate_scenflipper_rejects_non_contiguous_scenario_rows() -> None:
    stage_one = _stage_one()
    values, formulas = _scenario_grid()
    values[6][0] = None

    with pytest.raises(BondsLiteError, match="contiguous"):
        stage_one.validate_scenflipper(values, formulas)


@pytest.mark.parametrize(
    "formula",
    ["='Other'!A1", "=[linked.xlsx]Sheet1!A1"],
)
def test_validate_scenflipper_rejects_external_formula_references(
    formula: str,
) -> None:
    stage_one = _stage_one()
    values, formulas = _scenario_grid()
    formulas[5][1] = formula

    with pytest.raises(BondsLiteError, match="another sheet or workbook"):
        stage_one.validate_scenflipper(values, formulas)


@pytest.mark.parametrize("error_value", ["#VALUE!", "#N/A", "#NAME?"])
def test_validate_scenflipper_rejects_error_values(error_value: str) -> None:
    stage_one = _stage_one()
    values, formulas = _scenario_grid()
    values[5][1] = error_value

    with pytest.raises(BondsLiteError, match="Excel error"):
        stage_one.validate_scenflipper(values, formulas)


@pytest.mark.parametrize(
    ("value", "expected"),
    [
        (datetime(2025, 12, 31), datetime(2025, 12, 31)),
        (date(2025, 12, 31), datetime(2025, 12, 31)),
    ],
)
def test_date_value_accepts_excel_dates(value: object, expected: datetime) -> None:
    assert _stage_one()._date_value(value) == expected


@pytest.mark.parametrize("value", ["31/12/2025", 46022, None])
def test_date_value_rejects_non_date_inputs(value: object) -> None:
    with pytest.raises(BondsLiteError, match="I_Valuation_Date"):
        _stage_one()._date_value(value)


def test_resolve_template_path_accepts_draft_relative_workbook(
    tmp_path: Path,
) -> None:
    stage_one = _stage_one()
    draft = tmp_path / "Draft.xlsm"
    template = tmp_path / "LTM.xlsb"
    draft.touch()
    template.touch()

    assert stage_one._resolve_template_path("LTM.xlsb", draft) == template


def test_resolve_template_path_rejects_unsupported_extension(
    tmp_path: Path,
) -> None:
    stage_one = _stage_one()
    draft = tmp_path / "Draft.xlsm"
    draft.touch()
    (tmp_path / "LTM.csv").touch()

    with pytest.raises(BondsLiteError, match="must use .xlsb, .xlsm or .xlsx"):
        stage_one._resolve_template_path("LTM.csv", draft)


def _named_valuation_book(value: object, *, writable: bool = True) -> Any:
    class NamedCell:
        size = 1
        api = SimpleNamespace(Areas=SimpleNamespace(Count=1))

        def __init__(self) -> None:
            self._value = value

        @property
        def value(self) -> object:
            return self._value

        @value.setter
        def value(self, new_value: object) -> None:
            if writable:
                self._value = new_value

    cell = NamedCell()
    return SimpleNamespace(names={"ValDate": SimpleNamespace(refers_to_range=cell)})


def test_set_valuation_date_writes_and_reads_back_excel_date() -> None:
    stage_one = _stage_one()
    valuation_date = datetime(2025, 12, 31)
    book = _named_valuation_book(datetime(2024, 12, 31))

    stage_one.set_valuation_date(book, valuation_date)

    assert book.names["ValDate"].refers_to_range.value == valuation_date


def test_set_valuation_date_rejects_a_failed_readback() -> None:
    with pytest.raises(BondsLiteError, match="did not retain the Draft date"):
        _stage_one().set_valuation_date(
            _named_valuation_book(datetime(2024, 12, 31), writable=False),
            datetime(2025, 12, 31),
        )


def test_set_valuation_date_rejects_a_missing_name() -> None:
    with pytest.raises(BondsLiteError, match="missing the ValDate name"):
        _stage_one().set_valuation_date(SimpleNamespace(names={}), datetime.now())


@pytest.mark.parametrize(
    ("source_rows", "source_columns", "target_rows", "target_columns"),
    [(8, 3, 10, 5), (10, 5, 8, 3)],
)
def test_replace_scenflipper_clears_union_and_copies_into_existing_sheet(
    source_rows: int,
    source_columns: int,
    target_rows: int,
    target_columns: int,
) -> None:
    stage_one = _stage_one()
    clear_calls: list[tuple[object, object]] = []
    width_calls: list[tuple[int, object]] = []
    copy_calls: list[object] = []

    class FakeRange:
        def __init__(self, address: object) -> None:
            self.address = address
            self.api = SimpleNamespace(Clear=lambda: clear_calls.append(address))
            self._column_width: object = 0

        @property
        def column_width(self) -> object:
            return self._column_width

        @column_width.setter
        def column_width(self, value: object) -> None:
            self._column_width = value
            width_calls.append((self.address[1], value))

    source_used_range = SimpleNamespace(
        Copy=lambda **kwargs: copy_calls.append(kwargs["Destination"])
    )
    source_sheet = SimpleNamespace(
        api=SimpleNamespace(UsedRange=source_used_range),
        range=lambda address: FakeRange(address),
    )
    target_sheet = SimpleNamespace(
        range=lambda *address: FakeRange(address[0] if len(address) == 1 else address)
    )
    destination_range = target_sheet.range("A1")
    destination = destination_range.api
    target_sheet.range = lambda *address: (
        destination_range
        if address == ("A1",)
        else FakeRange(address[0] if len(address) == 1 else address)
    )

    stage_one.replace_scenflipper(
        source_sheet,
        target_sheet,
        source_rows,
        source_columns,
        target_rows,
        target_columns,
    )

    assert clear_calls == [
        ((1, 1), (max(source_rows, target_rows), max(source_columns, target_columns)))
    ]
    assert copy_calls == [destination]
    assert width_calls == [(column, 0) for column in range(1, source_columns + 1)]


def test_scenflipper_formula_snapshot_records_referencing_cells() -> None:
    stage_one = _stage_one()

    class FormulaRange:
        def __init__(self, formula: object, address: str = "") -> None:
            self.formula = formula
            self._address = address

        def get_address(self) -> str:
            return self._address

    class FormulaSheet:
        name = "I. Inputs"
        api = SimpleNamespace(
            UsedRange=SimpleNamespace(
                Row=1,
                Column=1,
                Rows=SimpleNamespace(Count=1),
                Columns=SimpleNamespace(Count=2),
                HasFormula=True,
                SpecialCells=lambda _: SimpleNamespace(
                    Areas=SimpleNamespace(
                        Count=1,
                        Item=lambda index: SimpleNamespace(
                            Row=1,
                            Column=1,
                            Formula=[
                                [
                                    "=OFFSET('I. ScenFlipper'!$B$5,prm_CurrentScen,0)",
                                    "=1",
                                ]
                            ],
                        ),
                    )
                ),
            )
        )

        def range(self, *addresses: object) -> FormulaRange:
            row, column = addresses[0]
            return FormulaRange(None, f"${chr(64 + column)}${row}")

    assert stage_one._scenflipper_formulas(
        SimpleNamespace(sheets=[FormulaSheet()])
    ) == (
        (
            "I. Inputs",
            "$A$1",
            "=OFFSET('I. ScenFlipper'!$B$5,prm_CurrentScen,0)",
        ),
    )


def test_first_grid_difference_reports_one_based_cell_coordinates() -> None:
    stage_one = _stage_one()

    assert stage_one._first_grid_difference([[1, "changed"]], [[1, "expected"]]) == (
        1,
        2,
        "changed",
        "expected",
    )
    assert stage_one._first_grid_difference([[1]], [[1]]) is None


def test_grid_formulas_normalises_excel_empty_cells() -> None:
    sheet = SimpleNamespace(range=lambda *_: SimpleNamespace(formula=[["=1", ""]]))

    assert _stage_one()._grid_formulas(sheet, 1, 2) == [["=1", None]]


@pytest.mark.excel
@pytest.mark.skipif(
    os.environ.get("BONDS_LITE_RUN_PROPERTY_STAGE_ONE_EXCEL_TESTS") != "1",
    reason="Set BONDS_LITE_RUN_PROPERTY_STAGE_ONE_EXCEL_TESTS=1",
)
def test_excel_property_stage_one_uses_temporary_draft_and_template_copies(
    tmp_path: Path,
) -> None:
    import hashlib
    import shutil

    import xlwings as xw

    stage_one = _stage_one()
    project = Path(__file__).resolve().parents[1]
    original_draft = project / "Draft Input spreadsheet for SST.xlsm"
    original_template = project / "LTM_FlashModel_v3.2a.xlsb"
    draft_hash = hashlib.sha256(original_draft.read_bytes()).digest()
    template_hash = hashlib.sha256(original_template.read_bytes()).digest()
    draft = tmp_path / original_draft.name
    template = tmp_path / original_template.name
    shutil.copy2(original_draft, draft)
    shutil.copy2(original_template, template)

    first_backup = stage_one.install_property_stage_one_controls(draft)
    second_backup = stage_one.install_property_stage_one_controls(draft)
    assert first_backup.is_file()
    assert second_backup.is_file()
    application = xw.App(visible=False, add_book=True)
    try:
        application.api.AutomationSecurity = 3
        application.display_alerts = False
        book = application.books.open(str(draft), update_links=False)
        try:
            book.names[PROPERTY_FLAG].refers_to_range.value = True
            book.names["SST_LTMFlash_input_template"].refers_to_range.value = (
                template.name
            )
            book.names["SST_output_folder_suffix"].refers_to_range.value = (
                "PropertyTest"
            )
            book.save()

            inputs = stage_one.read_property_stage_one_inputs(book)
            with stage_one.shared.owned_book(
                application, template, read_only=True
            ) as template_book:
                template_state = stage_one._snapshot_template(
                    template_book, inputs.headers
                )
            output = stage_one.run_property_stage_one(book)
            assert output == (
                tmp_path
                / "SST Output_PropertyTest"
                / "Property"
                / "LTM_FlashModel_v3.2a_PropertyTest.xlsb"
            )
            assert output.is_file()
            assert (
                book.names["SST_PropStage1_Status"].refers_to_range.value == "Succeeded"
            )
            with stage_one.shared.owned_book(
                application, output, read_only=True
            ) as prepared:
                stage_one.verify_ltm_flash(
                    prepared,
                    inputs,
                    template_state,
                    inputs.valuation_date,
                    draft,
                )
        finally:
            book.close()
    finally:
        application.api.DisplayAlerts = False
        application.api.Quit()

    assert hashlib.sha256(original_draft.read_bytes()).digest() == draft_hash
    assert hashlib.sha256(original_template.read_bytes()).digest() == template_hash
