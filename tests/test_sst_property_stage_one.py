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
INTEREST_FLAG = "SST_Interest_standalone"


def _stage_one() -> Any:
    """Import lazily so contracts collect before Property Stage 1 is implemented."""
    import sst_property_stage_one

    return sst_property_stage_one


def _named_property_book(
    value: object = True, *, include_flag: bool = True, interest: object = False
) -> Any:
    """Expose the stress flags as scalar names on this workbook's I. Control."""
    book = SimpleNamespace(fullname=os.path.normcase(str(Path("Draft.xlsm").resolve())))
    names: dict[str, Any] = {}
    flags = {INTEREST_FLAG: interest}
    if include_flag:
        flags[PROPERTY_FLAG] = value
    for name, flag in flags.items():
        names[name] = SimpleNamespace(
            refers_to_range=SimpleNamespace(
                value=flag,
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


@pytest.mark.parametrize(
    ("interest", "property_", "block"),
    [
        (True, False, "Interest"),
        (False, True, "Property"),
        (True, True, "Interest_Property"),
    ],
)
def test_require_stress_selected_maps_flags_to_block(
    interest: bool, property_: bool, block: str
) -> None:
    assert _stage_one().require_stress_selected(interest, property_) == block


def test_require_stress_selected_rejects_no_flags() -> None:
    with pytest.raises(
        BondsLiteError,
        match="set SST_Interest_standalone and/or SST_Property_standalone to TRUE",
    ):
        _stage_one().require_stress_selected(False, False)


def test_interest_only_flag_does_not_start_property_stage_one(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Interest-only selection is reserved for the Interest workstream."""
    stage_one = _stage_one()
    monkeypatch.setattr(
        stage_one.shared, "resolve_caller_path", lambda caller: Path("Draft.xlsm")
    )

    with pytest.raises(BondsLiteError, match="Property stress is not selected"):
        stage_one.read_property_stage_one_inputs(
            _named_property_book(False, interest=True)
        )


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


def test_stage_one_blank_entity_fails_before_folder_replacement(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    import contextlib

    stage_one = _stage_one()
    draft_path = tmp_path / "Draft.xlsm"
    template = tmp_path / "Flash.xlsb"
    template.write_bytes(b"template")
    book = SimpleNamespace(
        fullname=str(draft_path.resolve()),
        names={},
        sheets={},
        app=SimpleNamespace(api=SimpleNamespace(Calculate=lambda: None)),
    )

    def named_value(value: object) -> Any:
        return SimpleNamespace(
            refers_to_range=SimpleNamespace(
                value=value,
                size=1,
                sheet=SimpleNamespace(name="I. Control", book=book),
                api=SimpleNamespace(Areas=SimpleNamespace(Count=1)),
            )
        )

    book.names.update(
        {
            PROPERTY_FLAG: named_value(True),
            INTEREST_FLAG: named_value(False),
            "SST_Entity": named_value(" "),
            "SST_LTMFlash_input_template": named_value(template.name),
            "I_Valuation_Date": named_value(datetime(2026, 12, 31)),
            "SST_output_folder_suffix": named_value("Test"),
        }
    )
    events: list[dict[str, object]] = []
    monkeypatch.setattr(stage_one.shared, "resolve_caller_path", lambda _: draft_path)
    monkeypatch.setattr(stage_one.shared, "workbook_run_guard", lambda *a, **k: contextlib.nullcontext())
    monkeypatch.setattr(
        stage_one, "_read_property_stage_one_fields", lambda _: {"last_output": "old.xlsb"}
    )
    monkeypatch.setattr(
        stage_one,
        "write_property_stage_one_audit",
        lambda _, fields: events.append(fields.copy()),
    )
    monkeypatch.setattr(
        stage_one,
        "allocate_stress_folder",
        lambda *args, **kwargs: pytest.fail("Must fail before replacing the output folder"),
    )

    with pytest.raises(BondsLiteError, match="SST_Entity"):
        stage_one.run_property_stage_one(book)

    assert events[-1]["status"] == "Failed"
    assert events[-1]["last_output"] == "old.xlsb"


def test_stage_one_output_name_has_entity_prefix(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    import contextlib

    stage_one = _stage_one()
    draft_path = tmp_path / "Draft.xlsm"
    template = tmp_path / "LTM_FlashModel_v3.2a.xlsb"
    template.write_bytes(b"template")
    output_folder = tmp_path / "SST Output_Test" / "Property"
    inputs = stage_one.PropertyStageOneInputs(
        caller_path=draft_path,
        entity="JRL",
        template_path=template,
        valuation_date=datetime(2025, 12, 31),
        output_suffix="Test",
        values=[],
        rows=0,
        columns=0,
        headers=(),
        stress_block="Property",
    )
    caller = SimpleNamespace(
        fullname=str(draft_path),
        app=SimpleNamespace(),
    )
    workbook = SimpleNamespace(
        api=SimpleNamespace(Saved=True),
        save=lambda: None,
        sheets={"I. ScenFlipper": object()},
    )
    audits: list[dict[str, object]] = []

    def allocate(*args: Any, **kwargs: Any) -> tuple[Path, bool, None]:
        output_folder.mkdir(parents=True)
        return output_folder, True, None

    @contextlib.contextmanager
    def owned_book(*args: Any, **kwargs: Any) -> Any:
        yield workbook

    monkeypatch.setattr(stage_one.shared, "resolve_caller_path", lambda _: draft_path)
    monkeypatch.setattr(
        stage_one.shared, "workbook_run_guard", lambda *args, **kwargs: contextlib.nullcontext()
    )
    monkeypatch.setattr(
        stage_one, "_read_property_stage_one_fields", lambda _: {"last_output": None}
    )
    monkeypatch.setattr(stage_one, "read_property_stage_one_inputs", lambda _: inputs)
    monkeypatch.setattr(stage_one, "allocate_stress_folder", allocate)
    monkeypatch.setattr(stage_one.shared, "owned_book", owned_book)
    monkeypatch.setattr(
        stage_one.shared, "excel_state", lambda _: contextlib.nullcontext()
    )
    monkeypatch.setattr(
        stage_one, "_manual_calculation", lambda _: contextlib.nullcontext()
    )
    monkeypatch.setattr(
        stage_one,
        "_snapshot_template",
        lambda *_: SimpleNamespace(rows=0, columns=0),
    )
    monkeypatch.setattr(stage_one, "replace_scenflipper", lambda *args: None)
    monkeypatch.setattr(stage_one, "set_valuation_date", lambda *args: None)
    monkeypatch.setattr(stage_one, "verify_ltm_flash", lambda *args: None)
    monkeypatch.setattr(
        stage_one,
        "write_property_stage_one_audit",
        lambda _, fields: audits.append(fields.copy()),
    )
    monkeypatch.setattr(stage_one, "discard_replaced_stress_folder", lambda _: None)

    output = stage_one.run_property_stage_one(caller)

    assert output == output_folder / "JRL_LTM_FlashModel_v3.2a_Test.xlsb"
    assert audits[-1]["last_output"] == (
        "SST Output_Test/Property/JRL_LTM_FlashModel_v3.2a_Test.xlsb"
    )


def test_entity_control_installs_sst_entity_at_d9() -> None:
    stage_one = _stage_one()

    class Names(list[Any]):
        def add(self, name: str, refers_to: str) -> None:
            self.append(SimpleNamespace(name=name, refers_to=refers_to))

    book = SimpleNamespace(names=Names())

    stage_one._entity_control(book)

    assert book.names[0].name == "SST_Entity"
    assert book.names[0].refers_to == "='I. Control'!$D$9"


def test_entity_control_rejects_an_existing_name_at_another_cell() -> None:
    stage_one = _stage_one()
    book = SimpleNamespace(fullname=str(Path("Draft.xlsm").resolve()))
    region = SimpleNamespace(
        row=10,
        column=4,
        size=1,
        sheet=SimpleNamespace(name="I. Control", book=book),
        api=SimpleNamespace(Areas=SimpleNamespace(Count=1)),
    )

    class Names(list[Any]):
        def __getitem__(self, key: str | int) -> Any:
            if isinstance(key, str):
                return next(name for name in self if name.name == key)
            return super().__getitem__(key)

    book.names = Names(
        [SimpleNamespace(name="SST_Entity", refers_to_range=region)]
    )

    with pytest.raises(BondsLiteError, match="I. Control!\\$D\\$9"):
        stage_one._entity_control(book)


def _scenario_grid() -> list[list[object]]:
    """Two-row Base, Property, Interest and Interest_Property blocks."""
    headers = ["Scenario", "run type"] + [f"h{index}" for index in range(3, 10)]
    headers += ["scen desc", "PC1", "PC2", "PC3", "h14", "h15", "h16", "LSE", "RUK"]

    def row(number: int, desc: str, pc: float = 0, lse: float = 0) -> list[object]:
        cells: list[object] = [number, "IFRS"] + [0] * 7
        return cells + [desc, pc, 0, 0, 0, 0, 0, lse, 0]

    values: list[list[object]] = [
        ["Output type:", "detailed"] + [None] * 16,
        ["Number of scenarios:", 4] + [None] * 16,
        ["Current scenario:", 4] + [None] * 16,
        [None] * 18,
        headers,
        row(1, "Base"),
        row(2, "Base"),
    ]
    for desc, pc, lse in [
        ("Property", 0, -0.1),
        ("Interest", 0.01, 0),
        ("Interest_Property", 0.01, -0.1),
    ]:
        values += [row(3, desc, pc, lse), row(4, desc, pc, lse)]
    return values


@pytest.mark.parametrize(
    ("interest", "property_", "block", "first_row"),
    [
        (False, True, "Property", 7),
        (True, False, "Interest", 9),
        (True, True, "Interest_Property", 11),
    ],
)
def test_validate_scenflipper_selects_base_and_flagged_block(
    interest: bool, property_: bool, block: str, first_row: int
) -> None:
    values = _scenario_grid()

    headers, width, selected, stress_block = _stage_one().validate_scenflipper(
        values, interest, property_
    )

    assert headers[9] == "scen desc"
    assert width == 18
    assert stress_block == block
    assert selected == values[:7] + values[first_row : first_row + 2]


@pytest.mark.parametrize(
    ("interest", "property_", "row", "column", "message"),
    [
        (False, True, 7, 16, "Property stress is selected"),
        (True, False, 10, 10, "Interest stress is selected"),
        (True, True, 11, 16, "Property stress is selected"),
        (True, True, 12, 10, "Interest stress is selected"),
    ],
)
def test_validate_scenflipper_requires_non_zero_stress_per_row(
    interest: bool, property_: bool, row: int, column: int, message: str
) -> None:
    values = _scenario_grid()
    values[row][column] = 0

    with pytest.raises(BondsLiteError, match=message):
        _stage_one().validate_scenflipper(values, interest, property_)


def test_validate_scenflipper_accepts_any_one_non_zero_stress_column() -> None:
    values = _scenario_grid()
    for row in (9, 10):
        values[row][10], values[row][12] = 0, 0.02

    assert _stage_one().validate_scenflipper(values, True, False)[3] == "Interest"


@pytest.mark.parametrize(
    ("interest", "property_", "row", "column", "block"),
    [
        (False, True, 5, 10, "Base"),
        (True, False, 6, 17, "Base"),
        (False, True, 8, 12, "Property"),
        (True, False, 9, 17, "Interest"),
    ],
)
def test_validate_scenflipper_rejects_stress_outside_selected_type(
    interest: bool, property_: bool, row: int, column: int, block: str
) -> None:
    values = _scenario_grid()
    values[row][column] = 0.01

    with pytest.raises(BondsLiteError, match=f"{block} block must not be stressed"):
        _stage_one().validate_scenflipper(values, interest, property_)


@pytest.mark.parametrize(
    ("row", "column", "value", "message"),
    [
        (0, 1, "TRUE", "B1"),
        (1, 1, 6, "B2 must equal"),
        (2, 1, 0, "B3"),
        (2, 1, 5, "B3"),
        (4, 1, None, "headers"),
        (4, 2, "Scenario", "unique"),
        (4, 9, "desc", "column J header"),
        (8, 9, "Interest", "same number of rows"),
        (10, 9, "Property", "contiguous"),
        (5, 9, "Typo", "scen desc must be one of"),
        (8, 0, 9, "number 1 to 4"),
        (7, 16, "x", "must be numeric"),
    ],
)
def test_validate_scenflipper_rejects_invalid_layouts(
    row: int, column: int, value: object, message: str
) -> None:
    values = _scenario_grid()
    values[row][column] = value

    with pytest.raises(BondsLiteError, match=message):
        _stage_one().validate_scenflipper(values, False, True)


def test_validate_scenflipper_rejects_non_contiguous_scenario_rows() -> None:
    values = _scenario_grid()
    values[6][0] = None

    with pytest.raises(BondsLiteError, match="contiguous"):
        _stage_one().validate_scenflipper(values, False, True)


def test_validate_scenflipper_requires_base_as_first_block() -> None:
    values = _scenario_grid()[:11]
    values[5][9] = values[6][9] = "Interest_Property"

    with pytest.raises(BondsLiteError, match="first scenario block must be Base"):
        _stage_one().validate_scenflipper(values, False, True)


@pytest.mark.parametrize("error_value", ["#VALUE!", "#N/A", "#NAME?"])
def test_validate_scenflipper_rejects_error_values_in_selected_rows(
    error_value: str,
) -> None:
    values = _scenario_grid()
    values[5][1] = error_value

    with pytest.raises(BondsLiteError, match="Excel error"):
        _stage_one().validate_scenflipper(values, False, True)


def test_validate_scenflipper_ignores_errors_in_unselected_blocks() -> None:
    values = _scenario_grid()
    values[9][1] = "#N/A"

    assert _stage_one().validate_scenflipper(values, False, True)[3] == "Property"


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
def test_replace_scenflipper_clears_union_and_writes_values_only(
    source_rows: int,
    source_columns: int,
    target_rows: int,
    target_columns: int,
) -> None:
    clear_calls: list[object] = []
    writes: list[tuple[object, object]] = []

    class FakeRange:
        def __init__(self, address: object) -> None:
            self.address = address
            self.api = SimpleNamespace(
                ClearContents=lambda: clear_calls.append(address)
            )

        @property
        def value(self) -> object:
            return None

        @value.setter
        def value(self, new_value: object) -> None:
            writes.append((self.address, new_value))

    target_sheet = SimpleNamespace(range=lambda *address: FakeRange(address))
    values = [[row] for row in range(source_rows)]

    _stage_one().replace_scenflipper(
        target_sheet, values, source_columns, target_rows, target_columns
    )

    assert clear_calls == [
        ((1, 1), (max(source_rows, target_rows), max(source_columns, target_columns)))
    ]
    assert writes == [
        (
            ((1, 1), (source_rows, source_columns)),
            [[row] + [None] * (source_columns - 1) for row in range(source_rows)],
        )
    ]


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
            book.names[INTEREST_FLAG].refers_to_range.value = False
            book.names["SST_LTMFlash_input_template"].refers_to_range.value = (
                template.name
            )
            book.names["SST_output_folder_suffix"].refers_to_range.value = (
                "PropertyTest"
            )
            book.names["SST_Entity"].refers_to_range.value = "JRL"
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
                / "JRL_LTM_FlashModel_v3.2a_PropertyTest.xlsb"
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
