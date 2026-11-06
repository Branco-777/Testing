"""Pure contracts for Stage 3 template validation and funded-leg transfer."""

from __future__ import annotations

import hashlib
import os
import sys
from collections.abc import Sequence
from contextlib import nullcontext
from datetime import datetime
from pathlib import Path
from types import SimpleNamespace
from typing import Any

import pytest

MONTH_COUNT = 1200


def _stage_three() -> Any:
    """Import lazily so contract tests collect before the module is implemented."""
    import sst_stage_three

    return sst_stage_three


def _mace_row(
    isin: str = "XS0000000001",
    cashflows: Sequence[object] | None = None,
    **overrides: object,
) -> dict[str, object]:
    """Build a minimal synthetic row using semantic MACE field names."""
    row: dict[str, object] = {
        "isin": isin,
        "zspread": 0.0125,
        "type": "Bond",
        "fsao": 0.25,
        "ai": -125.0,
        "funding_cashflows": list(cashflows or [0.0] * MONTH_COUNT),
    }
    row.update(overrides)
    return row


def _map_row(
    isin: str,
    cashflows: Sequence[object] | None = None,
) -> dict[str, object]:
    """Build one synthetic MAP record; MAP cashflows are reconciliation-only."""
    return {
        "isin": isin,
        "currency": "GBP",
        "sector": "Infrastructure",
        "rating": "A",
        "index_linked": True,
        "cashflows": list(cashflows or [1.0] * MONTH_COUNT),
    }


def _map_bonds(*rows: dict[str, object]) -> dict[str, object]:
    """Represent the MAP description lookup."""
    return {"by_isin": {row["isin"]: row for row in rows}}


def _bonds_negative(*rows: dict[str, object]) -> set[str]:
    """Return the negative-cashflow ISINs of synthetic Stage 1 Bonds input rows."""
    return {
        str(row["isin"])
        for row in rows
        if min(row["cashflows"]) < 0  # type: ignore[arg-type]
    }


def _allocations(
    values: dict[str, tuple[object, object, object]],
    duplicates: frozenset[str] = frozenset(),
    valuation_date: object = None,
) -> Any:
    """Build a synthetic MACE Control allocation lookup (Comp. A %, B %, flag)."""
    from datetime import datetime

    stage_three = _stage_three()
    return stage_three.MaceAllocations(
        by_isin={
            isin: {"a": a, "b": b, "c": c} for isin, (a, b, c) in values.items()
        },
        duplicates=duplicates,
        valuation_date=valuation_date or datetime(2026, 6, 30),
    )


def _funded_isins(funded: Sequence[object]) -> set[str]:
    """Read ISINs from either mapping-like or attribute-based result records."""
    return {
        item["isin"] if isinstance(item, dict) else item.isin  # type: ignore[index, union-attr]
        for item in funded
    }


def test_excel_application_starts_with_a_workbook_for_application_settings(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import xlwings as xw

    stage_three = _stage_three()
    created: dict[str, Any] = {}

    class ApplicationApi:
        def __init__(self, has_workbook: bool) -> None:
            object.__setattr__(self, "has_workbook", has_workbook)
            object.__setattr__(self, "settings", {})
            object.__setattr__(self, "quit_called", False)

        def __setattr__(self, name: str, value: object) -> None:
            if name == "CalculateBeforeSave" and not self.has_workbook:
                raise RuntimeError("CalculateBeforeSave requires a workbook")
            self.settings[name] = value

        def Quit(self) -> None:
            object.__setattr__(self, "quit_called", True)

    class Application:
        def __init__(self, *, visible: bool, add_book: bool) -> None:
            created["visible"] = visible
            created["add_book"] = add_book
            self.api = ApplicationApi(add_book)
            created["application"] = self

    monkeypatch.setattr(xw, "App", Application)

    with stage_three._excel_application() as application:
        assert application.api.settings["CalculateBeforeSave"] is False
        assert application.api.settings["Calculation"] == -4135

    assert created["visible"] is False
    assert created["add_book"] is True
    assert created["application"].api.quit_called is True


@pytest.mark.excel
@pytest.mark.skipif(
    os.environ.get("BONDS_LITE_RUN_STAGE_THREE_EXCEL_TESTS") != "1",
    reason="Set BONDS_LITE_RUN_STAGE_THREE_EXCEL_TESTS=1",
)
def test_excel_macalc_write_preserves_template_links_and_protected_rows(
    tmp_path: Path,
) -> None:
    stage_three = _stage_three()
    template = Path(__file__).resolve().parents[1] / "MACalcinPython_v4.0f.xlsb"
    template_hash = hashlib.sha256(template.read_bytes()).digest()
    output = tmp_path / "MACalc output.xlsb"
    layout = stage_three._read_template_snapshot(template).layout

    cleared_through = stage_three.write_macalc_output(
        template, output, layout, []
    )
    stage_three.verify_macalc_output(
        output,
        [],
        template_path=template,
        layout=layout,
        cleared_through=cleared_through,
    )

    assert hashlib.sha256(template.read_bytes()).digest() == template_hash


SLICE_B_GOLDEN = {
    "funded": [
        "GR0000000013_JRINFRA_Aviva",
        "IL0000000122_JRINFRA_Aviva",
        "IL0000000167_JRINFRA_BlackRock",
        "IL0000000167A_JRINFRA_BlackRock",
        "JRCM60_JRINFRA_Rothschild",
    ],
    "asset_count": 988,
    "first_asset": "DEPOSIT00013_JRRLAM_Just_A",
    "last_asset": "XS3363388086_JRABACUS_Just_AA",
    "sum_m": 20810222938.965466,
    "sum_f": 204171.0806028967,
    "sum_a": 969.3,
    "sum_b": 7.1,
}


SLICE_C_GOLDEN = {
    # size, kept, remapped, unmatched, added, matched keys (plan C6 table).
    "tbl_ForcedCompAperc": (24, 19, 5, 0, 0, 24),
    "tbl_MinCompAperc": (164, 0, 106, 58, 0, 106),
    "lst_pre16assets": (322, 0, 31, 291, 0, 31),
    "lst_REITs": (16, 0, 16, 0, 0, 16),
}


@pytest.mark.excel
@pytest.mark.skipif(
    os.environ.get("BONDS_LITE_RUN_STAGE_THREE_EXCEL_TESTS") != "1",
    reason="Set BONDS_LITE_RUN_STAGE_THREE_EXCEL_TESTS=1",
)
def test_excel_slice_b_real_data_regression_on_handoff_111001(
    tmp_path: Path,
) -> None:
    """Real 111001 CSV/MAP plus central MACE: IDs, counts and checksums only."""
    import math

    import bonds_lite_input as shared
    import xlwings as xw

    stage_three = _stage_three()
    root = Path(__file__).resolve().parents[1]
    run = root / "Bonds Lite output" / "20261102_111001"
    template = root / "MACalcinPython_v4.0f.xlsb"
    mace_path = Path(
        r"\\server090\internal model$\7 ALM\2026\7. July\MA\JRL"
        r"\202606 - JRL - MACE v5.20 c2.xlsb"
    )
    if not (run.is_dir() and template.is_file() and mace_path.is_file()):
        pytest.skip("Real 111001 handoff, template or central MACE not available")
    template_hash = hashlib.sha256(template.read_bytes()).digest()
    snapshot = stage_three._read_template_snapshot(template)
    layout = snapshot.layout
    map_bonds = stage_three.read_map_bonds(run / shared.INPUT_FILENAMES["MAP"])
    bonds_negative = stage_three.read_bonds_negative_isins(
        run / shared.INPUT_FILENAMES["Bonds"]
    )
    base = stage_three.read_map_base_output(stage_three.resolve_map_base_output(run))
    mace = stage_three.read_mace_inputs(mace_path)
    funded = stage_three.identify_funded(mace.funding)
    stage_three.reconcile_funded(map_bonds, funded, bonds_negative)
    stage_three.reconcile_slice_b(
        base, mace.bonds, mace.allocations, map_bonds, funded, datetime(2026, 6, 30)
    )
    funded_rows = stage_three.build_funded_rows(map_bonds, funded, mace.allocations)
    asset_rows = stage_three.build_asset_rows(base, mace.bonds, mace.allocations)
    asset_ratings = [row.final_rating for row in base.rows]
    asset_keys = stage_three.build_asset_keys(
        layout, snapshot.fixed_keys, funded_rows, asset_rows,
        asset_ratings=asset_ratings,
    )
    alignment = stage_three.align_isin_lists(snapshot.isin_lists, asset_keys)
    assert {
        item.spec.name: (
            len(item.keys),
            item.kept,
            item.remapped,
            len(item.unmatched),
            item.added,
            item.matched_keys,
        )
        for item in alignment.lists
    } == SLICE_C_GOLDEN
    assert all(item.duplicates_removed == 0 for item in alignment.lists)
    # Informative parity: each matched (row identity, value) pair is unchanged.
    identity = {
        key.d: key.e if key.block == "asset" else key.d for key in asset_keys
    }
    for source, aligned in zip(snapshot.isin_lists.lists, alignment.lists):
        values = aligned.values or (None,) * len(aligned.keys)
        source_values = source.values or (None,) * len(source.keys)
        assert {
            (identity[key], value)
            for key, value in zip(aligned.keys, values)
            if key in identity
        } == {
            (key, value)
            for key, value in zip(source.keys, source_values)
            if key in set(identity.values())
        }
    output = tmp_path / "MACalc output.xlsb"

    cleared_through = stage_three.write_macalc_output(
        template, output, layout, funded_rows, asset_rows,
        isin_lists=alignment.lists,
    )
    stage_three.verify_macalc_output(
        output,
        [*funded_rows, *asset_rows],
        template_path=template,
        layout=layout,
        cleared_through=cleared_through,
        isin_lists=alignment.lists,
        template_isin_lists=snapshot.isin_lists,
        funded_count=len(funded_rows),
        asset_ratings=asset_ratings,
    )

    application = xw.App(visible=False, add_book=False)
    try:
        application.display_alerts = False
        book = application.books.open(str(output), update_links=False, read_only=True)
        try:
            start = layout.funded_start_row
            values = book.sheets[stage_three.MACALC_SHEET].range(
                f"A{start}:M{start + len(funded_rows) + len(asset_rows) - 1}"
            ).options(ndim=2).value
        finally:
            book.close()
    finally:
        application.api.Quit()
    funded_out = values[: len(funded_rows)]
    assets_out = values[len(funded_rows) :]
    golden = SLICE_B_GOLDEN
    assert [row[3] for row in funded_out] == golden["funded"]
    assert len(assets_out) == golden["asset_count"]
    assert assets_out[0][3] == golden["first_asset"]
    assert assets_out[-1][3] == golden["last_asset"]
    assert math.fsum(row[12] for row in assets_out) == golden["sum_m"]
    assert math.fsum(row[5] for row in assets_out) == golden["sum_f"]
    assert math.fsum(row[0] for row in assets_out) == golden["sum_a"]
    assert math.fsum(row[1] for row in assets_out) == golden["sum_b"]
    assert all(row[0] is not None for row in values)
    assert hashlib.sha256(template.read_bytes()).digest() == template_hash


@pytest.mark.parametrize("launcher_name", ["VBA_LAUNCHER", "PACKAGED_VBA_LAUNCHER"])
def test_launchers_define_stage_three_macro(launcher_name: str) -> None:
    import bonds_lite_input as workflow

    launcher = getattr(workflow, launcher_name)

    assert "Public Sub PrepareMACalc()\n    LaunchBondsLite 3\nEnd Sub" in launcher
    assert launcher.count("Public Sub PrepareMACalc()") == 1


@pytest.mark.parametrize(
    ("button_left", "expected_row"),
    [(100.0, 100), (700.0, 113)],
)
def test_stage_three_button_row_sits_beside_panel_or_below_it(
    button_left: float, expected_row: int
) -> None:
    stage_three = _stage_three()
    panel = SimpleNamespace(row=100, left=705.5, rows=SimpleNamespace(count=11))

    assert stage_three._stage_three_button_row(panel, button_left) == expected_row


@pytest.mark.excel
@pytest.mark.skipif(
    os.environ.get("BONDS_LITE_RUN_STAGE_THREE_EXCEL_TESTS") != "1",
    reason="Set BONDS_LITE_RUN_STAGE_THREE_EXCEL_TESTS=1",
)
def test_excel_install_stage_three_controls_on_draft_copy(tmp_path: Path) -> None:
    import shutil

    import xlwings as xw

    stage_three = _stage_three()
    draft = Path(__file__).resolve().parents[1] / "Draft Input spreadsheet for SST.xlsm"
    copy = tmp_path / draft.name
    shutil.copy2(draft, copy)
    draft_hash = hashlib.sha256(draft.read_bytes()).digest()

    def snapshot() -> dict[str, Any]:
        application = xw.App(visible=False, add_book=True)
        try:
            application.api.AutomationSecurity = 3
            application.display_alerts = False
            book = application.books.open(str(copy), update_links=False, read_only=True)
            try:
                names = {
                    name.name.split("!")[-1]: name.refers_to
                    for name in book.names
                    if name.name.split("!")[-1].startswith("SST_Stage")
                }
                sheet = book.sheets["I. Control"]
                shapes = {
                    shape.Name: (shape.Left, shape.Top, shape.OnAction)
                    for shape in sheet.api.Shapes
                    if shape.Name.startswith("btn")
                }
                code = book.api.VBProject.VBComponents.Item(
                    "modBondsLiteLauncher"
                ).CodeModule
                status = (
                    book.names["SST_Stage3_Status"].refers_to_range.value
                    if "SST_Stage3_Status" in names
                    else None
                )
                return {
                    "names": names,
                    "shapes": shapes,
                    "code": code.Lines(1, code.CountOfLines),
                    "status": status,
                }
            finally:
                book.close()
        finally:
            application.api.DisplayAlerts = False
            application.api.Quit()

    before = snapshot()
    first_backup = stage_three.install_stage_three_controls(copy)
    first = snapshot()
    second_backup = stage_three.install_stage_three_controls(copy)
    second = snapshot()

    assert first_backup.is_file() and second_backup.is_file()
    assert {
        name: ref for name, ref in first["names"].items()
        if not name.startswith("SST_Stage3")
    } == {
        name: ref for name, ref in before["names"].items()
        if not name.startswith("SST_Stage3")
    }
    assert {
        name: shape for name, shape in first["shapes"].items()
        if name != "btnPrepareMACalc"
    } == {
        name: shape for name, shape in before["shapes"].items()
        if name != "btnPrepareMACalc"
    }
    assert set(stage_three.STAGE_THREE_FIELD_NAMES.values()) <= first["names"].keys()
    assert first["names"]["SST_Stage3_Panel"] == "='I. Control'!$H$100:$O$110"
    assert first["status"] == (before["status"] or "Not Run")
    left, top, action = first["shapes"]["btnPrepareMACalc"]
    assert action.endswith("!PrepareMACalc")
    assert first["code"].count("Public Sub PrepareMACalc()") == 1
    assert second == first
    assert hashlib.sha256(draft.read_bytes()).digest() == draft_hash


@pytest.mark.parametrize(
    ("types", "funded_start_row"),
    [
        ([" mortgage ", "MORTGAGE", " ils ", "ILS", "Bond"], 6),
        (["Mortgage", "ILS", "Bond"], 4),
    ],
)
def test_validate_template_blocks_accepts_leading_mortgage_then_ils_blocks(
    types: list[str],
    funded_start_row: int,
) -> None:
    stage_three = _stage_three()

    layout = stage_three.validate_template_blocks(types)

    assert layout.funded_start_row == funded_start_row


@pytest.mark.parametrize(
    "types",
    [
        ["Bond", "ILS"],  # no Mortgage block
        ["Mortgage", "Bond"],  # no ILS block
        ["ILS", "Mortgage"],  # reversed blocks
        ["Mortgage", None, "ILS"],  # blank between blocks
        ["Mortgage", "ILS", "Mortgage"],  # Mortgage after ILS
        ["Mortgage", "Mortgage", " ", "ILS"],  # blank within Mortgage block
    ],
)
def test_validate_template_blocks_rejects_missing_blank_or_out_of_order_blocks(
    types: list[object],
) -> None:
    stage_three = _stage_three()

    with pytest.raises(stage_three.TemplateLayoutError):
        stage_three.validate_template_blocks(types)


@pytest.mark.parametrize(
    ("cashflows", "expected_funded"),
    [
        ([0.0] * MONTH_COUNT, False),
        ([-1.0] * MONTH_COUNT, True),
        ([-1.0, -0.5] + [0.0] * (MONTH_COUNT - 2), True),
        ([-1.0, 0.0, -0.25] + [0.0] * (MONTH_COUNT - 3), True),
    ],
)
def test_identify_funded_classifies_zero_and_negative_only_vectors(
    cashflows: list[float],
    expected_funded: bool,
) -> None:
    stage_three = _stage_three()

    funded = stage_three.identify_funded([_mace_row(cashflows=cashflows)])

    assert _funded_isins(funded) == (
        {"XS0000000001"} if expected_funded else set()
    )


@pytest.mark.parametrize(
    "cashflows",
    [
        ([0.0] * (MONTH_COUNT - 1) + [1.0]),  # positive-only
        ([-1.0] + [0.0] * (MONTH_COUNT - 2) + [1.0]),  # mixed signs
    ],
)
def test_identify_funded_rejects_any_positive_month(
    cashflows: list[float],
) -> None:
    stage_three = _stage_three()

    with pytest.raises(Exception, match="(?i)positive|anomal"):
        stage_three.identify_funded([_mace_row(cashflows=cashflows)])


@pytest.mark.parametrize("negative_value", [-0.0, -1e-12])
def test_identify_funded_respects_negative_boundary(
    negative_value: float,
) -> None:
    stage_three = _stage_three()
    cashflows = [negative_value] + [0.0] * (MONTH_COUNT - 1)

    funded = stage_three.identify_funded([_mace_row(cashflows=cashflows)])

    assert bool(_funded_isins(funded)) is (negative_value < 0)


@pytest.mark.parametrize(
    "bad_value",
    [None, "#VALUE!", "not-a-number"],
)
def test_identify_funded_rejects_blank_error_or_text_cashflows(
    bad_value: object,
) -> None:
    stage_three = _stage_three()
    cashflows: list[object] = [-1.0] + [0.0] * (MONTH_COUNT - 1)
    cashflows[1] = bad_value

    with pytest.raises(Exception):
        stage_three.identify_funded([_mace_row(cashflows=cashflows)])


@pytest.mark.parametrize(
    "overrides",
    [
        {"isin": None},
        {"isin": " "},
        {"zspread": None},
        {"fsao": "#VALUE!"},
        {"ai": "not-a-number"},
        {"type": ""},
        {"type": 123},
    ],
)
def test_identify_funded_rejects_invalid_funded_identity_or_required_values(
    overrides: dict[str, object],
) -> None:
    stage_three = _stage_three()
    row = _mace_row(cashflows=[-1.0] + [0.0] * (MONTH_COUNT - 1))
    row.update(overrides)

    with pytest.raises(Exception):
        stage_three.identify_funded([row])


@pytest.mark.parametrize(
    "rows",
    [
        [_mace_row(isin=""), _mace_row(isin="XS0000000002")],
        [_mace_row(), _mace_row()],
    ],
)
def test_identify_funded_rejects_blank_or_duplicate_mace_isins(
    rows: list[dict[str, object]],
) -> None:
    stage_three = _stage_three()

    with pytest.raises(Exception):
        stage_three.identify_funded(rows)


def test_reconcile_funded_compares_mace_with_bonds_input_not_map() -> None:
    """Stage 1 strips negatives from MAP, so MAP has none yet must still pass."""
    stage_three = _stage_three()
    negative = [-1.0] + [1.0] * (MONTH_COUNT - 1)
    bonds_rows = [
        _map_row("XS0000000001", negative),
        _map_row("XS0000000002"),
    ]
    map_bonds = _map_bonds(_map_row("XS0000000001"), _map_row("XS0000000002"))
    funded = [
        _mace_row("XS0000000001", [-1.0] + [0.0] * (MONTH_COUNT - 1))
    ]

    assert (
        stage_three.reconcile_funded(
            map_bonds, funded, _bonds_negative(*bonds_rows)
        )
        is False
    )


@pytest.mark.parametrize(
    ("bonds_rows", "map_rows", "funded", "unexpected_isin", "label"),
    [
        (
            [_map_row("XS0000000001", [-1.0] + [1.0] * (MONTH_COUNT - 1)),
             _map_row("BONDS_ONLY", [-0.5] + [1.0] * (MONTH_COUNT - 1))],
            [_map_row("XS0000000001"), _map_row("BONDS_ONLY")],
            [_mace_row("XS0000000001", [-1.0] + [0.0] * (MONTH_COUNT - 1))],
            "BONDS_ONLY",
            "Bonds input only",
        ),
        (
            [_map_row("XS0000000001", [-1.0] + [1.0] * (MONTH_COUNT - 1)),
             _map_row("MACE_ONLY")],
            [_map_row("XS0000000001"), _map_row("MACE_ONLY")],
            [
                _mace_row("XS0000000001", [-1.0] + [0.0] * (MONTH_COUNT - 1)),
                _mace_row("MACE_ONLY", [-0.5] + [0.0] * (MONTH_COUNT - 1)),
            ],
            "MACE_ONLY",
            "MACE only",
        ),
        (
            [_map_row("MISSING_FROM_MAP", [-1.0] + [1.0] * (MONTH_COUNT - 1))],
            [_map_row("XS0000000001")],
            [_mace_row("MISSING_FROM_MAP", [-1.0] + [0.0] * (MONTH_COUNT - 1))],
            "MISSING_FROM_MAP",
            "missing from MAP",
        ),
    ],
)
def test_reconcile_funded_rejects_and_identifies_set_mismatches(
    bonds_rows: list[dict[str, object]],
    map_rows: list[dict[str, object]],
    funded: list[dict[str, object]],
    unexpected_isin: str,
    label: str,
) -> None:
    stage_three = _stage_three()

    with pytest.raises(stage_three.StageThreeError) as error:
        stage_three.reconcile_funded(
            _map_bonds(*map_rows), funded, _bonds_negative(*bonds_rows)
        )

    assert "Bonds input/MACE funded set mismatch" in str(error.value)
    assert unexpected_isin in str(error.value)
    assert label in str(error.value)


def test_reconcile_funded_ignores_negative_cashflows_left_in_map() -> None:
    """Only the Bonds input defines the funded set; MAP cashflows are not used."""
    stage_three = _stage_three()
    map_bonds = _map_bonds(
        _map_row("XS0000000001", [-1.0] + [1.0] * (MONTH_COUNT - 1))
    )

    assert stage_three.reconcile_funded(map_bonds, [], set()) is True


def test_reconcile_funded_allows_both_sets_empty_for_zero_funded_warning() -> None:
    stage_three = _stage_three()

    stage_three.reconcile_funded(_map_bonds(), [], set())


def test_build_funded_rows_maps_values_and_only_mace_funding_cashflows() -> None:
    stage_three = _stage_three()
    isin = "XS0000000001"
    map_bonds = _map_bonds(
        _map_row(isin, [-999.0] + [1.0] * (MONTH_COUNT - 1))
    )
    funding_cashflows = [-2.5, -1.25] + [0.0] * (MONTH_COUNT - 2)
    mace = _mace_row(
        isin,
        funding_cashflows,
        zspread=0.034,
        type="Infrastructure",
        fsao=17.5,
        ai=125.0,
    )

    allocations = _allocations({isin: (0.6, 0.4, 1.1e-16)})

    rows = stage_three.build_funded_rows(map_bonds, [mace], allocations)

    assert len(rows) == 1
    assert len(rows[0]) == stage_three._column_number("ATQ")
    assert rows[0][:3] == [0.6, 0.4, 1.1e-16]  # A:C from MACE Control, exact
    assert rows[0][3:13] == [
        isin,  # D: ISIN
        isin,  # E: Name is deliberately the ISIN
        0.034,  # F: Zspread from MACE
        "Infrastructure",  # G: Type from MACE
        "Infrastructure",  # H: Sector from MAP
        "A",  # I: Rating from MAP
        True,  # J: cached Index Linked? value from MAP
        "GBP",  # K: Currency from MAP
        17.5,  # L: FSAO from MACE
        -125.0,  # M: PV is -AI
    ]
    assert len(rows[0][13:]) == MONTH_COUNT
    assert rows[0][13:] == funding_cashflows  # N:ATQ comes from MACE AJ:AUM
    assert rows[0][13:] != map_bonds["by_isin"][isin]["cashflows"]


def test_build_funded_rows_returns_no_rows_when_funded_set_is_empty() -> None:
    stage_three = _stage_three()

    assert stage_three.build_funded_rows(_map_bonds(), [], _allocations({})) == []


def test_build_funded_rows_stops_when_funded_isin_has_no_control_row() -> None:
    stage_three = _stage_three()
    isin = "XS0000000001"
    mace = _mace_row(isin, [-1.0] + [0.0] * (MONTH_COUNT - 1))

    with pytest.raises(stage_three.StageThreeError, match=isin):
        stage_three.build_funded_rows(
            _map_bonds(_map_row(isin)), [mace], _allocations({})
        )


def _template_headers(stage_three: Any) -> list[object]:
    headers: list[object] = [None] * stage_three._column_number("ATQ")
    for column, value in {
        "A": "Aallocations",
        "B": "Ballocations",
        "C": "PartialRecog",
        "D": "ISIN",
        "E": "Name",
        "F": "Zspread",
        "G": "Type",
        "H": "Sector",
        "I": "Rating",
        "J": "Index Linked?",
        "K": "Currency",
        "L": "FSAO",
        "M": "PV",
    }.items():
        headers[stage_three._column_number(column) - 1] = value
    first_month = stage_three._column_number("N") - 1
    headers[first_month : first_month + MONTH_COUNT] = list(range(1, MONTH_COUNT + 1))
    return headers


def _map_headers(stage_three: Any) -> list[object]:
    headers: list[object] = [None] * stage_three._column_number("ATT")
    for column, value in {
        "C": "ISIN",
        "E": "Currency",
        "F": "Sector",
        "H": "NotchedRating",
        "J": "Index Linked?",
    }.items():
        headers[stage_three._column_number(column) - 1] = value
    first_month = stage_three._column_number("Q") - 1
    headers[first_month : first_month + MONTH_COUNT] = [
        str(month) for month in range(1, MONTH_COUNT + 1)
    ]
    return headers


def _mace_headers(stage_three: Any) -> list[object]:
    headers: list[object] = [None] * stage_three._column_number("AUM")
    for column, value in {
        "B": "ISIN",
        "C": "Z-Spread",
        "D": "Type",
        "I": "FS Add On",
        "AI": 0,
    }.items():
        headers[stage_three._column_number(column) - 1] = value
    first_month = stage_three._column_number("AJ") - 1
    headers[first_month : first_month + MONTH_COUNT] = list(range(1, MONTH_COUNT + 1))
    return headers


def test_validate_template_headers_enforces_fixed_columns_and_month_ordinals() -> None:
    stage_three = _stage_three()
    headers = _template_headers(stage_three)

    stage_three.validate_template_headers(headers)

    headers[stage_three._column_number("ATQ") - 1] = 1199
    with pytest.raises(stage_three.StageThreeError, match="ordinals"):
        stage_three.validate_template_headers(headers)


@pytest.mark.parametrize(
    ("column", "value"),
    [("D", "Asset ID"), ("J", "Index Linked"), ("M", "Present Value")],
)
def test_validate_template_headers_rejects_header_drift(
    column: str,
    value: str,
) -> None:
    stage_three = _stage_three()
    headers = _template_headers(stage_three)
    headers[stage_three._column_number(column) - 1] = value

    with pytest.raises(stage_three.StageThreeError, match="header mismatch"):
        stage_three.validate_template_headers(headers)


def test_map_and_mace_header_contracts_match_source_month_sequences() -> None:
    stage_three = _stage_three()
    map_headers = _map_headers(stage_three)
    mace_headers = _mace_headers(stage_three)

    stage_three._require_map_headers(map_headers)
    stage_three._require_mace_headers(mace_headers)

    map_headers[stage_three._column_number("ATT") - 1] = 1199
    with pytest.raises(stage_three.StageThreeError, match="Q:ATT"):
        stage_three._require_map_headers(map_headers)
    mace_headers[stage_three._column_number("AUM") - 1] = 1199
    with pytest.raises(stage_three.StageThreeError, match="AJ:AUM"):
        stage_three._require_mace_headers(mace_headers)


def test_reconcile_funded_returns_true_only_for_two_empty_sets() -> None:
    stage_three = _stage_three()

    assert stage_three.reconcile_funded(_map_bonds(), [], set()) is True
    assert stage_three.reconcile_funded(
        _map_bonds(_map_row("XS0000000001")),
        [_mace_row("XS0000000001", [-1.0] + [0.0] * (MONTH_COUNT - 1))],
        {"XS0000000001"},
    ) is False


class _NamedRange:
    def __init__(self, value: object) -> None:
        self.value = value


class _GuardCell:
    def __init__(self, value: object) -> None:
        self.value = value

    def options(self, **kwargs: object) -> _GuardCell:
        return self


class _GuardSheet:
    def __init__(self, warning: object) -> None:
        self.warning = warning

    def range(self, address: str) -> _GuardCell:
        assert address == "H44"
        return _GuardCell(self.warning)


class _GuardCaller:
    def __init__(
        self,
        draft: Path,
        template: Path,
        mace: Path,
        warning: object = None,
    ) -> None:
        self.names = {
            "SST_MACalc_Input_Template": SimpleNamespace(
                refers_to_range=_NamedRange(template.name)
            ),
            "SST_MACE_Central_Template": SimpleNamespace(
                refers_to_range=_NamedRange(str(mace))
            ),
            "I_Valuation_Date": SimpleNamespace(
                refers_to_range=_NamedRange(datetime(2026, 6, 30))
            ),
        }
        self.sheets = {"I. Control": _GuardSheet(warning)}
        self.draft = draft


def _stage_three_resolver(
    monkeypatch: pytest.MonkeyPatch,
    stage_three: Any,
    draft: Path,
    run_folder: Path,
    status: str = "Succeeded",
    stage_two_run: str = "SST Output_a/Credit",
) -> _GuardCaller:
    import bonds_lite_input as shared

    template = draft.parent / "MACalc template.xlsb"
    mace = draft.parent / "MACE.xlsb"
    template.write_bytes(b"template")
    mace.write_bytes(b"MACE")
    run_folder.mkdir()
    map_path = run_folder / shared.INPUT_FILENAMES["MAP"]
    map_path.write_bytes(b"MAP")
    (run_folder / shared.INPUT_FILENAMES["Bonds"]).write_bytes(b"Bonds")
    map_folder = run_folder / "JRL MAP"
    map_folder.mkdir()
    (map_folder / "py_bonds_Base_Scenario_nozeros_BWS.csv").write_text("x")
    monkeypatch.setattr(
        shared,
        "read_model_inputs",
        lambda folder: {"MAP": SimpleNamespace(output_folder=map_folder)},
    )
    monkeypatch.setattr(shared, "require_stage_two_panel", lambda caller: None)
    monkeypatch.setattr(shared, "resolve_caller_path", lambda caller: draft)
    monkeypatch.setattr(shared, "resolve_stage_input", lambda caller: run_folder)
    monkeypatch.setattr(
        shared,
        "portable_output_path",
        lambda caller_path, output: "SST Output_a/Credit",
    )
    monkeypatch.setattr(
        shared,
        "read_stage_two_audit",
        lambda caller: {"status": status, "input_run": stage_two_run},
    )
    return _GuardCaller(draft, template, mace)


def test_resolve_stage_three_inputs_accepts_current_stage_two_and_relative_macalc_path(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    stage_three = _stage_three()
    draft = tmp_path / "Draft.xlsm"
    draft.write_bytes(b"Draft")
    run_folder = tmp_path / "SST Output_a" / "Credit"
    run_folder.parent.mkdir()
    caller = _stage_three_resolver(monkeypatch, stage_three, draft, run_folder)

    inputs = stage_three.resolve_stage_three_inputs(caller)

    assert inputs.run_folder == run_folder
    assert inputs.map_path == run_folder / "Bonds Lite input_MAP.xlsx"
    assert inputs.bonds_path == run_folder / "Bonds Lite input_Bonds.xlsx"
    assert inputs.template_path == tmp_path / "MACalc template.xlsb"
    assert inputs.mace_path == tmp_path / "MACE.xlsb"
    assert inputs.map_base_output_path == (
        run_folder / "JRL MAP" / "py_bonds_Base_Scenario_nozeros_BWS.csv"
    )
    assert inputs.valuation_date == datetime(2026, 6, 30)


def test_resolve_stage_three_inputs_stops_when_base_scenario_csv_is_missing(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    stage_three = _stage_three()
    draft = tmp_path / "Draft.xlsm"
    draft.write_bytes(b"Draft")
    run_folder = tmp_path / "SST Output_a" / "Credit"
    run_folder.parent.mkdir()
    caller = _stage_three_resolver(monkeypatch, stage_three, draft, run_folder)
    (run_folder / "JRL MAP" / "py_bonds_Base_Scenario_nozeros_BWS.csv").unlink()

    with pytest.raises(stage_three.StageThreeError, match="nozeros"):
        stage_three.resolve_stage_three_inputs(caller)


def test_resolve_stage_three_inputs_stops_when_bonds_input_is_missing(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    stage_three = _stage_three()
    draft = tmp_path / "Draft.xlsm"
    draft.write_bytes(b"Draft")
    run_folder = tmp_path / "SST Output_a" / "Credit"
    run_folder.parent.mkdir()
    caller = _stage_three_resolver(monkeypatch, stage_three, draft, run_folder)
    (run_folder / "Bonds Lite input_Bonds.xlsx").unlink()

    with pytest.raises(stage_three.StageThreeError, match="Bonds input is missing"):
        stage_three.resolve_stage_three_inputs(caller)


def test_resolve_stage_three_inputs_refuses_credit_replaced_by_stage_one(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    """Decision 59: a Stage 1 replace removes Stage 2 results; Stage 3 must stop."""
    import bonds_lite_input as shared

    stage_three = _stage_three()
    draft = tmp_path / "Draft.xlsm"
    draft.write_bytes(b"Draft")
    run_folder = tmp_path / "SST Output_a" / "Credit"
    run_folder.parent.mkdir()
    # The earlier Stage 2 succeeded for this same relative run path.
    caller = _stage_three_resolver(monkeypatch, stage_three, draft, run_folder)
    assert stage_three.resolve_stage_three_inputs(caller).run_folder == run_folder

    folder, _, backup = shared.allocate_stress_folder(tmp_path, "a", "Credit")
    assert folder == run_folder
    (folder / "JRL MAP").mkdir()
    (folder / shared.INPUT_FILENAMES["MAP"]).write_bytes(b"New MAP")
    (folder / shared.INPUT_FILENAMES["Bonds"]).write_bytes(b"New Bonds")
    shared.discard_replaced_stress_folder(backup)

    with pytest.raises(stage_three.StageThreeError) as error:
        stage_three.resolve_stage_three_inputs(caller)
    assert "rerun Stage 2" in str(error.value)


@pytest.mark.parametrize(
    ("status", "stage_two_run"),
    [
        ("Failed", "SST Output_a/Credit"),
        ("Succeeded", "SST Output_b/Credit"),
    ],
)
def test_resolve_stage_three_inputs_rejects_failed_or_stale_stage_two(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    status: str,
    stage_two_run: str,
) -> None:
    stage_three = _stage_three()
    draft = tmp_path / "Draft.xlsm"
    draft.write_bytes(b"Draft")
    run_folder = tmp_path / "SST Output_a" / "Credit"
    run_folder.parent.mkdir()
    caller = _stage_three_resolver(
        monkeypatch, stage_three, draft, run_folder, status, stage_two_run
    )

    with pytest.raises(stage_three.StageThreeError) as error:
        stage_three.resolve_stage_three_inputs(caller)

    assert "SST Output_a/Credit" in str(error.value)
    assert stage_two_run in str(error.value)


def test_resolve_stage_three_inputs_rejects_relative_mace_and_h44_warning(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    stage_three = _stage_three()
    draft = tmp_path / "Draft.xlsm"
    draft.write_bytes(b"Draft")
    run_folder = tmp_path / "SST Output_a" / "Credit"
    run_folder.parent.mkdir()
    caller = _stage_three_resolver(monkeypatch, stage_three, draft, run_folder)
    caller.names["SST_MACE_Central_Template"].refers_to_range.value = "MACE.xlsb"

    with pytest.raises(stage_three.StageThreeError, match="must be absolute"):
        stage_three.resolve_stage_three_inputs(caller)

    caller.names["SST_MACE_Central_Template"].refers_to_range.value = str(
        tmp_path / "MACE.xlsb"
    )
    caller.sheets["I. Control"] = _GuardSheet("WARNING: valuation date mismatch")
    with pytest.raises(stage_three.StageThreeError, match="H44"):
        stage_three.resolve_stage_three_inputs(caller)


def test_resolve_stage_three_inputs_rejects_missing_map_file(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    stage_three = _stage_three()
    draft = tmp_path / "Draft.xlsm"
    draft.write_bytes(b"Draft")
    run_folder = tmp_path / "SST Output_a" / "Credit"
    run_folder.parent.mkdir()
    caller = _stage_three_resolver(monkeypatch, stage_three, draft, run_folder)
    (run_folder / "Bonds Lite input_MAP.xlsx").unlink()

    with pytest.raises(stage_three.StageThreeError, match="MAP input is missing"):
        stage_three.resolve_stage_three_inputs(caller)


def test_output_filename_joins_output_suffix_and_token() -> None:
    stage_three = _stage_three()

    name = stage_three.output_filename(
        Path("Testing") / "SST Output_a" / "Credit", "3f9c2a7be104d65e"
    )

    assert name == "MACalc input_a_3f9c2a7be104d65e.xlsb"


@pytest.mark.parametrize("suffix", ["Q2 2026", "7", "run_1"])
def test_output_filename_keeps_suffix_unchanged(suffix: str) -> None:
    stage_three = _stage_three()

    name = stage_three.output_filename(
        Path(f"SST Output_{suffix}") / "Credit", "3f9c2a7be104d65e"
    )

    assert name == f"MACalc input_{suffix}_3f9c2a7be104d65e.xlsb"


@pytest.mark.parametrize(
    "folder",
    [
        "",
        "Credit",
        "Bonds Lite output/20261103_161318",
        "SST Output_/Credit",
        "Output_a/Credit",
        "SST Outputa/Credit",
    ],
)
def test_output_filename_rejects_unexpected_run_folder(folder: str) -> None:
    stage_three = _stage_three()

    with pytest.raises(stage_three.StageThreeError, match="run folder"):
        stage_three.output_filename(Path(folder), "3f9c2a7be104d65e")


def test_new_output_token_is_16_lowercase_hex_and_varies() -> None:
    import re

    stage_three = _stage_three()

    tokens = [stage_three._token() for _ in range(1000)]

    assert all(re.fullmatch(r"[0-9a-f]{16}", token) for token in tokens)
    assert len(set(tokens)) == 1000


def test_new_output_path_regenerates_token_when_name_exists(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    stage_three = _stage_three()
    run_folder = tmp_path / "SST Output_a" / "Credit"
    run_folder.mkdir(parents=True)
    (run_folder / "MACalc input_a_aaaaaaaaaaaaaaaa.xlsb").write_bytes(
        b"existing"
    )
    tokens = iter(["aaaaaaaaaaaaaaaa", "bbbbbbbbbbbbbbbb"])
    monkeypatch.setattr(stage_three, "_token", lambda: next(tokens))

    path = stage_three.new_output_path(run_folder)

    assert path == run_folder / "MACalc input_a_bbbbbbbbbbbbbbbb.xlsb"


def test_new_output_path_stops_when_no_unused_name_is_found(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    stage_three = _stage_three()
    run_folder = tmp_path / "SST Output_a" / "Credit"
    run_folder.mkdir(parents=True)
    (run_folder / "MACalc input_a_aaaaaaaaaaaaaaaa.xlsb").write_bytes(
        b"existing"
    )
    monkeypatch.setattr(stage_three, "_token", lambda: "aaaaaaaaaaaaaaaa")

    with pytest.raises(stage_three.StageThreeError, match="unused"):
        stage_three.new_output_path(run_folder)


NEW_TOKEN = "0123456789abcdef"
NEW_OUTPUT_NAME = f"MACalc input_a_{NEW_TOKEN}.xlsb"
EARLIER_OUTPUT_NAME = "MACalc input_a_aaaaaaaaaaaaaaaa.xlsb"


def _mock_stage_three_run(
    monkeypatch: pytest.MonkeyPatch,
    stage_three: Any,
    tmp_path: Path,
    existing_output: bool,
) -> tuple[Any, Path, list[dict[str, object]], list[str]]:
    """Return the caller, the expected new output path, audit writes and messages.

    With ``existing_output`` an earlier tokenised output of the same run holds
    ``b"previous"`` at ``EARLIER_OUTPUT_NAME``.
    """
    import contextlib

    import bonds_lite_input as shared

    caller_path = tmp_path / "Draft.xlsm"
    caller_path.write_bytes(b"Draft")
    run_folder = tmp_path / "SST Output_a" / "Credit"
    run_folder.mkdir(parents=True)
    output = run_folder / NEW_OUTPUT_NAME
    if existing_output:
        (run_folder / EARLIER_OUTPUT_NAME).write_bytes(b"previous")
    tokens = iter([NEW_TOKEN, "fedcba9876543210"])
    monkeypatch.setattr(stage_three, "_token", lambda: next(tokens))
    caller = SimpleNamespace()
    writes: list[dict[str, object]] = []
    messages: list[str] = []
    stage_inputs = stage_three.StageThreeInputs(
        caller_path=caller_path,
        run_folder=run_folder,
        input_run="SST Output_a/Credit",
        map_path=run_folder / "Bonds Lite input_MAP.xlsx",
        template_path=tmp_path / "MACalc template.xlsb",
        mace_path=tmp_path / "MACE.xlsb",
        map_base_output_path=run_folder / "JRL MAP" / "base.csv",
        valuation_date=datetime(2026, 6, 30),
        bonds_path=run_folder / "Bonds Lite input_Bonds.xlsx",
    )
    monkeypatch.setattr(shared, "resolve_caller_path", lambda book: caller_path)
    monkeypatch.setattr(
        shared,
        "workbook_run_guard",
        lambda path: contextlib.nullcontext(),
    )
    monkeypatch.setattr(shared, "resolve_stage_input", lambda book: run_folder)
    monkeypatch.setattr(
        shared, "portable_output_path",
        lambda path, folder: "SST Output_a/Credit",
    )
    monkeypatch.setattr(stage_three, "require_stage_three_panel", lambda book: None)
    monkeypatch.setattr(
        stage_three, "_read_stage_three_fields",
        lambda book: {"last_output": "Bonds Lite output/previous"},
    )
    monkeypatch.setattr(
        stage_three,
        "resolve_stage_three_inputs",
        lambda book: stage_inputs,
    )
    monkeypatch.setattr(
        stage_three,
        "_read_template_snapshot",
        lambda path: stage_three.TemplateSnapshot(
            stage_three.TemplateLayout(2, 2, 3, 3, 4)
        ),
    )
    monkeypatch.setattr(stage_three, "read_map_bonds", lambda path: object())
    monkeypatch.setattr(stage_three, "read_bonds_negative_isins", lambda path: set())
    monkeypatch.setattr(
        stage_three, "read_map_base_output", lambda path: SimpleNamespace(rows=())
    )
    monkeypatch.setattr(
        stage_three,
        "read_mace_inputs",
        lambda path: SimpleNamespace(
            funding=object(), bonds=object(), allocations=object()
        ),
    )
    monkeypatch.setattr(stage_three, "identify_funded", lambda table: [])
    monkeypatch.setattr(stage_three, "reconcile_funded", lambda *args: True)
    monkeypatch.setattr(
        stage_three,
        "reconcile_slice_b",
        lambda *args: stage_three.SliceBReconciliation(),
    )
    monkeypatch.setattr(
        stage_three, "build_funded_rows", lambda map_data, funded, allocations: []
    )
    monkeypatch.setattr(
        stage_three, "build_asset_rows", lambda base, bonds, allocations: []
    )

    def write_output(
        template: Path,
        staged: Path,
        layout: object,
        rows: object,
        asset_rows: object = (),
        isin_lists: object = (),
    ) -> int:
        staged.write_bytes(b"new")
        return 20

    monkeypatch.setattr(stage_three, "write_macalc_output", write_output)
    monkeypatch.setattr(
        stage_three,
        "verify_macalc_output",
        lambda *args, **kwargs: None,
    )
    monkeypatch.setattr(
        stage_three, "write_stage_three_audit",
        lambda book, fields: writes.append(fields),
    )
    monkeypatch.setattr(
        stage_three, "_message",
        lambda book, text, warning=False: messages.append(text),
    )
    return caller, output, writes, messages


def test_run_stage_three_zero_funded_succeeds_and_warns_without_overwriting_template(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    stage_three = _stage_three()
    caller, output, writes, messages = _mock_stage_three_run(
        monkeypatch, stage_three, tmp_path, existing_output=False
    )

    result = stage_three.run_stage_three(caller)

    assert result == output
    assert output.read_bytes() == b"new"
    assert writes[-1]["status"] == "Succeeded"
    assert writes[-1]["input_run"] == "SST Output_a/Credit"
    assert writes[-1]["last_output"] == (
        f"SST Output_a/Credit/{NEW_OUTPUT_NAME}"
    )
    assert "0 funded assets" in messages[-1]
    assert [path.name for path in output.parent.iterdir()] == [NEW_OUTPUT_NAME]


def test_run_stage_three_publishes_one_tokenised_output_and_audits_it(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import re

    stage_three = _stage_three()
    caller, output, writes, _ = _mock_stage_three_run(
        monkeypatch, stage_three, tmp_path, existing_output=False
    )
    tokens = iter(["3f9c2a7be104d65e"])
    monkeypatch.setattr(stage_three, "_token", lambda: next(tokens))

    result = stage_three.run_stage_three(caller)

    published = list(output.parent.iterdir())
    assert len(published) == 1 and published == [result]
    assert re.fullmatch(
        r"MACalc input_a_[0-9a-f]{16}\.xlsb", result.name
    )
    assert writes[-1]["status"] == "Succeeded"
    assert writes[-1]["last_output"] == (
        f"SST Output_a/Credit/{result.name}"
    )


def _staging_root(tmp_path: Path) -> Path:
    return tmp_path / "SST Output_a" / "Stage 3 staging"


def test_run_stage_three_stages_in_output_root_outside_run_folder_and_cleans_up(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    stage_three = _stage_three()
    caller, output, _, _ = _mock_stage_three_run(
        monkeypatch, stage_three, tmp_path, existing_output=False
    )
    staged_paths: list[Path] = []
    run_folder_during_write: list[list[str]] = []

    def write_output(template, staged, layout, rows, assets=(), isin_lists=()):
        staged_paths.append(staged)
        staged.write_bytes(b"new")
        run_folder_during_write.append(
            [path.name for path in output.parent.iterdir()]
        )
        return 20

    monkeypatch.setattr(stage_three, "write_macalc_output", write_output)

    stage_three.run_stage_three(caller)

    # Staying in the synced output root keeps Excel's link paths unchanged.
    assert staged_paths == [_staging_root(tmp_path) / NEW_TOKEN / NEW_OUTPUT_NAME]
    assert run_folder_during_write == [[]]
    assert not _staging_root(tmp_path).exists()
    assert output.read_bytes() == b"new"


def test_run_stage_three_keeps_staging_root_holding_other_runs(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    stage_three = _stage_three()
    caller, output, _, _ = _mock_stage_three_run(
        monkeypatch, stage_three, tmp_path, existing_output=False
    )
    other = _staging_root(tmp_path) / "fedcba9876543210" / "other.xlsb"
    other.parent.mkdir(parents=True)
    other.write_bytes(b"other run")

    stage_three.run_stage_three(caller)

    assert other.read_bytes() == b"other run"
    assert [path.name for path in _staging_root(tmp_path).iterdir()] == [
        "fedcba9876543210"
    ]
    assert output.read_bytes() == b"new"


def test_run_stage_three_removes_staging_folders_marked_read_only(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import os
    import stat

    stage_three = _stage_three()
    caller, output, _, _ = _mock_stage_three_run(
        monkeypatch, stage_three, tmp_path, existing_output=False
    )

    def write_output(template, staged, layout, rows, assets=(), isin_lists=()):
        staged.write_bytes(b"new")
        # OneDrive marks synced folders read-only, which blocks a plain rmdir.
        for folder in (staged.parent, staged.parent.parent):
            os.chmod(folder, stat.S_IREAD)
        assert not os.stat(staged.parent).st_mode & stat.S_IWRITE
        return 20

    monkeypatch.setattr(stage_three, "write_macalc_output", write_output)

    stage_three.run_stage_three(caller)

    assert not _staging_root(tmp_path).exists()
    assert output.read_bytes() == b"new"


def test_run_stage_three_warns_when_staging_folder_cannot_be_removed(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import shutil

    stage_three = _stage_three()
    caller, output, writes, _ = _mock_stage_three_run(
        monkeypatch, stage_three, tmp_path, existing_output=False
    )
    warnings: list[str] = []

    def failing_rmtree(path, *args, **kwargs):
        raise OSError("locked by sync")

    monkeypatch.setattr(shutil, "rmtree", failing_rmtree)
    monkeypatch.setattr(
        stage_three.logger,
        "warning",
        lambda message, *args: warnings.append(message % args),
    )

    result = stage_three.run_stage_three(caller)

    assert result == output and writes[-1]["status"] == "Succeeded"
    assert any(str(_staging_root(tmp_path) / NEW_TOKEN) in text for text in warnings)


def test_run_stage_three_rejects_staged_path_over_windows_limit_before_any_write(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    stage_three = _stage_three()
    caller, output, writes, messages = _mock_stage_three_run(
        monkeypatch, stage_three, tmp_path, existing_output=False
    )
    staged = _staging_root(tmp_path) / NEW_TOKEN / NEW_OUTPUT_NAME
    assert len(str(staged)) > len(str(output))
    monkeypatch.setattr(stage_three, "MAX_OUTPUT_PATH", len(str(staged)) - 1)

    with pytest.raises(stage_three.StageThreeError, match="characters"):
        stage_three.run_stage_three(caller)

    assert writes == [] and messages == []
    assert list(output.parent.iterdir()) == []
    assert not _staging_root(tmp_path).exists()


def test_run_stage_three_rerun_keeps_earlier_output_without_prompting(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    stage_three = _stage_three()
    caller, output, writes, _ = _mock_stage_three_run(
        monkeypatch, stage_three, tmp_path, existing_output=True
    )
    earlier = output.with_name(EARLIER_OUTPUT_NAME)

    result = stage_three.run_stage_three(caller)

    assert not hasattr(stage_three, "_confirm_overwrite")
    assert result == output
    assert earlier.read_bytes() == b"previous"
    assert output.read_bytes() == b"new"
    assert sorted(path.name for path in output.parent.iterdir()) == sorted(
        [EARLIER_OUTPUT_NAME, NEW_OUTPUT_NAME]
    )
    assert writes[-1]["last_output"] == (
        f"SST Output_a/Credit/{NEW_OUTPUT_NAME}"
    )


def test_run_stage_three_leaves_legacy_output_untouched(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    stage_three = _stage_three()
    caller, output, _, _ = _mock_stage_three_run(
        monkeypatch, stage_three, tmp_path, existing_output=False
    )
    legacy = output.with_name("MACalc input.xlsb")
    legacy.write_bytes(b"legacy")

    stage_three.run_stage_three(caller)

    assert legacy.read_bytes() == b"legacy"
    assert output.read_bytes() == b"new"


def test_run_stage_three_rejects_unexpected_run_folder_before_any_write(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import bonds_lite_input as shared

    stage_three = _stage_three()
    caller, _, writes, messages = _mock_stage_three_run(
        monkeypatch, stage_three, tmp_path, existing_output=False
    )
    copied = tmp_path / "SST Output copy" / "Credit"
    copied.mkdir(parents=True)
    monkeypatch.setattr(shared, "resolve_stage_input", lambda book: copied)

    with pytest.raises(stage_three.StageThreeError, match="run folder"):
        stage_three.run_stage_three(caller)

    assert writes == [] and messages == []
    assert list(copied.iterdir()) == []


def test_run_stage_three_rejects_output_path_over_windows_limit_before_any_write(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    stage_three = _stage_three()
    caller, output, writes, messages = _mock_stage_three_run(
        monkeypatch, stage_three, tmp_path, existing_output=False
    )
    monkeypatch.setattr(stage_three, "MAX_OUTPUT_PATH", len(str(output)) - 1)

    with pytest.raises(stage_three.StageThreeError, match="characters"):
        stage_three.run_stage_three(caller)

    assert writes == [] and messages == []
    assert list(output.parent.iterdir()) == []


def test_max_output_path_is_the_excel_limit() -> None:
    assert _stage_three().MAX_OUTPUT_PATH == 218


def test_run_stage_three_rolls_back_output_when_success_audit_fails(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    stage_three = _stage_three()
    caller, output, writes, _ = _mock_stage_three_run(
        monkeypatch, stage_three, tmp_path, existing_output=True
    )

    def fail_succeeded_audit(book: object, fields: dict[str, object]) -> None:
        writes.append(fields)
        if fields.get("status") == "Succeeded":
            raise OSError("audit save failed")

    monkeypatch.setattr(stage_three, "write_stage_three_audit", fail_succeeded_audit)

    with pytest.raises(OSError, match="audit save failed"):
        stage_three.run_stage_three(caller)

    assert output.with_name(EARLIER_OUTPUT_NAME).read_bytes() == b"previous"
    assert [path.name for path in output.parent.iterdir()] == [EARLIER_OUTPUT_NAME]
    assert not _staging_root(tmp_path).exists()
    assert writes[-1]["status"] == "Failed"


def test_run_stage_three_removes_partial_staged_file_after_write_failure(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    stage_three = _stage_three()
    caller, output, writes, _ = _mock_stage_three_run(
        monkeypatch, stage_three, tmp_path, existing_output=False
    )
    staged_paths: list[Path] = []

    def partial_write(
        template: Path,
        staged: Path,
        layout: object,
        rows: object,
        asset_rows: object = (),
        isin_lists: object = (),
    ) -> int:
        staged_paths.append(staged)
        staged.write_bytes(b"partial")
        raise OSError("write failed")

    monkeypatch.setattr(stage_three, "write_macalc_output", partial_write)

    with pytest.raises(OSError, match="write failed"):
        stage_three.run_stage_three(caller)

    assert not output.exists()
    assert staged_paths and not staged_paths[0].exists()
    assert not _staging_root(tmp_path).exists()
    assert writes[-1]["status"] == "Failed"


class _ReaderRange:
    def __init__(self, value: object = None, row: int = 1) -> None:
        self.value = value
        self.row = row

    def options(self, **kwargs: object) -> _ReaderRange:
        return self

    def end(self, direction: str) -> _ReaderRange:
        assert direction == "up"
        return self


class _MapSheet:
    def __init__(self, headers: list[object], body: list[list[object]]) -> None:
        self.headers = headers
        self.body = body
        self.api = SimpleNamespace(
            ListObjects=lambda name: SimpleNamespace(
                HeaderRowRange=SimpleNamespace(Row=8),
                DataBodyRange=SimpleNamespace(
                    Row=9,
                    Rows=SimpleNamespace(Count=len(body)),
                ),
            )
        )

    def range(self, address: str) -> _ReaderRange:
        if address == "A8:ATT8":
            return _ReaderRange(self.headers)
        if address == f"C9:ATT{8 + len(self.body)}":
            return _ReaderRange(self.body)
        raise AssertionError(f"Unexpected MAP range: {address}")


class _MaceSheet:
    def __init__(self, headers: list[object], body: list[list[object]]) -> None:
        self.headers = headers
        self.body = body

    def range(self, address: str) -> _ReaderRange:
        if address == "A9:AUM9":
            return _ReaderRange(self.headers)
        if address == "B1048576":
            return _ReaderRange(row=10 + len(self.body) - 1)
        if address == f"B10:AUM{9 + len(self.body)}":
            return _ReaderRange(self.body)
        raise AssertionError(f"Unexpected MACE range: {address}")


def test_read_map_bonds_and_bonds_input_split_descriptions_from_funded_cashflows(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import bonds_lite_input as shared

    stage_three = _stage_three()
    path = tmp_path / "MAP.xlsx"
    path.write_bytes(b"synthetic")
    headers = _map_headers(stage_three)
    body = [0.0] * (stage_three._column_number("ATT") - 2)
    body[0] = "XS0000000001"
    body[2], body[3], body[5], body[7] = "GBP", "Infrastructure", "A", True
    body[stage_three._column_number("Q") - stage_three._column_number("C")] = -1.0
    body[stage_three._column_number("Q") - stage_three._column_number("C") + 1 :] = [
        1.0
    ] * (MONTH_COUNT - 1)
    book = SimpleNamespace(sheets={"I. Bonds Data": _MapSheet(headers, [body])})
    monkeypatch.setattr(
        stage_three,
        "_excel_application",
        lambda: nullcontext(object()),
    )
    monkeypatch.setattr(
        shared, "owned_book",
        lambda application, workbook_path, read_only=True: nullcontext(book),
    )

    result = stage_three.read_map_bonds(path)

    assert stage_three.read_bonds_negative_isins(path) == {"XS0000000001"}
    assert not hasattr(result, "map_negative_isins")
    assert result.by_isin["XS0000000001"] == {
        "isin": "XS0000000001",
        "currency": "GBP",
        "sector": "Infrastructure",
        "rating": "A",
        "index_linked": True,
    }
    assert "cashflows" not in result.by_isin["XS0000000001"]


def test_read_mace_funding_maps_months_and_identifies_negative_leg(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import bonds_lite_input as shared

    stage_three = _stage_three()
    path = tmp_path / "MACE.xlsb"
    path.write_bytes(b"synthetic")
    headers = _mace_headers(stage_three)
    body = [0.0] * (stage_three._column_number("AUM") - 1)
    body[0], body[1], body[2], body[7], body[33] = (
        "XS0000000001", 0.034, "Infrastructure", 17.5, 125.0
    )
    body[34], body[35] = -2.5, -1.25
    book = SimpleNamespace(
        sheets={"C. Assets - Funding": _MaceSheet(headers, [body])}
    )
    monkeypatch.setattr(
        stage_three,
        "_excel_application",
        lambda: nullcontext(object()),
    )
    monkeypatch.setattr(
        shared, "owned_book",
        lambda application, workbook_path, read_only=True: nullcontext(book),
    )

    table = stage_three.read_mace_funding(path)
    funded = stage_three.identify_funded(table)

    assert len(table.rows) == 1
    assert table.rows[0]["funding_cashflows"][:3] == [-2.5, -1.25, 0.0]
    assert _funded_isins(funded) == {"XS0000000001"}


def test_main_dispatches_stage_three_to_the_separate_module(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import xlwings as xw
    import bonds_lite_input as workflow
    import sst_stage_three as stage_three

    calls: list[object] = []
    caller = SimpleNamespace(api=SimpleNamespace(FullName="Draft.xlsm"))
    application = SimpleNamespace(
        api=SimpleNamespace(Hwnd=1234),
        books=[caller],
    )
    monkeypatch.setattr(xw, "apps", [application])
    monkeypatch.setattr(
        stage_three, "run_stage_three",
        lambda book: calls.append(book) or Path("MACalc input.xlsb"),
    )
    monkeypatch.setattr(
        sys,
        "argv",
        [
            "bonds_lite_input",
            "--caller",
            "Draft.xlsm",
            "--excel-hwnd",
            "1234",
            "--stage",
            "3",
        ],
    )

    assert workflow.main() == 0
    assert calls == [caller]


# --------------------------------------------------------------------------
# Slice B: asset block from the Stage 2 base-scenario CSV and MACE allocations
# --------------------------------------------------------------------------

CSV_NAMED = [
    "ISIN",
    "Unique_ISIN",
    "Proxy_Grouping",
    "IM_Zspreads",
    "Currency",
    "Sector",
    "Start_Rating",
    "Final_Rating",
    "Final_Notched_Rating",
    "Base_PV",
]


def _csv_headers() -> list[str]:
    return [*CSV_NAMED, *[str(month) for month in range(MONTH_COUNT)], "Last_CF"]


def _pv_vector(seed: int) -> list[float]:
    """Base_PV followed by months 0-1199, generated reproducibly."""
    import numpy as np

    generator = np.random.default_rng(seed)
    return [float(value) for value in generator.normal(0.0, 1e6, MONTH_COUNT + 1)]


def _csv_record(isin: str, rating: str = "A", seed: int = 0) -> dict[str, str]:
    values = _pv_vector(seed)
    record = {
        "ISIN": isin,
        "Unique_ISIN": f"{isin}_{rating}",
        "Proxy_Grouping": "GROUP",
        "IM_Zspreads": repr(0.0123456789 + seed * 1e-4),
        "Currency": "GBP",
        "Sector": "financial",
        "Start_Rating": rating,
        "Final_Rating": rating,
        "Final_Notched_Rating": rating,
        "Base_PV": repr(values[0]),
        "Last_CF": "0",
    }
    for month, value in enumerate(values[1:]):
        record[str(month)] = repr(value)
    return record


def _write_csv(
    path: Path,
    records: Sequence[dict[str, str]],
    headers: Sequence[str] | None = None,
) -> Path:
    import csv

    names = list(headers or _csv_headers())
    canonical = _csv_headers()
    with path.open("w", newline="", encoding="utf-8") as stream:
        writer = csv.writer(stream)
        writer.writerow(names)
        for record in records:
            writer.writerow([record.get(name, "") for name in canonical])
    return path


def _bond_attributes(
    values: dict[str, tuple[object, object, object]],
    duplicates: frozenset[str] = frozenset(),
) -> Any:
    stage_three = _stage_three()
    return stage_three.MaceBondAttributes(
        by_isin={
            isin: {"type": asset_type, "index_linked": linked, "fsao": fsao}
            for isin, (asset_type, linked, fsao) in values.items()
        },
        duplicates=duplicates,
    )


def _slice_b_case(
    tmp_path: Path, isins: Sequence[str] = ("XS1", "XS2")
) -> dict[str, Any]:
    """A consistent CSV, MACE and MAP set that passes every Slice B guard."""
    stage_three = _stage_three()
    path = _write_csv(
        tmp_path / "py_bonds_Base_Scenario_nozeros_X.csv",
        [_csv_record(isin, seed=index) for index, isin in enumerate(isins)],
    )
    return {
        "base": stage_three.read_map_base_output(path),
        "bonds": _bond_attributes(
            {isin: ("Bond", "FXD", 0.0) for isin in isins}
        ),
        "allocations": _allocations({isin: (1.0, 0.0, 0.0) for isin in isins}),
        "map_bonds": _map_bonds(*[_map_row(isin) for isin in isins]),
        "funded": [],
        "valuation_date": datetime(2026, 6, 30),
    }


def _reconcile(stage_three: Any, case: dict[str, Any]) -> Any:
    return stage_three.reconcile_slice_b(
        case["base"],
        case["bonds"],
        case["allocations"],
        case["map_bonds"],
        case["funded"],
        case["valuation_date"],
    )


def test_map_base_output_pattern_is_one_changeable_module_constant() -> None:
    stage_three = _stage_three()

    assert stage_three.MAP_BASE_OUTPUT_PATTERN == "py_bonds_Base_Scenario_nozeros"


@pytest.mark.parametrize("count", [0, 2])
def test_find_map_base_output_requires_exactly_one_match(
    tmp_path: Path, count: int
) -> None:
    stage_three = _stage_three()
    (tmp_path / "py_bonds_Base_Scenario_BWS.csv").write_text("x")
    for index in range(count):
        (tmp_path / f"py_bonds_Base_Scenario_nozeros_{index}.csv").write_text("x")

    with pytest.raises(stage_three.StageThreeError) as error:
        stage_three.find_map_base_output(tmp_path)

    assert "py_bonds_Base_Scenario_nozeros" in str(error.value)
    if count == 2:
        assert "nozeros_0.csv" in str(error.value)
        assert "nozeros_1.csv" in str(error.value)


def test_find_map_base_output_returns_the_single_match_without_recursing(
    tmp_path: Path,
) -> None:
    stage_three = _stage_three()
    expected = tmp_path / "py_bonds_Base_Scenario_nozeros_BWS.csv"
    expected.write_text("x")
    (tmp_path / "nested").mkdir()
    (tmp_path / "nested" / "py_bonds_Base_Scenario_nozeros_old.csv").write_text("x")

    assert stage_three.find_map_base_output(tmp_path) == expected


def test_resolve_map_base_output_uses_the_stage_two_map_folder_only(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import bonds_lite_input as shared

    stage_three = _stage_three()
    run_folder = tmp_path / "20261102_111001"
    map_folder = run_folder / "JRL MAP"
    bonds_folder = run_folder / "JRL Bonds"
    map_folder.mkdir(parents=True)
    bonds_folder.mkdir()
    (bonds_folder / "py_bonds_Base_Scenario_nozeros_BWS.csv").write_text("x")
    calls: list[Path] = []

    def model_inputs(folder: Path) -> dict[str, object]:
        calls.append(folder)
        return {
            "MAP": SimpleNamespace(output_folder=map_folder),
            "Bonds": SimpleNamespace(output_folder=bonds_folder),
        }

    monkeypatch.setattr(shared, "read_model_inputs", model_inputs)

    with pytest.raises(stage_three.StageThreeError, match="JRL MAP"):
        stage_three.resolve_map_base_output(run_folder)

    expected = map_folder / "py_bonds_Base_Scenario_nozeros_BWS.csv"
    expected.write_text("x")
    assert stage_three.resolve_map_base_output(run_folder) == expected
    assert calls == [run_folder, run_folder]


def test_read_map_base_output_keeps_file_order_and_exact_values(
    tmp_path: Path,
) -> None:
    stage_three = _stage_three()
    records = [_csv_record("XS2", "BBB", seed=2), _csv_record("XS1", "A", seed=1)]
    path = _write_csv(tmp_path / "base.csv", records)

    base = stage_three.read_map_base_output(path)

    assert [row.isin for row in base.rows] == ["XS2", "XS1"]
    first = base.rows[0]
    assert first.unique_isin == "XS2_BBB"
    assert first.zspread == float(records[0]["IM_Zspreads"])
    assert (first.currency, first.sector, first.rating) == (
        "GBP",
        "financial",
        "BBB",
    )
    assert list(first.pv_and_cashflows) == _pv_vector(2)


def test_read_map_base_output_reads_final_rating_separately_from_notched_rating(
    tmp_path: Path,
) -> None:
    stage_three = _stage_three()
    record = _csv_record("XS1", "BBB")
    record["Final_Notched_Rating"] = "BBB+"

    row = stage_three.read_map_base_output(
        _write_csv(tmp_path / "base.csv", [record])
    ).rows[0]

    assert (row.unique_isin, row.rating, row.final_rating) == ("XS1_BBB", "BBB+", "BBB")


@pytest.mark.parametrize(
    ("column", "index"),
    [("A", 0), ("B", 1), ("D", 3), ("E", 4), ("F", 5), ("H", 7), ("I", 8), ("J", 9)],
)
def test_read_map_base_output_rejects_named_header_drift(
    tmp_path: Path, column: str, index: int
) -> None:
    stage_three = _stage_three()
    headers = _csv_headers()
    headers[index] = "Renamed"
    path = _write_csv(tmp_path / "base.csv", [_csv_record("XS1")], headers)

    with pytest.raises(stage_three.StageThreeError, match=f"{column} must be"):
        stage_three.read_map_base_output(path)


@pytest.mark.parametrize(
    "change",
    [
        lambda headers: headers.__setitem__(10, "1"),  # month 0 missing
        lambda headers: headers.__setitem__(1209, "1200"),  # last month wrong
    ],
)
def test_read_map_base_output_requires_month_headers_0_to_1199(
    tmp_path: Path, change: Any
) -> None:
    stage_three = _stage_three()
    headers = _csv_headers()
    change(headers)
    path = _write_csv(tmp_path / "base.csv", [_csv_record("XS1")], headers)

    with pytest.raises(stage_three.StageThreeError, match="K:ATN"):
        stage_three.read_map_base_output(path)


@pytest.mark.parametrize(
    ("field", "value"),
    [
        ("Unique_ISIN", ""),
        ("Sector", " "),
        ("Final_Notched_Rating", "#N/A"),
        ("Final_Rating", ""),
        ("Currency", ""),
        ("IM_Zspreads", "abc"),
        ("Base_PV", ""),
        ("0", "not-a-number"),
        ("1199", "inf"),
    ],
)
def test_read_map_base_output_rejects_blanks_errors_and_non_numeric_values(
    tmp_path: Path, field: str, value: str
) -> None:
    stage_three = _stage_three()
    bad = _csv_record("XS_BAD", seed=1)
    bad[field] = value
    path = _write_csv(tmp_path / "base.csv", [_csv_record("XS_OK"), bad])

    with pytest.raises(stage_three.StageThreeError) as error:
        stage_three.read_map_base_output(path)

    assert "XS_BAD" in str(error.value)
    assert "XS_OK" not in str(error.value)


@pytest.mark.parametrize("field", ["ISIN", "Unique_ISIN"])
def test_read_map_base_output_rejects_duplicate_identifiers(
    tmp_path: Path, field: str
) -> None:
    stage_three = _stage_three()
    first, second = _csv_record("XS1"), _csv_record("XS2", seed=1)
    second[field] = first[field]
    path = _write_csv(tmp_path / "base.csv", [first, second])

    with pytest.raises(stage_three.StageThreeError, match="(?i)duplicate") as error:
        stage_three.read_map_base_output(path)

    assert first[field] in str(error.value)


def test_build_asset_rows_maps_every_column_for_one_asset(tmp_path: Path) -> None:
    stage_three = _stage_three()
    record = _csv_record("XS1", "BBB+", seed=7)
    record["Sector"] = "utilities"
    record["Currency"] = "EUR"
    base = stage_three.read_map_base_output(
        _write_csv(tmp_path / "base.csv", [record])
    )
    bonds = _bond_attributes({"XS1": ("PP", "RPI(0.0,5.0)", 12.5)})
    allocations = _allocations({"XS1": (0.25, 0.75, 1.1e-16)})

    rows = stage_three.build_asset_rows(base, bonds, allocations)

    assert len(rows) == 1
    row = rows[0]
    assert len(row) == stage_three._column_number("ATQ")
    assert row[:13] == [
        0.25,  # A: Comp. A %
        0.75,  # B: Comp. B %
        1.1e-16,  # C: Partial CF Flag, copied exactly
        "XS1_BBB+",  # D: Unique_ISIN
        "XS1",  # E: ISIN as Name
        float(record["IM_Zspreads"]) * 10000,  # F: Zspread in basis points
        "PP",  # G: Type from I. Bonds
        "utilities",  # H: Sector
        "BBB+",  # I: Final_Notched_Rating
        "RPI(0.0,5.0)",  # J: Index Linked? from I. Bonds
        "EUR",  # K: Currency
        12.5,  # L: FSAO from I. Bonds
        _pv_vector(7)[0],  # M: Base_PV
    ]
    assert row[13:] == _pv_vector(7)[1:]  # N:ATQ from CSV months 0-1199


def test_reconcile_slice_b_accepts_consistent_inputs(tmp_path: Path) -> None:
    stage_three = _stage_three()

    result = _reconcile(stage_three, _slice_b_case(tmp_path))

    assert result.map_difference_warning() is None
    assert (result.mace_only_bonds, result.mace_only_allocations) == (0, 0)


def test_reconcile_slice_b_allows_and_counts_mace_only_rows(tmp_path: Path) -> None:
    stage_three = _stage_three()
    case = _slice_b_case(tmp_path)
    case["bonds"] = _bond_attributes(
        {
            "XS1": ("Bond", "FXD", 0.0),
            "XS2": ("Bond", "FXD", 0.0),
            "Inf_swap_FXD": ("ILS", "FXD", 0.0),
        },
        duplicates=frozenset({"Inf_swap_FXD"}),
    )
    case["allocations"] = _allocations(
        {
            "XS1": (1.0, 0.0, 0.0),
            "XS2": (1.0, 0.0, 0.0),
            "Inf_swap_FXD": (1.0, 0.0, 0.0),
            "Inf_swap_RPI": (1.0, 0.0, 0.0),
        }
    )

    result = _reconcile(stage_three, case)

    assert (result.mace_only_bonds, result.mace_only_allocations) == (1, 2)


@pytest.mark.parametrize(
    ("mutate", "category"),
    [
        (lambda case: case["bonds"].by_isin.pop("XS2"), "I. Bonds"),
        (lambda case: case["allocations"].by_isin.pop("XS2"), "Control"),
        (
            lambda case: case.update(
                bonds=_bond_attributes(
                    {"XS1": ("Bond", "FXD", 0.0), "XS2": ("Bond", "FXD", 0.0)},
                    duplicates=frozenset({"XS2"}),
                )
            ),
            "(?i)duplicate",
        ),
        (
            lambda case: case.update(
                allocations=_allocations(
                    {"XS1": (1.0, 0.0, 0.0), "XS2": (1.0, 0.0, 0.0)},
                    duplicates=frozenset({"XS2"}),
                )
            ),
            "(?i)duplicate",
        ),
        (lambda case: case["bonds"].by_isin["XS2"].update(type=" "), "Type"),
        (lambda case: case["bonds"].by_isin["XS2"].update(type=7), "Type"),
        (
            lambda case: case["bonds"].by_isin["XS2"].update(index_linked=None),
            "Index Linked",
        ),
        (lambda case: case["bonds"].by_isin["XS2"].update(fsao="#N/A"), "FSAO"),
        (lambda case: case["bonds"].by_isin["XS2"].update(fsao=True), "FSAO"),
        (lambda case: case["allocations"].by_isin["XS2"].update(a=None), "Comp. A"),
        (lambda case: case["allocations"].by_isin["XS2"].update(b="x"), "Comp. B"),
        (
            lambda case: case["allocations"].by_isin["XS2"].update(c="#VALUE!"),
            "Partial CF",
        ),
    ],
)
def test_reconcile_slice_b_stops_and_names_the_offending_isin(
    tmp_path: Path, mutate: Any, category: str
) -> None:
    stage_three = _stage_three()
    case = _slice_b_case(tmp_path)
    mutate(case)

    with pytest.raises(stage_three.StageThreeError, match=category) as error:
        _reconcile(stage_three, case)

    assert "XS2" in str(error.value)
    assert "XS1" not in str(error.value)


def test_reconcile_slice_b_reports_every_problem_in_one_message(
    tmp_path: Path,
) -> None:
    stage_three = _stage_three()
    case = _slice_b_case(tmp_path)
    case["bonds"].by_isin.pop("XS1")
    case["allocations"].by_isin["XS2"]["a"] = None

    with pytest.raises(stage_three.StageThreeError) as error:
        _reconcile(stage_three, case)

    message = str(error.value)
    assert "I. Bonds" in message and "XS1" in message
    assert "Comp. A" in message and "XS2" in message


def test_reconcile_slice_b_checks_funded_isins_against_control(
    tmp_path: Path,
) -> None:
    stage_three = _stage_three()
    case = _slice_b_case(tmp_path)
    case["funded"] = [_mace_row("FUNDED_ONLY", [-1.0] + [0.0] * (MONTH_COUNT - 1))]

    with pytest.raises(stage_three.StageThreeError, match="Control") as error:
        _reconcile(stage_three, case)

    assert "FUNDED_ONLY" in str(error.value)


def test_reconcile_slice_b_warns_but_does_not_stop_on_csv_map_difference(
    tmp_path: Path,
) -> None:
    stage_three = _stage_three()
    case = _slice_b_case(tmp_path)
    case["map_bonds"] = _map_bonds(_map_row("XS1"), _map_row("MAP_ONLY"))

    result = _reconcile(stage_three, case)

    warning = result.map_difference_warning()
    assert warning is not None
    assert "MAP_ONLY" in warning and "XS2" in warning


@pytest.mark.parametrize("mace_date", [datetime(2026, 3, 31), None, "30/06/2026"])
def test_reconcile_slice_b_stops_on_valuation_date_mismatch(
    tmp_path: Path, mace_date: object
) -> None:
    stage_three = _stage_three()
    case = _slice_b_case(tmp_path)
    case["allocations"] = stage_three.MaceAllocations(
        by_isin=case["allocations"].by_isin,
        duplicates=frozenset(),
        valuation_date=mace_date,
    )

    with pytest.raises(stage_three.StageThreeError, match="(?i)valuation date"):
        _reconcile(stage_three, case)


def test_funded_asset_appears_in_both_blocks_with_distinct_column_d(
    tmp_path: Path,
) -> None:
    stage_three = _stage_three()
    isin = "GR0000000013_JRINFRA_Aviva"
    case = _slice_b_case(tmp_path, (isin, "XS2"))
    funded = [
        _mace_row(
            isin, [-1.0] + [0.0] * (MONTH_COUNT - 1), type="Infrastructure"
        )
    ]
    case["funded"] = funded
    case["map_bonds"] = _map_bonds(
        _map_row(isin, [-1.0] + [1.0] * (MONTH_COUNT - 1)), _map_row("XS2")
    )

    _reconcile(stage_three, case)
    funded_rows = stage_three.build_funded_rows(
        case["map_bonds"], funded, case["allocations"]
    )
    asset_rows = stage_three.build_asset_rows(
        case["base"], case["bonds"], case["allocations"]
    )

    column_d = [row[3] for row in [*funded_rows, *asset_rows]]
    assert column_d == [isin, f"{isin}_A", "XS2_A"]
    assert len(set(column_d)) == len(column_d)
    assert funded_rows[0][:3] == asset_rows[0][:3] == [1.0, 0.0, 0.0]


@pytest.mark.parametrize(
    ("funded_count", "expected"),
    [(0, [(21, 0), (21, 3)]), (5, [(21, 5), (26, 3)])],
)
def test_output_blocks_place_assets_directly_below_funded_rows(
    funded_count: int, expected: list[tuple[int, int]]
) -> None:
    stage_three = _stage_three()
    layout = stage_three.TemplateLayout(2, 12, 13, 20, 21)

    blocks = stage_three.output_blocks(
        layout, [["f"]] * funded_count, [["a"]] * 3
    )

    assert [(start, len(rows)) for start, rows in blocks] == expected


def _bonds_headers() -> list[object]:
    return [
        None, "Inflow", None, "ISIN", "Name", "Zspread", "Type", "Sector",
        "Rating", "Index Linked?", "Currency", "FSAO",
    ]


def _bonds_body_row(isin: object, asset_type: object = "Bond") -> list[object]:
    # D:L = ISIN, Name, Zspread, Type, Sector, Rating, Index Linked?, Currency, FSAO
    return [isin, "Company", 1.0, asset_type, "S", "A", "FXD", "GBP", 0.5]


def test_parse_mace_bond_attributes_keys_on_column_d_and_records_duplicates() -> None:
    stage_three = _stage_three()
    body = [
        _bonds_body_row("XS1"),
        _bonds_body_row("XS2", "PP"),
        _bonds_body_row("XS2", "Bond"),
        _bonds_body_row(None),
    ]

    result = stage_three.parse_mace_bond_attributes(_bonds_headers(), body)

    assert result.by_isin["XS1"] == {
        "type": "Bond", "index_linked": "FXD", "fsao": 0.5
    }
    assert result.duplicates == frozenset({"XS2"})


@pytest.mark.parametrize(("index", "column"), [(3, "D"), (6, "G"), (9, "J"), (11, "L")])
def test_parse_mace_bond_attributes_rejects_header_drift(
    index: int, column: str
) -> None:
    stage_three = _stage_three()
    headers = _bonds_headers()
    headers[index] = "Renamed"

    with pytest.raises(stage_three.StageThreeError, match=f"{column}7"):
        stage_three.parse_mace_bond_attributes(headers, [_bonds_body_row("XS1")])


CONTROL_HEADERS = ["ISIN", "Comp. A % ", "Comp. B %", "Comp. C %", "Partial CF Flag"]


def _control_body(*isins: str, trailing_blank: int = 50) -> list[list[object]]:
    rows: list[list[object]] = [[isin, 1.0, 0.0, 0.0, 1e-16] for isin in isins]
    return rows + [[None] * 5 for _ in range(trailing_blank)]


def test_parse_control_table_reads_k_to_o_until_first_blank() -> None:
    stage_three = _stage_three()

    table = stage_three.parse_control_table(
        CONTROL_HEADERS, _control_body("XS1", "XS2", "XS2")
    )

    assert table.last_row == 12
    assert table.by_isin["XS1"] == {"a": 1.0, "b": 0.0, "c": 1e-16}
    assert table.duplicates == frozenset({"XS2"})


@pytest.mark.parametrize(
    "change",
    [
        lambda headers, body: headers.__setitem__(0, "Asset"),
        lambda headers, body: headers.__setitem__(1, "Comp A"),
        lambda headers, body: headers.__setitem__(4, "Partial Flag"),
        lambda headers, body: body.__setitem__(30, ["LATER", 1.0, 0.0, 0.0, 0.0]),
    ],
)
def test_parse_control_table_stops_when_table_is_not_contained_in_k_to_o(
    change: Any,
) -> None:
    stage_three = _stage_three()
    headers = list(CONTROL_HEADERS)
    body = _control_body("XS1", "XS2")
    change(headers, body)

    with pytest.raises(stage_three.StageThreeError, match="not contained within K:O"):
        stage_three.parse_control_table(headers, body)


def test_parse_control_table_ignores_keys_more_than_50_rows_below_the_gap() -> None:
    stage_three = _stage_three()
    body = _control_body("XS1", trailing_blank=50) + [["FAR", 1.0, 0.0, 0.0, 0.0]]

    table = stage_three.parse_control_table(CONTROL_HEADERS, body)

    assert set(table.by_isin) == {"XS1"}


def test_control_side_check_covers_column_p_only() -> None:
    assert _stage_three().CONTROL_SIDE_COLUMNS == ("P", "P")


def test_require_control_side_empty_rejects_data_in_column_p() -> None:
    stage_three = _stage_three()
    side = [[None] for _ in range(4)]
    stage_three.require_control_side_empty(side)

    side[2][0] = "Comp. D %"
    with pytest.raises(stage_three.StageThreeError, match=r"found data in P11"):
        stage_three.require_control_side_empty(side)


def test_read_mace_allocations_ignores_other_panels_beyond_column_p() -> None:
    stage_three = _stage_three()
    sheet = _control_sheet(stage_three)
    sheet.values["Q9:U11"] = [["panel"] * 5 for _ in range(3)]

    result = stage_three.read_mace_allocations(
        SimpleNamespace(sheets={"Control": sheet})
    )

    assert set(result.by_isin) == {"XS1", "XS2"}
    assert "P9:P11" in sheet.requested
    assert "Q9:U11" not in sheet.requested


class _AddressSheet:
    """Fake worksheet returning fixed values for exact range addresses."""

    def __init__(self, values: dict[str, object], ends: dict[str, int]) -> None:
        self.values = values
        self.ends = ends
        self.requested: list[str] = []

    def range(self, address: str) -> _ReaderRange:
        self.requested.append(address)
        if address in self.ends:
            return _ReaderRange(row=self.ends[address])
        if address in self.values:
            return _ReaderRange(self.values[address])
        raise AssertionError(f"Unexpected range: {address}")


def _control_sheet(stage_three: Any, side_value: object = None) -> _AddressSheet:
    first, last = stage_three.CONTROL_SIDE_COLUMNS
    width = stage_three._column_number(last) - stage_three._column_number(first) + 1
    side = [[None] * width for _ in range(3)]
    side[1][0] = side_value
    return _AddressSheet(
        {
            "C9": datetime(2026, 6, 30),
            "K9:O9": CONTROL_HEADERS,
            "K10:O61": _control_body("XS1", "XS2"),
            f"{first}9:{last}11": side,
        },
        {"K1048576": 11},
    )


def _bonds_sheet() -> _AddressSheet:
    return _AddressSheet(
        {
            "A7:L7": _bonds_headers(),
            "D8:L9": [_bonds_body_row("XS1"), _bonds_body_row("XS2")],
        },
        {"D1048576": 9},
    )


def test_read_mace_allocations_reads_table_and_valuation_date() -> None:
    stage_three = _stage_three()
    book = SimpleNamespace(sheets={"Control": _control_sheet(stage_three)})

    result = stage_three.read_mace_allocations(book)

    assert set(result.by_isin) == {"XS1", "XS2"}
    assert result.valuation_date == datetime(2026, 6, 30)


def test_read_mace_allocations_stops_on_data_beside_the_table() -> None:
    stage_three = _stage_three()
    book = SimpleNamespace(
        sheets={"Control": _control_sheet(stage_three, side_value="extra")}
    )

    with pytest.raises(stage_three.StageThreeError, match="not contained within K:O"):
        stage_three.read_mace_allocations(book)


def test_read_mace_bond_attributes_reads_rows_below_header_row_7() -> None:
    stage_three = _stage_three()
    book = SimpleNamespace(sheets={"I. Bonds": _bonds_sheet()})

    result = stage_three.read_mace_bond_attributes(book)

    assert set(result.by_isin) == {"XS1", "XS2"}


def test_read_mace_inputs_opens_mace_once_for_all_three_tabs(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    import bonds_lite_input as shared

    stage_three = _stage_three()
    path = tmp_path / "MACE.xlsb"
    path.write_bytes(b"synthetic")
    funding_body = [0.0] * (stage_three._column_number("AUM") - 1)
    funding_body[0], funding_body[2] = "XS1", "Bond"
    book = SimpleNamespace(
        sheets={
            "C. Assets - Funding": _MaceSheet(
                _mace_headers(stage_three), [funding_body]
            ),
            "I. Bonds": _bonds_sheet(),
            "Control": _control_sheet(stage_three),
        }
    )
    opened: list[tuple[Path, bool]] = []

    def owned_book(application: object, workbook: Path, read_only: bool = True):
        opened.append((workbook, read_only))
        return nullcontext(book)

    monkeypatch.setattr(
        stage_three, "_excel_application", lambda: nullcontext(object())
    )
    monkeypatch.setattr(shared, "owned_book", owned_book)

    result = stage_three.read_mace_inputs(path)

    assert opened == [(path.resolve(), True)]
    assert len(result.funding.rows) == 1
    assert set(result.bonds.by_isin) == {"XS1", "XS2"}
    assert set(result.allocations.by_isin) == {"XS1", "XS2"}


@pytest.mark.parametrize(
    ("a_values", "g_values", "row"),
    [
        (["x", None, "y"], ["Mortgage", "ILS", "Bond"], "3"),
        (["x", "", "y"], ["Mortgage", "ILS", "Bond"], "3"),
        (["x", 0.0, "y"], ["Mortgage", "Bonds", "Bond"], "3"),
        (["x", 0.0, "y"], ["Mortgage", "ILS", "infrastructure"], "4"),
    ],
)
def test_check_macalc_asset_table_stops_on_blank_a_or_unknown_type(
    a_values: list[object], g_values: list[object], row: str
) -> None:
    stage_three = _stage_three()

    with pytest.raises(stage_three.StageThreeError) as error:
        stage_three.check_macalc_asset_table(a_values, g_values)

    assert row in str(error.value)


def test_check_macalc_asset_table_accepts_zero_a_and_any_case_bond() -> None:
    stage_three = _stage_three()

    stage_three.check_macalc_asset_table(
        [1.0, 0.0, 0.0, 0.5, 0.0, 0.0, 1.0],
        ["Mortgage", "ILS", "BOND", "bond", "Infrastructure", "PP", "Swap"],
    )


def test_run_stage_three_slice_b_guard_failure_keeps_previous_output(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    stage_three = _stage_three()
    caller, output, writes, _ = _mock_stage_three_run(
        monkeypatch, stage_three, tmp_path, existing_output=True
    )

    def fail(*args: object) -> None:
        raise stage_three.StageThreeError("CSV ISINs missing from MACE I. Bonds: XS9")

    monkeypatch.setattr(stage_three, "reconcile_slice_b", fail)

    with pytest.raises(stage_three.StageThreeError, match="XS9"):
        stage_three.run_stage_three(caller)

    assert writes[0]["status"] == "Running"
    assert writes[-1]["status"] == "Failed"
    assert writes[-1]["last_output"] == "Bonds Lite output/previous"
    assert output.with_name(EARLIER_OUTPUT_NAME).read_bytes() == b"previous"
    assert [path.name for path in output.parent.iterdir()] == [EARLIER_OUTPUT_NAME]


def test_run_stage_three_writes_both_blocks_and_warns_on_csv_map_difference(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    stage_three = _stage_three()
    caller, output, writes, messages = _mock_stage_three_run(
        monkeypatch, stage_three, tmp_path, existing_output=False
    )
    funded_rows, asset_rows = [["funded"]], [["asset1"], ["asset2"]]
    written: dict[str, object] = {}
    monkeypatch.setattr(stage_three, "reconcile_funded", lambda *args: False)
    monkeypatch.setattr(
        stage_three, "build_funded_rows", lambda *args: funded_rows
    )
    monkeypatch.setattr(stage_three, "build_asset_rows", lambda *args: asset_rows)
    monkeypatch.setattr(
        stage_three, "build_asset_keys", lambda *args, **kwargs: ()
    )
    monkeypatch.setattr(
        stage_three,
        "reconcile_slice_b",
        lambda *args: stage_three.SliceBReconciliation(
            csv_only=("CSV_ONLY",), map_only=("MAP_ONLY",)
        ),
    )

    def write_output(template, staged, layout, rows, assets=(), isin_lists=()):
        written["funded"], written["assets"] = rows, assets
        staged.write_bytes(b"new")
        return 20

    def verify(path, expected, **kwargs):
        written["verified"] = expected

    monkeypatch.setattr(stage_three, "write_macalc_output", write_output)
    monkeypatch.setattr(stage_three, "verify_macalc_output", verify)

    assert stage_three.run_stage_three(caller) == output

    assert written == {
        "funded": funded_rows,
        "assets": asset_rows,
        "verified": [*funded_rows, *asset_rows],
    }
    assert writes[-1]["status"] == "Succeeded"
    assert any("CSV_ONLY" in text and "MAP_ONLY" in text for text in messages)

# --- Slice C: align the MACalc ISIN lists with I. Assets column D ---------------


def _list_spec(name: str) -> Any:
    return {spec.name: spec for spec in _stage_three().MACALC_ISIN_LISTS}[name]


def _isin_list(
    name: str, keys: Sequence[object], values: Sequence[object] | None = None
) -> Any:
    stage_three = _stage_three()
    spec = _list_spec(name)
    if spec.value_column is not None and values is None:
        values = [0.5] * len(keys)
    return stage_three.IsinList(
        spec,
        tuple(keys),
        None if values is None else tuple(values),
        stage_three.LIST_FIRST_ROW + len(keys) - 1,
    )


def _asset_key(
    row: int, d: str, e: str | None = None, block: str = "asset", rating: str = "A"
) -> Any:
    return _stage_three().AssetKey(row, d, d if e is None else e, block, rating)


def _real_shape_keys() -> list[Any]:
    """Mortgage, ILS, funded and suffixed asset rows mimicking the real D/E shapes."""
    return [
        _asset_key(2, "Note 1", block="template", rating="NR"),
        _asset_key(3, "Inf_swap_FXD", "Inf_swap", block="template", rating="NR"),
        _asset_key(4, "XSFUND1", block="funded"),
        _asset_key(5, "X_JRUSPP_Metlife_BBB", "X_JRUSPP_Metlife", rating="BBB"),
        _asset_key(6, "XS0000000001_A", "XS0000000001"),
    ]


def _align(lists: Sequence[Any], keys: Sequence[Any]) -> Any:
    return _stage_three().align_isin_lists(lists, keys)


def test_align_keeps_exact_d_matches_for_note_inf_swap_and_funded_isin() -> None:
    source = _isin_list("tbl_ForcedCompAperc", ["Note 1", "Inf_swap_FXD", "XSFUND1"])

    result = _align([source], _real_shape_keys()).get("tbl_ForcedCompAperc")

    assert result.keys == ("Note 1", "Inf_swap_FXD", "XSFUND1")
    assert (result.kept, result.remapped, result.added) == (3, 0, 0)
    assert result.unmatched == ()


def test_align_remaps_asset_e_to_unique_isin_copying_the_value_exactly() -> None:
    source = _isin_list(
        "tbl_MinCompAperc", ["X_JRUSPP_Metlife"], [0.10000000149011612]
    )

    result = _align([source], _real_shape_keys()).get("tbl_MinCompAperc")

    assert result.keys == ("X_JRUSPP_Metlife_BBB",)
    assert result.values == (0.10000000149011612,)
    assert (result.kept, result.remapped) == (0, 1)


def test_align_never_uses_an_ils_e_value_as_a_key() -> None:
    source = _isin_list("lst_pre16assets", ["Inf_swap"])

    result = _align([source], _real_shape_keys()).get("lst_pre16assets")

    assert result.keys == ("Inf_swap",)
    assert result.unmatched == ("Inf_swap",)


def test_align_keeps_unmatched_entries_in_place_and_counts_them() -> None:
    source = _isin_list(
        "tbl_MinCompAperc", ["STALE1", "X_JRUSPP_Metlife", "STALE2"], [0.1, 0.2, 0.3]
    )

    result = _align([source], _real_shape_keys()).get("tbl_MinCompAperc")

    assert result.keys == ("STALE1", "X_JRUSPP_Metlife_BBB", "STALE2")
    assert result.values == (0.1, 0.2, 0.3)
    assert result.unmatched == ("STALE1", "STALE2")
    assert result.matched_keys == 1


def test_align_fans_out_a_funded_isin_that_is_also_an_asset_e() -> None:
    keys = [*_real_shape_keys(), _asset_key(7, "XSFUND1_A", "XSFUND1")]
    source = _isin_list("tbl_ForcedCompAperc", ["XSFUND1"], [0.25])

    result = _align([source], keys).get("tbl_ForcedCompAperc")

    assert result.keys == ("XSFUND1", "XSFUND1_A")
    assert result.values == (0.25, 0.25)
    assert result.added == 1
    assert result.fan_outs == (("XSFUND1", ("XSFUND1", "XSFUND1_A")),)


def test_align_fans_out_one_e_at_two_ratings_in_row_order_not_sorted() -> None:
    keys = [
        _asset_key(40, "X_A", "X", rating="A"),
        _asset_key(31, "X_BBB", "X", rating="BBB"),
    ]
    source = _isin_list("tbl_MinCompAperc", ["BEFORE", "X", "AFTER"], [0.1, 0.2, 0.3])

    result = _align([source], keys).get("tbl_MinCompAperc")

    assert result.keys == ("BEFORE", "X_BBB", "X_A", "AFTER")
    assert result.values == (0.1, 0.2, 0.2, 0.3)
    assert result.added == 1


def test_align_fans_out_funded_and_two_asset_legs_in_row_order() -> None:
    keys = [
        _asset_key(21, "X", block="funded"),
        _asset_key(40, "X_A", "X", rating="A"),
        _asset_key(31, "X_BBB", "X", rating="BBB"),
    ]
    source = _isin_list("lst_REITs", ["X"])

    result = _align([source], keys).get("lst_REITs")

    assert result.keys == ("X", "X_BBB", "X_A")
    assert result.values is None
    assert result.added == 2


def test_align_output_depends_only_on_list_order_and_row_order() -> None:
    keys = [
        _asset_key(21, "X", block="funded"),
        _asset_key(31, "X_BBB", "X", rating="BBB"),
        _asset_key(40, "X_A", "X", rating="A"),
        _asset_key(41, "Y_AA", "Y", rating="AA"),
    ]
    lists = [
        _isin_list("tbl_MinCompAperc", ["Y", "X", "Z"], [0.1, 0.2, 0.3]),
        _isin_list("lst_pre16assets", ["X", "Y"]),
    ]

    first = _align(lists, keys)
    second = _align(lists, list(reversed(keys)))

    assert first == second == _align(lists, keys)
    assert first.get("tbl_MinCompAperc").keys == ("Y_AA", "X", "X_BBB", "X_A", "Z")


@pytest.mark.parametrize("name", ["tbl_ForcedCompAperc", "tbl_MinCompAperc"])
def test_align_removes_same_value_table_duplicates_keeping_the_first(
    name: str,
) -> None:
    keys = [
        _asset_key(31, "X_BBB", "X", rating="BBB"),
        _asset_key(40, "X_A", "X", rating="A"),
    ]
    source = _isin_list(name, ["K", "X", "K", "X_BBB"], [0.5, 0.2, 0.5, 0.2])

    result = _align([source], keys).get(name)

    assert result.keys == ("K", "X_BBB", "X_A")
    assert result.values == (0.5, 0.2, 0.2)
    assert result.duplicates_removed == 2


@pytest.mark.parametrize(
    ("keys", "values", "named"),
    [
        (["K", "K"], [0.5, 0.6], "'K'"),
        (["X", "X_BBB"], [0.5, 0.7], "'X_BBB'"),
    ],
)
def test_align_stops_on_conflicting_table_duplicates_naming_key_and_values(
    keys: list[str], values: list[float], named: str
) -> None:
    stage_three = _stage_three()
    asset_keys = [
        _asset_key(31, "X_BBB", "X", rating="BBB"),
        _asset_key(40, "X_A", "X", rating="A"),
    ]
    source = _isin_list("tbl_MinCompAperc", keys, values)

    with pytest.raises(stage_three.StageThreeError) as error:
        _align([source], asset_keys)

    message = str(error.value)
    assert "tbl_MinCompAperc" in message and named in message
    assert str(values[0]) in message and str(values[1]) in message


@pytest.mark.parametrize("name", ["lst_pre16assets", "lst_REITs"])
def test_align_removes_membership_duplicates_and_fan_out_collisions(
    name: str,
) -> None:
    keys = [
        _asset_key(31, "X_BBB", "X", rating="BBB"),
        _asset_key(40, "X_A", "X", rating="A"),
    ]
    source = _isin_list(name, ["P", "X", "P", "X_A"])

    result = _align([source], keys).get(name)

    assert result.keys == ("P", "X_BBB", "X_A")
    assert result.duplicates_removed == 2


def test_align_sizes_reconcile_added_and_removed_counts() -> None:
    keys = [
        _asset_key(31, "X_BBB", "X", rating="BBB"),
        _asset_key(40, "X_A", "X", rating="A"),
    ]
    source = _isin_list("lst_pre16assets", ["P", "X", "P", "X_A", "STALE"])

    result = _align([source], keys).get("lst_pre16assets")

    assert result.source_size == 5
    assert result.mapped_size == result.source_size + result.added == 6
    assert len(result.keys) == result.mapped_size - result.duplicates_removed == 4


@pytest.mark.parametrize("bad_key", [None, "", "  ", "#N/A", 12345.0])
def test_align_stops_on_blank_error_or_non_text_key_naming_the_cell(
    bad_key: object,
) -> None:
    stage_three = _stage_three()
    source = _isin_list("tbl_ForcedCompAperc", ["Note 1", bad_key])

    with pytest.raises(stage_three.StageThreeError, match=r"H8"):
        _align([source], _real_shape_keys())


@pytest.mark.parametrize(
    "bad_value", [None, "#N/A", "0.5", 1.5, -0.1, True, float("nan")]
)
def test_align_stops_on_invalid_table_value_naming_the_key(bad_value: object) -> None:
    stage_three = _stage_three()
    source = _isin_list("tbl_MinCompAperc", ["GOOD", "BAD"], [0.5, bad_value])

    with pytest.raises(stage_three.StageThreeError) as error:
        _align([source], _real_shape_keys())

    assert "L8" in str(error.value) and "'BAD'" in str(error.value)


def test_align_accepts_table_values_at_both_bounds() -> None:
    source = _isin_list("tbl_MinCompAperc", ["LOW", "HIGH"], [0.0, 1.0])

    assert _align([source], _real_shape_keys()).get("tbl_MinCompAperc").values == (
        0.0,
        1.0,
    )


def test_align_stops_when_an_asset_d_is_neither_e_nor_e_with_its_rating() -> None:
    stage_three = _stage_three()
    keys = [*_real_shape_keys(), _asset_key(9, "X_FOO", "X", rating="A")]

    with pytest.raises(stage_three.StageThreeError) as error:
        _align([_isin_list("lst_REITs", ["X"])], keys)

    assert "row 9" in str(error.value) and "X_FOO" in str(error.value)


@pytest.mark.parametrize("rating", ["GOV", "BBB", "A"])
def test_align_reaches_rows_whose_d_is_exactly_e_and_their_own_final_rating(
    rating: str,
) -> None:
    keys = [_asset_key(5, f"X_{rating}", "X", rating=rating)]

    result = _align([_isin_list("lst_REITs", ["X"])], keys).get("lst_REITs")

    assert result.keys == (f"X_{rating}",) and result.remapped == 1


@pytest.mark.parametrize(
    ("d", "rating"),
    [("X_BBB", "BBB+"), ("X_BBB+", "BBB"), ("X_A", "BBB"), ("X_BBB", "BBB ")],
)
def test_align_stops_when_the_suffix_is_not_exactly_the_rows_own_rating(
    d: str, rating: str
) -> None:
    stage_three = _stage_three()

    with pytest.raises(stage_three.StageThreeError, match="row 5"):
        _align([_isin_list("lst_REITs", ["X"])], [_asset_key(5, d, "X", rating=rating)])


def test_align_allows_an_empty_list() -> None:
    result = _align([_isin_list("lst_REITs", [])], _real_shape_keys()).get("lst_REITs")

    assert result.keys == () and result.source_size == 0


def test_align_returns_lists_unchanged_when_no_asset_row_is_suffixed() -> None:
    keys = [
        _asset_key(2, "Note 1", block="template", rating="NR"),
        _asset_key(4, "XSFUND1", block="funded"),
        _asset_key(5, "XS0000000001"),
        _asset_key(6, "XS0000000002"),
    ]
    lists = [
        _isin_list("tbl_ForcedCompAperc", ["Note 1", "XS0000000001"], [1.0, 0.3]),
        _isin_list("tbl_MinCompAperc", ["XS0000000002", "STALE"], [0.1, 0.2]),
        _isin_list("lst_pre16assets", ["XS0000000001", "STALE"]),
        _isin_list("lst_REITs", ["XS0000000002"]),
    ]

    result = _align(lists, keys)

    for source, aligned in zip(lists, result.lists):
        assert aligned.keys == source.keys
        assert aligned.values == source.values
        assert aligned.remapped == aligned.added == 0


def test_align_remaps_suffixed_rows_and_keeps_unsuffixed_rows() -> None:
    keys = [_asset_key(5, "Y"), _asset_key(6, "Z_A", "Z")]
    source = _isin_list("lst_pre16assets", ["Y", "Z"])

    result = _align([source], keys).get("lst_pre16assets")

    assert result.keys == ("Y", "Z_A")
    assert (result.kept, result.remapped) == (1, 1)


def test_build_asset_keys_takes_asset_ratings_from_the_csv_final_rating() -> None:
    stage_three = _stage_three()
    layout = stage_three.TemplateLayout(2, 2, 3, 3, 4)
    fixed = (_asset_key(2, "Note 1", block="template"),)
    funded = [[0, 0, 0, "F1", "F1", 0, "Bond", "S", "A"]]
    assets = [[0, 0, 0, "X_BBB", "X", 0, "Bond", "S", "BBB+"]]

    result = stage_three.build_asset_keys(
        layout, fixed, funded, assets, asset_ratings=["BBB"]
    )

    assert result == (
        fixed[0],
        stage_three.AssetKey(4, "F1", "F1", "funded", "A"),
        stage_three.AssetKey(5, "X_BBB", "X", "asset", "BBB"),
    )


def test_build_asset_keys_stops_when_asset_ratings_do_not_match_the_rows() -> None:
    stage_three = _stage_three()
    layout = stage_three.TemplateLayout(2, 2, 3, 3, 4)
    assets = [[0, 0, 0, "X_BBB", "X", 0, "Bond", "S", "BBB"]]

    with pytest.raises(stage_three.StageThreeError, match="Final_Rating"):
        stage_three.build_asset_keys(layout, (), [], assets, asset_ratings=[])


def test_align_stops_on_a_notched_csv_row_under_the_literal_column_i_rating(
    tmp_path: Path,
) -> None:
    """Column I is notched, so only the CSV Final_Rating can satisfy Guard 7."""
    stage_three = _stage_three()
    record = _csv_record("XS1", "BBB")
    record["Final_Notched_Rating"] = "BBB+"
    base = stage_three.read_map_base_output(_write_csv(tmp_path / "b.csv", [record]))
    bonds = _bond_attributes({"XS1": ("Bond", "FXD", 0.0)})
    rows = stage_three.build_asset_rows(base, bonds, _allocations({"XS1": (1, 0, 0)}))
    layout = stage_three.TemplateLayout(2, 2, 3, 3, 4)
    keys = stage_three.build_asset_keys(
        layout, (), [], rows, asset_ratings=[row.final_rating for row in base.rows]
    )

    result = _align([_isin_list("lst_REITs", ["XS1"])], keys).get("lst_REITs")

    assert rows[0][8] == "BBB+"
    assert result.keys == ("XS1_BBB",) and result.remapped == 1


def test_isin_list_specs_cover_the_four_d_keyed_lists_only() -> None:
    stage_three = _stage_three()

    assert [
        (spec.name, spec.sheet, spec.key_column, spec.value_column)
        for spec in stage_three.MACALC_ISIN_LISTS
    ] == [
        ("tbl_ForcedCompAperc", "I. Inputs", "H", "I"),
        ("tbl_MinCompAperc", "I. Inputs", "K", "L"),
        ("lst_pre16assets", "P. Parameters", "K", None),
        ("lst_REITs", "P. Parameters", "M", None),
    ]
    assert _list_spec("tbl_MinCompAperc").refers_to == (
        "=OFFSET('I. Inputs'!$K$7,0,0,COUNTA('I. Inputs'!$K:$K)-1,2)"
    )


class _GridSheet:
    """Fake worksheet backed by a sparse cell map, serving any A1 rectangle."""

    def __init__(self, cells: dict[str, object]) -> None:
        self.cells = dict(cells)

    @staticmethod
    def _split(cell: str) -> tuple[int, int]:
        letters = "".join(ch for ch in cell if ch.isalpha())
        return _stage_three()._column_number(letters), int(cell[len(letters) :])

    def range(self, address: str) -> _ReaderRange:
        stage_three = _stage_three()
        if ":" not in address:
            column, row = self._split(address)
            if row == 1048576:
                used = [
                    self._split(cell)[1]
                    for cell, value in self.cells.items()
                    if self._split(cell)[0] == column and value is not None
                ]
                return _ReaderRange(row=max(used, default=1))
            return _ReaderRange(self.cells.get(address))
        first, last = (self._split(part) for part in address.split(":"))
        grid = [
            [
                self.cells.get(f"{stage_three._column_letter(column)}{row}")
                for column in range(first[0], last[0] + 1)
            ]
            for row in range(first[1], last[1] + 1)
        ]
        return _ReaderRange(grid)


def _list_cells(
    column: str, title: str, keys: Sequence[object], values_column: str | None = None
) -> dict[str, object]:
    cells: dict[str, object] = {f"{column}5": title}
    for offset, key in enumerate(keys):
        cells[f"{column}{7 + offset}"] = key
        if values_column is not None:
            cells[f"{values_column}{7 + offset}"] = 0.5
    return cells


def _isin_book(
    inputs: dict[str, object] | None = None,
    parameters: dict[str, object] | None = None,
    formulas: dict[str, str] | None = None,
) -> Any:
    stage_three = _stage_three()
    specs = {spec.name: spec for spec in stage_three.MACALC_ISIN_LISTS}
    input_cells = {
        **_list_cells("H", specs["tbl_ForcedCompAperc"].title, ["Note 1", "F1"], "I"),
        **_list_cells("K", specs["tbl_MinCompAperc"].title, ["M1", "M2", "M3"], "L"),
        **_list_cells("N", "Partial Recognition restriction list (ISINs)", ["N1"]),
        **(inputs or {}),
    }
    parameter_cells = {
        **_list_cells("K", specs["lst_pre16assets"].title, ["P1", "P2"]),
        **_list_cells("M", specs["lst_REITs"].title, ["R1"]),
        **(parameters or {}),
    }
    names = {
        name: SimpleNamespace(refers_to=(formulas or {}).get(name, spec.refers_to))
        for name, spec in specs.items()
    }
    return SimpleNamespace(
        sheets={
            "I. Inputs": _GridSheet(input_cells),
            "P. Parameters": _GridSheet(parameter_cells),
        },
        names=names,
    )


def test_read_isin_lists_reads_each_list_from_row_7() -> None:
    stage_three = _stage_three()

    result = stage_three.read_isin_lists(_isin_book())

    by_name = {item.spec.name: item for item in result.lists}
    assert by_name["tbl_ForcedCompAperc"].keys == ("Note 1", "F1")
    assert by_name["tbl_ForcedCompAperc"].values == (0.5, 0.5)
    assert by_name["tbl_MinCompAperc"].keys == ("M1", "M2", "M3")
    assert by_name["lst_pre16assets"].keys == ("P1", "P2")
    assert by_name["lst_pre16assets"].values is None
    assert by_name["lst_REITs"].keys == ("R1",)
    assert by_name["tbl_MinCompAperc"].last_row == 9


def test_read_isin_lists_allows_an_empty_list() -> None:
    stage_three = _stage_three()
    book = _isin_book(parameters={"M7": None})

    result = stage_three.read_isin_lists(book)

    reits = {item.spec.name: item for item in result.lists}["lst_REITs"]
    assert reits.keys == () and reits.last_row == 6


@pytest.mark.parametrize(
    ("inputs", "parameters", "pattern"),
    [
        ({"H5": "Forced Comp A %"}, {}, r"title.*H5"),
        ({}, {"M5": None}, r"title.*M5"),
        ({"K6": "stray"}, {}, r"row 6.*K6"),
        ({"L6": 0.5}, {}, r"row 6.*L6"),
        ({"H2": "header"}, {}, r"COUNTA.*H2"),
    ],
)
def test_read_isin_lists_stops_on_title_or_header_row_drift(
    inputs: dict[str, object], parameters: dict[str, object], pattern: str
) -> None:
    stage_three = _stage_three()

    with pytest.raises(stage_three.StageThreeError, match=pattern):
        stage_three.read_isin_lists(_isin_book(inputs, parameters))


def test_read_isin_lists_stops_on_name_formula_drift() -> None:
    stage_three = _stage_three()
    book = _isin_book(
        formulas={"tbl_MinCompAperc": "='I. Inputs'!$K$7:$L$170"}
    )

    with pytest.raises(stage_three.StageThreeError, match="tbl_MinCompAperc"):
        stage_three.read_isin_lists(book)


def test_read_isin_lists_stops_on_a_missing_name() -> None:
    stage_three = _stage_three()
    book = _isin_book()
    del book.names["lst_REITs"]

    with pytest.raises(stage_three.StageThreeError, match="lst_REITs"):
        stage_three.read_isin_lists(book)


@pytest.mark.parametrize(
    ("inputs", "parameters", "pattern"),
    [
        ({"K11": "LATE"}, {}, r"K10"),
        ({}, {"K10": "LATE"}, r"K9"),
    ],
)
def test_read_isin_lists_stops_on_a_key_below_the_first_blank(
    inputs: dict[str, object], parameters: dict[str, object], pattern: str
) -> None:
    stage_three = _stage_three()

    with pytest.raises(stage_three.StageThreeError, match=pattern):
        stage_three.read_isin_lists(_isin_book(inputs, parameters))


@pytest.mark.parametrize(
    ("inputs", "parameters", "cell"),
    [
        ({"G8": "x"}, {}, "G8"),
        ({"J7": "x"}, {}, "J7"),
        ({"M9": 1.0}, {}, "M9"),
        ({}, {"J5": "x"}, "J5"),
        ({}, {"L8": "x"}, "L8"),
        ({"I12": 0.5}, {}, "I12"),
    ],
)
def test_read_isin_lists_stops_on_a_stray_value_beside_or_below_a_list(
    inputs: dict[str, object], parameters: dict[str, object], cell: str
) -> None:
    stage_three = _stage_three()

    with pytest.raises(stage_three.StageThreeError, match=rf"contained.*{cell}"):
        stage_three.read_isin_lists(_isin_book(inputs, parameters))


def test_read_isin_lists_reports_every_list_problem_in_one_message() -> None:
    stage_three = _stage_three()
    book = _isin_book({"G8": "x"}, {"K5": "Pre16"})

    with pytest.raises(stage_three.StageThreeError) as error:
        stage_three.read_isin_lists(book)

    assert "G8" in str(error.value) and "K5" in str(error.value)


def test_controller_logger_name_does_not_depend_on_how_it_is_launched() -> None:
    """Source mode runs ``-m bonds_lite_input`` as __main__; Stage 3 logs to its child."""
    import logging
    import runpy

    source = Path(__file__).resolve().parents[1] / "src" / "bonds_lite_input.py"
    namespace = runpy.run_path(str(source), run_name="launched_as_script")

    assert namespace["logger"].name == "bonds_lite_input"
    assert _stage_three().logger.parent is logging.getLogger("bonds_lite_input")


def test_stage_three_summary_reaches_the_controller_console(
    monkeypatch: pytest.MonkeyPatch, capsys: pytest.CaptureFixture[str]
) -> None:
    import bonds_lite_input as shared

    stage_three = _stage_three()
    level = shared.logger.level
    monkeypatch.setattr(shared.logger, "handlers", [])
    monkeypatch.setattr(shared.logger, "propagate", shared.logger.propagate)
    try:
        shared.configure_console_logging()
        alignment = _align(
            [_isin_list("lst_REITs", ["X"])], [_asset_key(5, "X_A", "X")]
        )
        stage_three.log_isin_alignment(alignment)
    finally:
        shared.logger.setLevel(level)

    assert "lst_REITs: kept 0, remapped 1, unmatched 0" in capsys.readouterr().out


def test_run_stage_three_slice_c_guard_failure_keeps_previous_output(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    stage_three = _stage_three()
    caller, output, writes, _ = _mock_stage_three_run(
        monkeypatch, stage_three, tmp_path, existing_output=True
    )
    snapshot = stage_three.TemplateSnapshot(
        stage_three.TemplateLayout(2, 2, 3, 3, 4),
        isin_lists=stage_three.TemplateIsinLists(
            (_isin_list("tbl_MinCompAperc", ["K", "K"], [0.5, 0.6]),)
        ),
    )
    monkeypatch.setattr(stage_three, "_read_template_snapshot", lambda path: snapshot)
    attempted: list[object] = []
    monkeypatch.setattr(
        stage_three,
        "write_macalc_output",
        lambda *args, **kwargs: attempted.append(args),
    )

    with pytest.raises(stage_three.StageThreeError, match="conflicting"):
        stage_three.run_stage_three(caller)

    assert attempted == []
    assert writes[0]["status"] == "Running"
    assert writes[-1]["status"] == "Failed"
    assert writes[-1]["last_output"] == "Bonds Lite output/previous"
    assert output.with_name(EARLIER_OUTPUT_NAME).read_bytes() == b"previous"
    assert [path.name for path in output.parent.iterdir()] == [EARLIER_OUTPUT_NAME]


def test_run_stage_three_passes_aligned_lists_to_the_single_write_and_verify(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    caplog: pytest.LogCaptureFixture,
) -> None:
    import logging

    stage_three = _stage_three()
    caller, output, writes, _ = _mock_stage_three_run(
        monkeypatch, stage_three, tmp_path, existing_output=False
    )
    source = stage_three.TemplateIsinLists(
        (
            _isin_list("tbl_ForcedCompAperc", ["Note 1"], [1.0]),
            _isin_list("lst_REITs", ["X", "STALE"]),
        )
    )
    snapshot = stage_three.TemplateSnapshot(
        stage_three.TemplateLayout(2, 2, 3, 3, 4),
        fixed_keys=(_asset_key(2, "Note 1", block="template", rating="NR"),),
        isin_lists=source,
    )
    asset_row = [1.0, 0.0, 0.0, "X_A", "X", 0.0, "Bond", "S", "A-"]
    monkeypatch.setattr(stage_three, "_read_template_snapshot", lambda path: snapshot)
    monkeypatch.setattr(
        stage_three,
        "read_map_base_output",
        lambda path: SimpleNamespace(rows=(SimpleNamespace(final_rating="A"),)),
    )
    monkeypatch.setattr(stage_three, "build_asset_rows", lambda *args: [asset_row])
    calls: dict[str, Any] = {}

    def write_output(template, staged, layout, rows, assets=(), isin_lists=()):
        calls.setdefault("writes", []).append(isin_lists)
        staged.write_bytes(b"new")
        return 20

    def verify(path, expected, **kwargs):
        calls.setdefault("verifies", []).append(kwargs)

    monkeypatch.setattr(stage_three, "write_macalc_output", write_output)
    monkeypatch.setattr(stage_three, "verify_macalc_output", verify)

    with caplog.at_level(logging.INFO, logger=stage_three.logger.name):
        monkeypatch.setattr(stage_three.shared.logger, "propagate", True)
        assert stage_three.run_stage_three(caller) == output

    assert len(calls["writes"]) == len(calls["verifies"]) == 1
    aligned = calls["writes"][0]
    assert [item.keys for item in aligned] == [("Note 1",), ("X_A", "STALE")]
    verified = calls["verifies"][0]
    assert verified["isin_lists"] == aligned
    assert verified["template_isin_lists"] == source
    assert verified["funded_count"] == 0
    assert verified["asset_ratings"] == ("A",)
    assert "lst_REITs: kept 0, remapped 1, unmatched 1, added 0" in caplog.text
    assert writes[-1]["status"] == "Succeeded"


@pytest.mark.excel
@pytest.mark.skipif(
    os.environ.get("BONDS_LITE_RUN_STAGE_THREE_EXCEL_TESTS") != "1",
    reason="Set BONDS_LITE_RUN_STAGE_THREE_EXCEL_TESTS=1",
)
def test_excel_slice_c_writes_resizes_and_verifies_lists_with_a_fan_out(
    tmp_path: Path,
) -> None:
    """Synthetic assets on a disposable template copy: remap, fan-out and a negative."""
    import xlwings as xw

    stage_three = _stage_three()
    template = Path(__file__).resolve().parents[1] / "MACalcinPython_v4.0f.xlsb"
    template_hash = hashlib.sha256(template.read_bytes()).digest()
    snapshot = stage_three._read_template_snapshot(template)
    sources = {item.spec.name: item for item in snapshot.isin_lists.lists}
    assert {name: len(item.keys) for name, item in sources.items()} == {
        "tbl_ForcedCompAperc": 24,
        "tbl_MinCompAperc": 164,
        "lst_pre16assets": 322,
        "lst_REITs": 16,
    }
    min_key = sources["tbl_MinCompAperc"].keys[0]
    reit_key = sources["lst_REITs"].keys[0]

    def asset(d: str, e: str, rating: str) -> list[object]:
        return [1.0, 0.0, 0.0, d, e, 100.0, "Bond", "Corporate", rating,
                "FXD", "GBP", 0.0, 1.0, *([0.0] * MONTH_COUNT)]

    asset_rows = [
        asset(f"{min_key}_A", min_key, "A"),
        asset(f"{reit_key}_BBB", reit_key, "BBB"),
        asset(f"{reit_key}_A", reit_key, "A"),
    ]
    asset_ratings = [row[8] for row in asset_rows]
    keys = stage_three.build_asset_keys(
        snapshot.layout, snapshot.fixed_keys, [], asset_rows,
        asset_ratings=asset_ratings,
    )
    alignment = stage_three.align_isin_lists(snapshot.isin_lists, keys)
    reits = alignment.get("lst_REITs")
    assert reits.keys[:2] == (f"{reit_key}_BBB", f"{reit_key}_A")
    assert len(reits.keys) == 17 and reits.added == 1
    assert alignment.get("tbl_MinCompAperc").keys[0] == f"{min_key}_A"
    output = tmp_path / "MACalc output.xlsb"

    cleared_through = stage_three.write_macalc_output(
        template, output, snapshot.layout, [], asset_rows,
        isin_lists=alignment.lists,
    )
    verify_arguments = {
        "template_path": template,
        "layout": snapshot.layout,
        "cleared_through": cleared_through,
        "isin_lists": alignment.lists,
        "template_isin_lists": snapshot.isin_lists,
        "funded_count": 0,
        "asset_ratings": asset_ratings,
    }
    stage_three.verify_macalc_output(output, asset_rows, **verify_arguments)

    application = xw.App(visible=False, add_book=False)
    try:
        application.display_alerts = False
        book = application.books.open(
            str(output), update_links=False, ignore_read_only_recommended=True
        )
        try:
            assert not book.api.ReadOnly
            resolved = book.names["lst_REITs"].refers_to_range
            assert (resolved.row, resolved.rows.count) == (7, 17)
            book.sheets["P. Parameters"].range("M8").value = "CORRUPTED"
            book.save()
        finally:
            book.close()
    finally:
        application.api.Quit()

    with pytest.raises(stage_three.StageThreeError, match="lst_REITs"):
        stage_three.verify_macalc_output(output, asset_rows, **verify_arguments)
    assert hashlib.sha256(template.read_bytes()).digest() == template_hash


TEMPLATE_LINKS = (
    "https://example.sharepoint.com/Inputs/Curves.xlsm",
    "https://example.sharepoint.com/MA/MACE.xlsb",
    r"\\server\share\Output\py_bonds_Base_Scenario_nozeros.csv",
    r"\\server\share\Regression\MACalcinPython_v4.0f.xlsb",
)


@pytest.mark.parametrize(
    ("current", "expected"),
    [
        (TEMPLATE_LINKS, TEMPLATE_LINKS),
        (TEMPLATE_LINKS[:2], TEMPLATE_LINKS[:2]),
        ((TEMPLATE_LINKS[0], TEMPLATE_LINKS[2]), (TEMPLATE_LINKS[0], TEMPLATE_LINKS[2])),
        ((), ()),
        ((*TEMPLATE_LINKS, r"\\server\share\New.xlsb"), TEMPLATE_LINKS),
        ((TEMPLATE_LINKS[1], TEMPLATE_LINKS[0]), TEMPLATE_LINKS),
        (("https://example.sharepoint.com/Inputs/Other.xlsm",), TEMPLATE_LINKS),
        (
            (r"\\server\share\Inputs\Curves.xlsm", r"\\server\share\MA\MACE.xlsb"),
            TEMPLATE_LINKS[:2],
        ),
    ],
    ids=["unchanged", "replaced-rows-only-links-dropped", "middle-link-dropped",
         "all-dropped", "added-link", "reordered", "unknown-file",
         "excel-repathed-url-to-unc"],
)
def test_retained_template_links_allows_only_dropped_links(
    current: tuple[str, ...], expected: tuple[str, ...]
) -> None:
    """Links used only by replaced asset rows may vanish; anything else is checked in full."""
    from sst_stage_three import _retained_template_links

    assert _retained_template_links(TEMPLATE_LINKS, current) == expected


def test_restore_accepts_retained_links_after_template_link_dropped() -> None:
    """Regression: a template link used only by replaced rows no longer fails the save."""
    from sst_stage_three import (
        StageThreeError,
        _restore_external_link_sources,
        _retained_template_links,
    )

    saved = [r"\\server\share\Inputs\Curves.xlsm", TEMPLATE_LINKS[1]]
    changes: list[tuple[str, str]] = []

    def change_link(current: str, original: str, kind: int) -> None:
        changes.append((current, original))
        saved[saved.index(current)] = original

    book = SimpleNamespace(
        api=SimpleNamespace(
            LinkSources=lambda kind: tuple(saved), ChangeLink=change_link
        )
    )
    with pytest.raises(StageThreeError, match="different number of external links"):
        _restore_external_link_sources(book, TEMPLATE_LINKS)
    expected = _retained_template_links(TEMPLATE_LINKS, saved)
    assert _restore_external_link_sources(book, expected) is True
    assert changes == [(r"\\server\share\Inputs\Curves.xlsm", TEMPLATE_LINKS[0])]
    assert tuple(saved) == TEMPLATE_LINKS[:2]
