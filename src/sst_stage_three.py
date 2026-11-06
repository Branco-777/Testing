"""Prepare and verify the funded-leg and asset blocks in a MACalc input workbook."""

from __future__ import annotations

import csv
import ctypes
import getpass
import logging
import math
import os
import re
import secrets
import shutil
import stat
import time
from collections.abc import Iterable, Mapping, Sequence
from contextlib import contextmanager, suppress
from dataclasses import dataclass
from datetime import date, datetime
from numbers import Real
from pathlib import Path
from typing import Any, Iterator

import bonds_lite_input as shared

# A child of the controller logger, so Stage 3 lines reach its console handler.
logger = shared.logger.getChild("sst_stage_three")

MONTH_COUNT = 1200
# A:M descriptive columns followed by N:ATQ months.
OUTPUT_WIDTH = 13 + MONTH_COUNT
MAP_HEADER_ROW = 8
MACE_HEADER_ROW = 9
MACE_FIRST_DATA_ROW = 10
MACALC_SHEET = "I. Assets"
MAP_SHEET = "I. Bonds Data"
MACE_SHEET = "C. Assets - Funding"
OUTPUT_STEM = "MACalc input"
OUTPUT_EXTENSION = ".xlsb"
# Excel cannot open or save a workbook whose full path exceeds 218 characters.
MAX_OUTPUT_PATH = 218
TOKEN_PATTERN = re.compile(r"[0-9a-f]{16}")
MAX_TOKEN_ATTEMPTS = 10
# Excel rewrites the template's SharePoint links when saving outside the synced
# folder, so staging stays in SST Output_<suffix> but outside every stress folder.
STAGING_FOLDER = "Stage 3 staging"
ASSETS_TABLE_NAME = "tbl_Assets"
# Change here if a Bonds Lite model update renames the Stage 2 base-scenario CSV.
MAP_BASE_OUTPUT_PATTERN = "py_bonds_Base_Scenario_nozeros"
MAP_BASE_CSV_HEADERS = {
    "A": "ISIN",
    "B": "Unique_ISIN",
    "D": "IM_Zspreads",
    "E": "Currency",
    "F": "Sector",
    "H": "Final_Rating",
    "I": "Final_Notched_Rating",
    "J": "Base_PV",
}
ZSPREAD_TO_BASIS_POINTS = 10000
MAX_LISTED_PROBLEMS = 50
MACE_BONDS_SHEET = "I. Bonds"
MACE_BONDS_HEADER_ROW = 7
MACE_BONDS_HEADERS = {"D": "ISIN", "G": "Type", "J": "Index Linked?", "L": "FSAO"}
MACE_CONTROL_SHEET = "Control"
CONTROL_HEADER_ROW = 9
CONTROL_HEADERS = ("ISIN", "Comp. A %", "Comp. B %", "Comp. C %", "Partial CF Flag")
CONTROL_LOOKAHEAD_ROWS = 50
# Column beside the K:O table that must be empty (containment rule I3, Decision 48).
# Columns Q onwards hold other MACE panels and are not checked.
CONTROL_SIDE_COLUMNS = ("P", "P")
LAYOUT_WARNING = (
    "The MACalc template must start with Mortgage assets followed by ILS assets. "
    "Please update the template to account for LTM/Swap assets and rerun Stage 3."
)
STAGE_THREE_FIELD_NAMES = {
    "last_run": "SST_Stage3_LastRun",
    "start": "SST_Stage3_Start",
    "finish": "SST_Stage3_Finish",
    "elapsed": "SST_Stage3_Elapsed",
    "user": "SST_Stage3_User",
    "status": "SST_Stage3_Status",
    "input_run": "SST_Stage3_InputRun",
    "last_output": "SST_Stage3_LastSuccessfulOutput",
}
STAGE_THREE_PANEL_NAME = "SST_Stage3_Panel"
STAGE_THREE_BUTTON_NAME = "btnPrepareMACalc"
STAGE_THREE_BUTTON_MACRO = "PrepareMACalc"
STAGE_THREE_LABELS = {
    "last_run": "Last Run",
    "start": "Start:",
    "finish": "Finish:",
    "elapsed": "Time",
    "user": "User",
    "status": "Status",
    "input_run": "Input Run",
    "last_output": "Last Successful Output",
}
STAGE_THREE_OFFSETS = {
    "last_run": 1,
    "start": 2,
    "finish": 3,
    "elapsed": 4,
    "user": 5,
    "status": 6,
    "input_run": 7,
    "last_output": 9,
}


class StageThreeError(shared.BondsLiteError):
    """An actionable Stage 3 validation or workbook-processing failure."""


class TemplateLayoutError(StageThreeError):
    """The MACalc template does not have the required Mortgage/ILS blocks."""


@dataclass(frozen=True)
class TemplateLayout:
    mortgage_start_row: int
    mortgage_end_row: int
    ils_start_row: int
    ils_end_row: int
    funded_start_row: int


@dataclass(frozen=True)
class MapBonds:
    by_isin: dict[str, dict[str, object]]


@dataclass(frozen=True)
class MaceFundingTable:
    rows: tuple[dict[str, object], ...]


@dataclass(frozen=True)
class MaceBondAttributes:
    by_isin: dict[str, dict[str, object]]
    duplicates: frozenset[str]


@dataclass(frozen=True)
class ControlTable:
    by_isin: dict[str, dict[str, object]]
    duplicates: frozenset[str]
    last_row: int


@dataclass(frozen=True)
class MaceAllocations:
    by_isin: dict[str, dict[str, object]]
    duplicates: frozenset[str]
    valuation_date: object


@dataclass(frozen=True)
class MaceInputs:
    funding: MaceFundingTable
    bonds: MaceBondAttributes
    allocations: MaceAllocations


@dataclass(frozen=True)
class MapBaseRow:
    isin: str
    unique_isin: str
    zspread: float
    currency: str
    sector: str
    rating: str
    final_rating: str
    pv_and_cashflows: tuple[float, ...]


@dataclass(frozen=True)
class MapBaseOutput:
    path: Path
    rows: tuple[MapBaseRow, ...]


@dataclass(frozen=True)
class StageThreeInputs:
    caller_path: Path
    run_folder: Path
    input_run: str
    map_path: Path
    template_path: Path
    mace_path: Path
    map_base_output_path: Path | None = None
    valuation_date: object = None
    bonds_path: Path | None = None


LIST_TITLE_ROW = 5
LIST_FIRST_ROW = 7


@dataclass(frozen=True)
class IsinListSpec:
    """One MACalc list keyed on I. Assets column D (Slice C, Decisions 49-52)."""

    name: str
    sheet: str
    key_column: str
    value_column: str | None
    title: str
    neighbour_columns: tuple[str, ...]

    @property
    def is_table(self) -> bool:
        return self.value_column is not None

    @property
    def last_column(self) -> str:
        return self.value_column or self.key_column

    @property
    def refers_to(self) -> str:
        sheet = f"'{self.sheet}'"
        column = self.key_column
        width = 2 if self.is_table else 1
        return (
            f"=OFFSET({sheet}!${column}${LIST_FIRST_ROW},0,0,"
            f"COUNTA({sheet}!${column}:${column})-1,{width})"
        )


# lst_noPRoC is deliberately absent: it holds ILS D values, which Stage 3 never
# changes, and the read-back proves it unchanged (Decision 50).
MACALC_ISIN_LISTS = (
    IsinListSpec(
        "tbl_ForcedCompAperc",
        "I. Inputs",
        "H",
        "I",
        "Forced Component A allocation percentages",
        ("G", "J"),
    ),
    IsinListSpec(
        "tbl_MinCompAperc",
        "I. Inputs",
        "K",
        "L",
        "Minimum Component A allocation percentages",
        ("J", "M"),
    ),
    IsinListSpec(
        "lst_pre16assets", "P. Parameters", "K", None, "Pre-16 asset ISINs", ("J", "L")
    ),
    IsinListSpec(
        "lst_REITs",
        "P. Parameters",
        "M",
        None,
        "REITs (classified as Fin for SCR)",
        ("L",),
    ),
)


@dataclass(frozen=True)
class IsinList:
    spec: IsinListSpec
    keys: tuple[object, ...]
    values: tuple[object, ...] | None = None
    # Last used row of the key/value columns in the template; the tail to clear.
    last_row: int = LIST_FIRST_ROW - 1


@dataclass(frozen=True)
class TemplateIsinLists:
    lists: tuple[IsinList, ...] = ()


@dataclass(frozen=True)
class AssetKey:
    """D, E and rating of one I. Assets row; block is template, funded or asset.

    Asset-block keys carry the CSV ``Final_Rating`` that suffixes ``Unique_ISIN``;
    template and funded keys carry column I.
    """

    row: int
    d: object
    e: object
    block: str
    rating: object = None


@dataclass(frozen=True)
class AlignedIsinList:
    spec: IsinListSpec
    keys: tuple[str, ...]
    values: tuple[object, ...] | None
    template_last_row: int
    source_size: int
    kept: int
    remapped: int
    unmatched: tuple[str, ...]
    added: int
    fan_outs: tuple[tuple[str, tuple[str, ...]], ...]
    duplicates_removed: int
    matched_keys: int

    @property
    def mapped_size(self) -> int:
        return self.source_size + self.added

    def summary(self) -> str:
        return (
            f"{self.spec.name}: kept {self.kept}, remapped {self.remapped}, "
            f"unmatched {len(self.unmatched)}, added {self.added}, "
            f"duplicates removed {self.duplicates_removed}"
        )


@dataclass(frozen=True)
class IsinListAlignment:
    lists: tuple[AlignedIsinList, ...]

    def get(self, name: str) -> AlignedIsinList:
        for item in self.lists:
            if item.spec.name == name:
                return item
        raise KeyError(name)


@dataclass(frozen=True)
class TemplateSnapshot:
    layout: TemplateLayout
    fixed_keys: tuple[AssetKey, ...] = ()
    isin_lists: TemplateIsinLists = TemplateIsinLists()


def _column_number(label: str) -> int:
    result = 0
    for character in label:
        result = result * 26 + ord(character.upper()) - 64
    return result


def _normalise_header(value: object) -> str:
    return value.strip().casefold() if isinstance(value, str) else ""


def _cell_error(value: object) -> bool:
    return isinstance(value, str) and value.startswith("#")


def _numeric(value: object, description: str) -> Real:
    try:
        finite = math.isfinite(float(value))
    except (OverflowError, TypeError, ValueError):
        finite = False
    if isinstance(value, bool) or not isinstance(value, Real) or not finite:
        raise StageThreeError(
            f"{description} must be a finite numeric value; got {value!r}"
        )
    return value


def _get(record: object, key: str) -> object:
    if isinstance(record, Mapping):
        return record[key]
    return getattr(record, key)


def validate_template_blocks(types: Sequence[object]) -> TemplateLayout:
    """Require contiguous Mortgage then ILS blocks at the top of the asset rows."""
    values = list(types)
    while values and (
        values[-1] is None or isinstance(values[-1], str) and not values[-1].strip()
    ):
        values.pop()
    normalised = [
        value.strip().casefold() if isinstance(value, str) else "" for value in values
    ]
    if not normalised or normalised[0] != "mortgage":
        raise TemplateLayoutError("Mortgage assets do not start at row 2")

    mortgage_end = 0
    while mortgage_end < len(normalised) and normalised[mortgage_end] == "mortgage":
        mortgage_end += 1
    if mortgage_end == len(normalised) or normalised[mortgage_end] != "ils":
        raise TemplateLayoutError(
            f"ILS must immediately follow the Mortgage block at row {mortgage_end + 2}"
        )

    ils_start = mortgage_end
    ils_end = ils_start
    while ils_end < len(normalised) and normalised[ils_end] == "ils":
        ils_end += 1
    if any(value == "ils" for value in normalised[ils_end:]):
        raise TemplateLayoutError(
            f"ILS block contains a gap or out-of-order row at row {ils_end + 2}"
        )
    if any(value == "mortgage" for value in normalised[ils_end:]):
        row = normalised.index("mortgage", ils_end) + 2
        raise TemplateLayoutError(f"Mortgage asset appears after ILS at row {row}")

    mortgage_start_row = 2
    mortgage_end_row = mortgage_end + 1
    ils_start_row = ils_start + 2
    ils_end_row = ils_end + 1
    return TemplateLayout(
        mortgage_start_row,
        mortgage_end_row,
        ils_start_row,
        ils_end_row,
        ils_end_row + 1,
    )


def validate_template_headers(headers: Sequence[object]) -> None:
    """Enforce the MACalc template's fixed descriptive and monthly column contract."""
    expected = {
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
    }
    errors = []
    for column, header in expected.items():
        index = _column_number(column) - 1
        if (
            index >= len(headers)
            or _normalise_header(headers[index]) != header.casefold()
        ):
            errors.append(f"{column}1 must be headed {header!r}")
    first_month = _column_number("N") - 1
    last_month = _column_number("ATQ") - 1
    if len(headers) <= last_month:
        errors.append("Month headers must extend through ATQ")
    else:
        months = headers[first_month : last_month + 1]
        if len(months) != MONTH_COUNT or any(
            isinstance(value, bool) or value != index + 1
            for index, value in enumerate(months)
        ):
            errors.append("Month headers N:ATQ must be the ordinals 1 through 1200")
    if errors:
        raise StageThreeError("MACalc template header mismatch: " + "; ".join(errors))


def identify_funded(table: MaceFundingTable | Sequence[object]) -> list[object]:
    """Classify strict negative-only MACE funding vectors and validate identities."""
    rows = table.rows if isinstance(table, MaceFundingTable) else tuple(table)
    seen: set[str] = set()
    funded: list[object] = []
    for row_number, row in enumerate(rows, start=MACE_FIRST_DATA_ROW):
        isin = _get(row, "isin")
        if not isinstance(isin, str) or not isin.strip():
            raise StageThreeError(f"MACE ISIN is blank or invalid in row {row_number}")
        if isin in seen:
            raise StageThreeError(f"Duplicate MACE ISIN {isin!r} in row {row_number}")
        seen.add(isin)
        cashflows = _get(row, "funding_cashflows")
        if not isinstance(cashflows, Sequence) or isinstance(cashflows, str):
            raise StageThreeError(f"MACE funding cashflows are invalid for {isin}")
        if len(cashflows) != MONTH_COUNT:
            raise StageThreeError(
                f"MACE funding cashflows for {isin} must contain {MONTH_COUNT} months"
            )
        values = [
            _numeric(value, f"MACE funding cashflow for {isin}, month {month}")
            for month, value in enumerate(cashflows, start=1)
        ]
        has_negative = any(value < 0 for value in values)
        has_positive = any(value > 0 for value in values)
        if has_positive:
            raise StageThreeError(
                f"Anomalous MACE funding cashflows for {isin}: positive month found"
            )
        if not has_negative:
            continue

        for key, label in (
            ("zspread", "Z-Spread"),
            ("fsao", "FS Add On"),
            ("ai", "AI/PV"),
        ):
            _numeric(_get(row, key), f"MACE {label} for funded ISIN {isin}")
        asset_type = _get(row, "type")
        if not isinstance(asset_type, str) or not asset_type.strip():
            raise StageThreeError(
                f"MACE Type is blank or invalid for funded ISIN {isin}"
            )
        funded.append(row)
    return funded


def reconcile_funded(
    map_bonds: MapBonds | Mapping[str, object],
    funded: Sequence[object],
    bonds_negative_isins: Iterable[str],
) -> bool:
    """Require Bonds-input/MACE funded-set equality; True when both sets are empty.

    Stage 1 removes negative cashflows from the MAP input, so the funded set is
    compared with the negative-cashflow ISINs of the Bonds input instead. MAP is
    still required to hold every funded ISIN because it supplies descriptions.
    """
    bonds_negative = set(bonds_negative_isins)
    map_by_isin = _get(map_bonds, "by_isin")
    mace_funded = {str(_get(row, "isin")) for row in funded}
    missing = mace_funded - set(map_by_isin)
    only_bonds = bonds_negative - mace_funded
    only_mace = mace_funded - bonds_negative
    if missing or only_bonds or only_mace:
        details = []
        if only_bonds:
            details.append("Bonds input only: " + ", ".join(sorted(only_bonds)))
        if only_mace:
            details.append("MACE only: " + ", ".join(sorted(only_mace)))
        if missing:
            details.append(
                "funded ISINs missing from MAP: " + ", ".join(sorted(missing))
            )
        raise StageThreeError(
            "Bonds input/MACE funded set mismatch (" + "; ".join(details) + ")"
        )
    return not bonds_negative and not mace_funded


def _allocation_values(allocations: object, isin: object) -> list[object]:
    try:
        allocation = _get(allocations, "by_isin")[isin]
    except KeyError as error:
        raise StageThreeError(
            f"ISIN {isin!r} is missing from MACE Control column K"
        ) from error
    return [allocation["a"], allocation["b"], allocation["c"]]


def build_funded_rows(
    map_bonds: MapBonds | Mapping[str, object],
    funded: Sequence[object],
    allocations: MaceAllocations,
) -> list[list[object]]:
    """Build A:ATQ rows: Control A:C, MAP descriptions and MACE funding values."""
    map_by_isin = _get(map_bonds, "by_isin")
    result = []
    for mace in funded:
        isin = _get(mace, "isin")
        try:
            map_row = map_by_isin[isin]
        except KeyError as error:
            raise StageThreeError(
                f"Funded ISIN {isin!r} is missing from MAP"
            ) from error
        allocation = _allocation_values(allocations, isin)
        cashflows = list(_get(mace, "funding_cashflows"))
        if len(cashflows) != MONTH_COUNT:
            raise StageThreeError(
                f"MACE funding cashflows for {isin} must contain 1200 months"
            )
        row = [
            *allocation,
            isin,
            isin,
            _get(mace, "zspread"),
            _get(mace, "type"),
            _get(map_row, "sector"),
            _get(map_row, "rating"),
            _get(map_row, "index_linked"),
            _get(map_row, "currency"),
            _get(mace, "fsao"),
            -_numeric(_get(mace, "ai"), f"MACE AI for funded ISIN {isin}"),
            *cashflows,
        ]
        if len(row) != OUTPUT_WIDTH:
            raise StageThreeError(f"Built MACalc row for {isin} has an invalid width")
        result.append(row)
    return result


@dataclass(frozen=True)
class SliceBReconciliation:
    csv_only: tuple[str, ...] = ()
    map_only: tuple[str, ...] = ()
    mace_only_bonds: int = 0
    mace_only_allocations: int = 0

    def map_difference_warning(self) -> str | None:
        """Describe CSV/MAP ISIN differences, which warn but never stop the run."""
        if not self.csv_only and not self.map_only:
            return None
        parts = []
        if self.csv_only:
            parts.append(
                f"In the base-scenario CSV but not MAP ({len(self.csv_only)}): "
                + _listed(list(self.csv_only))
            )
        if self.map_only:
            parts.append(
                f"In MAP but not the base-scenario CSV ({len(self.map_only)}): "
                + _listed(list(self.map_only))
            )
        return (
            "Stage 3 completed, but the base-scenario CSV and MAP ISINs differ.\n"
            + "\n".join(parts)
        )


def _as_date(value: object) -> date | None:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    return None


def reconcile_slice_b(
    base: MapBaseOutput,
    mace_bonds: MaceBondAttributes,
    allocations: MaceAllocations,
    map_bonds: MapBonds | Mapping[str, object],
    funded: Sequence[object],
    valuation_date: object,
) -> SliceBReconciliation:
    """Check the CSV and funded ISINs against MACE, collecting every problem."""
    csv_isins = list(dict.fromkeys(row.isin for row in base.rows))
    funded_isins = list(dict.fromkeys(str(_get(row, "isin")) for row in funded))
    control_isins = list(dict.fromkeys([*csv_isins, *funded_isins]))
    problems: dict[str, list[str]] = {}

    def add(category: str, isin: str) -> None:
        problems.setdefault(category, []).append(isin)

    for isin in csv_isins:
        if isin in mace_bonds.duplicates:
            add("ISINs duplicated in MACE I. Bonds column D", isin)
        attributes = mace_bonds.by_isin.get(isin)
        if attributes is None:
            add("ISINs missing from MACE I. Bonds column D", isin)
            continue
        if not _valid_text(attributes["type"]):
            add("Invalid Type in MACE I. Bonds", isin)
        if not _valid_text(attributes["index_linked"]):
            add("Invalid Index Linked? in MACE I. Bonds", isin)
        if not _valid_number(attributes["fsao"]):
            add("Invalid FSAO in MACE I. Bonds", isin)
    for isin in control_isins:
        if isin in allocations.duplicates:
            add("ISINs duplicated in MACE Control column K", isin)
        allocation = allocations.by_isin.get(isin)
        if allocation is None:
            add("ISINs missing from MACE Control column K", isin)
            continue
        for key, label in (
            ("a", "Comp. A %"),
            ("b", "Comp. B %"),
            ("c", "Partial CF Flag"),
        ):
            if not _valid_number(allocation[key]):
                add(f"Invalid {label} in MACE Control", isin)

    draft_date = _as_date(valuation_date)
    control_date = _as_date(allocations.valuation_date)
    if draft_date is None or control_date is None or draft_date != control_date:
        problems["Valuation date mismatch"] = [
            f"Draft I_Valuation_Date {valuation_date!r} vs MACE Control "
            f"C{CONTROL_HEADER_ROW} {allocations.valuation_date!r}"
        ]
    _raise_problems("Stage 3 Slice B reconciliation failed", problems)

    known = {*csv_isins, *funded_isins}
    map_isins = set(_get(map_bonds, "by_isin"))
    return SliceBReconciliation(
        csv_only=tuple(sorted(set(csv_isins) - map_isins)),
        map_only=tuple(sorted(map_isins - set(csv_isins))),
        mace_only_bonds=len(set(mace_bonds.by_isin) - known),
        mace_only_allocations=len(set(allocations.by_isin) - known),
    )


def build_asset_rows(
    base: MapBaseOutput,
    mace_bonds: MaceBondAttributes,
    allocations: MaceAllocations,
) -> list[list[object]]:
    """Build one A:ATQ row per CSV row, in CSV order."""
    result = []
    for csv_row in base.rows:
        try:
            attributes = mace_bonds.by_isin[csv_row.isin]
        except KeyError as error:
            raise StageThreeError(
                f"ISIN {csv_row.isin!r} is missing from MACE I. Bonds column D"
            ) from error
        row = [
            *_allocation_values(allocations, csv_row.isin),
            csv_row.unique_isin,
            csv_row.isin,
            csv_row.zspread * ZSPREAD_TO_BASIS_POINTS,
            attributes["type"],
            csv_row.sector,
            csv_row.rating,
            attributes["index_linked"],
            csv_row.currency,
            attributes["fsao"],
            *csv_row.pv_and_cashflows,
        ]
        if len(row) != OUTPUT_WIDTH:
            raise StageThreeError(
                f"Built MACalc asset row for {csv_row.unique_isin} has an invalid width"
            )
        result.append(row)
    return result


def output_blocks(
    layout: TemplateLayout,
    funded_rows: Sequence[Sequence[object]],
    asset_rows: Sequence[Sequence[object]],
) -> list[tuple[int, Sequence[Sequence[object]]]]:
    """Place funded rows after ILS and asset rows directly below them."""
    start = layout.funded_start_row
    return [(start, funded_rows), (start + len(funded_rows), asset_rows)]


def build_asset_keys(
    layout: TemplateLayout,
    fixed_keys: Sequence[AssetKey],
    funded_rows: Sequence[Sequence[object]],
    asset_rows: Sequence[Sequence[object]],
    *,
    asset_ratings: Sequence[object],
) -> tuple[AssetKey, ...]:
    """Key every I. Assets row on its D, E and rating, at its output row number.

    ``asset_ratings`` holds each asset row's CSV ``Final_Rating``, in CSV order.
    """
    if len(asset_ratings) != len(asset_rows):
        raise StageThreeError(
            f"Stage 3 has {len(asset_ratings)} CSV Final_Rating values for "
            f"{len(asset_rows)} asset rows"
        )
    keys = list(fixed_keys)
    (funded_start, _), (asset_start, _) = output_blocks(
        layout, funded_rows, asset_rows
    )
    keys.extend(
        AssetKey(funded_start + offset, row[3], row[4], "funded", row[8])
        for offset, row in enumerate(funded_rows)
    )
    keys.extend(
        AssetKey(asset_start + offset, row[3], row[4], "asset", rating)
        for offset, (row, rating) in enumerate(zip(asset_rows, asset_ratings))
    )
    return tuple(keys)


def _check_list_entries(item: IsinList, problems: dict[str, list[str]]) -> None:
    spec = item.spec
    for offset, key in enumerate(item.keys):
        if not isinstance(key, str) or not _valid_text(key):
            problems.setdefault(
                f"{spec.name}: blank, error or non-text keys", []
            ).append(f"{spec.key_column}{LIST_FIRST_ROW + offset} ({key!r})")
    if not spec.is_table:
        return
    values = item.values or ()
    for offset, key in enumerate(item.keys):
        value = values[offset] if offset < len(values) else None
        if not (_valid_number(value) and 0 <= value <= 1):  # type: ignore[operator]
            problems.setdefault(
                f"{spec.name}: blank, error, non-numeric or out-of-[0,1] values", []
            ).append(
                f"{spec.value_column}{LIST_FIRST_ROW + offset} for {key!r} "
                f"({value!r})"
            )


def _align_one(
    item: IsinList,
    by_d: Mapping[object, Sequence[AssetKey]],
    by_e: Mapping[object, Sequence[AssetKey]],
    conflicts: dict[str, list[str]],
) -> AlignedIsinList:
    spec = item.spec
    values = item.values if spec.is_table else None
    emitted: list[tuple[str, object]] = []
    kept = remapped = added = 0
    unmatched: list[str] = []
    fan_outs: list[tuple[str, tuple[str, ...]]] = []
    for offset, key in enumerate(item.keys):
        value = values[offset] if values is not None else None
        rows = {asset.row: asset for asset in by_d.get(key, ())}
        rows.update({asset.row: asset for asset in by_e.get(key, ())})
        matched = tuple(str(rows[row].d) for row in sorted(rows))
        if not matched:
            unmatched.append(str(key))
            emitted.append((str(key), value))
            continue
        if matched == (key,):
            kept += 1
        else:
            remapped += 1
        if len(matched) > 1:
            added += len(matched) - 1
            fan_outs.append((str(key), matched))
        emitted.extend((match, value) for match in matched)

    # Decision 52: keep the first occurrence; conflicting table values stop the run.
    seen: dict[str, list[object]] = {}
    final: list[tuple[str, object]] = []
    for key, value in emitted:
        if key not in seen:
            seen[key] = [value]
            final.append((key, value))
        else:
            seen[key].append(value)
    if spec.is_table:
        for key, found in seen.items():
            if any(value != found[0] for value in found[1:]):
                conflicts.setdefault(
                    f"{spec.name}: duplicate keys with conflicting values", []
                ).append(f"{key!r} (" + ", ".join(repr(v) for v in found) + ")")
    d_values = set(by_d)
    return AlignedIsinList(
        spec=spec,
        keys=tuple(key for key, _ in final),
        values=tuple(value for _, value in final) if spec.is_table else None,
        template_last_row=item.last_row,
        source_size=len(item.keys),
        kept=kept,
        remapped=remapped,
        unmatched=tuple(unmatched),
        added=added,
        fan_outs=tuple(fan_outs),
        duplicates_removed=len(emitted) - len(final),
        matched_keys=sum(key in d_values for key, _ in final),
    )


def align_isin_lists(
    lists: TemplateIsinLists | Sequence[IsinList], asset_keys: Sequence[AssetKey]
) -> IsinListAlignment:
    """Re-key each MACalc list on the I. Assets D values its entries matched.

    An entry ``k`` matches every row whose D is ``k`` and every rating-suffixed
    asset row (D exactly E + "_" + its own CSV Final_Rating) whose E is ``k``; it
    is replaced by those rows' D values in row order, keeping its value.
    Unmatched entries stay put.
    """
    source = lists.lists if isinstance(lists, TemplateIsinLists) else tuple(lists)
    problems: dict[str, list[str]] = {}
    by_d: dict[object, list[AssetKey]] = {}
    by_e: dict[object, list[AssetKey]] = {}
    for asset in sorted(asset_keys, key=lambda key: key.row):
        by_d.setdefault(asset.d, []).append(asset)
        if asset.block != "asset" or asset.d == asset.e:
            continue
        if (
            isinstance(asset.e, str)
            and isinstance(asset.rating, str)
            and asset.d == f"{asset.e}_{asset.rating}"
        ):
            by_e.setdefault(asset.e, []).append(asset)
        else:
            problems.setdefault(
                "I. Assets rows whose D is neither E nor E_<CSV Final_Rating>", []
            ).append(
                f"row {asset.row} (D {asset.d!r}, E {asset.e!r}, "
                f"Final_Rating {asset.rating!r})"
            )
    for item in source:
        _check_list_entries(item, problems)
    _raise_problems("Stage 3 could not align the MACalc ISIN lists", problems)
    aligned = tuple(_align_one(item, by_d, by_e, problems) for item in source)
    _raise_problems("Stage 3 could not align the MACalc ISIN lists", problems)
    return IsinListAlignment(aligned)


def log_isin_alignment(alignment: IsinListAlignment) -> None:
    """Print the per-list Slice C summary, unmatched keys and fan-outs."""
    for item in alignment.lists:
        logger.info("%s", item.summary())
        if item.unmatched and len(item.unmatched) <= MAX_LISTED_PROBLEMS:
            logger.info(
                "  %s unmatched (kept in place): %s",
                item.spec.name,
                ", ".join(item.unmatched),
            )
        for key, emitted in item.fan_outs:
            logger.warning(
                "  %s fan-out: %s -> %s", item.spec.name, key, ", ".join(emitted)
            )


@contextmanager
def _excel_application() -> Iterator[Any]:
    import xlwings as xw

    application = xw.App(visible=False, add_book=True)
    try:
        application.api.AutomationSecurity = 3
        application.api.EnableEvents = False
        application.api.DisplayAlerts = False
        application.api.ScreenUpdating = False
        application.api.CalculateBeforeSave = False
        application.api.Calculation = -4135
        yield application
    finally:
        try:
            application.api.DisplayAlerts = False
            application.api.Quit()
        except Exception:
            logger.exception("Could not close the dedicated Stage 3 Excel instance")


def _external_link_sources(book: Any) -> tuple[str, ...]:
    """Return a workbook's external Excel-link paths in their stored order."""
    sources = book.api.LinkSources(1)
    if sources is None:
        return ()
    if isinstance(sources, str):
        return (sources,)
    return tuple(sources)


def _restore_external_link_sources(
    book: Any, expected_sources: Sequence[str]
) -> bool:
    """Undo Excel's path canonicalisation of external links when saving a copy."""
    current_sources = _external_link_sources(book)
    expected = tuple(expected_sources)
    if len(current_sources) != len(expected):
        raise StageThreeError(
            "Saved MACalc workbook has a different number of external links "
            "from the template"
        )
    changed = False
    for current, original in zip(current_sources, expected):
        if current != original:
            book.api.ChangeLink(current, original, 1)
            changed = True
    if _external_link_sources(book) != expected:
        raise StageThreeError(
            "Could not preserve the MACalc template's external link paths"
        )
    return changed


def _retained_template_links(
    template_links: Sequence[str], current_links: Sequence[str]
) -> tuple[str, ...]:
    """Return the template links still in use after Stage 3 replaced the asset rows.

    Excel drops a link once no cell refers to it, so links used only by the
    replaced rows (for example an old model CSV) legitimately disappear. Excel
    may also re-path the remaining links (SharePoint URL to UNC), so links are
    matched in order by file name and the template's original paths returned.
    Any added or reordered link keeps the full template set, so the save-time
    checks still reject it (Decision 84).
    """

    def name(link: str) -> str:
        return re.split(r"[\\/]", link)[-1].casefold()

    remaining = iter(template_links)
    retained = []
    for current in current_links:
        match = next((link for link in remaining if name(link) == name(current)), None)
        if match is None:
            return tuple(template_links)
        retained.append(match)
    return tuple(retained)


def _grid(value: object) -> list[list[object]]:
    if not isinstance(value, (tuple, list)):
        return [[value]]
    values = list(value)
    if not values:
        return []
    if isinstance(values[0], (tuple, list)):
        return [list(row) for row in values]
    return [values]


def _check_path_distinct(paths: Sequence[Path]) -> None:
    for index, path in enumerate(paths):
        if any(path.samefile(other) for other in paths[:index]):
            raise StageThreeError(
                "MAP, MACalc template and MACE must be distinct files"
            )


def _named_value(caller: Any, name: str) -> object:
    try:
        return caller.names[name].refers_to_range.value
    except Exception as error:
        raise StageThreeError(
            f"Required Draft name is missing or invalid: {name}"
        ) from error


def _resolve_control_path(
    value: object, caller_path: Path, label: str, relative_ok: bool
) -> Path:
    if not isinstance(value, str) or not value.strip():
        raise StageThreeError(f"{label} path is blank")
    path = Path(os.path.expandvars(value.strip())).expanduser()
    if not relative_ok and not path.is_absolute():
        raise StageThreeError(f"{label} path must be absolute: {value}")
    if relative_ok and not path.is_absolute():
        path = caller_path.parent / path
    path = path.resolve()
    if not path.is_file():
        raise StageThreeError(f"{label} workbook does not exist: {path}")
    return path


def resolve_stage_three_inputs(caller: Any) -> StageThreeInputs:
    """Require a current successful Stage 2 and resolve its MAP/template/MACE inputs."""
    shared.require_stage_two_panel(caller)
    caller_path = shared.resolve_caller_path(caller)
    run_folder = shared.resolve_stage_input(caller)
    input_run = shared.portable_output_path(caller_path, run_folder)
    stage_two = shared.read_stage_two_audit(caller)
    if (
        stage_two.get("status") != "Succeeded"
        or stage_two.get("input_run") != input_run
    ):
        raise StageThreeError(
            "Stage 2 is not Succeeded for the current Stage 1 run: "
            f"Stage 1 run {input_run!r}; Stage 2 input run "
            f"{stage_two.get('input_run')!r} (status {stage_two.get('status')!r})"
        )
    map_path = run_folder / shared.INPUT_FILENAMES["MAP"]
    if not map_path.is_file():
        raise StageThreeError(f"Stage 2 MAP input is missing: {map_path}")
    bonds_path = run_folder / shared.INPUT_FILENAMES["Bonds"]
    if not bonds_path.is_file():
        raise StageThreeError(f"Stage 1 Bonds input is missing: {bonds_path}")
    template_path = _resolve_control_path(
        _named_value(caller, "SST_MACalc_Input_Template"),
        caller_path,
        "MACalc template",
        relative_ok=True,
    )
    mace_path = _resolve_control_path(
        _named_value(caller, "SST_MACE_Central_Template"),
        caller_path,
        "MACE",
        relative_ok=False,
    )
    _check_path_distinct(
        (map_path.resolve(), bonds_path.resolve(), template_path, mace_path)
    )
    warning = caller.sheets["I. Control"].range("H44").options(err_to_str=True).value
    if isinstance(warning, str) and (
        "warning" in warning.casefold() or "mismatch" in warning.casefold()
    ):
        raise StageThreeError(
            f"Valuation-date guard failed at I. Control!H44: {warning}"
        )
    if _cell_error(warning):
        raise StageThreeError(
            f"Valuation-date guard returned an Excel error: {warning}"
        )
    return StageThreeInputs(
        caller_path,
        run_folder,
        input_run,
        map_path.resolve(),
        template_path,
        mace_path,
        map_base_output_path=resolve_map_base_output(run_folder),
        valuation_date=_named_value(caller, "I_Valuation_Date"),
        bonds_path=bonds_path.resolve(),
    )


def find_map_base_output(map_output_folder: Path) -> Path:
    """Return the single base-scenario CSV directly inside the MAP output folder."""
    folder = Path(map_output_folder)
    pattern = f"*{MAP_BASE_OUTPUT_PATTERN}*.csv"
    matches = sorted(path for path in folder.glob(pattern) if path.is_file())
    if not matches:
        raise StageThreeError(
            f"No Stage 2 base-scenario CSV matching {pattern!r} was found in "
            f"{folder}. If Stage 1 has replaced this output folder since Stage 2 "
            "ran, rerun Stage 2 before Stage 3. If the Bonds Lite model renamed "
            "the CSV, update MAP_BASE_OUTPUT_PATTERN in sst_stage_three.py"
        )
    if len(matches) > 1:
        raise StageThreeError(
            f"More than one Stage 2 base-scenario CSV matches {pattern!r} in "
            f"{folder}: " + ", ".join(path.name for path in matches)
        )
    return matches[0].resolve()


def resolve_map_base_output(run_folder: Path) -> Path:
    """Locate the base-scenario CSV in the MAP output folder Stage 2 validated."""
    map_input = shared.read_model_inputs(Path(run_folder))["MAP"]
    return find_map_base_output(map_input.output_folder)


def _blank(value: object) -> bool:
    return value is None or isinstance(value, str) and not value.strip()


def _valid_text(value: object) -> bool:
    return isinstance(value, str) and bool(value.strip()) and not _cell_error(
        value.strip()
    )


def _valid_number(value: object) -> bool:
    try:
        _numeric(value, "value")
    except StageThreeError:
        return False
    return True


def _csv_number(text: str) -> float | None:
    try:
        value = float(text)
    except ValueError:
        return None
    return value if math.isfinite(value) else None


def _listed(values: Sequence[str]) -> str:
    shown = ", ".join(values[:MAX_LISTED_PROBLEMS])
    hidden = len(values) - MAX_LISTED_PROBLEMS
    return shown + (f" (and {hidden} more)" if hidden > 0 else "")


def _raise_problems(title: str, problems: Mapping[str, Sequence[str]]) -> None:
    if not problems:
        return
    for category, labels in problems.items():
        logger.error("%s - %s (%d): %s", title, category, len(labels), labels)
    raise StageThreeError(
        title
        + ":\n"
        + "\n".join(
            f"- {category} ({len(labels)}): {_listed(list(labels))}"
            for category, labels in problems.items()
        )
    )


def _require_map_base_headers(headers: Sequence[object], path: Path) -> None:
    errors = [
        f"{column} must be {name!r}"
        for column, name in MAP_BASE_CSV_HEADERS.items()
        if _normalise_header(_header_value(headers, column)) != name.casefold()
    ]
    months = list(headers[_column_number("K") - 1 : _column_number("ATN")])
    if len(months) != MONTH_COUNT or any(
        not isinstance(value, str) or value.strip() != str(index)
        for index, value in enumerate(months)
    ):
        errors.append("K:ATN must contain month headers 0 through 1199")
    if errors:
        raise StageThreeError(
            f"Base-scenario CSV header mismatch in {path.name}: " + "; ".join(errors)
        )


def read_map_base_output(path: Path) -> MapBaseOutput:
    """Read the Stage 2 base-scenario CSV in file order, failing on any bad value."""
    path = Path(path).resolve(strict=True)
    with path.open(newline="", encoding="utf-8-sig") as stream:
        reader = csv.reader(stream)
        headers = next(reader, None)
        if headers is None:
            raise StageThreeError(f"Base-scenario CSV is empty: {path}")
        _require_map_base_headers(headers, path)
        records = list(reader)
    if not records:
        raise StageThreeError(f"Base-scenario CSV has no asset rows: {path}")

    width = _column_number("ATN")
    first_value = _column_number("J") - 1
    text_fields = {"A": "isin", "B": "unique_isin", "E": "currency", "F": "sector",
                   "H": "final_rating", "I": "rating"}
    problems: dict[str, list[str]] = {}
    rows: list[MapBaseRow] = []
    for line, record in enumerate(records, start=2):
        label = record[0] if record and record[0].strip() else f"CSV line {line}"
        if len(record) < width:
            problems.setdefault("rows shorter than column ATN", []).append(label)
            continue
        texts: dict[str, str] = {}
        for column, key in text_fields.items():
            value = record[_column_number(column) - 1]
            if not _valid_text(value):
                problems.setdefault(
                    f"blank or error {MAP_BASE_CSV_HEADERS[column]}", []
                ).append(label)
            texts[key] = value
        zspread = _csv_number(record[_column_number("D") - 1])
        if zspread is None:
            problems.setdefault("blank or non-numeric IM_Zspreads", []).append(label)
        values = [_csv_number(text) for text in record[first_value:width]]
        if any(value is None for value in values):
            problems.setdefault(
                "blank or non-numeric Base_PV or monthly values (J:ATN)", []
            ).append(label)
        rows.append(
            MapBaseRow(
                pv_and_cashflows=tuple(values),  # type: ignore[arg-type]
                zspread=zspread,  # type: ignore[arg-type]
                **texts,
            )
        )
    for key, header in (("isin", "ISIN"), ("unique_isin", "Unique_ISIN")):
        seen: set[str] = set()
        duplicates: list[str] = []
        for row in rows:
            value = getattr(row, key)
            if value in seen and value not in duplicates and value.strip():
                duplicates.append(value)
            seen.add(value)
        if duplicates:
            problems[f"duplicate {header}"] = duplicates
    _raise_problems(f"Base-scenario CSV {path.name} failed validation", problems)
    return MapBaseOutput(path, tuple(rows))


def _header_value(headers: Sequence[object], column: str) -> object:
    index = _column_number(column) - 1
    return headers[index] if index < len(headers) else None


def _require_map_headers(headers: Sequence[object], label: str = "MAP") -> None:
    expected = {
        "C": "ISIN",
        "E": "Currency",
        "F": "Sector",
        "H": "NotchedRating",
        "J": "Index Linked?",
    }
    errors = [
        f"{column} must be {value!r}"
        for column, value in expected.items()
        if _normalise_header(_header_value(headers, column)) != value.casefold()
    ]
    first_month = _column_number("Q") - 1
    last_month = _column_number("ATT") - 1
    months = headers[first_month : last_month + 1]
    if len(months) != MONTH_COUNT or any(
        value != str(index + 1)
        for index, value in enumerate(months)
    ):
        errors.append("Q:ATT must contain month ordinals 1 through 1200")
    if errors:
        raise StageThreeError(
            f"{label} Bonds Data header mismatch: " + "; ".join(errors)
        )


def _read_bonds_data_rows(
    path: Path, label: str
) -> list[tuple[int, list[object]]]:
    """Bulk-read C:ATT of I_tbl_Monthly_Bonds from a Stage 1 Bonds Lite input."""
    path = Path(path).resolve(strict=True)
    with _excel_application() as application:
        with shared.owned_book(application, path, read_only=True) as book:
            try:
                sheet = book.sheets[MAP_SHEET]
                table = sheet.api.ListObjects("I_tbl_Monthly_Bonds")
                if int(table.HeaderRowRange.Row) != MAP_HEADER_ROW:
                    raise StageThreeError(
                        f"{label} table I_tbl_Monthly_Bonds must have headers on row 8"
                    )
            except StageThreeError:
                raise
            except Exception as error:
                raise StageThreeError(
                    f"{label} workbook has no I_tbl_Monthly_Bonds table on {MAP_SHEET}"
                ) from error
            headers = (
                sheet.range(f"A{MAP_HEADER_ROW}:ATT{MAP_HEADER_ROW}")
                .options(err_to_str=True)
                .value
            )
            _require_map_headers(headers, label)
            first_data = int(table.DataBodyRange.Row)
            last_data = first_data + int(table.DataBodyRange.Rows.Count) - 1
            if last_data < first_data:
                raise StageThreeError(f"{label} Bonds Data table is empty")
            values = _grid(
                sheet.range(f"C{first_data}:ATT{last_data}")
                .options(ndim=2, err_to_str=True)
                .value
            )
    return list(enumerate(values, start=first_data))


def _row_has_negative_cashflow(
    values_row: Sequence[object], isin: str, label: str
) -> bool:
    month_offset = _column_number("Q") - _column_number("C")
    cashflows = values_row[month_offset : month_offset + MONTH_COUNT]
    if len(cashflows) != MONTH_COUNT:
        raise StageThreeError(f"{label} cashflow vector for {isin} is incomplete")
    numeric_cashflows = [
        _numeric(value, f"{label} cashflow for {isin}, month {month}")
        for month, value in enumerate(cashflows, start=1)
    ]
    return min(numeric_cashflows) < 0


def read_bonds_negative_isins(path: Path) -> set[str]:
    """Return ISINs with any negative monthly cashflow in the Stage 1 Bonds input.

    These are the funded assets that MACE must agree with; Stage 1 removes the
    negative cashflows before writing the MAP input.
    """
    label = "Bonds input"
    negative_isins: set[str] = set()
    for row_number, values_row in _read_bonds_data_rows(path, label):
        isin = values_row[0] if values_row else None
        if not isinstance(isin, str) or not isin.strip():
            raise StageThreeError(
                f"{label} ISIN is blank or invalid in row {row_number}"
            )
        if _row_has_negative_cashflow(values_row, isin, label):
            negative_isins.add(isin)
    return negative_isins


def read_map_bonds(path: Path) -> MapBonds:
    """Read cached MAP descriptions in one bulk read."""
    by_isin: dict[str, dict[str, object]] = {}
    for row_number, values_row in _read_bonds_data_rows(path, "MAP"):
        isin = values_row[0] if values_row else None
        if not isinstance(isin, str) or not isin.strip():
            raise StageThreeError(f"MAP ISIN is blank or invalid in row {row_number}")
        if isin in by_isin:
            raise StageThreeError(f"Duplicate MAP ISIN {isin!r} in row {row_number}")
        currency, sector, rating, index_linked = (
            values_row[2],
            values_row[3],
            values_row[5],
            values_row[7],
        )
        if any(
            _cell_error(value) for value in (currency, sector, rating, index_linked)
        ):
            raise StageThreeError(
                f"MAP descriptive data contains an Excel error for {isin}"
            )
        if index_linked is None:
            raise StageThreeError(f"MAP Index Linked? has no cached value for {isin}")
        by_isin[isin] = {
            "isin": isin,
            "currency": currency,
            "sector": sector,
            "rating": rating,
            "index_linked": index_linked,
        }
    return MapBonds(by_isin)


def _require_mace_headers(headers: Sequence[object]) -> None:
    expected = {"B": "ISIN", "C": "Z-Spread", "D": "Type", "I": "FS Add On"}
    errors = [
        f"{column} must be {value!r}"
        for column, value in expected.items()
        if _normalise_header(_header_value(headers, column)) != value.casefold()
    ]
    if _header_value(headers, "AI") != 0:
        errors.append("AI must be header ordinal 0")
    first_month = _column_number("AJ") - 1
    last_month = _column_number("AUM") - 1
    months = headers[first_month : last_month + 1]
    if len(months) != MONTH_COUNT or any(
        isinstance(value, bool) or value != index + 1
        for index, value in enumerate(months)
    ):
        errors.append("AJ:AUM must contain month ordinals 1 through 1200")
    if errors:
        raise StageThreeError("MACE Funding header mismatch: " + "; ".join(errors))


def _read_mace_funding_sheet(book: Any) -> MaceFundingTable:
    try:
        sheet = book.sheets[MACE_SHEET]
    except Exception as error:
        raise StageThreeError(f"MACE workbook has no {MACE_SHEET} sheet") from error
    headers = (
        sheet.range(f"A{MACE_HEADER_ROW}:AUM{MACE_HEADER_ROW}")
        .options(err_to_str=True)
        .value
    )
    _require_mace_headers(headers)
    last_row = sheet.range("B1048576").end("up").row
    if last_row < MACE_FIRST_DATA_ROW:
        raise StageThreeError("MACE C. Assets - Funding has no data rows")
    values = _grid(
        sheet.range(f"B{MACE_FIRST_DATA_ROW}:AUM{last_row}")
        .options(ndim=2, err_to_str=True)
        .value
    )
    rows = []
    for row_number, values_row in enumerate(values, start=MACE_FIRST_DATA_ROW):
        isin = values_row[0] if values_row else None
        if not isinstance(isin, str) or not isin.strip():
            raise StageThreeError(f"MACE ISIN is blank or invalid in row {row_number}")
        cashflows = values_row[34 : 34 + MONTH_COUNT]
        if len(cashflows) != MONTH_COUNT:
            raise StageThreeError(f"MACE funding vector for {isin} is incomplete")
        rows.append(
            {
                "isin": isin,
                "zspread": values_row[1],
                "type": values_row[2],
                "fsao": values_row[7],
                "ai": values_row[33],
                "funding_cashflows": cashflows,
            }
        )
    return MaceFundingTable(tuple(rows))


def read_mace_funding(path: Path) -> MaceFundingTable:
    """Read saved MACE values read-only without updating links or recalculating."""
    path = Path(path).resolve(strict=True)
    with _excel_application() as application:
        with shared.owned_book(application, path, read_only=True) as book:
            return _read_mace_funding_sheet(book)


def _column_letter(number: int) -> str:
    letters = ""
    while number:
        number, remainder = divmod(number - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def _mace_sheet(book: Any, name: str) -> Any:
    try:
        return book.sheets[name]
    except Exception as error:
        raise StageThreeError(f"MACE workbook has no {name} sheet") from error


def parse_mace_bond_attributes(
    headers: Sequence[object], body: Sequence[Sequence[object]]
) -> MaceBondAttributes:
    """Key MACE I. Bonds rows (D:L) on the ISIN in D, recording duplicate keys."""
    errors = [
        f"{column}{MACE_BONDS_HEADER_ROW} must be {name!r}"
        for column, name in MACE_BONDS_HEADERS.items()
        if _normalise_header(_header_value(headers, column)) != name.casefold()
    ]
    if errors:
        raise StageThreeError("MACE I. Bonds header mismatch: " + "; ".join(errors))
    by_isin: dict[str, dict[str, object]] = {}
    duplicates: set[str] = set()
    for row in body:
        isin = row[0] if row else None
        if not _valid_text(isin):
            continue
        if isin in by_isin:
            duplicates.add(isin)  # type: ignore[arg-type]
            continue
        by_isin[isin] = {  # type: ignore[index]
            "type": row[3],
            "index_linked": row[6],
            "fsao": row[8],
        }
    return MaceBondAttributes(by_isin, frozenset(duplicates))


def read_mace_bond_attributes(book: Any) -> MaceBondAttributes:
    """Read Type, Index Linked? and FSAO from MACE I. Bonds in one bulk read."""
    sheet = _mace_sheet(book, MACE_BONDS_SHEET)
    row = MACE_BONDS_HEADER_ROW
    headers = sheet.range(f"A{row}:L{row}").options(err_to_str=True).value
    parse_mace_bond_attributes(headers, [])
    last_row = sheet.range("D1048576").end("up").row
    if last_row <= row:
        raise StageThreeError("MACE I. Bonds has no data rows")
    body = _grid(
        sheet.range(f"D{row + 1}:L{last_row}").options(ndim=2, err_to_str=True).value
    )
    return parse_mace_bond_attributes(headers, body)


def _not_contained(detail: str) -> StageThreeError:
    return StageThreeError(
        f"The MACE Control allocation table is not contained within K:O: {detail}"
    )


def _require_control_headers(headers: Sequence[object]) -> None:
    if len(headers) != len(CONTROL_HEADERS) or any(
        _normalise_header(actual) != expected.casefold()
        for actual, expected in zip(headers, CONTROL_HEADERS)
    ):
        raise _not_contained(
            f"row {CONTROL_HEADER_ROW} K:O must be headed {CONTROL_HEADERS}; "
            f"found {list(headers)}"
        )


def parse_control_table(
    headers: Sequence[object], body: Sequence[Sequence[object]]
) -> ControlTable:
    """Read the K:O allocation table from row 10 down to the first blank ISIN."""
    _require_control_headers(headers)
    first_row = CONTROL_HEADER_ROW + 1
    end = next(
        (index for index, row in enumerate(body) if _blank(row[0])), len(body)
    )
    if end == 0:
        raise StageThreeError("The MACE Control allocation table has no rows")
    later = [
        first_row + end + offset
        for offset, row in enumerate(body[end : end + CONTROL_LOOKAHEAD_ROWS])
        if not _blank(row[0])
    ]
    if later:
        raise _not_contained(
            f"a second block of ISINs starts at K{later[0]}, below the blank K"
            f"{first_row + end}"
        )
    by_isin: dict[str, dict[str, object]] = {}
    duplicates: set[str] = set()
    for row in body[:end]:
        isin = row[0]
        if not _valid_text(isin):
            continue
        if isin in by_isin:
            duplicates.add(isin)  # type: ignore[arg-type]
            continue
        by_isin[isin] = {"a": row[1], "b": row[2], "c": row[4]}  # type: ignore[index]
    return ControlTable(by_isin, frozenset(duplicates), first_row + end - 1)


def require_control_side_empty(
    side: Sequence[Sequence[object]], first_row: int = CONTROL_HEADER_ROW
) -> None:
    """Stop if any cell in the configured side columns beside the table is used."""
    first_column = _column_number(CONTROL_SIDE_COLUMNS[0])
    used = [
        f"{_column_letter(first_column + offset)}{first_row + row_offset}"
        for row_offset, row in enumerate(side)
        for offset, value in enumerate(row)
        if not _blank(value)
    ]
    if used:
        raise _not_contained(
            f"{CONTROL_SIDE_COLUMNS[0]}:{CONTROL_SIDE_COLUMNS[1]} beside the table "
            "must be empty; found data in " + _listed(used)
        )


def read_mace_allocations(book: Any) -> MaceAllocations:
    """Read the Control K:O allocation table and the MACE valuation date (C9)."""
    sheet = _mace_sheet(book, MACE_CONTROL_SHEET)
    row = CONTROL_HEADER_ROW
    headers = sheet.range(f"K{row}:O{row}").options(err_to_str=True).value
    _require_control_headers(headers)
    last_row = max(sheet.range("K1048576").end("up").row, row + 1)
    body = _grid(
        sheet.range(f"K{row + 1}:O{last_row + CONTROL_LOOKAHEAD_ROWS}")
        .options(ndim=2, err_to_str=True)
        .value
    )
    table = parse_control_table(headers, body)
    first, last = CONTROL_SIDE_COLUMNS
    require_control_side_empty(
        _grid(
            sheet.range(f"{first}{row}:{last}{table.last_row}")
            .options(ndim=2, err_to_str=True)
            .value
        )
    )
    valuation_date = sheet.range(f"C{row}").options(err_to_str=True).value
    return MaceAllocations(table.by_isin, table.duplicates, valuation_date)


def read_mace_inputs(path: Path) -> MaceInputs:
    """Open MACE once, read-only, and read its funding, I. Bonds and Control tabs."""
    path = Path(path).resolve(strict=True)
    with _excel_application() as application:
        with shared.owned_book(application, path, read_only=True) as book:
            return MaceInputs(
                funding=_read_mace_funding_sheet(book),
                bonds=read_mace_bond_attributes(book),
                allocations=read_mace_allocations(book),
            )


def parse_isin_list(
    spec: IsinListSpec, block: Sequence[Sequence[object]], problems: dict[str, list[str]]
) -> IsinList:
    """Parse one list from a grid read from row 1, spanning its neighbour columns.

    ``block`` starts at the leftmost of the key, value and neighbour columns and
    ends on the last used row of the key/value columns.
    """
    columns = [spec.key_column, *([spec.value_column] if spec.value_column else [])]
    first = min(
        _column_number(column) for column in (*columns, *spec.neighbour_columns)
    )

    def cell(column: str, row: int) -> object:
        values = block[row - 1] if row <= len(block) else ()
        index = _column_number(column) - first
        return values[index] if index < len(values) else None

    def add(category: str, label: str) -> None:
        problems.setdefault(f"{spec.name} ({spec.sheet}): {category}", []).append(
            label
        )

    key = spec.key_column
    for row in range(1, LIST_TITLE_ROW):
        if cell(key, row) is not None:
            add("cells above the title must be empty (COUNTA would misalign)",
                f"{key}{row}")
    title = cell(key, LIST_TITLE_ROW)
    if not isinstance(title, str) or title.strip() != spec.title:
        add(f"title must be {spec.title!r}", f"{key}{LIST_TITLE_ROW} ({title!r})")
    for column in columns:
        if cell(column, LIST_TITLE_ROW + 1) is not None:
            add("row 6 must be blank", f"{column}{LIST_TITLE_ROW + 1}")

    last_row = max(len(block), LIST_FIRST_ROW - 1)
    rows = range(LIST_FIRST_ROW, last_row + 1)
    end = next((row for row in rows if cell(key, row) is None), last_row + 1) - 1
    later = [f"{key}{row}" for row in range(end + 1, last_row + 1)
             if cell(key, row) is not None]
    if later:
        add(
            "keys below the first blank (COUNTA would misalign the range)",
            f"{', '.join(later)} (first blank {key}{end + 1})",
        )
    stray = [
        f"{column}{row}"
        for column in spec.neighbour_columns
        for row in range(LIST_TITLE_ROW, end + 1)
        if cell(column, row) is not None
    ]
    if spec.value_column:
        stray.extend(
            f"{spec.value_column}{row}"
            for row in range(end + 1, last_row + 1)
            if cell(spec.value_column, row) is not None
        )
    for label in stray:
        add("list is not contained (data beside or below it)", label)
    entries = range(LIST_FIRST_ROW, end + 1)
    return IsinList(
        spec,
        tuple(cell(key, row) for row in entries),
        tuple(cell(spec.value_column, row) for row in entries)
        if spec.value_column
        else None,
        last_row,
    )


def read_isin_lists(book: Any) -> TemplateIsinLists:
    """Read the D-keyed MACalc lists in one bulk read each, applying Guards 1-3 and 6."""
    problems: dict[str, list[str]] = {}
    lists = []
    for spec in MACALC_ISIN_LISTS:
        try:
            refers_to = book.names[spec.name].refers_to
        except Exception:
            refers_to = None
        if refers_to != spec.refers_to:
            problems.setdefault("Names that are missing or not the expected OFFSET",
                                []).append(f"{spec.name} ({refers_to!r})")
        try:
            sheet = book.sheets[spec.sheet]
        except Exception as error:
            raise StageThreeError(
                f"MACalc template has no {spec.sheet} sheet"
            ) from error
        columns = [spec.key_column, spec.last_column]
        last_row = max(
            sheet.range(f"{column}1048576").end("up").row for column in columns
        )
        numbers = [_column_number(column)
                   for column in (*columns, *spec.neighbour_columns)]
        address = (
            f"{_column_letter(min(numbers))}1:{_column_letter(max(numbers))}{last_row}"
        )
        block = _grid(sheet.range(address).options(ndim=2, err_to_str=True).value)
        lists.append(parse_isin_list(spec, block, problems))
    _raise_problems("The MACalc template ISIN lists failed validation", problems)
    return TemplateIsinLists(tuple(lists))


def _fixed_asset_keys(
    rows: Sequence[Sequence[object]], first_row: int = 2
) -> tuple[AssetKey, ...]:
    """Key the Mortgage/ILS rows from a D:I grid."""
    return tuple(
        AssetKey(first_row + offset, row[0], row[1], "template", row[5])
        for offset, row in enumerate(rows)
    )


def _read_template_snapshot(path: Path) -> TemplateSnapshot:
    with _excel_application() as application:
        with shared.owned_book(application, path, read_only=True) as book:
            try:
                sheet = book.sheets[MACALC_SHEET]
            except Exception as error:
                raise StageThreeError(
                    f"MACalc template has no {MACALC_SHEET} sheet"
                ) from error
            headers = sheet.range("A1:ATQ1").options(err_to_str=True).value
            validate_template_headers(headers)
            last_used = int(
                sheet.api.UsedRange.Row + sheet.api.UsedRange.Rows.Count - 1
            )
            if last_used < 2:
                raise TemplateLayoutError("MACalc template has no asset rows")
            types = (
                sheet.range(f"G2:G{last_used}").options(ndim=2, err_to_str=True).value
            )
            layout = validate_template_blocks([row[0] for row in _grid(types)])
            fixed_keys = _fixed_asset_keys(
                _grid(
                    sheet.range(f"D2:I{layout.ils_end_row}")
                    .options(ndim=2, err_to_str=True)
                    .value
                )
            )
            isin_lists = read_isin_lists(book)
    return TemplateSnapshot(layout, fixed_keys, isin_lists)


MACALC_EXACT_TYPES = frozenset({"Mortgage", "Infrastructure", "PP", "Swap", "ILS"})


def check_macalc_asset_table(
    a_values: Sequence[object], g_values: Sequence[object], first_row: int = 2
) -> None:
    """Stop if tbl_Assets would drop rows (blank A) or MACalc would reject a Type."""
    blank_a = [
        str(first_row + offset)
        for offset, value in enumerate(a_values)
        if value is None or isinstance(value, str) and not value.strip()
    ]
    bad_type = [
        f"{first_row + offset} ({value!r})"
        for offset, value in enumerate(g_values)
        if not (
            isinstance(value, str)
            and (value in MACALC_EXACT_TYPES or value.casefold() == "bond")
        )
    ]
    problems = {}
    if blank_a:
        problems["Rows with a blank column A (tbl_Assets would drop rows)"] = blank_a
    if bad_type:
        problems[
            "Rows whose Type (G) is not Mortgage, Bond, Infrastructure, PP, Swap or ILS"
        ] = bad_type
    _raise_problems("The MACalc I. Assets table would not load correctly", problems)


def _isin_list_end(aligned: AlignedIsinList) -> tuple[int, int]:
    """Return the last row of the new list and the last row to clear or check."""
    last = LIST_FIRST_ROW + len(aligned.keys) - 1
    return last, max(last, aligned.template_last_row)


def _isin_list_rows(aligned: AlignedIsinList, through: int) -> list[list[object]]:
    """Expected key (and value) cells from row 7 to ``through``, blank-padded."""
    width = 2 if aligned.spec.is_table else 1
    rows: list[list[object]] = [
        [key, aligned.values[offset]] if aligned.values is not None else [key]
        for offset, key in enumerate(aligned.keys)
    ]
    rows.extend([None] * width for _ in range(through - LIST_FIRST_ROW + 1 - len(rows)))
    return rows


def _write_isin_list(book: Any, aligned: AlignedIsinList) -> None:
    """Rewrite one list's key/value block from row 7 and clear the old tail."""
    spec = aligned.spec
    sheet = book.sheets[spec.sheet]
    last, clear_through = _isin_list_end(aligned)
    if clear_through >= LIST_FIRST_ROW:
        sheet.range(
            f"{spec.key_column}{LIST_FIRST_ROW}:{spec.last_column}{clear_through}"
        ).clear_contents()
    if aligned.keys:
        sheet.range(
            f"{spec.key_column}{LIST_FIRST_ROW}:{spec.last_column}{last}"
        ).value = _isin_list_rows(aligned, last)


def write_macalc_output(
    template_path: Path,
    staged_path: Path,
    layout: TemplateLayout,
    funded_rows: Sequence[Sequence[object]],
    asset_rows: Sequence[Sequence[object]] = (),
    isin_lists: Sequence[AlignedIsinList] = (),
) -> int:
    """Copy the template, clear post-ILS rows, and write both blocks and the lists."""
    blocks = output_blocks(layout, funded_rows, asset_rows)
    if any(len(row) != OUTPUT_WIDTH for _, rows in blocks for row in rows):
        raise StageThreeError("MACalc rows do not span columns A:ATQ")
    last_row = layout.funded_start_row + len(funded_rows) + len(asset_rows) - 1
    shutil.copy2(template_path, staged_path)
    with _excel_application() as application:
        with shared.owned_book(
            application, template_path, read_only=True
        ) as template:
            template_links = _external_link_sources(template)
        with shared.owned_book(application, staged_path, read_only=False) as book:
            sheet = book.sheets[MACALC_SHEET]
            last_used = int(
                sheet.api.UsedRange.Row + sheet.api.UsedRange.Rows.Count - 1
            )
            clear_through = max(last_used, last_row)
            if clear_through >= layout.funded_start_row:
                sheet.range(
                    f"A{layout.funded_start_row}:ATQ{clear_through}"
                ).clear_contents()
            for start, rows in blocks:
                if rows:
                    sheet.range(f"A{start}:ATQ{start + len(rows) - 1}").value = [
                        list(row) for row in rows
                    ]
            last_row = max(last_row, layout.ils_end_row)
            check_macalc_asset_table(
                [row[0] for row in _grid(
                    sheet.range(f"A2:A{last_row}").options(ndim=2).value
                )],
                [row[0] for row in _grid(
                    sheet.range(f"G2:G{last_row}")
                    .options(ndim=2, err_to_str=True)
                    .value
                )],
            )
            for aligned in isin_lists:
                _write_isin_list(book, aligned)
            expected_links = _retained_template_links(
                template_links, _external_link_sources(book)
            )
            for dropped in (
                link for link in template_links if link not in expected_links
            ):
                logger.info(
                    "Template link used only by the replaced asset rows was removed: %s",
                    dropped,
                )
            book.save()
            links_changed = _restore_external_link_sources(book, expected_links)
            if links_changed:
                book.save()
            if _external_link_sources(book) != expected_links:
                raise StageThreeError(
                    "Saved MACalc workbook changed the template's external links"
                )
            if not book.api.Saved:
                raise StageThreeError("Staged MACalc workbook was not saved")
    return clear_through


def _used_formulas(sheet: Any) -> list[list[object]]:
    """Read a sheet's formulas from A1 to the end of its used range."""
    used = sheet.api.UsedRange
    last_row = int(used.Row + used.Rows.Count - 1)
    last_column = int(used.Column + used.Columns.Count - 1)
    return _grid(
        sheet.range(f"A1:{_column_letter(last_column)}{last_row}")
        .options(ndim=2, err_to_str=True)
        .formula
    )


def _verify_isin_lists(
    book: Any,
    aligned_lists: Sequence[AlignedIsinList],
    realigned: IsinListAlignment,
    d_values: set[object],
    template_sheets: Mapping[str, Sequence[Sequence[object]]],
) -> None:
    """Prove the written lists, their names and every other cell on their sheets."""
    problems: dict[str, list[str]] = {}

    def add(category: str, label: str) -> None:
        problems.setdefault(category, []).append(label)

    excluded: dict[str, list[tuple[int, int, int]]] = {}
    for aligned in aligned_lists:
        spec = aligned.spec
        try:
            again = realigned.get(spec.name)
        except KeyError:
            again = None
        if again is None or (again.keys, again.values) != (
            aligned.keys,
            aligned.values,
        ):
            add("Lists that differ from the alignment re-run on the read-back "
                "I. Assets", spec.name)
        if sum(key in d_values for key in aligned.keys) != aligned.matched_keys:
            add("Lists whose matched-key count differs from the alignment",
                spec.name)
        sheet = book.sheets[spec.sheet]
        last, through = _isin_list_end(aligned)
        first_column = _column_number(spec.key_column)
        excluded.setdefault(spec.sheet, []).append(
            (first_column, _column_number(spec.last_column), through)
        )
        if through >= LIST_FIRST_ROW:
            block = sheet.range(
                f"{spec.key_column}{LIST_FIRST_ROW}:{spec.last_column}{through}"
            )
            if _grid(block.options(ndim=2, err_to_str=True).value) != (
                _isin_list_rows(aligned, through)
            ):
                add("Lists whose read-back differs from the written keys/values",
                    spec.name)
            if block.api.HasFormula is not False:
                add("Lists containing formulas", spec.name)
        if aligned.keys:
            try:
                resolved = book.names[spec.name].refers_to_range
                found = (
                    resolved.sheet.name,
                    resolved.row,
                    resolved.column,
                    resolved.rows.count,
                    resolved.columns.count,
                )
            except Exception:
                found = None
            wanted = (
                spec.sheet,
                LIST_FIRST_ROW,
                first_column,
                len(aligned.keys),
                2 if spec.is_table else 1,
            )
            if found != wanted:
                add("Names that do not resolve to the written list",
                    f"{spec.name} (expected {spec.sheet}!{spec.key_column}"
                    f"{LIST_FIRST_ROW}:{spec.last_column}{last}; got {found})")

    for sheet_name, template_grid in template_sheets.items():
        output_grid = _used_formulas(book.sheets[sheet_name])
        blocks = excluded.get(sheet_name, [])
        changed = []
        for row in range(1, max(len(template_grid), len(output_grid)) + 1):
            before = template_grid[row - 1] if row <= len(template_grid) else []
            after = output_grid[row - 1] if row <= len(output_grid) else []
            for column in range(1, max(len(before), len(after)) + 1):
                if any(
                    first <= column <= last and LIST_FIRST_ROW <= row <= through
                    for first, last, through in blocks
                ):
                    continue
                old = before[column - 1] if column <= len(before) else None
                new = after[column - 1] if column <= len(after) else None
                if (old if old is not None else "") != (new if new is not None else ""):
                    changed.append(f"{sheet_name}!{_column_letter(column)}{row}")
        if changed:
            problems[f"Cells outside the lists changed on {sheet_name}"] = changed
    _raise_problems("MACalc ISIN list read-back failed", problems)


def verify_macalc_output(
    path: Path,
    expected: Sequence[Sequence[object]],
    template_path: Path | None = None,
    layout: TemplateLayout | None = None,
    cleared_through: int | None = None,
    isin_lists: Sequence[AlignedIsinList] = (),
    template_isin_lists: TemplateIsinLists | None = None,
    funded_count: int = 0,
    asset_ratings: Sequence[object] = (),
) -> None:
    """Reopen and compare every written value, protected row, cleared row and size.

    With ``isin_lists``, also re-run the Slice C alignment of
    ``template_isin_lists`` on the read-back I. Assets (the first
    ``funded_count`` written rows are the funded block, and ``asset_ratings``
    holds the CSV ``Final_Rating`` of each asset row) and prove the lists.
    """
    if isin_lists and (template_path is None or template_isin_lists is None):
        raise StageThreeError(
            "MACalc ISIN list verification needs the template and its lists"
        )
    template_sheets: dict[str, list[list[object]]] = {}
    actual_rows: list[list[object]] = []
    with _excel_application() as application:
        with shared.owned_book(application, path, read_only=True) as book:
            sheet = book.sheets[MACALC_SHEET]
            if layout is None:
                headers = sheet.range("A1:ATQ1").options(err_to_str=True).value
                validate_template_headers(headers)
                last_type_row = sheet.range("G1048576").end("up").row
                types = _grid(
                    sheet.range(f"G2:G{last_type_row}")
                    .options(ndim=2, err_to_str=True)
                    .value
                )
                layout = validate_template_blocks([row[0] for row in types])
            if template_path is not None:
                with shared.owned_book(
                    application, template_path, read_only=True
                ) as template:
                    template_rows = _grid(
                        template.sheets[MACALC_SHEET]
                        .range(f"A2:ATQ{layout.ils_end_row}")
                        .options(ndim=2, err_to_str=True)
                        .formula
                    )
                    if isin_lists:
                        template_sheets = {
                            name: _used_formulas(template.sheets[name])
                            for name in dict.fromkeys(
                                item.spec.sheet for item in isin_lists
                            )
                        }
                output_rows = _grid(
                    sheet.range(f"A2:ATQ{layout.ils_end_row}")
                    .options(ndim=2, err_to_str=True)
                    .formula
                )
                if output_rows != template_rows:
                    raise StageThreeError(
                        "Mortgage/ILS rows changed from the MACalc template"
                    )
            start = layout.funded_start_row
            if expected:
                end = start + len(expected) - 1
                block = sheet.range(f"A{start}:ATQ{end}")
                actual_rows = _grid(block.options(ndim=2, err_to_str=True).value)
                for offset, (actual_row, expected_row) in enumerate(
                    zip(actual_rows, expected)
                ):
                    if actual_row != list(expected_row):
                        raise StageThreeError(
                            f"MACalc output read-back mismatch on row {start + offset}"
                        )
                if block.api.HasFormula is not False:
                    raise StageThreeError(
                        f"MACalc rows {start}-{end} must not contain formulas"
                    )
            used = sheet.api.UsedRange
            used_end = int(used.Row + used.Rows.Count - 1)
            cleared_through = max(cleared_through or 0, used_end)
            first_empty = start + len(expected)
            if first_empty <= cleared_through:
                remaining = _grid(
                    sheet.range(f"A{first_empty}:ATQ{cleared_through}")
                    .options(ndim=2, err_to_str=True)
                    .formula
                )
                if any(value not in (None, "") for row in remaining for value in row):
                    raise StageThreeError(
                        "MACalc rows below the asset block are not empty"
                    )
            table = _assets_table_range(book, sheet)
            expected_rows = first_empty - 1
            if table.row != 1 or table.rows.count != expected_rows:
                raise StageThreeError(
                    f"tbl_Assets covers rows {table.row}-"
                    f"{table.row + table.rows.count - 1}; expected 1-{expected_rows}"
                )
            if isin_lists:
                fixed_keys = _fixed_asset_keys(
                    _grid(
                        sheet.range(f"D2:I{layout.ils_end_row}")
                        .options(ndim=2, err_to_str=True)
                        .value
                    )
                )
                asset_keys = build_asset_keys(
                    layout,
                    fixed_keys,
                    actual_rows[:funded_count],
                    actual_rows[funded_count:],
                    asset_ratings=asset_ratings,
                )
                _verify_isin_lists(
                    book,
                    isin_lists,
                    align_isin_lists(template_isin_lists, asset_keys),  # type: ignore[arg-type]
                    {key.d for key in asset_keys},
                    template_sheets,
                )


def _assets_table_range(book: Any, sheet: Any) -> Any:
    for names in (sheet.names, book.names):
        try:
            return names[ASSETS_TABLE_NAME].refers_to_range
        except Exception:
            continue
    raise StageThreeError(f"MACalc output has no {ASSETS_TABLE_NAME} name")


def _stage_three_cells(caller: Any) -> dict[str, Any]:
    cells = {}
    for field, name in STAGE_THREE_FIELD_NAMES.items():
        region = caller.names[name].refers_to_range
        merged = region.api.Cells(1, 1).MergeArea
        cells[field] = region.sheet.range((merged.Row, merged.Column))
    return cells


def require_stage_three_panel(caller: Any) -> None:
    if caller.api.ReadOnly:
        raise StageThreeError("Draft is read-only; its Stage 3 audit cannot be saved")
    sheet = caller.sheets["I. Control"]
    if sheet.api.ProtectContents:
        raise StageThreeError(
            "I. Control is protected; its Stage 3 audit cannot be updated"
        )
    try:
        panel = caller.names[STAGE_THREE_PANEL_NAME].refers_to_range
        if panel.sheet.name != "I. Control" or panel.api.Areas.Count != 1:
            raise StageThreeError("Stage 3 audit panel name is invalid")
        bounds = (
            panel.row,
            panel.column,
            panel.row + panel.rows.count - 1,
            panel.column + panel.columns.count - 1,
        )
        occupied: list[tuple[int, int, int, int]] = []
        for field, name in STAGE_THREE_FIELD_NAMES.items():
            region = caller.names[name].refers_to_range
            if region.sheet.name != "I. Control":
                raise StageThreeError("Stage 3 audit fields must remain on I. Control")
            merge = region.api.Cells(1, 1).MergeArea
            field_bounds = (
                merge.Row,
                merge.Column,
                merge.Row + merge.Rows.Count - 1,
                merge.Column + merge.Columns.Count - 1,
            )
            if not (
                bounds[0] <= field_bounds[0] <= field_bounds[2] <= bounds[2]
                and bounds[1] <= field_bounds[1] <= field_bounds[3] <= bounds[3]
            ):
                raise StageThreeError(
                    f"Stage 3 audit field is outside its panel: {name}"
                )
            if any(
                field_bounds[0] <= old[2]
                and field_bounds[2] >= old[0]
                and field_bounds[1] <= old[3]
                and field_bounds[3] >= old[1]
                for old in occupied
            ):
                raise StageThreeError(f"Stage 3 audit fields overlap: {field}")
            occupied.append(field_bounds)
    except StageThreeError:
        raise
    except Exception as error:
        raise StageThreeError(
            "Stage 3 audit panel names are missing or invalid"
        ) from error


def _read_stage_three_fields(caller: Any) -> dict[str, object]:
    require_stage_three_panel(caller)
    return {field: cell.value for field, cell in _stage_three_cells(caller).items()}


def write_stage_three_audit(caller: Any, fields: dict[str, object]) -> None:
    require_stage_three_panel(caller)
    if set(fields) - STAGE_THREE_FIELD_NAMES.keys():
        raise StageThreeError("Unknown Stage 3 audit field")
    cells = _stage_three_cells(caller)
    events = caller.app.api.EnableEvents
    try:
        caller.app.api.EnableEvents = False
        for field, value in fields.items():
            cell = cells[field]
            if field in ("input_run", "last_output"):
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
            raise StageThreeError("Draft's Stage 3 audit was not saved")
    finally:
        caller.app.api.EnableEvents = events


def _message(caller: Any, text: str, warning: bool = False) -> None:
    if os.name == "nt":
        owner = int(getattr(caller.app.api, "Hwnd", 0))
        flags = (0x30 if warning else 0x10) | 0x1000
        user = ctypes.WinDLL("user32", use_last_error=True)
        user.MessageBoxW.argtypes = [
            ctypes.c_void_p,
            ctypes.c_wchar_p,
            ctypes.c_wchar_p,
            ctypes.c_uint,
        ]
        user.MessageBoxW(owner, text, "Bonds Lite - Stage 3", flags)
    elif warning:
        logger.warning("%s", text)


def _token() -> str:
    """Return a fresh 16-character lowercase hexadecimal output token."""
    return secrets.token_hex(8)


def output_filename(run_folder: Path, token: str) -> str:
    """Return the MACalc input name for a Stage 1 stress folder and token."""
    suffix = shared.output_root_suffix(run_folder.parent.name)
    if suffix is None:
        raise StageThreeError(
            "Stage 1 run folder is not inside an SST Output_<suffix> folder: "
            f"{run_folder}"
        )
    if not TOKEN_PATTERN.fullmatch(token):
        raise StageThreeError(f"MACalc input token is not 16 hex characters: {token}")
    return f"{OUTPUT_STEM}_{suffix}_{token}{OUTPUT_EXTENSION}"


def new_output_path(run_folder: Path) -> Path:
    """Return an unused, token-suffixed output path in the run folder."""
    for _ in range(MAX_TOKEN_ATTEMPTS):
        path = run_folder / output_filename(run_folder, _token())
        if not path.exists():
            return path
    raise StageThreeError(f"Could not find an unused MACalc input name in {run_folder}")


def run_stage_three(caller: Any) -> Path:
    """Prepare a verified MACalc workbook and commit it with a matching audit."""
    caller_path = shared.resolve_caller_path(caller)
    with shared.workbook_run_guard(caller_path):
        require_stage_three_panel(caller)
        current_run = shared.resolve_stage_input(caller)
        input_run = shared.portable_output_path(caller_path, current_run)
        previous = _read_stage_three_fields(caller)["last_output"]
        output_path = new_output_path(current_run)
        staging_root = current_run.parent / STAGING_FOLDER
        staging_folder = staging_root / output_path.stem.rsplit("_", 1)[1]
        for path in (output_path, staging_folder / output_path.name):
            if len(str(path.absolute())) > MAX_OUTPUT_PATH:
                raise StageThreeError(
                    f"MACalc input path is longer than {MAX_OUTPUT_PATH} characters; "
                    f"shorten the output folder path:\n{path}"
                )

        started = datetime.now()
        clock = time.perf_counter()
        write_stage_three_audit(
            caller,
            {
                "last_run": started.replace(hour=0, minute=0, second=0, microsecond=0),
                "start": started,
                "finish": None,
                "elapsed": None,
                "user": getpass.getuser(),
                "status": "Running",
                "input_run": input_run,
                "last_output": previous,
            },
        )
        staging_created = False
        staged_path: Path | None = None
        published = False
        committed = False
        try:
            inputs = resolve_stage_three_inputs(caller)
            template = _read_template_snapshot(inputs.template_path)
            map_bonds = read_map_bonds(inputs.map_path)
            bonds_negative = read_bonds_negative_isins(inputs.bonds_path)
            base = read_map_base_output(inputs.map_base_output_path)
            logger.info("Base-scenario CSV: %s", inputs.map_base_output_path)
            mace = read_mace_inputs(inputs.mace_path)
            funded = identify_funded(mace.funding)
            zero_funded = reconcile_funded(map_bonds, funded, bonds_negative)
            reconciliation = reconcile_slice_b(
                base,
                mace.bonds,
                mace.allocations,
                map_bonds,
                funded,
                inputs.valuation_date,
            )
            funded_rows = build_funded_rows(map_bonds, funded, mace.allocations)
            asset_rows = build_asset_rows(base, mace.bonds, mace.allocations)
            asset_ratings = tuple(row.final_rating for row in base.rows)
            # Slice C runs before any write, so a list guard leaves no output.
            alignment = align_isin_lists(
                template.isin_lists,
                build_asset_keys(
                    template.layout,
                    template.fixed_keys,
                    funded_rows,
                    asset_rows,
                    asset_ratings=asset_ratings,
                ),
            )
            log_isin_alignment(alignment)
            logger.info(
                "Stage 3 rows: %d funded, %d assets; MACE-only rows ignored: "
                "%d in I. Bonds, %d in Control",
                len(funded_rows),
                len(asset_rows),
                reconciliation.mace_only_bonds,
                reconciliation.mace_only_allocations,
            )
            map_warning = reconciliation.map_difference_warning()
            if map_warning:
                logger.warning("%s", map_warning)
            if zero_funded:
                warning = (
                    "WARNING: Stage 3 found 0 funded assets; "
                    "the funded block will be empty."
                )
                logger.warning("%s", warning)

            staging_folder.mkdir(parents=True)
            staging_created = True
            staged_path = staging_folder / output_path.name
            cleared_through = write_macalc_output(
                inputs.template_path,
                staged_path,
                template.layout,
                funded_rows,
                asset_rows,
                isin_lists=alignment.lists,
            )
            verify_macalc_output(
                staged_path,
                [*funded_rows, *asset_rows],
                template_path=inputs.template_path,
                layout=template.layout,
                cleared_through=cleared_through,
                isin_lists=alignment.lists,
                template_isin_lists=template.isin_lists,
                funded_count=len(funded_rows),
                asset_ratings=asset_ratings,
            )
            if output_path.exists():
                raise StageThreeError(f"MACalc input already exists: {output_path}")
            # From here output_path belongs to this run, so a failure may remove it.
            published = True
            shutil.move(staged_path, output_path)
            staged_path = None
            write_stage_three_audit(
                caller,
                {
                    "finish": datetime.now(),
                    "elapsed": (time.perf_counter() - clock) / 86400,
                    "status": "Succeeded",
                    "input_run": input_run,
                    "last_output": output_path.relative_to(
                        caller_path.parent
                    ).as_posix(),
                },
            )
            committed = True
            if zero_funded:
                _message(caller, warning, warning=True)
            if map_warning:
                _message(caller, map_warning, warning=True)
            logger.info(
                "Stage 3 succeeded%s; saved MACalc input:\n%s",
                " (0 funded assets)" if zero_funded else "",
                output_path,
            )
            return output_path
        except BaseException as error:
            # Only remove an output this run published; earlier ones are never touched.
            if published and not committed and output_path.exists():
                output_path.unlink()
            status = "Interrupted" if isinstance(error, KeyboardInterrupt) else "Failed"
            try:
                write_stage_three_audit(
                    caller,
                    {
                        "finish": datetime.now(),
                        "elapsed": (time.perf_counter() - clock) / 86400,
                        "status": status,
                        "input_run": input_run,
                        "last_output": previous,
                    },
                )
            except Exception as audit_error:
                logger.error(
                    "Stage 3 failed and its audit could not be saved: %s", audit_error
                )
            if isinstance(error, TemplateLayoutError):
                _message(caller, LAYOUT_WARNING)
                logger.error("%s", error)
            raise
        finally:
            if staged_path is not None and staged_path.exists():
                staged_path.unlink()
            if staging_created:
                _remove_staging(staging_folder, staging_root)


def _clear_read_only_and_retry(
    function: Any, path: str, error: BaseException
) -> None:
    """Retry a removal that failed only because OneDrive marked the path read-only."""
    mode = os.stat(path).st_mode
    if not isinstance(error, PermissionError) or mode & stat.S_IWRITE:
        raise error
    os.chmod(path, mode | stat.S_IWRITE)
    function(path)


def _remove_staging(staging_folder: Path, staging_root: Path) -> None:
    """Remove this run's staging folder, and the root once no other run uses it."""
    try:
        shutil.rmtree(staging_folder, onexc=_clear_read_only_and_retry)
    except OSError as error:
        logger.warning(
            "Could not remove the Stage 3 staging folder %s: %s", staging_folder, error
        )
        return
    try:
        staging_root.rmdir()
    except PermissionError as error:
        # Another Draft may be staging at the same time; keep the root then.
        with suppress(OSError):
            _clear_read_only_and_retry(os.rmdir, str(staging_root), error)
    except OSError:
        pass


def _install_stage_three_panel(book: Any) -> None:
    shared.require_stage_two_panel(book)
    expected_names = {STAGE_THREE_PANEL_NAME, *STAGE_THREE_FIELD_NAMES.values()}
    existing = {name.name.split("!")[-1] for name in book.names}
    if existing.intersection(expected_names):
        if not expected_names.issubset(existing):
            raise StageThreeError(
                "Stage 3 audit names are incomplete; installation was not changed"
            )
        require_stage_three_panel(book)
        return
    sheet = book.sheets["I. Control"]
    previous = book.names["SST_Stage2_Panel"].refers_to_range
    row = previous.row + previous.rows.count + 2
    column = previous.column
    for shape in sheet.api.Shapes:
        if shape.Name == "btnRunBondsLite":
            while sheet.range((row, column)).top < shape.Top + shape.Height + 12:
                row += 1
    area = sheet.range((row, column), (row + 10, column + 7))
    if (
        book.app.api.WorksheetFunction.CountA(area.api)
        or area.api.MergeCells is not False
    ):
        raise StageThreeError(
            "Proposed Stage 3 audit area is occupied; nothing was overwritten"
        )
    for shape in sheet.api.Shapes:
        if (
            shape.Left < area.left + area.width
            and shape.Left + shape.Width > area.left
            and shape.Top < area.top + area.height
            and shape.Top + shape.Height > area.top
        ):
            raise StageThreeError(
                "Proposed Stage 3 audit area overlaps a shape or button"
            )
    events = book.app.api.EnableEvents
    added: list[str] = []
    try:
        book.app.api.EnableEvents = False
        area.color, area.font.name, area.font.size = (230, 230, 255), "Calibri", 10
        area.api.VerticalAlignment = -4108
        heading = sheet.range((row, column), (row, column + 7))
        heading.merge()
        heading.value, heading.color = "Stage 3 - Prepare MACalc", (127, 127, 255)
        heading.font.color, heading.font.bold = (255, 255, 255), True
        book.names.add(STAGE_THREE_PANEL_NAME, f"='I. Control'!{area.get_address()}")
        added.append(STAGE_THREE_PANEL_NAME)
        for field, offset in STAGE_THREE_OFFSETS.items():
            end = row + offset + (1 if field in ("input_run", "last_output") else 0)
            label = sheet.range((row + offset, column), (end, column + 3))
            value = sheet.range((row + offset, column + 4), (end, column + 7))
            label.merge()
            value.merge()
            label.value, label.color, label.font.color = (
                STAGE_THREE_LABELS[field],
                (127, 127, 255),
                (255, 255, 255),
            )
            label.api.WrapText = value.api.WrapText = True
            value.font.color, value.api.HorizontalAlignment = (0, 0, 0), -4131
            value.number_format = (
                "dd/mm/yyyy"
                if field == "last_run"
                else (
                    "hh:mm:ss"
                    if field in ("start", "finish")
                    else "[h]:mm:ss" if field == "elapsed" else "General"
                )
            )
            if field == "status":
                value.value = "Not Run"
            name = STAGE_THREE_FIELD_NAMES[field]
            book.names.add(name, f"='I. Control'!{value[0, 0].get_address()}")
            added.append(name)
    except Exception:
        for name in reversed(added):
            book.names[name].delete()
        raise
    finally:
        book.app.api.EnableEvents = events
    require_stage_three_panel(book)


def _stage_three_button_row(panel: Any, button_left: float) -> int:
    """Mirror Stage 2: beside the panel when a 205pt button fits, else below."""
    if button_left + 205 <= panel.left:
        return panel.row
    return panel.row + panel.rows.count + 2


def install_stage_three_controls(caller: Path) -> Path:
    """Install Stage 3 panel/button into a closed Draft after creating a backup."""
    import xlwings as xw

    caller = Path(caller).resolve(strict=True)
    for application in xw.apps:
        for book in application.books:
            if book.name.casefold() == caller.name.casefold():
                raise StageThreeError(
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
                raise StageThreeError("Draft workbook is read-only or locked")
            _install_stage_three_panel(book)
            sheet = book.sheets["I. Control"]
            panel = book.names[STAGE_THREE_PANEL_NAME].refers_to_range
            row = _stage_three_button_row(panel, sheet.range((panel.row, 4)).left)
            anchor = sheet.range((row, 4), (row + 2, 6))
            _install_button(book, anchor)
            components = book.api.VBProject.VBComponents
            module = components.Item("modBondsLiteLauncher")
            code = module.CodeModule
            text = code.Lines(1, code.CountOfLines) if code.CountOfLines else ""
            if "Public Sub PrepareMACalc()" not in text:
                code.InsertLines(
                    code.CountOfLines + 1,
                    "\nPublic Sub PrepareMACalc()\n    LaunchBondsLite 3\nEnd Sub\n",
                )
            book.save()
        finally:
            book.close()
    finally:
        application.api.DisplayAlerts = False
        application.api.Quit()
    return backup


def _install_button(book: Any, anchor: Any) -> None:
    sheet = anchor.sheet
    if STAGE_THREE_BUTTON_NAME in [shape.Name for shape in sheet.api.Shapes]:
        button = sheet.api.Buttons().Item(STAGE_THREE_BUTTON_NAME)
    else:
        width = max(anchor.width, 205)
        end_column = anchor.column
        while sheet.range((anchor.row, end_column)).left < anchor.left + width:
            end_column += 1
        covered = sheet.range(
            (anchor.row, anchor.column),
            (anchor.row + anchor.rows.count - 1, end_column - 1),
        )
        if (
            book.app.api.WorksheetFunction.CountA(covered.api)
            or covered.api.MergeCells is not False
        ):
            raise StageThreeError(
                "Proposed Prepare MACalc button covers occupied cells or merges"
            )
        for shape in sheet.api.Shapes:
            if (
                shape.Left < anchor.left + width
                and shape.Left + shape.Width > anchor.left
                and shape.Top < anchor.top + anchor.height
                and shape.Top + shape.Height > anchor.top
            ):
                raise StageThreeError(
                    "Proposed Prepare MACalc button overlaps an existing shape"
                )
        button = sheet.api.Buttons().Add(anchor.left, anchor.top, width, anchor.height)
        button.Name = STAGE_THREE_BUTTON_NAME
    button.Caption = "Prepare MACalc"
    button.OnAction = (
        "'" + book.name.replace("'", "''") + "'!" + STAGE_THREE_BUTTON_MACRO
    )
