"""Prepare a Property-stress LTM Flash workbook from the saved Draft."""

from __future__ import annotations

import getpass
import logging
import os
import shutil
import tempfile
import time
from contextlib import contextmanager
from dataclasses import dataclass
from datetime import date, datetime
from numbers import Real
from pathlib import Path
from typing import Any, Iterator

import bonds_lite_input as shared

logger = shared.logger.getChild("sst_property_stage_one")
allocate_stress_folder = shared.allocate_stress_folder
discard_replaced_stress_folder = shared.discard_replaced_stress_folder
remove_incomplete_stress_folder = shared.remove_incomplete_stress_folder

PROPERTY_FLAG = "SST_Property_standalone"
INTEREST_FLAG = "SST_Interest_standalone"
BASE_BLOCK = "Base"
STRESS_BLOCKS = {
    (True, False): "Interest",
    (False, True): "Property",
    (True, True): "Interest_Property",
}
SCEN_DESC_COLUMN = 10
SCEN_DESC_HEADER = "scen desc"
INTEREST_STRESS_COLUMNS = (11, 12, 13)
PROPERTY_STRESS_COLUMNS = (17, 18)
TEMPLATE_NAME = "SST_LTMFlash_input_template"
SCENFLIPPER_SHEET = "I. ScenFlipper"
PROPERTY_PANEL_NAME = "SST_PropStage1_Panel"
PROPERTY_BUTTON_NAME = "btnPrepareLTMFlash"
PROPERTY_BUTTON_MACRO = "PrepareLTMFlash"
PROPERTY_FIELD_NAMES = {
    "last_run": "SST_PropStage1_LastRun",
    "start": "SST_PropStage1_Start",
    "finish": "SST_PropStage1_Finish",
    "elapsed": "SST_PropStage1_Elapsed",
    "user": "SST_PropStage1_User",
    "status": "SST_PropStage1_Status",
    "input_template": "SST_PropStage1_InputTemplate",
    "last_output": "SST_PropStage1_LastSuccessfulOutput",
}
PROPERTY_LABELS = {
    "last_run": "Last Run",
    "start": "Start:",
    "finish": "Finish:",
    "elapsed": "Time",
    "user": "User",
    "status": "Status",
    "input_template": "Input Template",
    "last_output": "Last Successful Output",
}
PROPERTY_OFFSETS = {
    "last_run": 1,
    "start": 2,
    "finish": 3,
    "elapsed": 4,
    "user": 5,
    "status": 6,
    "input_template": 7,
    "last_output": 9,
}
VALID_OUTPUT_TYPES = {"detailed", "GrossandNNEG", "total"}
ALLOWED_TEMPLATE_SUFFIXES = {".xlsb", ".xlsm", ".xlsx"}
XL_CELL_TYPE_FORMULAS = -4123


class PropertyStageOneError(shared.BondsLiteError):
    """An actionable Property Stage 1 input, workbook or audit failure."""


@dataclass(frozen=True)
class PropertyStageOneInputs:
    """Validated Draft controls and the selected ScenFlipper values.

    ``values`` holds rows 1-5 of the Draft ScenFlipper followed by the Base
    block and the stress block selected by the Interest and Property flags.
    """

    caller_path: Path
    entity: str
    template_path: Path
    valuation_date: datetime
    output_suffix: str
    values: list[list[object]]
    rows: int
    columns: int
    headers: tuple[object, ...]
    stress_block: str


@dataclass(frozen=True)
class TemplateSnapshot:
    """Template invariants that must remain unchanged during workbook editing."""

    rows: int
    columns: int
    prm_names: tuple[tuple[str, str], ...]
    link_sources: tuple[str, ...]
    dependent_formulas: tuple[tuple[str, str, str], ...]


def _read_flag(caller: Any, name: str, kind: str) -> bool:
    """Read a named I. Control flag as an actual Excel Boolean."""
    value = shared._control_cell(caller, name, f"{kind} flag name").value
    if not isinstance(value, bool):
        raise PropertyStageOneError(
            f"{kind} flag {name} must be TRUE or FALSE, not {value!r}"
        )
    return value


def read_property_flag(caller: Any) -> bool:
    """Read the Property standalone flag as an actual Excel Boolean."""
    return _read_flag(caller, PROPERTY_FLAG, "Property")


def read_interest_flag(caller: Any) -> bool:
    """Read the Interest standalone flag as an actual Excel Boolean."""
    return _read_flag(caller, INTEREST_FLAG, "Interest")


def read_stress_flags(caller: Any) -> tuple[bool, bool]:
    """Return the ``(interest, property)`` standalone flags."""
    return read_interest_flag(caller), read_property_flag(caller)


def require_property_selected(flag: bool) -> None:
    """Stop the Property workstream when its standalone flag is disabled."""
    if flag is not True:
        raise PropertyStageOneError(
            f"Property stress is not selected: set {PROPERTY_FLAG} to TRUE"
        )


def require_stress_selected(interest: bool, property_: bool) -> str:
    """Return the ScenFlipper stress block selected by the flags.

    Any non-empty flag combination is supported so that workstreams other
    than Property (e.g. a future Interest-only run) can reuse the selection.

    Raises:
        PropertyStageOneError: If neither flag is TRUE.
    """
    if interest is not True and property_ is not True:
        raise PropertyStageOneError(
            "No ScenFlipper stress is selected: set "
            f"{INTEREST_FLAG} and/or {PROPERTY_FLAG} to TRUE"
        )
    return STRESS_BLOCKS[(interest is True, property_ is True)]


def _named_value(caller: Any, name: str, kind: str) -> object:
    """Read one named scalar control from the Draft."""
    return shared._control_cell(caller, name, kind).value


def _resolve_template_path(value: object, caller_path: Path) -> Path:
    """Resolve and validate the Draft-relative LTM Flash template path."""
    if not isinstance(value, str) or not value.strip():
        raise PropertyStageOneError(f"{TEMPLATE_NAME} path is blank")
    path = Path(os.path.expandvars(value.strip())).expanduser()
    if not path.is_absolute():
        path = caller_path.parent / path
    try:
        resolved = path.resolve(strict=True)
    except OSError as error:
        raise PropertyStageOneError(
            f"{TEMPLATE_NAME} workbook does not exist: {path}"
        ) from error
    if not resolved.is_file():
        raise PropertyStageOneError(f"{TEMPLATE_NAME} path is not a file: {resolved}")
    if resolved.suffix.casefold() not in ALLOWED_TEMPLATE_SUFFIXES:
        raise PropertyStageOneError(
            f"{TEMPLATE_NAME} must use .xlsb, .xlsm or .xlsx: {resolved.name}"
        )
    return resolved


def _as_grid(value: object) -> list[list[object]]:
    """Normalise an Excel scalar or array value to a two-dimensional list."""
    if not isinstance(value, (tuple, list)):
        return [[value]]
    if value and not isinstance(value[0], (tuple, list)):
        return [list(value)]
    return [list(row) for row in value]


def _used_dimensions(sheet: Any) -> tuple[int, int, int, int]:
    """Return the one-based inclusive bounds of a sheet's used range."""
    used = sheet.api.UsedRange
    top = int(used.Row)
    left = int(used.Column)
    rows = int(used.Rows.Count)
    columns = int(used.Columns.Count)
    return top, left, top + rows - 1, left + columns - 1


def _read_scenflipper(
    sheet: Any,
) -> tuple[list[list[object]], list[list[object]], int, int]:
    """Read ScenFlipper values and formulas with their used-range origin."""
    top, left, bottom, right = _used_dimensions(sheet)
    region = sheet.range((top, left), (bottom, right))
    values = _as_grid(region.options(err_to_str=True).value)
    formulas = _normalise_formula_grid(_as_grid(region.formula))
    return values, formulas, top, left


def _is_excel_error(value: object) -> bool:
    """Identify a string representation returned for an Excel error value."""
    return isinstance(value, str) and value.startswith("#")


def _whole_number(value: object) -> bool:
    """Return whether an Excel numeric control is a finite whole number."""
    return (
        isinstance(value, Real)
        and not isinstance(value, bool)
        and float(value).is_integer()
    )


def _scenario_blocks(
    padded_values: list[list[object]],
) -> dict[str, list[int]]:
    """Group contiguous scenario rows by their ``scen desc`` label.

    Returns:
        Zero-based row indices for each block label, in sheet order.
    """
    blocks: dict[str, list[int]] = {}
    previous: str | None = None
    for row_index in range(5, len(padded_values)):
        scenario = padded_values[row_index][0]
        if scenario is None or scenario == "":
            break
        label = padded_values[row_index][SCEN_DESC_COLUMN - 1]
        label = label.strip() if isinstance(label, str) else ""
        valid = {BASE_BLOCK, *STRESS_BLOCKS.values()}
        if label not in valid:
            raise PropertyStageOneError(
                f"Draft I. ScenFlipper row {row_index + 1} scen desc must be one of "
                f"{', '.join(sorted(valid))}, not {label!r}"
            )
        if label != previous and label in blocks:
            raise PropertyStageOneError(
                f"Draft I. ScenFlipper {label} block rows must be contiguous"
            )
        blocks.setdefault(label, []).append(row_index)
        previous = label
    return blocks


def _require_stressed_rows(
    padded_values: list[list[object]],
    rows: list[int],
    columns: tuple[int, ...],
    kind: str,
) -> None:
    """Require every selected row to have a non-zero value in ``columns``."""
    letters = ", ".join(chr(64 + column) for column in columns)
    for row_index in rows:
        cells = [padded_values[row_index][column - 1] for column in columns]
        for column, value in zip(columns, cells, strict=True):
            if value is not None and (
                not isinstance(value, Real) or isinstance(value, bool)
            ):
                raise PropertyStageOneError(
                    f"Draft I. ScenFlipper {kind} stress at row {row_index + 1}, "
                    f"column {chr(64 + column)} must be numeric, not {value!r}"
                )
        if all(value is None or value == 0 for value in cells):
            raise PropertyStageOneError(
                f"{kind} stress is selected but Draft I. ScenFlipper row "
                f"{row_index + 1} has no non-zero value in columns {letters}"
            )


def _require_unstressed_rows(
    padded_values: list[list[object]],
    rows: list[int],
    columns: tuple[int, ...],
    block: str,
) -> None:
    """Require every row in ``rows`` to be blank or zero in ``columns``."""
    for row_index in rows:
        for column in columns:
            value = padded_values[row_index][column - 1]
            if value is not None and value != 0:
                raise PropertyStageOneError(
                    f"Draft I. ScenFlipper {block} block must not be stressed in "
                    f"column {chr(64 + column)}: row {row_index + 1} is {value!r}"
                )


def validate_scenflipper(
    values: list[list[object]], interest: bool, property_: bool
) -> tuple[tuple[object, ...], int, list[list[object]], str]:
    """Validate a Draft ScenFlipper grid and select the blocks to copy.

    Args:
        values: Draft ScenFlipper values from A1. Rows 1-3 hold the controls,
            row 5 the headers and rows 6 onwards contiguous scenario blocks
            labelled in column J (``scen desc``).
        interest: ``SST_Interest_standalone`` flag.
        property_: ``SST_Property_standalone`` flag.

    Returns:
        Row 5 headers, grid width, the output grid (rows 1-5, the Base block
        and the selected stress block) and the selected stress block label.
    """
    stress_block = require_stress_selected(interest, property_)
    if not values or len(values) < 5 or not values[0]:
        raise PropertyStageOneError("Draft I. ScenFlipper used range is empty")
    width = max(len(row) for row in values)
    padded_values = [row + [None] * (width - len(row)) for row in values]

    headers = tuple(padded_values[4])
    if any(not isinstance(header, str) or not header.strip() for header in headers):
        raise PropertyStageOneError(
            "Draft I. ScenFlipper row 5 headers must be non-blank"
        )
    if len({header.casefold() for header in headers}) != len(headers):
        raise PropertyStageOneError("Draft I. ScenFlipper row 5 headers must be unique")

    output_type = padded_values[0][1] if width > 1 else None
    scenario_count = padded_values[1][1] if width > 1 and len(values) > 1 else None
    current_scenario = padded_values[2][1] if width > 1 and len(values) > 2 else None
    if output_type not in VALID_OUTPUT_TYPES:
        raise PropertyStageOneError(
            "Draft I. ScenFlipper B1 must be detailed, GrossandNNEG or total"
        )
    if not _whole_number(scenario_count) or scenario_count <= 0:
        raise PropertyStageOneError(
            "Draft I. ScenFlipper B2 must be a positive whole number"
        )
    if not _whole_number(current_scenario) or not (
        1 <= current_scenario <= scenario_count
    ):
        raise PropertyStageOneError(
            "Draft I. ScenFlipper B3 must be a whole number from 1 to B2"
        )

    blank_seen = False
    for row_index in range(5, len(padded_values)):
        scenario = padded_values[row_index][0]
        if scenario is None or scenario == "":
            blank_seen = True
        elif blank_seen:
            raise PropertyStageOneError(
                "Draft I. ScenFlipper scenarios in column A must be contiguous"
            )
    if width < max(PROPERTY_STRESS_COLUMNS) or not (
        str(headers[SCEN_DESC_COLUMN - 1]).strip().casefold() == SCEN_DESC_HEADER
    ):
        raise PropertyStageOneError(
            f"Draft I. ScenFlipper column J header must be {SCEN_DESC_HEADER!r}"
        )
    blocks = _scenario_blocks(padded_values)
    if next(iter(blocks), None) != BASE_BLOCK:
        raise PropertyStageOneError(
            f"Draft I. ScenFlipper first scenario block must be {BASE_BLOCK}"
        )
    if stress_block not in blocks:
        raise PropertyStageOneError(
            f"Draft I. ScenFlipper has no {stress_block} block in column J"
        )
    base_rows = blocks[BASE_BLOCK]
    stress_rows = blocks[stress_block]
    if len(base_rows) != len(stress_rows):
        raise PropertyStageOneError(
            f"Draft I. ScenFlipper {BASE_BLOCK} and {stress_block} blocks must "
            f"have the same number of rows ({len(base_rows)} and {len(stress_rows)})"
        )
    all_stress_columns = INTEREST_STRESS_COLUMNS + PROPERTY_STRESS_COLUMNS
    _require_unstressed_rows(padded_values, base_rows, all_stress_columns, BASE_BLOCK)
    if interest:
        _require_stressed_rows(
            padded_values, stress_rows, INTEREST_STRESS_COLUMNS, "Interest"
        )
    else:
        _require_unstressed_rows(
            padded_values, stress_rows, INTEREST_STRESS_COLUMNS, stress_block
        )
    if property_:
        _require_stressed_rows(
            padded_values, stress_rows, PROPERTY_STRESS_COLUMNS, "Property"
        )
    else:
        _require_unstressed_rows(
            padded_values, stress_rows, PROPERTY_STRESS_COLUMNS, stress_block
        )

    selected = [padded_values[index] for index in range(5)] + [
        padded_values[index] for index in base_rows + stress_rows
    ]
    for row_index, row in enumerate(selected, start=1):
        for column_index, value in enumerate(row, start=1):
            if _is_excel_error(value):
                raise PropertyStageOneError(
                    "Draft I. ScenFlipper selected rows contain an Excel error at "
                    f"output row {row_index}, column {column_index}: {value}"
                )
    selected_count = len(base_rows) + len(stress_rows)
    if scenario_count != selected_count:
        raise PropertyStageOneError(
            "Draft I. ScenFlipper B2 must equal the number of Base and "
            f"{stress_block} rows (B2={scenario_count}, rows={selected_count})"
        )
    numbers = [row[0] for row in selected[5:]]
    if numbers != list(range(1, selected_count + 1)):
        raise PropertyStageOneError(
            f"Draft I. ScenFlipper column A of the {BASE_BLOCK} and {stress_block} "
            f"blocks must number 1 to {selected_count} in order"
        )
    return headers, width, selected, stress_block


def _date_value(value: object) -> datetime:
    """Convert an Excel date value to a timezone-naive datetime."""
    if isinstance(value, datetime):
        return value
    if isinstance(value, date):
        return datetime.combine(value, datetime.min.time())
    raise PropertyStageOneError("I_Valuation_Date must contain an Excel date value")


def read_property_stage_one_inputs(caller: Any) -> PropertyStageOneInputs:
    """Read the stress flags, template, valuation date and scenario grid."""
    caller_path = shared.resolve_caller_path(caller)
    interest, property_ = read_stress_flags(caller)
    # The Property workstream only runs Property stresses; Interest-only
    # selection is kept in validate_scenflipper for the Interest workstream.
    require_property_selected(property_)
    entity = shared.read_entity(caller)
    template_path = _resolve_template_path(
        _named_value(caller, TEMPLATE_NAME, "LTM Flash template name"),
        caller_path,
    )
    valuation_date = _date_value(
        _named_value(caller, "I_Valuation_Date", "Valuation date name")
    )
    output_suffix = shared.read_output_suffix(caller)
    replaced_folder = (
        caller_path.parent
        / f"{shared.OUTPUT_ROOT_PREFIX}{output_suffix}"
        / shared.PROPERTY_OUTPUT_FOLDER
    ).resolve()
    try:
        template_path.relative_to(replaced_folder)
    except ValueError:
        pass
    else:
        raise PropertyStageOneError(
            "LTM Flash template cannot be inside the Property folder replaced "
            f"by this run: {template_path}"
        )
    try:
        sheet = caller.sheets[SCENFLIPPER_SHEET]
    except Exception as error:
        raise PropertyStageOneError(
            f"Draft is missing the {SCENFLIPPER_SHEET} sheet"
        ) from error
    # Values are pasted rather than formulas, so refresh any manual-mode results.
    caller.app.api.Calculate()
    values, _, top, left = _read_scenflipper(sheet)
    if (top, left) != (1, 1):
        raise PropertyStageOneError("Draft I. ScenFlipper used range must start at A1")
    headers, columns, selected, stress_block = validate_scenflipper(
        values, interest, property_
    )
    logger.info("Property Stage 1 copies Base and %s blocks", stress_block)
    return PropertyStageOneInputs(
        caller_path,
        entity,
        template_path,
        valuation_date,
        output_suffix,
        selected,
        len(selected),
        columns,
        headers,
        stress_block,
    )


def _property_output_path(caller_path: Path, folder: Path, filename: str) -> Path:
    """Build a final output path after validating the Property run folder."""
    suffix = shared.output_root_suffix(folder.parent.name)
    if (
        suffix is None
        or folder.name != shared.PROPERTY_OUTPUT_FOLDER
        or folder.resolve().parent.parent != caller_path.parent.resolve()
    ):
        raise PropertyStageOneError(
            f"Property output folder is not beside Draft: {folder}"
        )
    return folder / filename


def _portable_path(caller_path: Path, path: Path) -> str:
    """Return a path relative to the Draft folder for workbook audit links."""
    try:
        return path.resolve().relative_to(caller_path.parent.resolve()).as_posix()
    except ValueError:
        return os.path.relpath(path.resolve(), caller_path.parent.resolve()).replace(
            os.sep, "/"
        )


def _audit_region(caller: Any, name: str) -> Any:
    """Resolve a local Property audit range and normalise its error type."""
    try:
        return shared._named_audit_region(caller, name)
    except shared.BondsLiteError as error:
        raise PropertyStageOneError(str(error)) from error


def _property_field_cells(caller: Any) -> dict[str, Any]:
    """Resolve audit fields to the first cell of each merged named range."""
    cells = {}
    for field, name in PROPERTY_FIELD_NAMES.items():
        region = _audit_region(caller, name)
        merged = region.api.Cells(1, 1).MergeArea
        cells[field] = region.sheet.range((merged.Row, merged.Column))
    return cells


def require_property_stage_one_panel(caller: Any) -> None:
    """Require a writable, non-overlapping named Property Stage 1 audit panel."""
    if caller.api.ReadOnly:
        raise PropertyStageOneError(
            "Draft is read-only; its Property Stage 1 audit cannot be saved"
        )
    sheet = caller.sheets["I. Control"]
    if sheet.api.ProtectContents:
        raise PropertyStageOneError(
            "I. Control is protected; its Property Stage 1 audit cannot be updated"
        )
    panel = _audit_region(caller, PROPERTY_PANEL_NAME)
    bounds = (
        panel.row,
        panel.column,
        panel.row + panel.rows.count - 1,
        panel.column + panel.columns.count - 1,
    )
    occupied: list[tuple[int, int, int, int]] = []
    for field, name in PROPERTY_FIELD_NAMES.items():
        region = _audit_region(caller, name)
        first = region.api.Cells(1, 1).MergeArea
        field_bounds = (
            int(first.Row),
            int(first.Column),
            int(first.Row + first.Rows.Count - 1),
            int(first.Column + first.Columns.Count - 1),
        )
        if not (
            bounds[0] <= field_bounds[0] <= field_bounds[2] <= bounds[2]
            and bounds[1] <= field_bounds[1] <= field_bounds[3] <= bounds[3]
        ):
            raise PropertyStageOneError(
                f"Property Stage 1 audit field is outside its panel: {name}"
            )
        if any(
            field_bounds[0] <= old[2]
            and field_bounds[2] >= old[0]
            and field_bounds[1] <= old[3]
            and field_bounds[3] >= old[1]
            for old in occupied
        ):
            raise PropertyStageOneError(
                f"Property Stage 1 audit fields overlap: {field}"
            )
        occupied.append(field_bounds)


def _read_property_stage_one_fields(caller: Any) -> dict[str, object]:
    """Read the current Property audit values after validating the panel."""
    require_property_stage_one_panel(caller)
    return {field: cell.value for field, cell in _property_field_cells(caller).items()}


def write_property_stage_one_audit(caller: Any, fields: dict[str, object]) -> None:
    """Write named audit values and save the Draft after each state change."""
    require_property_stage_one_panel(caller)
    if set(fields) - PROPERTY_FIELD_NAMES.keys():
        raise PropertyStageOneError("Unknown Property Stage 1 audit field")
    cells = _property_field_cells(caller)
    events = caller.app.api.EnableEvents
    try:
        caller.app.api.EnableEvents = False
        for field, value in fields.items():
            cell = cells[field]
            if field in ("input_template", "last_output"):
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
            raise PropertyStageOneError("Draft's Property Stage 1 audit was not saved")
    finally:
        caller.app.api.EnableEvents = events


def _normalise_link_sources(book: Any) -> tuple[str, ...]:
    """Return the workbook's external link sources in their stored order."""
    sources = book.api.LinkSources(1)
    if sources is None:
        return ()
    if isinstance(sources, str):
        sources = (sources,)
    return tuple(str(source) for source in sources)


def _prm_names(book: Any) -> tuple[tuple[str, str], ...]:
    """Snapshot all workbook names beginning with ``prm_``."""
    result = []
    for name in book.names:
        full_name = str(getattr(name, "name", getattr(name, "Name", "")))
        key = full_name.split("!")[-1]
        if key.casefold().startswith("prm_"):
            refers_to = str(getattr(name, "refers_to", getattr(name, "RefersTo", "")))
            result.append((key.casefold(), refers_to))
    return tuple(sorted(result))


def _scenflipper_formulas(book: Any) -> tuple[tuple[str, str, str], ...]:
    """Capture formulas that refer to ScenFlipper with their sheet and address."""
    references = []
    for sheet in book.sheets:
        used = sheet.api.UsedRange
        if used.HasFormula is False:
            continue
        try:
            areas = used.SpecialCells(XL_CELL_TYPE_FORMULAS).Areas
        except Exception as error:
            raise PropertyStageOneError(
                f"Cannot verify formulas on LTM Flash sheet {sheet.name}"
            ) from error
        for area_index in range(1, int(areas.Count) + 1):
            area = areas.Item(area_index)
            formulas = _as_grid(area.Formula)
            for row_offset, row in enumerate(formulas):
                for column_offset, formula in enumerate(row):
                    if _formula_references_scenflipper(formula):
                        if "#REF!" in str(formula).upper():
                            raise PropertyStageOneError(
                                "LTM Flash contains a #REF! formula referencing "
                                "I. ScenFlipper"
                            )
                        references.append(
                            (
                                sheet.name,
                                sheet.range(
                                    (
                                        int(area.Row) + row_offset,
                                        int(area.Column) + column_offset,
                                    )
                                ).get_address(),
                                str(formula),
                            )
                        )
    return tuple(sorted(references))


def _snapshot_template(
    book: Any, draft_headers: tuple[object, ...]
) -> TemplateSnapshot:
    """Validate the template sheet and capture names and external links."""
    try:
        sheet = book.sheets[SCENFLIPPER_SHEET]
    except Exception as error:
        raise PropertyStageOneError(
            f"LTM Flash template is missing the {SCENFLIPPER_SHEET} sheet"
        ) from error
    values, _, top, left = _read_scenflipper(sheet)
    if (top, left) != (1, 1):
        raise PropertyStageOneError(
            "LTM Flash template I. ScenFlipper used range must start at A1"
        )
    headers = tuple(values[4]) if len(values) >= 5 else ()
    if headers != draft_headers:
        raise PropertyStageOneError(
            "LTM Flash template and Draft I. ScenFlipper row 5 headers do not match"
        )
    _valuation_date_range(book)
    return TemplateSnapshot(
        len(values),
        max((len(row) for row in values), default=0),
        _prm_names(book),
        _normalise_link_sources(book),
        _scenflipper_formulas(book),
    )


def _valuation_date_range(book: Any) -> Any:
    """Resolve the LTM Flash ValDate name as a single cell."""
    try:
        region = book.names["ValDate"].refers_to_range
        if int(region.api.Areas.Count) != 1 or int(region.size) != 1:
            raise PropertyStageOneError("LTM Flash ValDate must identify one cell")
        return region
    except PropertyStageOneError:
        raise
    except Exception as error:
        raise PropertyStageOneError(
            "LTM Flash template is missing the ValDate name"
        ) from error


def set_valuation_date(book: Any, valuation_date: datetime) -> None:
    """Set the named LTM Flash valuation date and verify Excel read-back."""
    cell = _valuation_date_range(book)
    old_value = cell.value
    cell.value = valuation_date
    actual = cell.value
    if isinstance(actual, date) and not isinstance(actual, datetime):
        actual = datetime.combine(actual, datetime.min.time())
    if actual != valuation_date:
        raise PropertyStageOneError(
            f"LTM Flash ValDate did not retain the Draft date: "
            f"expected {valuation_date!r}, received {actual!r}"
        )
    logger.info("LTM Flash ValDate updated from %r to %r", old_value, actual)


def _grid_values(sheet: Any, rows: int, columns: int) -> list[list[object]]:
    """Read cell values from A1 through the requested grid dimensions."""
    if rows == 0 or columns == 0:
        return []
    region = sheet.range((1, 1), (rows, columns))
    return _as_grid(region.options(err_to_str=True).value)


def _grid_formulas(sheet: Any, rows: int, columns: int) -> list[list[object]]:
    """Read formulas from A1 through the requested grid dimensions."""
    if rows == 0 or columns == 0:
        return []
    formulas = _as_grid(sheet.range((1, 1), (rows, columns)).formula)
    return _normalise_formula_grid(formulas)


def _normalise_formula_grid(formulas: list[list[object]]) -> list[list[object]]:
    """Treat Excel's empty formula strings as blank cells."""
    return [[None if formula == "" else formula for formula in row] for row in formulas]


def _padded_grid(
    grid: list[list[object]], rows: int, columns: int
) -> list[list[object]]:
    """Pad a grid with empty cells to the requested row and column counts."""
    return [
        (grid[index] if index < len(grid) else [])[:columns]
        + [None] * max(0, columns - (len(grid[index]) if index < len(grid) else 0))
        for index in range(rows)
    ]


def _first_grid_difference(
    actual: list[list[object]], expected: list[list[object]]
) -> tuple[int, int, object, object] | None:
    """Return the first one-based row/column and mismatched cell values."""
    for row_index, (actual_row, expected_row) in enumerate(
        zip(actual, expected, strict=True), start=1
    ):
        for column_index, (actual_value, expected_value) in enumerate(
            zip(actual_row, expected_row, strict=True), start=1
        ):
            if actual_value != expected_value:
                return row_index, column_index, actual_value, expected_value
    return None


def _formula_references_scenflipper(formula: object) -> bool:
    """Return whether a formula text references the ScenFlipper sheet."""
    return (
        isinstance(formula, str) and SCENFLIPPER_SHEET.casefold() in formula.casefold()
    )


def verify_ltm_flash(
    book: Any,
    expected: PropertyStageOneInputs,
    template: TemplateSnapshot,
    valuation_date: datetime,
    draft_path: Path,
) -> None:
    """Verify copied scenario cells and preservation of template dependencies."""
    sheet = book.sheets[SCENFLIPPER_SHEET]
    rows = max(expected.rows, template.rows)
    columns = max(expected.columns, template.columns)
    actual_values = _padded_grid(_grid_values(sheet, rows, columns), rows, columns)
    expected_values = _padded_grid(expected.values, rows, columns)
    actual_formulas = _padded_grid(_grid_formulas(sheet, rows, columns), rows, columns)
    difference = _first_grid_difference(actual_values, expected_values)
    if difference is not None:
        row, column, actual, expected_value = difference
        raise PropertyStageOneError(
            "LTM Flash I. ScenFlipper value mismatch at "
            f"row {row}, column {column}: expected {expected_value!r}, "
            f"received {actual!r}"
        )
    for row, formula_row in enumerate(actual_formulas, start=1):
        for column, formula in enumerate(formula_row, start=1):
            if isinstance(formula, str) and formula.startswith("="):
                raise PropertyStageOneError(
                    "LTM Flash I. ScenFlipper must contain values only; formula at "
                    f"row {row}, column {column}: {formula!r}"
                )
    if _prm_names(book) != template.prm_names:
        raise PropertyStageOneError("LTM Flash prm_* names changed during preparation")
    if _scenflipper_formulas(book) != template.dependent_formulas:
        raise PropertyStageOneError(
            "LTM Flash formulas referencing I. ScenFlipper changed during preparation"
        )

    try:
        dependent_formula = book.sheets["I. Inputs"].range("C1").formula
    except Exception as error:
        raise PropertyStageOneError(
            "LTM Flash I. Inputs!C1 cannot be checked"
        ) from error
    if (
        not _formula_references_scenflipper(dependent_formula)
        or "#REF!" in str(dependent_formula).upper()
    ):
        raise PropertyStageOneError(
            "LTM Flash I. Inputs!C1 no longer references I. ScenFlipper correctly"
        )
    links = _normalise_link_sources(book)
    if links != template.link_sources:
        raise PropertyStageOneError(
            "LTM Flash external link sources changed during preparation"
        )
    draft_identity = os.path.normcase(str(draft_path.resolve()))
    if any(
        draft_identity in os.path.normcase(source)
        or source.replace("\\", "/").split("/")[-1].split("?")[0].casefold()
        == draft_path.name.casefold()
        for source in links
    ):
        raise PropertyStageOneError(
            "LTM Flash must not retain an external link to the Draft workbook"
        )
    actual_date = _valuation_date_range(book).value
    if isinstance(actual_date, date) and not isinstance(actual_date, datetime):
        actual_date = datetime.combine(actual_date, datetime.min.time())
    if actual_date != valuation_date:
        raise PropertyStageOneError(
            f"LTM Flash ValDate does not match the Draft date: {actual_date!r}"
        )


def replace_scenflipper(
    target_sheet: Any,
    values: list[list[object]],
    columns: int,
    target_rows: int,
    target_columns: int,
) -> None:
    """Overwrite the template sheet contents in place with selected values.

    Values only are written, keeping the template's formatting and avoiding
    names or external links to the Draft being carried across by formulas.
    """
    rows = max(len(values), target_rows)
    width = max(columns, target_columns)
    target_sheet.range((1, 1), (rows, width)).api.ClearContents()
    grid = _padded_grid(values, len(values), columns)
    target_sheet.range((1, 1), (len(values), columns)).value = grid


@contextmanager
def _manual_calculation(application: Any) -> Iterator[None]:
    """Temporarily put the caller's Excel application into manual calculation."""
    original = application.api.Calculation
    try:
        application.api.Calculation = -4135
        yield
    finally:
        application.api.Calculation = original


def _write_failure_audit(
    caller: Any,
    started: datetime,
    clock: float,
    previous_output: object,
    status: str,
    template_path: str | None = None,
) -> None:
    """Write a terminal failure status while retaining the last successful output."""
    values: dict[str, object] = {
        "last_run": started.replace(hour=0, minute=0, second=0, microsecond=0),
        "start": started,
        "finish": datetime.now(),
        "elapsed": (time.perf_counter() - clock) / 86400,
        "user": getpass.getuser(),
        "status": status,
        "last_output": previous_output,
    }
    if template_path is not None:
        values["input_template"] = template_path
    write_property_stage_one_audit(caller, values)


def run_property_stage_one(caller: Any) -> Path | None:
    """Prepare a verified LTM Flash workbook in the current Property folder."""
    caller_path = shared.resolve_caller_path(caller)
    with shared.workbook_run_guard(
        caller_path,
        conflict_message="A stress run is already in progress for this workbook",
    ):
        previous_output = _read_property_stage_one_fields(caller)["last_output"]
        started = datetime.now()
        clock = time.perf_counter()
        template_portable: str | None = None
        try:
            inputs = read_property_stage_one_inputs(caller)
            template_portable = _portable_path(caller_path, inputs.template_path)
        except BaseException as error:
            status = "Interrupted" if isinstance(error, KeyboardInterrupt) else "Failed"
            try:
                _write_failure_audit(caller, started, clock, previous_output, status)
            except Exception as audit_error:
                logger.error(
                    "Property Stage 1 input validation failed and its audit "
                    "could not be saved: %s",
                    audit_error,
                )
            raise

        try:
            write_property_stage_one_audit(
                caller,
                {
                    "last_run": started.replace(
                        hour=0, minute=0, second=0, microsecond=0
                    ),
                    "start": started,
                    "finish": None,
                    "elapsed": None,
                    "user": getpass.getuser(),
                    "status": "Running",
                    "input_template": template_portable,
                    "last_output": previous_output,
                },
            )
        except BaseException as error:
            status = "Interrupted" if isinstance(error, KeyboardInterrupt) else "Failed"
            try:
                _write_failure_audit(
                    caller,
                    started,
                    clock,
                    previous_output,
                    status,
                    template_portable,
                )
            except Exception as audit_error:
                logger.error(
                    "Property Stage 1 could not start and its failure audit "
                    "could not be saved: %s",
                    audit_error,
                )
            raise

        output_folder: Path | None = None
        root_created = False
        backup: Path | None = None
        temporary_folder: Path | None = None
        output_path: Path | None = None
        committed = False
        try:
            output_folder, root_created, backup = allocate_stress_folder(
                caller_path.parent,
                inputs.output_suffix,
                shared.PROPERTY_OUTPUT_FOLDER,
            )
            template_name = shared.entity_output_name(
                inputs.entity, inputs.template_path, inputs.output_suffix
            )
            output_path = _property_output_path(
                caller_path, output_folder, template_name
            )
            if output_path.exists():
                raise PropertyStageOneError(
                    f"Property output already exists after folder allocation: {output_path}"
                )
            temporary_folder = Path(
                tempfile.mkdtemp(prefix="~property-stage1-", dir=output_folder)
            )
            working_path = temporary_folder / template_name
            shutil.copy2(inputs.template_path, working_path)
            if working_path.stat().st_size != inputs.template_path.stat().st_size:
                raise PropertyStageOneError(
                    "LTM Flash template copy has an unexpected file size"
                )

            application = caller.app
            with shared.excel_state(application), _manual_calculation(application):
                with shared.owned_book(
                    application, working_path, read_only=False
                ) as book:
                    template = _snapshot_template(book, inputs.headers)
                    target_sheet = book.sheets[SCENFLIPPER_SHEET]
                    replace_scenflipper(
                        target_sheet,
                        inputs.values,
                        inputs.columns,
                        template.rows,
                        template.columns,
                    )
                    set_valuation_date(book, inputs.valuation_date)
                    verify_ltm_flash(
                        book,
                        inputs,
                        template,
                        inputs.valuation_date,
                        caller_path,
                    )
                    book.save()
                    if not book.api.Saved:
                        raise PropertyStageOneError(
                            "LTM Flash working copy was not saved"
                        )

            working_path.rename(output_path)
            portable_output = _portable_path(caller_path, output_path)
            write_property_stage_one_audit(
                caller,
                {
                    "finish": datetime.now(),
                    "elapsed": (time.perf_counter() - clock) / 86400,
                    "status": "Succeeded",
                    "input_template": template_portable,
                    "last_output": portable_output,
                },
            )
            committed = True
            discard_replaced_stress_folder(backup)
            logger.info("Property Stage 1 succeeded; saved LTM Flash: %s", output_path)
            return output_path
        except BaseException as error:
            if output_folder is not None and not committed:
                try:
                    remove_incomplete_stress_folder(output_folder, root_created, backup)
                except Exception as cleanup_error:
                    logger.error(
                        "Property Stage 1 failed and its output folder could not "
                        "be restored: %s",
                        cleanup_error,
                    )
            status = "Interrupted" if isinstance(error, KeyboardInterrupt) else "Failed"
            try:
                _write_failure_audit(
                    caller,
                    started,
                    clock,
                    previous_output,
                    status,
                    template_portable,
                )
            except Exception as audit_error:
                logger.error(
                    "Property Stage 1 failed and its audit could not be saved: %s",
                    audit_error,
                )
            raise
        finally:
            if temporary_folder is not None and temporary_folder.exists():
                try:
                    shutil.rmtree(
                        temporary_folder,
                        onexc=shared._retry_readonly_backup_removal,
                    )
                except OSError as cleanup_error:
                    logger.warning(
                        "Could not remove Property Stage 1 staging folder %s: %s",
                        temporary_folder,
                        cleanup_error,
                    )


def _name_keys(book: Any) -> set[str]:
    """Return the unqualified names in a workbook."""
    return {
        str(getattr(name, "name", getattr(name, "Name", ""))).split("!")[-1]
        for name in book.names
    }


def _flag_control(book: Any) -> None:
    """Preserve or add the Property flag name after validating its label."""
    names = _name_keys(book)
    if PROPERTY_FLAG in names:
        shared._control_cell(book, PROPERTY_FLAG, "Property flag name")
        return
    sheet = book.sheets["I. Control"]
    label = sheet.range("C18").value
    if not isinstance(label, str) or label.strip() != "Property Value":
        raise PropertyStageOneError(
            "I. Control!C18 must be labelled Property Value; no controls were installed"
        )
    book.names.add(PROPERTY_FLAG, "='I. Control'!$D$18")


def _entity_control(book: Any) -> None:
    """Ensure SST_Entity names the entity cell on I. Control."""
    if "SST_Entity" in _name_keys(book):
        try:
            cell = shared._control_cell(book, "SST_Entity", "Entity name")
        except shared.BondsLiteError as error:
            raise PropertyStageOneError(str(error)) from error
        if cell.row != 9 or cell.column != 4:
            raise PropertyStageOneError(
                "SST_Entity must refer to I. Control!$D$9; no controls were installed"
            )
        return
    book.names.add("SST_Entity", "='I. Control'!$D$9")


def _overlaps(
    first: tuple[float, float, float, float], second: tuple[float, float, float, float]
) -> bool:
    """Return whether two left/top/width/height rectangles intersect."""
    return (
        first[0] < second[0] + second[2]
        and first[0] + first[2] > second[0]
        and first[1] < second[1] + second[3]
        and first[1] + first[3] > second[1]
    )


def _install_property_panel(book: Any) -> None:
    """Create the named Property audit panel to the right of Credit Stage 1."""
    existing = _name_keys(book)
    expected = {PROPERTY_PANEL_NAME, *PROPERTY_FIELD_NAMES.values()}
    overlap = expected.intersection(existing)
    if overlap:
        if not expected.issubset(existing):
            raise PropertyStageOneError(
                "Property Stage 1 audit names are incomplete; installation was not changed"
            )
        require_property_stage_one_panel(book)
        return

    sheet = book.sheets["I. Control"]
    try:
        control_panel = _audit_region(book, "SST_Stage1_Panel")
        button = sheet.api.Shapes.Item("btnGenerateBondsLiteInput")
    except Exception as error:
        raise PropertyStageOneError(
            "Cannot locate the Credit Stage 1 panel and button; no controls were installed"
        ) from error
    button_row = int(button.TopLeftCell.Row)
    button_bottom_row = int(button.BottomRightCell.Row)
    target_left_column = control_panel.column + control_panel.columns.count + 1
    button_left = sheet.range((button_row, target_left_column)).left
    target_top = control_panel.top + button.Top - control_panel.top
    button_width = float(button.Width)
    button_height = float(button.Height)
    last_button_column = target_left_column
    while (
        sheet.range((button_row, last_button_column)).left < button_left + button_width
    ):
        last_button_column += 1
    panel_left_column = last_button_column
    row_count = 11
    column_count = 8
    panel = sheet.range(
        (control_panel.row, panel_left_column),
        (control_panel.row + row_count - 1, panel_left_column + column_count - 1),
    )
    button_cells = sheet.range(
        (button_row, target_left_column),
        (button_bottom_row, last_button_column - 1),
    )
    gap = sheet.range(
        (control_panel.row, control_panel.column + control_panel.columns.count),
        (
            control_panel.row + row_count - 1,
            control_panel.column + control_panel.columns.count,
        ),
    )
    for description, region in (
        ("Property Stage 1 audit area", panel),
        ("Prepare LTM Flash button area", button_cells),
        ("gap beside the Credit panel", gap),
    ):
        if (
            book.app.api.WorksheetFunction.CountA(region.api)
            or region.api.MergeCells is not False
        ):
            raise PropertyStageOneError(
                f"Proposed {description} is occupied or merged; nothing was overwritten"
            )
    new_button_box = (
        button_left,
        target_top,
        button_width,
        button_height,
    )
    panel_box = (panel.left, panel.top, panel.width, panel.height)
    gap_box = (gap.left, gap.top, gap.width, gap.height)
    for shape in sheet.api.Shapes:
        shape_box = (
            float(shape.Left),
            float(shape.Top),
            float(shape.Width),
            float(shape.Height),
        )
        if any(
            _overlaps(shape_box, proposed)
            for proposed in (new_button_box, panel_box, gap_box)
        ):
            raise PropertyStageOneError(
                "Proposed Property Stage 1 area overlaps a shape or button"
            )

    events = book.app.api.EnableEvents
    added: list[str] = []
    try:
        book.app.api.EnableEvents = False
        panel.color = (230, 230, 255)
        panel.font.name = "Calibri"
        panel.font.size = 10
        panel.api.VerticalAlignment = -4108
        heading = sheet.range(
            (control_panel.row, panel_left_column),
            (control_panel.row, panel_left_column + column_count - 1),
        )
        heading.merge()
        heading.value = "Property Stage 1 - Prepare LTM Flash"
        heading.color = (127, 127, 255)
        heading.font.color = (255, 255, 255)
        heading.font.bold = True
        book.names.add(PROPERTY_PANEL_NAME, f"='I. Control'!{panel.get_address()}")
        added.append(PROPERTY_PANEL_NAME)
        for field, offset in PROPERTY_OFFSETS.items():
            extra_row = 1 if field in ("input_template", "last_output") else 0
            end_row = control_panel.row + offset + extra_row
            label = sheet.range(
                (control_panel.row + offset, panel_left_column),
                (end_row, panel_left_column + 3),
            )
            value = sheet.range(
                (control_panel.row + offset, panel_left_column + 4),
                (end_row, panel_left_column + 7),
            )
            label.merge()
            value.merge()
            label.value = PROPERTY_LABELS[field]
            label.color = (127, 127, 255)
            label.font.color = (255, 255, 255)
            label.api.WrapText = True
            value.api.WrapText = True
            value.font.color = (0, 0, 0)
            value.api.HorizontalAlignment = -4131
            if field == "last_run":
                value.number_format = "dd/mm/yyyy"
            elif field in ("start", "finish"):
                value.number_format = "hh:mm:ss"
            elif field == "elapsed":
                value.number_format = "[h]:mm:ss"
            else:
                value.number_format = "General"
            if field == "status":
                value.value = "Not Run"
            name = PROPERTY_FIELD_NAMES[field]
            book.names.add(
                name,
                f"='I. Control'!{value[0, 0].get_address()}",
            )
            added.append(name)
    except Exception:
        for name in reversed(added):
            book.names[name].delete()
        raise
    finally:
        book.app.api.EnableEvents = events
    require_property_stage_one_panel(book)


def _install_property_button(book: Any) -> None:
    """Install or refresh the Prepare LTM Flash button in the planned position."""
    sheet = book.sheets["I. Control"]
    control_panel = book.names["SST_Stage1_Panel"].refers_to_range
    control_button = sheet.api.Shapes.Item("btnGenerateBondsLiteInput")
    left_column = control_panel.column + control_panel.columns.count + 1
    row = int(control_button.TopLeftCell.Row)
    left = sheet.range((row, left_column)).left
    top = control_panel.top + control_button.Top - control_panel.top
    button_width = float(control_button.Width)
    button_height = float(control_button.Height)
    if PROPERTY_BUTTON_NAME in [shape.Name for shape in sheet.api.Shapes]:
        button = sheet.api.Buttons().Item(PROPERTY_BUTTON_NAME)
    else:
        right_column = left_column
        while sheet.range((row, right_column)).left < left + button_width:
            right_column += 1
        button_cells = sheet.range(
            (row, left_column),
            (int(control_button.BottomRightCell.Row), right_column - 1),
        )
        if (
            book.app.api.WorksheetFunction.CountA(button_cells.api)
            or button_cells.api.MergeCells is not False
        ):
            raise PropertyStageOneError(
                "Proposed Prepare LTM Flash button covers occupied cells or merges"
            )
        proposed = (left, top, button_width, button_height)
        for shape in sheet.api.Shapes:
            shape_box = (
                float(shape.Left),
                float(shape.Top),
                float(shape.Width),
                float(shape.Height),
            )
            if _overlaps(shape_box, proposed):
                raise PropertyStageOneError(
                    "Proposed Prepare LTM Flash button overlaps an existing shape"
                )
        button = sheet.api.Buttons().Add(left, top, button_width, button_height)
        button.Name = PROPERTY_BUTTON_NAME
    button.Caption = "Prepare LTM Flash"
    button.OnAction = "'" + book.name.replace("'", "''") + "'!" + PROPERTY_BUTTON_MACRO


def _install_property_macro(book: Any) -> None:
    """Add the Property macro and upgrade the shared launcher signature."""
    try:
        module = book.api.VBProject.VBComponents.Item("modBondsLiteLauncher")
        code = module.CodeModule
        source = code.Lines(1, code.CountOfLines) if code.CountOfLines else ""
    except Exception as error:
        raise PropertyStageOneError(
            "Cannot access modBondsLiteLauncher in the Draft VBA project"
        ) from error
    signature = "Private Sub LaunchBondsLite(ByVal stage As Long)"
    if signature.casefold() in source.casefold():
        line = next(
            index
            for index in range(1, code.CountOfLines + 1)
            if code.Lines(index, 1).strip().casefold() == signature.casefold()
        )
        code.ReplaceLine(line, "Private Sub LaunchBondsLite(ByVal stage As String)")
        source = code.Lines(1, code.CountOfLines)
    if (
        "Private Sub LaunchBondsLite(ByVal stage As String)".casefold()
        not in source.casefold()
    ):
        raise PropertyStageOneError(
            "Draft launcher does not contain a supported LaunchBondsLite signature"
        )
    macro = "Public Sub PrepareLTMFlash()\n" '    LaunchBondsLite "P1"\n' "End Sub"
    if "Public Sub PrepareLTMFlash()".casefold() in source.casefold():
        if 'LaunchBondsLite "P1"'.casefold() not in source.casefold():
            raise PropertyStageOneError(
                "Draft PrepareLTMFlash macro has an unexpected implementation"
            )
    else:
        code.InsertLines(code.CountOfLines + 1, "\n" + macro + "\n")



def install_property_stage_one_controls(caller: Path) -> Path:
    """Install the Property flag, audit panel and VBA button into a closed Draft."""
    import xlwings as xw

    caller = Path(caller).resolve(strict=True)
    for application in xw.apps:
        for book in application.books:
            if book.name.casefold() == caller.name.casefold():
                raise PropertyStageOneError(
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
                raise PropertyStageOneError("Draft workbook is read-only or locked")
            _flag_control(book)
            _entity_control(book)
            _install_property_panel(book)
            _install_property_button(book)
            _install_property_macro(book)
            book.save()
        finally:
            book.close()
    finally:
        application.api.DisplayAlerts = False
        application.api.Quit()
    return backup
