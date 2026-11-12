"""Prepare a NoteVal copy linked to the successful Property Stage 2 workbook."""

from __future__ import annotations

import getpass
import logging
import math
import os
import re
import shutil
import tempfile
import time
from dataclasses import dataclass
from datetime import datetime
from numbers import Real
from pathlib import Path
from typing import Any
from urllib.parse import unquote

import bonds_lite_input as shared
import sst_property_stage_one as stage_one
import sst_property_stage_two as stage_two

logger = shared.logger.getChild("sst_property_stage_three")

STAGE_THREE_PANEL_NAME = "SST_PropStage3_Panel"
STAGE_THREE_PANEL_RANGE = "AB35:AI45"
STAGE_THREE_BUTTON_NAME = "btnPrepareNoteVal"
STAGE_THREE_BUTTON_MACRO = "PrepareNoteVal"
STAGE_THREE_FIELD_NAMES = {
    "last_run": "SST_PropStage3_LastRun",
    "start": "SST_PropStage3_Start",
    "finish": "SST_PropStage3_Finish",
    "elapsed": "SST_PropStage3_Elapsed",
    "user": "SST_PropStage3_User",
    "status": "SST_PropStage3_Status",
    "input_workbook": "SST_PropStage3_InputWorkbook",
    "last_successful_output": "SST_PropStage3_LastSuccessfulOutput",
}
STAGE_THREE_LABELS = {
    "last_run": "Last Run",
    "start": "Start:",
    "finish": "Finish:",
    "elapsed": "Time",
    "user": "User",
    "status": "Status",
    "input_workbook": "Input Workbook",
    "last_successful_output": "Last Successful Output",
}
STAGE_THREE_OFFSETS = {
    "last_run": 1,
    "start": 2,
    "finish": 3,
    "elapsed": 4,
    "user": 5,
    "status": 6,
    "input_workbook": 7,
    "last_successful_output": 9,
}
XL_CELL_TYPE_FORMULAS = -4123
XL_LINK_TYPE_EXCEL = 1
NOTEVAL_TEMPLATE_NAME = "SST_LTMNote_input_template"
NOTEVAL_DATE_NAME = "I_ValuationDate"


class PropertyStageThreeError(stage_two.PropertyStageTwoError):
    """An actionable Property Stage 3 input, workbook or audit failure."""


@dataclass(frozen=True)
class PropertyStageThreeInputs:
    """Validated Stage 2 handoff and NoteVal output contract."""

    caller_path: Path
    flash_path: Path
    template_path: Path
    output_path: Path
    valuation_date: datetime
    entity: str
    output_suffix: str


@dataclass(frozen=True)
class FlashLinkSnapshot:
    """The Flash links and formulas that Stage 3 is permitted to change."""

    old_source: str
    link_sources: tuple[str, ...]
    other_sources: tuple[str, ...]
    formulas: tuple[tuple[str, str, str], ...]


def _named_value(caller: Any, name: str, kind: str) -> object:
    """Read one named scalar control from the Draft."""
    return shared._control_cell(caller, name, kind).value


def _name_keys(book: Any) -> set[str]:
    """Return workbook-level Excel names without sheet-qualified prefixes."""
    return {
        str(getattr(name, "name", getattr(name, "Name", ""))).split("!")[-1]
        for name in book.names
    }


def _property_stage_three_field_cells(caller: Any) -> dict[str, Any]:
    """Resolve the Stage 3 audit fields to the first cells of merged ranges."""
    cells: dict[str, Any] = {}
    for field, name in STAGE_THREE_FIELD_NAMES.items():
        region = stage_one._audit_region(caller, name)
        merged = region.api.Cells(1, 1).MergeArea
        cells[field] = region.sheet.range((merged.Row, merged.Column))
    return cells


def require_property_stage_three_panel(caller: Any) -> None:
    """Validate Stage 3's named panel and ensure it does not overlap prior panels."""
    try:
        shared._require_named_stage_panel(
            caller, STAGE_THREE_PANEL_NAME, STAGE_THREE_FIELD_NAMES
        )
        stage_two.require_property_stage_two_panel(caller)
        first = shared._audit_bounds(
            shared._named_audit_region(caller, stage_two.STAGE_TWO_PANEL_NAME)
        )
        third = shared._audit_bounds(
            shared._named_audit_region(caller, STAGE_THREE_PANEL_NAME)
        )
    except shared.BondsLiteError as error:
        raise PropertyStageThreeError(str(error)) from error
    if (
        first[0] <= third[2]
        and first[2] >= third[0]
        and first[1] <= third[3]
        and first[3] >= third[1]
    ):
        raise PropertyStageThreeError(
            "Property Stage 2 and Stage 3 audit panels overlap"
        )


def _read_property_stage_three_fields(caller: Any) -> dict[str, object]:
    """Read current Stage 3 audit values after validating the named panel."""
    require_property_stage_three_panel(caller)
    return {
        field: cell.value
        for field, cell in _property_stage_three_field_cells(caller).items()
    }


def write_property_stage_three_audit(caller: Any, fields: dict[str, object]) -> None:
    """Write and save Stage 3 audit values through the named cells."""
    require_property_stage_three_panel(caller)
    if set(fields) - STAGE_THREE_FIELD_NAMES.keys():
        raise PropertyStageThreeError("Unknown Property Stage 3 audit field")
    cells = _property_stage_three_field_cells(caller)
    events = caller.app.api.EnableEvents
    try:
        caller.app.api.EnableEvents = False
        for field, value in fields.items():
            cell = cells[field]
            if field in ("input_workbook", "last_successful_output"):
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
            raise PropertyStageThreeError(
                "Draft's Property Stage 3 audit was not saved"
            )
    finally:
        caller.app.api.EnableEvents = events


def _audit_values(
    started: datetime,
    clock: float,
    status: str,
    input_workbook: str | None,
    last_successful_output: object,
    *,
    finish: datetime | None = None,
) -> dict[str, object]:
    """Build the Stage 3 audit fields while retaining the last successful output."""
    return {
        "last_run": started.replace(hour=0, minute=0, second=0, microsecond=0),
        "start": started,
        "finish": finish,
        "elapsed": (time.perf_counter() - clock) / 86400 if finish else None,
        "user": getpass.getuser(),
        "status": status,
        "input_workbook": input_workbook,
        "last_successful_output": last_successful_output,
    }


def _record_failure(
    caller: Any,
    started: datetime,
    clock: float,
    input_workbook: str | None,
    previous_success: object,
    status: str,
) -> None:
    """Persist a terminal failure while retaining the previous successful output."""
    write_property_stage_three_audit(
        caller,
        _audit_values(
            started,
            clock,
            status,
            input_workbook,
            previous_success,
            finish=datetime.now(),
        ),
    )


def resolve_property_stage_three_inputs(caller: Any) -> PropertyStageThreeInputs:
    """Validate successful Property Stage 2 output and resolve the NoteVal copy."""
    caller_path = shared.resolve_caller_path(caller)
    try:
        stage_one.require_property_selected(stage_one.read_property_flag(caller))
        stage_one.require_property_stage_one_panel(caller)
        stage_two.require_property_stage_two_panel(caller)
        stage_one_fields = stage_one._read_property_stage_one_fields(caller)
        stage_two_fields = stage_two._read_property_stage_two_fields(caller)
    except shared.BondsLiteError as error:
        raise PropertyStageThreeError(str(error)) from error
    if stage_two_fields["status"] != "Succeeded":
        raise PropertyStageThreeError(
            "Property Stage 3 requires Property Stage 2 status Succeeded"
        )
    try:
        stage_one_path = stage_two._resolve_handoff_path(
            stage_one_fields["last_output"],
            caller_path,
            "Property Stage 1 output workbook",
        )
        flash_path = stage_two._resolve_handoff_path(
            stage_two_fields["input_workbook"],
            caller_path,
            "Property Stage 2 Flash input workbook",
        )
    except stage_two.PropertyStageTwoError as error:
        raise PropertyStageThreeError(str(error)) from error
    if os.path.normcase(str(flash_path)) != os.path.normcase(str(stage_one_path)):
        raise PropertyStageThreeError(
            "Property Stage 2 input workbook does not match the successful "
            "Property Stage 1 output"
        )
    try:
        stage_one_finish = stage_two._excel_datetime(
            stage_one_fields["finish"], "Property Stage 1 finish"
        )
        stage_two_finish = stage_two._excel_datetime(
            stage_two_fields["last_successful_run"],
            "Property Stage 2 last successful run",
        )
    except stage_two.PropertyStageTwoError as error:
        raise PropertyStageThreeError(str(error)) from error
    if stage_two_finish < stage_one_finish:
        raise PropertyStageThreeError(
            "Property Stage 2 last successful run is earlier than the "
            "Property Stage 1 finish"
        )

    try:
        template_path = stage_one._resolve_template_path(
            _named_value(caller, NOTEVAL_TEMPLATE_NAME, "NoteVal template name"),
            caller_path,
        )
        if template_path.suffix.casefold() not in stage_one.ALLOWED_TEMPLATE_SUFFIXES:
            raise PropertyStageThreeError(
                "NoteVal template must be an .xlsb, .xlsm or .xlsx workbook"
            )
        valuation_date = stage_one._date_value(
            _named_value(caller, "I_Valuation_Date", "Valuation date name")
        )
        entity = shared.read_entity(caller)
        output_suffix = shared.read_output_suffix(caller)
        filename = shared.entity_output_name(entity, template_path, output_suffix)
    except shared.BondsLiteError as error:
        raise PropertyStageThreeError(str(error)) from error
    try:
        output_path = stage_one._property_output_path(
            caller_path, flash_path.parent, filename
        ).resolve()
    except shared.BondsLiteError as error:
        raise PropertyStageThreeError(str(error)) from error
    if os.path.normcase(str(output_path)) == os.path.normcase(
        str(template_path.resolve())
    ):
        raise PropertyStageThreeError(
            "NoteVal template cannot be the Stage 3 output workbook"
        )
    return PropertyStageThreeInputs(
        caller_path=caller_path,
        flash_path=flash_path,
        template_path=template_path,
        output_path=output_path,
        valuation_date=valuation_date,
        entity=entity,
        output_suffix=output_suffix,
    )


def _source_filename(source: str) -> str:
    """Return a decoded external workbook filename from a Windows path or URL."""
    return unquote(str(source).replace("\\", "/").rstrip("/").split("/")[-1])


def _formula_external_filenames(formula: str) -> tuple[str, ...]:
    """Find the workbook names in Excel external-reference brackets."""
    return tuple(re.findall(r"\[([^\]]+)\]", formula))


def _unescaped_quote_positions(value: str) -> list[int]:
    """Return positions of single quotes, ignoring Excel's doubled apostrophes."""
    positions = []
    index = 0
    while index < len(value):
        if value[index] == "'":
            if index + 1 < len(value) and value[index + 1] == "'":
                index += 2
                continue
            positions.append(index)
        index += 1
    return positions


def _replace_external_path(
    formula: str, old_filename: str, new_path: Path
) -> tuple[str, int]:
    """Replace each quoted old workbook path while preserving the sheet and cells."""
    brackets = list(re.finditer(r"\[([^\]]+)\]", formula))
    if not brackets:
        return formula, 0
    quotes = _unescaped_quote_positions(formula)
    replacements: list[tuple[int, int, str]] = []
    for bracket in brackets:
        filename = unquote(bracket.group(1)).casefold()
        if filename != old_filename.casefold():
            continue
        start = next(
            (position for position in reversed(quotes) if position < bracket.start()),
            None,
        )
        end = next(
            (position for position in quotes if position > bracket.end()),
            None,
        )
        if start is None or end is None:
            continue
        tail = formula[bracket.end() : end]
        parent = str(new_path.parent).replace("/", "\\").rstrip("\\")
        filename = new_path.name.replace("'", "''")
        new_reference = (
            "'" + parent.replace("'", "''") + "\\[" + filename + "]" + tail + "'"
        )
        replacements.append((start, end + 1, new_reference))
    for start, end, replacement in reversed(replacements):
        formula = formula[:start] + replacement + formula[end:]
    return formula, len(replacements)


def rewrite_external_reference(formula: str, old_source: str, new_source: Path) -> str:
    """Repoint a formula's old Flash link without changing its formula structure."""
    if not isinstance(formula, str):
        raise PropertyStageThreeError("Cannot rewrite a non-text Excel formula")
    old_filename = _source_filename(old_source)
    rewritten, count = _replace_external_path(formula, old_filename, Path(new_source))
    if count == 0:
        raise PropertyStageThreeError(
            f"Formula does not contain the expected Flash source {old_filename}"
        )
    return rewritten


def _cell_address(row: int, column: int) -> str:
    """Format a one-based Excel cell coordinate without absolute markers."""
    letters = ""
    while column:
        column, remainder = divmod(column - 1, 26)
        letters = chr(65 + remainder) + letters
    return f"{letters}{row}"


def _formula_grid(value: object) -> list[list[object]]:
    """Normalise an Excel formula return value to a two-dimensional grid."""
    if isinstance(value, (tuple, list)):
        if value and isinstance(value[0], (tuple, list)):
            return [list(row) for row in value]
        return [list(value)]
    return [[value]]


def _iter_formula_cells(book: Any) -> list[tuple[str, str, str]]:
    """Read formulas from the NoteVal sheets that may contain the Flash link."""
    result: list[tuple[str, str, str]] = []
    for sheet in book.sheets:
        if str(sheet.name).casefold() not in {"i. inputs", "i. control"}:
            continue
        used = sheet.api.UsedRange
        if used.HasFormula is False:
            continue
        try:
            areas = used.SpecialCells(XL_CELL_TYPE_FORMULAS).Areas
            area_count = int(areas.Count)
        except Exception as error:
            if "protected sheet" not in str(error).casefold():
                raise PropertyStageThreeError(
                    f"Cannot inspect formulas on NoteVal sheet {sheet.name}: {error}"
                ) from error
            logger.warning(
                "Scanning the read-only used-range formula grid on protected "
                "NoteVal sheet %s",
                sheet.name,
            )
            try:
                grid = _formula_grid(used.Formula)
                for row_offset, row in enumerate(grid):
                    for column_offset, formula in enumerate(row):
                        if isinstance(formula, str) and formula.startswith("="):
                            result.append(
                                (
                                    str(sheet.name),
                                    _cell_address(
                                        int(used.Row) + row_offset,
                                        int(used.Column) + column_offset,
                                    ),
                                    formula,
                                )
                            )
            except Exception as fallback_error:
                raise PropertyStageThreeError(
                    f"Cannot read formulas on protected NoteVal sheet "
                    f"{sheet.name}: {fallback_error}"
                ) from fallback_error
            continue
        for area_index in range(1, area_count + 1):
            area = areas.Item(area_index)
            for row_offset, row in enumerate(_formula_grid(area.Formula)):
                for column_offset, formula in enumerate(row):
                    if not isinstance(formula, str) or not formula.startswith("="):
                        continue
                    result.append(
                        (
                            str(sheet.name),
                            _cell_address(
                                int(area.Row) + row_offset,
                                int(area.Column) + column_offset,
                            ),
                            formula,
                        )
                    )
    return result


def _normalise_address(address: str) -> str:
    """Remove Excel absolute markers from an A1 address."""
    return address.replace("$", "").upper()


def _scope_cell(sheet: str, address: str) -> bool:
    """Return whether a formula cell is within the approved NoteVal link scope."""
    address = _normalise_address(address)
    if sheet.casefold() == "i. control":
        return address == "D12"
    if sheet.casefold() != "i. inputs":
        return False
    if address == "I10":
        return True
    match = re.fullmatch(r"([CEG])(\d+)", address)
    return match is not None and int(match.group(2)) >= 10


def _normalise_external_paths(formula: str) -> str:
    """Remove directory differences from external references for formula checks."""
    brackets = list(re.finditer(r"\[([^\]]+)\]", formula))
    quotes = _unescaped_quote_positions(formula)
    for bracket in reversed(brackets):
        start = next(
            (position for position in reversed(quotes) if position < bracket.start()),
            None,
        )
        if start is None:
            continue
        formula = formula[: start + 1] + formula[bracket.start() :]
    return formula


def _normalise_source(source: str) -> str:
    """Normalise a workbook link path for cross-format path comparisons."""
    return unquote(source).replace("\\", "/").rstrip("/").casefold()


def _source_matches(source: str, expected: str | Path) -> bool:
    """Match a LinkSources item to a local workbook or previously recorded link."""
    expected_text = str(expected)
    if _normalise_source(source) == _normalise_source(expected_text):
        return True
    return _source_filename(source).casefold() == Path(expected_text).name.casefold()


def snapshot_flash_links(book: Any) -> FlashLinkSnapshot:
    """Capture and validate the exact formula cells referencing the old Flash."""
    links_value = book.api.LinkSources(1)
    if links_value is None:
        raise PropertyStageThreeError("NoteVal has no external workbook links")
    links = (
        (str(links_value),)
        if isinstance(links_value, str)
        else tuple(str(source) for source in links_value)
    )
    formulas = _iter_formula_cells(book)
    control_formula = next(
        (
            formula
            for sheet, address, formula in formulas
            if sheet.casefold() == "i. control" and _normalise_address(address) == "D12"
        ),
        None,
    )
    if control_formula is None:
        raise PropertyStageThreeError("NoteVal I. Control!D12 has no formula")
    referenced_files = {
        unquote(filename).casefold()
        for filename in _formula_external_filenames(control_formula)
    }
    if len(referenced_files) != 1:
        raise PropertyStageThreeError(
            "NoteVal I. Control!D12 must reference one external Flash workbook"
        )
    old_filename = next(iter(referenced_files))
    matching_sources = [
        source
        for source in links
        if _source_filename(source).casefold() == old_filename
    ]
    if len(matching_sources) != 1:
        raise PropertyStageThreeError(
            "Cannot identify one Flash workbook in NoteVal's external links"
        )
    old_source = matching_sources[0]
    old_references = {
        (sheet, _normalise_address(address))
        for sheet, address, formula in formulas
        if any(
            unquote(filename).casefold() == old_filename
            for filename in _formula_external_filenames(formula)
        )
    }

    in_scope_formulas = [
        (sheet, _normalise_address(address), formula)
        for sheet, address, formula in formulas
        if _scope_cell(sheet, address)
    ]
    expected_scope = {(sheet, address) for sheet, address, _ in in_scope_formulas}
    if ("I. Control", "D12") not in expected_scope or (
        "I. Inputs",
        "I10",
    ) not in expected_scope:
        raise PropertyStageThreeError(
            "NoteVal Flash link scope is missing I. Control!D12 or I. Inputs!I10"
        )
    for sheet, address, formula in in_scope_formulas:
        filenames = _formula_external_filenames(formula)
        if not filenames or any(
            unquote(filename).casefold() != old_filename for filename in filenames
        ):
            raise PropertyStageThreeError(
                f"Unexpected external workbook reference in {sheet}!{address}"
            )
    if old_references != expected_scope:
        unexpected = sorted(old_references - expected_scope)
        missing = sorted(expected_scope - old_references)
        details = unexpected or missing
        location = f"{details[0][0]}!{details[0][1]}" if details else "unknown cell"
        raise PropertyStageThreeError(
            f"Flash reference is outside the approved NoteVal scope at {location}"
        )
    for name in book.names:
        refers_to = str(getattr(name, "refers_to", getattr(name, "RefersTo", "")))
        if any(
            unquote(filename).casefold() == old_filename
            for filename in _formula_external_filenames(refers_to)
        ):
            raise PropertyStageThreeError(
                f"Defined name also refers to the old Flash source: {name.name}"
            )
    return FlashLinkSnapshot(
        old_source=old_source,
        link_sources=links,
        other_sources=tuple(source for source in links if source != old_source),
        formulas=tuple(
            (sheet, address, formula) for sheet, address, formula in in_scope_formulas
        ),
    )


def _get_sheet(workbook: Any, name: str) -> Any:
    """Resolve a workbook sheet by name for xlwings and small test doubles."""
    sheets = workbook.sheets
    try:
        return sheets[name]
    except (KeyError, TypeError, IndexError):
        for sheet in sheets:
            if str(sheet.name).casefold() == name.casefold():
                return sheet
    raise PropertyStageThreeError(f"Flash workbook is missing the {name} sheet")


def _grid(value: object) -> list[list[object]]:
    """Normalise a sheet range's values to a two-dimensional list."""
    if isinstance(value, (tuple, list)):
        if value and isinstance(value[0], (tuple, list)):
            return [list(row) for row in value]
        return [list(value)]
    return [[value]]


def _scenario_description(rows: list[list[object]], scenario: int) -> str:
    """Find a ScenFlipper description for one scenario number."""
    for row in rows:
        if len(row) <= 9 or not isinstance(row[0], Real) or isinstance(row[0], bool):
            continue
        if float(row[0]).is_integer() and int(row[0]) == scenario:
            value = row[9]
            return str(value).strip() if value is not None else ""
    return "missing"


def check_stressed_scenario(flash_workbook: Any) -> None:
    """Require NoteVal's fixed CJ columns to identify the first stressed scenario."""
    cf = _get_sheet(flash_workbook, "O. CFOutput")
    pv = _get_sheet(flash_workbook, "O. PVOutput")
    cf_scenario = cf.range("CJ1").value
    pv_scenario = pv.range("CJ1").value
    if cf_scenario != 9 or pv_scenario != 9:
        raise PropertyStageThreeError(
            "NoteVal reads scenario 9 from Flash column CJ, but the headers "
            f"are CF={cf_scenario!r} and PV={pv_scenario!r}"
        )
    flipper = _get_sheet(flash_workbook, "I. ScenFlipper")
    try:
        rows = _grid(flipper.api.UsedRange.Value)
    except Exception as error:
        raise PropertyStageThreeError(
            f"Cannot inspect Flash I. ScenFlipper scenarios: {error}"
        ) from error
    eighth = _scenario_description(rows, 8)
    ninth = _scenario_description(rows, 9)
    if eighth.casefold() != "base" or ninth not in {
        "Interest",
        "Property",
        "Interest_Property",
    }:
        raise PropertyStageThreeError(
            "NoteVal reads Flash scenario 9, which must be the first stressed "
            f"scenario; scenario 8 is {eighth!r} and scenario 9 is {ninth!r}"
        )


def _noteval_date_range(book: Any) -> Any:
    """Resolve NoteVal's named valuation-date cell."""
    try:
        region = book.names[NOTEVAL_DATE_NAME].refers_to_range
        if region.api.Areas.Count != 1 or region.size != 1:
            raise PropertyStageThreeError(
                f"NoteVal name must refer to one cell: {NOTEVAL_DATE_NAME}"
            )
        return region
    except PropertyStageThreeError:
        raise
    except Exception as error:
        raise PropertyStageThreeError(
            f"NoteVal is missing a usable {NOTEVAL_DATE_NAME} name: {error}"
        ) from error


def set_noteval_valuation_date(book: Any, valuation_date: datetime) -> None:
    """Write and verify the Draft valuation date without changing cell formatting."""
    cell = _noteval_date_range(book)
    cell.value = valuation_date
    try:
        readback = stage_one._date_value(cell.value)
    except shared.BondsLiteError as error:
        raise PropertyStageThreeError(
            f"NoteVal {NOTEVAL_DATE_NAME} did not retain the valuation date"
        ) from error
    if readback != valuation_date:
        raise PropertyStageThreeError(
            f"NoteVal {NOTEVAL_DATE_NAME} did not retain the valuation date"
        )


def _formula_cell(book: Any, sheet_name: str, address: str) -> Any:
    """Return one formula cell from a named sheet."""
    try:
        return _get_sheet(book, sheet_name).range(address)
    except PropertyStageThreeError:
        raise
    except Exception as error:
        raise PropertyStageThreeError(
            f"Cannot read NoteVal {sheet_name}!{address}: {error}"
        ) from error


def _range_sum(sheet: Any, address: str) -> float:
    """Sum a referenced Flash range, rejecting non-numeric cells."""
    values = _grid(sheet.range(address).value)
    total = 0.0
    for row in values:
        for value in row:
            if not isinstance(value, Real) or isinstance(value, bool):
                raise PropertyStageThreeError(
                    f"Flash range {sheet.name}!{address} contains a non-numeric value"
                )
            total += float(value)
    return total


_CELL_TOKEN = re.compile(
    r"(?<![A-Z0-9])\$?([A-Z]{1,3})\$?(\d+)" r"(?::\$?([A-Z]{1,3})\$?(\d+))?",
    re.IGNORECASE,
)


def _formula_ranges(formula: str) -> list[str]:
    """Extract A1 ranges and cells from an external-reference formula."""
    ranges = []
    for segment in formula.split("!")[1:]:
        reference_text = segment.split("'")[0]
        for match in _CELL_TOKEN.finditer(reference_text):
            first = f"{match.group(1).upper()}{match.group(2)}"
            second = (
                f"{match.group(3).upper()}{match.group(4)}" if match.group(3) else ""
            )
            ranges.append(f"{first}:{second}" if second else first)
    return ranges


def _expected_value(sheet: Any, address: str, formula: str, sheet_name: str) -> float:
    """Calculate the simple fixed-range value represented by a NoteVal formula."""
    ranges = _formula_ranges(formula)
    if not ranges:
        raise PropertyStageThreeError(
            f"Cannot interpret NoteVal formula at {sheet_name}!{address}"
        )
    if sheet_name.casefold() == "i. control":
        if len(ranges) != 2:
            raise PropertyStageThreeError(
                "NoteVal I. Control!D12 does not contain two PV ranges"
            )
        return _range_sum(sheet, ranges[0]) - _range_sum(sheet, ranges[1])
    if sheet_name.casefold() == "i. inputs" and _normalise_address(address) == "I10":
        raise PropertyStageThreeError(
            "NoteVal I. Inputs!I10 must be verified as a spill range"
        )
    if _normalise_address(address).startswith("C"):
        if len(ranges) != 2:
            raise PropertyStageThreeError(
                f"NoteVal {sheet_name}!{address} does not contain two CF ranges"
            )
        return _range_sum(sheet, ranges[0]) - _range_sum(sheet, ranges[1])
    return -_range_sum(sheet, ranges[0])


def _numeric_value(value: object, sheet_name: str, address: str) -> float:
    """Require a calculated numeric NoteVal result."""
    if (
        not isinstance(value, Real)
        or isinstance(value, bool)
        or not math.isfinite(value)
    ):
        raise PropertyStageThreeError(
            f"NoteVal {sheet_name}!{address} is not a valid calculated number: {value!r}"
        )
    return float(value)


def _verify_i10_spill(cell: Any, flash_sheet: Any, formula: str) -> None:
    """Verify every value in the I10 spill against the negated Flash range."""
    ranges = _formula_ranges(formula)
    if len(ranges) != 1 or ":" not in ranges[0]:
        raise PropertyStageThreeError(
            "NoteVal I. Inputs!I10 must reference one contiguous Flash range"
        )
    source = flash_sheet.range(ranges[0])
    source_api = source.api
    rows = int(source_api.Rows.Count)
    columns = int(source_api.Columns.Count)
    source_start_row = int(source_api.Row)
    source_start_column = int(source_api.Column)
    anchor = cell.api
    if not anchor.HasSpill:
        raise PropertyStageThreeError(
            "NoteVal I. Inputs!I10 does not spill the Flash source range"
        )
    spill = anchor.SpillingToRange
    start_row, start_column = int(anchor.Row), int(anchor.Column)
    if (
        int(spill.Row) != start_row
        or int(spill.Column) != start_column
        or int(spill.Rows.Count) != rows
        or int(spill.Columns.Count) != columns
    ):
        raise PropertyStageThreeError(
            "NoteVal I. Inputs!I10 spill dimensions do not match the Flash range"
        )
    destination = cell.sheet.range(
        (start_row, start_column),
        (start_row + rows - 1, start_column + columns - 1),
    )
    source_values = _grid(source.options(ndim=2).value)
    output_values = _grid(destination.options(ndim=2).value)
    if any(
        len(values) != rows or any(len(row) != columns for row in values)
        for values in (source_values, output_values)
    ):
        raise PropertyStageThreeError(
            "NoteVal I. Inputs!I10 or its Flash source range has incomplete values"
        )
    for row_index, (source_row, output_row) in enumerate(
        zip(source_values, output_values)
    ):
        for column_index, (source_value, output_value) in enumerate(
            zip(source_row, output_row)
        ):
            source_address = _cell_address(
                source_start_row + row_index,
                source_start_column + column_index,
            )
            output_address = _cell_address(
                start_row + row_index, start_column + column_index
            )
            if (
                not isinstance(source_value, Real)
                or isinstance(source_value, bool)
                or not math.isfinite(source_value)
            ):
                raise PropertyStageThreeError(
                    f"Flash {flash_sheet.name}!{source_address} is not a valid "
                    f"calculated number: {source_value!r}"
                )
            actual = _numeric_value(output_value, "I. Inputs", output_address)
            expected = -float(source_value)
            if not math.isclose(actual, expected, rel_tol=1e-6, abs_tol=1e-6):
                raise PropertyStageThreeError(
                    f"NoteVal value at I. Inputs!{output_address} is {actual}; "
                    f"Flash formula value is {expected}"
                )


def verify_noteval(
    workbook: Any,
    snapshot: FlashLinkSnapshot,
    flash_workbook: Any,
    valuation_date: datetime,
) -> None:
    """Verify links, formulas, calculated values and the valuation date."""
    current_links_value = workbook.api.LinkSources(1)
    current_links = (
        ()
        if current_links_value is None
        else (
            (current_links_value,)
            if isinstance(current_links_value, str)
            else tuple(str(source) for source in current_links_value)
        )
    )
    if any(_source_matches(source, snapshot.old_source) for source in current_links):
        raise PropertyStageThreeError("Old Flash source remains in NoteVal LinkSources")
    if not any(
        _source_matches(source, snapshot_link_target)
        for source in current_links
        for snapshot_link_target in (flash_workbook.fullname,)
    ):
        raise PropertyStageThreeError(
            "Updated Flash workbook is missing from NoteVal LinkSources"
        )
    current_other = tuple(
        source
        for source in current_links
        if not _source_matches(source, flash_workbook.fullname)
    )
    if tuple(map(_normalise_source, current_other)) != tuple(
        map(_normalise_source, snapshot.other_sources)
    ):
        raise PropertyStageThreeError(
            "NoteVal external links other than the Flash workbook have changed"
        )
    for sheet_name, address, original in snapshot.formulas:
        cell = _formula_cell(workbook, sheet_name, address)
        current = cell.formula
        if isinstance(current, (tuple, list)):
            current = (
                current[0][0] if isinstance(current[0], (tuple, list)) else current[0]
            )
        expected = rewrite_external_reference(
            original, snapshot.old_source, Path(flash_workbook.fullname)
        )
        if _normalise_external_paths(str(current)) != _normalise_external_paths(
            expected
        ):
            raise PropertyStageThreeError(
                f"NoteVal formula changed beyond its Flash path at "
                f"{sheet_name}!{address}"
            )
        if "#REF!" in str(cell.value).upper() or "#VALUE!" in str(cell.value).upper():
            raise PropertyStageThreeError(
                f"NoteVal formula returned an Excel error at {sheet_name}!{address}"
            )
        flash_sheet = flash_workbook.sheets[
            "O. PVOutput" if sheet_name.casefold() == "i. control" else "O. CFOutput"
        ]
        if (
            sheet_name.casefold() == "i. inputs"
            and _normalise_address(address) == "I10"
        ):
            _verify_i10_spill(cell, flash_sheet, original)
            continue
        expected_value = _expected_value(flash_sheet, address, original, sheet_name)
        actual_value = _numeric_value(cell.value, sheet_name, address)
        if not math.isclose(actual_value, expected_value, rel_tol=1e-6, abs_tol=1e-6):
            raise PropertyStageThreeError(
                f"NoteVal value at {sheet_name}!{address} is {actual_value}; "
                f"Flash formula value is {expected_value}"
            )
    if stage_one._date_value(_noteval_date_range(workbook).value) != valuation_date:
        raise PropertyStageThreeError(
            f"NoteVal {NOTEVAL_DATE_NAME} does not match the Draft valuation date"
        )


def _manual_calculation(application: Any) -> Any:
    """Temporarily put the caller's Excel application into manual calculation."""
    return stage_one._manual_calculation(application)


def _write_failure_audit(
    caller: Any,
    started: datetime,
    clock: float,
    input_workbook: str | None,
    previous_success: object,
    status: str,
) -> None:
    """Write a terminal failure status while retaining the last successful output."""
    _record_failure(caller, started, clock, input_workbook, previous_success, status)


def _portable_workbook_path(caller_path: Path, workbook_path: Path) -> str:
    """Return an audit link relative to the Draft when possible."""
    return stage_two._portable_workbook_path(caller_path, workbook_path)


def run_property_stage_three(caller: Any) -> Path:
    """Prepare and verify a local NoteVal copy linked to the successful Flash run."""
    caller_path = shared.resolve_caller_path(caller)
    with shared.workbook_run_guard(
        caller_path,
        conflict_message="A stress run is already in progress for this workbook",
    ):
        try:
            previous_success = _read_property_stage_three_fields(caller)[
                "last_successful_output"
            ]
        except Exception as error:
            raise PropertyStageThreeError(
                f"Property Stage 3 audit panel is unavailable: {error}"
            ) from error
        started = datetime.now()
        clock = time.perf_counter()
        flash_link: str | None = None
        temporary_folder: Path | None = None
        output_backup: Path | None = None
        output_preexisted = False
        output_replaced = False
        try:
            inputs = resolve_property_stage_three_inputs(caller)
            flash_link = _portable_workbook_path(caller_path, inputs.flash_path)
            if inputs.valuation_date.weekday() >= 5:
                logger.warning(
                    "Draft valuation date %s falls on a weekend",
                    inputs.valuation_date.date(),
                )
            stage_two._ensure_workbook_not_open(inputs.flash_path, caller, caller_path)
            stage_two._ensure_workbook_not_open(inputs.output_path, caller, caller_path)
            write_property_stage_three_audit(
                caller,
                _audit_values(
                    started,
                    clock,
                    "Running",
                    flash_link,
                    previous_success,
                ),
            )
            if not inputs.template_path.is_file():
                raise PropertyStageThreeError(
                    f"NoteVal template does not exist: {inputs.template_path}"
                )
            temporary_folder = Path(
                tempfile.mkdtemp(
                    prefix="~property-stage3-", dir=inputs.flash_path.parent
                )
            )
            working_path = temporary_folder / inputs.output_path.name
            try:
                shutil.copy2(inputs.template_path, working_path)
            except OSError as error:
                raise PropertyStageThreeError(
                    "Cannot copy the NoteVal template; make the file available offline: "
                    f"{error}"
                ) from error
            if working_path.stat().st_size != inputs.template_path.stat().st_size:
                raise PropertyStageThreeError(
                    "NoteVal template copy has an unexpected file size; "
                    "make the file available offline"
                )

            application = caller.app
            with shared.excel_state(application), _manual_calculation(application):
                with shared.owned_book(
                    application, working_path, read_only=False
                ) as noteval:
                    snapshot = snapshot_flash_links(noteval)
                    with shared.owned_book(
                        application, inputs.flash_path, read_only=True
                    ) as flash:
                        check_stressed_scenario(flash)
                        noteval.api.ChangeLink(
                            snapshot.old_source,
                            str(inputs.flash_path),
                            XL_LINK_TYPE_EXCEL,
                        )
                        set_noteval_valuation_date(noteval, inputs.valuation_date)
                        application.api.CalculateFull()
                        verify_noteval(
                            noteval,
                            snapshot,
                            flash,
                            inputs.valuation_date,
                        )
                    noteval.save()
                    if not noteval.api.Saved:
                        raise PropertyStageThreeError(
                            "NoteVal working copy was not saved"
                        )

            output_preexisted = inputs.output_path.exists()
            if output_preexisted:
                output_backup = temporary_folder / (
                    inputs.output_path.name + ".previous"
                )
                shutil.copy2(inputs.output_path, output_backup)
                if shared._file_hash(inputs.output_path) != shared._file_hash(
                    output_backup
                ):
                    raise PropertyStageThreeError(
                        "Previous NoteVal output backup could not be verified"
                    )
            os.replace(working_path, inputs.output_path)
            output_replaced = True
            finished = datetime.now()
            output_link = _portable_workbook_path(caller_path, inputs.output_path)
            write_property_stage_three_audit(
                caller,
                _audit_values(
                    started,
                    clock,
                    "Succeeded",
                    flash_link,
                    output_link,
                    finish=finished,
                ),
            )
            output_replaced = False
            logger.info(
                "Property Stage 3 succeeded; saved NoteVal: %s", inputs.output_path
            )
            return inputs.output_path
        except BaseException as error:
            status = "Interrupted" if isinstance(error, KeyboardInterrupt) else "Failed"
            restoration_error: OSError | None = None
            if output_replaced:
                try:
                    if output_preexisted and output_backup is not None:
                        os.replace(output_backup, inputs.output_path)
                    else:
                        inputs.output_path.unlink(missing_ok=True)
                except OSError as restore_error:
                    restoration_error = restore_error
                    logger.error(
                        "Property Stage 3 failed and the previous NoteVal output "
                        "could not be restored: %s",
                        restore_error,
                    )
            try:
                _write_failure_audit(
                    caller,
                    started,
                    clock,
                    flash_link,
                    previous_success,
                    status,
                )
            except Exception as audit_error:
                logger.error(
                    "Property Stage 3 failed and its audit could not be saved: %s",
                    audit_error,
                )
            if restoration_error is not None:
                raise PropertyStageThreeError(
                    "Property Stage 3 failed and could not restore the previous "
                    f"NoteVal output: {restoration_error}"
                ) from error
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
                        "Could not remove Property Stage 3 staging folder %s: %s",
                        temporary_folder,
                        cleanup_error,
                    )


def _initialise_property_stage_three_panel(book: Any) -> None:
    """Install the fixed Property Stage 3 audit panel without overwriting cells."""
    existing = _name_keys(book)
    expected = {STAGE_THREE_PANEL_NAME, *STAGE_THREE_FIELD_NAMES.values()}
    if existing.intersection(expected):
        if not expected.issubset(existing):
            raise PropertyStageThreeError(
                "Property Stage 3 audit names are incomplete; installation was not changed"
            )
        require_property_stage_three_panel(book)
        return
    sheet = book.sheets["I. Control"]
    area = sheet.range(STAGE_THREE_PANEL_RANGE)
    if (
        book.app.api.WorksheetFunction.CountA(area.api)
        or area.api.MergeCells is not False
    ):
        raise PropertyStageThreeError(
            "Proposed Property Stage 3 audit area contains cells or merges"
        )
    proposed = (area.left, area.top, area.width, area.height)
    for shape in sheet.api.Shapes:
        shape_box = (
            float(shape.Left),
            float(shape.Top),
            float(shape.Width),
            float(shape.Height),
        )
        if stage_two._panel_overlap(shape_box, proposed):
            raise PropertyStageThreeError(
                "Proposed Property Stage 3 audit area overlaps a shape or button"
            )
    row, column = area.row, area.column
    events = book.app.api.EnableEvents
    added: list[str] = []
    try:
        book.app.api.EnableEvents = False
        area.color = (230, 230, 255)
        area.font.name = "Calibri"
        area.font.size = 10
        area.api.VerticalAlignment = -4108
        heading = sheet.range((row, column), (row, column + 7))
        heading.merge()
        heading.value = "Property Stage 3 - Prepare NoteVal"
        heading.color = (127, 127, 255)
        heading.font.color = (255, 255, 255)
        heading.font.bold = True
        book.names.add(
            STAGE_THREE_PANEL_NAME,
            f"='I. Control'!{area.get_address()}",
        )
        added.append(STAGE_THREE_PANEL_NAME)
        for field, offset in STAGE_THREE_OFFSETS.items():
            extra_row = (
                1 if field in ("input_workbook", "last_successful_output") else 0
            )
            label = sheet.range(
                (row + offset, column),
                (row + offset + extra_row, column + 3),
            )
            value = sheet.range(
                (row + offset, column + 4),
                (row + offset + extra_row, column + 7),
            )
            label.merge()
            value.merge()
            label.value = STAGE_THREE_LABELS[field]
            label.color = (127, 127, 255)
            label.font.color = (255, 255, 255)
            label.api.WrapText = True
            value.api.WrapText = True
            value.font.color = (0, 0, 0)
            value.api.HorizontalAlignment = -4131
            if field == "last_run":
                value.number_format = "dd/mm/yyyy"
            elif field in ("start", "finish"):
                value.number_format = "dd/mm/yyyy hh:mm:ss"
            elif field == "elapsed":
                value.number_format = "[h]:mm:ss"
            else:
                value.number_format = "General"
            if field == "status":
                value.value = "Not Run"
            name = STAGE_THREE_FIELD_NAMES[field]
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
    require_property_stage_three_panel(book)


def _install_property_stage_three_button(book: Any) -> None:
    """Install the Stage 3 button beside its audit panel."""
    sheet = book.sheets["I. Control"]
    if STAGE_THREE_BUTTON_NAME in [shape.Name for shape in sheet.api.Shapes]:
        button = sheet.api.Buttons().Item(STAGE_THREE_BUTTON_NAME)
    else:
        stage_two_button = sheet.api.Shapes.Item(stage_two.STAGE_TWO_BUTTON_NAME)
        panel = book.names[STAGE_THREE_PANEL_NAME].refers_to_range
        left = float(stage_two_button.Left)
        top = float(panel.top)
        width = float(stage_two_button.Width)
        height = float(stage_two_button.Height)
        row = panel.row
        column = int(stage_two_button.TopLeftCell.Column)
        end_column = column
        while sheet.range((row, end_column)).left < left + width:
            end_column += 1
        cells = sheet.range((row, column), (row + 2, end_column - 1))
        if (
            book.app.api.WorksheetFunction.CountA(cells.api)
            or cells.api.MergeCells is not False
        ):
            raise PropertyStageThreeError(
                "Proposed Prepare NoteVal button covers occupied cells or merges"
            )
        proposed = (left, top, width, height)
        for shape in sheet.api.Shapes:
            shape_box = (
                float(shape.Left),
                float(shape.Top),
                float(shape.Width),
                float(shape.Height),
            )
            if stage_two._panel_overlap(shape_box, proposed):
                raise PropertyStageThreeError(
                    "Proposed Prepare NoteVal button overlaps an existing shape"
                )
        button = sheet.api.Buttons().Add(left, top, width, height)
        button.Name = STAGE_THREE_BUTTON_NAME
    button.Caption = "Prepare NoteVal"
    button.OnAction = (
        "'" + book.name.replace("'", "''") + "'!" + STAGE_THREE_BUTTON_MACRO
    )


def _install_property_stage_three_macro(book: Any) -> None:
    """Add the P3 VBA entry point without altering the existing stage macros."""
    try:
        module = book.api.VBProject.VBComponents.Item("modBondsLiteLauncher")
        code = module.CodeModule
        source = code.Lines(1, code.CountOfLines) if code.CountOfLines else ""
    except Exception as error:
        raise PropertyStageThreeError(
            "Cannot access modBondsLiteLauncher in the Draft VBA project"
        ) from error
    if (
        'LaunchBondsLite "P3"'.casefold() in source.casefold()
        and "Public Sub PrepareNoteVal()".casefold() in source.casefold()
    ):
        return
    if (
        "Private Sub LaunchBondsLite(ByVal stage As String)".casefold()
        not in source.casefold()
    ):
        raise PropertyStageThreeError(
            "Draft launcher does not contain a supported LaunchBondsLite signature"
        )
    if "Public Sub PrepareNoteVal()".casefold() in source.casefold():
        raise PropertyStageThreeError(
            "Draft PrepareNoteVal macro has an unexpected implementation"
        )
    code.InsertLines(
        code.CountOfLines + 1,
        '\nPublic Sub PrepareNoteVal()\n    LaunchBondsLite "P3"\nEnd Sub\n',
    )


def install_property_stage_three_controls(caller: Path) -> Path:
    """Install Stage 3 controls and the entity name into a closed Draft backup."""
    import xlwings as xw

    caller = Path(caller).resolve(strict=True)
    for application in xw.apps:
        for book in application.books:
            if book.name.casefold() == caller.name.casefold():
                raise PropertyStageThreeError(
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
                raise PropertyStageThreeError("Draft workbook is read-only or locked")
            stage_one.require_property_stage_one_panel(book)
            stage_two.require_property_stage_two_panel(book)
            stage_one._entity_control(book)
            _initialise_property_stage_three_panel(book)
            _install_property_stage_three_button(book)
            _install_property_stage_three_macro(book)
            book.save()
        finally:
            book.close()
    finally:
        application.api.DisplayAlerts = False
        application.api.Quit()
    return backup
