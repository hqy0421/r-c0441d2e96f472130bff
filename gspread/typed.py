"""
gspread.typed
~~~~~~~~~~~~~

Typed cell records decoded from the Sheets API grid data.

This module powers :meth:`gspread.worksheet.Worksheet.get_typed_records`.
It decodes the raw ``userEnteredValue`` / ``effectiveValue`` pairs returned
by ``spreadsheets.get`` into typed records suitable for data-warehouse
ingestion, keeping the distinction between genuinely blank cells, explicit
empty strings, formulas evaluating to an empty string and formula errors —
a distinction that formatted (displayed) values cannot provide.
"""

from dataclasses import dataclass
from datetime import date, datetime, timedelta
from typing import Any, Dict, Iterable, List, Mapping, Optional, Set, Tuple, Union

from .exceptions import IncorrectCellLabel, InvalidInputValue, ValueConflictError
from .utils import (
    StrEnum,
    a1_range_to_grid_range,
    column_letter_to_index,
    get_a1_from_absolute_range,
    rowcol_to_a1,
)


class CellKind(StrEnum):
    """The kind of a :class:`TypedCell`.

    Discriminates values that the formatted (displayed) cell values would
    render identically, e.g. a genuinely blank cell, an explicit empty
    string and a formula evaluating to an empty string.
    """

    #: genuinely blank cell: nothing was ever entered in it
    BLANK = "blank"

    #: explicit empty string: the user entered ``""`` as a literal value
    EMPTY_STRING = "empty_string"

    #: a formula whose computed result is an empty string
    FORMULA_EMPTY_STRING = "formula_empty_string"

    #: a formula whose evaluation failed
    ERROR = "error"

    #: a number (``0`` is kept as-is)
    NUMBER = "number"

    #: a boolean (``False`` is kept as-is)
    BOOL = "bool"

    #: a text value (text like ``"001"`` is kept as-is)
    STRING = "string"

    #: a number decoded as a date, in a column declared as a date column
    DATE = "date"

    #: a number decoded as a datetime, in a column declared as a datetime column
    DATETIME = "datetime"


#: The Sheets serial number epoch: serial ``1`` is ``1899-12-31``.
SHEETS_SERIAL_EPOCH = datetime(1899, 12, 30)


def serial_to_datetime(serial: Union[int, float]) -> datetime:
    """Decode a Sheets serial number to a naive :class:`datetime.datetime`
    in the spreadsheet's local time.

    No timezone conversion is applied: the returned datetime is the local
    wall-clock time as displayed in the spreadsheet.
    """
    return SHEETS_SERIAL_EPOCH + timedelta(days=serial)


def serial_to_date(serial: Union[int, float]) -> date:
    """Decode a Sheets serial number to a :class:`datetime.date` in the
    spreadsheet's local time."""
    return serial_to_datetime(serial).date()


@dataclass(frozen=True)
class TypedCell:
    """A single physical cell decoded from the Sheets API grid data.

    .. note::

       This class should never be instantiated manually.
       It is built by :meth:`gspread.worksheet.Worksheet.get_typed_records`.
    """

    #: 1-based absolute row index of the cell.
    row: int

    #: 1-based absolute column index of the cell.
    col: int

    #: The cell address in A1 notation, e.g. ``"B3"``.
    address: str

    #: The kind of the cell, see :class:`CellKind`.
    kind: CellKind

    #: The typed value of the cell:
    #:
    #: * ``None`` for ``BLANK`` and ``ERROR`` cells
    #: * ``""`` for ``EMPTY_STRING`` and ``FORMULA_EMPTY_STRING`` cells
    #: * ``bool`` for ``BOOL`` cells
    #: * ``int``/``float`` for ``NUMBER`` cells
    #: * ``str`` for ``STRING`` cells
    #: * :class:`datetime.date` for ``DATE`` cells
    #: * naive :class:`datetime.datetime` (spreadsheet local time) for
    #:   ``DATETIME`` cells
    value: Any = None

    #: The value as entered in the cell: the formula string (``"=..."``)
    #: for formula cells, the literal value otherwise, ``None`` for blank
    #: cells.
    input_value: Any = None

    #: The raw computed value of the cell as returned by the API — the
    #: serial number for ``DATE``/``DATETIME`` cells. ``None`` for blank
    #: and error cells.
    result_value: Any = None

    #: The formula of the cell (``"=..."``) or ``None``.
    formula: Optional[str] = None

    #: The error type for ``ERROR`` cells, e.g. ``"DIVIDE_BY_ZERO"``,
    #: ``"N_A"``, ``"REF"``.
    error_type: Optional[str] = None

    #: The error message for ``ERROR`` cells.
    error_message: Optional[str] = None

    #: The spreadsheet timezone (e.g. ``"Asia/Shanghai"``) for ``DATETIME``
    #: cells.
    timezone: Optional[str] = None

    #: The original Sheets serial number for ``DATE``/``DATETIME`` cells.
    serial: Optional[Union[int, float]] = None


def parse_finite_a1_range(range_name: str) -> Tuple[int, int, int, int]:
    """Parse a *finite* (fully bounded) A1 notation range.

    Accepts ranges like ``"A1:B2"`` and single cells like ``"A1"``.
    Unbounded ranges (``"A:B"``, ``"A1:B"``, ``"1:2"`` ...) and invalid
    ranges are rejected with :exc:`gspread.exceptions.InvalidInputValue`.

    :returns: a ``(start_row, start_column, end_row, end_column)`` tuple of
        0-based, half-open indexes.
    """
    a1 = get_a1_from_absolute_range(range_name)
    if ":" in a1:
        start_label, _, end_label = a1.partition(":")
        if not start_label or not end_label:
            raise InvalidInputValue(
                "range {!r} is unbounded, ranges must be finite A1 notations like 'A1:B2'".format(
                    range_name
                )
            )
    try:
        grid_range = a1_range_to_grid_range(a1)
    except IncorrectCellLabel:
        raise InvalidInputValue(
            "invalid range: {!r}, ranges must be finite A1 notations like 'A1:B2'".format(
                range_name
            )
        )

    try:
        return (
            grid_range["startRowIndex"],
            grid_range["startColumnIndex"],
            grid_range["endRowIndex"],
            grid_range["endColumnIndex"],
        )
    except KeyError:
        raise InvalidInputValue(
            "range {!r} is unbounded, ranges must be finite A1 notations like 'A1:B2'".format(
                range_name
            )
        )


def normalize_column_indexes(columns: Optional[Iterable[Union[str, int]]]) -> Set[int]:
    """Normalize a column declaration to a set of 1-based column indexes.

    Accepts column letters (``"A"``, ``"BC"``) and 1-based indexes
    (``1``, ``28``) mixed together.
    """
    normalized: Set[int] = set()
    for column in columns or []:
        if isinstance(column, int):
            if column < 1:
                raise InvalidInputValue(
                    "invalid column index: {!r}, must be >= 1".format(column)
                )
            normalized.add(column)
        else:
            normalized.add(column_letter_to_index(str(column)))
    return normalized


def decode_cell_data(
    cell_data: Mapping[str, Any],
    row: int,
    col: int,
    date_columns: Optional[Set[int]] = None,
    datetime_columns: Optional[Set[int]] = None,
    timezone: Optional[str] = None,
) -> TypedCell:
    """Decode one ``CellData`` mapping from the Sheets API into a
    :class:`TypedCell`.

    :param cell_data: the raw ``CellData`` mapping, holding optional
        ``userEnteredValue`` and ``effectiveValue`` extended values.
    :param row: 1-based absolute row index of the cell.
    :param col: 1-based absolute column index of the cell.
    :param date_columns: 1-based indexes of the columns declared as dates.
    :param datetime_columns: 1-based indexes of the columns declared as
        datetimes.
    :param timezone: the spreadsheet timezone, attached to ``DATETIME``
        records.
    """
    date_columns = date_columns or set()
    datetime_columns = datetime_columns or set()
    address = rowcol_to_a1(row, col)

    user_entered = cell_data.get("userEnteredValue")
    effective = cell_data.get("effectiveValue")

    formula = user_entered.get("formulaValue") if user_entered else None

    # formula errors: keep the error type, the formula and the source address
    if effective and "errorValue" in effective:
        error = effective["errorValue"]
        return TypedCell(
            row=row,
            col=col,
            address=address,
            kind=CellKind.ERROR,
            input_value=formula,
            formula=formula,
            error_type=error.get("type"),
            error_message=error.get("message"),
        )

    if not user_entered:
        # no user entered value: the cell is genuinely blank
        return TypedCell(row=row, col=col, address=address, kind=CellKind.BLANK)

    if formula is not None:
        # formula cell: the input is the formula, the value is the computed result
        return _decode_extended_value(
            effective or {},
            row=row,
            col=col,
            address=address,
            date_columns=date_columns,
            datetime_columns=datetime_columns,
            timezone=timezone,
            formula=formula,
        )

    # literal cell: the input and the result are the entered value
    return _decode_extended_value(
        user_entered,
        row=row,
        col=col,
        address=address,
        date_columns=date_columns,
        datetime_columns=datetime_columns,
        timezone=timezone,
        formula=None,
    )


def _decode_extended_value(
    extended_value: Mapping[str, Any],
    row: int,
    col: int,
    address: str,
    date_columns: Set[int],
    datetime_columns: Set[int],
    timezone: Optional[str],
    formula: Optional[str],
) -> TypedCell:
    """Decode one ``ExtendedValue`` mapping into a :class:`TypedCell`.

    ``formula`` is the cell formula for formula cells (their input value),
    ``None`` for literal cells.
    """
    serial: Optional[Union[int, float]] = None
    cell_timezone: Optional[str] = None

    if "stringValue" in extended_value:
        string = extended_value["stringValue"]
        if string == "":
            kind = CellKind.FORMULA_EMPTY_STRING if formula else CellKind.EMPTY_STRING
        else:
            kind = CellKind.STRING
        result = value = string
    elif "boolValue" in extended_value:
        kind = CellKind.BOOL
        result = value = extended_value["boolValue"]
    elif "numberValue" in extended_value:
        number = extended_value["numberValue"]
        result = number
        if col in datetime_columns:
            kind = CellKind.DATETIME
            value = serial_to_datetime(number)
            serial = number
            cell_timezone = timezone
        elif col in date_columns:
            kind = CellKind.DATE
            value = serial_to_date(number)
            serial = number
        else:
            kind = CellKind.NUMBER
            value = number
    else:
        # no value at all (e.g. a formula without an effective value): blank
        return TypedCell(
            row=row,
            col=col,
            address=address,
            kind=CellKind.BLANK,
            input_value=formula,
            formula=formula,
        )

    return TypedCell(
        row=row,
        col=col,
        address=address,
        kind=kind,
        value=value,
        input_value=formula if formula is not None else result,
        result_value=result,
        formula=formula,
        serial=serial,
        timezone=cell_timezone,
    )


def extract_typed_cells(
    spreadsheet: Mapping[str, Any],
    sheet_id: int,
    rectangles: Iterable[Tuple[int, int, int, int]],
    date_columns: Optional[Set[int]] = None,
    datetime_columns: Optional[Set[int]] = None,
) -> List[TypedCell]:
    """Build the typed cell records from a ``spreadsheets.get`` response.

    Cells are aligned on the absolute ``startRow``/``startColumn``
    coordinates returned by the API for each grid data block. Cells from
    overlapping ranges are only returned once — conflicting values for the
    same physical cell raise :exc:`gspread.exceptions.ValueConflictError`.
    Cells omitted by the API (trailing empty rows/columns) inside the
    requested rectangles are filled in as blank records.

    :param spreadsheet: the raw ``spreadsheets.get`` response body.
    :param sheet_id: the ID of the worksheet to extract.
    :param rectangles: the requested ranges as ``(start_row, start_column,
        end_row, end_column)`` tuples of 0-based, half-open indexes.
    :param date_columns: 1-based indexes of the columns declared as dates.
    :param datetime_columns: 1-based indexes of the columns declared as
        datetimes.
    :returns: the typed records, sorted by row then column.
    """
    timezone = spreadsheet.get("properties", {}).get("timeZone")

    cells: Dict[Tuple[int, int], TypedCell] = {}

    for sheet in spreadsheet.get("sheets", []):
        if sheet.get("properties", {}).get("sheetId") != sheet_id:
            continue
        for grid_data in sheet.get("data", []):
            _collect_grid_data(
                grid_data,
                cells,
                date_columns=date_columns,
                datetime_columns=datetime_columns,
                timezone=timezone,
            )

    # the API omits trailing empty rows and cells: fill the requested
    # rectangles with blank records
    for start_row, start_column, end_row, end_column in rectangles:
        for row in range(start_row + 1, end_row + 1):
            for col in range(start_column + 1, end_column + 1):
                if (row, col) not in cells:
                    cells[(row, col)] = TypedCell(
                        row=row,
                        col=col,
                        address=rowcol_to_a1(row, col),
                        kind=CellKind.BLANK,
                    )

    return [cells[key] for key in sorted(cells)]


def _collect_grid_data(
    grid_data: Mapping[str, Any],
    cells: Dict[Tuple[int, int], TypedCell],
    date_columns: Optional[Set[int]],
    datetime_columns: Optional[Set[int]],
    timezone: Optional[str],
) -> None:
    """Decode one grid data block into ``cells``, keyed by absolute
    ``(row, col)``.

    Cells are aligned on the absolute ``startRow``/``startColumn``
    coordinates returned by the API. A physical cell already collected from
    an overlapping range is kept once; a conflicting value raises
    :exc:`gspread.exceptions.ValueConflictError`.
    """
    start_row = grid_data.get("startRow", 0)
    start_column = grid_data.get("startColumn", 0)
    for row_offset, row_data in enumerate(grid_data.get("rowData", [])):
        for col_offset, cell_data in enumerate(row_data.get("values", [])):
            row = start_row + row_offset + 1
            col = start_column + col_offset + 1
            record = decode_cell_data(
                cell_data,
                row,
                col,
                date_columns=date_columns,
                datetime_columns=datetime_columns,
                timezone=timezone,
            )
            existing = cells.get((row, col))
            if existing is not None:
                # overlapping ranges: the same physical cell is only
                # kept once, conflicting values are an error
                if existing != record:
                    raise ValueConflictError(
                        "conflicting values for cell {}: {!r} != {!r}".format(
                            record.address, existing, record
                        )
                    )
                continue
            cells[(row, col)] = record
