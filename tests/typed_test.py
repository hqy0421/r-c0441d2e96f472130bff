import unittest
from datetime import date, datetime
from unittest.mock import Mock

from gspread.exceptions import InvalidInputValue, ValueConflictError
from gspread.http_client import HTTPClient
from gspread.typed import (
    CellKind,
    decode_cell_data,
    extract_typed_cells,
    normalize_column_indexes,
    parse_finite_a1_range,
    serial_to_date,
    serial_to_datetime,
)
from gspread.worksheet import Worksheet


def make_response(data_blocks, timezone="Asia/Shanghai", sheet_id=0):
    return {
        "properties": {"timeZone": timezone},
        "sheets": [{"properties": {"sheetId": sheet_id}, "data": data_blocks}],
    }


def make_worksheet(client):
    return Worksheet(
        spreadsheet=Mock(),
        properties={
            "sheetId": 0,
            "title": "Sheet1",
            "index": 0,
            "gridProperties": {"rowCount": 100, "columnCount": 26},
        },
        spreadsheet_id="spreadsheet-id",
        client=client,
    )


class SerialTest(unittest.TestCase):
    def test_serial_epoch(self):
        # Sheets/Lotus serial 1 is 1899-12-31
        self.assertEqual(serial_to_datetime(1), datetime(1899, 12, 31))
        self.assertEqual(serial_to_date(1), date(1899, 12, 31))

    def test_serial_with_time_part(self):
        self.assertEqual(serial_to_datetime(45000), datetime(2023, 3, 15))
        self.assertEqual(serial_to_datetime(45000.5), datetime(2023, 3, 15, 12, 0, 0))
        self.assertEqual(serial_to_date(45000.5), date(2023, 3, 15))


class DecodeCellDataTest(unittest.TestCase):
    def test_blank_cell(self):
        record = decode_cell_data({}, 1, 1)
        self.assertEqual(record.kind, CellKind.BLANK)
        self.assertIsNone(record.value)
        self.assertIsNone(record.input_value)
        self.assertIsNone(record.result_value)
        self.assertEqual(record.address, "A1")

    def test_explicit_empty_string(self):
        record = decode_cell_data({"userEnteredValue": {"stringValue": ""}}, 1, 1)
        self.assertEqual(record.kind, CellKind.EMPTY_STRING)
        self.assertEqual(record.value, "")
        self.assertEqual(record.input_value, "")
        self.assertEqual(record.result_value, "")

    def test_formula_empty_string(self):
        record = decode_cell_data(
            {
                "userEnteredValue": {"formulaValue": '=IF(FALSE,"x","")'},
                "effectiveValue": {"stringValue": ""},
            },
            1,
            1,
        )
        self.assertEqual(record.kind, CellKind.FORMULA_EMPTY_STRING)
        self.assertEqual(record.value, "")
        self.assertEqual(record.formula, '=IF(FALSE,"x","")')
        self.assertEqual(record.input_value, '=IF(FALSE,"x","")')
        self.assertEqual(record.result_value, "")

    def test_formula_error_keeps_type_formula_and_address(self):
        record = decode_cell_data(
            {
                "userEnteredValue": {"formulaValue": "=1/0"},
                "effectiveValue": {
                    "errorValue": {
                        "type": "DIVIDE_BY_ZERO",
                        "message": "Function DIVIDE parameter 2 cannot be zero.",
                    }
                },
            },
            2,
            2,
        )
        self.assertEqual(record.kind, CellKind.ERROR)
        self.assertIsNone(record.value)
        self.assertEqual(record.error_type, "DIVIDE_BY_ZERO")
        self.assertEqual(
            record.error_message, "Function DIVIDE parameter 2 cannot be zero."
        )
        self.assertEqual(record.formula, "=1/0")
        self.assertEqual(record.address, "B2")

    def test_zero_is_preserved(self):
        record = decode_cell_data(
            {
                "userEnteredValue": {"numberValue": 0},
                "effectiveValue": {"numberValue": 0},
            },
            1,
            1,
        )
        self.assertEqual(record.kind, CellKind.NUMBER)
        self.assertEqual(record.value, 0)
        self.assertIsNotNone(record.value)
        self.assertEqual(record.result_value, 0)

    def test_false_is_preserved(self):
        record = decode_cell_data(
            {
                "userEnteredValue": {"boolValue": False},
                "effectiveValue": {"boolValue": False},
            },
            1,
            1,
        )
        self.assertEqual(record.kind, CellKind.BOOL)
        self.assertEqual(record.value, False)
        self.assertIsNotNone(record.value)

    def test_text_001_is_preserved(self):
        record = decode_cell_data(
            {
                "userEnteredValue": {"stringValue": "001"},
                "effectiveValue": {"stringValue": "001"},
            },
            1,
            1,
        )
        self.assertEqual(record.kind, CellKind.STRING)
        self.assertEqual(record.value, "001")

    def test_date_column_is_decoded_from_serial(self):
        record = decode_cell_data(
            {
                "userEnteredValue": {"numberValue": 45000},
                "effectiveValue": {"numberValue": 45000},
            },
            1,
            1,
            date_columns={1},
        )
        self.assertEqual(record.kind, CellKind.DATE)
        self.assertEqual(record.value, date(2023, 3, 15))
        self.assertEqual(record.serial, 45000)
        self.assertEqual(record.input_value, 45000)
        self.assertEqual(record.result_value, 45000)

    def test_datetime_column_keeps_local_time_timezone_and_serial(self):
        record = decode_cell_data(
            {
                "userEnteredValue": {"numberValue": 45000.5},
                "effectiveValue": {"numberValue": 45000.5},
            },
            1,
            3,
            datetime_columns={3},
            timezone="Asia/Shanghai",
        )
        self.assertEqual(record.kind, CellKind.DATETIME)
        self.assertEqual(record.value, datetime(2023, 3, 15, 12, 0, 0))
        self.assertEqual(record.timezone, "Asia/Shanghai")
        self.assertEqual(record.serial, 45000.5)

    def test_number_outside_declared_columns_stays_a_number(self):
        record = decode_cell_data(
            {"userEnteredValue": {"numberValue": 45000}},
            1,
            2,
            date_columns={1},
            datetime_columns={3},
        )
        self.assertEqual(record.kind, CellKind.NUMBER)
        self.assertEqual(record.value, 45000)
        self.assertIsNone(record.serial)

    def test_formula_with_number_result(self):
        record = decode_cell_data(
            {
                "userEnteredValue": {"formulaValue": "=1+1"},
                "effectiveValue": {"numberValue": 2},
            },
            1,
            1,
        )
        self.assertEqual(record.kind, CellKind.NUMBER)
        self.assertEqual(record.value, 2)
        self.assertEqual(record.input_value, "=1+1")
        self.assertEqual(record.result_value, 2)
        self.assertEqual(record.formula, "=1+1")

    def test_formula_result_in_datetime_column_is_decoded(self):
        record = decode_cell_data(
            {
                "userEnteredValue": {"formulaValue": "=NOW()"},
                "effectiveValue": {"numberValue": 45000.5},
            },
            1,
            3,
            datetime_columns={3},
            timezone="Asia/Shanghai",
        )
        self.assertEqual(record.kind, CellKind.DATETIME)
        self.assertEqual(record.value, datetime(2023, 3, 15, 12, 0, 0))
        self.assertEqual(record.formula, "=NOW()")
        self.assertEqual(record.serial, 45000.5)


class ParseFiniteA1RangeTest(unittest.TestCase):
    def test_full_range(self):
        self.assertEqual(parse_finite_a1_range("A1:B2"), (0, 0, 2, 2))
        self.assertEqual(parse_finite_a1_range("B3:D10"), (2, 1, 10, 4))

    def test_single_cell(self):
        self.assertEqual(parse_finite_a1_range("B3"), (2, 1, 3, 2))

    def test_sheet_prefix_is_stripped(self):
        self.assertEqual(parse_finite_a1_range("Sheet1!A1:B2"), (0, 0, 2, 2))

    def test_unbounded_ranges_are_rejected(self):
        for range_name in ["A:B", "A1:B", "1:2", "A1:", ""]:
            with self.assertRaises(InvalidInputValue, msg=range_name):
                parse_finite_a1_range(range_name)

    def test_invalid_ranges_are_rejected(self):
        with self.assertRaises(InvalidInputValue):
            parse_finite_a1_range("not a range")


class NormalizeColumnIndexesTest(unittest.TestCase):
    def test_letters_and_indexes(self):
        self.assertEqual(normalize_column_indexes(["A", 3, "bc"]), {1, 3, 55})

    def test_none_and_empty(self):
        self.assertEqual(normalize_column_indexes(None), set())
        self.assertEqual(normalize_column_indexes([]), set())

    def test_invalid_values(self):
        with self.assertRaises(InvalidInputValue):
            normalize_column_indexes([0])
        with self.assertRaises(InvalidInputValue):
            normalize_column_indexes(["!@#$%^&"])


class ExtractTypedCellsTest(unittest.TestCase):
    def test_aligns_on_absolute_start_coordinates(self):
        response = make_response(
            [
                {
                    "startRow": 4,
                    "startColumn": 2,
                    "rowData": [{"values": [{"userEnteredValue": {"numberValue": 7}}]}],
                }
            ]
        )
        records = extract_typed_cells(response, 0, [(4, 2, 5, 3)])
        self.assertEqual(len(records), 1)
        record = records[0]
        self.assertEqual((record.row, record.col), (5, 3))
        self.assertEqual(record.address, "C5")
        self.assertEqual(record.value, 7)

    def test_missing_start_coordinates_default_to_origin(self):
        response = make_response(
            [{"rowData": [{"values": [{"userEnteredValue": {"numberValue": 7}}]}]}]
        )
        records = extract_typed_cells(response, 0, [(0, 0, 1, 1)])
        self.assertEqual(records[0].address, "A1")
        self.assertEqual(records[0].value, 7)

    def test_overlapping_ranges_are_returned_once(self):
        block_cell = {
            "userEnteredValue": {"stringValue": "x"},
            "effectiveValue": {"stringValue": "x"},
        }
        response = make_response(
            [
                {
                    "startRow": 0,
                    "startColumn": 0,
                    "rowData": [
                        {"values": [block_cell, block_cell]},
                        {"values": [block_cell, block_cell]},
                    ],
                },
                # second range overlaps B2 with identical data
                {
                    "startRow": 1,
                    "startColumn": 1,
                    "rowData": [{"values": [block_cell]}],
                },
            ]
        )
        records = extract_typed_cells(response, 0, [(0, 0, 2, 2), (1, 1, 2, 2)])
        self.assertEqual([r.address for r in records], ["A1", "B1", "A2", "B2"])

    def test_conflicting_values_raise(self):
        response = make_response(
            [
                {
                    "startRow": 0,
                    "startColumn": 0,
                    "rowData": [{"values": [{"userEnteredValue": {"numberValue": 1}}]}],
                },
                {
                    "startRow": 0,
                    "startColumn": 0,
                    "rowData": [{"values": [{"userEnteredValue": {"numberValue": 2}}]}],
                },
            ]
        )
        with self.assertRaises(ValueConflictError):
            extract_typed_cells(response, 0, [(0, 0, 1, 1), (0, 0, 1, 1)])

    def test_omitted_cells_are_padded_with_blanks(self):
        response = make_response(
            [
                {
                    "startRow": 0,
                    "startColumn": 0,
                    "rowData": [
                        # third cell of the row omitted
                        {
                            "values": [
                                {"userEnteredValue": {"numberValue": 1}},
                                {"userEnteredValue": {"numberValue": 2}},
                            ]
                        },
                        # trailing cells omitted
                        {"values": [{"userEnteredValue": {"numberValue": 3}}]},
                        # third row omitted entirely
                    ],
                }
            ]
        )
        records = extract_typed_cells(response, 0, [(0, 0, 3, 3)])
        self.assertEqual(len(records), 9)
        by_address = {r.address: r for r in records}
        self.assertEqual(by_address["A1"].value, 1)
        self.assertEqual(by_address["B1"].value, 2)
        self.assertEqual(by_address["A2"].value, 3)
        for address in ["C1", "B2", "C2", "A3", "B3", "C3"]:
            self.assertEqual(by_address[address].kind, CellKind.BLANK, address)
            self.assertIsNone(by_address[address].value, address)

    def test_range_without_data_block_is_all_blank(self):
        response = make_response([])
        records = extract_typed_cells(response, 0, [(4, 4, 6, 6)])
        self.assertEqual(len(records), 4)
        self.assertTrue(all(r.kind is CellKind.BLANK for r in records))
        self.assertEqual(records[0].address, "E5")

    def test_other_sheets_are_ignored(self):
        response = make_response(
            [
                {
                    "startRow": 0,
                    "startColumn": 0,
                    "rowData": [{"values": [{"userEnteredValue": {"numberValue": 1}}]}],
                }
            ],
            sheet_id=999,
        )
        records = extract_typed_cells(response, 0, [(0, 0, 1, 1)])
        self.assertEqual(records[0].kind, CellKind.BLANK)


class GetTypedRecordsTest(unittest.TestCase):
    def make_response(self):
        return {
            "properties": {"timeZone": "Asia/Shanghai"},
            "sheets": [
                {
                    "properties": {"sheetId": 0},
                    "data": [
                        {
                            # block for 'Sheet1'!A1:B2 — the API omitted the
                            # second (empty) row
                            "startRow": 0,
                            "startColumn": 0,
                            "rowData": [
                                {
                                    "values": [
                                        {
                                            "userEnteredValue": {"numberValue": 45000},
                                            "effectiveValue": {"numberValue": 45000},
                                        },
                                        {
                                            "userEnteredValue": {"stringValue": "001"},
                                            "effectiveValue": {"stringValue": "001"},
                                        },
                                    ]
                                },
                            ],
                        },
                        {
                            # block for 'Sheet1'!B2:C3
                            "startRow": 1,
                            "startColumn": 1,
                            "rowData": [
                                {
                                    "values": [
                                        {
                                            "userEnteredValue": {
                                                "formulaValue": "=1/0"
                                            },
                                            "effectiveValue": {
                                                "errorValue": {
                                                    "type": "DIVIDE_BY_ZERO",
                                                    "message": "div by 0",
                                                }
                                            },
                                        },
                                        {
                                            "userEnteredValue": {"boolValue": False},
                                            "effectiveValue": {"boolValue": False},
                                        },
                                    ]
                                },
                                {
                                    "values": [
                                        {
                                            "userEnteredValue": {
                                                "formulaValue": '=IF(FALSE,1,"")'
                                            },
                                            "effectiveValue": {"stringValue": ""},
                                        },
                                        {
                                            "userEnteredValue": {
                                                "numberValue": 45000.5
                                            },
                                            "effectiveValue": {"numberValue": 45000.5},
                                        },
                                    ]
                                },
                            ],
                        },
                    ],
                }
            ],
        }

    def test_get_typed_records(self):
        client = Mock(spec=HTTPClient)
        client.spreadsheets_get.return_value = self.make_response()
        worksheet = make_worksheet(client)

        records = worksheet.get_typed_records(
            ["A1:B2", "B2:C3"],
            date_columns=["A"],
            datetime_columns=["C"],
        )

        # a single API request for all ranges, asking for both the entered
        # values and the computed results
        client.spreadsheets_get.assert_called_once()
        args, kwargs = client.spreadsheets_get.call_args
        self.assertEqual(args[0], "spreadsheet-id")
        params = kwargs["params"]
        self.assertEqual(params["includeGridData"], "true")
        self.assertEqual(params["ranges"], ["'Sheet1'!A1:B2", "'Sheet1'!B2:C3"])
        self.assertIn("userEnteredValue", params["fields"])
        self.assertIn("effectiveValue", params["fields"])

        self.assertEqual(
            [r.address for r in records],
            ["A1", "B1", "A2", "B2", "C2", "B3", "C3"],
        )
        by_address = {r.address: r for r in records}

        a1 = by_address["A1"]
        self.assertEqual(a1.kind, CellKind.DATE)
        self.assertEqual(a1.value, date(2023, 3, 15))
        self.assertEqual(a1.serial, 45000)

        b1 = by_address["B1"]
        self.assertEqual(b1.kind, CellKind.STRING)
        self.assertEqual(b1.value, "001")

        # row omitted by the API is padded with a blank record
        a2 = by_address["A2"]
        self.assertEqual(a2.kind, CellKind.BLANK)
        self.assertIsNone(a2.value)

        b2 = by_address["B2"]
        self.assertEqual(b2.kind, CellKind.ERROR)
        self.assertEqual(b2.error_type, "DIVIDE_BY_ZERO")
        self.assertEqual(b2.formula, "=1/0")

        c2 = by_address["C2"]
        self.assertEqual(c2.kind, CellKind.BOOL)
        self.assertEqual(c2.value, False)

        b3 = by_address["B3"]
        self.assertEqual(b3.kind, CellKind.FORMULA_EMPTY_STRING)
        self.assertEqual(b3.value, "")

        c3 = by_address["C3"]
        self.assertEqual(c3.kind, CellKind.DATETIME)
        self.assertEqual(c3.value, datetime(2023, 3, 15, 12, 0, 0))
        self.assertEqual(c3.timezone, "Asia/Shanghai")
        self.assertEqual(c3.serial, 45000.5)

    def test_single_range_string_is_accepted(self):
        client = Mock(spec=HTTPClient)
        client.spreadsheets_get.return_value = make_response([])
        worksheet = make_worksheet(client)

        records = worksheet.get_typed_records("A1:A1")

        self.assertEqual(len(records), 1)
        params = client.spreadsheets_get.call_args[1]["params"]
        self.assertEqual(params["ranges"], ["'Sheet1'!A1:A1"])

    def test_empty_ranges_do_not_call_the_api(self):
        client = Mock(spec=HTTPClient)
        worksheet = make_worksheet(client)

        self.assertEqual(worksheet.get_typed_records([]), [])
        client.spreadsheets_get.assert_not_called()

    def test_unbounded_ranges_are_rejected_before_any_call(self):
        client = Mock(spec=HTTPClient)
        worksheet = make_worksheet(client)

        with self.assertRaises(InvalidInputValue):
            worksheet.get_typed_records(["A:B"])
        client.spreadsheets_get.assert_not_called()

    def test_conflicting_overlap_raises(self):
        client = Mock(spec=HTTPClient)
        client.spreadsheets_get.return_value = make_response(
            [
                {
                    "startRow": 0,
                    "startColumn": 0,
                    "rowData": [{"values": [{"userEnteredValue": {"numberValue": 1}}]}],
                },
                {
                    "startRow": 0,
                    "startColumn": 0,
                    "rowData": [{"values": [{"userEnteredValue": {"numberValue": 2}}]}],
                },
            ]
        )
        worksheet = make_worksheet(client)

        with self.assertRaises(ValueConflictError):
            worksheet.get_typed_records(["A1", "A1"])
