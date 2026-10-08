from datetime import date, datetime

import pytest

from merge_excel.convert import to_datetime, transformer, validate
from merge_excel.profile import ResultColumn


@pytest.mark.parametrize(
    ("value", "expected"),
    [
        (45839, datetime(2025, 7, 1)),
        (45839.5, datetime(2025, 7, 1, 12)),
        ("45839", datetime(2025, 7, 1)),
        ("45839,25", datetime(2025, 7, 1, 6)),
        (2003, datetime(2003, 1, 1)),
        ("2003", datetime(2003, 1, 1)),
        (20240603, datetime(2024, 6, 3)),
        ("20240603", datetime(2024, 6, 3)),
        ("03.06.2024 11:23:10", datetime(2024, 6, 3, 11, 23, 10)),
        ("3.6.2024 9:05", datetime(2024, 6, 3, 9, 5)),
        ("03/06/2024", datetime(2024, 6, 3)),
        ("06/13/2024", datetime(2024, 6, 13)),
        ("03-06-24", datetime(2024, 6, 3)),
        ("2024-06-03", datetime(2024, 6, 3)),
        ("2024-06-03T11:23:10Z", datetime(2024, 6, 3, 11, 23, 10)),
        ("2024/06/03 11:23", datetime(2024, 6, 3, 11, 23)),
        ("3 июня 2024 г.", datetime(2024, 6, 3)),
        ("15 мая 2025", datetime(2025, 5, 15)),
        ("03-Jun-2024", datetime(2024, 6, 3)),
        ("June 3, 2024", datetime(2024, 6, 3)),
        ("Jun 3rd 2024 10:30 PM", datetime(2024, 6, 3, 22, 30)),
        (date(2024, 6, 3), datetime(2024, 6, 3)),
        (datetime(2024, 6, 3, 1, 2, 3), datetime(2024, 6, 3, 1, 2, 3)),
        ("not a date", None),
        ("31.02.2024", None),
        ("50 217", None),
        (True, None),
        (0, None),
    ],
)
def test_to_datetime(value, expected):
    assert to_datetime(value) == expected


@pytest.mark.parametrize(
    ("column", "value", "expected"),
    [
        (ResultColumn("VIN", extract="last", pattern="17"), "Body No: JTMABABJ904074567", "JTMABABJ904074567"),
        (ResultColumn("Short", extract="first", pattern=" 5 "), "Hereketlendirijiniň", "Herek"),
        (ResultColumn("Order", extract="regex", pattern=r"TY\d+"), "Заявка на ремонт TY00000002 от 03.01.2023", "TY00000002"),
        (ResultColumn("Order", extract="regex", pattern=r"TY\d+"), "no order here", None),
        (ResultColumn("Opened", extract="regex", pattern=r"от (\S+)", date=True), "Заявка TY1 от 03.01.2023", datetime(2023, 1, 3)),
        (ResultColumn("Opened", date=True), "unknown", "unknown"),
        (ResultColumn("Plain"), 42, 42),
    ],
)
def test_transformer(column, value, expected):
    assert transformer(column)(value) == expected


@pytest.mark.parametrize(
    ("column", "message"),
    [
        (ResultColumn("Ok", extract="last", pattern="17"), ""),
        (ResultColumn("Bad", extract="last", pattern="x"), "“Bad”: enter how many characters to keep."),
        (ResultColumn("Bad", extract="regex", pattern=""), "“Bad”: enter a regular expression."),
    ],
)
def test_validate(column, message):
    assert validate(column) == message


def test_validate_reports_invalid_regex():
    assert validate(ResultColumn("Bad", extract="regex", pattern="(")).startswith("“Bad”: invalid regular expression")
