import re
from collections.abc import Iterator
from pathlib import Path

import openpyxl
import xlrd
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openpyxl.utils.cell import range_boundaries

EXTENSIONS = {".xls", ".xlsx", ".xlsm"}
OLE_SIGNATURE = b"\xd0\xcf\x11\xe0"
MERGE_CELL = re.compile(rb'<(?:\w+:)?mergeCell\b[^>]*?\bref="([A-Z]+\d+(?::[A-Z]+\d+)?)"')

Merge = tuple[int, int, int, int]


def clean(value):
    if isinstance(value, str):
        return ILLEGAL_CHARACTERS_RE.sub("", value).strip() or None
    if isinstance(value, float) and value.is_integer():
        return int(value)
    return value


class XlsBook:
    def __init__(self, path: Path, formatting: bool):
        self._book = xlrd.open_workbook(str(path), formatting_info=formatting, on_demand=True)

    def __enter__(self):
        return self

    def __exit__(self, *_):
        self._book.release_resources()

    @property
    def sheet_names(self) -> list[str]:
        return self._book.sheet_names()

    def rows(self, sheet: str, start: int = 0, stop: int | None = None) -> Iterator[list]:
        source = self._book.sheet_by_name(sheet)
        end = source.nrows if stop is None else min(stop, source.nrows)
        for index in range(start, end):
            yield [self._value(cell) for cell in source.row(index)]

    def merges(self, sheet: str) -> list[Merge]:
        return [(r0, r1 - 1, c0, c1 - 1) for r0, r1, c0, c1 in self._book.sheet_by_name(sheet).merged_cells]

    def _value(self, cell):
        if cell.ctype == xlrd.XL_CELL_DATE:
            try:
                return xlrd.xldate_as_datetime(cell.value, self._book.datemode)
            except xlrd.xldate.XLDateError:
                return clean(cell.value)
        if cell.ctype == xlrd.XL_CELL_BOOLEAN:
            return bool(cell.value)
        if cell.ctype in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK, xlrd.XL_CELL_ERROR):
            return None
        return clean(cell.value)


class XlsxBook:
    def __init__(self, path: Path, formatting: bool):
        self._book = openpyxl.load_workbook(path, read_only=True, data_only=True)

    def __enter__(self):
        return self

    def __exit__(self, *_):
        self._book.close()

    @property
    def sheet_names(self) -> list[str]:
        return [sheet.title for sheet in self._book.worksheets]

    def rows(self, sheet: str, start: int = 0, stop: int | None = None) -> Iterator[list]:
        source = self._book[sheet]
        source.reset_dimensions()
        for row in source.iter_rows(min_row=start + 1, max_row=stop, values_only=True):
            yield [clean(value) for value in row]

    def merges(self, sheet: str) -> list[Merge]:
        with self._book[sheet]._get_source() as source:
            xml = source.read()
        found = []
        for ref in MERGE_CELL.findall(xml, max(xml.find(b"mergeCell"), 0)):
            c0, r0, c1, r1 = range_boundaries(ref.decode())
            found.append((r0 - 1, r1 - 1, c0 - 1, c1 - 1))
        return found


def open_book(path: Path, formatting: bool = False) -> XlsBook | XlsxBook:
    with path.open("rb") as file:
        legacy = file.read(4) == OLE_SIGNATURE
    return (XlsBook if legacy else XlsxBook)(path, formatting)
