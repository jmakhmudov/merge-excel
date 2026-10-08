import hashlib
import re
from collections.abc import Callable
from dataclasses import dataclass
from datetime import date, datetime
from functools import cache
from pathlib import Path

from openpyxl.utils import get_column_letter

from .reader import EXTENSIONS, Merge, open_book

SCAN_DEPTH = 30
MIN_HEADER_CELLS = 3
SAMPLE_COUNT = 3
FILE_KEY = "file"
SHEET_KEY = "sheet"
NUMERIC = re.compile(r"[\d\s.,:;/+\-()%]+")
DIGIT = re.compile(r"\d")

Progress = Callable[[int, int, str], None]


def norm(text: str) -> str:
    return " ".join(text.split()).casefold()


def text(value) -> str:
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.strftime("%d.%m.%Y %H:%M:%S" if value.time() != datetime.min.time() else "%d.%m.%Y")
    if isinstance(value, date):
        return value.strftime("%d.%m.%Y")
    return str(value)


def is_skipped(key, prefixes: tuple[str, ...]) -> bool:
    return isinstance(key, str) and key.casefold().startswith(prefixes)


def _is_label(value) -> bool:
    return isinstance(value, str) and not NUMERIC.fullmatch(value)


@dataclass(frozen=True)
class Header:
    start: int
    end: int
    body: int
    left: int
    names: tuple[str, ...]
    numbered: bool


def _expand(rows: list[list], merges: list[Merge]) -> list[list]:
    width = max(map(len, rows), default=0)
    grid = [row + [None] * (width - len(row)) for row in rows]
    for r0, r1, c0, c1 in merges:
        if r0 < len(grid) and c0 < width:
            value = grid[r0][c0]
            for r in range(r0, min(r1, len(grid) - 1) + 1):
                grid[r][c0:min(c1, width - 1) + 1] = [value] * (min(c1, width - 1) + 1 - c0)
    return grid


def _merged_columns(merges: list[Merge], row: int) -> set[int]:
    return {c for r0, r1, c0, c1 in merges if r0 <= row <= r1 and c1 > c0 for c in range(c0, c1 + 1)}


def _is_group_row(grid: list[list], merges: list[Merge], row: int) -> bool:
    filled = [value for value in grid[row] if value is not None]
    return bool(_merged_columns(merges, row)) and len(set(filled)) >= 2 and all(map(_is_label, filled))


def _extend_down(grid: list[list], merges: list[Merge], start: int, main: int) -> int:
    end = main
    grown = True
    while grown:
        grown = False
        for r0, r1, _, _ in merges:
            if start <= r0 <= end < r1 < len(grid):
                end, grown = r1, True
    parents = _merged_columns(merges, main)
    if end == main and parents and main + 1 < len(grid):
        below = {c: value for c, value in enumerate(grid[main + 1]) if value is not None}
        if below and set(below) <= parents and all(map(_is_label, below.values())):
            end += 1
    return end


def _is_typed(values: list) -> bool:
    return sum(not _is_label(value) for value in values) * 2 > len(values)


def _is_description(row: list, below: list[list], columns: range) -> bool:
    filled = {c: row[c] for c in columns if row[c] is not None}
    if len(set(filled.values())) < MIN_HEADER_CELLS or len(filled) * 2 < len(columns):
        return False
    if not all(_is_label(value) for value in row if value is not None):
        return False
    data = {c: [line[c] for line in below if line[c] is not None] for c in columns}
    if not any(data.values()):
        return not has_digit(row[columns.start])
    typed = [c for c, values in data.items() if _is_typed(values)]
    return bool(typed) and sum(c in filled for c in typed) * 2 >= len(typed)


def has_digit(value) -> bool:
    return bool(DIGIT.search(text(value)))


def _is_banner(row: list, columns: range) -> bool:
    key = row[columns.start]
    return key is not None and not has_digit(key) and all(_is_label(row[c]) for c in columns if row[c] is not None)


def _column_name(block: list[list], column: int) -> str:
    parts: list[str] = []
    for row in block:
        label = " ".join(text(row[column]).split())
        if label and label not in parts:
            parts.append(label)
    return " / ".join(parts)


def detect_header(rows: list[list], merges: list[Merge]) -> Header | None:
    grid = _expand(rows, merges)
    scores = [len({value for value in row if _is_label(value)}) for row in grid]
    if max(scores, default=0) < MIN_HEADER_CELLS:
        return None
    main = scores.index(max(scores))
    start = main
    while start and _is_group_row(grid, merges, start - 1):
        start -= 1
    end = _extend_down(grid, merges, start, main)
    block = grid[start:end + 1]
    named = [c for c in range(len(grid[0])) if any(row[c] is not None for row in block)]
    columns = range(named[0], named[-1] + 1)
    body = end + 1
    while body < len(grid) and _is_description(grid[body], grid[body + 1:], columns):
        body += 1
    keys = [row[columns.start] for row in grid[body:] if row[columns.start] is not None]
    numbered = bool(keys) and sum(map(has_digit, keys)) * 2 >= len(keys)
    while numbered and body < len(grid) and _is_banner(grid[body], columns):
        body += 1
    return Header(start, end, body, named[0], tuple(_column_name(block, c) for c in columns), numbered)


@dataclass(frozen=True)
class Column:
    key: str
    letter: str
    name: str
    match: str
    samples: tuple[str, ...]


@dataclass(frozen=True)
class Sheet:
    path: Path
    name: str
    header: Header
    width: int
    sample_rows: tuple[tuple, ...]


@dataclass(frozen=True, eq=False)
class Template:
    key: str
    label: str
    names: tuple[str, ...]
    sheets: tuple[Sheet, ...]

    @property
    def width(self) -> int:
        return max(sheet.width for sheet in self.sheets)

    @cache
    def columns(self, skip: tuple[str, ...]) -> tuple[Column, ...]:
        rows = [row for sheet in self.sheets for row in sheet.sample_rows if not is_skipped(row[0] if row else None, skip)]
        left = self.sheets[0].header.left
        repeated = {norm(name) for name in self.names if self.names.count(name) > 1}
        result = []
        for index in range(self.width):
            name = self.names[index] if index < len(self.names) else ""
            samples = tuple(dict.fromkeys(text(row[index]) for row in rows if index < len(row) and row[index] is not None))
            if name or samples:
                match = "" if not name or norm(name) in repeated else norm(name)
                result.append(Column(str(index), get_column_letter(left + index + 1), name, match, samples[:SAMPLE_COUNT]))
        result.append(Column(FILE_KEY, "", "‹File name›", f"\0{FILE_KEY}", (self.sheets[0].path.name,)))
        result.append(Column(SHEET_KEY, "", "‹Sheet name›", f"\0{SHEET_KEY}", (self.sheets[0].name,)))
        return tuple(result)


@dataclass(frozen=True)
class Ignored:
    path: Path
    sheet: str
    reason: str


@dataclass(frozen=True)
class Scan:
    files: tuple[Path, ...]
    templates: tuple[Template, ...]
    ignored: tuple[Ignored, ...]


def _read_sheet(book, path: Path, name: str) -> Sheet | Ignored:
    rows = list(book.rows(name, 0, SCAN_DEPTH))
    header = detect_header(rows, book.merges(name)) if rows else None
    if header is None:
        filled = any(value is not None for row in rows for value in row)
        return Ignored(path, name, "no header found" if filled else "empty")
    data = [row[header.left:] for row in rows[header.body:]]
    last = max((i for row in data for i, value in enumerate(row) if value is not None), default=-1)
    return Sheet(path, name, header, max(len(header.names), last + 1), tuple(map(tuple, data)))


def scan(folder: Path, progress: Progress) -> Scan:
    files = sorted(
        (p for p in folder.iterdir() if p.is_file() and p.suffix.lower() in EXTENSIONS and not p.name.startswith("~$")),
        key=lambda p: p.name.casefold(),
    )
    groups: dict[tuple[str, ...], list[Sheet]] = {}
    ignored: list[Ignored] = []
    for done, path in enumerate(files):
        progress(done, len(files), path.name)
        try:
            with open_book(path, formatting=True) as book:
                for name in book.sheet_names:
                    found = _read_sheet(book, path, name)
                    if isinstance(found, Ignored):
                        ignored.append(found)
                    else:
                        groups.setdefault(tuple(map(norm, found.header.names)), []).append(found)
        except Exception as error:
            ignored.append(Ignored(path, "", f"cannot read: {error}"))
    ordered = sorted(groups.items(), key=lambda item: (-len(item[1]), item[1][0].path.name.casefold()))
    templates = tuple(
        Template(hashlib.sha1("\x1f".join(signature).encode()).hexdigest()[:16], f"T{number}", sheets[0].header.names, tuple(sheets))
        for number, (signature, sheets) in enumerate(ordered, 1)
    )
    return Scan(tuple(files), templates, tuple(ignored))
