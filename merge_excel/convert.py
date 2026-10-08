import re
from collections.abc import Callable
from datetime import date, datetime, time, timedelta

from .profile import ResultColumn
from .scanner import text

Transform = Callable[[object], object]
EXCEL_EPOCH = datetime(1899, 12, 30)
MAX_SERIAL = 2958465
YEARS = (1900, 2100)
NUMBER = re.compile(r"\d+(?:[.,]\d+)?")
COMPACT = re.compile(r"(\d{4})(\d{2})(\d{2})")
DATE_TIME = re.compile(
    r"(?P<date>.+?)(?:[T\s]+(?P<time>\d{1,2}:\d{2}(?::\d{2})?(?:[.,]\d+)?(?:\s*[AaPp][Mm])?)(?:\s*(?:Z|[+-]\d{2}:?\d{2}))?)?"
)
NUMERIC_DATE = re.compile(r"(\d{1,4})[./-](\d{1,2})[./-](\d{1,4})")
DAY_MONTH = re.compile(r"(\d{1,2})[\s./-]*([^\W\d_]{3,})\.?[\s./,-]*(\d{4}|\d{2})(?:\s*г\.?)?")
MONTH_DAY = re.compile(r"([^\W\d_]{3,})\.?\s*(\d{1,2})(?:st|nd|rd|th)?,?\s*(\d{4}|\d{2})")
CLOCK = re.compile(r"(\d{1,2}):(\d{2})(?::(\d{2}))?(?:[.,]\d+)?\s*([AaPp][Mm])?")
MONTHS = {
    name: number
    for number, names in enumerate(
        [("jan", "янв"), ("feb", "фев"), ("mar", "мар"), ("apr", "апр"), ("may", "май", "мая"), ("jun", "июн"),
         ("jul", "июл"), ("aug", "авг"), ("sep", "сен"), ("oct", "окт"), ("nov", "ноя"), ("dec", "дек")],
        1,
    )
    for name in names
}


def _year(digits: str) -> int:
    year = int(digits)
    if len(digits) == 2:
        return year + (2000 if year < 50 else 1900)
    return year


def _date(year: int, month: int | None, day: int) -> date | None:
    try:
        return date(year, month, day) if month else None
    except ValueError:
        return None


def _from_number(number: float) -> datetime | None:
    if number.is_integer() and YEARS[0] <= number <= YEARS[1]:
        return datetime(int(number), 1, 1)
    if number.is_integer() and (match := COMPACT.fullmatch(str(int(number)))):
        found = _date(int(match[1]), int(match[2]), int(match[3]))
        return datetime.combine(found, time()) if found else None
    if 1 <= number <= MAX_SERIAL:
        return EXCEL_EPOCH + timedelta(seconds=round(number * 86400))
    return None


def _parse_date(value: str) -> date | None:
    if match := NUMERIC_DATE.fullmatch(value):
        first, month, last = match.groups()
        if len(first) == 4:
            return _date(int(first), int(month), int(last))
        if len(last) not in (2, 4):
            return None
        day, month = int(first), int(month)
        if month > 12 >= day:
            day, month = month, day
        return _date(_year(last), month, day)
    if match := DAY_MONTH.fullmatch(value):
        return _date(_year(match[3]), MONTHS.get(match[2].casefold()[:3]), int(match[1]))
    if match := MONTH_DAY.fullmatch(value):
        return _date(_year(match[3]), MONTHS.get(match[1].casefold()[:3]), int(match[2]))
    return None


def _parse_time(value: str | None) -> time | None:
    if not value:
        return time()
    match = CLOCK.fullmatch(value)
    hour, minute, second = int(match[1]), int(match[2]), int(match[3] or 0)
    if match[4]:
        hour = hour % 12 + (12 if match[4].casefold() == "pm" else 0)
    return time(hour, minute, second) if hour < 24 and minute < 60 and second < 60 else None


def to_datetime(value) -> datetime | None:
    if isinstance(value, datetime):
        return value
    if isinstance(value, date):
        return datetime.combine(value, time())
    if isinstance(value, bool) or not isinstance(value, (int, float, str)):
        return None
    if isinstance(value, (int, float)):
        return _from_number(float(value))
    value = value.strip()
    if NUMBER.fullmatch(value):
        return _from_number(float(value.replace(",", ".")))
    match = DATE_TIME.fullmatch(value)
    if not match:
        return None
    day, clock = _parse_date(match["date"].strip()), _parse_time(match["time"])
    return datetime.combine(day, clock) if day and clock else None


def validate(column: ResultColumn) -> str:
    if column.extract in ("first", "last") and not (column.pattern.strip().isdigit() and int(column.pattern) > 0):
        return f"“{column.name}”: enter how many characters to keep."
    if column.extract == "regex":
        if not column.pattern:
            return f"“{column.name}”: enter a regular expression."
        try:
            re.compile(column.pattern)
        except re.error as error:
            return f"“{column.name}”: invalid regular expression — {error}."
    return ""


def transformer(column: ResultColumn) -> Transform:
    steps: list[Transform] = []
    if column.extract == "first":
        steps.append(lambda value, size=int(column.pattern): text(value)[:size] or None)
    elif column.extract == "last":
        steps.append(lambda value, size=int(column.pattern): text(value)[-size:] or None)
    elif column.extract == "regex":
        pattern = re.compile(column.pattern)

        def search(value):
            match = pattern.search(text(value))
            if not match:
                return None
            return (match.group(1) if pattern.groups else match.group(0)) or None

        steps.append(search)
    if column.date:
        steps.append(lambda value: to_datetime(value) or value)

    def transform(value):
        for step in steps:
            if value is not None:
                value = step(value)
        return value

    return transform
