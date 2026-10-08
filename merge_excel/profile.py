import json
import os
import uuid
from dataclasses import asdict, dataclass, field
from pathlib import Path

from .scanner import Template

PROFILE_PATH = Path(os.environ.get("APPDATA", Path.home())) / "MergeExcel" / "profile.json"
DEFAULT_SKIP = ["Повторн", "Итого", "Total"]


@dataclass
class ResultColumn:
    name: str
    join: bool = False
    separator: str = "\n"
    date: bool = False
    extract: str = ""
    pattern: str = ""
    id: str = field(default_factory=lambda: uuid.uuid4().hex[:8])


@dataclass
class Profile:
    folder: str = ""
    columns: list[ResultColumn] = field(default_factory=list)
    skip: list[str] = field(default_factory=lambda: list(DEFAULT_SKIP))
    links: dict[str, dict[str, str]] = field(default_factory=dict)
    unique: str = ""
    latest: bool = False
    latest_by: str = ""

    @classmethod
    def load(cls) -> "Profile":
        try:
            data = json.loads(PROFILE_PATH.read_text(encoding="utf-8"))
            return cls(
                folder=data["folder"],
                columns=[ResultColumn(**column) for column in data["columns"]],
                skip=data["skip"],
                links=data["links"],
                unique=data.get("unique", ""),
                latest=data.get("latest", False),
                latest_by=data.get("latest_by", ""),
            )
        except (OSError, ValueError, TypeError, KeyError):
            return cls()

    def column(self, column_id: str) -> ResultColumn | None:
        return next((column for column in self.columns if column.id == column_id), None)

    def save(self) -> None:
        PROFILE_PATH.parent.mkdir(parents=True, exist_ok=True)
        PROFILE_PATH.write_text(json.dumps(asdict(self), ensure_ascii=False, indent=2), encoding="utf-8")

    @property
    def skip_prefixes(self) -> tuple[str, ...]:
        return tuple(prefix.casefold() for prefix in self.skip if prefix)

    def is_auto(self, template: Template) -> bool:
        return template.key not in self.links

    def links_for(self, template: Template, templates: tuple[Template, ...]) -> dict[str, str]:
        valid = {column.id for column in self.columns}
        if template.key in self.links:
            return {key: target for key, target in self.links[template.key].items() if target in valid}
        known = {
            column.match: self.links[other.key][column.key]
            for other in templates
            if other.key in self.links
            for column in other.columns(self.skip_prefixes)
            if column.match and self.links[other.key].get(column.key) in valid
        }
        return {column.key: known[column.match] for column in template.columns(self.skip_prefixes) if column.match in known}

    def link(self, template: Template, templates: tuple[Template, ...], key: str, target: str | None) -> None:
        links = self.links.setdefault(template.key, self.links_for(template, templates))
        if target:
            links[key] = target
        else:
            links.pop(key, None)

    def remove_column(self, column: ResultColumn) -> None:
        self.columns.remove(column)
        for links in self.links.values():
            for key in [key for key, target in links.items() if target == column.id]:
                del links[key]
        if self.unique == column.id:
            self.unique = ""
        if self.latest_by == column.id:
            self.latest_by = ""
