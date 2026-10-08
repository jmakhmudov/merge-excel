import sys
from contextlib import suppress
from pathlib import Path

from PySide6.QtCore import QThread, Qt, QUrl, Signal
from PySide6.QtGui import QDesktopServices
from PySide6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QCheckBox,
    QComboBox,
    QFileDialog,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QProgressDialog,
    QPushButton,
    QSplitter,
    QStackedWidget,
    QStyle,
    QTableWidget,
    QTableWidgetItem,
    QTreeWidget,
    QTreeWidgetItem,
    QVBoxLayout,
    QWidget,
)

from .convert import validate
from .merger import Output, export
from .profile import Profile, ResultColumn
from .scanner import Scan, Template, scan

TITLE = "Merge Excel"
SEPARATORS = {"New line": "\n", "Semicolon": "; ", "Comma": ", ", "Space": " ", "Slash": " / "}
EXTRACTS = {"Whole value": "", "First characters": "first", "Last characters": "last", "Regular expression": "regex"}
EXTRACT_HINTS = {"": "", "first": "How many, e.g. 17", "last": "How many, e.g. 17", "regex": r"e.g. TY\d+  or  VIN: (\w+)"}
NAME, JOIN, SEPARATOR, DATE, EXTRACT, VALUE = range(6)


class Task(QThread):
    progressed = Signal(int, int, str)
    succeeded = Signal(object)
    failed = Signal(str)

    def __init__(self, job, parent):
        super().__init__(parent)
        self._job = job

    def run(self):
        try:
            self.succeeded.emit(self._job(self.progressed.emit))
        except Exception as error:
            self.failed.emit(str(error) or type(error).__name__)


class ComboBox(QComboBox):
    def __init__(self):
        super().__init__()
        self.setFocusPolicy(Qt.StrongFocus)

    def wheelEvent(self, event):
        if self.hasFocus():
            super().wheelEvent(event)
        else:
            event.ignore()


def _table(headers: list[str]) -> QTableWidget:
    table = QTableWidget(0, len(headers))
    table.setHorizontalHeaderLabels(headers)
    table.verticalHeader().setVisible(False)
    table.setSelectionBehavior(QAbstractItemView.SelectRows)
    table.setSelectionMode(QAbstractItemView.SingleSelection)
    return table


def _row(*widgets, stretch_at: int | None = None) -> QHBoxLayout:
    layout = QHBoxLayout()
    for index, widget in enumerate(widgets):
        if index == stretch_at:
            layout.addStretch()
        layout.addWidget(widget)
    if stretch_at == len(widgets):
        layout.addStretch()
    return layout


def _count(number: int, noun: str) -> str:
    return f"{number} {noun}{'' if number == 1 else 's'}"


def _check_item(checked: bool) -> QTableWidgetItem:
    item = QTableWidgetItem()
    item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsUserCheckable)
    item.setCheckState(Qt.Checked if checked else Qt.Unchecked)
    return item


def _index(options: dict[str, str], value: str) -> int:
    values = list(options.values())
    return values.index(value) if value in values else 0


def _button(text: str, slot) -> QPushButton:
    button = QPushButton(text)
    button.clicked.connect(slot)
    return button


class SetupPage(QWidget):
    scan_requested = Signal()

    def __init__(self, profile: Profile):
        super().__init__()
        self._profile = profile
        self._folder = QLineEdit(profile.folder)
        self._folder.setPlaceholderText("Folder with Excel files")
        self._skip = QLineEdit(", ".join(profile.skip))
        self._table = _table(["Result column", "Join rows", "Separator", "Date", "Extract", "Extract value"])
        for index in range(self._table.columnCount()):
            stretch = index in (NAME, VALUE)
            self._table.horizontalHeader().setSectionResizeMode(index, QHeaderView.Stretch if stretch else QHeaderView.ResizeToContents)
        self._table.itemChanged.connect(self._item_changed)

        source = QGroupBox("1. Source folder")
        source.setLayout(_row(self._folder, _button("Browse…", self._browse)))

        columns = QGroupBox("2. Result columns")
        columns_layout = QVBoxLayout(columns)
        columns_layout.addWidget(QLabel(
            "Join rows — combine values from the rows below a record (rows whose first column is empty).\n"
            "Date — turn serial numbers (45839), text dates and years into real dates.\n"
            "Extract (optional) — keep only part of each value: the first or last N characters, "
            "or what a regular expression matches (its first group, if it has one)."
        ))
        columns_layout.addWidget(self._table)
        columns_layout.addLayout(_row(
            _button("Add", self._add),
            _button("Remove", self._remove),
            _button("Move up", lambda: self._move(-1)),
            _button("Move down", lambda: self._move(1)),
            stretch_at=4,
        ))

        self._unique = ComboBox()
        self._latest = QCheckBox("Keep the latest record for each value, by the date in")
        self._latest_by = ComboBox()
        for box in (self._unique, self._latest_by):
            box.setSizeAdjustPolicy(QComboBox.AdjustToContents)
        self._unique.currentIndexChanged.connect(self._extras_changed)
        self._latest.toggled.connect(self._extras_changed)
        self._latest_by.currentIndexChanged.connect(self._extras_changed)
        extras = QGroupBox("3. Extra files (optional)")
        extras.setLayout(_row(QLabel("Unique values of"), self._unique, self._latest, self._latest_by, stretch_at=4))

        skip = QGroupBox("Skip rows whose first column starts with (comma separated)")
        skip.setLayout(_row(self._skip))

        layout = QVBoxLayout(self)
        layout.addWidget(source)
        layout.addWidget(columns, 1)
        layout.addWidget(extras)
        layout.addWidget(skip)
        layout.addLayout(_row(_button("Scan headers →", self._request_scan), stretch_at=0))
        self._render()

    def _render(self, select: int | None = None) -> None:
        self._table.blockSignals(True)
        self._table.setRowCount(len(self._profile.columns))
        for row, column in enumerate(self._profile.columns):
            self._table.setItem(row, NAME, QTableWidgetItem(column.name))
            self._table.setItem(row, JOIN, _check_item(column.join))
            self._table.setItem(row, DATE, _check_item(column.date))
            separator = ComboBox()
            separator.addItems(list(SEPARATORS))
            separator.setCurrentIndex(_index(SEPARATORS, column.separator))
            separator.setEnabled(column.join)
            separator.currentTextChanged.connect(lambda label, target=column: setattr(target, "separator", SEPARATORS[label]))
            self._table.setCellWidget(row, SEPARATOR, separator)
            value = QLineEdit(column.pattern)
            value.textChanged.connect(lambda pattern, target=column: setattr(target, "pattern", pattern))
            extract = ComboBox()
            extract.addItems(list(EXTRACTS))
            extract.setCurrentIndex(_index(EXTRACTS, column.extract))
            extract.currentTextChanged.connect(lambda label, target=column, field=value: self._set_extract(target, field, EXTRACTS[label]))
            self._set_extract(column, value, column.extract)
            self._table.setCellWidget(row, EXTRACT, extract)
            self._table.setCellWidget(row, VALUE, value)
        self._table.blockSignals(False)
        if select is not None:
            self._table.selectRow(select)
        self._render_extras()

    def _render_extras(self) -> None:
        profile = self._profile
        if not profile.column(profile.latest_by):
            profile.latest_by = next((c.id for c in profile.columns if c.date), profile.columns[0].id if profile.columns else "")
        for box, selected, empty in ((self._unique, profile.unique, "— none —"), (self._latest_by, profile.latest_by, None)):
            box.blockSignals(True)
            box.clear()
            if empty:
                box.addItem(empty, "")
            for column in profile.columns:
                box.addItem(column.name, column.id)
            box.setCurrentIndex(max(box.findData(selected), 0))
            box.blockSignals(False)
        self._latest.blockSignals(True)
        self._latest.setChecked(profile.latest)
        self._latest.blockSignals(False)
        self._extras_changed()

    def _extras_changed(self) -> None:
        self._profile.unique = self._unique.currentData() or ""
        self._profile.latest = self._latest.isChecked()
        self._profile.latest_by = self._latest_by.currentData() or ""
        self._latest.setEnabled(bool(self._profile.unique))
        self._latest_by.setEnabled(bool(self._profile.unique) and self._profile.latest)

    @staticmethod
    def _set_extract(column: ResultColumn, field: QLineEdit, mode: str) -> None:
        column.extract = mode
        field.setEnabled(bool(mode))
        field.setPlaceholderText(EXTRACT_HINTS[mode])

    def _item_changed(self, item: QTableWidgetItem) -> None:
        column = self._profile.columns[item.row()]
        checked = item.checkState() == Qt.Checked
        if item.column() == NAME:
            column.name = item.text().strip()
            self._render_extras()
        elif item.column() == JOIN:
            column.join = checked
            self._table.cellWidget(item.row(), SEPARATOR).setEnabled(checked)
        elif item.column() == DATE:
            column.date = checked

    def _browse(self) -> None:
        folder = QFileDialog.getExistingDirectory(self, "Select folder", self._folder.text())
        if folder:
            self._folder.setText(str(Path(folder)))

    def _add(self) -> None:
        self._profile.columns.append(ResultColumn(f"Column {len(self._profile.columns) + 1}"))
        row = len(self._profile.columns) - 1
        self._render(row)
        self._table.editItem(self._table.item(row, 0))

    def _remove(self) -> None:
        row = self._table.currentRow()
        if row >= 0:
            self._profile.remove_column(self._profile.columns[row])
            self._render(min(row, len(self._profile.columns) - 1))

    def _move(self, step: int) -> None:
        row, columns = self._table.currentRow(), self._profile.columns
        if row >= 0 and 0 <= row + step < len(columns):
            columns[row], columns[row + step] = columns[row + step], columns[row]
            self._render(row + step)

    def _request_scan(self) -> None:
        folder = self._folder.text().strip()
        names = [column.name.casefold() for column in self._profile.columns]
        invalid = [problem for problem in map(validate, self._profile.columns) if problem]
        if not folder or not Path(folder).is_dir():
            problem = "Choose an existing folder."
        elif not names:
            problem = "Add at least one result column."
        elif "" in names or len(set(names)) != len(names):
            problem = "Result column names must be filled in and unique."
        elif invalid:
            problem = "\n".join(invalid)
        else:
            self._profile.folder = folder
            self._profile.skip = [prefix.strip() for prefix in self._skip.text().split(",") if prefix.strip()]
            self.scan_requested.emit()
            return
        QMessageBox.warning(self, TITLE, problem)


class LinkPage(QWidget):
    back_requested = Signal()
    generate_requested = Signal()

    def __init__(self, profile: Profile):
        super().__init__()
        self._profile = profile
        self._scan: Scan | None = None
        self._shown: Template | None = None
        self._tree = QTreeWidget()
        self._tree.setHeaderLabels(["Template / sheet", "Status"])
        self._tree.header().setSectionResizeMode(0, QHeaderView.Stretch)
        self._tree.header().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self._tree.header().setStretchLastSection(False)
        self._tree.currentItemChanged.connect(self._select)
        self._caption = QLabel()
        self._caption.setWordWrap(True)
        self._table = _table(["Column", "Header", "Example", "Result column"])
        self._table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self._table.setWordWrap(False)
        header = self._table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.Stretch)
        header.setSectionResizeMode(2, QHeaderView.Stretch)
        header.setSectionResizeMode(3, QHeaderView.ResizeToContents)
        self._summary = QLabel()

        details = QWidget()
        details_layout = QVBoxLayout(details)
        details_layout.setContentsMargins(0, 0, 0, 0)
        details_layout.addWidget(self._caption)
        details_layout.addWidget(self._table)
        splitter = QSplitter()
        splitter.addWidget(self._tree)
        splitter.addWidget(details)
        splitter.setSizes([420, 780])

        layout = QVBoxLayout(self)
        layout.addWidget(QLabel("4. Templates found — select one and link its columns to your result columns."))
        layout.addWidget(splitter, 1)
        layout.addLayout(_row(
            _button("← Back", self.back_requested),
            self._summary,
            _button("Generate…", self.generate_requested),
            stretch_at=1,
        ))

    def set_scan(self, scan: Scan) -> None:
        self._scan, self._shown = scan, None
        self._tree.clear()
        for template in scan.templates:
            item = QTreeWidgetItem([f"{template.label} · {_count(template.width, 'column')} · {_count(len(template.sheets), 'sheet')}", ""])
            item.setData(0, Qt.UserRole, template)
            for sheet in template.sheets:
                start, end = sheet.header.start + 1, sheet.header.body
                rows = f"header row {start}" if start == end else f"header rows {start}–{end}"
                child = QTreeWidgetItem([f"{sheet.path.name} — {sheet.name}", rows])
                child.setData(0, Qt.UserRole, template)
                item.addChild(child)
            self._tree.addTopLevelItem(item)
        if scan.ignored:
            group = QTreeWidgetItem([f"Ignored · {_count(len(scan.ignored), 'sheet')}", ""])
            group.addChildren([
                QTreeWidgetItem([" — ".join(filter(None, (item.path.name, item.sheet))), item.reason]) for item in scan.ignored
            ])
            self._tree.addTopLevelItem(group)
        self._refresh_status()
        self._tree.setCurrentItem(self._tree.topLevelItem(0))

    def _refresh_status(self) -> None:
        templates = self._scan.templates
        warning = self.style().standardIcon(QStyle.SP_MessageBoxWarning)
        done = self.style().standardIcon(QStyle.SP_DialogApplyButton)
        linked = 0
        for index, template in enumerate(templates):
            item = self._tree.topLevelItem(index)
            has_links = bool(self._profile.links_for(template, templates))
            linked += has_links
            item.setText(1, "Not linked" if not has_links else "Auto-linked" if self._profile.is_auto(template) else "Linked")
            item.setIcon(1, done if has_links else warning)
        self._summary.setText(f"{linked} of {len(templates)} templates linked")

    def _select(self, item: QTreeWidgetItem | None) -> None:
        template = item.data(0, Qt.UserRole) if item else None
        if template is not self._shown:
            self._shown = template
            self._render()

    def _render(self) -> None:
        template = self._shown
        self._table.setRowCount(0)
        if template is None:
            self._caption.setText("Ignored sheets have no recognizable header and are not merged.")
            return
        self._caption.setText(
            f"<b>{template.label}</b> — {_count(len(template.sheets), 'sheet')}. "
            "Choose the result column for every template column you need; leave the rest as “—”."
        )
        links = self._profile.links_for(template, self._scan.templates)
        columns = template.columns(self._profile.skip_prefixes)
        self._table.setRowCount(len(columns))
        for row, column in enumerate(columns):
            cells = (column.letter, column.name or "(no header)", " · ".join(column.samples))
            for index, value in enumerate(cells):
                cell = QTableWidgetItem(value)
                cell.setToolTip("\n".join(column.samples) if index == 2 else value)
                self._table.setItem(row, index, cell)
            choice = ComboBox()
            choice.addItem("—", None)
            for result in self._profile.columns:
                choice.addItem(result.name, result.id)
            choice.setCurrentIndex(max(choice.findData(links.get(column.key)), 0))
            choice.currentIndexChanged.connect(lambda _, box=choice, key=column.key: self._link(key, box.currentData()))
            self._table.setCellWidget(row, 3, choice)

    def _link(self, key: str, target: str | None) -> None:
        self._profile.link(self._shown, self._scan.templates, key, target)
        self._refresh_status()


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(TITLE)
        self.resize(1200, 720)
        self._profile = Profile.load()
        self._scan: Scan | None = None
        self._setup = SetupPage(self._profile)
        self._links = LinkPage(self._profile)
        self._stack = QStackedWidget()
        self._stack.addWidget(self._setup)
        self._stack.addWidget(self._links)
        self.setCentralWidget(self._stack)
        self._setup.scan_requested.connect(self._start_scan)
        self._links.back_requested.connect(lambda: self._stack.setCurrentWidget(self._setup))
        self._links.generate_requested.connect(self._generate)

    def closeEvent(self, event):
        with suppress(OSError):
            self._profile.save()
        super().closeEvent(event)

    def _run(self, label: str, job, done) -> None:
        dialog = QProgressDialog(self)
        dialog.setWindowTitle(TITLE)
        dialog.setLabelText(label)
        dialog.setCancelButton(None)
        dialog.setAutoReset(False)
        dialog.setAutoClose(False)
        dialog.setRange(0, 0)
        dialog.setMinimumDuration(0)
        dialog.setWindowModality(Qt.WindowModal)
        task = Task(job, self)

        def progressed(value: int, total: int, name: str) -> None:
            dialog.setRange(0, total)
            dialog.setValue(value)
            dialog.setLabelText(f"{label}\n{name}")

        def finish(callback):
            def handle(result) -> None:
                dialog.close()
                dialog.deleteLater()
                callback(result)
            return handle

        task.progressed.connect(progressed)
        task.succeeded.connect(finish(done))
        task.failed.connect(finish(lambda message: QMessageBox.critical(self, TITLE, message)))
        task.finished.connect(task.deleteLater)
        dialog.show()
        task.start()

    def _start_scan(self) -> None:
        folder = Path(self._profile.folder)
        self._run("Scanning headers…", lambda progress: scan(folder, progress), self._scanned)

    def _scanned(self, result: Scan) -> None:
        if not result.templates:
            QMessageBox.warning(self, TITLE, "No tables with a recognizable header were found in this folder.")
            return
        self._scan = result
        self._links.set_scan(result)
        self._stack.setCurrentWidget(self._links)

    def _generate(self) -> None:
        templates = self._scan.templates
        unlinked = [template.label for template in templates if not self._profile.links_for(template, templates)]
        if len(unlinked) == len(templates):
            QMessageBox.warning(self, TITLE, "Link at least one template column to a result column.")
            return
        if unlinked and QMessageBox.question(
            self, TITLE, f"These templates have no links and will be skipped: {', '.join(unlinked)}.\n\nContinue?"
        ) != QMessageBox.Yes:
            return
        default = Path(self._profile.folder).parent / "Merged.xlsx"
        chosen, _ = QFileDialog.getSaveFileName(self, "Save result", str(default), "Excel workbook (*.xlsx)")
        if not chosen:
            return
        path = Path(chosen)
        if path.suffix.lower() != ".xlsx":
            path = path.with_name(f"{path.name}.xlsx")
        if path.resolve() in {file.resolve() for file in self._scan.files}:
            QMessageBox.warning(self, TITLE, "The result cannot overwrite one of the source files.")
            return
        with suppress(OSError):
            self._profile.save()
        scan_result, profile = self._scan, self._profile
        self._run("Merging files…", lambda progress: export(scan_result, profile, path, progress), self._generated)

    def _generated(self, outputs: list[Output]) -> None:
        folder = outputs[0].path.parent
        summary = "\n".join(f"{output.label}: {output.rows} → {output.path.name}" for output in outputs)
        box = QMessageBox(QMessageBox.Information, TITLE, f"{summary}\n\nSaved to {folder}", parent=self)
        open_folder = box.addButton("Open folder", QMessageBox.AcceptRole)
        box.addButton(QMessageBox.Close)
        box.exec()
        if box.clickedButton() is open_folder:
            QDesktopServices.openUrl(QUrl.fromLocalFile(str(folder)))


def run() -> int:
    app = QApplication(sys.argv)
    app.setApplicationName(TITLE)
    window = MainWindow()
    window.show()
    return app.exec()
