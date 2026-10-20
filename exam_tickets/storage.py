"""Append-only Excel history with atomic saves and a lock between app instances."""

from __future__ import annotations

import contextlib
import os
import tempfile
from contextlib import contextmanager
from io import BytesIO
from pathlib import Path
from typing import Iterator

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from .models import Assignment, JournalError, Student

HEADERS = ("Группа", "Фамилия", "Имя", "Номер билета", "Дата и время", "Повтор (да/нет)")
DATE_FORMAT = "yyyy-mm-dd hh:mm:ss"


@contextmanager
def _exclusive_lock(path: Path) -> Iterator[None]:
    """OS locks are released even if a process crashes; the sidecar may remain."""
    with open(path.with_name(f".{path.name}.lock"), "a+b") as handle:
        handle.seek(0, os.SEEK_END)
        if handle.tell() == 0:
            handle.write(b"\0")
            handle.flush()
        handle.seek(0)
        if os.name == "nt":
            import msvcrt
            msvcrt.locking(handle.fileno(), msvcrt.LK_NBLCK, 1)
        else:
            import fcntl
            fcntl.flock(handle.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
        try:
            yield
        finally:
            handle.seek(0)
            if os.name == "nt":
                msvcrt.locking(handle.fileno(), msvcrt.LK_UNLCK, 1)
            else:
                fcntl.flock(handle.fileno(), fcntl.LOCK_UN)


def _last_row(sheet: Worksheet) -> int:
    for index in range(sheet.max_row, 0, -1):
        if any(cell.value is not None for cell in sheet[index]):
            return index
    return 0


class ResultsJournal:
    def __init__(self, path: Path):
        self.path = path

    def ensure_exists(self) -> None:
        """Create the header at startup, without rewriting an existing workbook."""
        with self.transaction() as (workbook, _sheet):
            if not self.path.exists():
                self.save(workbook)

    @contextmanager
    def transaction(self) -> Iterator[tuple[Workbook, Worksheet]]:
        """Serialize read/first-assignment/append/write as one transaction."""
        try:
            with _exclusive_lock(self.path):
                workbook = self._open()
                try:
                    yield workbook, workbook.worksheets[0]
                finally:
                    workbook.close()
        except JournalError:
            raise
        except (PermissionError, BlockingIOError) as exc:
            raise JournalError(
                f"Файл {self.path.name} занят или недоступен для записи.\n"
                "Закройте его в Excel и другие экземпляры приложения, затем повторите попытку."
            ) from exc
        except OSError as exc:
            raise JournalError(f"Не удалось сохранить {self.path.name}.\nПроверьте папку, права и место на диске.\n{exc}") from exc

    def _open(self) -> Workbook:
        if self.path.exists():
            try:
                workbook = load_workbook(BytesIO(self.path.read_bytes()))
            except OSError:
                raise
            except Exception as exc:
                raise JournalError(f"Файл {self.path.name} повреждён. Восстановите его или выберите другую папку данных.") from exc
            sheet = workbook.worksheets[0]
            if tuple(sheet.cell(1, col).value for col in range(1, 7)) != HEADERS:
                workbook.close()
                raise JournalError(f"В {self.path.name} неверные заголовки. Существующий файл не изменён.")
            return workbook
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Результаты"
        sheet.append(HEADERS)
        sheet.freeze_panes = "A2"
        sheet.sheet_view.showGridLines = False
        for column, width in enumerate((20, 25, 25, 18, 23, 21), 1):
            cell = sheet.cell(1, column)
            cell.font = Font(name="Calibri", bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="254A78")
            sheet.column_dimensions[get_column_letter(column)].width = width
        sheet.row_dimensions[1].height = 26
        return workbook

    @staticmethod
    def first_ticket(sheet: Worksheet, student: Student) -> int | None:
        """Use physical row order, never the latest row or a timestamp sort."""
        for row in sheet.iter_rows(min_row=2, max_col=6, values_only=True):
            if tuple(row[:3]) == student.key:
                number = row[3]
                if isinstance(number, bool) or not isinstance(number, (int, float)) or number < 1 or int(number) != number:
                    raise JournalError(f"В первой записи студента {student.full_name} некорректный номер билета.")
                return int(number)
        return None

    def append(self, workbook: Workbook, sheet: Worksheet, assignment: Assignment) -> None:
        index = _last_row(sheet) + 1
        for column, value in enumerate(assignment.as_row(), 1):
            cell = sheet.cell(index, column, value)
            if isinstance(value, str):
                cell.data_type = "s"
        sheet.cell(index, 5).number_format = DATE_FORMAT
        self.save(workbook)

    def save(self, workbook: Workbook) -> None:
        """Flush a temporary file on the same disk, then replace the destination."""
        temp_path: Path | None = None
        try:
            with tempfile.NamedTemporaryFile(
                prefix=f".{self.path.stem}-", suffix=".tmp", dir=self.path.parent, delete=False
            ) as handle:
                temp_path = Path(handle.name)
                workbook.save(handle)
                handle.flush()
                os.fsync(handle.fileno())
            os.replace(temp_path, self.path)
        finally:
            if temp_path is not None:
                with contextlib.suppress(OSError):
                    temp_path.unlink(missing_ok=True)
