"""Read groups from XLSX and explicitly numbered tickets from DOCX."""

from __future__ import annotations

import logging
import re
from io import BytesIO
from pathlib import Path

from docx import Document
from openpyxl import load_workbook

from .models import SourceError, Student, Ticket

LOG = logging.getLogger(__name__)
HEADER = re.compile(r"Билет\s+([1-9]\d*)", re.IGNORECASE)
QUESTION = re.compile(r"([1-9]\d*)\.\s*(\S.*)")


def _source_bytes(path: Path) -> bytes:
    try:
        return path.read_bytes()
    except FileNotFoundError as exc:
        raise SourceError(f"Не найден файл {path.name}.\nПоместите его в папку:\n{path.parent}") from exc
    except PermissionError as exc:
        raise SourceError(f"Нет доступа к {path.name}. Закройте файл в другой программе и повторите загрузку.") from exc
    except OSError as exc:
        raise SourceError(f"Не удалось прочитать {path.name}: {exc}") from exc


def load_students(path: Path) -> dict[str, tuple[Student, ...]]:
    """Keep worksheet/row order; retain empty groups and skip invalid student rows."""
    data = _source_bytes(path)
    try:
        workbook = load_workbook(BytesIO(data), read_only=True, data_only=False)
    except Exception as exc:
        raise SourceError(f"Файл {path.name} повреждён или не является книгой Excel.") from exc
    groups: dict[str, tuple[Student, ...]] = {}
    try:
        for sheet in workbook.worksheets:
            rows = sheet.iter_rows(max_col=2)
            header = next(rows, ())
            titles = tuple(str(cell.value or "").strip().casefold() for cell in header)
            if titles != ("фамилия", "имя"):
                raise SourceError(
                    f"{path.name}, лист «{sheet.title}»: в A1 должно быть «Фамилия», в B1 — «Имя»."
                )
            students: list[Student] = []
            seen: set[tuple[str, str, str]] = set()
            for index, cells in enumerate(rows, start=2):
                values = tuple(cell.value for cell in cells)
                if all(value is None or value == "" for value in values):
                    continue
                if any(cell.data_type == "f" for cell in cells) or any(
                    not isinstance(value, str) or not value.strip() for value in values
                ):
                    LOG.warning("%s / %s, строка %s: некорректное имя, строка пропущена.",
                                path.name, sheet.title, index)
                    continue
                student = Student(sheet.title, values[0].strip(), values[1].strip())
                if student.key in seen:
                    LOG.warning("%s / %s, строка %s: дубликат студента пропущен.",
                                path.name, sheet.title, index)
                    continue
                seen.add(student.key)
                students.append(student)
            groups[sheet.title] = tuple(students)
    except SourceError:
        raise
    except Exception as exc:
        raise SourceError(f"Не удалось прочитать листы {path.name}. Проверьте структуру книги.") from exc
    finally:
        workbook.close()
    if not groups:
        raise SourceError(f"В {path.name} нет листов с группами.")
    return groups


def load_tickets(path: Path) -> dict[int, Ticket]:
    """Parse 'Билет N' and 'N. text'; never merge questions across broken blocks."""
    data = _source_bytes(path)
    try:
        document = Document(BytesIO(data))
    except Exception as exc:
        raise SourceError(f"Файл {path.name} повреждён или не является документом Word.") from exc
    tickets: dict[int, Ticket] = {}
    number: int | None = None
    questions: list[tuple[int, str]] = []
    def finish() -> None:
        if number is None:
            return
        if len(questions) < 3 or [q[0] for q in questions] != list(range(1, len(questions) + 1)):
            LOG.warning("Билет %s пропущен: нужны как минимум три последовательных вопроса 1., 2., 3.",
                        number)
        elif number in tickets:
            LOG.warning("Билет %s пропущен: такой номер уже встречался.", number)
        else:
            tickets[number] = Ticket(number, tuple(q[1] for q in questions[:3]))
            if len(questions) > 3:
                LOG.warning("Билет %s: используются первые три вопроса из %s.", number, len(questions))

    for paragraph in document.paragraphs:
        # A manual line break in a paragraph is also accepted.
        for raw in paragraph.text.splitlines():
            text = raw.strip()
            if not text:
                continue
            heading = HEADER.fullmatch(text)
            question = QUESTION.fullmatch(text)
            if heading:
                finish()
                number, questions = int(heading[1]), []
            elif not question:
                # An unknown heading also ends the preceding block. Its questions
                # must not corrupt an otherwise valid preceding ticket.
                finish()
                number, questions = None, []
                LOG.warning("Нераспознанный заголовок «%s»: блок пропущен.", text)
            elif number is not None:
                questions.append((int(question[1]), question[2].strip()))
            else:
                LOG.warning("Абзац без заголовка билета пропущен: «%s».", text)
    finish()
    if not tickets:
        raise SourceError(f"В {path.name} нет корректных билетов с тремя вопросами. Подробности в app.log.")
    return tickets
