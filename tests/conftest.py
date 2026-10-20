from datetime import datetime

import pytest
from docx import Document
from openpyxl import Workbook

from exam_tickets.models import Student, Ticket
from exam_tickets.service import TicketService
from exam_tickets.storage import ResultsJournal


@pytest.fixture
def student():
    return Student("ИС-41", "Иванов", "Иван")


@pytest.fixture
def catalog():
    return {n: Ticket(n, (f"Вопрос {n}.1", f"Вопрос {n}.2", f"Вопрос {n}.3")) for n in (1, 2, 7)}


@pytest.fixture
def service(tmp_path, student, catalog):
    journal = ResultsJournal(tmp_path / "results.xlsx")
    journal.ensure_exists()
    return TicketService({student.group: (student,), "Пустая": ()}, catalog, journal,
                         clock=lambda: datetime(2026, 10, 20, 20, 30, 15))


@pytest.fixture
def sources(tmp_path):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "ИС-41"
    sheet.append(["Фамилия", "Имя"])
    sheet.append(["Иванов", "Иван"])
    sheet.append(["Петрова", "Анна"])
    other = workbook.create_sheet("ИС-42")
    other.append(["Фамилия", "Имя"])
    other.append(["Иванов", "Иван"])
    workbook.create_sheet("Пустая").append(["Фамилия", "Имя"])
    workbook.save(tmp_path / "students.xlsx")
    workbook.close()
    document = Document()
    for n in (1, 2, 7):
        document.add_paragraph(f"Билет {n}")
        for i in range(1, 4):
            document.add_paragraph(f"{i}. Вопрос {n}.{i}")
    document.save(tmp_path / "tickets.docx")
    return tmp_path
