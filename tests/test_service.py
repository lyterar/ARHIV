from datetime import datetime
import random

import pytest
from openpyxl import load_workbook
from openpyxl.styles import Font

from exam_tickets.models import AppError, JournalError, Student
from exam_tickets.service import TicketService
from exam_tickets.storage import DATE_FORMAT, HEADERS, ResultsJournal, _exclusive_lock
import exam_tickets.storage as storage


def rows(path):
    workbook = load_workbook(path)
    try:
        return list(workbook.active.values)
    finally:
        workbook.close()


def test_journal_created_with_exact_header_before_any_generation(service):
    assert rows(service.journal.path) == [HEADERS]


def test_group_lookup_is_exact_and_unknown_group_is_empty(service, student):
    assert service.students_for_group(student.group) == (student,)
    assert service.students_for_group("ИС-42") == ()
    assert service.students_for_group("Пустая") == ()


def test_immediate_save_and_repeated_click_append_history(service, student):
    first = service.generate(student)
    assert rows(service.journal.path) == [HEADERS, first.as_row()]
    for _ in range(3):
        repeated = service.generate(student)
        assert repeated.repeated
        assert repeated.ticket == first.ticket
    saved = rows(service.journal.path)
    assert len(saved) == 5
    assert [row[5] for row in saved[1:]] == ["нет", "да", "да", "да"]
    assert saved[1] == first.as_row()


def test_same_student_after_application_restart_reuses_first_ticket(service, student, catalog):
    first = service.generate(student)
    restarted = TicketService(service.groups, catalog, ResultsJournal(service.journal.path), random.Random(9))
    assert restarted.generate(student).ticket == first.ticket
    assert len(rows(service.journal.path)) == 3


def test_uses_first_row_even_when_later_row_has_different_ticket(service, student):
    first = service.generate(student)
    workbook = load_workbook(service.journal.path)
    different = next(n for n in service.tickets if n != first.ticket.number)
    workbook.active.append([*student.key, different, datetime(2026, 10, 20), "да"])
    workbook.save(service.journal.path)
    workbook.close()
    before = rows(service.journal.path)
    repeated = service.generate(student)
    assert repeated.ticket == first.ticket
    assert rows(service.journal.path)[:-1] == before


def test_same_name_in_two_groups_are_different_students(service, student):
    other = Student("ИС-42", student.last_name, student.first_name)
    service.groups[other.group] = (other,)
    assert not service.generate(student).repeated
    assert not service.generate(other).repeated
    assert service.generate(other).repeated
    assert service.generate(student).repeated


def test_draw_only_uses_surviving_ticket_numbers(service):
    students = tuple(Student("Новая", f"Студент{i}", "Имя") for i in range(25))
    service.groups["Новая"] = students
    service.rng = random.Random(42)
    assert {service.generate(s).ticket.number for s in students} == {1, 2, 7}


def test_unknown_student_is_rejected_without_writing(service):
    with pytest.raises(AppError, match="Выберите студента"):
        service.generate(Student("ИС-41", "Неизвестный", "Иван"))
    assert rows(service.journal.path) == [HEADERS]


def test_first_assigned_ticket_removed_from_source_is_not_reassigned(service, student):
    first = service.generate(student)
    del service.tickets[first.ticket.number]
    before = service.journal.path.read_bytes()
    with pytest.raises(JournalError, match="Восстановите"):
        service.generate(student)
    assert service.journal.path.read_bytes() == before


def test_excel_date_is_typed_and_precise_to_seconds(service, student):
    service.clock = lambda: datetime(2026, 10, 20, 21, 2, 3, 987654)
    service.generate(student)
    workbook = load_workbook(service.journal.path)
    cell = workbook.active["E2"]
    assert cell.value == datetime(2026, 10, 20, 21, 2, 3)
    assert cell.number_format == DATE_FORMAT
    workbook.close()


def test_save_failure_retry_keeps_ticket_and_writes_exactly_once(service, student, monkeypatch):
    real_replace = storage.os.replace
    attempts = []
    before = service.journal.path.read_bytes()

    def fail_once(source, target):
        attempts.append(service._pending)
        if len(attempts) == 1:
            raise PermissionError("locked by Excel")
        real_replace(source, target)

    monkeypatch.setattr(storage.os, "replace", fail_once)
    with pytest.raises(JournalError, match="занят"):
        service.generate(student)
    assert service.journal.path.read_bytes() == before
    result = service.generate(student)
    assert attempts[0] == attempts[1] == result
    assert rows(service.journal.path) == [HEADERS, result.as_row()]
    assert list(service.journal.path.parent.glob("*.tmp")) == []


def test_actual_windows_file_lock_recovers(service, student):
    if storage.os.name != "nt":
        pytest.skip("Windows replacement sharing semantics")
    with service.journal.path.open("rb"):
        with pytest.raises(JournalError, match="занят"):
            service.generate(student)
    result = service.generate(student)
    assert len(rows(service.journal.path)) == 2
    assert not result.repeated


def test_parallel_instance_lock_preserves_history(service, student):
    with _exclusive_lock(service.journal.path):
        with pytest.raises(JournalError):
            service.generate(student)
    assert rows(service.journal.path) == [HEADERS]
    service.generate(student)
    assert len(rows(service.journal.path)) == 2


def test_corrupt_file_is_untouched_and_recoverable(service, student):
    service.journal.path.write_bytes(b"broken")
    with pytest.raises(JournalError, match="поврежд"):
        service.generate(student)
    assert service.journal.path.read_bytes() == b"broken"


def test_invalid_header_is_not_overwritten(service):
    workbook = load_workbook(service.journal.path)
    workbook.active["A1"] = "wrong"
    workbook.save(service.journal.path)
    workbook.close()
    before = service.journal.path.read_bytes()
    with pytest.raises(JournalError, match="заголовки"):
        service.journal.ensure_exists()
    assert service.journal.path.read_bytes() == before


def test_existing_formatting_and_extra_sheets_survive(service, student):
    workbook = load_workbook(service.journal.path)
    workbook.active["A50"].font = Font(italic=True)
    workbook.create_sheet("Примечания")["A1"] = "Сохранить"
    workbook.save(service.journal.path)
    workbook.close()
    service.generate(student)
    workbook = load_workbook(service.journal.path)
    assert workbook.active["A2"].value == student.group
    assert workbook.active["A50"].font.italic
    assert workbook["Примечания"]["A1"].value == "Сохранить"
    workbook.close()


def test_name_resembling_formula_is_stored_as_text(service):
    student = Student("ИС-41", "=1+1", "+Имя")
    service.groups[student.group] = (student,)
    service.generate(student)
    workbook = load_workbook(service.journal.path)
    assert workbook.active["B2"].value == "=1+1"
    assert workbook.active["B2"].data_type == "s"
    workbook.close()


def test_initialization_never_rewrites_existing_journal(service):
    before = service.journal.path.read_bytes()
    service.journal.ensure_exists()
    assert service.journal.path.read_bytes() == before
