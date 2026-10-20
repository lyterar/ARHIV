import pytest
from docx import Document
from openpyxl import load_workbook

from exam_tickets.models import SourceError
from exam_tickets.sources import load_students, load_tickets


def write_doc(path, lines):
    document = Document()
    for line in lines:
        document.add_paragraph(line)
    document.save(path)


def block(number, count=3):
    return [f"Билет {number}", *(f"{i}. Вопрос {number}.{i}" for i in range(1, count + 1))]


def test_reads_three_questions_and_preserves_original_numbers(sources):
    tickets = load_tickets(sources / "tickets.docx")
    assert list(tickets) == [1, 2, 7]
    assert tickets[7].questions == ("Вопрос 7.1", "Вопрос 7.2", "Вопрос 7.3")


def test_skips_incomplete_ticket_and_logs(tmp_path, caplog):
    path = tmp_path / "tickets.docx"
    write_doc(path, block(1, 2) + block(2))
    assert list(load_tickets(path)) == [2]
    assert "Билет 1 пропущен" in caplog.text


@pytest.mark.parametrize("heading", ["Билет два", "Билет 0", "Билет -5", "Билет № 3", "Экзаменационный билет 9", "Билет4", "Ticket 5"])
def test_malformed_heading_does_not_merge_questions(tmp_path, caplog, heading):
    path = tmp_path / "tickets.docx"
    write_doc(path, block(1, 2) + [heading, "1. Чужой", "2. Чужой", "3. Чужой"] + block(7))
    assert list(load_tickets(path)) == [7]
    assert "Нераспознанный заголовок" in caplog.text


def test_unknown_heading_does_not_discard_previous_valid_ticket(tmp_path):
    path = tmp_path / "tickets.docx"
    write_doc(path, block(1) + ["Билт 3", "1. A", "2. B", "3. C"] + block(7))
    assert list(load_tickets(path)) == [1, 7]


def test_bad_question_numbering_and_duplicate_ids_are_logged(tmp_path, caplog):
    path = tmp_path / "tickets.docx"
    write_doc(path, ["Билет 1", "1. A", "1. B", "3. C"] + block(2) + block(2))
    assert list(load_tickets(path)) == [2]
    assert "такой номер уже встречался" in caplog.text


def test_more_than_three_questions_uses_first_three(tmp_path, caplog):
    path = tmp_path / "tickets.docx"
    write_doc(path, block(1, 4))
    assert len(load_tickets(path)[1].questions) == 3
    assert "первые три" in caplog.text


def test_manual_line_breaks_and_surrounding_whitespace(tmp_path):
    path = tmp_path / "tickets.docx"
    write_doc(path, ["  Билет 4 \n 1. Один \n2. Два\n 3. Три "])
    assert load_tickets(path)[4].questions == ("Один", "Два", "Три")


def test_no_valid_tickets_is_a_friendly_error(tmp_path):
    path = tmp_path / "tickets.docx"
    write_doc(path, block(1, 2))
    with pytest.raises(SourceError, match="нет корректных билетов"):
        load_tickets(path)


def test_groups_match_worksheets_and_empty_group_survives(sources):
    groups = load_students(sources / "students.xlsx")
    assert list(groups) == ["ИС-41", "ИС-42", "Пустая"]
    assert [s.full_name for s in groups["ИС-41"]] == ["Иванов Иван", "Петрова Анна"]
    assert groups["ИС-42"][0].group == "ИС-42"
    assert groups["Пустая"] == ()


def test_bad_rows_and_duplicates_are_skipped_names_trimmed(sources, caplog):
    path = sources / "students.xlsx"
    workbook = load_workbook(path)
    sheet = workbook.active
    sheet.append([" Иванов ", " Иван "])
    sheet.append(["Сидоров", None])
    sheet.append(["=1+1", "Иван"])
    sheet.append([" Соколов ", " Андрей "])
    workbook.save(path)
    workbook.close()
    students = load_students(path)["ИС-41"]
    assert [s.full_name for s in students] == ["Иванов Иван", "Петрова Анна", "Соколов Андрей"]
    assert "дубликат" in caplog.text and "некорректное имя" in caplog.text


def test_wrong_student_headers_fail(sources):
    path = sources / "students.xlsx"
    workbook = load_workbook(path)
    workbook.active["A1"] = "Не фамилия"
    workbook.save(path)
    workbook.close()
    with pytest.raises(SourceError, match="A1"):
        load_students(path)


@pytest.mark.parametrize("reader, filename", [(load_students, "students.xlsx"), (load_tickets, "tickets.docx")])
def test_missing_and_corrupt_sources(reader, filename, tmp_path):
    path = tmp_path / filename
    with pytest.raises(SourceError, match="Не найден файл"):
        reader(path)
    path.write_bytes(b"not an office document")
    with pytest.raises(SourceError, match="поврежд"):
        reader(path)
