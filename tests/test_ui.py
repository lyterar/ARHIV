import tkinter as tk
import os

import pytest
from openpyxl import load_workbook

from exam_tickets.models import JournalError
from exam_tickets.ui import TicketApp
import exam_tickets.ui as ui

pytestmark = pytest.mark.gui


@pytest.fixture(scope="module")
def tk_root():
    if os.name != "nt" and not os.environ.get("DISPLAY") and not os.environ.get("WAYLAND_DISPLAY"):
        pytest.skip("GUI tests require a graphical desktop")
    root = tk.Tk()
    yield root
    root.destroy()


@pytest.fixture
def app(sources, monkeypatch, tk_root):
    root = tk_root
    # The real widget tree and event loop are tested; dialogs are intercepted.
    errors = []
    monkeypatch.setattr(ui.messagebox, "showerror", lambda *args, **kwargs: errors.append(args))
    application = TicketApp(root, sources, auto_load=False)
    application.reload_data()
    root.update()
    application.errors = errors
    yield application
    application.close_result()
    for child in root.winfo_children():
        child.destroy()


def select_student(app, group=0, student=0):
    app.group_combo.current(group)
    app.group_combo.event_generate("<<ComboboxSelected>>")
    app.student_combo.current(student)
    app.student_combo.event_generate("<<ComboboxSelected>>")
    app.root.update()


def test_button_requires_both_choices_and_group_switch_resets_student(app):
    assert app.generate_button.instate(["disabled"])
    app.group_combo.current(0)
    app.on_group_changed()
    assert app.generate_button.instate(["disabled"])
    app.student_combo.current(0)
    app.on_student_changed()
    assert app.generate_button.instate(["!disabled"])
    app.group_combo.current(1)
    app.on_group_changed()
    assert app.student.get() == ""
    assert app.generate_button.instate(["disabled"])
    assert app.student_combo["values"] == ("Иванов Иван",)


def test_empty_group_keeps_list_empty_and_generation_disabled(app):
    select_student(app)
    app.group_combo.current(2)
    app.on_group_changed()
    assert not app.student_combo["values"]
    assert app.student_combo.instate(["disabled"])
    assert app.generate_button.instate(["disabled"])


def test_result_is_saved_before_escape_and_main_window_survives(app):
    select_student(app)
    app.generate_button.invoke()
    app.root.update()
    result = app.result_window
    assert result is not None
    assert "Вопрос" in app.result_text.get("1.0", "end")
    workbook = load_workbook(app.data_dir / "results.xlsx")
    assert workbook.active.max_row == 2
    first_ticket = workbook.active["D2"].value
    workbook.close()
    assert result.title() == f"Ваш билет № {first_ticket}"
    result.focus_force()
    result.event_generate("<Escape>")
    app.root.update()
    assert app.result_window is None and app.root.winfo_exists()
    assert app.generate_button.instate(["disabled"])
    select_student(app)
    app.generate()
    workbook = load_workbook(app.data_dir / "results.xlsx")
    assert workbook.active.max_row == 3
    assert workbook.active["D3"].value == first_ticket
    assert workbook.active["F3"].value == "да"
    workbook.close()
    app.close_result()


def test_missing_source_keeps_app_open_and_reload_recovers(app):
    path = app.data_dir / "students.xlsx"
    data = path.read_bytes()
    path.unlink()
    app.reload_data()
    assert app.errors and app.root.winfo_exists()
    assert app.generate_button.instate(["disabled"])
    path.write_bytes(data)
    app.reload_data()
    assert app.service is not None
    assert app.group_combo.instate(["readonly"])


def test_cancel_failed_save_does_not_show_result(app, monkeypatch):
    select_student(app)
    monkeypatch.setattr(app.service.journal, "append", lambda *args: (_ for _ in ()).throw(JournalError("locked")))
    monkeypatch.setattr(ui.messagebox, "askretrycancel", lambda *args, **kwargs: False)
    app.generate()
    assert app.result_window is None
    assert "не сохранена" in app.status.get()
    assert app.generate_button.instate(["!disabled"])


def test_retry_dialog_saves_once_then_shows_result(app, monkeypatch):
    select_student(app)
    real_append = app.service.journal.append
    attempts = []

    def fail_once(*args):
        attempts.append(args[-1])
        if len(attempts) == 1:
            raise JournalError("locked")
        return real_append(*args)

    monkeypatch.setattr(app.service.journal, "append", fail_once)
    monkeypatch.setattr(ui.messagebox, "askretrycancel", lambda *args, **kwargs: True)
    app.generate()
    assert attempts[0] == attempts[1]
    assert app.result_window is not None
    workbook = load_workbook(app.data_dir / "results.xlsx")
    assert workbook.active.max_row == 2
    workbook.close()
    app.close_result()
