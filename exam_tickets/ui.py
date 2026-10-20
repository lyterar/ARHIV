"""Russian desktop UI. Only the result window handles Escape."""

from __future__ import annotations

import logging
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from .models import AppError, Assignment, JournalError, Student
from .service import TicketService
from .sources import load_students, load_tickets
from .storage import ResultsJournal

LOG = logging.getLogger(__name__)
BACKGROUND = "#F2F4F8"
INK = "#172B46"
MUTED = "#65748A"
ACCENT = "#2559A7"


class TicketApp:
    def __init__(self, root: tk.Tk, data_dir: Path, *, auto_load: bool = True):
        self.root = root
        self.data_dir = data_dir.resolve()
        self.service: TicketService | None = None
        self.result_window: tk.Toplevel | None = None
        self.busy = False
        self.group = tk.StringVar(root)
        self.student = tk.StringVar(root)
        self.status = tk.StringVar(root, "Загрузка источников…")
        self.summary = tk.StringVar(root, "Списки групп и билетов")
        self.folder = tk.StringVar(root, str(self.data_dir))
        self._students: tuple[Student, ...] = ()
        self._build()
        if auto_load:
            root.after(50, self.reload_data)

    def _build(self) -> None:
        self.root.title("Лабораторная работа 2 — Генератор билетов")
        self.root.geometry("860x660")
        self.root.minsize(720, 600)
        self.root.configure(bg=BACKGROUND)
        style = ttk.Style(self.root)
        style.theme_use("clam")
        style.configure(".", font=("Segoe UI", 11))
        style.configure("TFrame", background=BACKGROUND)
        style.configure("Card.TFrame", background="white")
        style.configure("TLabel", background=BACKGROUND, foreground=INK)
        style.configure("Muted.TLabel", foreground=MUTED)
        style.configure("Card.TLabel", background="white", foreground=INK)
        style.configure("CardMuted.TLabel", background="white", foreground=MUTED, font=("Segoe UI", 10))
        style.configure("Title.TLabel", font=("Segoe UI", 25, "bold"))
        style.configure("TCombobox", padding=9, fieldbackground="white", arrowsize=16)
        style.map("TCombobox", fieldbackground=[("readonly", "white"), ("disabled", "#F2F4F8")],
                  foreground=[("disabled", "#7E899A")])
        style.configure("TButton", padding=(15, 10), background="#E6EBF2", foreground=INK)
        style.configure("Primary.TButton", padding=(20, 14), background=ACCENT,
                        foreground="white", font=("Segoe UI", 12, "bold"), borderwidth=0)
        style.map("Primary.TButton", background=[("disabled", "#DCE3ED"), ("active", "#1D4684")],
                  foreground=[("disabled", "#8491A5")])

        outer = ttk.Frame(self.root, padding=(36, 27))
        outer.pack(fill="both", expand=True)
        ttk.Label(outer, text="ЛАБОРАТОРНАЯ РАБОТА 02", style="Muted.TLabel",
                  font=("Segoe UI", 10, "bold")).pack(anchor="w")
        ttk.Label(outer, text="Экзаменационные билеты", style="Title.TLabel").pack(anchor="w", pady=(8, 6))
        ttk.Label(outer, text="Выберите группу и студента, чтобы получить три вопроса.",
                  style="Muted.TLabel").pack(anchor="w")

        meta = ttk.Frame(outer)
        meta.pack(fill="x", pady=(23, 12))
        ttk.Label(meta, textvariable=self.summary, style="Muted.TLabel").pack(side="left")

        card = ttk.Frame(outer, style="Card.TFrame", padding=26)
        card.pack(fill="both", expand=True)
        ttk.Label(card, text="01   Группа", style="Card.TLabel",
                  font=("Segoe UI", 11, "bold")).pack(anchor="w")
        self.group_combo = ttk.Combobox(card, textvariable=self.group, state="disabled")
        self.group_combo.pack(fill="x", pady=(8, 20))
        self.group_combo.bind("<<ComboboxSelected>>", self.on_group_changed)
        ttk.Label(card, text="02   Студент", style="Card.TLabel",
                  font=("Segoe UI", 11, "bold")).pack(anchor="w")
        self.student_combo = ttk.Combobox(card, textvariable=self.student, state="disabled")
        self.student_combo.pack(fill="x", pady=(8, 10))
        self.student_combo.bind("<<ComboboxSelected>>", self.on_student_changed)
        self.hint = ttk.Label(card, text="Сначала выберите группу.", style="CardMuted.TLabel")
        self.hint.pack(anchor="w")
        self.generate_button = ttk.Button(card, text="Сгенерировать билет", style="Primary.TButton",
                                          command=self.generate, state="disabled")
        self.generate_button.pack(fill="x", pady=(22, 12))
        ttk.Label(card, text="При повторном выборе сохраняется первый назначенный билет.",
                  style="CardMuted.TLabel", wraplength=650).pack(anchor="w")

        ttk.Label(outer, textvariable=self.status, style="Muted.TLabel",
                  wraplength=750).pack(anchor="w", pady=(15, 9))
        footer = ttk.Frame(outer)
        footer.pack(fill="x")
        self.reload_button = ttk.Button(footer, text="Обновить данные", command=self.reload_data)
        self.reload_button.pack(side="left")
        self.folder_button = ttk.Button(footer, text="Выбрать папку", command=self.choose_folder)
        self.folder_button.pack(side="left", padx=8)
        ttk.Label(footer, text="ESC — закрыть результат", style="Muted.TLabel").pack(side="right")

    def reload_data(self) -> None:
        if self.busy or self.result_window is not None:
            return
        self.service = None
        self._students = ()
        self.group.set("")
        self.student.set("")
        self.group_combo.configure(values=(), state="disabled")
        self.student_combo.configure(values=(), state="disabled")
        self.generate_button.state(["disabled"])
        self.summary.set("Источники не загружены")
        try:
            journal = ResultsJournal(self.data_dir / "results.xlsx")
            journal.ensure_exists()
            groups = load_students(self.data_dir / "students.xlsx")
            tickets = load_tickets(self.data_dir / "tickets.docx")
            self.service = TicketService(groups, tickets, journal)
        except AppError as exc:
            LOG.warning("Ошибка загрузки: %s", exc)
            self.status.set("Исправьте файлы и нажмите «Обновить данные».")
            self.hint.configure(text="Для выбора студента загрузите источники.")
            messagebox.showerror("Не удалось загрузить данные", str(exc), parent=self.root)
            return
        self.group_combo.configure(values=tuple(groups), state="readonly")
        self.summary.set(f"Групп: {len(groups)}     Студентов: {sum(map(len, groups.values()))}     Билетов: {len(tickets)}")
        self.status.set("Данные загружены. Результат сохраняется сразу после выдачи.")
        self.hint.configure(text="Сначала выберите группу.")
        self.group_combo.focus_set()

    def choose_folder(self) -> None:
        selected = filedialog.askdirectory(parent=self.root, initialdir=self.data_dir,
                                           title="Папка со students.xlsx и tickets.docx")
        if selected:
            self.data_dir = Path(selected).resolve()
            self.folder.set(str(self.data_dir))
            self.reload_data()

    def on_group_changed(self, _event: tk.Event | None = None) -> None:
        self.student.set("")
        self._students = self.service.students_for_group(self.group.get()) if self.service else ()
        self.student_combo.configure(values=tuple(student.full_name for student in self._students),
                                     state="readonly" if self._students else "disabled")
        self.hint.configure(text=f"Студентов в группе: {len(self._students)}. Выберите студента."
                            if self._students else "В этой группе пока нет студентов.")
        self._sync_button()

    def on_student_changed(self, _event: tk.Event | None = None) -> None:
        self._sync_button()

    def _sync_button(self) -> None:
        ready = (self.service is not None and self.group.get() in self.service.groups
                 and 0 <= self.student_combo.current() < len(self._students)
                 and not self.busy and self.result_window is None)
        self.generate_button.state(["!disabled"] if ready else ["disabled"])

    def generate(self) -> None:
        index = self.student_combo.current()
        if self.service is None or self.busy or self.result_window is not None or not 0 <= index < len(self._students):
            return
        student = self._students[index]
        self.busy = True
        self._sync_button()
        self.status.set("Сохраняем результат…")
        self.root.update_idletasks()
        try:
            while True:
                try:
                    result = self.service.generate(student)
                    break
                except JournalError as exc:
                    LOG.warning("Запись результата: %s", exc)
                    if not messagebox.askretrycancel("Результат не сохранён", str(exc), parent=self.root):
                        self.status.set("Выдача отменена: запись не сохранена. Можно повторить попытку.")
                        return
                except AppError as exc:
                    messagebox.showerror("Не удалось выдать билет", str(exc), parent=self.root)
                    self.status.set("Выдача не выполнена.")
                    return
            self.status.set(f"Сохранено: {student.full_name}, билет № {result.ticket.number}.")
            self.show_result(result)
        finally:
            self.busy = False
            self._sync_button()

    def show_result(self, assignment: Assignment) -> None:
        window = tk.Toplevel(self.root)
        self.result_window = window
        window.title(f"Ваш билет № {assignment.ticket.number}")
        window.geometry("760x590")
        window.minsize(600, 480)
        window.configure(bg=BACKGROUND)
        window.transient(self.root)
        window.protocol("WM_DELETE_WINDOW", self.close_result)
        window.bind("<Escape>", self.close_result)
        content = ttk.Frame(window, padding=30)
        content.pack(fill="both", expand=True)
        ttk.Label(content, text="ПОВТОРНАЯ ВЫДАЧА" if assignment.repeated else "БИЛЕТ НАЗНАЧЕН",
                  style="Muted.TLabel", font=("Segoe UI", 10, "bold")).pack(anchor="w")
        ttk.Label(content, text=f"Ваш билет № {assignment.ticket.number}",
                  style="Title.TLabel").pack(anchor="w", pady=(8, 10))
        ttk.Label(content, text=f"{assignment.student.full_name}  ·  {assignment.student.group}",
                  wraplength=680).pack(anchor="w", pady=(0, 20))
        questions_frame = ttk.Frame(content)
        questions_frame.pack(fill="both", expand=True)
        text = tk.Text(questions_frame, wrap="word", font=("Segoe UI", 13), bg="white",
                       fg=INK, relief="flat", padx=20, pady=18, cursor="arrow", height=9)
        scrollbar = ttk.Scrollbar(questions_frame, orient="vertical", command=text.yview)
        text.configure(yscrollcommand=scrollbar.set)
        text.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        text.tag_configure("number", foreground=ACCENT, font=("Segoe UI", 13, "bold"))
        text.tag_configure("question", spacing3=20)
        for index, question in enumerate(assignment.ticket.questions, 1):
            text.insert("end", f"{index:02d}   ", "number")
            text.insert("end", question + "\n", "question")
        text.configure(state="disabled")
        self.result_text = text
        ttk.Label(content, text=f"Записано в results.xlsx  ·  {assignment.timestamp:%d.%m.%Y %H:%M:%S}",
                  style="Muted.TLabel").pack(anchor="w", pady=(16, 12))
        close = ttk.Button(content, text="Следующий студент  ·  ESC", style="Primary.TButton",
                           command=self.close_result)
        close.pack(fill="x")
        window.wait_visibility()
        window.grab_set()
        close.focus_set()

    def close_result(self, _event: tk.Event | None = None) -> str:
        if self.result_window is not None:
            self.result_window.grab_release()
            self.result_window.destroy()
            self.result_window = None
            self.student.set("")
            self._sync_button()
            self.student_combo.focus_set()
        return "break"
