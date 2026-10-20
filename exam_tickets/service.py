"""Assignment rules and retry state, independent of Tkinter."""

from __future__ import annotations

import random
from datetime import datetime
from typing import Callable

from .models import Assignment, AppError, JournalError, Student, Ticket
from .storage import ResultsJournal


class TicketService:
    def __init__(
        self,
        groups: dict[str, tuple[Student, ...]],
        tickets: dict[int, Ticket],
        journal: ResultsJournal,
        rng: random.Random | None = None,
        clock: Callable[[], datetime] = datetime.now,
    ):
        self.groups = groups
        self.tickets = tickets
        self.journal = journal
        self.rng = rng or random.Random()
        self.clock = clock
        self._pending: Assignment | None = None

    def students_for_group(self, group: str) -> tuple[Student, ...]:
        return self.groups.get(group, ())

    def generate(self, student: Student) -> Assignment:
        """Return only after the row is saved; failed saves retain the same draw."""
        if student not in self.students_for_group(student.group):
            raise AppError("Выберите студента из выбранной группы.")
        if not self.tickets:
            raise AppError("Нет доступных билетов.")
        with self.journal.transaction() as (workbook, sheet):
            first_number = self.journal.first_ticket(sheet, student)
            if first_number is not None:
                if first_number not in self.tickets:
                    raise JournalError(
                        f"Студенту ранее назначен билет № {first_number}, которого нет среди корректных "
                        "билетов в tickets.docx. Восстановите этот билет и обновите данные."
                    )
                ticket = self.tickets[first_number]
            elif self._pending is not None and self._pending.student == student:
                ticket = self._pending.ticket
            else:
                # Original numbers survive skipped blocks: never draw a missing ID.
                ticket = self.rng.choice(tuple(self.tickets.values()))
            if self._pending is not None and self._pending.student == student:
                timestamp = self._pending.timestamp
            else:
                timestamp = self.clock().replace(microsecond=0)
            assignment = Assignment(student, ticket, timestamp, first_number is not None)
            self._pending = assignment
            self.journal.append(workbook, sheet, assignment)
        self._pending = None
        return assignment
