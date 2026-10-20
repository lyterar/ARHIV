"""Immutable application data, independent of the interface and file formats."""

from dataclasses import dataclass
from datetime import datetime


class AppError(Exception):
    """An expected problem that can be shown to the user without a traceback."""


class SourceError(AppError):
    """A source file cannot be used."""


class JournalError(AppError):
    """The results journal cannot be read or saved."""


@dataclass(frozen=True)
class Student:
    group: str
    last_name: str
    first_name: str

    @property
    def full_name(self) -> str:
        return f"{self.last_name} {self.first_name}"

    @property
    def key(self) -> tuple[str, str, str]:
        return self.group, self.last_name, self.first_name


@dataclass(frozen=True)
class Ticket:
    number: int
    questions: tuple[str, str, str]


@dataclass(frozen=True)
class Assignment:
    student: Student
    ticket: Ticket
    timestamp: datetime
    repeated: bool

    def as_row(self) -> tuple[str, str, str, int, datetime, str]:
        return (*self.student.key, self.ticket.number, self.timestamp,
                "да" if self.repeated else "нет")
