"""Exam ticket number generation."""

from __future__ import annotations

import random

MIN_TICKET = 1
MAX_TICKET = 20


def generate_ticket(rng: random.Random | None = None) -> int:
    """Return a random ticket number from MIN_TICKET to MAX_TICKET inclusive.

    Pass a seeded ``random.Random`` to get a reproducible sequence (e.g. in tests);
    by default the module-level generator is used.
    """
    if rng is None:
        return random.randint(MIN_TICKET, MAX_TICKET)
    return rng.randint(MIN_TICKET, MAX_TICKET)


# Two questions per ticket, index 0 is ticket № 1.
QUESTIONS: tuple[tuple[str, str], ...] = (
    ("Переменные и базовые типы данных Python", "Условный оператор if/elif/else"),
    ("Строки: методы и форматирование", "Цикл for и функция range"),
    ("Списки и срезы", "Цикл while, break и continue"),
    ("Кортежи и множества", "Определение функций, аргументы по умолчанию"),
    ("Словари и их методы", "Область видимости переменных"),
    ("Генераторы списков и словарей", "Функции *args и **kwargs"),
    ("Работа с файлами: open и with", "Исключения: try/except/finally"),
    ("Модули и импорт", "Пакеты и виртуальные окружения"),
    ("Классы и объекты", "Конструктор __init__ и атрибуты"),
    ("Наследование и super()", "Полиморфизм и переопределение методов"),
    ("Инкапсуляция и свойства (property)", "Магические методы __str__ и __repr__"),
    ("Итераторы и генераторы", "Декораторы"),
    ("Лямбда-функции, map, filter, sorted", "Рекурсия"),
    ("Работа с форматами JSON и CSV", "Регулярные выражения (модуль re)"),
    ("Работа с датой и временем (datetime)", "Модуль random"),
    ("Аннотации типов", "Датаклассы (dataclasses)"),
    ("Тестирование: pytest и assert", "Отладка и логирование"),
    ("Работа с Excel через openpyxl", "Запись и чтение xlsx-журнала"),
    ("Контекстные менеджеры", "Изменяемые и неизменяемые типы"),
    ("Основы алгоритмов: поиск и сортировка", "Оценка сложности алгоритмов"),
)


def ticket_questions(ticket: int) -> tuple[str, str]:
    """Return the two exam questions of *ticket*; raise ValueError if out of range."""
    if not MIN_TICKET <= ticket <= MAX_TICKET:
        raise ValueError(f"ticket must be between {MIN_TICKET} and {MAX_TICKET}")
    return QUESTIONS[ticket - MIN_TICKET]


def ticket_text(ticket: int, last_name: str = "", first_name: str = "") -> str:
    """Build the printable text of an exam ticket (optionally addressed to a student)."""
    first, second = ticket_questions(ticket)
    lines = [f"ЭКЗАМЕНАЦИОННЫЙ БИЛЕТ № {ticket}"]
    student = f"{last_name} {first_name}".strip()
    if student:
        lines.append(f"Студент: {student}")
    lines += ["", f"1. {first}", f"2. {second}"]
    return "\n".join(lines)
