"""Console entry point: registers students for exam tickets in journal.xlsx."""

from __future__ import annotations

import random
import sys
from pathlib import Path
from typing import Callable, Protocol

import readchar

from journal import JOURNAL_FILENAME, Record, SaveResult, SaveStatus, append_record, clean_name
from tickets import generate_ticket

EXIT_HINT = "Для выхода нажмите ESC."
LAST_NAME_PROMPT = "Last name: "
FIRST_NAME_PROMPT = "First name: "
EMPTY_INPUT_MESSAGE = "Значение не может быть пустым. Повторите ввод."
RETRY_EXIT_HINT = "Для выхода без сохранения этой записи нажмите ESC."
GOODBYE_MESSAGE = "Работа завершена. До свидания!"

ESC = "\x1b"
ENTER_KEYS = frozenset({"\r", "\n"})

# (prompt, first typed character) -> the whole line entered by the user
Prefill = Callable[[str, str], str]


class ExitRequested(Exception):
    """The user asked to quit (ESC or end of input)."""


class Console(Protocol):
    """Source of user input; lets the main loop run against a fake in tests."""

    def read_line(self, prompt: str) -> str:
        """Show *prompt* and return the entered line; raise ExitRequested on ESC."""
        ...

    def wait_for_retry(self) -> None:
        """Block until Enter is pressed; raise ExitRequested on ESC."""
        ...


def is_escape(key: str) -> bool:
    """True if *key* is ESC.

    On POSIX readchar waits for one more byte after a lone ESC, so it arrives
    together with the next key as a two-character string.
    """
    return key == ESC or (len(key) == 2 and key[0] == ESC)


def is_character(key: str) -> bool:
    """True if *key* types a visible character."""
    return len(key) == 1 and key.isprintable()


def read_input(prompt: str = "") -> str:
    """input() that turns end of input (Ctrl+Z / Ctrl+D) into ExitRequested."""
    try:
        return input(prompt)
    except EOFError:
        print()
        raise ExitRequested from None


def echo_prefill(prompt: str, first: str) -> str:
    """Fallback: print *first* and read the rest of the line after it."""
    print(first, end="", flush=True)
    return first + read_input()


def readline_prefill(prompt: str, first: str) -> str:
    """POSIX: let readline put *first* into the editable line, so Backspace works."""
    import readline

    readline.set_startup_hook(lambda: readline.insert_text(first))
    try:
        print("\r", end="")  # input() redraws the prompt, so readline knows its width
        return read_input(prompt)
    finally:
        readline.set_startup_hook()


def no_character_pending() -> bool:
    """Default peek: say nothing is queued, so the first key goes through readchar."""
    return False


if sys.platform == "win32":
    import ctypes
    from ctypes import wintypes

    _KEY_EVENT = 0x0001
    _STD_INPUT_HANDLE = -10
    _WAIT_OBJECT_0 = 0x0000
    _WAIT_TIMEOUT = 0x0102
    _POLL_MS = 100  # wake up regularly so Ctrl+C is handled promptly

    class _KeyEventRecord(ctypes.Structure):
        _fields_ = [
            ("bKeyDown", wintypes.BOOL),
            ("wRepeatCount", wintypes.WORD),
            ("wVirtualKeyCode", wintypes.WORD),
            ("wVirtualScanCode", wintypes.WORD),
            ("uChar", wintypes.WCHAR),
            ("dwControlKeyState", wintypes.DWORD),
        ]

    class _EventUnion(ctypes.Union):
        # KEY_EVENT_RECORD is as large as any other event record (16 bytes),
        # so this has the full INPUT_RECORD size for every event type.
        _fields_ = [("KeyEvent", _KeyEventRecord)]

    class _InputRecord(ctypes.Structure):
        _fields_ = [("EventType", wintypes.WORD), ("Event", _EventUnion)]

    _kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    _kernel32.GetStdHandle.restype = wintypes.HANDLE
    _kernel32.WaitForSingleObject.argtypes = [wintypes.HANDLE, wintypes.DWORD]
    _kernel32.WaitForSingleObject.restype = wintypes.DWORD
    for _func in (_kernel32.PeekConsoleInputW, _kernel32.ReadConsoleInputW):
        _func.argtypes = [
            wintypes.HANDLE,
            ctypes.POINTER(_InputRecord),
            wintypes.DWORD,
            ctypes.POINTER(wintypes.DWORD),
        ]

    def windows_character_pending() -> bool:
        """Wait for the next key press without taking a typed character off the queue.

        True means the next key types a printable character. It stays queued,
        so input() echoes it and Backspace can erase it like any typed text.
        False means another key (ESC, Enter, Backspace, ...) that the caller
        reads with readchar. Key releases, Shift alone, mouse and focus events
        are discarded on the way.
        """
        handle = _kernel32.GetStdHandle(_STD_INPUT_HANDLE)
        event = _InputRecord()
        count = wintypes.DWORD()
        while True:
            status = _kernel32.WaitForSingleObject(handle, _POLL_MS)
            if status == _WAIT_TIMEOUT:
                continue
            if status != _WAIT_OBJECT_0 or not _kernel32.PeekConsoleInputW(
                handle, ctypes.byref(event), 1, ctypes.byref(count)
            ):
                return False  # not a console we can inspect: let readchar handle it
            if count.value == 0:
                continue
            key = event.Event.KeyEvent
            if event.EventType == _KEY_EVENT and key.bKeyDown and key.uChar != "\0":
                return key.uChar.isprintable()
            _kernel32.ReadConsoleInputW(handle, ctypes.byref(event), 1, ctypes.byref(count))


class KeyboardConsole:
    """Interactive terminal input: ESC is checked with readchar before each line."""

    def __init__(
        self, prefill: Prefill, character_pending: Callable[[], bool] = no_character_pending
    ) -> None:
        self._prefill = prefill
        self._character_pending = character_pending

    def read_line(self, prompt: str) -> str:
        """Show *prompt*; ESC exits, Enter gives "", a character starts the line."""
        print(prompt, end="", flush=True)
        while True:
            if self._character_pending():
                return read_input()  # the first character is still queued for input()
            key = readchar.readkey()
            if is_escape(key):
                print()
                raise ExitRequested
            if key in ENTER_KEYS:
                print()
                return ""
            if is_character(key):
                return self._prefill(prompt, key)
            # Backspace with nothing typed yet, arrows, F-keys, Tab: nothing to do.

    def wait_for_retry(self) -> None:
        """Wait for Enter (retry) or ESC (exit); other keys are ignored."""
        while True:
            key = readchar.readkey()
            if is_escape(key):
                raise ExitRequested
            if key in ENTER_KEYS:
                return


def keyboard_console() -> KeyboardConsole:
    """Build a KeyboardConsole that lets Backspace erase the first character.

    Windows: the first character is only peeked at and left for input().
    POSIX: readline inserts it into the editable line. Otherwise it is echoed.
    """
    if sys.platform == "win32":
        return KeyboardConsole(echo_prefill, windows_character_pending)
    try:
        import readline  # noqa: F401  (importing it enables line editing in input())
    except ImportError:
        return KeyboardConsole(echo_prefill)
    return KeyboardConsole(readline_prefill)


class PlainConsole:
    """Line-based input for when stdin is not a terminal (pipe, file, IDE runner)."""

    def read_line(self, prompt: str) -> str:
        """Show *prompt* and read one line."""
        return read_input(prompt)

    def wait_for_retry(self) -> None:
        """Wait for an empty line (Enter)."""
        read_input()


def ask_name(console: Console, prompt: str) -> str:
    """Repeat *prompt* until a non-blank value is entered; return it trimmed."""
    while True:
        name = clean_name(console.read_line(prompt))
        if name is not None:
            return name
        print(EMPTY_INPUT_MESSAGE)


def failure_message(result: SaveResult, file_name: str) -> str:
    """Explain why saving failed and what the user should do."""
    if result.status is SaveStatus.LOCKED:
        text = (
            f"Файл {file_name} открыт в другой программе. "
            "Закройте его и нажмите Enter для повторной попытки."
        )
    elif result.status is SaveStatus.CORRUPTED:
        text = (
            f"Файл {file_name} повреждён или не является книгой Excel ({result.error}). "
            "Восстановите его или переименуйте (будет создан новый журнал) "
            "и нажмите Enter для повторной попытки."
        )
    else:
        text = (
            f"Не удалось записать файл {file_name}: {result.error}. "
            "Устраните причину и нажмите Enter для повторной попытки."
        )
    return f"{text}\n{RETRY_EXIT_HINT}"


def save_with_retry(console: Console, path: Path, record: Record) -> None:
    """Append *record*; on failure explain, wait for Enter and retry the same record."""
    while True:
        result = append_record(path, record)
        if result.ok:
            return
        print(failure_message(result, path.name))
        try:
            console.wait_for_retry()
        except ExitRequested:
            print(
                f"Запись не сохранена: {record.last_name} {record.first_name}, "
                f"билет № {record.ticket}."
            )
            raise


def run(console: Console, journal_path: Path, rng: random.Random | None = None) -> None:
    """Register students one by one until the user presses ESC."""
    print(EXIT_HINT)
    try:
        while True:
            last_name = ask_name(console, LAST_NAME_PROMPT)
            first_name = ask_name(console, FIRST_NAME_PROMPT)
            ticket = generate_ticket(rng)
            print(f"Билет № {ticket}")
            save_with_retry(console, journal_path, Record(last_name, first_name, ticket))
    except ExitRequested:
        pass
    except KeyboardInterrupt:
        print()
    print(GOODBYE_MESSAGE)


def main() -> None:
    """Start the app with the journal in the current working directory."""
    console: Console = keyboard_console() if sys.stdin.isatty() else PlainConsole()
    run(console, Path.cwd() / JOURNAL_FILENAME)


if __name__ == "__main__":
    main()
