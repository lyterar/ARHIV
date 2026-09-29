# Exam ticket registration

A console app that registers students for exam tickets. For each student it asks for
the last and first name, gives out a random ticket number from 1 to 20 and adds a row
to `journal.xlsx` in the current working directory.

## Requirements

- Python 3.10 or newer
- `openpyxl`, `readchar` (and `pytest` for the tests), listed in `requirements.txt`

## Installation

```bash
python -m venv .venv
# Windows (cmd / PowerShell):
.venv\Scripts\activate
# Linux / macOS:
source .venv/bin/activate

pip install -r requirements.txt
```

## Running

Run the app from the folder where the journal should be kept:

```bash
python main.py
```

```
Для выхода нажмите ESC.
Last name: Иванов
First name: Иван
Билет № 7
Last name:
```

- **ESC** pressed at the start of a prompt (before typing anything) exits the app.
  On Linux/macOS, readchar waits for one more key after ESC, so press ESC twice.
- **Enter** on an empty line, or a line of spaces only, shows a message and repeats
  the same prompt. Leading and trailing spaces are removed before saving.
- **Backspace** works as usual, including on the first character you typed.
- **Ctrl+C** also exits cleanly.

Use a real terminal: Windows Terminal, cmd, PowerShell, the VS Code terminal or any
Linux/macOS terminal. If standard input is not a terminal (piped input, some IDE output
panes, Git Bash's mintty without `winpty`), the app falls back to plain line input.
ESC is not available there; end the input with Ctrl+Z, Enter (Windows) or Ctrl+D.

## The journal file

`journal.xlsx` is created on the first record, with this header row:

| A         | B          | C            | D            |
|-----------|------------|--------------|--------------|
| Last name | First name | Номер билета | Дата и время |

Column D holds a real Excel date/time shown as `YYYY-MM-DD HH:MM:SS`, so it can be
sorted and filtered. If the file already exists, new rows go after the last filled row.
Existing rows are never changed.

### Reliability

- **Saved after every student.** For each record the app opens the file, adds the
  row, saves and closes it. The new version is written to a temporary file first and
  then swapped in, so killing the app or losing power mid-save cannot damage records
  that were already saved.
- **File open in Excel.** The app prints
  `Файл journal.xlsx открыт в другой программе. Закройте его и нажмите Enter для повторной попытки.`
  and waits. After Enter it tries to save the same record again, so nothing typed is
  lost. ESC at this point exits and prints the record that was not saved.
- **Damaged or unreadable file.** The app says so and waits for Enter as above. The
  damaged file is never overwritten: repair it, or rename it to start a new journal.

## Tests

```bash
python -m pytest
```

The tests create their Excel files in temporary directories and never touch your
`journal.xlsx`.

## Project structure

```
main.py               entry point, console input loop, ESC/Backspace handling
journal.py            Excel logic: create, append, save safely, report errors
tickets.py            ticket number generation
tests/test_journal.py tests for the Excel journal and input validation
tests/test_tickets.py tests for ticket generation
tests/test_main.py    tests for the prompt loop and key handling
requirements.txt      dependencies
pytest.ini            pytest settings (test path, import path)
```
