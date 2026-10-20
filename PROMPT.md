# Lab 2 implementation specification

Implement a desktop exam-ticket application using Python 3.10+ and Tkinter/ttk.
Codex replaces Copilot for implementation and test generation in this assignment.
Do not use JavaFX, Windows Forms, a browser UI or a web server.

## Inputs

- `students.xlsx`: each worksheet is a group, using its exact worksheet name.
  Row 1 contains `Фамилия` and `Имя` in A and B. Each following row is one student.
  Trim surrounding whitespace. Preserve worksheet and student order. Keep groups
  with no students. Warn about invalid rows and exact duplicate students.
- `tickets.docx`: a heading is a separate paragraph `Билет N`, where N is a
  positive integer. Questions are separate paragraphs `1. text`, `2. text`,
  `3. text`. Manual line breaks may also delimit lines.
  Skip malformed headings and tickets with fewer than three questions, log every
  error, and continue with valid tickets. Do not join questions across ticket
  boundaries. Use the first three questions if more are supplied and log it.
- Missing/unreadable/corrupt sources must produce a friendly Russian message.
  Keep the window open and allow the user to reload after fixing the files.

## Interface

Show two read-only combo boxes, for the group and student. Initially neither is
selected and the Generate button is disabled. Changing a group clears the student
selection and loads only that group's students as `Фамилия Имя`. An empty group
keeps the student list empty and generation disabled. Enable generation only when
both choices are valid. Display the assigned number and three questions in a
separate, scrollable result window. Escape closes that window, keeps the main app
alive, and prepares it for the next student. Include reload and data-folder controls.
Default file paths must be relative to main.py, not the shell's working directory.

## Assignment and persistence

Create `results.xlsx` on first launch if missing, with exactly these six columns:
`Группа`, `Фамилия`, `Имя`, `Номер билета`, `Дата и время`, `Повтор (да/нет)`.
Identify a student by the tuple (group, last name, first name). Read the journal
from the first data row onwards. If the student exists, reuse the ticket number
from the first matching row, regardless of later rows, and append `да`. Otherwise
choose a ticket uniformly and append `нет`.

For a valid contiguous catalog this is a draw from 1..N. If invalid ticket blocks
have been skipped, choose uniformly among surviving original ticket IDs, even if
they contain gaps; never renumber tickets or select a missing ticket. If a
previously assigned number no longer exists in the valid catalog, report the
problem instead of silently reassigning the student.

Append one row per successful click, including repeated selections. Never modify
previous rows. Save immediately, before presenting the result. Store a real Excel
datetime with second precision and `yyyy-mm-dd hh:mm:ss` formatting. Save via a
temporary file in the same directory, flush, then atomically replace the journal.
Serialize the read/assign/append/save transaction between application instances.
Preserve the chosen ticket during a failed write and Retry. Offer Retry/Cancel
when Excel blocks the file. Never report an unsaved result as saved. Do not
overwrite a corrupt or incompatible journal. Treat names as text, not formulas.

## Deliverables and validation

Separate typed data models, source parsing, persistence, assignment service and UI.
Provide Russian README, this English prompt, requirements, a Windows launcher,
synthetic student data and 20 sample Python tickets based on Lab 1 plus a third
question each. Do not pretend the sample students are the user's actual roster.
Add pytest tests for Word parsing, group lookup, repeated assignments, first-row
precedence, independent groups, immediate persistence, unchanged old rows,
corruption, locked-file retry, and the UI acceptance scenarios including Escape.
Run tests on the available environment and record actual results and limitations.
Create branch `lab2-ui-tickets`, commit with message `Lab 2: UI-генератор билетов`,
and push to the existing GitHub repository if authentication is available.
