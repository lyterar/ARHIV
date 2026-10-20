"""Launch the desktop app; input files default to the directory beside main.py."""

import argparse
import logging
from logging.handlers import RotatingFileHandler
from pathlib import Path
import tkinter as tk

from exam_tickets.ui import TicketApp


def main() -> None:
    parser = argparse.ArgumentParser(description="Генератор экзаменационных билетов")
    parser.add_argument("--data-dir", type=Path, default=Path(__file__).resolve().parent,
                        help="Папка с students.xlsx, tickets.docx и results.xlsx")
    args = parser.parse_args()
    handlers: list[logging.Handler] = [logging.StreamHandler()]
    try:
        handlers.append(RotatingFileHandler(Path(__file__).resolve().parent / "app.log",
                                           maxBytes=1_000_000, backupCount=2, encoding="utf-8"))
    except OSError:
        pass  # A read-only application folder must not prevent startup.
    logging.basicConfig(level=logging.INFO, handlers=handlers,
                        format="%(asctime)s %(levelname)s %(name)s: %(message)s")
    root = tk.Tk()
    TicketApp(root, args.data_dir)
    root.mainloop()


if __name__ == "__main__":
    main()
