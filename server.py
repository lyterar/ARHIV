"""Small JSON API (standard library only) used by the React frontend.

POST /api/tickets   {"last_name": ..., "first_name": ...} -> draws a ticket, writes it
                    to journal.xlsx and returns the ticket number and generated text.
GET  /api/tickets/N -> the text of ticket N without registering anybody.
"""

from __future__ import annotations

import json
import random
import re
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any

from journal import JOURNAL_FILENAME, Record, append_record, clean_name
from tickets import generate_ticket, ticket_questions, ticket_text

MAX_BODY = 4096
TICKET_PATH = re.compile(r"^/api/tickets/(\d+)$")


def ticket_payload(ticket: int, last_name: str = "", first_name: str = "") -> dict[str, Any]:
    """JSON-ready description of a ticket."""
    return {
        "ticket": ticket,
        "questions": list(ticket_questions(ticket)),
        "text": ticket_text(ticket, last_name, first_name),
    }


def make_handler(journal_path: Path, rng: random.Random | None = None) -> type:
    """Create a request handler class bound to *journal_path*."""

    class Handler(BaseHTTPRequestHandler):
        def _send(self, status: int, body: dict[str, Any]) -> None:
            data = json.dumps(body, ensure_ascii=False).encode("utf-8")
            self.send_response(status)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)

        def do_GET(self) -> None:  # noqa: N802
            match = TICKET_PATH.match(self.path)
            if not match:
                self._send(404, {"error": "not found"})
                return
            try:
                self._send(200, ticket_payload(int(match.group(1))))
            except ValueError as exc:
                self._send(404, {"error": str(exc)})

        def do_POST(self) -> None:  # noqa: N802
            if self.path != "/api/tickets":
                self._send(404, {"error": "not found"})
                return
            try:
                length = int(self.headers.get("Content-Length", 0))
                if not 0 < length <= MAX_BODY:
                    raise ValueError
                data = json.loads(self.rfile.read(length))
                last_name = clean_name(str(data["last_name"]))
                first_name = clean_name(str(data["first_name"]))
                if last_name is None or first_name is None:
                    raise ValueError
            except (ValueError, KeyError, TypeError):
                self._send(400, {"error": "Укажите фамилию и имя"})
                return
            ticket = generate_ticket(rng)
            result = append_record(journal_path, Record(last_name, first_name, ticket))
            if not result.ok:
                self._send(500, {"error": f"Не удалось сохранить журнал: {result.status.value}"})
                return
            self._send(201, ticket_payload(ticket, last_name, first_name))

        def log_message(self, format: str, *args: Any) -> None:  # noqa: A002
            pass

    return Handler


def main() -> None:
    """Serve the API on http://127.0.0.1:8000."""
    server = ThreadingHTTPServer(("127.0.0.1", 8000), make_handler(Path.cwd() / JOURNAL_FILENAME))
    print("API: http://127.0.0.1:8000")
    server.serve_forever()


if __name__ == "__main__":
    main()
