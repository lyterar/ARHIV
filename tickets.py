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
