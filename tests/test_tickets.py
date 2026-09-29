"""Tests for ticket number generation."""

import random

import pytest

from tickets import MAX_TICKET, MIN_TICKET, generate_ticket


def test_range_constants_match_spec() -> None:
    assert (MIN_TICKET, MAX_TICKET) == (1, 20)


@pytest.mark.parametrize("seed", range(10))
def test_ticket_is_always_between_1_and_20(seed: int) -> None:
    rng = random.Random(seed)
    for _ in range(500):
        ticket = generate_ticket(rng)
        assert isinstance(ticket, int)
        assert 1 <= ticket <= 20


def test_default_random_source_stays_in_range() -> None:
    assert all(1 <= generate_ticket() <= 20 for _ in range(1000))


def test_every_ticket_including_bounds_can_be_drawn() -> None:
    rng = random.Random(0)
    assert {generate_ticket(rng) for _ in range(2000)} == set(range(1, 21))


def test_same_seed_gives_same_sequence() -> None:
    first, second = random.Random(42), random.Random(42)
    assert [generate_ticket(first) for _ in range(50)] == [
        generate_ticket(second) for _ in range(50)
    ]
