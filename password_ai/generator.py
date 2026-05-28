"""Secure password generation using cryptographic randomness."""

from __future__ import annotations

import secrets
import string


SYMBOLS = "!@#$%^&*_-+=?"
MIN_LENGTH = 12
MAX_LENGTH = 32
DEFAULT_LENGTH = 18


def generate_secure_password(length: int = DEFAULT_LENGTH) -> str:
    """Generate a strong password with all major character categories."""

    length = max(MIN_LENGTH, min(MAX_LENGTH, int(length)))
    required_characters = [
        secrets.choice(string.ascii_lowercase),
        secrets.choice(string.ascii_uppercase),
        secrets.choice(string.digits),
        secrets.choice(SYMBOLS),
    ]
    alphabet = string.ascii_letters + string.digits + SYMBOLS
    remaining = [secrets.choice(alphabet) for _ in range(length - len(required_characters))]
    return _secure_shuffle(required_characters + remaining)


def _secure_shuffle(characters: list[str]) -> str:
    shuffled: list[str] = []
    pool = characters[:]
    while pool:
        index = secrets.randbelow(len(pool))
        shuffled.append(pool.pop(index))
    return "".join(shuffled)
