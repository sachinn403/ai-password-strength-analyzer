"""Feature extraction for the password strength model.

The analyzer keeps every calculation local. Passwords are transformed into
numeric signals and the original password is never stored by the app.
"""

from __future__ import annotations

import math
import re
from dataclasses import dataclass
from typing import Mapping


COMMON_WORDS = {
    "admin",
    "bharat",
    "college",
    "computer",
    "cricket",
    "dragon",
    "football",
    "google",
    "hello",
    "india",
    "iloveyou",
    "krishna",
    "letmein",
    "master",
    "monkey",
    "nishad",
    "password",
    "princess",
    "qwerty",
    "sachin",
    "student",
    "welcome",
}

KEYBOARD_LINES = (
    "1234567890",
    "qwertyuiop",
    "asdfghjkl",
    "zxcvbnm",
)

LEET_TRANSLATION = str.maketrans(
    {
        "0": "o",
        "1": "i",
        "3": "e",
        "4": "a",
        "5": "s",
        "7": "t",
        "@": "a",
        "$": "s",
        "!": "i",
    }
)

YEAR_PATTERN = re.compile(r"(19[5-9]\d|20[0-4]\d)")
DATE_PATTERN = re.compile(
    r"(\d{1,2}[-_/]\d{1,2}[-_/]\d{2,4})|(\d{2,4}[-_/]\d{1,2}[-_/]\d{1,2})"
)


@dataclass(frozen=True)
class FeatureBundle:
    """Container for raw facts and model-ready features."""

    raw: dict[str, float | int | str | list[str]]
    model: dict[str, float]


def normalize_text(value: str) -> str:
    return re.sub(r"[^a-z0-9]", "", value.lower().translate(LEET_TRANSLATION))


def _charset_size(password: str) -> int:
    size = 0
    if any(ch.islower() for ch in password):
        size += 26
    if any(ch.isupper() for ch in password):
        size += 26
    if any(ch.isdigit() for ch in password):
        size += 10
    if any(not ch.isalnum() for ch in password):
        size += 33
    return max(size, 1)


def _max_run(password: str) -> int:
    if not password:
        return 0
    longest = current = 1
    for index in range(1, len(password)):
        if password[index] == password[index - 1]:
            current += 1
            longest = max(longest, current)
        else:
            current = 1
    return longest


def _sequence_triplets(password: str) -> int:
    count = 0
    lowered = password.lower()
    for index in range(max(0, len(lowered) - 2)):
        chars = lowered[index : index + 3]
        numbers = [ord(ch) for ch in chars]
        if numbers[1] - numbers[0] == numbers[2] - numbers[1] and abs(numbers[1] - numbers[0]) == 1:
            count += 1
    return count


def _keyboard_walks(password: str) -> int:
    count = 0
    lowered = password.lower()
    for line in KEYBOARD_LINES:
        reverse = line[::-1]
        for index in range(max(0, len(lowered) - 2)):
            token = lowered[index : index + 3]
            if token in line or token in reverse:
                count += 1
    return count


def _suffix_digit_count(password: str) -> int:
    match = re.search(r"\d+$", password)
    return len(match.group(0)) if match else 0


def _context_matches(password: str, context: Mapping[str, str] | None) -> list[str]:
    if not context:
        return []

    normalized_password = normalize_text(password)
    hits: list[str] = []
    for key, value in context.items():
        cleaned = normalize_text(str(value))
        if len(cleaned) >= 3 and cleaned in normalized_password:
            hits.append(key)
    return hits


def _common_word_hits(password: str) -> list[str]:
    normalized = normalize_text(password)
    return sorted(word for word in COMMON_WORDS if len(word) >= 4 and word in normalized)


def extract_features(password: str, context: Mapping[str, str] | None = None) -> FeatureBundle:
    """Convert a password into explainable numeric features."""

    length = len(password)
    lower_count = sum(ch.islower() for ch in password)
    upper_count = sum(ch.isupper() for ch in password)
    digit_count = sum(ch.isdigit() for ch in password)
    symbol_count = sum(not ch.isalnum() for ch in password)
    character_classes = sum(
        value > 0 for value in (lower_count, upper_count, digit_count, symbol_count)
    )
    unique_chars = len(set(password))
    charset_size = _charset_size(password)
    entropy_bits = length * math.log2(charset_size) if password else 0.0

    longest_run = _max_run(password)
    sequence_triplets = _sequence_triplets(password)
    keyboard_walks = _keyboard_walks(password)
    suffix_digits = _suffix_digit_count(password)
    common_hits = _common_word_hits(password)
    context_hits = _context_matches(password, context)

    has_year = bool(YEAR_PATTERN.search(password))
    date_like = bool(DATE_PATTERN.search(password))
    starts_cap_ends_digits = bool(re.match(r"^[A-Z][a-z]+[\W_]*\d+$", password))
    only_digits = password.isdigit()
    only_letters = password.isalpha()

    repeat_ratio = (length - unique_chars) / length if length else 0.0
    model_features = {
        "length_norm": min(length, 32) / 32,
        "entropy_norm": min(entropy_bits, 140) / 140,
        "unique_ratio": unique_chars / length if length else 0.0,
        "class_ratio": character_classes / 4,
        "has_lower": float(lower_count > 0),
        "has_upper": float(upper_count > 0),
        "has_digit": float(digit_count > 0),
        "has_symbol": float(symbol_count > 0),
        "repeat_penalty": min(repeat_ratio, 1.0),
        "run_penalty": min(longest_run / 6, 1.0),
        "sequence_penalty": min(sequence_triplets / 4, 1.0),
        "keyboard_penalty": min(keyboard_walks / 3, 1.0),
        "common_word_penalty": min(len(common_hits) / 2, 1.0),
        "year_penalty": float(has_year),
        "date_penalty": float(date_like),
        "shape_penalty": float(starts_cap_ends_digits),
        "only_digits_penalty": float(only_digits),
        "only_letters_penalty": float(only_letters),
        "suffix_digit_penalty": min(suffix_digits / 4, 1.0),
        "personal_penalty": min(len(context_hits) / 2, 1.0),
    }

    raw = {
        "length": length,
        "lower_count": lower_count,
        "upper_count": upper_count,
        "digit_count": digit_count,
        "symbol_count": symbol_count,
        "character_classes": character_classes,
        "unique_chars": unique_chars,
        "charset_size": charset_size,
        "entropy_bits": round(entropy_bits, 2),
        "longest_run": longest_run,
        "repeat_ratio": round(repeat_ratio, 3),
        "sequence_triplets": sequence_triplets,
        "keyboard_walks": keyboard_walks,
        "suffix_digits": suffix_digits,
        "common_words": common_hits,
        "context_matches": context_hits,
        "has_year": int(has_year),
        "date_like": int(date_like),
        "starts_cap_ends_digits": int(starts_cap_ends_digits),
        "only_digits": int(only_digits),
        "only_letters": int(only_letters),
    }

    return FeatureBundle(raw=raw, model=model_features)
