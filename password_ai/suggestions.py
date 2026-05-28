"""Intelligent improvement suggestions for weak passwords."""

from __future__ import annotations


def create_improvement_suggestions(raw: dict, score: int) -> list[str]:
    """Create targeted password improvement suggestions from extracted features."""

    suggestions: list[str] = []

    def add(message: str) -> None:
        if message not in suggestions:
            suggestions.append(message)

    length = int(raw["length"])
    missing_classes = _missing_character_classes(raw)

    if score < 45:
        add(
            "Start with a new passphrase instead of editing this weak pattern: use 4 unrelated words plus one number or symbol."
        )
    elif score < 65:
        add("Improve the weakest detected pattern first, then test the password again.")

    if length < 14:
        add(f"Increase length from {length} to at least 14 characters.")
    elif length < 18 and score < 65:
        add("Add a few more characters to make brute-force guessing harder.")

    if missing_classes and int(raw["character_classes"]) < 3:
        add("Add " + _join_words(missing_classes) + " to improve character variety.")

    if raw["common_words"]:
        add("Replace common word(s) " + ", ".join(raw["common_words"]) + " with unrelated words.")

    if int(raw["keyboard_walks"]) > 0:
        add("Remove keyboard patterns such as qwerty, asdf, or 123-style movement.")

    if int(raw["sequence_triplets"]) > 0:
        add("Avoid simple alphabetical or numeric sequences like abc, 123, or 987.")

    if raw["context_matches"]:
        add("Remove personal details found in the password: " + ", ".join(raw["context_matches"]) + ".")

    if raw["has_year"] or raw["date_like"]:
        add("Replace dates and years with random words or characters that are not connected to you.")

    if int(raw["suffix_digits"]) >= 2:
        add("Avoid simply placing digits at the end; mix numbers inside a longer passphrase instead.")

    if float(raw["repeat_ratio"]) >= 0.34 or int(raw["longest_run"]) >= 3:
        add("Reduce repeated characters and use more unique characters.")

    if raw["only_digits"]:
        add("Do not use only numbers; combine words, letters, symbols, and digits.")

    if raw["only_letters"]:
        add("Do not use only letters; add a number or symbol inside a longer password.")

    if not suggestions:
        add("Password looks solid; use a password manager to keep it unique.")

    return suggestions[:6]


def _missing_character_classes(raw: dict) -> list[str]:
    missing: list[str] = []
    if int(raw["lower_count"]) == 0:
        missing.append("lowercase letters")
    if int(raw["upper_count"]) == 0:
        missing.append("uppercase letters")
    if int(raw["digit_count"]) == 0:
        missing.append("digits")
    if int(raw["symbol_count"]) == 0:
        missing.append("special characters")
    return missing


def _join_words(words: list[str]) -> str:
    if len(words) <= 1:
        return "".join(words)
    if len(words) == 2:
        return words[0] + " and " + words[1]
    return ", ".join(words[:-1]) + ", and " + words[-1]
