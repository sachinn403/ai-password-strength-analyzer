"""Human-readable explanation for model decisions."""

from __future__ import annotations


def explain_ai_decision(
    raw: dict,
    score: int,
    ai_prediction: str,
    ai_confidence: float,
    ai_probability: float,
    pattern_penalty: float,
) -> list[str]:
    """Explain why the model and rule layer produced the current result."""

    explanation: list[str] = [
        f"Model prediction is {ai_prediction} with {round(ai_confidence * 100)}% confidence.",
    ]

    entropy_bits = float(raw["entropy_bits"])
    if entropy_bits < 45:
        explanation.append("Entropy is low, so an attacker needs fewer guesses.")
    elif entropy_bits < 70:
        explanation.append("Entropy is moderate, but pattern checks still affect the final score.")
    else:
        explanation.append("Raw entropy is high, so length and character range are helping the password.")

    pattern_reasons = _pattern_reasons(raw)
    if pattern_reasons:
        explanation.append("Pattern penalty was applied for " + _join_words(pattern_reasons) + ".")
    else:
        explanation.append("No major dictionary, keyboard, date, or personal-information pattern was detected.")

    if ai_probability >= 0.5 and score >= 65:
        explanation.append("The ML score and entropy score agree that this password is difficult to guess.")
    elif ai_probability < 0.5 and score < 45:
        explanation.append("The ML model and pattern checks both indicate a weak password.")
    else:
        explanation.append("The final score combines ML probability, entropy, and detected pattern penalties.")

    if pattern_penalty > 0:
        explanation.append(f"Pattern penalty reduced the final score by about {round(pattern_penalty)} points.")

    return explanation[:5]


def _pattern_reasons(raw: dict) -> list[str]:
    reasons: list[str] = []
    if raw["common_words"]:
        reasons.append("common words")
    if int(raw["keyboard_walks"]) > 0:
        reasons.append("keyboard patterns")
    if int(raw["sequence_triplets"]) > 0:
        reasons.append("simple sequences")
    if raw["context_matches"]:
        reasons.append("personal details")
    if raw["has_year"] or raw["date_like"]:
        reasons.append("dates or years")
    if float(raw["repeat_ratio"]) >= 0.34 or int(raw["longest_run"]) >= 3:
        reasons.append("repeated characters")
    if raw["starts_cap_ends_digits"]:
        reasons.append("capitalized word with ending digits")
    return reasons


def _join_words(words: list[str]) -> str:
    if len(words) <= 1:
        return "".join(words)
    if len(words) == 2:
        return words[0] + " and " + words[1]
    return ", ".join(words[:-1]) + ", and " + words[-1]
