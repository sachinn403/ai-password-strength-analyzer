"""High-level password analysis service."""

from __future__ import annotations

from pathlib import Path
from typing import Mapping

from .explainability import explain_ai_decision
from .feature_engineering import extract_features
from .model import LogisticPasswordModel
from .suggestions import create_improvement_suggestions


PROJECT_ROOT = Path(__file__).resolve().parents[1]
DEFAULT_MODEL_PATH = PROJECT_ROOT / "data" / "model.json"


class PasswordStrengthAnalyzer:
    """Combines ML probability, entropy, and explainable security feedback."""

    def __init__(self, model: LogisticPasswordModel):
        self.model = model

    @classmethod
    def from_default_model(cls) -> "PasswordStrengthAnalyzer":
        return cls(LogisticPasswordModel.from_json(DEFAULT_MODEL_PATH))

    def analyze(self, password: str, context: Mapping[str, str] | None = None) -> dict:
        password = password or ""
        bundle = extract_features(password, context)

        if not password:
            return {
                "score": 0,
                "label": "Empty",
                "ai_probability": 0,
                "ai_confidence": 0,
                "ai_prediction": "Waiting",
                "entropy_bits": 0,
                "effective_entropy_bits": 0,
                "estimated_crack_time": {
                    "offline_fast_hash": "Waiting",
                    "online_throttled": "Waiting",
                },
                "strengths": [],
                "risks": ["Password field is empty."],
                "suggestions": ["Create a password or passphrase before testing strength."],
                "ai_explanation": ["Enter a password to run the AI explanation."],
                "features": bundle.raw,
                "model_version": self.model.version,
            }

        prediction_probabilities = self.model.predict_proba(bundle.model)
        ai_probability = prediction_probabilities["strong"]
        entropy_bits = float(bundle.raw["entropy_bits"])
        entropy_score = min(100.0, entropy_bits * 1.15)
        ml_score = ai_probability * 100
        pattern_penalty = self._pattern_penalty(bundle.raw)
        score = round(max(0, min(100, (ml_score * 0.68) + (entropy_score * 0.32) - pattern_penalty)))
        ai_prediction = "Strong" if ai_probability >= 0.5 else "Weak"
        ai_confidence = max(ai_probability, 1 - ai_probability)
        crack_time = self._crack_time_estimate(entropy_bits, ai_probability, pattern_penalty)

        return {
            "score": score,
            "label": self._label(score),
            "ai_probability": round(ai_probability, 3),
            "ai_confidence": round(ai_confidence, 3),
            "ai_prediction": ai_prediction,
            "entropy_bits": round(entropy_bits, 2),
            "effective_entropy_bits": crack_time["effective_entropy_bits"],
            "estimated_crack_time": crack_time["estimated_crack_time"],
            "strengths": self._strengths(bundle.raw),
            "risks": self._risks(bundle.raw),
            "suggestions": create_improvement_suggestions(bundle.raw, score),
            "ai_explanation": explain_ai_decision(
                bundle.raw,
                score,
                ai_prediction,
                ai_confidence,
                ai_probability,
                pattern_penalty,
            ),
            "features": bundle.raw,
            "model_version": self.model.version,
        }

    @staticmethod
    def _label(score: int) -> str:
        if score < 25:
            return "Very Weak"
        if score < 45:
            return "Weak"
        if score < 65:
            return "Fair"
        if score < 85:
            return "Strong"
        return "Excellent"

    @staticmethod
    def _pattern_penalty(raw: dict) -> float:
        penalty = 0.0
        repeat_excess = max(0.0, float(raw["repeat_ratio"]) - 0.18)
        penalty += min(repeat_excess * 35, 16)
        if int(raw["longest_run"]) >= 3:
            penalty += min(int(raw["longest_run"]) * 2, 12)
        penalty += min(int(raw["sequence_triplets"]) * 4, 16)
        penalty += min(int(raw["keyboard_walks"]) * 5, 15)
        penalty += min(len(raw["common_words"]) * 10, 20)
        penalty += min(len(raw["context_matches"]) * 12, 24)
        if raw["has_year"]:
            penalty += 5
        if raw["date_like"]:
            penalty += 8
        if raw["starts_cap_ends_digits"]:
            penalty += 7
        if raw["only_digits"] or raw["only_letters"]:
            penalty += 9
        return penalty

    @staticmethod
    def _strengths(raw: dict) -> list[str]:
        strengths: list[str] = []
        if int(raw["length"]) >= 12:
            strengths.append("Good length for resisting brute-force guessing.")
        if int(raw["character_classes"]) >= 3:
            strengths.append("Uses multiple character categories.")
        if float(raw["repeat_ratio"]) < 0.18 and int(raw["unique_chars"]) >= 8:
            strengths.append("Characters are reasonably varied.")
        if not raw["common_words"] and not raw["keyboard_walks"]:
            strengths.append("No obvious common word or keyboard walk was detected.")
        return strengths

    @staticmethod
    def _risks(raw: dict) -> list[str]:
        risks: list[str] = []
        if int(raw["length"]) < 10:
            risks.append("Short length makes brute-force attacks easier.")
        if int(raw["character_classes"]) < 3:
            risks.append("Limited character variety reduces search complexity.")
        if float(raw["repeat_ratio"]) >= 0.34 or int(raw["longest_run"]) >= 3:
            risks.append("Repeated characters reduce unpredictability.")
        if int(raw["sequence_triplets"]) > 0:
            risks.append("Alphabetical or numeric sequences are easy to guess.")
        if int(raw["keyboard_walks"]) > 0:
            risks.append("Keyboard patterns such as qwerty-style walks were found.")
        if raw["common_words"]:
            risks.append("Contains common words: " + ", ".join(raw["common_words"]) + ".")
        if raw["context_matches"]:
            risks.append("Contains personal context: " + ", ".join(raw["context_matches"]) + ".")
        if raw["has_year"] or raw["date_like"]:
            risks.append("Dates or years are commonly tried by attackers.")
        if raw["starts_cap_ends_digits"]:
            risks.append("Capitalized word plus ending digits is a predictable shape.")
        return risks or ["No major pattern risk detected."]

    @staticmethod
    def _crack_time_estimate(entropy_bits: float, ai_probability: float, pattern_penalty: float) -> dict[str, object]:
        effective_entropy = max(0.0, entropy_bits * (0.22 + ai_probability * 0.72) - pattern_penalty * 0.8)
        return {
            "effective_entropy_bits": round(effective_entropy, 2),
            "estimated_crack_time": {
                "offline_fast_hash": _offline_attack_band(effective_entropy),
                "online_throttled": _online_attack_band(effective_entropy),
            },
        }


def _offline_attack_band(effective_entropy: float) -> str:
    if effective_entropy < 30:
        return "Seconds"
    if effective_entropy < 50:
        return "Minutes to hours"
    if effective_entropy < 70:
        return "Months to years"
    if effective_entropy < 90:
        return "Centuries"
    return "Thousands of years"


def _online_attack_band(effective_entropy: float) -> str:
    if effective_entropy < 30:
        return "Seconds to minutes"
    if effective_entropy < 50:
        return "Days to months"
    if effective_entropy < 70:
        return "Years"
    if effective_entropy < 90:
        return "Centuries"
    return "Thousands of years"
