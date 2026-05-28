"""Small logistic model used by the password analyzer."""

from __future__ import annotations

import json
import math
from pathlib import Path


DEFAULT_MODEL = {
    "version": "fallback-1.0",
    "bias": -4.15,
    "thresholds": {
        "very_weak": 25,
        "weak": 45,
        "fair": 65,
        "strong": 85,
    },
    "weights": {
        "length_norm": 2.2,
        "entropy_norm": 3.6,
        "unique_ratio": 1.2,
        "class_ratio": 1.5,
        "has_lower": 0.2,
        "has_upper": 0.55,
        "has_digit": 0.5,
        "has_symbol": 0.8,
        "repeat_penalty": -2.4,
        "run_penalty": -1.1,
        "sequence_penalty": -1.4,
        "keyboard_penalty": -1.7,
        "common_word_penalty": -2.2,
        "year_penalty": -0.8,
        "date_penalty": -1.1,
        "shape_penalty": -0.9,
        "only_digits_penalty": -2.2,
        "only_letters_penalty": -0.9,
        "suffix_digit_penalty": -0.8,
        "personal_penalty": -1.9,
    },
}


class LogisticPasswordModel:
    """Predicts whether a password looks resistant to common attacks."""

    def __init__(self, weights: dict[str, float], bias: float, version: str = "custom"):
        self.weights = weights
        self.bias = bias
        self.version = version

    @classmethod
    def from_dict(cls, payload: dict) -> "LogisticPasswordModel":
        return cls(
            weights={key: float(value) for key, value in payload["weights"].items()},
            bias=float(payload["bias"]),
            version=str(payload.get("version", "custom")),
        )

    @classmethod
    def from_json(cls, path: str | Path) -> "LogisticPasswordModel":
        model_path = Path(path)
        if not model_path.exists():
            return cls.from_dict(DEFAULT_MODEL)
        with model_path.open("r", encoding="utf-8") as handle:
            return cls.from_dict(json.load(handle))

    def to_dict(self) -> dict:
        return {
            "version": self.version,
            "bias": self.bias,
            "weights": self.weights,
        }

    def predict_probability(self, features: dict[str, float]) -> float:
        value = self.bias
        for key, weight in self.weights.items():
            value += weight * float(features.get(key, 0.0))
        return 1 / (1 + math.exp(-value))

    def predict_proba(self, features: dict[str, float]) -> dict[str, float]:
        """Return weak/strong class probabilities similar to ML predict_proba."""

        strong_probability = self.predict_probability(features)
        return {
            "weak": 1 - strong_probability,
            "strong": strong_probability,
        }
