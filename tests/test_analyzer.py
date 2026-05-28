import unittest

from password_ai import PasswordStrengthAnalyzer, generate_secure_password


class AnalyzerTests(unittest.TestCase):
    def setUp(self):
        self.analyzer = PasswordStrengthAnalyzer.from_default_model()

    def test_common_password_scores_low(self):
        result = self.analyzer.analyze("password123")
        self.assertLess(result["score"], 45)
        self.assertEqual(result["ai_prediction"], "Weak")
        self.assertGreater(result["ai_confidence"], 0.5)
        self.assertTrue(any("common" in item.lower() for item in result["risks"]))
        self.assertTrue(any("passphrase" in item.lower() for item in result["suggestions"]))

    def test_strong_password_scores_high(self):
        result = self.analyzer.analyze("River-Matrix-Signal-47!")
        self.assertGreaterEqual(result["score"], 65)
        self.assertTrue(any("model prediction" in item.lower() for item in result["ai_explanation"]))

    def test_personal_context_is_detected(self):
        result = self.analyzer.analyze("Sachin@2026", context={"name": "Sachin"})
        self.assertTrue(any("personal" in item.lower() for item in result["risks"]))
        self.assertTrue(any("personal details" in item.lower() for item in result["suggestions"]))

    def test_repeated_password_gets_targeted_suggestion(self):
        result = self.analyzer.analyze("aaaa1111")
        self.assertTrue(any("repeated" in item.lower() for item in result["suggestions"]))

    def test_secure_generator_creates_complex_password(self):
        password = generate_secure_password(18)
        self.assertEqual(len(password), 18)
        self.assertTrue(any(ch.islower() for ch in password))
        self.assertTrue(any(ch.isupper() for ch in password))
        self.assertTrue(any(ch.isdigit() for ch in password))
        self.assertTrue(any(not ch.isalnum() for ch in password))

    def test_secure_generator_clamps_short_length(self):
        password = generate_secure_password(5)
        self.assertGreaterEqual(len(password), 12)


if __name__ == "__main__":
    unittest.main()
