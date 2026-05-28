# AI-Powered Password Strength Analyzer

This is an original BCA final-year project that checks password strength using a local AI-style logistic model, entropy calculation, and explainable pattern detection. It runs without third-party Python packages.

## Main Features

- Real-time web interface for password analysis.
- Machine-learning model trained from synthetic password examples.
- Entropy, character variety, repetition, keyboard pattern, date, year, and dictionary-word checks.
- Optional name, email, and mobile context detection.
- Intelligent improvement suggestions for weak passwords.
- AI model confidence and weak/strong prediction display.
- Animated red/orange/green strength circle.
- Loading spinner while password analysis is running.
- Secure password generator using Python cryptographic randomness.
- Copy generated or typed password to clipboard.
- AI explanation section that explains why the model predicted weak or strong.
- Crack-time estimate for offline and online attack scenarios.
- No password storage.

## How To Run

```bash
python app.py
```

Open:

```text
http://127.0.0.1:8000
```

## Train The Model Again

```bash
python scripts/train_model.py
```

The script writes a fresh model to `data/model.json`.

## Run Tests

```bash
python -m unittest discover -s tests
```

## Project Structure

```text
app.py                         Local HTTP server and API
password_ai/analyzer.py        Main password analysis logic
password_ai/explainability.py  AI decision explanation logic
password_ai/feature_engineering.py  Numeric feature extraction
password_ai/generator.py       Secure password generator
password_ai/model.py           Logistic model loader and predictor
password_ai/suggestions.py     Intelligent improvement suggestions
scripts/train_model.py         Synthetic model training script
static/                        Web interface files
tests/                         Unit tests
PROJECT_REPORT.md              Final-year project report draft
PLAGIARISM_GUIDE.md            Originality and submission guidance
```

## Academic Note

The code and report are written freshly for this project. A plagiarism percentage cannot be guaranteed because each college may use a different checker, database, and exclusion policy. You should add your own screenshots, college formatting, acknowledgements, and viva observations before final submission.
