# 🇩🇪 DeutschFlash — German Vocabulary Learning App

A premium Streamlit-based flashcard application for learning German vocabulary with visual mnemonics, MCQ quizzes, and auto-fetched images.

## Features

- **Day-wise learning** — Vocabulary organized into 5 daily sessions
- **MCQ quiz** — 4 options per card with instant feedback
- **Auto-images** — DuckDuckGo image search, cached in Excel
- **Marathon mode** — Practice all words in one session
- **Score tracking** — Streak counter, accuracy, and final summary
- **Multi-level** — Supports A1, A2, B1, B2 vocabulary files

## Quick Start

```bash
# Install dependencies
pip install -r requirements.txt

# Generate Excel data files (first time only)
python -m lib.data_generator

# Run the app
streamlit run app.py
```

## Project Structure

```
PROJECT VOCAB/
├── app.py                  # Main Streamlit application
├── requirements.txt        # Python dependencies
├── .streamlit/config.toml  # Dark theme config
├── data/
│   ├── A1_vocab.xlsx       # A1 vocabulary (111 words, 5 days)
│   ├── A2_vocab.xlsx       # A2 template (empty)
│   ├── B1_vocab.xlsx       # B1 template (empty)
│   └── B2_vocab.xlsx       # B2 template (empty)
└── lib/
    ├── __init__.py
    ├── data_generator.py   # Excel file generator
    ├── excel_manager.py    # Read/write Excel operations
    └── image_fetcher.py    # DuckDuckGo image search

Additional: A lightweight Flask clone is available at `clone_app/` that provides
an HTML/Tailwind mobile-friendly UI and JSON APIs reusing the same `lib/`
helpers. See `clone_app/README.md` for quick start instructions.
```

## Excel Column Structure

| Column | Description |
|--------|-------------|
| german_word | German word or phrase |
| pronunciation | Phonetic pronunciation guide |
| meaning | English translation |
| example_sentence | Example usage in German |
| option_1–4 | MCQ options (one is the correct answer) |
| image_url | Auto-fetched image URL (cached) |
| gender | Grammatical category |
| emoji | Fallback emoji |
| keyword | Image search keyword |
| note | Word class information |

## Adding New Vocabulary

1. Open the appropriate Excel file in `data/`
2. Add words to existing day sheets or create new sheets
3. Leave `image_url` empty — the app auto-fetches images on first view
# PROJECT-VOCAB_DAILY
