"""
Translation service for DeutschFlash — dict.cc integration.

Provides German→English word translations for hover tooltips in the
Random Test section. Uses the ``dict.cc.py`` library (web-scraping based)
with Streamlit caching to minimise network calls.

Usage:
    from lib.translation_service import translate_word, translate_sentence_words

    meaning = translate_word("Vater")           # → "father"
    word_map = translate_sentence_words(
        "Mein Vater arbeitet viel.",
        exclude_word="Vater",
    )
    # → {"Mein": "My", "arbeitet": "works", "viel": "much"}
"""

import re
from concurrent.futures import ThreadPoolExecutor, as_completed

try:
    import streamlit as st
    _has_st = True
except ImportError:
    _has_st = False

try:
    from dictcc import Dict
    _has_dictcc = True
except ImportError:
    _has_dictcc = False

# ── German stop-words that are too common / short to look up ──
# These are skipped during batch translation to avoid unnecessary requests.
STOP_WORDS = frozenset({
    "der", "die", "das", "den", "dem", "des",
    "ein", "eine", "einen", "einem", "einer", "eines",
    "und", "oder", "aber", "doch", "denn", "weil",
    "ist", "bin", "bist", "sind", "seid", "war", "hat",
    "es", "er", "sie", "wir", "ihr", "ich", "du",
    "in", "im", "an", "am", "auf", "aus", "bei", "mit",
    "von", "zu", "zum", "zur", "für", "um", "über",
    "nicht", "kein", "keine", "keinen", "keinem",
    "so", "sehr", "auch", "noch", "schon", "ja", "nein",
    "mein", "meine", "meinen", "dein", "deine", "sein", "seine",
    "sich", "mir", "dir", "uns", "euch",
})

# Minimum word length worth translating (skip 1-2 char tokens).
_MIN_WORD_LEN = 3


def _extract_first_english(translation_tuples: list) -> str:
    """Pick the cleanest, shortest English meaning from dict.cc results.

    Dict.cc returns tuples like ``("Vater {m}", "father")``.
    We strip grammar annotations and return just the plain English word(s).

    Returns:
        A short English string, or ``""`` if nothing usable.
    """
    for _de, en in translation_tuples[:5]:
        # Remove grammar tags like {m}, {f}, {adj}, [attr.], (past-p), etc.
        clean = re.sub(r'\{[^}]*\}', '', en)
        clean = re.sub(r'\[[^\]]*\]', '', clean)
        clean = re.sub(r'\([^)]*\)', '', clean)
        # Remove leading "sb./sth." prefixes
        clean = re.sub(r'^(sb\.\s*|sth\.\s*)+', '', clean)
        clean = clean.strip().strip(',').strip()
        if clean and len(clean) < 40:
            return clean
    return ""


# ── Cached single-word translation ──
if _has_st:
    @st.cache_data(show_spinner=False, ttl=3600)
    def _cached_translate(word: str) -> str:
        """Look up a German word on dict.cc and return a short English meaning.

        Cached for 1 hour via Streamlit's ``@cache_data``.
        Returns ``""`` on any failure (rate-limit, timeout, no results).
        """
        if not _has_dictcc:
            return ""
        try:
            d = Dict()
            result = d.translate(word, from_language="de", to_language="en")
            if result and result.translation_tuples:
                return _extract_first_english(result.translation_tuples)
        except Exception:
            pass
        return ""
else:
    def _cached_translate(word: str) -> str:  # noqa: F811 – fallback when no Streamlit
        if not _has_dictcc:
            return ""
        try:
            d = Dict()
            result = d.translate(word, from_language="de", to_language="en")
            if result and result.translation_tuples:
                return _extract_first_english(result.translation_tuples)
        except Exception:
            pass
        return ""


def translate_word(word: str) -> str:
    """Translate a single German word to English via dict.cc.

    Args:
        word: A German word (e.g. ``"Vater"``).

    Returns:
        A short English meaning (e.g. ``"father"``), or ``""`` if not found.
    """
    if not word or len(word.strip()) < _MIN_WORD_LEN:
        return ""
    return _cached_translate(word.strip())


if _has_st:
    @st.cache_data(show_spinner=False, ttl=86400)
    def translate_sentence_words(sentence: str, exclude_word: str = "") -> dict:
        """Translate each significant word in a German sentence — parallelised + cached."""
        if not sentence:
            return {}
        exclude_tokens: set[str] = set()
        if exclude_word:
            exclude_tokens.add(exclude_word.lower())
            bare = re.sub(
                r'^(der|die|das|ein|eine|einen)\s+',
                '', exclude_word, flags=re.IGNORECASE,
            ).strip()
            if bare:
                exclude_tokens.add(bare.lower())

        raw_tokens = sentence.split()
        to_translate = []
        for raw in raw_tokens:
            token = raw.strip(".,!?;:\"'()[]\u2013\u2014-\u2026")
            lower = token.lower()
            if len(token) < _MIN_WORD_LEN:
                continue
            if lower in STOP_WORDS:
                continue
            if lower in exclude_tokens:
                continue
            if token not in to_translate:
                to_translate.append(token)

        if not to_translate:
            return {}

        # Fire all lookups concurrently
        result: dict[str, str] = {}
        with ThreadPoolExecutor(max_workers=6) as ex:
            future_to_word = {ex.submit(translate_word, w): w for w in to_translate}
            for future in as_completed(future_to_word):
                word = future_to_word[future]
                try:
                    meaning = future.result()
                    if meaning:
                        result[word] = meaning
                except Exception:
                    pass
        return result
else:
    def translate_sentence_words(
        sentence: str,
        exclude_word: str = "",
    ) -> dict[str, str]:
        """Translate each significant word in a German sentence (no-Streamlit fallback)."""
        if not sentence:
            return {}
        exclude_tokens: set[str] = set()
        if exclude_word:
            exclude_tokens.add(exclude_word.lower())
            bare = re.sub(
                r'^(der|die|das|ein|eine|einen)\s+',
                '', exclude_word, flags=re.IGNORECASE,
            ).strip()
            if bare:
                exclude_tokens.add(bare.lower())

        raw_tokens = sentence.split()
        result: dict[str, str] = {}
        for raw in raw_tokens:
            token = raw.strip(".,!?;:\"'()[]\u2013\u2014-\u2026")
            lower = token.lower()
            if len(token) < _MIN_WORD_LEN:
                continue
            if lower in STOP_WORDS:
                continue
            if lower in exclude_tokens:
                continue
            if token in result:
                continue
            meaning = translate_word(token)
            if meaning:
                result[token] = meaning
        return result
