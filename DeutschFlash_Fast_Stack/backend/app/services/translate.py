from __future__ import annotations

import re
from functools import lru_cache

try:
    from dictcc import Dict
except Exception:  # pragma: no cover
    Dict = None


STOP_WORDS = {
    "der", "die", "das", "den", "dem", "des", "ein", "eine", "einen", "einem", "einer",
    "und", "oder", "aber", "weil", "ist", "bin", "bist", "sind", "seid", "war", "hat",
    "es", "er", "sie", "wir", "ihr", "ich", "du", "in", "im", "an", "am", "auf", "aus",
    "bei", "mit", "von", "zu", "zum", "zur", "für", "um", "nicht", "kein", "keine",
    "so", "sehr", "auch", "noch", "schon", "mein", "meine", "dein", "deine", "sich",
}


def _clean(text: str) -> str:
    text = re.sub(r"\{[^}]*\}", "", text)
    text = re.sub(r"\[[^\]]*\]", "", text)
    text = re.sub(r"\([^)]*\)", "", text)
    text = re.sub(r"^(sb\.\s*|sth\.\s*)+", "", text)
    return text.strip().strip(",").strip()


@lru_cache(maxsize=2048)
def translate_word(word: str) -> str:
    if Dict is None or len(word.strip()) < 3:
        return ""
    try:
        result = Dict().translate(word.strip(), from_language="de", to_language="en")
        for _, english in (result.translation_tuples or [])[:5]:
            clean = _clean(english)
            if clean and len(clean) < 40:
                return clean
    except Exception:
        return ""
    return ""


def sentence_words(sentence: str, exclude_word: str = "") -> dict[str, str]:
    exclude = {exclude_word.lower()}
    bare = re.sub(r"^(der|die|das|ein|eine|einen)\s+", "", exclude_word, flags=re.IGNORECASE).strip()
    if bare:
        exclude.add(bare.lower())
    result: dict[str, str] = {}
    for raw in sentence.split():
        token = raw.strip(".,!?;:\"'()[]–—-…")
        lower = token.lower()
        if len(token) < 3 or lower in STOP_WORDS or lower in exclude or token in result:
            continue
        meaning = translate_word(token)
        if meaning:
            result[token] = meaning
    return result
