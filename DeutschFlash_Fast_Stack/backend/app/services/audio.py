from __future__ import annotations

import asyncio
import re
from pathlib import Path

import edge_tts


ROOT = Path(__file__).resolve().parents[3]
AUDIO_DIR = ROOT / "data" / "audio_cache"
VOICE = "de-DE-ConradNeural"


def sanitize_filename(word: str) -> str:
    text = word.strip().replace(" ", "_").replace("/", "_")
    text = re.sub(r"[^a-zA-Z0-9_\-äöüÄÖÜß]", "", text)
    return re.sub(r"_{2,}", "_", text)


def structured_path(level: str, sheet: str, row_idx: int, word: str, mode: str) -> Path:
    row_num = row_idx + 2
    return AUDIO_DIR / (level or "UnknownLevel") / (sheet or "UnknownDay") / mode / f"row_{row_num}_{sanitize_filename(word)}.mp3"


def sentence_path(text: str, mode: str) -> Path:
    import hashlib

    slug = hashlib.md5(f"{mode}:{text}".encode()).hexdigest()
    return AUDIO_DIR / f"{slug}.mp3"


async def _generate(text: str, rate: str = "-10%") -> bytes:
    communicate = edge_tts.Communicate(text, VOICE, rate=rate)
    chunks = bytearray()
    async for chunk in communicate.stream():
        if chunk["type"] == "audio":
            chunks.extend(chunk["data"])
    return bytes(chunks)


async def _generate_spell(word: str) -> bytes:
    raw = re.sub(r"^(der|die|das)\s+", "", word, flags=re.IGNORECASE).strip()
    raw = raw.split("/")[0].strip()
    letters = [f"{char.upper()}..." for char in raw if char not in (" ", "-", "?", "!", ",", "'", ".", "(", ")")]
    return await _generate(f"{word}. ...... Buchstabiert: {', '.join(letters)}", rate="-10%")


def get_audio_file(
    text: str,
    mode: str = "pronounce",
    level: str | None = None,
    sheet: str | None = None,
    row_idx: int | None = None,
) -> Path:
    mode = mode if mode in {"pronounce", "slow", "spell"} else "pronounce"
    if level and sheet and row_idx is not None:
        path = structured_path(level, sheet, row_idx, text, mode)
    else:
        path = sentence_path(text, mode)
    if path.exists() and path.stat().st_size > 0:
        return path

    path.parent.mkdir(parents=True, exist_ok=True)
    if mode == "spell":
        audio = asyncio.run(_generate_spell(text))
    elif mode == "slow":
        audio = asyncio.run(_generate(text, rate="-35%"))
    else:
        audio = asyncio.run(_generate(text, rate="-10%"))
    path.write_bytes(audio)
    return path
