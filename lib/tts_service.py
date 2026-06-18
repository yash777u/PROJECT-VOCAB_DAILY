"""
TTS service for DeutschFlash — edge-tts for native German pronunciation.

Modes:
  1. Pronounce: Full native pronunciation
  2. Spell: Says word then spells each letter

Audio is saved to data/audio_cache/<hash>.mp3 on first generation.
Subsequent calls (even after app restart) read from disk — no edge-tts call needed.
"""

import asyncio
import io
import base64
import hashlib
import os
import re

import edge_tts

VOICE = "de-DE-ConradNeural"

_HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
AUDIO_CACHE_DIR = os.path.join(_HERE, "data", "audio_cache")
os.makedirs(AUDIO_CACHE_DIR, exist_ok=True)


def sanitize_filename(word: str) -> str:
    """Sanitize the word to be filesystem-friendly."""
    import re
    s = word.strip()
    s = s.replace(" ", "_").replace("/", "_")
    # Retain German letters, alphanumeric characters, hyphens, and underscores
    s = re.sub(r'[^a-zA-Z0-9_\-äöüÄÖÜß]', '', s)
    s = re.sub(r'_{2,}', '_', s)  # collapse multiple underscores
    return s


def get_audio_paths(level: str, day: str, row_idx: int, word: str) -> dict:
    """Get the expected file paths for pronunciation (normal & slow) and spelling audio."""
    row_num = row_idx + 2
    sanitized = sanitize_filename(word)
    day_folder = day if day else "UnknownDay"
    level_folder = level if level else "UnknownLevel"
    
    # Target folder: audio_cache/<Level>/<Day>/
    base_folder = os.path.join(AUDIO_CACHE_DIR, level_folder, day_folder)
    
    pronounce_path = os.path.join(base_folder, "pronounce", f"row_{row_num}_{sanitized}.mp3")
    slow_path = os.path.join(base_folder, "slow", f"row_{row_num}_{sanitized}.mp3")
    spell_path = os.path.join(base_folder, "spell", f"row_{row_num}_{sanitized}.mp3")
    
    return {
        "pronounce": pronounce_path,
        "slow": slow_path,
        "spell": spell_path
    }


def _audio_cache_path(text: str, mode: str) -> str:
    slug = hashlib.md5(f"{mode}:{text}".encode()).hexdigest()
    return os.path.join(AUDIO_CACHE_DIR, f"{slug}.mp3")


def _run_async(coro):
    try:
        loop = asyncio.get_event_loop()
        if loop.is_running():
            import nest_asyncio
            nest_asyncio.apply()
            return loop.run_until_complete(coro)
        else:
            return loop.run_until_complete(coro)
    except RuntimeError:
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        try:
            return loop.run_until_complete(coro)
        finally:
            loop.close()


async def _generate_audio(text: str, rate: str = "+0%", pitch: str = "+0Hz",
                           volume: str = "+0%") -> bytes:
    communicate = edge_tts.Communicate(text, VOICE, rate=rate, pitch=pitch, volume=volume)
    buffer = io.BytesIO()
    async for chunk in communicate.stream():
        if chunk["type"] == "audio":
            buffer.write(chunk["data"])
    return buffer.getvalue()


async def _generate_spell_audio(word: str) -> bytes:
    raw = re.sub(r'^(der|die|das)\s+', '', word, flags=re.IGNORECASE).strip()
    raw = raw.split('/')[0].strip()

    spelled_letters = []
    for char in raw:
        if char in (' ', '-', '?', '!', ',', "'", '.', '(', ')'):
            continue
        spelled_letters.append(f"{char.upper()}...")

    spelled_phrase = ", ".join(spelled_letters)
    combined_text = f"{word}. ...... Buchstabiert: {spelled_phrase}"
    return await _generate_audio(combined_text, rate="-10%")


def _bytes_to_data_uri(audio_bytes: bytes) -> str:
    b64 = base64.b64encode(audio_bytes).decode("utf-8")
    return f"data:audio/mp3;base64,{b64}"


def get_audio_uri(word: str, mode: str, level: str, day: str, row_idx: int) -> str:
    """Get base64 MP3 data URI. Load from structured disk path if exists, otherwise generate."""
    paths = get_audio_paths(level, day, row_idx, word)
    if mode == "spell":
        target_path = paths["spell"]
    elif mode == "slow":
        target_path = paths["slow"]
    else:
        target_path = paths["pronounce"]
    
    # If exists on disk, read and return
    if os.path.exists(target_path) and os.path.getsize(target_path) > 0:
        try:
            with open(target_path, "rb") as f:
                return _bytes_to_data_uri(f.read())
        except Exception:
            pass # fallback to regenerate if corrupted
            
    # Otherwise generate
    try:
        if mode == "spell":
            audio_bytes = _run_async(_generate_spell_audio(word))
        elif mode == "slow":
            audio_bytes = _run_async(_generate_audio(word.strip(), rate="-35%"))
        else:
            audio_bytes = _run_async(_generate_audio(word.strip(), rate="-10%"))
            
        if audio_bytes:
            # Save to disk
            os.makedirs(os.path.dirname(target_path), exist_ok=True)
            with open(target_path, "wb") as f:
                f.write(audio_bytes)
            return _bytes_to_data_uri(audio_bytes)
    except Exception as e:
        print(f"Error in get_audio_uri: {e}")
        
    return ""


def _load_or_generate(word: str, mode: str, generate_fn) -> str:
    """Check disk cache first; generate + save if missing."""
    cache_path = _audio_cache_path(word, mode)

    # Disk hit
    if os.path.exists(cache_path) and os.path.getsize(cache_path) > 0:
        try:
            with open(cache_path, "rb") as f:
                return _bytes_to_data_uri(f.read())
        except Exception:
            pass  # Corrupted — fall through to regenerate

    # Generate
    try:
        audio_bytes = _run_async(generate_fn(word))
        if not audio_bytes:
            return ""
        # Save to disk
        try:
            with open(cache_path, "wb") as f:
                f.write(audio_bytes)
        except Exception as e:
            print(f"TTS disk-cache write error: {e}")
        return _bytes_to_data_uri(audio_bytes)
    except Exception as e:
        print(f"TTS generate error ({mode}): {e}")
        return ""


def pronounce_word(word: str, level: str = None, day: str = None, row_idx: int = None) -> str:
    """Return base64 MP3 data URI for native pronunciation. Cached to disk."""
    if not word or not word.strip():
        return ""

    if level and day and row_idx is not None:
        return get_audio_uri(word.strip(), "pronounce", level, day, row_idx)

    async def _gen(w):
        return await _generate_audio(w.strip(), rate="-10%")

    return _load_or_generate(word.strip(), "pronounce", _gen)


def slow_word(word: str, level: str = None, day: str = None, row_idx: int = None) -> str:
    """Return base64 MP3 data URI for slow native pronunciation. Cached to disk."""
    if not word or not word.strip():
        return ""

    if level and day and row_idx is not None:
        return get_audio_uri(word.strip(), "slow", level, day, row_idx)

    async def _gen(w):
        return await _generate_audio(w.strip(), rate="-35%")

    return _load_or_generate(word.strip(), "slow", _gen)


def spell_word(word: str, level: str = None, day: str = None, row_idx: int = None) -> str:
    """Return base64 MP3 data URI for spell-out audio. Cached to disk."""
    if not word or not word.strip():
        return ""

    if level and day and row_idx is not None:
        return get_audio_uri(word.strip(), "spell", level, day, row_idx)

    async def _gen(w):
        return await _generate_spell_audio(w.strip())

    return _load_or_generate(word.strip(), "spell", _gen)