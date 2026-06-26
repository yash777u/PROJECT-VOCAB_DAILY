"""
Image fetcher for DeutschFlash — DuckDuckGo image search.
Searches by keyword column. Falls back to german_word.

NOTE: image cache / disk persistence has been removed. This module will
return image URLs from DuckDuckGo results and will not download or save
images to data/image_cache/. The Streamlit UI can display remote URLs
directly via `st.image(url)`.
"""

import os
import hashlib
import requests
try:
    import streamlit as st
except Exception:
    # Provide a minimal no-op replacement for st.cache_data used by this module
    class _NoopCache:
        def __call__(self, *args, **kwargs):
            def _decorator(f):
                return f
            return _decorator

    st = type('st', (), {'cache_data': _NoopCache()})
from io import BytesIO
from ddgs import DDGS
from ddgs.exceptions import RatelimitException, TimeoutException, DDGSException

_HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
IMAGE_CACHE_DIR = os.path.join(_HERE, "data", "image_cache")
# directory left for backwards-compatibility but not written to by this module
os.makedirs(IMAGE_CACHE_DIR, exist_ok=True)

_MAX_PX = 480
_JPEG_Q = 78


def _cache_path(key: str) -> str:
    # kept for API compatibility; not used for downloads anymore
    slug = hashlib.md5(key.lower().strip().encode()).hexdigest()
    return os.path.join(IMAGE_CACHE_DIR, f"{slug}.jpg")


@st.cache_data(show_spinner=False, ttl=3600)
def _ddg_search(keywords: str, max_results: int = 5):
    with DDGS() as ddgs:
        return ddgs.images(keywords, max_results=max_results)


def _url_cache_path(url: str) -> str:
    # kept for compatibility; downloads are disabled so this is unused
    slug = hashlib.md5(url.encode()).hexdigest()
    return os.path.join(IMAGE_CACHE_DIR, f"url_{slug}.jpg")


def fetch_image_for_row(row: dict, session_results: dict = None, max_retries: int = 3) -> tuple[str, dict]:
    """Return (image_url_or_empty, debug_info) for the image matching this vocab row.

    Only uses the 'keyword' column for searching — never falls back to german_word.
    Retries the DDG search up to `max_retries` times on failure before giving up.
    """
    keyword = str(row.get("keyword", "")).strip()
    german = str(row.get("german_word", "")).strip()
    search_term = keyword  # Only use keyword column, no fallback

    info = {
        "search_term": search_term,
        "keyword_col": keyword,
        "german_word": german,
        "source": "none",
        "chosen_url": "",
    }

    if not search_term:
        info["error"] = "keyword column is empty — no image search performed"
        return "", info

    # Attempt to reuse session cache if available
    results = None
    if session_results is not None and search_term in session_results:
        results = session_results[search_term]

    if not results:
        last_error = None
        for attempt in range(1, max_retries + 1):
            try:
                results = _ddg_search(search_term, max_results=5)
                if results:
                    break
            except (RatelimitException, TimeoutException, DDGSException, Exception) as e:
                last_error = e
                import time as _time
                _time.sleep(1)  # brief pause before retry
                continue

        if session_results is not None and results:
            session_results[search_term] = results

        if not results:
            info["error"] = f"DDG returned no results after {max_retries} attempts"
            if last_error:
                info["error"] += f" (last error: {last_error})"
            return "", info

    # Collect valid URLs from results
    valid_urls = []
    for r in results:
        url = r.get("thumbnail") or r.get("image", "")
        if url:
            valid_urls.append(url)

    if not valid_urls:
        info["error"] = "No valid image URLs in results"
        return "", info

    # Pick a random URL and return it (do not download)
    import random
    chosen_url = random.choice(valid_urls)
    info["chosen_url"] = chosen_url
    info["all_urls"] = valid_urls
    info["source"] = "ddg_url"

    return chosen_url, info


def fetch_image_url(keyword: str, max_results: int = 3) -> str:
    if not keyword or not keyword.strip():
        return ""
    try:
        results = _ddg_search(keyword.strip(), max_results=max_results)
        if not results:
            return ""
        for img_data in results:
            img_url = img_data.get("thumbnail") or img_data.get("image", "")
            if img_url:
                return img_url
        return ""
    except Exception:
        return ""


def validate_image_url(url: str, timeout: int = 5) -> bool:
    if not url or not url.strip():
        return False
    try:
        resp = requests.head(url, timeout=timeout, allow_redirects=True)
        if resp.status_code == 200:
            ct = resp.headers.get("content-type", "").lower()
            if "image" in ct:
                return True
            if any(url.lower().endswith(ext) for ext in [".jpg", ".jpeg", ".png", ".webp", ".gif"]):
                return True
        if resp.status_code in (301, 302):
            return True
        return False
    except Exception:
        return False