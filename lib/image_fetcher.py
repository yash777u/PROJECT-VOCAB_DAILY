"""
Image fetcher for DeutschFlash — DuckDuckGo image search.
Searches by keyword column. Falls back to german_word.
Caches downscaled images to data/image_cache/.
"""

import os
import hashlib
import requests
import streamlit as st
from io import BytesIO
from ddgs import DDGS
from ddgs.exceptions import RatelimitException, TimeoutException, DDGSException

_HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
IMAGE_CACHE_DIR = os.path.join(_HERE, "data", "image_cache")
os.makedirs(IMAGE_CACHE_DIR, exist_ok=True)

_MAX_PX = 480
_JPEG_Q = 78


def _cache_path(key: str) -> str:
    slug = hashlib.md5(key.lower().strip().encode()).hexdigest()
    return os.path.join(IMAGE_CACHE_DIR, f"{slug}.jpg")


def _downscale_and_save(image_bytes: bytes, dest_path: str) -> bool:
    try:
        from PIL import Image
        img = Image.open(BytesIO(image_bytes)).convert("RGB")
        w, h = img.size
        if max(w, h) > _MAX_PX:
            scale = _MAX_PX / max(w, h)
            img = img.resize((max(1, int(w * scale)), max(1, int(h * scale))), Image.LANCZOS)
        img.save(dest_path, "JPEG", quality=_JPEG_Q, optimize=True)
        return True
    except Exception:
        return False


@st.cache_data(show_spinner=False, ttl=3600)
def _ddg_search(keywords: str, max_results: int = 5):
    with DDGS() as ddgs:
        return ddgs.images(keywords, max_results=max_results)


def _url_cache_path(url: str) -> str:
    slug = hashlib.md5(url.encode()).hexdigest()
    return os.path.join(IMAGE_CACHE_DIR, f"url_{slug}.jpg")


def fetch_image_for_row(row: dict, session_results: dict = None) -> tuple[str, dict]:
    """Return (local_file_path, debug_info) for the image matching this vocab row.

    Pick a random URL from the top 5 image results from DDG.
    Cache downloaded files on disk by URL hash to avoid redundant downloads.
    Use session_results cache to prevent redundant DDG queries in the same session.
    """
    keyword = str(row.get("keyword", "")).strip()
    german  = str(row.get("german_word", "")).strip()
    search_term = keyword if keyword else german

    info = {
        "search_term": search_term,
        "keyword_col": keyword,
        "german_word": german,
        "source": "failed",
        "chosen_url": "",
        "cache_path": "",
    }

    if not search_term:
        return "", info

    # Fetch/Search results
    results = None
    if session_results is not None and search_term in session_results:
        results = session_results[search_term]

    if not results:
        try:
            results = _ddg_search(search_term, max_results=5)
            if session_results is not None:
                session_results[search_term] = results
        except (RatelimitException, TimeoutException, DDGSException, Exception) as e:
            info["error"] = f"DDG search error: {e}"
            return "", info

    if not results:
        info["error"] = "DDG returned no results"
        return "", info

    # Filter valid URLs
    valid_urls = []
    for r in results:
        url = r.get("thumbnail") or r.get("image", "")
        if url:
            valid_urls.append(url)

    if not valid_urls:
        info["error"] = "No valid image URLs in results"
        return "", info

    # Pick a random URL to satisfy 'let it generate each time new result'
    import random
    chosen_url = random.choice(valid_urls)
    info["chosen_url"] = chosen_url
    info["all_urls"] = valid_urls

    dest = _url_cache_path(chosen_url)
    info["cache_path"] = dest

    # Disk cache hit for this specific URL
    if os.path.exists(dest) and os.path.getsize(dest) > 0:
        info["source"] = "disk_cache"
        return dest, info

    # Download
    try:
        resp = requests.get(chosen_url, timeout=6, headers={"User-Agent": "Mozilla/5.0"})
        if resp.status_code == 200 and resp.content:
            if _downscale_and_save(resp.content, dest):
                info["source"] = "ddg_download"
                return dest, info
        info["error"] = f"Download failed: HTTP {resp.status_code}"
    except Exception as e:
        info["error"] = f"Download exception: {e}"

    return "", info


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