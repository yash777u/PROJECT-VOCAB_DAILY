from __future__ import annotations

from functools import lru_cache

from ddgs import DDGS


@lru_cache(maxsize=512)
def image_url(keyword: str) -> str:
    keyword = keyword.strip()
    if not keyword:
        return ""
    try:
        with DDGS() as ddgs:
            results = ddgs.images(keyword, max_results=5)
        for result in results or []:
            url = result.get("thumbnail") or result.get("image") or ""
            if url:
                return url
    except Exception:
        return ""
    return ""
