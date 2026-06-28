from __future__ import annotations

from pathlib import Path

from fastapi import FastAPI, HTTPException, Query
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles

from app.services.audio import get_audio_file
from app.services.images import image_url
from app.services.translate import sentence_words
from app.services.vocab import DATA_DIR, deck, level_summary, search, test_deck


ROOT = Path(__file__).resolve().parents[2]
SLIDES_DIR = DATA_DIR / "slides_cache"

app = FastAPI(title="DeutschFlash Fast API")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)

if SLIDES_DIR.exists():
    app.mount("/slides", StaticFiles(directory=SLIDES_DIR), name="slides")


@app.get("/api/health")
def health() -> dict[str, str]:
    return {"status": "ok"}


@app.get("/api/levels")
def get_levels() -> dict:
    return {"levels": level_summary()}


@app.get("/api/deck")
def get_deck(level: str, sheet: str = "__all__", shuffle: bool = True) -> dict:
    return {"cards": deck(level, sheet=sheet, shuffle=shuffle)}


@app.get("/api/search")
def get_search(q: str = Query("", min_length=0), limit: int = 25) -> dict:
    return {"results": search(q, limit=limit)}


@app.get("/api/test")
def get_test(level: str, sheet: str = "__all__", count: int = 25) -> dict:
    return {"cards": test_deck(level, sheet=sheet, count=count)}


@app.get("/api/audio")
def get_audio(
    text: str,
    mode: str = "pronounce",
    level: str | None = None,
    sheet: str | None = None,
    row_idx: int | None = None,
):
    if not text.strip():
        raise HTTPException(status_code=400, detail="text is required")
    path = get_audio_file(text, mode=mode, level=level, sheet=sheet, row_idx=row_idx)
    return FileResponse(path, media_type="audio/mpeg")


@app.get("/api/image")
def get_image(keyword: str = "") -> dict[str, str]:
    return {"url": image_url(keyword)}


@app.get("/api/translations")
def get_translations(sentence: str, exclude: str = "") -> dict:
    return {"words": sentence_words(sentence, exclude_word=exclude)}


@app.get("/api/slides")
def get_slides() -> dict:
    slides = []
    if SLIDES_DIR.exists():
        slides = [f"/slides/{path.name}" for path in sorted(SLIDES_DIR.glob("slide-*.png"))]
    return {"slides": slides}
