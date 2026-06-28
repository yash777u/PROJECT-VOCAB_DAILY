"""
DeutschFlash — Streamlit German Vocabulary Learning App.

A mobile-first, glassmorphism-styled flashcard app for learning German.
Features three tabs:
  • Practice — Day-by-day or marathon flashcard quiz with TTS & images
  • Search   — Full-text search across every level and sheet
  • Random Test — Timed sentence-context quiz with dict.cc hover tooltips

Dependencies: streamlit, pandas, openpyxl, edge-tts, ddgs, dict.cc.py
"""
import random
import os
import time
import base64
import threading
import streamlit as st
import pandas as pd
from lib.excel_manager import get_available_levels, get_day_info, load_vocab, load_all_vocab, is_level_empty
from lib.image_fetcher import fetch_image_for_row
from lib.tts_service import pronounce_word, slow_word, spell_word, get_audio_paths
from lib.translation_service import translate_sentence_words

# ── Page Config ──
st.set_page_config(page_title="DeutschFlash Master", page_icon="🇩🇪", layout="centered")

# ── Custom CSS ──
st.markdown("""<style>

html, body, [class*="css"] {
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;
    color: #f8fafc;
}

.stApp {
    background: radial-gradient(circle at top right, #1e1b4b, #0f172a) !important;
}

/* Hide Streamlit elements */
#MainMenu, footer, header { visibility: hidden; }
div[data-testid="stToolbar"] { display: none; }

/* Custom Container Width & Padding for responsiveness */
.block-container {
    max-width: 600px !important;
    padding: 1.5rem 1rem !important;
}

/* Glassmorphism Card styling */
.glass-card {
    background: rgba(15, 23, 42, 0.65);
    border: 1px solid rgba(255, 255, 255, 0.08);
    border-radius: 24px;
    padding: 1.75rem;
    backdrop-filter: blur(20px);
    -webkit-backdrop-filter: blur(20px);
    box-shadow: 0 12px 40px 0 rgba(0, 0, 0, 0.3);
    margin-bottom: 1rem;
    animation: fadeIn 0.4s ease-out;
}

/* Day selector cards */
.day-card {
    background: rgba(30, 41, 59, 0.45);
    border: 1px solid rgba(255, 255, 255, 0.05);
    border-radius: 18px;
    padding: 1.25rem 1.5rem;
    margin: 0.5rem 0;
    transition: all 0.25s cubic-bezier(0.4, 0, 0.2, 1);
    cursor: pointer;
}
.day-card:hover {
    border-color: rgba(99, 102, 241, 0.4);
    background: rgba(99, 102, 241, 0.08);
    transform: translateY(-2px);
    box-shadow: 0 8px 24px rgba(99, 102, 241, 0.15);
}

/* Badges */
.badge {
    display: inline-block;
    font-size: 0.725rem;
    font-weight: 700;
    padding: 0.3rem 0.85rem;
    border-radius: 999px;
    letter-spacing: 0.05em;
    text-transform: uppercase;
}
.badge-m { background: rgba(59, 130, 246, 0.12); color: #3b82f6; border: 1px solid rgba(59, 130, 246, 0.25); }
.badge-f { background: rgba(236, 72, 153, 0.12); color: #ec4899; border: 1px solid rgba(236, 72, 153, 0.25); }
.badge-n { background: rgba(16, 185, 129, 0.12); color: #10b981; border: 1px solid rgba(16, 185, 129, 0.25); }
.badge-p { background: rgba(139, 92, 246, 0.12); color: #8b5cf6; border: 1px solid rgba(139, 92, 246, 0.25); }
.badge-v { background: rgba(245, 158, 11, 0.12); color: #f59e0b; border: 1px solid rgba(245, 158, 11, 0.25); }
.badge-t { background: rgba(6, 182, 212, 0.12); color: #06b6d4; border: 1px solid rgba(6, 182, 212, 0.25); }

/* Typography */
.german-word {
    font-size: 2.25rem;
    font-weight: 800;
    color: #ffffff;
    text-align: center;
    margin: 0.5rem 0;
    letter-spacing: -0.02em;
    word-break: break-word;
}
.phonetic {
    font-size: 0.95rem;
    color: #22d3ee;
    background: rgba(6, 182, 212, 0.08);
    padding: 0.35rem 0.85rem;
    border-radius: 8px;
    border: 1px solid rgba(6, 182, 212, 0.15);
    display: inline-block;
    font-weight: 500;
}
.note-text {
    font-size: 0.775rem;
    color: #64748b;
    text-transform: uppercase;
    letter-spacing: 0.1em;
}

/* MCQ Options & Standard buttons styling */
.opt-btn, .stButton > button {
    display: inline-flex !important;
    align-items: center !important;
    justify-content: center !important;
    text-align: center !important;
    background: rgba(30, 41, 59, 0.45) !important;
    border: 1px solid rgba(255, 255, 255, 0.08) !important;
    border-radius: 12px !important;
    padding: 0.75rem 1rem !important;
    color: #e2e8f0 !important;
    font-weight: 500 !important;
    font-size: 0.925rem !important;
    min-height: 46px !important;
    width: 100% !important;
    transition: all 0.2s ease !important;
    box-sizing: border-box !important;
    box-shadow: 0 4px 12px rgba(0, 0, 0, 0.15) !important;
    margin-bottom: 0.5rem;
}
.opt-btn:hover, .stButton > button:hover {
    background: rgba(99, 102, 241, 0.15) !important;
    border-color: rgba(99, 102, 241, 0.45) !important;
    color: #ffffff !important;
    transform: translateY(-1px);
}
@media (max-width: 600px) {
    .opt-btn, .stButton > button {
        font-size: 0.825rem !important;
        padding: 0.6rem 0.8rem !important;
        min-height: 40px !important;
        border-radius: 10px !important;
    }
}

/* Hide Streamlit Audio Player visually */
div[data-testid="stAudio"] {
    display: none !important;
    height: 0px !important;
    width: 0px !important;
    padding: 0px !important;
    margin: 0px !important;
}

.opt-correct {
    background: rgba(16, 185, 129, 0.15) !important;
    border-color: rgba(16, 185, 129, 0.5) !important;
    color: #34d399 !important;
    font-weight: 700 !important;
}
.opt-wrong {
    background: rgba(239, 68, 68, 0.15) !important;
    border-color: rgba(239, 68, 68, 0.5) !important;
    color: #f87171 !important;
    font-weight: 600 !important;
}

/* Feedback blocks */
.fb-correct {
    background: rgba(16, 185, 129, 0.06);
    border: 1px solid rgba(16, 185, 129, 0.2);
    border-radius: 16px;
    padding: 1.25rem;
    color: #34d399;
    margin-top: 1rem;
    animation: slideUp 0.3s ease-out;
}
.fb-wrong {
    background: rgba(239, 68, 68, 0.06);
    border: 1px solid rgba(239, 68, 68, 0.2);
    border-radius: 16px;
    padding: 1.25rem;
    color: #f87171;
    margin-top: 1rem;
    animation: slideUp 0.3s ease-out;
}

/* Streak box */
.streak-box {
    background: rgba(245, 158, 11, 0.1);
    border: 1px solid rgba(245, 158, 11, 0.25);
    border-radius: 999px;
    padding: 0.35rem 0.95rem;
    font-size: 0.875rem;
    font-weight: 700;
    color: #fbbf24;
    display: inline-flex;
    align-items: center;
    gap: 0.3rem;
    box-shadow: 0 4px 12px rgba(245, 158, 11, 0.1);
}

/* Progress bar */
.stProgress > div > div {
    background: linear-gradient(90deg, #6366f1, #8b5cf6, #06b6d4) !important;
    border-radius: 999px;
    height: 8px !important;
}

/* Stat Card */
.stat-card {
    background: rgba(30, 41, 59, 0.4);
    border: 1px solid rgba(255, 255, 255, 0.05);
    border-radius: 18px;
    padding: 1.25rem;
    text-align: center;
}
.stat-label {
    font-size: 0.725rem;
    color: #64748b;
    text-transform: uppercase;
    letter-spacing: 0.08em;
    font-weight: 700;
}
.stat-value {
    font-size: 1.5rem;
    font-weight: 800;
    margin-top: 0.25rem;
}

/* Primary/Next button styling override */
div[data-testid="stBaseButton-primary"] > button {
    background: linear-gradient(90deg, #4f46e5, #7c3aed) !important;
    border: none !important;
    color: #ffffff !important;
    box-shadow: 0 4px 14px rgba(79, 70, 229, 0.3) !important;
    justify-content: center !important;
    text-align: center !important;
}
div[data-testid="stBaseButton-primary"] > button:hover {
    background: linear-gradient(90deg, #4338ca, #6d28d9) !important;
    box-shadow: 0 6px 20px rgba(79, 70, 229, 0.4) !important;
}

/* Keep top-bar navigation inline on mobile */
div[data-testid="element-container"]:has(.top-nav-container) + div[data-testid="element-container"] > div[data-testid="stHorizontalBlock"] {
    display: flex !important;
    flex-direction: row !important;
    flex-wrap: nowrap !important;
    align-items: center !important;
    justify-content: space-between !important;
    gap: 0.5rem !important;
}
div[data-testid="element-container"]:has(.top-nav-container) + div[data-testid="element-container"] > div[data-testid="stHorizontalBlock"] > div[data-testid="column"] {
    flex: 1 1 50% !important;
    width: 50% !important;
    max-width: 50% !important;
    min-width: 0px !important;
}

/* Keep streak line inline on mobile */
div[data-testid="element-container"]:has(.streak-line-container) + div[data-testid="element-container"] > div[data-testid="stHorizontalBlock"] {
    display: flex !important;
    flex-direction: row !important;
    flex-wrap: nowrap !important;
    align-items: center !important;
    justify-content: space-between !important;
    gap: 0.5rem !important;
}
div[data-testid="element-container"]:has(.streak-line-container) + div[data-testid="element-container"] > div[data-testid="stHorizontalBlock"] > div[data-testid="column"] {
    flex: 1 1 50% !important;
    width: 50% !important;
    max-width: 50% !important;
    min-width: 0px !important;
}

/* Keep TTS buttons inline on mobile and style them for spacing/padding */
div[data-testid="element-container"]:has(.tts-container) + div[data-testid="element-container"] > div[data-testid="stHorizontalBlock"] {
    display: flex !important;
    flex-direction: row !important;
    flex-wrap: nowrap !important;
    align-items: center !important;
    justify-content: space-between !important;
    gap: 0.5rem !important;
}
div[data-testid="element-container"]:has(.tts-container) + div[data-testid="element-container"] > div[data-testid="stHorizontalBlock"] div[data-testid="column"] {
    flex: 1 1 33.33% !important;
    width: 33.33% !important;
    max-width: 33.33% !important;
    min-width: 0px !important;
}
div[data-testid="element-container"]:has(.tts-container) + div[data-testid="element-container"] > div[data-testid="stHorizontalBlock"] button {
    padding: 0.6rem 0.5rem !important;
    font-size: 0.85rem !important;
    min-height: 38px !important;
    border-radius: 8px !important;
    width: 100% !important;
}

/* MCQ Option layout adjustment: wrap to 1 or 2 columns based on text size */
div[data-testid="element-container"]:has(.mcq-container) + div[data-testid="element-container"] > div[data-testid="stHorizontalBlock"] {
    display: flex !important;
    flex-direction: row !important;
    flex-wrap: wrap !important;
    align-items: stretch !important;
    gap: 0.5rem !important;
}
div[data-testid="element-container"]:has(.mcq-container) + div[data-testid="element-container"] > div[data-testid="stHorizontalBlock"] > div[data-testid="column"] {
    flex: 1 1 45% !important;
    min-width: 140px !important;
}

/* Fixed square flashcard image (300x300, cropped to fill) */
.flashcard-img-box {
    width: 300px;
    height: 300px;
    margin: 0.75rem auto;
    border-radius: 16px;
    overflow: hidden;
    border: 1px solid rgba(255, 255, 255, 0.1);
    box-shadow: 0 4px 15px rgba(0, 0, 0, 0.3);
}
.flashcard-img-box img {
    width: 100%;
    height: 100%;
    object-fit: cover;
    display: block;
    pointer-events: none;
    -webkit-user-drag: none;
}
@media (max-width: 600px) {
    .flashcard-img-box {
        width: 80vw;
        height: 80vw;
        max-width: 300px;
        max-height: 300px;
    }
}

/* Search input box */
div[data-testid="stTextInput"] input {
    background: rgba(30, 41, 59, 0.55) !important;
    border: 1px solid rgba(255, 255, 255, 0.1) !important;
    border-radius: 14px !important;
    color: #f8fafc !important;
    padding: 0.85rem 1.1rem !important;
    font-size: 1rem !important;
    box-shadow: 0 4px 14px rgba(0, 0, 0, 0.2) !important;
    transition: border-color 0.2s ease, box-shadow 0.2s ease !important;
}
div[data-testid="stTextInput"] input::placeholder {
    color: #64748b !important;
}
div[data-testid="stTextInput"] input:focus {
    border-color: rgba(99, 102, 241, 0.5) !important;
    box-shadow: 0 0 0 3px rgba(99, 102, 241, 0.15) !important;
}
div[data-testid="stTextInput"] > div {
    background: transparent !important;
    border: none !important;
}

/* Search tab — result card */
.search-card {
    background: rgba(15, 23, 42, 0.65);
    border: 1px solid rgba(255, 255, 255, 0.08);
    border-radius: 20px;
    padding: 1.5rem;
    backdrop-filter: blur(20px);
    -webkit-backdrop-filter: blur(20px);
    box-shadow: 0 12px 32px 0 rgba(0, 0, 0, 0.28);
    margin-bottom: 1.25rem;
    animation: fadeIn 0.35s ease-out;
}
.search-meta-row {
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 0.5rem;
    margin-bottom: 0.5rem;
}
.search-loc {
    font-size: 0.7rem;
    color: #818cf8;
    font-weight: 600;
    letter-spacing: 0.08em;
    text-transform: uppercase;
}
.search-meaning {
    text-align: center;
    font-size: 1.15rem;
    font-weight: 600;
    color: #c7d2fe;
    margin: 0.4rem 0 0.2rem 0;
}
.search-example {
    background: rgba(99, 102, 241, 0.06);
    border: 1px solid rgba(99, 102, 241, 0.15);
    border-radius: 12px;
    padding: 0.85rem 1rem;
    color: #cbd5e1;
    font-size: 0.95rem;
    margin-top: 0.75rem;
}
.search-empty {
    text-align: center;
    color: #64748b;
    padding: 2.5rem 1rem;
}

/* Scrap proofing rules */
.glass-card, .search-card, .german-word, .phonetic, .note-text, img, button, div[data-testid="stMarkdownContainer"] {
    user-select: none !important;
    -webkit-user-select: none !important;
    -moz-user-select: none !important;
    -ms-user-select: none !important;
}

/* Animations */
@keyframes fadeIn { from {opacity:0; transform:translateY(6px)} to {opacity:1; transform:translateY(0)} }
@keyframes slideUp { from {opacity:0; transform:translateY(12px)} to {opacity:1; transform:translateY(0)} }
@keyframes pulse-glow { 0%,100%{box-shadow:0 0 8px rgba(239,68,68,0.3)} 50%{box-shadow:0 0 20px rgba(239,68,68,0.6)} }

/* Random Test — sentence display */
.test-sentence {
    font-size: 1.35rem;
    font-weight: 500;
    color: #e2e8f0;
    text-align: center;
    line-height: 2.2;
    margin: 0.75rem 0;
    word-break: break-word;
}
/* Quiz word — highlighted, NO tooltip (would give away the answer) */
.test-sentence .highlight-word {
    background: linear-gradient(135deg, rgba(99,102,241,0.25), rgba(139,92,246,0.25));
    border: 1px solid rgba(139,92,246,0.45);
    border-radius: 8px;
    padding: 0.2rem 0.55rem;
    color: #a78bfa;
    font-weight: 700;
}
/* Dict.cc tooltip — hover to see English meaning */
.tt-word {
    position: relative;
    cursor: help;
    border-bottom: 1px dashed rgba(255,255,255,0.25);
    transition: color 0.15s ease, border-color 0.15s ease;
}
.tt-word:hover {
    color: #22d3ee;
    border-bottom-color: #22d3ee;
}
.tt-word:hover::after {
    content: attr(data-tip);
    position: absolute;
    bottom: 110%;
    left: 50%;
    transform: translateX(-50%);
    background: rgba(15, 23, 42, 0.95);
    color: #22d3ee;
    font-size: 0.78rem;
    font-weight: 600;
    padding: 0.3rem 0.65rem;
    border-radius: 8px;
    border: 1px solid rgba(34, 211, 238, 0.25);
    white-space: nowrap;
    z-index: 100;
    pointer-events: none;
    box-shadow: 0 4px 12px rgba(0,0,0,0.4);
    animation: fadeIn 0.15s ease-out;
}
/* Mobile: tap to toggle tooltip via checkbox hack */
@media (hover: none) {
    .tt-word:active::after {
        content: attr(data-tip);
        position: absolute;
        bottom: 110%;
        left: 50%;
        transform: translateX(-50%);
        background: rgba(15, 23, 42, 0.95);
        color: #22d3ee;
        font-size: 0.78rem;
        font-weight: 600;
        padding: 0.3rem 0.65rem;
        border-radius: 8px;
        border: 1px solid rgba(34, 211, 238, 0.25);
        white-space: nowrap;
        z-index: 100;
        box-shadow: 0 4px 12px rgba(0,0,0,0.4);
    }
}

/* Timer bar */
.test-timer {
    display: flex;
    align-items: center;
    justify-content: center;
    gap: 0.5rem;
    font-size: 1.1rem;
    font-weight: 700;
    padding: 0.5rem 1rem;
    border-radius: 999px;
    margin: 0.5rem auto;
    width: fit-content;
}
.timer-ok {
    color: #34d399;
    background: rgba(16,185,129,0.08);
    border: 1px solid rgba(16,185,129,0.2);
}
.timer-warn {
    color: #fbbf24;
    background: rgba(245,158,11,0.08);
    border: 1px solid rgba(245,158,11,0.2);
}
.timer-danger {
    color: #f87171;
    background: rgba(239,68,68,0.1);
    border: 1px solid rgba(239,68,68,0.3);
    animation: pulse-glow 1s ease-in-out infinite;
}

/* Responsive adjustments */
@media (max-width: 768px) {
    .block-container { padding: 1rem 0.5rem !important; }
    .glass-card { padding: 1.25rem; border-radius: 20px; }
    .search-card { padding: 1.1rem; border-radius: 18px; }
    .german-word { font-size: 1.85rem; }
    .test-sentence { font-size: 1.15rem; }
    .day-card { padding: 1rem 1.15rem; }
    .stat-card { padding: 0.85rem; }
    .stat-value { font-size: 1.25rem; }
}

</style>""", unsafe_allow_html=True)

# ── Session State Init ──
DEFAULTS = {
    "screen": "home", "level": None, "day": None, "deck": [], "idx": 0,
    "score": 0, "attempts": 0, "streak": 0, "max_streak": 0,
    "answered": False, "selected": None, "show_image": False,
    # tts_cache removed (not used) to avoid holding unused session data
    "debug_images": False,
    "image_cache": {},   # keyword -> (local_path, debug_info)
    "quiz_tooltip_map": {},
    "quiz_tooltip_example": None,
    "test_tooltip_map": {},
    "test_tooltip_sentence": None,
    # Search tab state
    "search_query": "",
    "search_results_cache": None,
    "search_shown_images": set(),
    # Random Test tab state
    "test_screen": "setup",  # setup | running | summary
    "test_level": None,
    "test_deck": [],
    "test_idx": 0,
    "test_score": 0,
    "test_attempts": 0,
    "test_answered": False,
    "test_selected": None,
    "test_start_time": None,
    "test_time_limit": 125,  # seconds
    "test_answers_log": [],  # list of dicts for review
}
for k, v in DEFAULTS.items():
    if k not in st.session_state:
        st.session_state[k] = v


def _badge(gender):
    m = {
        "masculine": ("Masculine (der)", "m"), "feminine": ("Feminine (die)", "f"),
        "neuter": ("Neuter (das)", "n"), "phrase": ("Social Phrase", "p"),
        "pronoun": ("Pronoun", "p"), "verb": ("Verb", "v"), "time": ("Time Component", "t"),
    }
    label, cls = m.get(gender, ("—", "p"))
    return f'<span class="badge badge-{cls}">{label}</span>'


def _image_to_data_uri(local_path):
    """Read a local image file and return a base64 data URI, or None on failure."""
    try:
        ext = os.path.splitext(local_path)[1].lower().lstrip(".")
        mime = "jpeg" if ext in ("jpg", "jpeg") else (ext or "png")
        with open(local_path, "rb") as f:
            b64 = base64.b64encode(f.read()).decode()
        return f"data:image/{mime};base64,{b64}"
    except Exception:
        return None


def _render_square_image(src):
    """Render an image (URL or data URI) inside the fixed 300x300 cropped box."""
    st.markdown(
        f'<div class="flashcard-img-box"><img src="{src}" /></div>',
        unsafe_allow_html=True,
    )


def _get_and_show_image(row, level, day, row_idx):
    row_dict = row.to_dict() if hasattr(row, "to_dict") else dict(row)
    keyword = str(row_dict.get("keyword", "")).strip()

    if keyword and keyword in st.session_state.image_cache:
        local_path, debug_info = st.session_state.image_cache[keyword]
    else:
        session_results = {}
        local_path, debug_info = fetch_image_for_row(row_dict, session_results)
        if keyword:
            st.session_state.image_cache[keyword] = (local_path, debug_info)

    if st.session_state.debug_images:
        with st.expander("🔍 Image search debug", expanded=True):
            st.write(f"**keyword col:** `{debug_info['keyword_col']}`")
            st.write(f"**german_word:** `{debug_info['german_word']}`")
            st.write(f"**➡️ Sent to DDG:** `{debug_info['search_term']}`")
            st.write(f"**Source:** `{debug_info['source']}`")
            st.write(f"**Chosen URL:** {debug_info.get('chosen_url', 'none')}")
            if "all_urls" in debug_info:
                st.write("**All DDG results:**")
                for u in debug_info["all_urls"]:
                    st.write(f"  - {u}")
            if "error" in debug_info:
                st.error(f"Error: {debug_info['error']}")

    if local_path:
        if str(local_path).lower().startswith(("http://", "https://")):
            _render_square_image(local_path)
            return
        elif os.path.exists(local_path):
            data_uri = _image_to_data_uri(local_path)
            if data_uri:
                _render_square_image(data_uri)
                return
        st.caption("Could not display image.")
        return

    st.caption("No image available.")


def _compute_audio_paths(deck, level, day_name=None):
    def _paths(row):
        d = day_name if day_name else str(row.get("_day", ""))
        p = get_audio_paths(level, d, int(row["_row_idx"]), str(row["german_word"]))
        return p["pronounce"], p["slow"], p["spell"]

    results = [_paths(row) for _, row in deck.iterrows()]
    deck["audio_pronounce_path"] = [r[0] for r in results]
    deck["audio_slow_path"]      = [r[1] for r in results]
    deck["audio_spell_path"]     = [r[2] for r in results]
    return deck


def _prewarm_audio(word, level, day, row_idx, sentence=None):
    """Fire-and-forget: generate and cache audio to disk before user needs it."""
    def _bg():
        try:
            if word:
                pronounce_word(word, level, day, row_idx)
                slow_word(word, level, day, row_idx)
                spell_word(word, level, day, row_idx)
            if sentence:
                pronounce_word(sentence)
        except Exception:
            pass
    threading.Thread(target=_bg, daemon=True).start()


def _prewarm_deck_cards(deck, level, indices):
    for i in indices:
        if 0 <= i < len(deck):
            row = deck.iloc[i]
            word = str(row.get("german_word", ""))
            actual_day = str(row.get("_day", ""))
            row_idx = int(row.get("_row_idx", 0))
            example = str(row.get("example_sentence", ""))
            if example == "nan":
                example = None
            if word:
                _prewarm_audio(word, level, actual_day, row_idx, example)


# ══════════════════════════════════════
# SCREEN: HOME
# ══════════════════════════════════════
def render_home():
    st.markdown('<div style="text-align:center"><span style="font-size:2rem">🇩🇪</span></div>', unsafe_allow_html=True)
    st.markdown('<h1 style="text-align:center;font-weight:800;font-size:2.8rem;background:linear-gradient(90deg,#818cf8,#a78bfa,#22d3ee);-webkit-background-clip:text;-webkit-text-fill-color:transparent;margin:0">DeutschFlash Master</h1>', unsafe_allow_html=True)
    st.markdown('<p style="text-align:center;color:#94a3b8;font-size:0.9rem;margin-bottom:2rem">Select a level and day to practice active retrieval & visual association</p>', unsafe_allow_html=True)

    levels = get_available_levels()
    if not levels:
        st.error("No vocabulary files found in `data/`. Run `python -m lib.data_generator` first.")
        return

    level = st.selectbox("Select Level", levels, key="level_select", help="Choose your CEFR level")

    # Check if this level is empty (template only)
    if is_level_empty(level):
        st.markdown(f"""<div class="glass-card" style="text-align:center;border-color:rgba(245,158,11,0.3)">
            <div style="font-size:2rem;margin-bottom:0.5rem">📝</div>
            <div style="font-weight:700;color:#fbbf24;font-size:1.1rem;margin-bottom:0.5rem">{level} — Empty Template</div>
            <div style="color:#94a3b8;font-size:0.9rem">
                This level file exists but has no vocabulary data yet.<br>
                 </div>
        </div>""", unsafe_allow_html=True)
        return

    day_info = get_day_info(level)
    if not day_info:
        st.warning(f"No data found for level {level}.")
        return

    st.markdown("---")
    total_words = sum(day_info.values())

    DAY_META = {
        "Day 1": ("Essential Nouns I", "Family, nature, weather & stationery", "#818cf8"),
        "Day 2": ("Social & Greetings", "Greetings, farewells, politeness & feelings", "#a78bfa"),
        "Day 3": ("Essential Nouns II", "Furniture, household, people & transport", "#34d399"),
        "Day 4": ("Directions, Seasons & Pronouns", "Cardinal directions, months & pronouns", "#fbbf24"),
        "Day 5": ("Verbs & Time Phrases", "Essential verbs & telling time", "#22d3ee"),
    }

    # Override for ENT level
    if level == "ENT":
        DAY_META = {
            "Day 1": ("Movie & TV Phrases", "Common subtitle lines from movies & TV shows", "#f472b6"),
            "Day 2": ("Emotions & Reactions", "Drama, reality TV emotions & reactions", "#fb923c"),
            "Day 3": ("Streaming & Pop Culture", "Social media, music & streaming vocabulary", "#a78bfa"),
        }

    for day_name, count in day_info.items():
        title, desc, color = DAY_META.get(day_name, (day_name, "", "#818cf8"))
        col1, col2 = st.columns([5, 1])
        with col1:
            st.markdown(f"""<div class="day-card">
                <div style="font-weight:700;font-size:1.1rem;color:#f1f5f9">{day_name}: {title}
                    <span style="font-size:0.7rem;background:rgba(99,102,241,0.15);color:{color};
                    padding:0.2rem 0.6rem;border-radius:999px;margin-left:0.5rem;
                    border:1px solid {color}30">{count} items</span>
                </div>
                <div style="font-size:0.9rem;color:#94a3b8;margin-top:0.4rem">{desc}</div>
            </div>""", unsafe_allow_html=True)
        with col2:
            if st.button("Start ➔", key=f"start_{day_name}", use_container_width=True):
                st.session_state.level = level
                st.session_state.day = day_name
                deck = load_vocab(level, day_name)
                deck = deck.sample(frac=1).reset_index(drop=True)
                deck = _compute_audio_paths(deck, level, day_name)

                st.session_state.deck = deck
                st.session_state.idx = 0
                st.session_state.score = 0
                st.session_state.attempts = 0
                st.session_state.streak = 0
                st.session_state.max_streak = 0
                st.session_state.answered = False
                st.session_state.selected = None
                st.session_state.show_image = False
                st.session_state.screen = "quiz"
                # image cache keys removed; no-op
                _prewarm_deck_cards(deck, level, [0, 1])
                st.rerun()

    st.markdown("---")
    col1, col2 = st.columns([5, 1])
    with col1:
        st.markdown(f"""<div class="day-card" style="background:linear-gradient(135deg,rgba(99,102,241,0.1),rgba(15,23,42,0.7));border-color:rgba(99,102,241,0.3)">
            <div style="font-weight:700;font-size:1.1rem;color:#a5b4fc">Ultimate Marathon Mode
                <span style="font-size:0.7rem;background:#6366f1;color:#0f172a;
                padding:0.2rem 0.6rem;border-radius:999px;margin-left:0.5rem;font-weight:800">{total_words} items</span>
            </div>
            <div style="font-size:0.9rem;color:#818cf8;margin-top:0.4rem">Study the entire deck in one session</div>
        </div>""", unsafe_allow_html=True)
    with col2:
        if st.button("Go!", key="start_marathon", use_container_width=True):
            st.session_state.level = level
            st.session_state.day = "__all__"
            deck = load_all_vocab(level)
            deck = deck.sample(frac=1).reset_index(drop=True)
            deck = _compute_audio_paths(deck, level)

            st.session_state.deck = deck
            st.session_state.idx = 0
            st.session_state.score = 0
            st.session_state.attempts = 0
            st.session_state.streak = 0
            st.session_state.max_streak = 0
            st.session_state.answered = False
            st.session_state.selected = None
            st.session_state.show_image = False
            st.session_state.screen = "quiz"
            # image cache keys removed; no-op
            _prewarm_deck_cards(deck, level, [0, 1])
            st.rerun()


# ══════════════════════════════════════
# SCREEN: QUIZ
# ══════════════════════════════════════
def render_quiz():
    deck = st.session_state.deck
    idx = st.session_state.idx

    if idx >= len(deck):
        st.session_state.screen = "summary"
        st.rerun()
        return

    # Trigger background audio pre-warming for current and upcoming cards
    prewarm_key = f"prewarmed_{idx}"
    if prewarm_key not in st.session_state:
        st.session_state[prewarm_key] = True
        _prewarm_deck_cards(deck, st.session_state.level, [idx, idx + 1, idx + 2])

    row = deck.iloc[idx]
    total = len(deck)

    # Level/Day label
    day_label = "Marathon" if st.session_state.day == "__all__" else st.session_state.day
    st.markdown(f'<div style="text-align:center;font-size:0.7rem;color:#818cf8;font-weight:500;letter-spacing:0.15em;text-transform:uppercase;">{st.session_state.level} — {day_label}</div>', unsafe_allow_html=True)

    # Progress / card counter
    st.markdown(f'<div style="font-size:0.85rem;color:#94a3b8;text-align:center;margin-top:0.1rem">Card {idx+1} of {total}</div>', unsafe_allow_html=True)

    gender = str(row.get("gender", ""))
    st.markdown(f'<div style="text-align:right;margin:0.5rem 0">{_badge(gender)}</div>', unsafe_allow_html=True)

    # Image
    if st.session_state.show_image or st.session_state.answered:
        actual_day = str(row.get("_day", st.session_state.day))
        if actual_day == "__all__":
            actual_day = st.session_state.day
        _get_and_show_image(row, st.session_state.level, actual_day, idx)
    else:
        if st.button("👁️ Show Image", key="reveal_img", use_container_width=True):
            st.session_state.show_image = True
            st.rerun()

    # Word display
    word = str(row.get("german_word", ""))
    phonetic = str(row.get("pronunciation", ""))
    note = str(row.get("note", ""))
    st.markdown(
        f'<div class="german-word">{word}'
        f' <span id="word-audio-trigger-{idx}" title="Play pronunciation"'
        f' style="cursor:pointer;font-size:1.5rem;vertical-align:middle;opacity:0.6;'
        f'transition:opacity 0.2s;display:inline-block"'
        f' onmouseover="this.style.opacity=1" onmouseout="this.style.opacity=0.6">🔊</span>'
        f'</div>',
        unsafe_allow_html=True,
    )
    st.markdown(f'<div style="text-align:center"><span class="phonetic">Pronounced: [{phonetic}]</span></div>', unsafe_allow_html=True)
    if note and note != "nan":
        st.markdown(f'<div style="text-align:center;margin-top:0.4rem"><span class="note-text">{note}</span></div>', unsafe_allow_html=True)
    # ── Word Audio (auto-play on load + 🔊 icon tap) ──
    actual_day = str(row.get("_day", st.session_state.day))
    if actual_day == "__all__":
        actual_day = st.session_state.day

    # Pre-generate word pronunciation audio (cached to disk)
    word_b64 = ""
    try:
        with st.spinner(""):
            word_audio_uri = pronounce_word(
                word,
                level=st.session_state.level,
                day=actual_day,
                row_idx=int(row["_row_idx"])
            )
        if word_audio_uri and "," in word_audio_uri:
            word_b64 = word_audio_uri.split(",", 1)[1]
    except Exception:
        pass

    # Auto-play word audio once when card first loads
    auto_key = f"auto_played_{idx}"
    if auto_key not in st.session_state and word_b64:
        st.session_state[auto_key] = True
        import streamlit.components.v1 as _c
        _c.html(
            f'<audio autoplay src="data:audio/mp3;base64,{word_b64}"></audio>',
            height=0
        )

    # JS: wire 🔊 word icon — tap to replay, no rerun needed
    import streamlit.components.v1 as _comp_word
    _comp_word.html(f"""
    <script>
    (function() {{
        const audioData = "{word_b64}";
        const pd = window.parent.document;
        function wire() {{
            const icon = pd.getElementById('word-audio-trigger-{idx}');
            if (!icon) return false;
            icon.onclick = function(e) {{
                e.preventDefault();
                if (!audioData) return;
                const audio = new Audio('data:audio/mp3;base64,' + audioData);
                audio.play();
            }};
            return true;
        }}
        if (!wire()) {{
            let tries = 0;
            const iv = setInterval(function() {{
                tries++;
                if (wire() || tries > 30) clearInterval(iv);
            }}, 150);
        }}
    }})();
    </script>
    """, height=0)

    st.markdown("<br>", unsafe_allow_html=True)

    # Example sentence with hover tooltips + 🔊 icon
    example = str(row.get("example_sentence", ""))
    if example and example != "nan":
        if "quiz_tooltip_map" not in st.session_state or st.session_state.quiz_tooltip_example != example:
            st.session_state.quiz_tooltip_map = translate_sentence_words(example, exclude_word=word)
            st.session_state.quiz_tooltip_example = example
        tooltips = st.session_state.quiz_tooltip_map
        highlighted_example = _highlight_word_in_sentence(example, word, tooltips=tooltips)

        # Sentence audio — generate once, cache in session_state (disk read ~10ms after first gen)
        sent_cache_key = f"quiz_audio_b64_{idx}"
        sentence_b64 = st.session_state.get(sent_cache_key, "")
        if not sentence_b64:
            try:
                _uri = pronounce_word(example)
                if _uri and "," in _uri:
                    sentence_b64 = _uri.split(",", 1)[1]
                    st.session_state[sent_cache_key] = sentence_b64
            except Exception:
                pass

        # Sentence with inline 🔊 at the end
        sentence_audio_icon = (
            f' <span id="sentence-audio-trigger-{idx}" title="Play sentence" '
            f'style="cursor:pointer;font-size:1rem;vertical-align:middle;opacity:0.6;'
            f'transition:opacity 0.2s;display:inline-block" '
            f'onmouseover="this.style.opacity=1" onmouseout="this.style.opacity=0.6">🔊</span>'
        )
        st.markdown(
            f'<div class="test-sentence" style="font-size:1.1rem;line-height:2;margin:0.25rem 0 0.75rem 0">'
            f'{highlighted_example}{sentence_audio_icon}</div>',
            unsafe_allow_html=True,
        )

        # JS: wire 🔊 icon to play audio client-side
        import streamlit.components.v1 as _comp_sent
        _comp_sent.html(f"""
        <script>
        (function() {{
            const audioData = "{sentence_b64}";
            const pd = window.parent.document;
            function wire() {{
                const icon = pd.getElementById('sentence-audio-trigger-{idx}');
                if (!icon) return false;
                icon.onclick = function(e) {{
                    e.preventDefault();
                    if (audioData) new Audio('data:audio/mp3;base64,' + audioData).play();
                }};
                return true;
            }}
            if (!wire()) {{
                let tries = 0;
                const iv = setInterval(function() {{
                    tries++;
                    if (wire() || tries > 30) clearInterval(iv);
                }}, 150);
            }}
        }})();
        </script>
        """, height=0)



    # MCQ
    correct = str(row.get("meaning", ""))
    options = [str(row.get(f"option_{i}", "") ) for i in range(1, 5)]
    options = [o for o in options if o and o != "nan"]
    if correct not in options:
        if len(options) >= 4:
            options[0] = correct
        else:
            options.append(correct)
    random.seed(hash(word + str(idx)))
    random.shuffle(options)

    if not st.session_state.answered:
        cols = st.columns(2)
        for i, opt in enumerate(options):
            with cols[i % 2]:
                if st.button(opt, key=f"opt_{i}_{idx}", use_container_width=True):
                    st.session_state.answered = True
                    st.session_state.selected = opt
                    st.session_state.attempts += 1
                    st.session_state.show_image = True
                    if opt == correct:
                        st.session_state.score += 1
                        st.session_state.streak += 1
                        st.session_state.max_streak = max(st.session_state.max_streak, st.session_state.streak)
                    else:
                        st.session_state.streak = 0
                    st.rerun()
    else:
        selected = st.session_state.selected
        is_correct = selected == correct
        cols = st.columns(2)
        for i, opt in enumerate(options):
            with cols[i % 2]:
                if opt == correct:
                    st.markdown(f'<div class="opt-btn opt-correct">✓ {opt}</div>', unsafe_allow_html=True)
                elif opt == selected and not is_correct:
                    st.markdown(f'<div class="opt-btn opt-wrong">✗ {opt}</div>', unsafe_allow_html=True)
                else:
                    st.markdown(f'<div class="opt-btn" style="opacity:0.4">{opt}</div>', unsafe_allow_html=True)

        if is_correct:
            st.markdown(f"""<div class="fb-correct">
                <div style="font-weight:700;font-size:0.85rem;letter-spacing:0.1em;text-transform:uppercase;margin-bottom:0.4rem">Richtig! (Correct)</div>
                <div style="font-size:1.05rem">"{word}" = "{correct}"</div>
            </div>""", unsafe_allow_html=True)
        else:
            st.markdown(f"""<div class="fb-wrong">
                <div style="font-weight:700;font-size:0.85rem;letter-spacing:0.1em;text-transform:uppercase;margin-bottom:0.4rem">Falsch! (Incorrect)</div>
                <div style="font-size:1.05rem">"{word}" = "{correct}" (not "{selected}")</div>
            </div>""", unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("Next Card ➔", key="next_btn", type="primary", use_container_width=True):
            st.session_state.idx += 1
            st.session_state.answered = False
            st.session_state.selected = None
            st.session_state.show_image = False
            st.rerun()

    # Audio trace debugging
    if st.session_state.debug_images:
        actual_day = str(row.get("_day", st.session_state.day))
        if actual_day == "__all__":
            actual_day = st.session_state.day
        paths = get_audio_paths(st.session_state.level, actual_day, int(row["_row_idx"]), word)
        with st.expander("🔊 Audio Cache Trace", expanded=True):
            st.write(f"**Pronounce path:** `{paths['pronounce']}`")
            st.write(f"**Pronounce exists:** `{os.path.exists(paths['pronounce'])}`")
            st.write(f"**Slow path:** `{paths['slow']}`")
            st.write(f"**Slow exists:** `{os.path.exists(paths['slow'])}`")
            st.write(f"**Spell path:** `{paths['spell']}`")
            st.write(f"**Spell exists:** `{os.path.exists(paths['spell'])}`")


# ══════════════════════════════════════
# SCREEN: SUMMARY
# ══════════════════════════════════════
def render_summary():
    st.markdown('<h1 style="text-align:center;font-weight:800;font-size:2.5rem;color:#f8fafc;margin-top:1rem">Glückwunsch!</h1>', unsafe_allow_html=True)
    st.markdown('<p style="text-align:center;color:#94a3b8;font-size:1.1rem">You\'ve completed this practice session!</p>', unsafe_allow_html=True)

    total = len(st.session_state.deck)
    score = st.session_state.score
    attempts = st.session_state.attempts
    accuracy = round((score / attempts) * 100) if attempts > 0 else 0
    max_streak = st.session_state.max_streak

    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown(f"""<div class="stat-card">
            <div class="stat-label">Score</div>
            <div class="stat-value" style="color:#818cf8">{score}/{total}</div>
        </div>""", unsafe_allow_html=True)
    with c2:
        st.markdown(f"""<div class="stat-card">
            <div class="stat-label">Accuracy</div>
            <div class="stat-value" style="color:#34d399">{accuracy}%</div>
        </div>""", unsafe_allow_html=True)
    with c3:
        st.markdown(f"""<div class="stat-card">
            <div class="stat-label">Max Streak</div>
            <div class="stat-value" style="color:#fbbf24">{max_streak}</div>
        </div>""", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1:
        if st.button("Restart This Day", use_container_width=True, type="primary"):
            st.session_state.deck = st.session_state.deck.sample(frac=1).reset_index(drop=True)
            st.session_state.idx = 0
            st.session_state.score = 0
            st.session_state.attempts = 0
            st.session_state.streak = 0
            st.session_state.max_streak = 0
            st.session_state.answered = False
            st.session_state.selected = None
            st.session_state.show_image = False
            st.session_state.screen = "quiz"
            st.rerun()
    with c2:
        if st.button("Back to Home", use_container_width=True):
            st.session_state.screen = "home"
            st.rerun()


# ══════════════════════════════════════
# TAB: SEARCH
# ══════════════════════════════════════
@st.cache_data(show_spinner=False, ttl=600)
def _build_search_index():
    """Load every level's full vocab into one combined DataFrame, tagged with level."""
    levels = get_available_levels()
    frames = []
    for level in levels:
        if is_level_empty(level):
            continue
        df = load_all_vocab(level)
        if df is None or df.empty:
            continue
        df = df.copy()
        df["_level"] = level
        frames.append(df)
    if not frames:
        return pd.DataFrame()
    return pd.concat(frames, ignore_index=True)


def _search_matches(query: str, df: pd.DataFrame) -> pd.DataFrame:
    if df.empty or not query.strip():
        return df.iloc[0:0]
    q = query.strip().lower()
    word_col = df.get("german_word", pd.Series([""] * len(df))).astype(str).str.lower()
    meaning_col = df.get("meaning", pd.Series([""] * len(df))).astype(str).str.lower()
    mask = word_col.str.contains(q, na=False, regex=False) | meaning_col.str.contains(q, na=False, regex=False)
    matches = df[mask].copy()
    # Rank: exact german_word match first, then "starts with", then everything else
    def _rank(w):
        w = str(w).lower()
        if w == q:
            return 0
        if w.startswith(q):
            return 1
        return 2
    matches["_rank"] = matches["german_word"].apply(_rank)
    matches = matches.sort_values("_rank").drop(columns="_rank").reset_index(drop=True)
    return matches


def _render_search_card(row, card_key: str):
    word = str(row.get("german_word", ""))
    meaning = str(row.get("meaning", ""))
    phonetic = str(row.get("pronunciation", ""))
    note = str(row.get("note", ""))
    gender = str(row.get("gender", ""))
    example = str(row.get("example_sentence", ""))
    level = str(row.get("_level", ""))
    day = str(row.get("_day", ""))

    st.markdown('<div class="search-card">', unsafe_allow_html=True)

    loc_label = f"{level}" + (f" — {day}" if day and day != "nan" else "")
    st.markdown(
        f'<div class="search-meta-row"><span class="search-loc">{loc_label}</span>{_badge(gender)}</div>',
        unsafe_allow_html=True,
    )

    st.markdown(f'<div class="search-meaning">= {meaning}</div>', unsafe_allow_html=True)

    if example and example != "nan":
        st.markdown(f'<div class="search-example">📝 {example}</div>', unsafe_allow_html=True)

    # Image — on demand
    if card_key in st.session_state.search_shown_images:
        _get_and_show_image(row, level, day, card_key)
    else:
        if st.button("👁️ Show Image", key=f"search_img_{card_key}", use_container_width=True):
            st.session_state.search_shown_images.add(card_key)
            st.rerun()

    row_idx = int(row["_row_idx"]) if "_row_idx" in row and pd.notna(row["_row_idx"]) else 0
    actual_day = day if day and day != "nan" else "Search"

    # Pre-generate word audio for this search card
    search_word_b64 = ""
    try:
        search_audio_uri = pronounce_word(word, level=level, day=actual_day, row_idx=row_idx)
        if search_audio_uri and "," in search_audio_uri:
            search_word_b64 = search_audio_uri.split(",", 1)[1]
    except Exception:
        pass

    # Unique trigger ID per search card (card_key is already unique)
    trigger_id = f"search-word-audio-{card_key}"

    # Word + inline 🔊 icon
    st.markdown(
        f'<div class="german-word">{word}'
        f' <span id="{trigger_id}" title="Play pronunciation" '
        f'style="cursor:pointer;font-size:1.5rem;vertical-align:middle;opacity:0.6;'
        f'transition:opacity 0.2s;display:inline-block" '
        f'onmouseover="this.style.opacity=1" onmouseout="this.style.opacity=0.6">🔊</span>'
        f'</div>',
        unsafe_allow_html=True,
    )

    # JS: wire 🔊 icon — plays client-side, no rerun needed
    import streamlit.components.v1 as _comp_sw
    _comp_sw.html(f"""
    <script>
    (function() {{
        const audioData = "{search_word_b64}";
        const pd = window.parent.document;
        function wire() {{
            const icon = pd.getElementById('{trigger_id}');
            if (!icon) return false;
            icon.onclick = function(e) {{
                e.preventDefault();
                if (!audioData) return;
                const audio = new Audio('data:audio/mp3;base64,' + audioData);
                audio.play();
            }};
            return true;
        }}
        if (!wire()) {{
            let tries = 0;
            const iv = setInterval(function() {{
                tries++;
                if (wire() || tries > 30) clearInterval(iv);
            }}, 150);
        }}
    }})();
    </script>
    """, height=0)

    # Clean up old play_mode state if any lingered
    st.session_state.pop(f"search_play_{card_key}", None)

    st.markdown('</div>', unsafe_allow_html=True)


def render_search():
    st.markdown(
        '<h2 style="text-align:center;font-weight:800;font-size:1.8rem;color:#f1f5f9;margin-bottom:0.25rem">🔍 Word Search</h2>',
        unsafe_allow_html=True,
    )
    st.markdown(
        '<p style="text-align:center;color:#94a3b8;font-size:0.875rem;margin-bottom:1.5rem">Search any German word or English meaning across all levels</p>',
        unsafe_allow_html=True,
    )

    query = st.text_input(
        "Search",
        value=st.session_state.search_query,
        placeholder="e.g. Haus, mother, danke...",
        label_visibility="collapsed",
        key="search_input",
    )
    st.session_state.search_query = query

    index_df = _build_search_index()

    if index_df.empty:
        st.markdown(
            '<div class="search-empty">No vocabulary data found. Add words under <code>data/</code> first.</div>',
            unsafe_allow_html=True,
        )
        return

    if not query.strip():
        st.markdown(
            '<div class="search-empty">👆 Start typing to search across every level and day.</div>',
            unsafe_allow_html=True,
        )
        return

    matches = _search_matches(query, index_df)

    if matches.empty:
        st.markdown(
            f'<div class="search-empty">No results for "<strong>{query}</strong>". Try a different word.</div>',
            unsafe_allow_html=True,
        )
        return

    st.markdown(
        f'<div style="color:#64748b;font-size:0.8rem;margin-bottom:0.75rem">{len(matches)} result(s) found</div>',
        unsafe_allow_html=True,
    )

    MAX_RESULTS = 25
    for i, (_, row) in enumerate(matches.head(MAX_RESULTS).iterrows()):
        card_key = f"{row.get('_level','')}_{row.get('_day','')}_{row.get('_row_idx', i)}_{i}"
        _render_search_card(row, card_key)

    if len(matches) > MAX_RESULTS:
        st.caption(f"Showing first {MAX_RESULTS} of {len(matches)} results — refine your search for more precise matches.")


# ══════════════════════════════════════
# TAB: RANDOM TEST
# Timed sentence-context quiz with dict.cc hover tooltips.
# The student reads a German sentence, identifies the highlighted
# word, and picks its English meaning from four MCQ options.
# ══════════════════════════════════════
TEST_WORD_COUNT = 25
TEST_TIME_LIMIT = 250  # 25 words × 10s each


def _highlight_word_in_sentence(sentence: str, word: str, tooltips: dict[str, str] | None = None) -> str:
    """Build the test sentence HTML with highlight + dict.cc tooltips.

    1. The **quiz word** is wrapped in ``<span class="highlight-word">`` — no
       tooltip, because showing its meaning would give away the answer.
    2. Every other significant word gets a ``<span class="tt-word"
       data-tip="english">`` so the student can hover for help.
    3. Articles, pronouns, and punctuation are left plain.

    Args:
        sentence:  The German example sentence.
        word:      The German quiz word (to highlight, NOT tooltip).
        tooltips:  Dict of ``{german_token: english_meaning}`` from
                   :func:`translate_sentence_words`. ``None`` = no tooltips.

    Returns:
        HTML string ready for ``st.markdown(unsafe_allow_html=True)``.
    """
    import re
    if tooltips is None:
        tooltips = {}

    # Determine which tokens belong to the quiz word (to exclude from tooltips)
    quiz_tokens: set[str] = set()
    quiz_tokens.add(word.lower())
    bare = re.sub(r'^(der|die|das|ein|eine|einen)\s+', '', word, flags=re.IGNORECASE).strip()
    if bare:
        quiz_tokens.add(bare.lower())
    # Also add individual tokens of multi-word quiz words (e.g. "der Vater" → "Vater")
    for part in word.split():
        quiz_tokens.add(part.lower())

    # First pass: mark the quiz word position in the sentence
    # Try exact match, then bare noun match
    quiz_marked = False
    MARKER = '\x00QUIZ_START\x00'
    MARKER_END = '\x00QUIZ_END\x00'

    pattern = re.compile(re.escape(word), re.IGNORECASE)
    if pattern.search(sentence):
        sentence = pattern.sub(lambda m: f'{MARKER}{m.group()}{MARKER_END}', sentence, count=1)
        quiz_marked = True
    elif bare:
        pattern = re.compile(re.escape(bare), re.IGNORECASE)
        if pattern.search(sentence):
            sentence = pattern.sub(lambda m: f'{MARKER}{m.group()}{MARKER_END}', sentence, count=1)
            quiz_marked = True

    # Split into segments around the quiz word marker
    if quiz_marked:
        parts = re.split(r'\x00QUIZ_START\x00|\x00QUIZ_END\x00', sentence)
        # parts = [before, quiz_word_text, after]
    else:
        parts = [sentence]

    def _wrap_tokens(text: str) -> str:
        """Wrap individual words with tooltips where available."""
        tokens = text.split(' ')
        result_tokens = []
        for tok in tokens:
            # Strip punctuation to match against tooltip dict
            stripped = tok.strip('.,!?;:"\'()[]–—-…')
            if stripped and stripped in tooltips and stripped.lower() not in quiz_tokens:
                tip = tooltips[stripped].replace('"', '&quot;')
                # Re-attach surrounding punctuation
                prefix = tok[:tok.index(stripped)] if stripped in tok else ''
                suffix = tok[tok.index(stripped) + len(stripped):] if stripped in tok else ''
                result_tokens.append(f'{prefix}<span class="tt-word" data-tip="{tip}">{stripped}</span>{suffix}')
            else:
                result_tokens.append(tok)
        return ' '.join(result_tokens)

    # Build final HTML
    html_parts = []
    if quiz_marked and len(parts) >= 3:
        html_parts.append(_wrap_tokens(parts[0]))
        html_parts.append(f'<span class="highlight-word">{parts[1]}</span>')
        html_parts.append(_wrap_tokens(parts[2]))
    elif quiz_marked and len(parts) == 2:
        html_parts.append(_wrap_tokens(parts[0]))
        html_parts.append(f'<span class="highlight-word">{parts[1]}</span>')
    else:
        # Fallback: no match found — append quiz word at end
        html_parts.append(_wrap_tokens(sentence))
        html_parts.append(f' <span class="highlight-word">[{word}]</span>')

    return ''.join(html_parts)


SHEET_TEST_WORD_COUNT = 10
SHEET_TEST_TIME_LIMIT = 100  # 10 words × 10s each


def render_test_setup():
    st.markdown('<div style="text-align:center"><span style="font-size:2rem">🎯</span></div>', unsafe_allow_html=True)
    st.markdown(
        '<h1 style="text-align:center;font-weight:800;font-size:2.4rem;'
        'background:linear-gradient(90deg,#f472b6,#fb923c,#fbbf24);'
        '-webkit-background-clip:text;-webkit-text-fill-color:transparent;margin:0">'
        'Random Test</h1>',
        unsafe_allow_html=True,
    )
    st.markdown(
        '<p style="text-align:center;color:#94a3b8;font-size:0.9rem;margin-bottom:1.5rem">'
        'Timed challenge — read the sentence, spot the highlighted word, pick the meaning!</p>',
        unsafe_allow_html=True,
    )

    levels = get_available_levels()
    non_empty = [l for l in levels if not is_level_empty(l)]
    if not non_empty:
        st.error("No vocabulary data found.")
        return

    level = st.selectbox("Select Level", non_empty, key="test_level_select")

    day_info = get_day_info(level)
    total_words = sum(day_info.values())

    # Test mode toggle
    test_mode = st.radio(
        "Test Mode",
        ["📚 Level-wise (25 words · 125s)", "📄 Sheet-wise (10 words · 50s)"],
        key="test_mode_radio",
        horizontal=True,
    )
    is_sheet_mode = "Sheet-wise" in test_mode

    # Sheet selector (only shown in sheet mode)
    selected_sheet = None
    sheet_word_count = 0
    if is_sheet_mode:
        sheet_names = list(day_info.keys())
        if not sheet_names:
            st.warning("No sheets with data found in this level.")
            return
        selected_sheet = st.selectbox(
            "Select Sheet",
            sheet_names,
            key="test_sheet_select",
            format_func=lambda s: f"{s}  ({day_info.get(s, 0)} words)",
        )
        sheet_word_count = day_info.get(selected_sheet, 0)

    # Determine counts for display
    q_count = SHEET_TEST_WORD_COUNT if is_sheet_mode else TEST_WORD_COUNT
    t_limit = SHEET_TEST_TIME_LIMIT if is_sheet_mode else TEST_TIME_LIMIT
    pool_words = sheet_word_count if is_sheet_mode else total_words
    pool_label = f"{selected_sheet}" if is_sheet_mode else f"{len(day_info)} sheet(s)"

    # Info card
    st.markdown(f"""
    <div class="glass-card" style="text-align:center">
        <div style="font-size:1.5rem;margin-bottom:0.3rem">📊</div>
        <div style="font-weight:700;color:#f1f5f9;font-size:1.05rem">{level} — {pool_words} words available</div>
        <div style="color:#94a3b8;font-size:0.85rem;margin-top:0.35rem">
            {pool_label} · {q_count} random questions · {t_limit}s time limit
        </div>
    </div>
    """, unsafe_allow_html=True)

    if pool_words < 4:
        st.warning("Need at least 4 words to generate a test.")
        return

    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("🚀 Start Test", key="start_test_btn", type="primary", use_container_width=True):
        # Load vocabulary based on mode
        if is_sheet_mode and selected_sheet:
            vocab = load_vocab(level, selected_sheet)
            if vocab.empty:
                st.error("Could not load vocabulary for this sheet.")
                return
            # Add _day column for consistency
            vocab["_day"] = selected_sheet
        else:
            vocab = load_all_vocab(level)
            if vocab.empty:
                st.error("Could not load vocabulary.")
                return

        # Filter rows that have a valid example_sentence
        valid = vocab[
            vocab["example_sentence"].astype(str).str.strip().ne("")
            & vocab["example_sentence"].notna()
            & vocab["example_sentence"].astype(str).ne("nan")
        ].copy()
        if len(valid) < 4:
            st.error("Not enough words with example sentences to run a test.")
            return
        n = min(q_count, len(valid))
        test_deck = valid.sample(n=n).reset_index(drop=True)
        st.session_state.test_level = level
        st.session_state.test_deck = test_deck
        st.session_state.test_idx = 0
        st.session_state.test_score = 0
        st.session_state.test_attempts = 0
        st.session_state.test_answered = False
        st.session_state.test_selected = None
        st.session_state.test_start_time = time.time()
        st.session_state.test_time_limit = t_limit
        st.session_state.test_answers_log = []
        st.session_state.test_screen = "running"
        st.rerun()


def render_test_running():
    deck = st.session_state.test_deck
    idx = st.session_state.test_idx
    level = st.session_state.test_level

    # ── Check time (server-side) ──
    elapsed = time.time() - st.session_state.test_start_time
    remaining = max(0, st.session_state.test_time_limit - elapsed)

    if remaining <= 0 or idx >= len(deck):
        st.session_state.test_screen = "summary"
        st.rerun()
        return

    row = deck.iloc[idx]
    total = len(deck)

    # Hidden time-up button — JS will click this when countdown hits zero
    if st.button("⏰ Time Up", key="test_time_up_btn", type="secondary", use_container_width=True):
        st.session_state.test_screen = "summary"
        st.rerun()
        return

    # Timer placeholder — JS will update this live
    remaining_int = int(remaining)
    mins = remaining_int // 60
    secs = remaining_int % 60
    st.markdown(
        f'<div class="test-timer timer-ok" id="live-timer">⏱️ <span id="timer-value">{mins}:{secs:02d}</span></div>',
        unsafe_allow_html=True,
    )

    # Client-side JS: live countdown + auto-click time-up button when done
    import streamlit.components.v1 as components
    components.html(f"""
    <script>
    (function() {{
        let remaining = {remaining_int};
        const parentDoc = window.parent.document;

        function findTimerEl() {{
            return parentDoc.getElementById('timer-value');
        }}
        function findTimerBox() {{
            return parentDoc.getElementById('live-timer');
        }}
        function findTimeUpBtn() {{
            const buttons = parentDoc.querySelectorAll('button');
            for (const btn of buttons) {{
                if (btn.textContent.includes('Time Up')) return btn;
            }}
            return null;
        }}

        // Hide the time-up button visually
        const hideBtn = findTimeUpBtn();
        if (hideBtn) {{
            hideBtn.style.display = 'none';
        }}

        function tick() {{
            remaining--;
            if (remaining <= 0) {{
                const btn = findTimeUpBtn();
                if (btn) {{
                    btn.style.display = '';
                    btn.click();
                }}
                return;
            }}
            const el = findTimerEl();
            const box = findTimerBox();
            if (el) {{
                const m = Math.floor(remaining / 60);
                const s = remaining % 60;
                el.textContent = m + ':' + String(s).padStart(2, '0');
            }}
            if (box) {{
                box.className = 'test-timer ' + (remaining > 60 ? 'timer-ok' : remaining > 30 ? 'timer-warn' : 'timer-danger');
            }}
            setTimeout(tick, 1000);
        }}
        setTimeout(tick, 1000);
    }})();
    </script>
    """, height=0)

    # Header
    st.markdown(
        f'<div style="text-align:center;font-size:0.7rem;color:#f472b6;font-weight:500;'
        f'letter-spacing:0.15em;text-transform:uppercase;">{level} — Random Test</div>',
        unsafe_allow_html=True,
    )
    st.markdown(
        f'<div style="font-size:0.85rem;color:#94a3b8;text-align:center;margin-top:0.1rem">'
        f'Question {idx + 1} of {total}</div>',
        unsafe_allow_html=True,
    )

    # Progress bar — shown once, reflects current question position
    st.progress((idx + 1) / total)


    gender = str(row.get("gender", ""))
    st.markdown(f'<div style="text-align:right;margin:0.5rem 0">{_badge(gender)}</div>', unsafe_allow_html=True)

    # Show sentence with highlighted word + dict.cc hover tooltips
    word = str(row.get("german_word", ""))
    sentence = str(row.get("example_sentence", ""))

    # Translate non-quiz words for hover tooltips (cached — only hits dict.cc once per word)
    if "test_tooltip_map" not in st.session_state or st.session_state.test_tooltip_sentence != sentence:
        st.session_state.test_tooltip_map = translate_sentence_words(sentence, exclude_word=word)
        st.session_state.test_tooltip_sentence = sentence
    tooltips = st.session_state.test_tooltip_map

    highlighted = _highlight_word_in_sentence(sentence, word, tooltips=tooltips)

    # Audio icon appended inline at end of sentence
    audio_inline_btn = (
        ' <span class="inline-audio-btn" id="inline-audio-trigger" '
        'title="Play pronunciation" '
        'style="cursor:pointer;font-size:1.1rem;vertical-align:middle;'
        'margin-left:0.4rem;opacity:0.7;transition:opacity 0.2s;display:inline-block" '
        'onmouseover="this.style.opacity=1" onmouseout="this.style.opacity=0.7">'
        '🔊</span>'
    )
    st.markdown(
        f'<div class="test-sentence">{highlighted}{audio_inline_btn}</div>',
        unsafe_allow_html=True,
    )
    st.markdown(
        '<div style="color:#64748b;font-size:0.75rem;text-align:center;margin-top:-0.25rem;margin-bottom:0.5rem">'
        'What does the highlighted word mean?</div>',
        unsafe_allow_html=True,
    )

    # Sentence audio — generate once, cache in session_state (disk read ~10ms after first gen)
    sentence_cache_key = f"test_audio_b64_{idx}"
    sentence_b64 = st.session_state.get(sentence_cache_key, "")
    if not sentence_b64:
        try:
            _uri = pronounce_word(sentence)
            if _uri and "," in _uri:
                sentence_b64 = _uri.split(",", 1)[1]
                st.session_state[sentence_cache_key] = sentence_b64
        except Exception:
            pass

    # JS: wire 🔊 icon to play audio client-side — no button, no rerun
    import streamlit.components.v1 as _comp_audio
    _comp_audio.html(f"""
    <script>
    (function() {{
        const audioData = "{sentence_b64}";
        const pd = window.parent.document;
        function wire() {{
            const icon = pd.getElementById('inline-audio-trigger');
            if (!icon) return false;
            icon.onclick = function(e) {{
                e.preventDefault();
                if (audioData) new Audio('data:audio/mp3;base64,' + audioData).play();
            }};
            return true;
        }}
        if (!wire()) {{
            let tries = 0;
            const iv = setInterval(function() {{
                tries++;
                if (wire() || tries > 30) clearInterval(iv);
            }}, 150);
        }}
    }})();
    </script>
    """, height=0)


    # MCQ options
    correct = str(row.get("meaning", ""))
    options = [str(row.get(f"option_{i}", "")) for i in range(1, 5)]
    options = [o for o in options if o and o != "nan"]
    if correct not in options:
        if len(options) >= 4:
            options[0] = correct
        else:
            options.append(correct)
    random.seed(hash(word + str(idx) + "test"))
    random.shuffle(options)

    if not st.session_state.test_answered:
        st.markdown('<div class="mcq-container"></div>', unsafe_allow_html=True)
        cols = st.columns(2)
        for i, opt in enumerate(options):
            with cols[i % 2]:
                if st.button(opt, key=f"test_opt_{i}_{idx}", use_container_width=True):
                    # Re-check time on answer click
                    if time.time() - st.session_state.test_start_time >= st.session_state.test_time_limit:
                        st.session_state.test_screen = "summary"
                        st.rerun()
                        return
                    st.session_state.test_answered = True
                    st.session_state.test_selected = opt
                    st.session_state.test_attempts += 1
                    is_right = (opt == correct)
                    if is_right:
                        st.session_state.test_score += 1
                    st.session_state.test_answers_log.append({
                        "word": word,
                        "sentence": sentence,
                        "correct": correct,
                        "selected": opt,
                        "is_correct": is_right,
                    })
                    st.rerun()
    else:
        selected = st.session_state.test_selected
        is_correct = selected == correct
        cols = st.columns(2)
        for i, opt in enumerate(options):
            with cols[i % 2]:
                if opt == correct:
                    st.markdown(f'<div class="opt-btn opt-correct">✓ {opt}</div>', unsafe_allow_html=True)
                elif opt == selected and not is_correct:
                    st.markdown(f'<div class="opt-btn opt-wrong">✗ {opt}</div>', unsafe_allow_html=True)
                else:
                    st.markdown(f'<div class="opt-btn" style="opacity:0.4">{opt}</div>', unsafe_allow_html=True)

        phonetic = str(row.get("pronunciation", ""))
        if is_correct:
            st.markdown(f"""<div class="fb-correct">
                <div style="font-weight:700;font-size:0.85rem;letter-spacing:0.1em;text-transform:uppercase;margin-bottom:0.4rem">Richtig!</div>
                <div style="font-size:1.05rem">"{word}" [{phonetic}] = "{correct}"</div>
            </div>""", unsafe_allow_html=True)
        else:
            st.markdown(f"""<div class="fb-wrong">
                <div style="font-weight:700;font-size:0.85rem;letter-spacing:0.1em;text-transform:uppercase;margin-bottom:0.4rem">Falsch!</div>
                <div style="font-size:1.05rem">"{word}" [{phonetic}] = "{correct}" (not "{selected}")</div>
            </div>""", unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("Next ➔", key="test_next_btn", type="primary", use_container_width=True):
            # Re-check time on next click
            if time.time() - st.session_state.test_start_time >= st.session_state.test_time_limit:
                st.session_state.test_screen = "summary"
                st.rerun()
                return
            st.session_state.test_idx += 1
            st.session_state.test_answered = False
            st.session_state.test_selected = None
            st.rerun()


def render_test_summary():
    elapsed = time.time() - st.session_state.test_start_time
    time_up = elapsed >= st.session_state.test_time_limit
    total_answered = st.session_state.test_attempts
    total_questions = len(st.session_state.test_deck)
    score = st.session_state.test_score
    accuracy = round((score / total_answered) * 100) if total_answered > 0 else 0
    time_taken = min(elapsed, st.session_state.test_time_limit)
    mins = int(time_taken // 60)
    secs = int(time_taken % 60)

    # Header
    if time_up:
        st.markdown(
            '<div style="text-align:center;font-size:2.5rem;margin-bottom:0.25rem">⏰</div>',
            unsafe_allow_html=True,
        )
        st.markdown(
            '<h1 style="text-align:center;font-weight:800;font-size:2.2rem;color:#fbbf24;margin:0">Time\'s Up!</h1>',
            unsafe_allow_html=True,
        )
    else:
        st.markdown(
            '<div style="text-align:center;font-size:2.5rem;margin-bottom:0.25rem">🏆</div>',
            unsafe_allow_html=True,
        )
        st.markdown(
            '<h1 style="text-align:center;font-weight:800;font-size:2.2rem;color:#34d399;margin:0">Test Complete!</h1>',
            unsafe_allow_html=True,
        )

    st.markdown(
        f'<p style="text-align:center;color:#94a3b8;font-size:0.95rem;margin-bottom:1.5rem">{st.session_state.test_level} Level Random Test</p>',
        unsafe_allow_html=True,
    )

    # Stats row
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.markdown(f"""<div class="stat-card">
            <div class="stat-label">Score</div>
            <div class="stat-value" style="color:#818cf8">{score}/{total_questions}</div>
        </div>""", unsafe_allow_html=True)
    with c2:
        st.markdown(f"""<div class="stat-card">
            <div class="stat-label">Answered</div>
            <div class="stat-value" style="color:#a78bfa">{total_answered}/{total_questions}</div>
        </div>""", unsafe_allow_html=True)
    with c3:
        st.markdown(f"""<div class="stat-card">
            <div class="stat-label">Accuracy</div>
            <div class="stat-value" style="color:#34d399">{accuracy}%</div>
        </div>""", unsafe_allow_html=True)
    with c4:
        st.markdown(f"""<div class="stat-card">
            <div class="stat-label">Time</div>
            <div class="stat-value" style="color:#fbbf24">{mins}:{secs:02d}</div>
        </div>""", unsafe_allow_html=True)

    # Answers review
    log = st.session_state.test_answers_log
    if log:
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown(
            '<div style="font-weight:700;color:#f1f5f9;font-size:1.1rem;margin-bottom:0.75rem">📋 Answer Review</div>',
            unsafe_allow_html=True,
        )
        for i, entry in enumerate(log):
            icon = "✅" if entry["is_correct"] else "❌"
            word = entry["word"]
            correct_ans = entry["correct"]
            selected_ans = entry["selected"]
            sentence = entry["sentence"]
            if entry["is_correct"]:
                st.markdown(
                    f'<div style="background:rgba(16,185,129,0.06);border:1px solid rgba(16,185,129,0.15);'
                    f'border-radius:12px;padding:0.75rem 1rem;margin-bottom:0.5rem;font-size:0.9rem">'
                    f'{icon} <strong>{word}</strong> = {correct_ans}</div>',
                    unsafe_allow_html=True,
                )
            else:
                st.markdown(
                    f'<div style="background:rgba(239,68,68,0.06);border:1px solid rgba(239,68,68,0.15);'
                    f'border-radius:12px;padding:0.75rem 1rem;margin-bottom:0.5rem;font-size:0.9rem">'
                    f'{icon} <strong>{word}</strong> = {correct_ans}'
                    f'<span style="color:#f87171"> (you: {selected_ans})</span></div>',
                    unsafe_allow_html=True,
                )

    # Action buttons
    st.markdown("<br>", unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1:
        if st.button("🔄 Retake Test", key="test_retake", type="primary", use_container_width=True):
            st.session_state.test_screen = "setup"
            st.rerun()
    with c2:
        if st.button("🏠 Back to Setup", key="test_home", use_container_width=True):
            st.session_state.test_screen = "setup"
            st.rerun()
def render_test_tab():
    screen = st.session_state.test_screen
    if screen == "setup":
        render_test_setup()
    elif screen == "running":
        render_test_running()
    elif screen == "summary":
        render_test_summary()


def render_visuals_tab():
    import json
    from lib.pptx_loader import get_slide_images
    
    st.markdown(
        '<h2 style="text-align:center;font-weight:800;font-size:1.8rem;color:#f1f5f9;margin-bottom:0.25rem">🖼️ Learn by Visuals</h2>',
        unsafe_allow_html=True,
    )
    st.markdown(
        '<p style="text-align:center;color:#94a3b8;font-size:0.875rem;margin-bottom:1.5rem">Study slides, swipe left/right, and select slides directly</p>',
        unsafe_allow_html=True,
    )
    
    with st.spinner("Loading visual slides..."):
        slides = get_slide_images()
        
    if not slides:
        st.markdown(
            '<div class="search-empty">Please place <code>German.pptx</code> in the <code>data/</code> folder to start visual learning.</div>',
            unsafe_allow_html=True,
        )
        return

    # Embed the custom slide viewer component
    slides_json = json.dumps(slides)
    
    html_code = f"""
    <!DOCTYPE html>
    <html>
    <head>
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <style>
    * {{ box-sizing: border-box; margin: 0; padding: 0; }}
    body {{
        background: transparent;
        color: #f8fafc;
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
        overflow: hidden;
        height: 520px;
        display: flex;
        flex-direction: column;
    }}
    .viewer-container {{
        flex: 1;
        display: flex;
        flex-direction: column;
        position: relative;
        width: 100%;
        height: 100%;
        align-items: center;
        justify-content: center;
    }}
    /* Top Controls Header */
    .top-bar {{
        width: 100%;
        display: flex;
        align-items: center;
        justify-content: space-between;
        padding: 8px 12px;
        background: rgba(15, 23, 42, 0.45);
        backdrop-filter: blur(12px);
        border: 1px solid rgba(255, 255, 255, 0.08);
        border-radius: 12px;
        z-index: 10;
        gap: 10px;
        margin-bottom: 12px;
    }}
    .select-wrapper {{
        position: relative;
        flex: 1;
    }}
    select.slide-select {{
        width: 100%;
        padding: 8px 30px 8px 12px;
        border-radius: 8px;
        background: rgba(30, 41, 59, 0.7);
        color: #f8fafc;
        border: 1px solid rgba(255, 255, 255, 0.15);
        font-size: 0.9rem;
        font-weight: 500;
        outline: none;
        cursor: pointer;
        appearance: none;
        -webkit-appearance: none;
    }}
    .select-wrapper::after {{
        content: "▼";
        position: absolute;
        right: 12px;
        top: 50%;
        transform: translateY(-50%);
        color: #94a3b8;
        font-size: 0.75rem;
        pointer-events: none;
    }}
    button.control-btn {{
        padding: 8px 12px;
        border-radius: 8px;
        background: rgba(30, 41, 59, 0.7);
        color: #f8fafc;
        border: 1px solid rgba(255, 255, 255, 0.15);
        cursor: pointer;
        font-size: 0.9rem;
        font-weight: 500;
        display: flex;
        align-items: center;
        gap: 6px;
        transition: all 0.2s ease;
    }}
    button.control-btn:hover {{
        background: rgba(51, 65, 85, 0.9);
        border-color: rgba(255, 255, 255, 0.3);
    }}
    
    /* Slide Area */
    .slide-area {{
        flex: 1;
        width: 100%;
        display: flex;
        align-items: center;
        justify-content: center;
        position: relative;
        overflow: hidden;
        touch-action: none;
        background: rgba(15, 23, 42, 0.3);
        border: 1px solid rgba(255, 255, 255, 0.05);
        border-radius: 16px;
        min-height: 380px;
    }}
    .slide-wrapper {{
        position: relative;
        max-width: 95%;
        max-height: 90%;
        display: flex;
        align-items: center;
        justify-content: center;
        transition: transform 0.3s cubic-bezier(0.4, 0, 0.2, 1);
    }}
    .slide-img {{
        max-width: 100%;
        max-height: 360px;
        object-fit: contain;
        border-radius: 8px;
        box-shadow: 0 10px 30px rgba(0, 0, 0, 0.5);
        user-select: none;
        pointer-events: none;
        transition: transform 0.3s ease;
    }}
    
    /* Rotating class */
    .rotated {{
        transform: rotate(90deg) scale(0.7);
    }}
    
    /* Desktop Arrow Buttons */
    .nav-arrow {{
        position: absolute;
        top: 50%;
        transform: translateY(-50%);
        width: 40px;
        height: 40px;
        border-radius: 50%;
        background: rgba(15, 23, 42, 0.6);
        border: 1px solid rgba(255, 255, 255, 0.1);
        color: #f8fafc;
        display: flex;
        align-items: center;
        justify-content: center;
        cursor: pointer;
        font-size: 1.1rem;
        z-index: 5;
        transition: all 0.2s ease;
        user-select: none;
    }}
    .nav-arrow:hover {{
        background: rgba(30, 41, 59, 0.85);
        border-color: rgba(255, 255, 255, 0.3);
        transform: translateY(-50%) scale(1.05);
    }}
    .nav-arrow.left-arrow {{ left: 10px; }}
    .nav-arrow.right-arrow {{ right: 10px; }}
    
    /* Hide arrows on mobile */
    @media (max-width: 768px) {{
        .nav-arrow {{
            display: none !important;
        }}
    }}
    
    /* Pagination indicator dots */
    .dots-container {{
        display: flex;
        gap: 6px;
        justify-content: center;
        padding-top: 12px;
        z-index: 2;
    }}
    .dot {{
        width: 8px;
        height: 8px;
        border-radius: 50%;
        background: rgba(255, 255, 255, 0.2);
        cursor: pointer;
        transition: all 0.2s ease;
    }}
    .dot.active {{
        background: #818cf8;
        transform: scale(1.25);
    }}

    /* Fullscreen Specific Styles */
    .viewer-container:fullscreen {{
        background: black !important;
        width: 100vw !important;
        height: 100vh !important;
        padding: 0 !important;
    }}
    .viewer-container:fullscreen .top-bar,
    .viewer-container:fullscreen .dots-container,
    .viewer-container:fullscreen .nav-arrow {{
        display: none !important;
    }}
    .viewer-container:fullscreen .slide-area {{
        width: 100% !important;
        height: 100% !important;
        flex: 1 !important;
        min-height: 0 !important;
        margin-bottom: 0 !important;
        background: black !important;
        border: none !important;
        border-radius: 0 !important;
    }}
    .viewer-container:fullscreen .slide-wrapper {{
        width: 100% !important;
        height: 100% !important;
        max-width: 100% !important;
        max-height: 100% !important;
    }}
    .viewer-container:fullscreen .slide-img {{
        width: 100% !important;
        height: 100% !important;
        max-width: none !important;
        max-height: none !important;
        object-fit: fill; /* Default to fill so it occupies the entire screen space */
        border-radius: 0 !important;
        box-shadow: none !important;
    }}

    .viewer-container:-webkit-full-screen {{
        background: black !important;
        width: 100vw !important;
        height: 100vh !important;
        padding: 0 !important;
    }}
    .viewer-container:-webkit-full-screen .top-bar,
    .viewer-container:-webkit-full-screen .dots-container,
    .viewer-container:-webkit-full-screen .nav-arrow {{
        display: none !important;
    }}
    .viewer-container:-webkit-full-screen .slide-area {{
        width: 100% !important;
        height: 100% !important;
        flex: 1 !important;
        min-height: 0 !important;
        margin-bottom: 0 !important;
        background: black !important;
        border: none !important;
        border-radius: 0 !important;
    }}
    .viewer-container:-webkit-full-screen .slide-wrapper {{
        width: 100% !important;
        height: 100% !important;
        max-width: 100% !important;
        max-height: 100% !important;
    }}
    .viewer-container:-webkit-full-screen .slide-img {{
        width: 100% !important;
        height: 100% !important;
        max-width: none !important;
        max-height: none !important;
        object-fit: fill; /* Default to fill so it occupies the entire screen space */
        border-radius: 0 !important;
        box-shadow: none !important;
    }}
    </style>
    </head>
    <body>
    <div class="viewer-container">
        <div class="top-bar">
            <div class="select-wrapper">
                <select class="slide-select" id="slideSelect">
                </select>
            </div>
            <button class="control-btn" id="rotateBtn">
                <span>🔄 Tilt</span>
            </button>
            <button class="control-btn" id="fullscreenBtn">
                <span>📺 Fullscreen</span>
            </button>
        </div>
        
        <div class="slide-area" id="slideArea">
            <div class="nav-arrow left-arrow" id="prevBtn">❮</div>
            <div class="slide-wrapper" id="slideWrapper">
                <img class="slide-img" id="slideImg" src="" alt="Slide">
            </div>
            <div class="nav-arrow right-arrow" id="nextBtn">❯</div>
        </div>
        
        <div class="dots-container" id="dotsContainer"></div>
        <div id="toast" style="position: absolute; bottom: 80px; background: rgba(15, 23, 42, 0.85); border: 1px solid rgba(255,255,255,0.15); color: #f8fafc; padding: 10px 20px; border-radius: 99px; font-size: 0.9rem; font-weight: 600; opacity: 0; transition: opacity 0.2s ease, transform 0.2s ease; transform: translateY(10px); pointer-events: none; z-index: 100;"></div>
    </div>
    
    <script>
    const slides = {slides_json};
    let currentIdx = 0;
    let isRotated = false;
    
    const slideImg = document.getElementById('slideImg');
    const slideWrapper = document.getElementById('slideWrapper');
    const slideSelect = document.getElementById('slideSelect');
    const dotsContainer = document.getElementById('dotsContainer');
    const prevBtn = document.getElementById('prevBtn');
    const nextBtn = document.getElementById('nextBtn');
    const rotateBtn = document.getElementById('rotateBtn');
    const fullscreenBtn = document.getElementById('fullscreenBtn');
    const slideArea = document.getElementById('slideArea');
    const viewer = document.querySelector('.viewer-container');
    
    slides.forEach((slide, idx) => {{
        const opt = document.createElement('option');
        opt.value = idx;
        opt.textContent = `Slide ${{idx + 1}} of ${{slides.length}}`;
        slideSelect.appendChild(opt);
        
        const dot = document.createElement('div');
        dot.className = `dot ${{idx === 0 ? 'active' : ''}}`;
        dot.onclick = () => showSlide(idx);
        dotsContainer.appendChild(dot);
    }});
    
    const preloadedImages = new Map();
    
    function updatePreloads(currentIdx) {{
        const start = Math.max(0, currentIdx - 4);
        const end = Math.min(slides.length - 1, currentIdx + 5);
        
        for (const [idx, img] of preloadedImages.entries()) {{
            if (idx < start || idx > end) {{
                img.src = '';
                preloadedImages.delete(idx);
            }}
        }}
        
        for (let i = start; i <= end; i++) {{
            if (!preloadedImages.has(i)) {{
                const img = new Image();
                img.src = slides[i];
                preloadedImages.set(i, img);
            }}
        }}
    }}
    
    function showSlide(idx) {{
        if (idx < 0 || idx >= slides.length) return;
        currentIdx = idx;
        slideImg.src = slides[currentIdx];
        slideSelect.value = currentIdx;
        
        const dots = dotsContainer.querySelectorAll('.dot');
        dots.forEach((dot, dIdx) => {{
            dot.className = `dot ${{dIdx === currentIdx ? 'active' : ''}}`;
        }});
        
        updatePreloads(currentIdx);
    }}
    
    function nextSlide() {{
        if (currentIdx < slides.length - 1) showSlide(currentIdx + 1);
    }}
    function prevSlide() {{
        if (currentIdx > 0) showSlide(currentIdx - 1);
    }}
    
    prevBtn.onclick = prevSlide;
    nextBtn.onclick = nextSlide;
    slideSelect.onchange = (e) => showSlide(parseInt(e.target.value));
    
    rotateBtn.onclick = () => {{
        isRotated = !isRotated;
        if (isRotated) {{
            slideImg.classList.add('rotated');
        }} else {{
            slideImg.classList.remove('rotated');
        }}
    }};

    function toggleFullscreen() {{
        if (!document.fullscreenElement && !document.webkitFullscreenElement) {{
            if (viewer.requestFullscreen) {{
                viewer.requestFullscreen();
            }} else if (viewer.webkitRequestFullscreen) {{
                viewer.webkitRequestFullscreen();
            }}
        }} else {{
            if (document.exitFullscreen) {{
                document.exitFullscreen();
            }} else if (document.webkitExitFullscreen) {{
                document.webkitExitFullscreen();
            }}
        }}
    }}
    
    fullscreenBtn.onclick = toggleFullscreen;
    
    function onFullscreenChange() {{
        const isFS = document.fullscreenElement || document.webkitFullscreenElement;
        if (isFS) {{
            fullscreenBtn.querySelector('span').textContent = '🚪 Exit';
        }} else {{
            fullscreenBtn.querySelector('span').textContent = '📺 Fullscreen';
        }}
    }}
    
    document.addEventListener('fullscreenchange', onFullscreenChange);
    document.addEventListener('webkitfullscreenchange', onFullscreenChange);
    
    window.addEventListener('keydown', (e) => {{
        if (e.key === 'ArrowRight') nextSlide();
        if (e.key === 'ArrowLeft') prevSlide();
    }});
    
    const fitModes = ['fill', 'cover', 'contain'];
    let fitIdx = 0;
    const toast = document.getElementById('toast');
    
    function showToast(text) {{
        toast.textContent = text;
        toast.style.opacity = '1';
        toast.style.transform = 'translateY(0)';
        setTimeout(() => {{
            toast.style.opacity = '0';
            toast.style.transform = 'translateY(10px)';
        }}, 1200);
    }}
    
    function cycleFitMode() {{
        fitIdx = (fitIdx + 1) % fitModes.length;
        const mode = fitModes[fitIdx];
        slideImg.style.setProperty('object-fit', mode, 'important');
        showToast(`Fit: ${{mode.toUpperCase()}}`);
    }}

    let touchStartX = 0;
    let touchStartY = 0;
    let touchEndX = 0;
    let touchEndY = 0;
    let lastTap = 0;
    
    viewer.addEventListener('touchstart', (e) => {{
        touchStartX = e.changedTouches[0].screenX;
        touchStartY = e.changedTouches[0].screenY;
    }}, {{passive: true}});
    
    viewer.addEventListener('touchend', (e) => {{
        touchEndX = e.changedTouches[0].screenX;
        touchEndY = e.changedTouches[0].screenY;
        
        const deltaX = touchEndX - touchStartX;
        const deltaY = touchEndY - touchStartY;
        
        if (Math.abs(deltaX) > Math.abs(deltaY) && Math.abs(deltaX) > 50) {{
            if (deltaX < 0) {{
                nextSlide();
            }} else {{
                prevSlide();
            }}
        }} else if (Math.abs(deltaX) < 10 && Math.abs(deltaY) < 10) {{
            const currentTime = new Date().getTime();
            const tapLength = currentTime - lastTap;
            const isClickOnControls = e.target.closest('.top-bar') || e.target.closest('.dots-container') || e.target.closest('.nav-arrow');
            
            if (tapLength < 300 && tapLength > 0) {{
                // Double tap: cycle fit modes!
                if (!isClickOnControls) {{
                    cycleFitMode();
                }}
                e.preventDefault();
            }} else {{
                // Single tap: toggle fullscreen (after a brief delay to check for double tap)
                setTimeout(() => {{
                    const newCurrentTime = new Date().getTime();
                    if (newCurrentTime - lastTap >= 300) {{
                        if (!isClickOnControls) {{
                            toggleFullscreen();
                        }}
                    }}
                }}, 300);
            }}
            lastTap = currentTime;
        }}
    }}, {{passive: false}});
    
    // Also toggle fullscreen on click (desktop) and double click to cycle fit mode
    slideArea.onclick = (e) => {{
        if (e.target.closest('.nav-arrow')) return;
        toggleFullscreen();
    }};
    
    slideArea.ondblclick = (e) => {{
        if (e.target.closest('.nav-arrow')) return;
        cycleFitMode();
    }};
    
    showSlide(0);
    </script>
    </body>
    </html>
    """
    
    import streamlit.components.v1 as components
    components.html(html_code, height=540)


# ── Router: top-level tabs separate Practice, Search, and Test ──
tab_practice, tab_search, tab_test, tab_visuals = st.tabs(["📚 Practice", "🔍 Search", "🎯 Random Test", "🖼️ Visuals"])

with tab_practice:
    screen = st.session_state.screen
    if screen == "home":
        render_home()
    elif screen == "quiz":
        render_quiz()
    elif screen == "summary":
        render_summary()

with tab_search:
    render_search()

with tab_test:
    render_test_tab()

with tab_visuals:
    render_visuals_tab()


# ── Debug Controls (developer-only, hidden from end users) ──
# Visit the app with ?debug=1 in the URL to reveal these.
if st.query_params.get("debug") == "1":
    st.markdown("<br><br><hr style='border:0;border-top:1px solid rgba(255,255,255,0.05)'>", unsafe_allow_html=True)
    col_dbg1, col_dbg2 = st.columns([1, 1])
    with col_dbg1:
        st.session_state.debug_images = st.checkbox("🔍 Debug image/audio details", value=st.session_state.debug_images)
    with col_dbg2:
        if st.button("🧹 Clear cache", key="clear_cache_btn", use_container_width=True):
            st.session_state.tts_cache = {}
            # image caches removed; clear other search helper state
            st.session_state.search_shown_images = set()
            try:
                _build_search_index.clear()
            except Exception:
                pass
            st.success("Session cache cleared!")
            time.sleep(1)
            st.rerun()