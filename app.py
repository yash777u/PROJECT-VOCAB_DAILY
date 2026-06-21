"""DeutschFlash — Streamlit German Vocabulary Learning App."""
import random
import os
import time
import base64
import streamlit as st
import pandas as pd
from lib.excel_manager import get_available_levels, get_day_info, load_vocab, load_all_vocab, is_level_empty
from lib.image_fetcher import fetch_image_for_row
from lib.tts_service import pronounce_word, slow_word, spell_word, get_audio_paths

# ── Page Config ──
st.set_page_config(page_title="DeutschFlash Master", page_icon="🇩🇪", layout="centered")

# ── Custom CSS ──
st.markdown("""<style>
@import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@300;400;500;600;700;800&display=swap');

html, body, [class*="css"] {
    font-family: 'Plus Jakarta Sans', -apple-system, BlinkMacSystemFont, sans-serif;
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

/* Responsive adjustments */
@media (max-width: 768px) {
    .block-container { padding: 1rem 0.5rem !important; }
    .glass-card { padding: 1.25rem; border-radius: 20px; }
    .search-card { padding: 1.1rem; border-radius: 18px; }
    .german-word { font-size: 1.85rem; }
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
    # Search tab state
    "search_query": "",
    "search_results_cache": None,
    "search_shown_images": set(),
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
    # Always fetch image dynamically on demand (no session caching)
    row_dict = row.to_dict() if hasattr(row, "to_dict") else dict(row)
    # Pass empty session_results so fetcher performs a fresh search
    session_results = {}

    local_path, debug_info = fetch_image_for_row(row_dict, session_results)

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
                deck["audio_pronounce_path"] = [get_audio_paths(level, day_name, row["_row_idx"], row["german_word"])["pronounce"] for _, row in deck.iterrows()]
                deck["audio_slow_path"] = [get_audio_paths(level, day_name, row["_row_idx"], row["german_word"])["slow"] for _, row in deck.iterrows()]
                deck["audio_spell_path"] = [get_audio_paths(level, day_name, row["_row_idx"], row["german_word"])["spell"] for _, row in deck.iterrows()]

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
            deck["audio_pronounce_path"] = [get_audio_paths(level, str(row["_day"]), row["_row_idx"], row["german_word"])["pronounce"] for _, row in deck.iterrows()]
            deck["audio_slow_path"] = [get_audio_paths(level, str(row["_day"]), row["_row_idx"], row["german_word"])["slow"] for _, row in deck.iterrows()]
            deck["audio_spell_path"] = [get_audio_paths(level, str(row["_day"]), row["_row_idx"], row["german_word"])["spell"] for _, row in deck.iterrows()]

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
    st.markdown(f'<div class="german-word">{word}</div>', unsafe_allow_html=True)
    st.markdown(f'<div style="text-align:center"><span class="phonetic">Pronounced: [{phonetic}]</span></div>', unsafe_allow_html=True)
    if note and note != "nan":
        st.markdown(f'<div style="text-align:center;margin-top:0.4rem"><span class="note-text">{note}</span></div>', unsafe_allow_html=True)

    # TTS buttons (kept in one row via .tts-container CSS hook)
    st.markdown('<div class="tts-container"></div>', unsafe_allow_html=True)
    tts_col1, tts_col2, tts_col3 = st.columns(3)
    with tts_col1:
        if st.button("Normal", key=f"tts_pron_{idx}", use_container_width=True):
            st.session_state[f"play_pronounce_{idx}"] = True
    with tts_col2:
        if st.button("Slow", key=f"tts_slow_{idx}", use_container_width=True):
            st.session_state[f"play_slow_{idx}"] = True
    with tts_col3:
        if st.button("Spell", key=f"tts_spell_{idx}", use_container_width=True):
            st.session_state[f"play_spell_{idx}"] = True

    # TTS playback
    auto_key = f"auto_played_{idx}"
    should_auto = not st.session_state.answered and auto_key not in st.session_state
    should_pronounce = st.session_state.get(f"play_pronounce_{idx}", False)
    should_slow = st.session_state.get(f"play_slow_{idx}", False)
    should_spell = st.session_state.get(f"play_spell_{idx}", False)

    if should_auto or should_pronounce or should_slow or should_spell:
        actual_day = str(row.get("_day", st.session_state.day))
        if actual_day == "__all__":
            actual_day = st.session_state.day

        if should_spell:
            mode_type = "spell"
        elif should_slow:
            mode_type = "slow"
        else:
            mode_type = "pronounce"

        paths = get_audio_paths(st.session_state.level, actual_day, int(row["_row_idx"]), word)
        target_path = paths[mode_type]

        if not os.path.exists(target_path) or os.path.getsize(target_path) == 0:
            with st.spinner(f"Generating {mode_type}..."):
                if should_spell:
                    spell_word(word, level=st.session_state.level, day=actual_day, row_idx=int(row["_row_idx"]))
                elif should_slow:
                    slow_word(word, level=st.session_state.level, day=actual_day, row_idx=int(row["_row_idx"]))
                else:
                    pronounce_word(word, level=st.session_state.level, day=actual_day, row_idx=int(row["_row_idx"]))

        if should_spell:
            st.session_state[f"play_spell_{idx}"] = False
        elif should_slow:
            st.session_state[f"play_slow_{idx}"] = False
        else:
            if should_pronounce:
                st.session_state[f"play_pronounce_{idx}"] = False
            if should_auto:
                st.session_state[auto_key] = True

        if os.path.exists(target_path) and os.path.getsize(target_path) > 0:
            with open(target_path, "rb") as f_audio:
                b64_audio = base64.b64encode(f_audio.read()).decode()
            st.markdown(f'<audio autoplay src="data:audio/mp3;base64,{b64_audio}"></audio>', unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # MCQ
    correct = str(row.get("meaning", ""))
    options = [str(row.get(f"option_{i}", "")) for i in range(1, 5)]
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

        example = str(row.get("example_sentence", ""))
        if is_correct:
            st.markdown(f"""<div class="fb-correct">
                <div style="font-weight:700;font-size:0.85rem;letter-spacing:0.1em;text-transform:uppercase;margin-bottom:0.4rem">Richtig! (Correct)</div>
                <div style="font-size:1.05rem">"{word}" = "{correct}"</div>
                <div style="font-size:1rem;margin-top:0.4rem;color:#6ee7b7">Example: {example}</div>
            </div>""", unsafe_allow_html=True)
        else:
            st.markdown(f"""<div class="fb-wrong">
                <div style="font-weight:700;font-size:0.85rem;letter-spacing:0.1em;text-transform:uppercase;margin-bottom:0.4rem">Falsch! (Incorrect)</div>
                <div style="font-size:1.05rem">"{word}" = "{correct}" (not "{selected}")</div>
                <div style="font-size:1rem;margin-top:0.4rem;color:#fca5a5">Example: {example}</div>
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

    st.markdown(f'<div class="german-word">{word}</div>', unsafe_allow_html=True)
    st.markdown(
        f'<div style="text-align:center"><span class="phonetic">Pronounced: [{phonetic}]</span></div>',
        unsafe_allow_html=True,
    )
    if note and note != "nan":
        st.markdown(
            f'<div style="text-align:center;margin-top:0.4rem"><span class="note-text">{note}</span></div>',
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

    # TTS buttons
    st.markdown('<div class="tts-container"></div>', unsafe_allow_html=True)
    tts_col1, tts_col2, tts_col3 = st.columns(3)
    row_idx = int(row["_row_idx"]) if "_row_idx" in row and pd.notna(row["_row_idx"]) else 0
    actual_day = day if day and day != "nan" else "Search"

    with tts_col1:
        if st.button("Normal", key=f"search_tts_pron_{card_key}", use_container_width=True):
            with st.spinner("Generating audio..."):
                pronounce_word(word, level=level, day=actual_day, row_idx=row_idx)
            st.session_state[f"search_play_{card_key}"] = "pronounce"
    with tts_col2:
        if st.button("Slow", key=f"search_tts_slow_{card_key}", use_container_width=True):
            with st.spinner("Generating audio..."):
                slow_word(word, level=level, day=actual_day, row_idx=row_idx)
            st.session_state[f"search_play_{card_key}"] = "slow"
    with tts_col3:
        if st.button("Spell", key=f"search_tts_spell_{card_key}", use_container_width=True):
            with st.spinner("Generating audio..."):
                spell_word(word, level=level, day=actual_day, row_idx=row_idx)
            st.session_state[f"search_play_{card_key}"] = "spell"

    play_mode = st.session_state.get(f"search_play_{card_key}")
    if play_mode:
        paths = get_audio_paths(level, actual_day, row_idx, word)
        target_path = paths.get(play_mode, "")
        if os.path.exists(target_path) and os.path.getsize(target_path) > 0:
            with open(target_path, "rb") as f_audio:
                b64_audio = base64.b64encode(f_audio.read()).decode()
            st.markdown(f'<audio autoplay src="data:audio/mp3;base64,{b64_audio}"></audio>', unsafe_allow_html=True)
        st.session_state[f"search_play_{card_key}"] = None

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


# ── Router: top-level tabs separate Practice from Search ──
tab_practice, tab_search = st.tabs(["📚 Practice", "🔍 Search"])

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