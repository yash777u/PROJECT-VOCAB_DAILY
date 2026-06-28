"""
Excel manager for DeutschFlash — data layer for vocabulary files.

All vocabulary lives in ``data/<LEVEL>_vocab.xlsx`` files.  Each workbook
contains one or more sheets (e.g. "Day 1", "Day 2" or topic names like
"Our Classroom").  Every sheet shares the same 12-column schema:

    german_word | pronunciation | meaning | example_sentence |
    option_1..4 | gender | emoji | keyword | note

Key design detail:
    ``_row_idx`` is stamped from the **original sheet index** before any
    shuffle so that audio cache paths and image lookups always resolve
    to the correct Excel row (``Excel row = _row_idx + 2`` accounting
    for the 1-based header row).
"""

import os
import glob
import pandas as pd
from openpyxl import load_workbook
import streamlit as st

DATA_DIR = os.path.join(os.path.dirname(os.path.dirname(__file__)), "data")


@st.cache_data(show_spinner=False)
def get_available_levels() -> list[str]:
    """Discover all vocabulary levels by scanning ``data/*_vocab.xlsx``.

    Returns:
        Sorted list of level names (e.g. ``["A1", "A2", "Movie"]``).
        Skips temporary lock-files created by Excel/LibreOffice.
    """
    pattern = os.path.join(DATA_DIR, "*_vocab.xlsx")
    files = glob.glob(pattern)
    levels = []
    for f in files:
        basename = os.path.basename(f)
        # Skip temp/lock files created by Excel/LibreOffice
        if basename.startswith(".~") or basename.startswith("~$"):
            continue
        level = basename.replace("_vocab.xlsx", "")
        try:
            wb = load_workbook(f, read_only=True)
            wb.close()
            levels.append(level)
        except Exception:
            continue
    return sorted(levels)


@st.cache_data(show_spinner=False)
def is_level_empty(level: str) -> bool:
    """Check whether a level file contains any vocabulary rows.

    Uses pandas to reliably count rows (openpyxl's ``max_row`` can
    return ``None`` in read-only mode for certain Excel files).

    Args:
        level: Level name prefix (e.g. ``"A1"``).

    Returns:
        ``True`` if the file doesn't exist or every sheet has zero data rows.
    """
    filepath = os.path.join(DATA_DIR, f"{level}_vocab.xlsx")
    if not os.path.exists(filepath):
        return True
    try:
        # Use pandas to reliably count rows (max_row can be None in read_only mode)
        all_sheets = pd.read_excel(filepath, sheet_name=None, engine="openpyxl")
        for sheet_name, df in all_sheets.items():
            if len(df) > 0:
                return False
    except Exception:
        pass
    return True


@st.cache_data(show_spinner=False)
def get_day_info(level: str) -> dict:
    filepath = os.path.join(DATA_DIR, f"{level}_vocab.xlsx")
    if not os.path.exists(filepath):
        return {}
    info = {}
    try:
        # Use pandas for reliable row counting (max_row can be None in read_only mode)
        all_sheets = pd.read_excel(filepath, sheet_name=None, engine="openpyxl")
        for sheet_name, df in all_sheets.items():
            count = len(df)
            if count > 0:
                info[sheet_name] = count
    except Exception:
        pass
    return info


@st.cache_data(show_spinner=False)
def load_vocab(level: str, day: str) -> pd.DataFrame:
    filepath = os.path.join(DATA_DIR, f"{level}_vocab.xlsx")
    if not os.path.exists(filepath):
        return pd.DataFrame()
    try:
        df = pd.read_excel(filepath, sheet_name=day, engine="openpyxl")
        # Stamp BEFORE any shuffle — Excel row = _row_idx + 2
        df["_row_idx"] = df.index
        return df
    except Exception:
        return pd.DataFrame()


@st.cache_data(show_spinner=False)
def load_all_vocab(level: str) -> pd.DataFrame:
    filepath = os.path.join(DATA_DIR, f"{level}_vocab.xlsx")
    if not os.path.exists(filepath):
        return pd.DataFrame()
    try:
        all_sheets = pd.read_excel(filepath, sheet_name=None, engine="openpyxl")
        frames = []
        for sheet_name, df in all_sheets.items():
            df = df.copy()
            df["_day"] = sheet_name
            df["_row_idx"] = df.index  # Stamp BEFORE concat+shuffle
            frames.append(df)
        if not frames:
            return pd.DataFrame()
        combined = pd.concat(frames, ignore_index=True)
        return combined
    except Exception:
        return pd.DataFrame()