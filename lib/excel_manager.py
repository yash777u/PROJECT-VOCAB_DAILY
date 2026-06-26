"""
Excel manager for DeutschFlash — all read/write operations on vocab Excel files.
_row_idx is stamped from original sheet position before any shuffle so
update_image_url() always writes to the correct Excel row.
"""

import os
import glob
import pandas as pd
from openpyxl import load_workbook

DATA_DIR = os.path.join(os.path.dirname(os.path.dirname(__file__)), "data")


def get_available_levels() -> list[str]:
    """Return all levels found in data/ directory (including empty templates)."""
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


def is_level_empty(level: str) -> bool:
    """Check if a level has zero vocabulary data (only headers or no rows)."""
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