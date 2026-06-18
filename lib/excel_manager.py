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
    pattern = os.path.join(DATA_DIR, "*_vocab.xlsx")
    files = glob.glob(pattern)
    levels = []
    for f in files:
        basename = os.path.basename(f)
        level = basename.replace("_vocab.xlsx", "")
        try:
            wb = load_workbook(f, read_only=True)
            has_data = False
            for ws in wb.worksheets:
                if ws.max_row and ws.max_row > 1:
                    has_data = True
                    break
            wb.close()
            if has_data:
                levels.append(level)
        except Exception:
            continue
    return sorted(levels)


def get_day_info(level: str) -> dict:
    filepath = os.path.join(DATA_DIR, f"{level}_vocab.xlsx")
    if not os.path.exists(filepath):
        return {}
    info = {}
    try:
        wb = load_workbook(filepath, read_only=True)
        for ws in wb.worksheets:
            count = max(0, (ws.max_row or 1) - 1)
            if count > 0:
                info[ws.title] = count
        wb.close()
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