from __future__ import annotations

import math
import random
from functools import lru_cache
from pathlib import Path
from typing import Any

import pandas as pd
from openpyxl import load_workbook


ROOT = Path(__file__).resolve().parents[3]
DATA_DIR = ROOT / "data"


def _clean(value: Any) -> Any:
    if value is None:
        return ""
    if isinstance(value, float) and math.isnan(value):
        return ""
    return value


def _record(row: pd.Series, level: str, sheet: str) -> dict[str, Any]:
    data = {str(k): _clean(v) for k, v in row.to_dict().items()}
    data["_level"] = level
    data["_day"] = sheet
    data["_row_idx"] = int(data.get("_row_idx", 0) or 0)
    data["id"] = f"{level}::{sheet}::{data['_row_idx']}"
    return data


def _workbook_path(level: str) -> Path:
    return DATA_DIR / f"{level}_vocab.xlsx"


@lru_cache(maxsize=1)
def levels() -> list[str]:
    names: list[str] = []
    for path in DATA_DIR.glob("*_vocab.xlsx"):
        if path.name.startswith((".~", "~$")):
            continue
        try:
            wb = load_workbook(path, read_only=True)
            wb.close()
            names.append(path.name.replace("_vocab.xlsx", ""))
        except Exception:
            continue
    return sorted(names)


@lru_cache(maxsize=64)
def sheets(level: str) -> dict[str, int]:
    path = _workbook_path(level)
    if not path.exists():
        return {}
    try:
        loaded = pd.read_excel(path, sheet_name=None, engine="openpyxl")
    except Exception:
        return {}
    return {name: int(len(df)) for name, df in loaded.items() if len(df) > 0}


@lru_cache(maxsize=128)
def sheet_frame(level: str, sheet: str) -> pd.DataFrame:
    path = _workbook_path(level)
    if not path.exists():
        return pd.DataFrame()
    try:
        df = pd.read_excel(path, sheet_name=sheet, engine="openpyxl")
    except Exception:
        return pd.DataFrame()
    df = df.copy()
    df["_row_idx"] = df.index
    return df


@lru_cache(maxsize=64)
def all_frame(level: str) -> pd.DataFrame:
    path = _workbook_path(level)
    if not path.exists():
        return pd.DataFrame()
    try:
        loaded = pd.read_excel(path, sheet_name=None, engine="openpyxl")
    except Exception:
        return pd.DataFrame()
    frames: list[pd.DataFrame] = []
    for name, df in loaded.items():
        if df.empty:
            continue
        frame = df.copy()
        frame["_day"] = name
        frame["_row_idx"] = frame.index
        frames.append(frame)
    if not frames:
        return pd.DataFrame()
    return pd.concat(frames, ignore_index=True)


def level_summary() -> list[dict[str, Any]]:
    result = []
    for level in levels():
        info = sheets(level)
        result.append(
            {
                "name": level,
                "empty": sum(info.values()) == 0,
                "total": sum(info.values()),
                "sheets": [{"name": name, "count": count} for name, count in info.items()],
            }
        )
    return result


def deck(level: str, sheet: str | None = None, shuffle: bool = True) -> list[dict[str, Any]]:
    if sheet and sheet != "__all__":
        df = sheet_frame(level, sheet)
        records = [_record(row, level, sheet) for _, row in df.iterrows()]
    else:
        df = all_frame(level)
        records = [_record(row, level, str(row.get("_day", ""))) for _, row in df.iterrows()]
    if shuffle:
        random.shuffle(records)
    return records


def search(query: str, limit: int = 25) -> list[dict[str, Any]]:
    q = query.strip().lower()
    if not q:
        return []
    matches: list[dict[str, Any]] = []
    for level in levels():
        df = all_frame(level)
        if df.empty:
            continue
        word_col = df.get("german_word", pd.Series([""] * len(df))).astype(str).str.lower()
        meaning_col = df.get("meaning", pd.Series([""] * len(df))).astype(str).str.lower()
        mask = word_col.str.contains(q, regex=False, na=False) | meaning_col.str.contains(q, regex=False, na=False)
        for _, row in df[mask].iterrows():
            rec = _record(row, level, str(row.get("_day", "")))
            word = str(rec.get("german_word", "")).lower()
            rec["_rank"] = 0 if word == q else 1 if word.startswith(q) else 2
            matches.append(rec)
    matches.sort(key=lambda item: (item.get("_rank", 2), str(item.get("german_word", ""))))
    for item in matches:
        item.pop("_rank", None)
    return matches[:limit]


def test_deck(level: str, sheet: str | None = None, count: int = 25) -> list[dict[str, Any]]:
    source = deck(level, sheet=sheet or "__all__", shuffle=False)
    valid = [
        row
        for row in source
        if str(row.get("example_sentence", "")).strip()
        and str(row.get("example_sentence", "")).strip().lower() != "nan"
    ]
    random.shuffle(valid)
    return valid[: min(count, len(valid))]
