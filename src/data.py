"""Load and clean the two Excel datasets."""
from __future__ import annotations

import pandas as pd

from . import config

# Spelling / spacing fixes found in the raw files.
_ITEM_FIXES = {
    "tshirt and jeans": "T-shirt and jeans",
    "north bengal trip": "North Bengal trip",
    "brother's tution fee": "Brother's tuition fee",
}


def clean_item(value: object) -> str:
    """Normalise whitespace and case so 'Online Food Order' == 'Online food order'."""
    text = " ".join(str(value).split())
    fixed = _ITEM_FIXES.get(text.lower())
    return fixed if fixed else text[:1].upper() + text[1:].lower()


def _tidy(df: pd.DataFrame) -> pd.DataFrame:
    df["Category"] = df["Category"].astype(str).str.strip()
    df["Item"] = df["Item"].map(clean_item)
    df["Expense"] = df["Expense"].astype(int)
    return df


def load_june(path=config.JUNE_FILE) -> pd.DataFrame:
    """Part 1: one row per purchase in June (Date, Category, Item, Expense)."""
    df = pd.read_excel(path, header=1, usecols=range(4))
    df.columns = ["Date", "Category", "Item", "Expense"]
    df = df.dropna(subset=["Category", "Expense"]).copy()
    df["Date"] = pd.to_datetime(df["Date"])
    return _tidy(df).reset_index(drop=True)


def load_six_months(path=config.SIX_MONTH_FILE) -> pd.DataFrame:
    """Part 2: one row per purchase, January to June (Month, Category, Item, Expense)."""
    df = pd.read_excel(path, header=1, usecols=range(4))
    df.columns = ["Month", "Category", "Item", "Expense"]
    df = df.dropna(subset=["Month", "Expense"]).copy()
    df["Month"] = pd.Categorical(df["Month"].str.strip(), categories=config.MONTHS, ordered=True)
    if df["Month"].isna().any():
        raise ValueError("Unexpected month name in the six-month file")
    return _tidy(df).reset_index(drop=True)
