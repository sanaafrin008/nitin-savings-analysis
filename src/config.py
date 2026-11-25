"""Project-wide settings. Change the assumptions here and re-run ``python main.py``."""
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
RAW_DIR = ROOT / "data" / "raw"
JUNE_FILE = RAW_DIR / "part1_june_expenses.xlsx"
SIX_MONTH_FILE = RAW_DIR / "part2_six_month_expenses.xlsx"

OUT_DIR = ROOT / "outputs"
TABLES_DIR = OUT_DIR / "tables"
CHARTS_DIR = OUT_DIR / "charts"
WORKBOOK_PATH = OUT_DIR / "Nitin_Expense_Analysis.xlsx"

# ---- Facts from the problem statement ------------------------------------
INCOME = 15_000  # Nitin's monthly income (INR)
MONTHS = ["January", "February", "March", "April", "May", "June"]

# ---- How categories are grouped (judgement call, see README) -------------
ESSENTIAL = ["Grocery", "Ticket and Bills", "Doctor and Medicine"]
LESS_ESSENTIAL = ["Entertainment", "Shopping", "Food"]
MIXED = ["Miscellaneous"]  # gifts, tuition fee, auto fare: mostly needed

# ---- Savings plan assumptions --------------------------------------------
ENTERTAINMENT_CAP = 500   # max INR per month
SHOPPING_CAP = 1_000      # max INR per month
FOOD_CUT = 0.50           # share of Food spend removed
SNACK_CUT = 0.50          # share of packaged-snack spend removed (items named "Snacks")
