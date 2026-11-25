"""All the calculations: the pivot-table logic of the project, written in pandas."""
from __future__ import annotations

import pandas as pd

from . import config


# --------------------------------------------------------------------------
# Part 1: June
# --------------------------------------------------------------------------
def category_summary(df: pd.DataFrame) -> pd.DataFrame:
    """Spend, number of purchases and % of total for each category (sorted high to low)."""
    out = (
        df.groupby("Category")["Expense"]
        .agg(Total="sum", Purchases="count")
        .sort_values("Total", ascending=False)
        .reset_index()
    )
    out["Share_%"] = (out["Total"] / out["Total"].sum() * 100).round(1)
    return out


def item_summary(df: pd.DataFrame) -> pd.DataFrame:
    """Spend and number of times spent on each item, grouped by category."""
    order = category_summary(df)["Category"].tolist()
    out = df.groupby(["Category", "Item"])["Expense"].agg(Total="sum", Times="count").reset_index()
    out["_rank"] = out["Category"].map({c: i for i, c in enumerate(order)})
    out = out.sort_values(["_rank", "Total"], ascending=[True, False]).drop(columns="_rank")
    return out.reset_index(drop=True)


def june_answers(df: pd.DataFrame) -> dict:
    """The five questions the project asks for Part 1."""
    cats = category_summary(df).set_index("Category")
    food = df[df["Item"] == "Online food order"]
    movies = df[df["Item"] == "Movie"]
    less_essential = cats.loc[[c for c in config.LESS_ESSENTIAL if c in cats.index]]
    return {
        "total_spent": int(df["Expense"].sum()),
        "saved": config.INCOME - int(df["Expense"].sum()),
        "top_category": cats.index[0],
        "top_category_total": int(cats.iloc[0]["Total"]),
        "top_category_share": float(cats.iloc[0]["Share_%"]),
        "entertainment_plus_shopping": int(cats.loc[["Entertainment", "Shopping"], "Total"].sum()),
        "online_food_times": int(len(food)),
        "online_food_amount": int(food["Expense"].sum()),
        "movie_times": int(len(movies)),
        "movie_amount": int(movies["Expense"].sum()),
        # largest of the less-essential categories = the one to drop first
        "category_to_remove": less_essential["Total"].idxmax(),
        "category_to_remove_total": int(less_essential["Total"].max()),
    }


# --------------------------------------------------------------------------
# Part 2: six months
# --------------------------------------------------------------------------
def monthly_totals(df: pd.DataFrame) -> pd.DataFrame:
    out = df.groupby("Month", observed=False)["Expense"].sum().rename("Total").reset_index()
    out["Saved_today"] = config.INCOME - out["Total"]
    return out


def category_totals(df: pd.DataFrame) -> pd.DataFrame:
    out = df.groupby("Category")["Expense"].sum().sort_values(ascending=False).rename("Total").reset_index()
    out["Share_%"] = (out["Total"] / out["Total"].sum() * 100).round(1)
    return out


def month_by_category(df: pd.DataFrame) -> pd.DataFrame:
    """Rows = category (sorted by total), columns = month, values = INR."""
    order = category_totals(df)["Category"].tolist()
    pt = df.pivot_table(index="Category", columns="Month", values="Expense", aggfunc="sum",
                        fill_value=0, observed=False)
    return pt.loc[order, config.MONTHS]


def top2_by_month(pivot: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for month in pivot.columns:
        ranked = pivot[month].sort_values(ascending=False)
        rows.append({"Month": month,
                     "Highest": ranked.index[0], "Highest_INR": int(ranked.iloc[0]),
                     "Second": ranked.index[1], "Second_INR": int(ranked.iloc[1])})
    return pd.DataFrame(rows)


def items_by_month(df: pd.DataFrame, categories=("Entertainment", "Food", "Shopping")) -> pd.DataFrame:
    sub = df[df["Category"].isin(categories)]
    pt = sub.pivot_table(index=["Category", "Item"], columns="Month", values="Expense",
                         aggfunc="sum", fill_value=0, observed=False)[config.MONTHS]
    pt["Total"] = pt.sum(axis=1)
    pt["Peak_month"] = pt[config.MONTHS].idxmax(axis=1)
    return pt.reset_index()


def essential_split(df: pd.DataFrame) -> pd.DataFrame:
    totals = category_totals(df).set_index("Category")["Total"]
    groups = {
        "Essential": config.ESSENTIAL,
        "Less essential": config.LESS_ESSENTIAL,
        "Miscellaneous (mixed)": config.MIXED,
    }
    out = pd.DataFrame({"Group": list(groups),
                        "Total": [int(totals.reindex(v).sum()) for v in groups.values()]})
    out["Share_%"] = (out["Total"] / out["Total"].sum() * 100).round(1)
    return out


def six_month_answers(df: pd.DataFrame) -> dict:
    months = monthly_totals(df)
    cats = category_totals(df)
    pivot = month_by_category(df)
    movies = pivot_item_peak(df, "Movie")
    dining = pivot_item_peak(df, "Dining out")
    return {
        "total_spent": int(df["Expense"].sum()),
        "average_month": float(months["Total"].mean()),
        "highest_month": months.loc[months["Total"].idxmax(), "Month"],
        "highest_month_total": int(months["Total"].max()),
        "lowest_month": months.loc[months["Total"].idxmin(), "Month"],
        "lowest_month_total": int(months["Total"].min()),
        "highest_category": cats.iloc[0]["Category"],
        "highest_category_total": int(cats.iloc[0]["Total"]),
        "lowest_category": cats.iloc[-1]["Category"],
        "lowest_category_total": int(cats.iloc[-1]["Total"]),
        "grocery_first_months": int((pivot.idxmax() == "Grocery").sum()),
        "movie_peak_month": movies[0], "movie_peak_amount": movies[1],
        "dining_peak_month": dining[0], "dining_peak_amount": dining[1],
        "snacks_total": int(df.loc[df["Item"] == "Snacks", "Expense"].sum()),
    }


def pivot_item_peak(df: pd.DataFrame, item: str) -> tuple[str, int]:
    s = df[df["Item"] == item].groupby("Month", observed=True)["Expense"].sum()
    return str(s.idxmax()), int(s.max())


# --------------------------------------------------------------------------
# Recommendations: the savings plan
# --------------------------------------------------------------------------
def savings_plan(
    df: pd.DataFrame,
    ent_cap: float = config.ENTERTAINMENT_CAP,
    shop_cap: float = config.SHOPPING_CAP,
    food_cut: float = config.FOOD_CUT,
    snack_cut: float = config.SNACK_CUT,
    income: float = config.INCOME,
) -> pd.DataFrame:
    """Month by month: what each cut saves and the saving with the plan.

    Caps are applied month by month, so a month already under the cap saves nothing.
    """
    pivot = month_by_category(df)
    snacks = df[df["Item"] == "Snacks"].groupby("Month", observed=False)["Expense"].sum()
    rows = []
    for month in config.MONTHS:
        spent = int(pivot[month].sum())
        cut_ent = max(0, pivot.loc["Entertainment", month] - ent_cap)
        cut_shop = max(0, pivot.loc["Shopping", month] - shop_cap)
        cut_food = pivot.loc["Food", month] * food_cut
        cut_snack = snacks.get(month, 0) * snack_cut
        total_cut = cut_ent + cut_shop + cut_food + cut_snack
        rows.append({
            "Month": month, "Spent": spent, "Saved_today": income - spent,
            "Cut_Entertainment": cut_ent, "Cut_Shopping": cut_shop,
            "Cut_Food": cut_food, "Cut_Snacks": cut_snack,
            "Total_cut": total_cut, "Saved_with_plan": income - spent + total_cut,
        })
    return pd.DataFrame(rows)


def plan_summary(plan: pd.DataFrame) -> dict:
    return {
        "entertainment": float(plan["Cut_Entertainment"].sum()),
        "shopping": float(plan["Cut_Shopping"].sum()),
        "food": float(plan["Cut_Food"].sum()),
        "snacks": float(plan["Cut_Snacks"].sum()),
        "total_cut_6m": float(plan["Total_cut"].sum()),
        "extra_per_month": float(plan["Total_cut"].sum() / len(plan)),
        "saved_today_6m": float(plan["Saved_today"].sum()),
        "saved_plan_6m": float(plan["Saved_with_plan"].sum()),
        "avg_saved_today": float(plan["Saved_today"].mean()),
        "avg_saved_plan": float(plan["Saved_with_plan"].mean()),
    }
