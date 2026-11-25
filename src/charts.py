"""Matplotlib charts used in the README. One function per chart, saved as PNG."""
from __future__ import annotations

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
from matplotlib.patches import Rectangle

from . import analysis, config

TEAL, DEEP, ORANGE, CORAL, GREY, BLUE = "#1DA1A0", "#0F4C4B", "#F2A33A", "#E4572E", "#9AA7AB", "#4472C4"
CATEGORY_COLORS = [TEAL, ORANGE, DEEP, CORAL, "#7FB5B0", GREY, "#B8D8D3"]

plt.rcParams.update({
    "font.family": "DejaVu Sans", "font.size": 11, "axes.spines.top": False,
    "axes.spines.right": False, "axes.edgecolor": "#B0B8BA", "axes.titleweight": "bold",
    "axes.titlesize": 14, "axes.titlecolor": DEEP, "figure.facecolor": "white",
})


def _save(fig, name):
    config.CHARTS_DIR.mkdir(parents=True, exist_ok=True)
    path = config.CHARTS_DIR / name
    fig.savefig(path, dpi=150, bbox_inches="tight")
    plt.close(fig)
    return path


def _inr(x, _pos=None):
    return f"{int(x):,}"


def june_category_share(june):
    cats = analysis.category_summary(june)
    fig, ax = plt.subplots(figsize=(8, 5.2))
    wedges, _ = ax.pie(cats["Total"], colors=CATEGORY_COLORS, startangle=90, counterclock=False,
                       wedgeprops=dict(width=0.42, edgecolor="white", linewidth=2))
    ax.legend(wedges, [f"{c}  ({s:.1f}%)" for c, s in zip(cats["Category"], cats["Share_%"])],
              loc="center left", bbox_to_anchor=(1.0, 0.5), frameon=False)
    ax.text(0, 0.06, f"Rs {cats['Total'].sum():,}", ha="center", fontsize=20, fontweight="bold", color=DEEP)
    ax.text(0, -0.14, "spent in June", ha="center", color=GREY)
    ax.set_title("June: share of total expense by category")
    return _save(fig, "01_june_category_share.png")


def june_items_data_bars(june):
    items = analysis.item_summary(june)
    items = items[items["Category"].isin(["Grocery", "Shopping"])]
    fig, ax = plt.subplots(figsize=(8, 5))
    labels = [f"{r.Item}  ({r.Category})" for r in items.itertuples()]
    colors = [TEAL if c == "Grocery" else ORANGE for c in items["Category"]]
    y = np.arange(len(items))[::-1]
    ax.barh(y, items["Total"], color=colors, height=0.65)
    ax.set_yticks(y, labels)
    for yi, v in zip(y, items["Total"]):
        ax.text(v + 25, yi, f"{v:,}", va="center", fontsize=10)
    ax.xaxis.set_major_formatter(_inr)
    ax.set_xlim(0, items["Total"].max() * 1.15)
    ax.set_title("June: items of the two biggest categories (Grocery, Shopping)")
    ax.set_xlabel("Expense (INR)")
    return _save(fig, "02_june_grocery_shopping_items.png")


def june_entertainment_bills(june):
    items = analysis.item_summary(june)
    sub = items[items["Category"].isin(["Entertainment", "Ticket and Bills"])]
    sub = sub.sort_values(["Category", "Total"], ascending=[True, False])
    fig, ax = plt.subplots(figsize=(8, 4.8))
    colors = [ORANGE if c == "Entertainment" else BLUE for c in sub["Category"]]
    ax.bar(sub["Item"], sub["Total"], color=colors, width=0.6)
    for x, v in enumerate(sub["Total"]):
        ax.text(x, v + 20, f"{v:,}", ha="center", fontsize=10)
    ax.yaxis.set_major_formatter(_inr)
    ax.set_title("June: Entertainment and Ticket and Bills items")
    ax.set_ylabel("Expense (INR)")
    plt.setp(ax.get_xticklabels(), rotation=15, ha="right")
    ax.legend(handles=[Rectangle((0, 0), 1, 1, color=ORANGE), Rectangle((0, 0), 1, 1, color=BLUE)],
              labels=["Entertainment", "Ticket and Bills"], frameon=False)
    return _save(fig, "03_june_entertainment_and_bills.png")


def monthly_trend(six):
    m = analysis.monthly_totals(six)
    fig, ax = plt.subplots(figsize=(8.5, 4.8))
    ax.plot(m["Month"], m["Total"], marker="o", color=BLUE, linewidth=2.5, markersize=8)
    ax.axhline(config.INCOME, color=GREY, linestyle="--", linewidth=1)
    ax.text(5.05, config.INCOME + 60, "Income Rs 15,000", ha="right", color=GREY, fontsize=9)
    peak = m["Total"].idxmax()
    for i, v in enumerate(m["Total"]):
        ax.annotate(f"{v:,}", (i, v), textcoords="offset points", xytext=(0, 12 if i in (1, 3) else -20), ha="center",
                    fontweight="bold" if i == peak else "normal", color=CORAL if i == peak else "#404040")
    ax.set_ylim(12000, 15800)
    ax.yaxis.set_major_formatter(_inr)
    ax.set_title("Month-wise total expense, January to June")
    ax.set_ylabel("Expense (INR)")
    return _save(fig, "04_monthly_trend.png")


def category_totals(six):
    c = analysis.category_totals(six)
    fig, ax = plt.subplots(figsize=(8.5, 4.8))
    colors = [CORAL if i == 0 else (GREY if i == len(c) - 1 else TEAL) for i in range(len(c))]
    y = np.arange(len(c))[::-1]
    ax.barh(y, c["Total"], color=colors, height=0.65)
    ax.set_yticks(y, c["Category"])
    for yi, v, s in zip(y, c["Total"], c["Share_%"]):
        ax.text(v + 300, yi, f"{v:,}  ({s:.1f}%)", va="center", fontsize=10)
    ax.xaxis.set_major_formatter(_inr)
    ax.set_xlim(0, c["Total"].max() * 1.25)
    ax.set_title("Six months: expense by category")
    ax.set_xlabel("Expense (INR)")
    return _save(fig, "05_category_totals.png")


def month_category_heatmap(six):
    pv = analysis.month_by_category(six)
    fig, ax = plt.subplots(figsize=(9.5, 5))
    ax.imshow(pv.values, cmap="GnBu", aspect="auto", vmax=pv.values.max() * 0.9)
    ax.set_xticks(range(6), [m[:3] for m in pv.columns])
    ax.set_yticks(range(len(pv)), pv.index)
    for i in range(pv.shape[0]):
        for j in range(pv.shape[1]):
            v = pv.iat[i, j]
            ax.text(j, i, f"{v:,}" if v else "", ha="center", va="center", fontsize=9,
                    color="white" if v > pv.values.max() * 0.55 else "#1B2A2F")
    for j, month in enumerate(pv.columns):  # outline the top two of each month
        for i in pv[month].reset_index(drop=True).nlargest(2).index:
            ax.add_patch(Rectangle((j - 0.5, i - 0.5), 1, 1, fill=False, edgecolor=ORANGE, linewidth=2.5))
    ax.set_title("Month by category (orange box = top 2 of that month)")
    for s in ax.spines.values():
        s.set_visible(False)
    return _save(fig, "06_month_by_category_heatmap.png")


def essential_split(six):
    e = analysis.essential_split(six)
    fig, ax = plt.subplots(figsize=(7.5, 4.8))
    wedges, _ = ax.pie(e["Total"], colors=[TEAL, ORANGE, GREY], startangle=90, counterclock=False,
                       wedgeprops=dict(width=0.42, edgecolor="white", linewidth=2))
    ax.legend(wedges, [f"{g}: Rs {t:,} ({s:.1f}%)" for g, t, s in zip(e["Group"], e["Total"], e["Share_%"])],
              loc="center left", bbox_to_anchor=(1.0, 0.5), frameon=False)
    ax.set_title("Essential vs less essential spending (6 months)")
    return _save(fig, "07_essential_vs_less_essential.png")


def savings_today_vs_plan(plan):
    x = np.arange(len(plan))
    fig, ax = plt.subplots(figsize=(9, 5))
    ax.bar(x - 0.2, plan["Saved_today"], 0.4, color=BLUE, label="Saved today")
    ax.bar(x + 0.2, plan["Saved_with_plan"], 0.4, color=ORANGE, label="Saved with the plan")
    for xi, a, b in zip(x, plan["Saved_today"], plan["Saved_with_plan"]):
        ax.text(xi - 0.2, a + (150 if a >= 0 else -450), f"{a:,.0f}", ha="center", fontsize=9)
        ax.text(xi + 0.2, b + 150, f"{b:,.0f}", ha="center", fontsize=9)
    ax.axhline(0, color="#707070", linewidth=0.8)
    ax.set_xticks(x, [m[:3] for m in plan["Month"]])
    ax.yaxis.set_major_formatter(_inr)
    ax.set_title("Money left each month: today vs with the savings plan")
    ax.set_ylabel("INR")
    ax.legend(frameon=False)
    return _save(fig, "08_savings_today_vs_plan.png")


def make_all(june, six, plan):
    return [
        june_category_share(june), june_items_data_bars(june), june_entertainment_bills(june),
        monthly_trend(six), category_totals(six), month_category_heatmap(six),
        essential_split(six), savings_today_vs_plan(plan),
    ]
