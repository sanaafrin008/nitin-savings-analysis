"""Run the whole project:  python main.py

1. loads and cleans the two Excel datasets
2. writes every pivot-style table to outputs/tables/*.csv
3. draws the charts to outputs/charts/*.png
4. builds the formula-driven Excel workbook outputs/Nitin_Expense_Analysis.xlsx
5. prints the key findings
"""
from src import analysis, charts, config, data, workbook


def main() -> None:
    june = data.load_june()
    six = data.load_six_months()
    plan = analysis.savings_plan(six)

    config.TABLES_DIR.mkdir(parents=True, exist_ok=True)
    tables = {
        "p1_category_summary": analysis.category_summary(june),
        "p1_item_summary": analysis.item_summary(june),
        "p2_monthly_totals": analysis.monthly_totals(six),
        "p2_category_totals": analysis.category_totals(six),
        "p2_month_by_category": analysis.month_by_category(six).reset_index(),
        "p2_top2_by_month": analysis.top2_by_month(analysis.month_by_category(six)),
        "p2_items_by_month": analysis.items_by_month(six),
        "p2_essential_split": analysis.essential_split(six),
        "savings_plan": plan.round(0),
    }
    for name, df in tables.items():
        df.to_csv(config.TABLES_DIR / f"{name}.csv", index=False)

    charts.make_all(june, six, plan)
    workbook.build(june, six, config.WORKBOOK_PATH)

    a1, a2, ps = analysis.june_answers(june), analysis.six_month_answers(six), analysis.plan_summary(plan)
    print("\nPART 1 - JUNE")
    print(f"  Spent Rs {a1['total_spent']:,} of Rs {config.INCOME:,} (saved Rs {a1['saved']:,})")
    print(f"  Highest category      : {a1['top_category']} (Rs {a1['top_category_total']:,}, {a1['top_category_share']}%)")
    print(f"  Entertainment+Shopping: Rs {a1['entertainment_plus_shopping']:,}")
    print(f"  Online food           : {a1['online_food_times']} times, Rs {a1['online_food_amount']:,}")
    print(f"  Movies                : {a1['movie_times']} times, Rs {a1['movie_amount']:,}")
    print(f"  Category to remove    : {a1['category_to_remove']} (Rs {a1['category_to_remove_total']:,})")
    print("\nPART 2 - SIX MONTHS")
    print(f"  Total Rs {a2['total_spent']:,}; average Rs {a2['average_month']:,.0f} a month")
    print(f"  Highest month         : {a2['highest_month']} (Rs {a2['highest_month_total']:,})")
    print(f"  Highest / lowest cat. : {a2['highest_category']} Rs {a2['highest_category_total']:,} / "
          f"{a2['lowest_category']} Rs {a2['lowest_category_total']:,}")
    print(f"  Movies peak           : {a2['movie_peak_month']} | Dining out peak: {a2['dining_peak_month']}")
    print("\nSAVINGS PLAN")
    print(f"  Entertainment Rs {ps['entertainment']:,.0f} | Shopping Rs {ps['shopping']:,.0f} | "
          f"Food Rs {ps['food']:,.0f} | Snacks Rs {ps['snacks']:,.0f}")
    print(f"  Total Rs {ps['total_cut_6m']:,.0f} over 6 months (about Rs {ps['extra_per_month']:,.0f} a month)")
    print(f"  Average saving per month: Rs {ps['avg_saved_today']:,.0f} today -> Rs {ps['avg_saved_plan']:,.0f} with the plan")
    print(f"\nOutputs written to: {config.OUT_DIR}")


if __name__ == "__main__":
    main()
