"""Sanity checks: the numbers quoted in the README and the presentation."""
import pytest

from src import analysis, config, data


@pytest.fixture(scope="module")
def june():
    return data.load_june()


@pytest.fixture(scope="module")
def six():
    return data.load_six_months()


def test_row_counts(june, six):
    assert len(june) == 33
    assert len(six) == 101


def test_no_missing_values(june, six):
    assert not june.isna().any().any()
    assert not six.isna().any().any()


def test_june_totals(june):
    assert june["Expense"].sum() == 13_560
    cats = analysis.category_summary(june).set_index("Category")
    assert cats.loc["Grocery", "Total"] == 4_690
    assert cats.loc["Shopping", "Total"] == 3_500
    assert cats["Share_%"].sum() == pytest.approx(100, abs=0.2)


def test_june_answers(june):
    a = analysis.june_answers(june)
    assert a["top_category"] == "Grocery"
    assert a["entertainment_plus_shopping"] == 4_500
    assert (a["online_food_times"], a["online_food_amount"]) == (4, 600)
    assert (a["movie_times"], a["movie_amount"]) == (4, 1_000)
    assert a["category_to_remove"] == "Shopping"
    assert a["saved"] == 1_440


def test_june_matches_the_june_rows_of_part2(june, six):
    """Part 1 (daily data) and Part 2 (June rows) describe the same month."""
    assert six[six["Month"] == "June"]["Expense"].sum() == june["Expense"].sum()


def test_six_month_answers(six):
    a = analysis.six_month_answers(six)
    assert a["total_spent"] == 84_390
    assert a["average_month"] == pytest.approx(14_065)
    assert (a["highest_month"], a["highest_month_total"]) == ("February", 15_620)
    assert (a["lowest_month"], a["lowest_month_total"]) == ("March", 13_140)
    assert (a["highest_category"], a["highest_category_total"]) == ("Grocery", 30_990)
    assert (a["lowest_category"], a["lowest_category_total"]) == ("Doctor and Medicine", 4_000)
    assert a["grocery_first_months"] == 5
    assert (a["movie_peak_month"], a["movie_peak_amount"]) == ("June", 1_000)
    assert (a["dining_peak_month"], a["dining_peak_amount"]) == ("January", 1_000)


def test_top2_by_month(six):
    top = analysis.top2_by_month(analysis.month_by_category(six)).set_index("Month")
    assert top.loc["February", "Highest"] == "Entertainment"
    assert top.loc["April", "Second"] == "Miscellaneous"
    assert top.loc["June", "Second"] == "Shopping"


def test_essential_split_adds_up(six):
    split = analysis.essential_split(six).set_index("Group")["Total"]
    assert split["Essential"] == 51_030
    assert split["Less essential"] == 25_640
    assert split["Miscellaneous (mixed)"] == 7_720
    assert split.sum() == 84_390


def test_savings_plan(six):
    plan = analysis.savings_plan(six)
    s = analysis.plan_summary(plan)
    assert s["entertainment"] == 9_250
    assert s["shopping"] == 4_700
    assert s["food"] == 2_470
    assert s["snacks"] == 2_720
    assert s["total_cut_6m"] == 19_140
    assert s["avg_saved_today"] == pytest.approx(935)
    assert s["avg_saved_plan"] == pytest.approx(4_125)
    assert (plan["Saved_with_plan"] >= plan["Saved_today"]).all()
    assert plan.loc[plan["Month"] == "February", "Saved_today"].iloc[0] == -620


def test_plan_responds_to_assumptions(six):
    stricter = analysis.plan_summary(analysis.savings_plan(six, ent_cap=0, shop_cap=0))
    default = analysis.plan_summary(analysis.savings_plan(six))
    assert stricter["total_cut_6m"] > default["total_cut_6m"]
