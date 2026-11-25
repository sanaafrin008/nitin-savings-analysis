"""Build the Excel workbook: the same analysis as the pandas code, but with live formulas.

Every summary table uses SUMIFS / COUNTIFS on the data sheets, so editing a number or an
assumption on the ``Inputs`` sheet updates every table, chart and highlight.
"""
from __future__ import annotations

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.series import DataPoint
from openpyxl.formatting.rule import DataBarRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.workbook.defined_name import DefinedName

from . import analysis, config

TEAL, DEEP, LIGHT = "1DA1A0", "0F4C4B", "E3F4F3"
FILL_HDR = PatternFill("solid", fgColor=TEAL)
FILL_TOTAL = PatternFill("solid", fgColor="DDEBF7")
FILL_INPUT = PatternFill("solid", fgColor="FFF2CC")
THIN = Side(style="thin", color="D9D9D9")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
PALETTE = ["4472C4", "ED7D31", "A5A5A5", "FFC000", "5B9BD5", "70AD47", "264478"]
NUM = "#,##0"


# ------------------------------------------------------------------ helpers
def _val_labels():
    """Data labels that show only the value (no series or category name)."""
    return DataLabelList(showVal=True, showSerName=False, showCatName=False, showLegendKey=False, showPercent=False)


def _title(ws, title, subtitle=""):
    ws.sheet_view.showGridLines = False
    ws["A1"] = title
    ws["A1"].font = Font(name="Calibri", size=16, bold=True, color=DEEP)
    ws["A2"] = subtitle
    ws["A2"].font = Font(name="Calibri", size=10, italic=True, color="707070")


def _header(ws, row, labels, col=1):
    for i, text in enumerate(labels):
        c = ws.cell(row=row, column=col + i, value=text)
        c.font = Font(name="Calibri", bold=True, color="FFFFFF")
        c.fill = FILL_HDR
        c.alignment = Alignment(horizontal="center" if i else "left", vertical="center", wrap_text=True)
        c.border = BORDER


def _cell(ws, row, col, value, fmt=None, bold=False, fill=None, align=None):
    c = ws.cell(row=row, column=col, value=value)
    c.font = Font(name="Calibri", bold=bold)
    c.border = BORDER
    if fmt:
        c.number_format = fmt
    if fill:
        c.fill = fill
    if align:
        c.alignment = Alignment(horizontal=align)
    return c


def _widths(ws, widths):
    for letter, w in widths.items():
        ws.column_dimensions[letter].width = w


def _databar(ws, rng, color="638EC6"):
    ws.conditional_formatting.add(
        rng, DataBarRule(start_type="num", start_value=0, end_type="max", color=color, showValue=True))


def _name(wb, name, ref):
    wb.defined_names[name] = DefinedName(name, attr_text=ref)


def _data_sheet(wb, title, frame, cols, widths):
    ws = wb.create_sheet(title)
    _header(ws, 1, cols)
    for r, row in enumerate(frame.itertuples(index=False), start=2):
        for c, v in enumerate(row, start=1):
            v = v.to_pydatetime() if hasattr(v, "to_pydatetime") else v
            cell = ws.cell(row=r, column=c, value=v if not hasattr(v, "item") else v.item())
            cell.border = BORDER
            if c == len(cols):
                cell.number_format = NUM
            if cols[c - 1] == "Date":
                cell.number_format = "dd-mmm-yyyy"
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{chr(64 + len(cols))}{len(frame) + 1}"
    _widths(ws, widths)
    return ws, len(frame) + 1


# ------------------------------------------------------------------ build
def build(june, six, path):
    wb = Workbook()
    wb.remove(wb.active)
    summary = wb.create_sheet("Summary")

    # ---- Inputs --------------------------------------------------------
    inp = wb.create_sheet("Inputs")
    _title(inp, "Inputs and assumptions", "Yellow cells can be changed. Every sheet recalculates.")
    rows = [("Monthly income (INR)", config.INCOME, NUM), ("Entertainment cap per month (INR)", config.ENTERTAINMENT_CAP, NUM),
            ("Shopping cap per month (INR)", config.SHOPPING_CAP, NUM), ("Share of Food spend cut", config.FOOD_CUT, "0%"),
            ("Share of packaged-snack spend cut", config.SNACK_CUT, "0%")]
    for i, (label, value, fmt) in enumerate(rows, start=3):
        _cell(inp, i, 1, label, bold=True)
        c = _cell(inp, i, 2, value, fmt=fmt, fill=FILL_INPUT)
        c.font = Font(name="Calibri", color="0000FF", bold=True)
    _header(inp, 11, ["Category", "Group"])
    groups = ([(c, "Essential") for c in config.ESSENTIAL] + [(c, "Less essential") for c in config.LESS_ESSENTIAL]
              + [(c, "Miscellaneous (mixed)") for c in config.MIXED])
    for i, (cat, grp) in enumerate(groups, start=12):
        _cell(inp, i, 1, cat)
        c = _cell(inp, i, 2, grp, fill=FILL_INPUT)
        c.font = Font(name="Calibri", color="0000FF")
    _widths(inp, {"A": 38, "B": 24})

    # ---- raw data ------------------------------------------------------
    jws, jn = _data_sheet(wb, "June_Data", june, ["Date", "Category", "Item", "Expense (INR)"],
                          {"A": 14, "B": 22, "C": 28, "D": 15})
    sws, sn = _data_sheet(wb, "SixMonth_Data", six.assign(Month=six["Month"].astype(str)),
                          ["Month", "Category", "Item", "Expense (INR)"], {"A": 14, "B": 22, "C": 28, "D": 15})
    _name(wb, "JUNE_CAT", f"June_Data!$B$2:$B${jn}")
    _name(wb, "JUNE_ITEM", f"June_Data!$C$2:$C${jn}")
    _name(wb, "JUNE_AMT", f"June_Data!$D$2:$D${jn}")
    _name(wb, "SIX_MONTH", f"SixMonth_Data!$A$2:$A${sn}")
    _name(wb, "SIX_CAT", f"SixMonth_Data!$B$2:$B${sn}")
    _name(wb, "SIX_ITEM", f"SixMonth_Data!$C$2:$C${sn}")
    _name(wb, "SIX_AMT", f"SixMonth_Data!$D$2:$D${sn}")

    # ---- Part 1: June_Category (pivot 1 + pie) ---------------------------
    cats = analysis.category_summary(june)["Category"].tolist()
    ws = wb.create_sheet("June_Category")
    _title(ws, "June: spend by category", "Pivot-style table (SUMIFS) and chart of each category's share of the total")
    _header(ws, 4, ["Category", "Sum of Expense (INR)", "Purchases", "% of total"])
    first, last = 5, 4 + len(cats)
    for i, cat in enumerate(cats, start=first):
        _cell(ws, i, 1, cat)
        _cell(ws, i, 2, f"=SUMIFS(JUNE_AMT,JUNE_CAT,$A{i})", NUM)
        _cell(ws, i, 3, f"=COUNTIFS(JUNE_CAT,$A{i})", NUM)
        _cell(ws, i, 4, f"=B{i}/$B${last + 1}", "0.0%")
    t = last + 1
    for col, f, fmt in [(1, "Grand Total", None), (2, f"=SUM(B{first}:B{last})", NUM), (3, f"=SUM(C{first}:C{last})", NUM), (4, f"=SUM(D{first}:D{last})", "0.0%")]:
        _cell(ws, t, col, f, fmt, bold=True, fill=FILL_TOTAL)
    _databar(ws, f"B{first}:B{last}")
    pie = PieChart()
    pie.title = "Share of total expense, June"
    pie.add_data(Reference(ws, min_col=2, min_row=4, max_row=last), titles_from_data=True)
    pie.set_categories(Reference(ws, min_col=1, min_row=first, max_row=last))
    pie.dataLabels = DataLabelList(showPercent=True, showVal=False, showCatName=False, showSerName=False, showLeaderLines=False)
    for i in range(len(cats)):
        pt = DataPoint(idx=i)
        pt.graphicalProperties.solidFill = PALETTE[i % 7]
        pie.series[0].dPt.append(pt)
    pie.height, pie.width = 9, 14
    ws.add_chart(pie, "F4")
    _widths(ws, {"A": 24, "B": 22, "C": 12, "D": 12})
    cat_rng = (first, last)

    # ---- Part 1: June_Items -------------------------------------------
    items = analysis.item_summary(june)
    ws = wb.create_sheet("June_Items")
    _title(ws, "June: items, number of times spent and amount",
           "Use the filter on the Category column to show only Grocery and Shopping. Data bars show the amount.")
    _header(ws, 4, ["Category", "Item", "Times spent", "Sum of Expense (INR)"])
    f1 = 5
    for i, r in enumerate(items.itertuples(index=False), start=f1):
        _cell(ws, i, 1, r.Category)
        _cell(ws, i, 2, r.Item)
        _cell(ws, i, 3, f"=COUNTIFS(JUNE_CAT,$A{i},JUNE_ITEM,$B{i})", NUM)
        _cell(ws, i, 4, f"=SUMIFS(JUNE_AMT,JUNE_CAT,$A{i},JUNE_ITEM,$B{i})", NUM)
    l1 = f1 + len(items) - 1
    for col, f, fmt in [(1, "Grand Total", None), (2, None, None), (3, f"=SUM(C{f1}:C{l1})", NUM), (4, f"=SUM(D{f1}:D{l1})", NUM)]:
        _cell(ws, l1 + 1, col, f, fmt, bold=True, fill=FILL_TOTAL)
    ws.auto_filter.ref = f"A4:D{l1}"
    _databar(ws, f"D{f1}:D{l1}")
    # chart block: Entertainment and Ticket and Bills items
    _header(ws, 4, ["Item", "Entertainment", "Ticket and Bills"], col=6)
    sel = items[items["Category"].isin(["Entertainment", "Ticket and Bills"])].sort_values(["Category", "Total"], ascending=[True, False])
    for i, name in enumerate(sel["Item"], start=5):
        _cell(ws, i, 6, name)
        _cell(ws, i, 7, f'=SUMIFS(JUNE_AMT,JUNE_CAT,"Entertainment",JUNE_ITEM,$F{i})', "#,##0;;")
        _cell(ws, i, 8, f'=SUMIFS(JUNE_AMT,JUNE_CAT,"Ticket and Bills",JUNE_ITEM,$F{i})', "#,##0;;")
    bar = BarChart()
    bar.type, bar.grouping, bar.overlap = "col", "clustered", 100
    bar.title = "Entertainment and Ticket and Bills items (INR)"
    bar.add_data(Reference(ws, min_col=7, max_col=8, min_row=4, max_row=4 + len(sel)), titles_from_data=True)
    bar.set_categories(Reference(ws, min_col=6, min_row=5, max_row=4 + len(sel)))
    bar.dataLabels = _val_labels()
    bar.dataLabels.numFmt = "#,##0;;"
    bar.series[0].graphicalProperties.solidFill = "ED7D31"
    bar.series[1].graphicalProperties.solidFill = "4472C4"
    bar.height, bar.width = 8.5, 16
    ws.add_chart(bar, "F12")
    _widths(ws, {"A": 22, "B": 26, "C": 13, "D": 22, "E": 3, "F": 24, "G": 15, "H": 17})

    # ---- Part 1: findings ------------------------------------------------
    ws = wb.create_sheet("June_Findings")
    _title(ws, "June: answers to the five questions", "All values are formulas on the June data")
    _header(ws, 4, ["Question", "Answer", "Amount (INR)"])
    c1, c2 = cat_rng
    qs = [
        ("1. Category with the highest expense", f"=INDEX(June_Category!A{c1}:A{c2},MATCH(MAX(June_Category!B{c1}:B{c2}),June_Category!B{c1}:B{c2},0))", f"=MAX(June_Category!B{c1}:B{c2})"),
        ("2. Total expense against Entertainment and Shopping", "Entertainment + Shopping", '=SUMIFS(JUNE_AMT,JUNE_CAT,"Entertainment")+SUMIFS(JUNE_AMT,JUNE_CAT,"Shopping")'),
        ("3. Times food was ordered online (and amount)", '=COUNTIFS(JUNE_ITEM,"Online food order")&" times"', '=SUMIFS(JUNE_AMT,JUNE_ITEM,"Online food order")'),
        ("4. Times Nitin watched a movie (and amount)", '=COUNTIFS(JUNE_ITEM,"Movie")&" times"', '=SUMIFS(JUNE_AMT,JUNE_ITEM,"Movie")'),
        ("5. Less essential category he may remove", "=INDEX(A12:A14,MATCH(MAX(B12:B14),B12:B14,0))", "=MAX(B12:B14)"),
    ]
    for i, (q, a, v) in enumerate(qs, start=5):
        _cell(ws, i, 1, q, bold=True)
        _cell(ws, i, 2, a)
        _cell(ws, i, 3, v, NUM)
    _header(ws, 11, ["Less essential categories", "Sum (INR)"])
    for i, cat in enumerate(config.LESS_ESSENTIAL, start=12):
        _cell(ws, i, 1, cat)
        _cell(ws, i, 2, f"=SUMIFS(JUNE_AMT,JUNE_CAT,$A{i})", NUM)
    _cell(ws, 16, 1, "June total spent", bold=True)
    _cell(ws, 16, 3, f"=June_Category!B{c2 + 1}", NUM)
    _cell(ws, 17, 1, "Saved in June (income - spent)", bold=True)
    _cell(ws, 17, 3, "=Inputs!B3-C16", NUM)
    _widths(ws, {"A": 52, "B": 28, "C": 16})

    # ---- Part 2: Monthly_Trend ------------------------------------------
    ws = wb.create_sheet("Monthly_Trend")
    _title(ws, "Six months: month-wise trend of expenses", "The highest month is highlighted")
    _header(ws, 4, ["Month", "Sum of Expense (INR)", "Income (INR)", "Saved today (INR)"])
    for i, m in enumerate(config.MONTHS, start=5):
        _cell(ws, i, 1, m)
        _cell(ws, i, 2, f"=SUMIFS(SIX_AMT,SIX_MONTH,$A{i})", NUM)
        _cell(ws, i, 3, "=Inputs!$B$3", NUM)
        _cell(ws, i, 4, f"=C{i}-B{i}", NUM)
    for col, f in [(1, "Total"), (2, "=SUM(B5:B10)"), (3, "=SUM(C5:C10)"), (4, "=SUM(D5:D10)")]:
        _cell(ws, 11, col, f, NUM if col > 1 else None, bold=True, fill=FILL_TOTAL)
    _cell(ws, 12, 1, "Average per month", bold=True)
    _cell(ws, 12, 2, "=AVERAGE(B5:B10)", NUM)
    _cell(ws, 12, 4, "=AVERAGE(D5:D10)", NUM)
    _cell(ws, 14, 1, "Month with the highest expense", bold=True)
    _cell(ws, 14, 2, "=INDEX(A5:A10,MATCH(MAX(B5:B10),B5:B10,0))", bold=True)
    _cell(ws, 15, 1, "Month with the lowest expense", bold=True)
    _cell(ws, 15, 2, "=INDEX(A5:A10,MATCH(MIN(B5:B10),B5:B10,0))")
    ws.conditional_formatting.add("A5:D10", FormulaRule(formula=["$B5=MAX($B$5:$B$10)"], fill=PatternFill("solid", bgColor="FFEB9C", fgColor="FFEB9C")))
    line = LineChart()
    line.title = "Month-wise total expense (INR)"
    line.add_data(Reference(ws, min_col=2, min_row=4, max_row=10), titles_from_data=True)
    line.set_categories(Reference(ws, min_col=1, min_row=5, max_row=10))
    s = line.series[0]
    s.smooth = False
    s.marker.symbol, s.marker.size = "circle", 8
    s.graphicalProperties.line.solidFill, s.graphicalProperties.line.width = "4472C4", 28000
    s.marker.graphicalProperties.solidFill = "4472C4"
    line.dataLabels = _val_labels()
    line.legend = None
    line.y_axis.scaling.min = 12000
    line.height, line.width = 8.5, 16
    ws.add_chart(line, "F4")
    _widths(ws, {"A": 32, "B": 22, "C": 15, "D": 18})

    # ---- Part 2: Category_Totals ----------------------------------------
    ctot = analysis.category_totals(six)["Category"].tolist()
    ws = wb.create_sheet("Category_Totals")
    _title(ws, "Six months: category-wise expenses", "Data bars show size; green = highest category, red = lowest")
    _header(ws, 4, ["Category", "Sum of Expense (INR)", "% of total", "Group"])
    g1, g2 = 5, 4 + len(ctot)
    for i, cat in enumerate(ctot, start=g1):
        _cell(ws, i, 1, cat)
        _cell(ws, i, 2, f"=SUMIFS(SIX_AMT,SIX_CAT,$A{i})", NUM)
        _cell(ws, i, 3, f"=B{i}/$B${g2 + 1}", "0.0%")
        _cell(ws, i, 4, f"=VLOOKUP($A{i},Inputs!$A$12:$B$18,2,FALSE)")
    for col, f, fmt in [(1, "Grand Total", None), (2, f"=SUM(B{g1}:B{g2})", NUM), (3, f"=SUM(C{g1}:C{g2})", "0.0%")]:
        _cell(ws, g2 + 1, col, f, fmt, bold=True, fill=FILL_TOTAL)
    _databar(ws, f"B{g1}:B{g2}")
    ws.conditional_formatting.add(f"A{g1}:A{g2}", FormulaRule(formula=[f"$B{g1}=MAX($B${g1}:$B${g2})"], fill=PatternFill("solid", bgColor="C6EFCE", fgColor="C6EFCE"), font=Font(color="006100", bold=True)))
    ws.conditional_formatting.add(f"A{g1}:A{g2}", FormulaRule(formula=[f"$B{g1}=MIN($B${g1}:$B${g2})"], fill=PatternFill("solid", bgColor="FFC7CE", fgColor="FFC7CE"), font=Font(color="9C0006", bold=True)))
    bc = BarChart()
    bc.type, bc.title, bc.legend = "bar", "Expense by category (INR)", None
    bc.add_data(Reference(ws, min_col=2, min_row=4, max_row=g2), titles_from_data=True)
    bc.set_categories(Reference(ws, min_col=1, min_row=g1, max_row=g2))
    bc.y_axis.majorGridlines = None
    bc.x_axis.scaling.orientation = "maxMin"
    bc.dataLabels = _val_labels()
    bc.series[0].graphicalProperties.solidFill = "4472C4"
    bc.height, bc.width = 8.5, 16
    ws.add_chart(bc, "G4")
    _widths(ws, {"A": 24, "B": 22, "C": 12, "D": 24})

    # ---- Part 2: Month_x_Category ---------------------------------------
    ws = wb.create_sheet("Month_x_Category")
    _title(ws, "Six months: month-wise expense of each category", "Yellow = the two highest categories of that month")
    _header(ws, 4, ["Category"] + config.MONTHS + ["Grand Total"])
    m1, m2 = 5, 4 + len(ctot)
    for i, cat in enumerate(ctot, start=m1):
        _cell(ws, i, 1, cat)
        for j in range(6):
            col = chr(66 + j)
            _cell(ws, i, 2 + j, f"=SUMIFS(SIX_AMT,SIX_CAT,$A{i},SIX_MONTH,{col}$4)", NUM)
        _cell(ws, i, 8, f"=SUM(B{i}:G{i})", NUM, bold=True)
    _cell(ws, m2 + 1, 1, "Grand Total", bold=True, fill=FILL_TOTAL)
    for j in range(7):
        col = chr(66 + j)
        _cell(ws, m2 + 1, 2 + j, f"=SUM({col}{m1}:{col}{m2})", NUM, bold=True, fill=FILL_TOTAL)
    ws.conditional_formatting.add(f"B{m1}:G{m2}", FormulaRule(formula=[f"B{m1}>=LARGE(B${m1}:B${m2},2)"], fill=PatternFill("solid", bgColor="FFEB9C", fgColor="FFEB9C"), font=Font(color="9C5700", bold=True)))
    r0 = m2 + 4
    _header(ws, r0, ["Top 2 categories"] + config.MONTHS)
    for k, label in enumerate(["Highest", "2nd highest"], start=1):
        _cell(ws, r0 + k, 1, label, bold=True)
        for j in range(6):
            col = chr(66 + j)
            _cell(ws, r0 + k, 2 + j, f"=INDEX($A${m1}:$A${m2},MATCH(LARGE({col}${m1}:{col}${m2},{k}),{col}${m1}:{col}${m2},0))")
    _widths(ws, {"A": 24, **{chr(66 + j): 15 for j in range(7)}})

    # ---- Part 2: Items_EFS -----------------------------------------------
    efs = analysis.items_by_month(six)
    ws = wb.create_sheet("Items_EFS")
    _title(ws, "Six months: items of Entertainment, Food and Shopping by month", "Yellow = the month each item peaks (movies: June, dining out: January)")
    _header(ws, 4, ["Category", "Item"] + config.MONTHS + ["Total", "Peak month"])
    e1 = 5
    for i, r in enumerate(efs.itertuples(index=False), start=e1):
        _cell(ws, i, 1, r.Category)
        _cell(ws, i, 2, r.Item)
        for j in range(6):
            col = chr(67 + j)
            _cell(ws, i, 3 + j, f"=SUMIFS(SIX_AMT,SIX_CAT,$A{i},SIX_ITEM,$B{i},SIX_MONTH,{col}$4)", "#,##0;;")
        _cell(ws, i, 9, f"=SUM(C{i}:H{i})", NUM, bold=True)
        _cell(ws, i, 10, f"=INDEX($C$4:$H$4,MATCH(MAX(C{i}:H{i}),C{i}:H{i},0))")
    e2 = e1 + len(efs) - 1
    _cell(ws, e2 + 1, 1, "Total", bold=True, fill=FILL_TOTAL)
    _cell(ws, e2 + 1, 2, None, fill=FILL_TOTAL)
    for j in range(7):
        col = chr(67 + j)
        _cell(ws, e2 + 1, 3 + j, f"=SUM({col}{e1}:{col}{e2})", NUM, bold=True, fill=FILL_TOTAL)
    ws.conditional_formatting.add(f"C{e1}:H{e2}", FormulaRule(formula=[f"AND(C{e1}>0,C{e1}=MAX($C{e1}:$H{e1}))"], fill=PatternFill("solid", bgColor="FFEB9C", fgColor="FFEB9C")))
    _widths(ws, {"A": 16, "B": 24, **{chr(67 + j): 11 for j in range(6)}, "I": 11, "J": 13})

    # ---- Essential_Split ---------------------------------------------------
    ws = wb.create_sheet("Essential_Split")
    _title(ws, "Essential vs less essential spending", "Groups come from the table on the Inputs sheet")
    _header(ws, 4, ["Group", "Sum of Expense (INR)", "% of total"])
    for i, grp in enumerate(["Essential", "Less essential", "Miscellaneous (mixed)"], start=5):
        _cell(ws, i, 1, grp)
        _cell(ws, i, 2, f"=SUMIFS(Category_Totals!$B${g1}:$B${g2},Category_Totals!$D${g1}:$D${g2},$A{i})", NUM)
        _cell(ws, i, 3, f"=B{i}/$B$8", "0.0%")
    _cell(ws, 8, 1, "Grand Total", bold=True, fill=FILL_TOTAL)
    _cell(ws, 8, 2, "=SUM(B5:B7)", NUM, bold=True, fill=FILL_TOTAL)
    _cell(ws, 8, 3, "=SUM(C5:C7)", "0.0%", bold=True, fill=FILL_TOTAL)
    pie2 = PieChart()
    pie2.title = "Essential vs less essential"
    pie2.add_data(Reference(ws, min_col=2, min_row=4, max_row=7), titles_from_data=True)
    pie2.set_categories(Reference(ws, min_col=1, min_row=5, max_row=7))
    pie2.dataLabels = DataLabelList(showPercent=True, showVal=False, showCatName=False, showSerName=False, showLeaderLines=False)
    for i, colr in enumerate(["4472C4", "ED7D31", "A5A5A5"]):
        pt = DataPoint(idx=i)
        pt.graphicalProperties.solidFill = colr
        pie2.series[0].dPt.append(pt)
    pie2.height, pie2.width = 8.5, 13
    ws.add_chart(pie2, "E4")
    _widths(ws, {"A": 26, "B": 22, "C": 12})

    # ---- Savings_Plan --------------------------------------------------------
    ws = wb.create_sheet("Savings_Plan")
    _title(ws, "Recommendations: savings plan", "Change the caps on the Inputs sheet to test other plans")
    _header(ws, 4, ["Month", "Spent", "Saved today", "Entertainment", "Shopping", "Food", "Snacks", "Cut: Entertainment", "Cut: Shopping", "Cut: Food", "Cut: Snacks", "Total cut", "Saved with plan"])
    ws.row_dimensions[4].height = 32
    for i, m in enumerate(config.MONTHS, start=5):
        _cell(ws, i, 1, m)
        _cell(ws, i, 2, f"=Monthly_Trend!B{i}", NUM)
        _cell(ws, i, 3, f"=Inputs!$B$3-B{i}", NUM)
        _cell(ws, i, 4, f'=SUMIFS(SIX_AMT,SIX_CAT,"Entertainment",SIX_MONTH,$A{i})', NUM)
        _cell(ws, i, 5, f'=SUMIFS(SIX_AMT,SIX_CAT,"Shopping",SIX_MONTH,$A{i})', NUM)
        _cell(ws, i, 6, f'=SUMIFS(SIX_AMT,SIX_CAT,"Food",SIX_MONTH,$A{i})', NUM)
        _cell(ws, i, 7, f'=SUMIFS(SIX_AMT,SIX_ITEM,"Snacks",SIX_MONTH,$A{i})', NUM)
        _cell(ws, i, 8, f"=MAX(0,D{i}-Inputs!$B$4)", NUM)
        _cell(ws, i, 9, f"=MAX(0,E{i}-Inputs!$B$5)", NUM)
        _cell(ws, i, 10, f"=F{i}*Inputs!$B$6", NUM)
        _cell(ws, i, 11, f"=G{i}*Inputs!$B$7", NUM)
        _cell(ws, i, 12, f"=SUM(H{i}:K{i})", NUM, bold=True)
        _cell(ws, i, 13, f"=C{i}+L{i}", NUM, bold=True)
    _cell(ws, 11, 1, "Total (6 months)", bold=True, fill=FILL_TOTAL)
    _cell(ws, 12, 1, "Average per month", bold=True, fill=FILL_TOTAL)
    for j in range(2, 14):
        col = chr(64 + j)
        _cell(ws, 11, j, f"=SUM({col}5:{col}10)", NUM, bold=True, fill=FILL_TOTAL)
        _cell(ws, 12, j, f"=AVERAGE({col}5:{col}10)", NUM, bold=True, fill=FILL_TOTAL)
    _header(ws, 15, ["Where to cut", "What to do", "", "", "Reason", "", "", "", "", "", "", "", "Saving (6 months)"])
    recs = [
        ("Entertainment", "Cap at the monthly limit; plan trips from a separate fund", "Most of the spend was one trip and two outings, not daily needs", "=H11"),
        ("Shopping", "Cap at the monthly limit; buy only when needed", "Clothes and shoes. No shopping in Feb and Apr, so it can be skipped", "=I11"),
        ("Food", "Cut dining out, online orders and chips; eat home food", "Eating out while groceries were already bought", "=J11"),
        ("Snacks (in Grocery)", "Cut packaged snacks; prefer fruit and home snacks", "Snacks cost more than the entire Food category", "=K11"),
    ]
    for i, (a, b, c, d) in enumerate(recs, start=16):
        _cell(ws, i, 1, a, bold=True)
        _cell(ws, i, 2, b)
        _cell(ws, i, 5, c)
        _cell(ws, i, 13, d, NUM, bold=True)
    _cell(ws, 20, 1, "Total saving", bold=True, fill=FILL_TOTAL)
    _cell(ws, 20, 13, "=SUM(M16:M19)", NUM, bold=True, fill=FILL_TOTAL)
    cc = BarChart()
    cc.type, cc.grouping, cc.title = "col", "clustered", "Money left each month: today vs with the plan (INR)"
    cc.add_data(Reference(ws, min_col=3, min_row=4, max_row=10), titles_from_data=True)
    cc.add_data(Reference(ws, min_col=13, min_row=4, max_row=10), titles_from_data=True)
    cc.set_categories(Reference(ws, min_col=1, min_row=5, max_row=10))
    cc.series[0].graphicalProperties.solidFill = "4472C4"
    cc.series[1].graphicalProperties.solidFill = "ED7D31"
    cc.dataLabels = _val_labels()
    cc.height, cc.width = 9, 20
    ws.add_chart(cc, "A23")
    _widths(ws, {"A": 24, "B": 12, "C": 13, "D": 14, "E": 12, "F": 10, "G": 10, "H": 12, "I": 12, "J": 10, "K": 10, "L": 11, "M": 16})

    # ---- Summary ------------------------------------------------------------
    ws = summary
    _title(ws, "Nitin's savings analysis: summary", "Helping Nitin save for his scooter  |  income Rs 15,000 a month")
    ws["A4"] = "Key numbers"
    ws["A4"].font = Font(bold=True, size=13, color=DEEP)
    kn = [
        ("June: total spent", f"=June_Category!B{cat_rng[1] + 1}", NUM),
        ("June: highest category", "=June_Findings!B5", None),
        ("Six months: total spent", "=Monthly_Trend!B11", NUM),
        ("Six months: average spend per month", "=Monthly_Trend!B12", NUM),
        ("Highest-spend month", "=Monthly_Trend!B14", None),
        ("Average saving per month today", "=Savings_Plan!C12", NUM),
        ("Average saving per month with the plan", "=Savings_Plan!M12", NUM),
        ("Extra saved over six months with the plan", "=Savings_Plan!L11", NUM),
    ]
    for i, (label, f, fmt) in enumerate(kn, start=5):
        _cell(ws, i, 1, label, bold=True)
        _cell(ws, i, 2, f, fmt, align="right")
    ws["A15"] = "Sheets in this workbook"
    ws["A15"].font = Font(bold=True, size=13, color=DEEP)
    guide = [("Inputs", "Income, caps and category groups (editable)"), ("June_Data", "Part 1 raw data"),
             ("June_Category", "Pivot 1: spend by category + pie chart"), ("June_Items", "Pivot 2/3/4: items, counts, filter, data bars, chart"),
             ("June_Findings", "Answers to the five Part 1 questions"), ("SixMonth_Data", "Part 2 raw data"),
             ("Monthly_Trend", "Month-wise trend + line chart"), ("Category_Totals", "Category-wise expenses with data bars"),
             ("Month_x_Category", "Month by category, top 2 each month"), ("Items_EFS", "Entertainment, Food, Shopping items by month"),
             ("Essential_Split", "Essential vs less essential"), ("Savings_Plan", "Recommendations and impact on savings")]
    for i, (sheet, desc) in enumerate(guide, start=16):
        c = ws.cell(row=i, column=1, value=sheet)
        c.hyperlink = f"#'{sheet}'!A1"
        c.font = Font(color="0563C1", underline="single")
        ws.cell(row=i, column=2, value=desc)
    _widths(ws, {"A": 44, "B": 52})

    order = ["Summary", "Inputs", "June_Data", "June_Category", "June_Items", "June_Findings", "SixMonth_Data",
             "Monthly_Trend", "Category_Totals", "Month_x_Category", "Items_EFS", "Essential_Split", "Savings_Plan"]
    wb._sheets = [wb[n] for n in order]
    for sheet in wb:  # print nicely: landscape, one page wide
        sheet.page_setup.orientation = "landscape"
        sheet.page_setup.fitToWidth, sheet.page_setup.fitToHeight = 1, 0
        sheet.sheet_properties.pageSetUpPr.fitToPage = True
    wb.calculation.fullCalcOnLoad = True
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    return path
