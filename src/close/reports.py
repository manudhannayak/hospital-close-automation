"""Generate the close package: system Excel pack, 32 department packs and a PDF executive summary.

Excel conventions (finance standard): Arial; blue = values loaded from the ledger/budget,
black = formulas; variances shown Fav/(Unfav) so a negative number is always bad news;
subtotals, variances and system totals are live formulas, so the pack recalculates if edited.
"""
from __future__ import annotations

from pathlib import Path

import matplotlib
import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as col

from .model import PL_LINES, REVENUE_LINES

matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402

NAVY = "1F3A5F"
F = "Arial"
INPUT = Font(name=F, size=10, color="0000FF")
CALC = Font(name=F, size=10, color="000000")
BOLD = Font(name=F, size=10, bold=True)
HEAD = Font(name=F, size=10, bold=True, color="FFFFFF")
TITLE = Font(name=F, size=14, bold=True, color=NAVY)
SUB = Font(name=F, size=10, italic=True, color="555555")
FILL_HEAD = PatternFill("solid", fgColor=NAVY)
FILL_SUB = PatternFill("solid", fgColor="E8EEF5")
THIN = Border(bottom=Side(style="thin", color="BFC7D1"))
TOP = Border(top=Side(style="thin", color="1F3A5F"))
USD = '$#,##0;($#,##0);"-"'
PCT = '0.0%;(0.0%);"-"'
SUBTOTALS = {"Net Patient Revenue", "Total Operating Expense", "EBIDA", "Operating Income"}
RED = Font(name=F, size=10, color="C00000")


def _print_setup(ws):
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_options.horizontalCentered = True
    ws.page_margins.left = ws.page_margins.right = 0.4


def _hdr(ws, row: int, labels: list[str], start: int = 1):
    for i, l in enumerate(labels):
        c = ws.cell(row, start + i, l)
        c.font, c.fill = HEAD, FILL_HEAD
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def _title(ws, title: str, subtitle: str):
    ws["A1"], ws["A2"] = title, subtitle
    ws["A1"].font, ws["A2"].font = TITLE, SUB
    ws.sheet_view.showGridLines = False


def _neg_red(ws, rng: str):
    ws.conditional_formatting.add(rng, CellIsRule(operator="lessThan", formula=["0"], font=RED))


def _pl_block(ws, top: int, left: int, data: dict[str, dict[str, float]], cols: list[str], label: str):
    """Write one P&L block. data[line][col] for detail lines; subtotals become formulas.
    cols = value columns (e.g. Actual, Budget); adds Var Fav/(Unfav) and Var % after Actual/Budget."""
    ws.cell(top, left, label).font = BOLD
    heads = cols + ["Var Fav/(Unfav)", "Var %"]
    _hdr(ws, top + 1, heads, left)
    r0 = top + 2
    rowof = {l: r0 + i for i, l in enumerate(PL_LINES)}
    for line, r in rowof.items():
        for j, c in enumerate(cols):
            cell = ws.cell(r, left + j)
            L = col(left + j)
            if line == "Net Patient Revenue":
                cell.value = f"={L}{rowof['Gross Patient Revenue']}-{L}{rowof['Deductions from Revenue']}"
            elif line == "Total Operating Expense":
                cell.value = f"=SUM({L}{rowof['Salaries & Wages']}:{L}{rowof['Other Operating Expense']})"
            elif line == "EBIDA":
                cell.value = f"={L}{rowof['Net Patient Revenue']}-{L}{rowof['Total Operating Expense']}"
            elif line == "Operating Income":
                cell.value = f"={L}{rowof['EBIDA']}-{L}{rowof['Depreciation & Interest']}"
            else:
                cell.value = round(float(data.get(line, {}).get(c, 0.0)), 2)
            cell.font = CALC if line in SUBTOTALS else INPUT
            cell.number_format = USD
        A, B = col(left), col(left + 1)
        v = ws.cell(r, left + len(cols))
        v.value = f"={A}{r}-{B}{r}" if line in REVENUE_LINES else f"={B}{r}-{A}{r}"
        v.number_format, v.font = USD, CALC
        p = ws.cell(r, left + len(cols) + 1)
        p.value = f"=IFERROR({col(left + len(cols))}{r}/ABS({B}{r}),0)"
        p.number_format, p.font = PCT, CALC
        if line in SUBTOTALS:
            for j in range(len(heads)):
                ws.cell(r, left + j).fill = FILL_SUB
                ws.cell(r, left + j).font = Font(name=F, size=10, bold=True,
                                                  color="0000FF" if False else "000000")
                ws.cell(r, left + j).border = TOP
    _neg_red(ws, f"{col(left + len(cols))}{r0}:{col(left + len(cols) + 1)}{r0 + len(PL_LINES) - 1}")
    return rowof


def system_pack(path: Path, period_label: str, org: str, fac_names: dict, fac_pl: pd.DataFrame, flags: pd.DataFrame,
                v: pd.DataFrame, kpi: pd.DataFrame, controls: dict, reclass: pd.DataFrame, narrative: str,
                settings: dict, run_info: dict) -> None:
    wb = Workbook()

    # --- Cover ---
    ws = wb.active
    ws.title = "Cover"
    _title(ws, f"{org} - Monthly Close Package", f"{period_label} · generated automatically {run_info['generated_at']} · synthetic data")
    rows = [("Contents", ""), ("Exec Summary", "Narrative, KPIs by facility"), ("P&L by Facility", "Actual vs budget, Fav/(Unfav)"),
            ("Variance Analysis", "Material variances: volume vs rate/efficiency, with commentary"),
            ("Department Detail", "All 32 departments x P&L lines (filterable)"), ("KPI Trend", "13-month KPI trend + chart"),
            ("Controls", "Close controls and proposed correcting entries"), ("Notes", "Definitions and assumptions"),
            ("", ""), ("Close controls", f"{controls['counts']['PASS']} passed · {controls['counts']['WARN']} review · {controls['counts']['FAIL']} failed"),
            ("Source extract rows", f"{run_info['gl_rows']:,} GL lines · {run_info['stat_rows']:,} statistics"),
            ("Input fingerprint", run_info["input_sha256"][:16])]
    for i, (a, b) in enumerate(rows, start=4):
        ws.cell(i, 1, a).font = BOLD if i in (4, 13, 14, 15) else CALC
        ws.cell(i, 2, b).font = CALC
    ws.column_dimensions["A"].width, ws.column_dimensions["B"].width = 26, 70

    # --- Exec summary ---
    ws = wb.create_sheet("Exec Summary")
    _title(ws, "Executive Summary", period_label)
    ws["A4"] = narrative
    ws["A4"].alignment = Alignment(wrap_text=True, vertical="top")
    ws["A4"].font = CALC
    ws.merge_cells("A4:H18")
    ws.row_dimensions[4].height = 15
    k = kpi[kpi.period == run_info["period"]].set_index("facility")
    labels = [("Net patient revenue", "net_revenue", USD), ("Operating income", "operating_income", USD),
              ("Operating margin", "operating_margin_pct", PCT), ("Net revenue / adj. patient day", "net_revenue_per_apd", USD),
              ("Expense / adj. patient day", "expense_per_apd", USD), ("Labor % of net revenue", "labor_pct_net_revenue", PCT),
              ("Contract labor % of labor", "contract_labor_pct_labor", PCT), ("Overtime % of salaries", "overtime_pct_salaries", PCT),
              ("Supply cost / surgical case", "supply_cost_per_case", USD), ("Worked hours / adj. patient day", "worked_hours_per_apd", "0.0")]
    _hdr(ws, 20, ["KPI"] + [fac_names[f] for f in fac_names])
    for i, (lab, key, fmt) in enumerate(labels, start=21):
        ws.cell(i, 1, lab).font = CALC
        for j, f in enumerate(fac_names, start=2):
            val = k.loc[f, key] if f in k.index else None
            if val is not None and fmt == PCT:
                val = val / 100
            c = ws.cell(i, j, None if pd.isna(val) else round(float(val), 4))
            c.number_format, c.font = fmt, INPUT
        ws.cell(i, 1).border = THIN
    ws.column_dimensions["A"].width = 34
    for j in range(2, 9):
        ws.column_dimensions[col(j)].width = 16

    # --- P&L by facility ---
    ws = wb.create_sheet("P&L by Facility")
    _title(ws, "Income Statement by Facility", f"{period_label} · Actual vs Budget · variances Fav/(Unfav)")
    left = 2
    blocks = {}
    for f, name in fac_names.items():
        d = fac_pl[fac_pl.facility == f].set_index("line")
        data = {l: {"Actual": d.loc[l, "actual"], "Budget": d.loc[l, "budget"]} for l in d.index}
        blocks[f] = (left, _pl_block(ws, 4, left, data, ["Actual", "Budget"], name))
        left += 5
    # System = formulas summing facility blocks
    ws.cell(4, left, "System Total").font = BOLD
    _hdr(ws, 5, ["Actual", "Budget", "Var Fav/(Unfav)", "Var %"], left)
    rowof = next(iter(blocks.values()))[1]
    for line, r in rowof.items():
        for j in range(2):
            refs = "+".join(f"{col(bl + j)}{r}" for bl, _ in blocks.values())
            c = ws.cell(r, left + j, f"={refs}")
            c.number_format, c.font = USD, BOLD if line in SUBTOTALS else CALC
        A, B = col(left), col(left + 1)
        ws.cell(r, left + 2, f"={A}{r}-{B}{r}" if line in REVENUE_LINES else f"={B}{r}-{A}{r}").number_format = USD
        ws.cell(r, left + 3, f"=IFERROR({col(left + 2)}{r}/ABS({B}{r}),0)").number_format = PCT
        for j in range(4):
            if line in SUBTOTALS:
                ws.cell(r, left + j).fill, ws.cell(r, left + j).border = FILL_SUB, TOP
    _neg_red(ws, f"{col(left + 2)}6:{col(left + 3)}{5 + len(PL_LINES)}")
    for i, line in enumerate(PL_LINES, start=6):
        ws.cell(i, 1, line).font = BOLD if line in SUBTOTALS else CALC
    ws.column_dimensions["A"].width = 26
    for j in range(2, left + 4):
        ws.column_dimensions[col(j)].width = 13
    ws.freeze_panes = "B6"

    # --- Variance analysis ---
    ws = wb.create_sheet("Variance Analysis")
    _title(ws, "Material Variances: Volume vs Rate/Efficiency",
           f"Flagged when |variance| >= ${settings['materiality']['dollars']:,} and >= {settings['materiality']['percent']}% of budget. "
           "Flexed budget = budget adjusted to actual volume.")
    heads = ["Facility", "Department", "P&L line", "Actual", "Budget", "Flexed budget", "Volume variance",
             "Rate / efficiency variance", "Total variance", "Fav / Unfav", "Main driver", "Commentary"]
    _hdr(ws, 4, heads)
    for i, r in enumerate(flags.itertuples(), start=5):
        vals = [r.facility_name, r.department, r.line, r.actual, r.budget, r.flexed_budget]
        for j, x in enumerate(vals, start=1):
            c = ws.cell(i, j, round(x, 2) if isinstance(x, float) else x)
            c.font = INPUT if j >= 4 else CALC
            if j >= 4:
                c.number_format = USD
        ws.cell(i, 7, f"=F{i}-E{i}").number_format = USD
        ws.cell(i, 8, f"=D{i}-F{i}").number_format = USD
        ws.cell(i, 9, f"=D{i}-E{i}").number_format = USD
        ws.cell(i, 10, "Favorable" if r.favorable else "Unfavorable").font = Font(name=F, size=10, color="1A7F4B" if r.favorable else "C00000")
        ws.cell(i, 11, r.driver).font = CALC
        c = ws.cell(i, 12, r.commentary)
        c.alignment, c.font = Alignment(wrap_text=True, vertical="top"), CALC
        ws.row_dimensions[i].height = 48
        for j in range(1, 12):
            ws.cell(i, j).alignment = Alignment(vertical="top", wrap_text=j <= 3)
    for j, w in enumerate([14, 18, 20, 13, 13, 13, 13, 14, 13, 12, 15, 90], start=1):
        ws.column_dimensions[col(j)].width = w
    ws.freeze_panes = "D5"

    # --- Department detail ---
    ws = wb.create_sheet("Department Detail")
    _title(ws, "Department Detail", "Every department x P&L line. Use the filters.")
    heads = ["Facility", "Cost center", "Department", "P&L line", "Actual", "Budget", "Flexed budget", "Prior year",
             "Var vs budget Fav/(Unfav)", "Var vs flexed Fav/(Unfav)"]
    _hdr(ws, 4, heads)
    det = v[v.line.isin(PL_LINES)].copy()
    det["order"] = det.line.map({l: i for i, l in enumerate(PL_LINES)})
    det = det.sort_values(["cost_center", "order"])
    for i, r in enumerate(det.itertuples(), start=5):
        f_, cc = r.cost_center.split("-")
        vals = [fac_names[f_], r.cost_center, settings["departments"][cc]["name"], r.line,
                r.actual, r.budget, r.flexed_budget, r.prior_year]
        for j, x in enumerate(vals, start=1):
            c = ws.cell(i, j, round(float(x), 2) if j >= 5 else x)
            c.font = INPUT if j >= 5 else CALC
            if j >= 5:
                c.number_format = USD
        rev = r.line in REVENUE_LINES
        ws.cell(i, 9, f"=E{i}-F{i}" if rev else f"=F{i}-E{i}").number_format = USD
        ws.cell(i, 10, f"=E{i}-G{i}" if rev else f"=G{i}-E{i}").number_format = USD
    last = 4 + len(det)
    ws.auto_filter.ref = f"A4:J{last}"
    _neg_red(ws, f"I5:J{last}")
    for j, w in enumerate([14, 11, 18, 24, 13, 13, 13, 13, 15, 15], start=1):
        ws.column_dimensions[col(j)].width = w
    ws.freeze_panes = "E5"

    # --- KPI trend + chart ---
    ws = wb.create_sheet("KPI Trend")
    _title(ws, "13-Month KPI Trend", "Operating margin, contract labor share and supply cost per surgical case by facility")
    months = sorted(kpi.period.unique())[-13:]
    blocks_k = [("Operating margin", "operating_margin_pct", PCT, True), ("Contract labor % of labor", "contract_labor_pct_labor", PCT, True),
                ("Supply cost / surgical case", "supply_cost_per_case", USD, False)]
    top = 4
    chart_ref = None
    for title, key, fmt, pct in blocks_k:
        ws.cell(top, 1, title).font = BOLD
        _hdr(ws, top + 1, ["Month"] + [fac_names[f] for f in fac_names])
        for i, mth in enumerate(months, start=top + 2):
            ws.cell(i, 1, mth).font = CALC
            for j, f in enumerate(fac_names, start=2):
                val = kpi[(kpi.period == mth) & (kpi.facility == f)][key]
                x = float(val.iloc[0]) if len(val) and pd.notna(val.iloc[0]) else None
                c = ws.cell(i, j, (x / 100 if pct else x) if x is not None else None)
                c.number_format, c.font = fmt, INPUT
        if chart_ref is None:
            chart_ref = (top + 1, top + 1 + len(months))
        top += len(months) + 4
    ch = LineChart()
    ch.title, ch.y_axis.title, ch.height, ch.width = "Operating margin by facility", "Margin", 8, 18
    ch.y_axis.number_format = "0%"
    ch.add_data(Reference(ws, min_col=2, max_col=1 + len(fac_names), min_row=chart_ref[0], max_row=chart_ref[1]), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=1, min_row=chart_ref[0] + 1, max_row=chart_ref[1]))
    ws.add_chart(ch, "G4")
    ws.column_dimensions["A"].width = 12
    for j in range(2, 6):
        ws.column_dimensions[col(j)].width = 15

    # --- Controls ---
    ws = wb.create_sheet("Controls")
    _title(ws, "Close Controls", "FAIL blocks distribution; REVIEW items must be signed off by the controller")
    _hdr(ws, 4, ["Control", "Status", "Detail", "Dollars"])
    for i, c in enumerate(controls["results"], start=5):
        ws.cell(i, 1, c["control"].replace("_", " ").title()).font = CALC
        s = ws.cell(i, 2, {"PASS": "Pass", "WARN": "Review", "FAIL": "FAIL"}[c["status"]])
        s.font = Font(name=F, size=10, bold=True, color={"PASS": "1A7F4B", "WARN": "B26B00", "FAIL": "C00000"}[c["status"]])
        d = ws.cell(i, 3, c["detail"])
        d.alignment, d.font = Alignment(wrap_text=True, vertical="top"), CALC
        ws.cell(i, 4, round(c["dollars"], 2) or None).number_format = USD
    r = 6 + len(controls["results"])
    ws.cell(r, 1, "Proposed correcting entries").font = BOLD
    _hdr(ws, r + 1, ["Account", "Cost center", "Debit", "Credit", "Memo"])
    for i, e in enumerate(reclass.itertuples(), start=r + 2):
        for j, x in enumerate([e.account, e.cost_center, e.debit or None, e.credit or None, e.memo], start=1):
            c = ws.cell(i, j, x)
            c.font = INPUT
            if j in (3, 4):
                c.number_format = USD
    if len(reclass):
        tr = r + 2 + len(reclass)
        ws.cell(tr, 2, "Total").font = BOLD
        ws.cell(tr, 3, f"=SUM(C{r + 2}:C{tr - 1})").number_format = USD
        ws.cell(tr, 4, f"=SUM(D{r + 2}:D{tr - 1})").number_format = USD
        ws.cell(tr, 5, f'=IF(ROUND(C{tr}-D{tr},2)=0,"Balanced","OUT OF BALANCE")').font = BOLD
    for j, w in enumerate([30, 12, 80, 14, 50], start=1):
        ws.column_dimensions[col(j)].width = w

    # --- Notes ---
    ws = wb.create_sheet("Notes")
    _title(ws, "Definitions & Assumptions", "Document every number that isn't straight from the ledger")
    notes = [
        ("Source", "Simulated ERP general-ledger and statistics extracts for a fictional health system (Desert Valley Health). Synthetic data."),
        ("Color code", "Blue = value loaded from ledger, budget or statistics. Black = formula."),
        ("Fav/(Unfav)", "Revenue and income: actual minus budget. Deductions and expenses: budget minus actual. Negative = unfavorable."),
        ("Flexed budget", "Budget adjusted to actual volume: budget x (1 + variable share x (actual volume / budget volume - 1))."),
        ("Variable share", "Revenue, deductions, supplies, drugs 100%; salaries, contract labor, benefits 80%; all other lines fixed."),
        ("Volume variance", "Flexed budget - budget: the part of the variance explained by more or fewer patients."),
        ("Rate / efficiency variance", "Actual - flexed budget: price, mix, staffing efficiency or spending decisions."),
        ("Materiality", f"|variance| >= ${settings['materiality']['dollars']:,} and >= {settings['materiality']['percent']}% of budget."),
        ("Adjusted patient days", "Inpatient patient days x (total gross revenue / inpatient gross revenue): converts outpatient work into inpatient-day equivalents."),
        ("EBIDA", "Earnings before interest, depreciation and amortization."),
    ]
    for i, (a, b) in enumerate(notes, start=4):
        ws.cell(i, 1, a).font = BOLD
        c = ws.cell(i, 2, b)
        c.font, c.alignment = CALC, Alignment(wrap_text=True, vertical="top")
    ws.column_dimensions["A"].width, ws.column_dimensions["B"].width = 26, 110

    for s in wb.worksheets:
        s.sheet_properties.tabColor = NAVY
        _print_setup(s)
    wb.calculation.fullCalcOnLoad = True
    wb.save(path)


def department_pack(path: Path, period_label: str, org: str, cc: str, dept: str, fac_name: str, owner: str,
                    v: pd.DataFrame, stats: pd.DataFrame, bstats: pd.DataFrame, flags: pd.DataFrame) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Summary"
    _title(ws, f"{dept} - {fac_name} ({cc})", f"{org} · {period_label} · prepared for {owner}")
    heads = ["P&L line", "Actual", "Budget", "Flexed budget", "Prior year", "Var vs budget Fav/(Unfav)", "Var vs flexed Fav/(Unfav)"]
    _hdr(ws, 4, heads)
    d = v[v.cost_center == cc].set_index("line")
    for i, line in enumerate(PL_LINES, start=5):
        ws.cell(i, 1, line).font = BOLD if line in SUBTOTALS else CALC
        for j, key in enumerate(["actual", "budget", "flexed_budget", "prior_year"], start=2):
            c = ws.cell(i, j, round(float(d.loc[line, key]), 2) if line in d.index else 0)
            c.number_format, c.font = USD, INPUT
        rev = line in REVENUE_LINES
        ws.cell(i, 6, f"=B{i}-C{i}" if rev else f"=C{i}-B{i}").number_format = USD
        ws.cell(i, 7, f"=B{i}-D{i}" if rev else f"=D{i}-B{i}").number_format = USD
        if line in SUBTOTALS:
            for j in range(1, 8):
                ws.cell(i, j).fill, ws.cell(i, j).border = FILL_SUB, TOP
    _neg_red(ws, f"F5:G{4 + len(PL_LINES)}")
    r = 6 + len(PL_LINES)
    ws.cell(r, 1, "Volume & labor").font = BOLD
    _hdr(ws, r + 1, ["Statistic", "Actual", "Budget", "Variance", "Variance %"])
    s_a = stats[stats.cost_center == cc].set_index("statistic").value
    s_b = bstats[bstats.cost_center == cc].set_index("statistic").budget
    for i, k in enumerate(s_a.index, start=r + 2):
        ws.cell(i, 1, k.replace("_", " ").capitalize()).font = CALC
        ws.cell(i, 2, float(s_a[k])).font = INPUT
        ws.cell(i, 3, float(s_b.get(k, 0))).font = INPUT
        ws.cell(i, 4, f"=B{i}-C{i}")
        ws.cell(i, 5, f"=IFERROR(D{i}/C{i},0)").number_format = PCT
        for j in (2, 3, 4):
            ws.cell(i, j).number_format = "#,##0"
    r = r + 4 + len(s_a)
    ws.cell(r, 1, "What needs your attention").font = BOLD
    mine = flags[flags.cost_center == cc]
    if len(mine):
        for i, f in enumerate(mine.itertuples(), start=r + 1):
            c = ws.cell(i, 1, f"{f.line}: {f.commentary}")
            c.alignment, c.font = Alignment(wrap_text=True, vertical="top"), CALC
            ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=7)
            ws.row_dimensions[i].height = 45
    else:
        ws.cell(r + 1, 1, "No material variances this month.").font = CALC
    ws.column_dimensions["A"].width = 28
    for j in range(2, 8):
        ws.column_dimensions[col(j)].width = 15
    ws.freeze_panes = "B5"
    _print_setup(ws)
    wb.calculation.fullCalcOnLoad = True
    wb.save(path)


def pdf_summary(path: Path, period_label: str, org: str, narrative: str, sys_pl: pd.DataFrame, flags: pd.DataFrame,
                kpi: pd.DataFrame, fac_names: dict, controls: dict, chart_dir: Path) -> None:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import letter
    from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
    from reportlab.lib.units import inch
    from reportlab.platypus import Image, KeepTogether, PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

    # Operating margin trend chart
    fig, ax = plt.subplots(figsize=(7.2, 2.6), dpi=150)
    months = sorted(kpi.period.unique())[-13:]
    palette = ["#1F3A5F", "#2F7DBF", "#C0661B", "#2A8A7A"]
    for (f, name), c in zip(fac_names.items(), palette):
        k = kpi[(kpi.facility == f) & kpi.period.isin(months)].sort_values("period")
        ax.plot(k.period, k.operating_margin_pct, marker="o", ms=3, lw=1.8, color=c, label=name)
    ax.set_ylabel("Operating margin %", fontsize=8)
    ax.tick_params(labelsize=7)
    ax.set_xticks(months[::2])
    ax.grid(axis="y", color="#eee")
    for s in ("top", "right"):
        ax.spines[s].set_visible(False)
    ax.legend(fontsize=7, ncol=4, frameon=False, loc="upper center", bbox_to_anchor=(.5, 1.18))
    fig.tight_layout()
    img = chart_dir / "margin_trend.png"
    fig.savefig(img)
    plt.close(fig)

    ss = getSampleStyleSheet()
    body = ParagraphStyle("b", parent=ss["Normal"], fontName="Helvetica", fontSize=9, leading=12.5)
    h1 = ParagraphStyle("h1", parent=ss["Title"], fontName="Helvetica-Bold", fontSize=16, textColor=colors.HexColor("#1F3A5F"), alignment=0)
    h2 = ParagraphStyle("h2", parent=ss["Heading2"], fontName="Helvetica-Bold", fontSize=11, textColor=colors.HexColor("#1F3A5F"), spaceBefore=10)
    small = ParagraphStyle("s", parent=body, fontSize=7.5, textColor=colors.grey)
    doc = SimpleDocTemplate(str(path), pagesize=letter, leftMargin=.6 * inch, rightMargin=.6 * inch, topMargin=.55 * inch, bottomMargin=.55 * inch)
    els = [Paragraph(f"{org}: Monthly Close Executive Summary", h1), Paragraph(f"{period_label} · auto-generated · synthetic data", small), Spacer(1, 8)]
    for para in narrative.split("\n"):
        els.append(Paragraph(para.replace("&", "&amp;"), body))
    els.append(Paragraph("System income statement", h2))
    g = sys_pl.set_index("line")
    data = [["", "Actual", "Budget", "Fav/(Unfav)", "%"]]
    for l in PL_LINES:
        r = g.loc[l]
        fav = r.actual - r.budget if l in REVENUE_LINES else r.budget - r.actual
        data.append([l, f"${r.actual:,.0f}", f"${r.budget:,.0f}", f"(${-fav:,.0f})" if fav < 0 else f"${fav:,.0f}",
                     f"{100 * fav / abs(r.budget):.1f}%" if r.budget else "-"])
    t = Table(data, colWidths=[2.3 * inch, 1.3 * inch, 1.3 * inch, 1.2 * inch, .7 * inch])
    sty = [("FONT", (0, 0), (-1, -1), "Helvetica", 8.5), ("FONT", (0, 0), (-1, 0), "Helvetica-Bold", 8.5),
           ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F3A5F")), ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
           ("ALIGN", (1, 0), (-1, -1), "RIGHT"), ("LINEBELOW", (0, 1), (-1, -1), .25, colors.HexColor("#DDE3EA"))]
    for i, l in enumerate(PL_LINES, start=1):
        if l in SUBTOTALS:
            sty += [("FONT", (0, i), (-1, i), "Helvetica-Bold", 8.5), ("BACKGROUND", (0, i), (-1, i), colors.HexColor("#E8EEF5"))]
        if data[i][3].startswith("("):
            sty.append(("TEXTCOLOR", (3, i), (4, i), colors.HexColor("#C00000")))
    t.setStyle(TableStyle(sty))
    els += [t, KeepTogether([Paragraph("Operating margin trend", h2), Image(str(img), width=7.2 * inch, height=2.6 * inch)]),
            Paragraph("Material variances (top 10)", h2)]
    vt = [["Facility / dept", "Line", "Variance", "Driver", "Commentary"]]
    for r in flags.head(10).itertuples():
        vt.append([Paragraph(f"{r.facility_name}<br/>{r.department}", small), Paragraph(r.line, small),
                   f"{'' if r.favorable else '('}${abs(r.variance):,.0f}{'' if r.favorable else ')'}",
                   Paragraph(r.driver, small), Paragraph(r.commentary.replace("&", "&amp;"), small)])
    t2 = Table(vt, colWidths=[1.15 * inch, .95 * inch, .8 * inch, .8 * inch, 3.6 * inch], repeatRows=1)
    t2.setStyle(TableStyle([("FONT", (0, 0), (-1, 0), "Helvetica-Bold", 8), ("FONT", (0, 1), (-1, -1), "Helvetica", 8),
                            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F3A5F")), ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                            ("VALIGN", (0, 0), (-1, -1), "TOP"), ("LINEBELOW", (0, 1), (-1, -1), .25, colors.HexColor("#DDE3EA"))]))
    els += [t2, Paragraph("Close controls", h2)]
    ct = [["Control", "Status", "Detail"]] + [[c["control"].replace("_", " "), c["status"], Paragraph(c["detail"].replace("&", "&amp;"), small)]
                                               for c in controls["results"]]
    t3 = Table(ct, colWidths=[1.8 * inch, .7 * inch, 4.8 * inch])
    t3.setStyle(TableStyle([("FONT", (0, 0), (-1, 0), "Helvetica-Bold", 8), ("FONT", (0, 1), (-1, -1), "Helvetica", 8),
                            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F3A5F")), ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                            ("VALIGN", (0, 0), (-1, -1), "TOP"), ("LINEBELOW", (0, 1), (-1, -1), .25, colors.HexColor("#DDE3EA"))]))
    els.append(t3)
    doc.build(els)
