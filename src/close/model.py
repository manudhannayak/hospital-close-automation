"""P&L model, KPIs and flexible-budget variance analysis."""
from __future__ import annotations

import pandas as pd

from . import EXTRACTS
from .simulate import ACCOUNTS

PL_LINES = ["Gross Patient Revenue", "Deductions from Revenue", "Net Patient Revenue", "Salaries & Wages",
            "Contract Labor", "Benefits", "Supplies", "Drugs", "Purchased Services", "Other Operating Expense",
            "Total Operating Expense", "EBIDA", "Depreciation & Interest", "Operating Income"]
REVENUE_LINES = {"Gross Patient Revenue", "Net Patient Revenue", "EBIDA", "Operating Income"}
# Share of each line that moves with volume (the rest is fixed) - drives the flexible budget
VARIABLE_SHARE = {"Gross Patient Revenue": 1.0, "Deductions from Revenue": 1.0, "Salaries & Wages": .8,
                  "Contract Labor": .8, "Benefits": .8, "Supplies": 1.0, "Drugs": 1.0}


def load_period(period: str):
    gl = pd.read_csv(EXTRACTS / f"gl_{period}.csv", dtype={"account": str})
    st = pd.read_csv(EXTRACTS / f"stats_{period}.csv")
    return gl, st


def actuals(gl: pd.DataFrame) -> pd.DataFrame:
    """Income statement amounts, positive in their natural direction."""
    g = gl[gl.account.map(lambda a: ACCOUNTS[a][2]) == "IS"].copy()
    g["line"] = g.account.map(lambda a: ACCOUNTS[a][1])
    g["amount"] = (g.debit - g.credit).where(g.account.str[0] != "4", g.credit - g.debit)
    g.loc[g.account.isin(["4800", "4900"]), "amount"] = g.debit - g.credit   # deductions: positive = reduction
    return g.groupby(["period", "facility", "cost_center", "line"], as_index=False).amount.sum()


def with_subtotals(df: pd.DataFrame, value_cols: list[str], keys: list[str]) -> pd.DataFrame:
    """Pivot detail lines into the full P&L with subtotals, for any key level."""
    piv = df.pivot_table(index=keys, columns="line", values=value_cols, aggfunc="sum", fill_value=0)
    out = []
    for v in value_cols:
        p = piv[v].copy()
        for c in PL_LINES:
            if c not in p:
                p[c] = 0.0
        p["Net Patient Revenue"] = p["Gross Patient Revenue"] - p["Deductions from Revenue"]
        p["Total Operating Expense"] = p[["Salaries & Wages", "Contract Labor", "Benefits", "Supplies", "Drugs",
                                          "Purchased Services", "Other Operating Expense"]].sum(axis=1)
        p["EBIDA"] = p["Net Patient Revenue"] - p["Total Operating Expense"]
        p["Operating Income"] = p["EBIDA"] - p["Depreciation & Interest"]
        out.append(p[PL_LINES].stack().rename(v))
    return pd.concat(out, axis=1).reset_index().rename(columns={"level_%d" % len(keys): "line"})


def budget_for(period: str) -> tuple[pd.DataFrame, pd.DataFrame]:
    year = period[:4]
    b = pd.read_csv(EXTRACTS / f"budget_FY{year}.csv", dtype={"account": str})
    bs = pd.read_csv(EXTRACTS / f"budget_stats_FY{year}.csv")
    b = b[b.period == period].copy()
    b["line"] = b.account.map(lambda a: ACCOUNTS[a][1])
    b["facility"] = b.cost_center.str[0]
    return b.groupby(["period", "facility", "cost_center", "line"], as_index=False).budget.sum(), bs[bs.period == period]


def flex_factors(stats: pd.DataFrame, bstats: pd.DataFrame, settings: dict) -> pd.Series:
    """Actual / budget volume per department (1.0 for overhead)."""
    vol = {c: d["stat"] for c, d in settings["departments"].items()}
    a = stats[stats.apply(lambda r: r.statistic == vol[r.cost_center[2:]], axis=1)].set_index("cost_center").value
    b = bstats[bstats.apply(lambda r: r.statistic == vol[r.cost_center[2:]], axis=1)].set_index("cost_center").budget
    return (a / b).rename("flex")


def variance_table(period: str, act: pd.DataFrame, stats: pd.DataFrame, prior_year: pd.DataFrame | None,
                   settings: dict) -> pd.DataFrame:
    """Cost center x P&L line: actual, budget, flexed budget, volume and rate/efficiency variances, prior year."""
    bud, bstats = budget_for(period)
    f = flex_factors(stats, bstats, settings)
    d = act.merge(bud, on=["period", "facility", "cost_center", "line"], how="outer").fillna({"amount": 0, "budget": 0})
    d["flex"] = d.cost_center.map(f).fillna(1.0)
    d["var_share"] = d.line.map(VARIABLE_SHARE).fillna(0.0)
    d["flexed_budget"] = d.budget * (1 + d.var_share * (d.flex - 1))
    if prior_year is not None:
        d = d.merge(prior_year.rename(columns={"amount": "prior_year"}).drop(columns="period"),
                    on=["facility", "cost_center", "line"], how="left").fillna({"prior_year": 0})
    else:
        d["prior_year"] = 0.0
    d = d.rename(columns={"amount": "actual"})
    full = with_subtotals(d, ["actual", "budget", "flexed_budget", "prior_year"], ["facility", "cost_center"])
    full["variance"] = full.actual - full.budget
    full["volume_variance"] = full.flexed_budget - full.budget
    full["rate_variance"] = full.actual - full.flexed_budget
    # favorable = higher revenue/income, lower deductions/expense
    sign = full.line.map(lambda l: 1 if l in REVENUE_LINES else -1)
    full["favorable"] = (full.variance * sign) >= 0
    full["variance_pct"] = (100 * full.variance / full.budget.abs().where(full.budget != 0)).round(1)
    return full


def roll_up(v: pd.DataFrame, by: list[str]) -> pd.DataFrame:
    cols = ["actual", "budget", "flexed_budget", "prior_year", "variance", "volume_variance", "rate_variance"]
    r = v.groupby(by + ["line"], as_index=False)[cols].sum()
    sign = r.line.map(lambda l: 1 if l in REVENUE_LINES else -1)
    r["favorable"] = (r.variance * sign) >= 0
    r["variance_pct"] = (100 * r.variance / r.budget.abs().where(r.budget != 0)).round(1)
    r["line"] = pd.Categorical(r.line, PL_LINES, ordered=True)
    return r.sort_values(by + ["line"]).reset_index(drop=True)


def kpis(act: pd.DataFrame, stats: pd.DataFrame, gl: pd.DataFrame) -> pd.DataFrame:
    """Facility-month KPIs used in the exec summary and trend tab."""
    rows = []
    for (period, fac), a in act.groupby(["period", "facility"]):
        s = stats[(stats.period == period) & stats.cost_center.str.startswith(fac)]
        g = gl[(gl.period == period) & (gl.facility == fac)]
        line = a.groupby("line").amount.sum()
        net = line.get("Gross Patient Revenue", 0) - line.get("Deductions from Revenue", 0)
        opex = sum(line.get(x, 0) for x in ["Salaries & Wages", "Contract Labor", "Benefits", "Supplies", "Drugs",
                                            "Purchased Services", "Other Operating Expense", "Depreciation & Interest"])
        ip = (g[g.account == "4000"].credit.sum())
        gross = g[g.account.isin(["4000", "4100"])].credit.sum()
        pdays = s[s.statistic == "patient_days"].value.sum()
        apd = pdays * gross / ip if ip else 0
        labor = line.get("Salaries & Wages", 0) + line.get("Contract Labor", 0) + line.get("Benefits", 0)
        ot = g[g.account == "5100"].debit.sum()
        sal = g[g.account.isin(["5000", "5010", "5020", "5100"])].debit.sum()
        cases = s[s.statistic == "surgical_cases"].value.sum()
        surg_sup = g[(g.cost_center == f"{fac}-6200") & g.account.isin(["6000", "6100"])].debit.sum()
        hours = s[s.statistic == "worked_hours"].value.sum()
        rows.append({"period": period, "facility": fac, "net_revenue": net, "operating_income": net - opex,
                     "operating_margin_pct": 100 * (net - opex) / net if net else None,
                     "adjusted_patient_days": apd, "net_revenue_per_apd": net / apd if apd else None,
                     "expense_per_apd": opex / apd if apd else None,
                     "labor_pct_net_revenue": 100 * labor / net if net else None,
                     "contract_labor_pct_labor": 100 * line.get("Contract Labor", 0) / labor if labor else None,
                     "overtime_pct_salaries": 100 * ot / sal if sal else None,
                     "supply_cost_per_case": surg_sup / cases if cases else None,
                     "worked_hours_per_apd": hours / apd if apd else None})
    return pd.DataFrame(rows)
