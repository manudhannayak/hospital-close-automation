"""Material variance flags with driver-based commentary.

Every sentence is built from the numbers (volume statistics, hours, unit costs), so the commentary
explains *why*, not just *how much*, and can be traced back to the data.
"""
from __future__ import annotations

import pandas as pd

from .model import PL_LINES, REVENUE_LINES

# Review net revenue (gross less deductions) rather than the two contra lines separately
DETAIL_LINES = [l for l in PL_LINES if l not in {"Gross Patient Revenue", "Deductions from Revenue",
                                                 "Total Operating Expense", "EBIDA", "Operating Income"}]
STAT_LABEL = {"patient_days": "patient days", "ed_visits": "ED visits", "surgical_cases": "surgical cases",
              "procedures": "procedures", "tests": "tests", "doses": "doses"}


def _m(x: float) -> str:
    return f"${abs(x) / 1e6:,.2f}M" if abs(x) >= 1e6 else f"${abs(x) / 1e3:,.0f}K"


def flag_variances(v: pd.DataFrame, stats: pd.DataFrame, bstats: pd.DataFrame, settings: dict,
                   facility_names: dict, gl: pd.DataFrame) -> pd.DataFrame:
    mat = settings["materiality"]
    d = v[v.line.isin(DETAIL_LINES)].copy()
    d = d[(d.variance.abs() >= mat["dollars"]) & (d.variance.abs() >= mat["percent"] / 100 * d.budget.abs())]
    rows = []
    for r in d.itertuples():
        fac, cc = r.cost_center.split("-")
        dept = settings["departments"][cc]
        stat = dept["stat"]
        a_vol = stats[(stats.cost_center == r.cost_center) & (stats.statistic == stat)].value.sum()
        b_vol = bstats[(bstats.cost_center == r.cost_center) & (bstats.statistic == stat)].budget.sum()
        driver = "Volume" if abs(r.volume_variance) > abs(r.rate_variance) else "Rate / efficiency"
        over = "over" if r.variance > 0 else "under"
        good = r.favorable
        pct = "no budget for this line" if pd.isna(r.variance_pct) else f"{r.variance_pct:+.0f}%"
        head = f"{r.line} {_m(r.variance)} {over} budget ({pct}), {'favorable' if good else 'unfavorable'}."
        why = []
        if stat != "none" and b_vol:
            lab = STAT_LABEL[stat]
            chg = 100 * (a_vol / b_vol - 1)
            chg_s = "flat" if abs(chg) < .05 else f"{chg:+.1f}%"
            why.append(f"{lab[0].upper() + lab[1:]} {a_vol:,.0f} vs {b_vol:,.0f} budget ({chg_s}) "
                       f"accounts for {_m(r.volume_variance)} {'of the variance' if driver == 'Volume' else 'only'}.")
        if r.line == "Contract Labor":
            ch = stats[(stats.cost_center == r.cost_center) & (stats.statistic == "contract_hours")].value.sum()
            cb = bstats[(bstats.cost_center == r.cost_center) & (bstats.statistic == "contract_hours")].budget.sum()
            why.append(f"Agency hours {ch:,.0f} vs {cb:,.0f} budgeted: contract staff are covering employed vacancies "
                       f"(check employed Salaries & Wages, likely under budget).")
        elif r.line == "Salaries & Wages" and r.variance < 0:
            why.append("Employed labor below plan; confirm whether vacancies are being backfilled with contract labor.")
        elif r.line == "Supplies" and stat != "none" and a_vol:
            per_a, per_b = r.actual / a_vol, r.budget / b_vol if b_vol else 0
            unit = STAT_LABEL[stat][:-1] if STAT_LABEL[stat].endswith("s") else STAT_LABEL[stat]
            tail = ": price or mix, not volume." if driver != "Volume" else "; mostly volume, unit cost close to plan."
            why.append(f"Cost per {unit} ${per_a:,.0f} vs ${per_b:,.0f} budget ({100 * (per_a / per_b - 1):+.1f}%){tail}")
        elif r.line == "Drugs" and cc == "8000":
            misc = gl[(gl.cost_center == r.cost_center) & (gl.account == "6200") & (gl.debit > 0)]
            why.append(f"Drug expense posted to Administration ({', '.join(misc.je_id)}); no drug budget here. "
                       f"Reclass to Pharmacy proposed in Controls.")
        elif driver == "Rate / efficiency" and r.line == "Net Patient Revenue":
            why.append(f"{_m(r.rate_variance)} from net rate / payer mix rather than volume.")
        rows.append({"facility": fac, "facility_name": facility_names[fac], "cost_center": r.cost_center,
                     "department": dept["name"], "line": r.line, "actual": r.actual, "budget": r.budget,
                     "flexed_budget": r.flexed_budget, "volume_variance": r.volume_variance,
                     "rate_variance": r.rate_variance, "variance": r.variance, "variance_pct": r.variance_pct,
                     "favorable": bool(good), "driver": driver, "commentary": " ".join([head] + why)})
    out = pd.DataFrame(rows)
    if len(out):
        out = out.assign(_abs=out.variance.abs()).sort_values("_abs", ascending=False).drop(columns="_abs")
    return out.reset_index(drop=True)


def exec_narrative(period_label: str, sys_pl: pd.DataFrame, flags: pd.DataFrame, k: pd.DataFrame, ctrl: dict) -> str:
    g = sys_pl.set_index("line")
    oi, ni = g.loc["Operating Income"], g.loc["Net Patient Revenue"]
    margin = 100 * oi.actual / ni.actual
    bmargin = 100 * oi.budget / ni.budget
    lines = [f"{period_label}: operating income {_m(oi.actual)} vs {_m(oi.budget)} budget "
             f"({'+' if oi.variance >= 0 else '-'}{_m(oi.variance)}); operating margin {margin:.1f}% vs {bmargin:.1f}% budget. "
             f"Net patient revenue {_m(ni.actual)} ({ni.variance_pct:+.1f}% vs budget)."]
    unfav = flags[~flags.favorable].head(4)
    if len(unfav):
        lines.append("Largest unfavorable variances:")
        lines += [f"- {r.facility_name} {r.department}: {r.commentary}" for r in unfav.itertuples()]
    fav = flags[flags.favorable].head(2)
    if len(fav):
        lines.append("Largest favorable variances:")
        lines += [f"- {r.facility_name} {r.department}: {r.line} {_m(r.variance)} ({r.variance_pct:+.0f}%), {r.driver.lower()}-driven."
                  for r in fav.itertuples()]
    warn = [c for c in ctrl["results"] if c["status"] != "PASS"]
    lines.append(f"Close controls: {ctrl['counts']['PASS']} passed, {ctrl['counts']['WARN']} need review, "
                 f"{ctrl['counts']['FAIL']} failed." + (" Review items: " + "; ".join(c["control"].replace("_", " ") for c in warn) + "." if warn else ""))
    return "\n".join(lines)
