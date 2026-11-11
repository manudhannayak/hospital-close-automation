"""Simulated ERP general-ledger extracts for a fictional 4-hospital system ("Desert Valley Health").

Real hospital ledgers are private, so this module produces what an ERP (Workday/Oracle/Lawson) GL
extract looks like: balanced journal entries by facility, cost center and natural account, a
monthly statistics file (patient days, ED visits, surgical cases, worked hours...) and an FY2026
budget. It is deterministic per month.

Injected so the close automation has real work to do in Sep-Oct 2026:
  * East Campus RN shortage: contract labor replaces employed nursing hours (rate-driven variance)
  * North Campus implant price increase in Surgical Services (rate variance, volume normal)
  * West Campus ED volume drop after a competitor opens (volume variance)
  * A $186,400 drug invoice mis-posted to West Administration instead of Pharmacy (control finding)
  * One duplicated journal line in the October extract (control finding)
"""
from __future__ import annotations

import calendar

import numpy as np
import pandas as pd

from . import EXTRACTS

SEED = 2026
ACCOUNTS = {  # account: (name, P&L group, statement)
    "1000": ("Cash", "Balance sheet", "BS"), "1200": ("Patient Accounts Receivable", "Balance sheet", "BS"),
    "1500": ("Accumulated Depreciation", "Balance sheet", "BS"), "2000": ("Accounts Payable", "Balance sheet", "BS"),
    "2100": ("Accrued Payroll", "Balance sheet", "BS"),
    "4000": ("Inpatient Revenue", "Gross Patient Revenue", "IS"), "4100": ("Outpatient Revenue", "Gross Patient Revenue", "IS"),
    "4800": ("Contractual Allowances", "Deductions from Revenue", "IS"), "4900": ("Bad Debt & Charity", "Deductions from Revenue", "IS"),
    "5000": ("Salaries - RN", "Salaries & Wages", "IS"), "5010": ("Salaries - Other Clinical", "Salaries & Wages", "IS"),
    "5020": ("Salaries - Non-Clinical", "Salaries & Wages", "IS"), "5100": ("Overtime", "Salaries & Wages", "IS"),
    "5200": ("Contract Labor", "Contract Labor", "IS"), "5300": ("Benefits", "Benefits", "IS"),
    "6000": ("Medical Supplies", "Supplies", "IS"), "6100": ("Implants", "Supplies", "IS"), "6200": ("Drugs", "Drugs", "IS"),
    "6300": ("Purchased Services", "Purchased Services", "IS"), "6400": ("Utilities", "Other Operating Expense", "IS"),
    "6500": ("Repairs & Maintenance", "Other Operating Expense", "IS"), "7000": ("Depreciation", "Depreciation & Interest", "IS"),
    "7100": ("Interest", "Depreciation & Interest", "IS"),
}
FAC_SIZE = {"N": 1.10, "C": 1.30, "E": 0.90, "W": 0.80}
SEASON = [1.10, 1.12, 1.08, 1.00, .95, .90, .88, .90, .95, 1.00, 1.03, 1.06]  # Phoenix winter visitors
# dept: (base monthly volume, gross charge per unit, inpatient share, RN?, hours per unit, fixed hours,
#        supplies per unit, implants per unit, drugs per unit, purchased services / month)
DEPT = {
    "6010": (2400, 4200, 1.00, True, 9.5, 900, 55, 0, 0, 45000),
    "6020": (520, 9800, 1.00, True, 19.0, 400, 210, 0, 0, 38000),
    "6100": (4200, 2600, .40, True, 2.6, 1200, 48, 0, 0, 90000),
    "6200": (620, 21000, .55, True, 14.0, 1500, 2600, 1900, 0, 120000),
    "6300": (5200, 1450, .40, False, .9, 600, 30, 0, 0, 85000),
    "6400": (38000, 95, .40, False, .07, 500, 6, 0, 0, 60000),
    "6500": (52000, 140, .40, False, .05, 700, 0, 0, 38, 30000),
}
STAT = {"6010": "patient_days", "6020": "patient_days", "6100": "ed_visits", "6200": "surgical_cases",
        "6300": "procedures", "6400": "tests", "6500": "doses"}


def periods(start: str, end: str) -> list[str]:
    return [str(p) for p in pd.period_range(start, end, freq="M")]


def _month_index(period: str) -> int:
    y, m = map(int, period.split("-"))
    return (y - 2025) * 12 + (m - 1)


def activity(period: str, fac: str, cc: str, rng, expected: bool = False) -> tuple[dict, dict]:
    """Return ({account: amount}, {stat: value}) for one cost center-month."""
    y, m = map(int, period.split("-"))
    t = _month_index(period)
    noise = (lambda sd: 1.0) if expected else (lambda sd: float(rng.normal(1, sd)))
    amt, stats = {}, {}
    price = 1.025 ** (t / 12)        # charge master increases ~2.5%/yr
    wage = 1.035 ** (t / 12)         # wage inflation ~3.5%/yr
    supply_inf = 1.04 ** (t / 12)

    if cc == "8000":                 # Administration: fixed costs
        s = FAC_SIZE[fac]
        amt["5020"] = 1_100_000 * s * wage * noise(.02)
        amt["5300"] = amt["5020"] * .24
        amt["6300"] = 620_000 * s * noise(.06)
        amt["6400"] = 180_000 * s * SEASON[m - 1] ** 2 * noise(.05)   # Phoenix cooling is not optional
        amt["6500"] = 90_000 * s * noise(.15)
        amt["7000"] = 650_000 * s
        amt["7100"] = 140_000 * s
        return {k: round(v, 2) for k, v in amt.items()}, stats

    base, charge, ip, rn, hpu, fixed_h, sup, imp, drug, purch = DEPT[cc]
    vol = base * FAC_SIZE[fac] * SEASON[m - 1] * 1.03 ** (t / 12) * noise(.03)
    if not expected and (y, m) == (2026, 10) and fac == "W" and cc == "6100":
        vol *= .86                   # competitor urgent care opened
    vol = round(vol)
    stats[STAT[cc]] = vol

    gross = vol * charge * price * noise(.02)
    amt["4000"], amt["4100"] = gross * ip, gross * (1 - ip)
    amt["4800"] = gross * .715 * noise(.012)
    amt["4900"] = gross * .03 * noise(.08)

    hours = (vol * hpu + fixed_h) * noise(.02)
    contract_share = .04
    if not expected and fac == "E" and cc in ("6010", "6020") and (y, m) in ((2026, 9), (2026, 10)):
        contract_share = .14 if m == 9 else .21    # RN vacancies filled by agency nurses
    ot_share = .06
    employed = hours * (1 - contract_share)
    rate = (52 if rn else 38) * wage
    amt["5000" if rn else "5010"] = employed * (1 - ot_share) * rate
    amt["5100"] = employed * ot_share * rate * 1.5
    amt["5200"] = hours * contract_share * 115 * wage
    amt["5300"] = (amt["5000" if rn else "5010"] + amt["5100"]) * .24
    stats["worked_hours"] = round(hours)
    stats["contract_hours"] = round(hours * contract_share)

    if sup:
        amt["6000"] = vol * sup * supply_inf * noise(.04)
    if imp:
        bump = 1.22 if (not expected and (y, m) == (2026, 10) and fac == "N") else 1.0   # vendor price increase
        amt["6100"] = vol * imp * supply_inf * bump * noise(.05)
    if drug:
        amt["6200"] = vol * drug * supply_inf * noise(.04)
    amt["6300"] = purch * FAC_SIZE[fac] * noise(.06)
    return {k: round(v, 2) for k, v in amt.items()}, stats


def journal_lines(period: str, fac: str, cc: str, amt: dict) -> list[dict]:
    """Turn natural-account amounts into balanced journal entries by source module."""
    y, m = map(int, period.split("-"))
    day = f"{period}-{calendar.monthrange(y, m)[1]:02d}"
    ccid = f"{fac}-{cc}"
    lines = []

    def je(src, offsets: list[tuple[str, float]], contra: str, revenue: bool = False):
        jid = f"{src}-{ccid}-{period.replace('-', '')}"
        total = round(sum(a for _, a in offsets), 2)
        if total == 0:
            return
        n = 1
        for acct, a in offsets:
            lines.append({"je_id": jid, "line": n, "period": period, "posting_date": day, "facility": fac,
                          "cost_center": ccid, "account": acct, "debit": 0.0 if revenue else a,
                          "credit": a if revenue else 0.0, "source": src, "description": ACCOUNTS[acct][0]})
            n += 1
        lines.append({"je_id": jid, "line": n, "period": period, "posting_date": day, "facility": fac,
                      "cost_center": ccid, "account": contra, "debit": total if revenue else 0.0,
                      "credit": 0.0 if revenue else total, "source": src, "description": ACCOUNTS[contra][0]})

    pick = lambda *ks: [(k, amt[k]) for k in ks if amt.get(k)]  # noqa: E731
    je("AR", pick("4000", "4100"), "1200", revenue=True)
    je("ARADJ", pick("4800", "4900"), "1200")
    je("PAYROLL", pick("5000", "5010", "5020", "5100", "5300"), "2100")
    je("AP", pick("5200", "6000", "6100", "6200", "6300", "6400", "6500"), "2000")
    je("FA", pick("7000"), "1500")
    je("TREAS", pick("7100"), "1000")
    return lines


def build_month(period: str, settings: dict) -> tuple[pd.DataFrame, pd.DataFrame]:
    rng = np.random.default_rng([SEED, _month_index(period)])
    gl, st = [], []
    for fac in settings["facilities"]:
        for cc in settings["departments"]:
            amt, stats = activity(period, fac, cc, rng)
            gl += journal_lines(period, fac, cc, amt)
            st += [{"period": period, "cost_center": f"{fac}-{cc}", "statistic": k, "value": v} for k, v in stats.items()]
    gl = pd.DataFrame(gl)
    if period == "2026-10":
        # Drug invoice keyed to West Administration instead of West Pharmacy
        gl = pd.concat([gl, pd.DataFrame([
            {"je_id": "AP-W-8000-INV88412", "line": 1, "period": period, "posting_date": "2026-10-31", "facility": "W",
             "cost_center": "W-8000", "account": "6200", "debit": 186400.0, "credit": 0.0, "source": "AP",
             "description": "Drugs - McKesson inv 88412"},
            {"je_id": "AP-W-8000-INV88412", "line": 2, "period": period, "posting_date": "2026-10-31", "facility": "W",
             "cost_center": "W-8000", "account": "2000", "debit": 0.0, "credit": 186400.0, "source": "AP",
             "description": "Accounts Payable"}])], ignore_index=True)
        dup = gl[(gl.je_id == "AP-C-6200-202610") & (gl.line == 1)]
        gl = pd.concat([gl, dup], ignore_index=True)       # extract glitch: one line sent twice
    gl[["debit", "credit"]] = gl[["debit", "credit"]].round(2)
    return gl, pd.DataFrame(st)


def build_budget(year: int, settings: dict) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Budget = expected activity (no noise, no events) under the plan's volume and rate assumptions."""
    rows, st = [], []
    for period in periods(f"{year}-01", f"{year}-12"):
        for fac in settings["facilities"]:
            for cc in settings["departments"]:
                amt, stats = activity(period, fac, cc, None, expected=True)
                rows += [{"period": period, "cost_center": f"{fac}-{cc}", "account": a, "budget": v} for a, v in amt.items()]
                st += [{"period": period, "cost_center": f"{fac}-{cc}", "statistic": k, "budget": v} for k, v in stats.items()]
    return pd.DataFrame(rows), pd.DataFrame(st)


def write_extracts(through: str, settings: dict) -> list[str]:
    """Write any missing monthly GL + statistics extracts, plus the FY budget files. Idempotent."""
    EXTRACTS.mkdir(parents=True, exist_ok=True)
    written = []
    for p in periods(settings["history_start"], through):
        gp = EXTRACTS / f"gl_{p}.csv"
        if gp.exists():
            continue
        gl, st = build_month(p, settings)
        gl.to_csv(gp, index=False)
        st.to_csv(EXTRACTS / f"stats_{p}.csv", index=False)
        written.append(p)
    for year in sorted({int(p[:4]) for p in periods(settings["history_start"], through)}):
        bp = EXTRACTS / f"budget_FY{year}.csv"
        if year >= 2026 and not bp.exists():
            b, bs = build_budget(year, settings)
            b.to_csv(bp, index=False)
            bs.to_csv(EXTRACTS / f"budget_stats_FY{year}.csv", index=False)
    return written
