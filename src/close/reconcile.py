"""Parallel run: prove the automated close matches the legacy manual workbook, and explain every difference.

Before switching off a manual report, finance runs both side by side for at least one period. Any
difference must be explained. The diagnosis uses classic accounting tests:
  * transposition: the difference is divisible by 9 and the two numbers contain the same digits
  * omitted component: the difference equals one department's amount (a SUM range that missed a row)
  * stale version: the legacy value is a constant ratio off the system value (an old budget version)
"""
from __future__ import annotations

import json
from pathlib import Path

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill

from . import LEGACY
from .reports import BOLD, CALC, HEAD, FILL_HEAD, USD

METRICS = ["Net Patient Revenue", "Salaries & Wages", "Contract Labor", "Supplies", "Drugs", "Operating Income"]
LEGACY_FILE = "Sep2026 Close Package FINAL v3 (manual).xlsx"


def build_legacy(fac_pl: pd.DataFrame, dept_v: pd.DataFrame, fac_names: dict) -> Path:
    """Recreate what the analyst's hand-built September workbook contained (values pasted, three human errors)."""
    LEGACY.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = "Summary"
    ws["A1"] = "Sep 2026 Close - Facility Summary (manual, values pasted from GL query + budget file)"
    ws["A1"].font = BOLD
    ws.append([])
    ws.append(["Facility", "Metric", "Actual", "Budget"])
    for f, name in fac_names.items():
        d = fac_pl[fac_pl.facility == f].set_index("line")
        for m in METRICS:
            a, b = round(float(d.loc[m, "actual"])), round(float(d.loc[m, "budget"]))
            if f == "E" and m == "Contract Labor":           # error 1: keyed with two digits swapped
                s = str(a)
                a = int(s[0] + s[2] + s[1] + s[3:])
            if f == "C" and m == "Supplies":                  # error 2: SUM range missed Laboratory
                lab = dept_v[(dept_v.cost_center == "C-6400") & (dept_v.line == "Supplies")].actual.sum()
                a = round(a - lab)
            if f == "W" and m == "Net Patient Revenue":       # error 3: budget pasted from the prior budget version
                b = round(b * .97)
            ws.append([name, m, a, b])
    p = LEGACY / LEGACY_FILE
    wb.save(p)
    return p


def _same_digits(a: int, b: int) -> bool:
    return sorted(str(abs(a))) == sorted(str(abs(b)))


def reconcile(fac_pl: pd.DataFrame, dept_v: pd.DataFrame, fac_names: dict, out_dir: Path) -> dict:
    p = LEGACY / LEGACY_FILE
    if not p.exists():
        build_legacy(fac_pl, dept_v, fac_names)
    ws = load_workbook(p, data_only=True)["Summary"]
    rows = [r for r in ws.iter_rows(min_row=4, values_only=True) if r[0]]
    by_name = {v: k for k, v in fac_names.items()}
    items = []
    for name, metric, la, lb in rows:
        f = by_name[name]
        d = fac_pl[(fac_pl.facility == f)].set_index("line").loc[metric]
        for kind, legacy_v, sys_v in (("Actual", la, d.actual), ("Budget", lb, d.budget)):
            sys_v = round(float(sys_v))
            diff = legacy_v - sys_v
            cause, explained = "Match", True
            if abs(diff) > 1:
                explained = False
                comp = dept_v[(dept_v.facility == f) & (dept_v.line == metric)]
                hit = comp[(comp[kind.lower()] - abs(diff)).abs() <= 1]
                if diff % 9 == 0 and _same_digits(legacy_v, sys_v):
                    cause, explained = "Transposition error in manual entry (difference divisible by 9, same digits)", True
                elif len(hit):
                    cause, explained = f"Manual SUM range omitted {hit.cost_center.iloc[0]} (difference equals that department exactly)", True
                elif sys_v and 0.9 < legacy_v / sys_v < 1.0 and kind == "Budget":
                    cause, explained = f"Legacy used a prior budget version ({legacy_v / sys_v:.1%} of approved budget)", True
                else:
                    cause = "UNEXPLAINED - investigate"
            items.append({"facility": name, "metric": metric, "type": kind, "legacy": legacy_v, "automated": sys_v,
                          "difference": diff, "cause": cause, "explained": explained})
    df = pd.DataFrame(items)
    diffs = df[df.cause != "Match"]
    # Downstream subtotal differences caused by the same errors (e.g. Operating Income) are expected
    summary = {"period": "2026-09", "compared": int(len(df)), "matched": int((df.cause == "Match").sum()),
               "differences": int(len(diffs)), "explained": int(diffs.explained.sum()),
               "unexplained": int((~diffs.explained).sum()), "items": diffs.to_dict("records")}

    wb = Workbook()
    ws = wb.active
    ws.title = "Parallel Run"
    ws["A1"] = "Parallel run reconciliation: automated close vs legacy manual workbook, September 2026"
    ws["A1"].font = Font(name="Arial", size=13, bold=True, color="1F3A5F")
    ws["A2"] = (f"{summary['compared']} figures compared · {summary['matched']} match · {summary['differences']} differ "
                f"({summary['explained']} explained by diagnosis rules, {summary['unexplained']} unexplained)")
    heads = ["Facility", "Metric", "Actual/Budget", "Legacy (manual)", "Automated", "Difference", "Diagnosis"]
    for j, h in enumerate(heads, start=1):
        c = ws.cell(4, j, h)
        c.font, c.fill = HEAD, FILL_HEAD
    for i, r in enumerate(df.itertuples(), start=5):
        for j, x in enumerate([r.facility, r.metric, r.type, r.legacy, r.automated], start=1):
            c = ws.cell(i, j, x)
            c.font = CALC
            if j >= 4:
                c.number_format = USD
        ws.cell(i, 6, f"=D{i}-E{i}").number_format = USD
        c = ws.cell(i, 7, r.cause)
        c.font = BOLD if r.cause != "Match" else CALC
        if r.cause != "Match":
            for j in range(1, 8):
                ws.cell(i, j).fill = PatternFill("solid", fgColor="FFF2CC")
    for j, w in enumerate([16, 22, 13, 16, 16, 14, 80], start=1):
        ws.column_dimensions[chr(64 + j)].width = w
    wb.calculation.fullCalcOnLoad = True
    out_dir.mkdir(parents=True, exist_ok=True)
    wb.save(out_dir / "Parallel_Run_Reconciliation_2026-09.xlsx")
    (out_dir / "parallel_run.json").write_text(json.dumps(summary, indent=2, default=str))
    return summary
