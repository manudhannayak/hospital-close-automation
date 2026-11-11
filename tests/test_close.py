from datetime import date

import pandas as pd
import pytest
from openpyxl import load_workbook

from close import load_settings
from close.commentary import flag_variances
from close.controls import run_controls, summarize
from close.model import PL_LINES, actuals, budget_for, roll_up, variance_table
from close.pipeline import bridge, clean
from close.schedule import nth_business_day, period_due
from close.simulate import build_month, write_extracts

P = "2026-10"


@pytest.fixture(scope="module")
def s():
    st = load_settings()
    write_extracts(P, st)
    return st


@pytest.fixture(scope="module")
def month(s):
    gl, stats = build_month(P, s)
    prior, _ = build_month("2026-09", s)
    return gl, stats, prior


@pytest.fixture(scope="module")
def closed(s, month):
    gl, stats, prior = month
    gl2, res, reclass = run_controls(gl, stats, P, clean(prior), s)
    v = variance_table(P, actuals(gl2), stats, actuals(clean(build_month("2025-10", s)[0])), s)
    v["facility"] = v.cost_center.str[0]
    return gl2, res, reclass, v


def test_ledger_is_deterministic(s):
    a, _ = build_month("2026-03", s)
    b, _ = build_month("2026-03", s)
    pd.testing.assert_frame_equal(a, b)


def test_margins_realistic(s, closed):
    v = closed[3]
    sys = roll_up(v.assign(x=1), ["x"]).set_index("line")
    margin = sys.loc["Operating Income", "actual"] / sys.loc["Net Patient Revenue", "actual"]
    assert .0 < margin < .12          # hospitals run thin margins


def test_controls_catch_planted_issues(closed):
    _, res, reclass, _ = closed
    r = {c.control: c for c in res}
    assert r["duplicate_lines"].status == "WARN"
    assert r["trial_balance"].status == "PASS"           # balances once the duplicate is removed
    assert r["account_department_fit"].status == "WARN" and "INV88412" in r["account_department_fit"].detail
    assert round(reclass.debit.sum(), 2) == round(reclass.credit.sum(), 2) == 186400.0
    assert summarize(res)["passed"]


def test_unbalanced_entry_fails_close(s, month):
    gl, stats, prior = month
    bad = gl.copy()
    bad.loc[bad.index[0], "debit"] += 100
    _, res, _ = run_controls(bad, stats, P, clean(prior), s)
    assert not summarize(res)["passed"]


def test_unmapped_account_fails_close(s, month):
    gl, stats, prior = month
    bad = gl.copy()
    bad.loc[bad.index[0], "account"] = "9999"
    _, res, _ = run_controls(bad, stats, P, clean(prior), s)
    assert not summarize(res)["passed"]


def test_variance_decomposition_adds_up(closed):
    v = closed[3]
    assert (v.volume_variance + v.rate_variance - v.variance).abs().max() < .01


def test_subtotals(closed):
    v = closed[3]
    g = v.groupby("line").actual.sum()
    assert g["Net Patient Revenue"] == pytest.approx(g["Gross Patient Revenue"] - g["Deductions from Revenue"])
    assert g["Operating Income"] == pytest.approx(g["EBIDA"] - g["Depreciation & Interest"])


def test_bridge_reconciles(closed):
    v = closed[3]
    b = bridge(roll_up(v.assign(x=1), ["x"]))
    assert b[0]["value"] + sum(x["value"] for x in b[1:-1]) == pytest.approx(b[-1]["value"], abs=1)


def test_commentary_finds_planted_events(s, closed, month):
    gl2, _, _, v = closed
    _, bstats = budget_for(P)
    flags = flag_variances(v, month[1], bstats, s, s["facilities"], gl2)
    find = lambda f, cc, line: flags[(flags.cost_center == f"{f}-{cc}") & (flags.line == line)]  # noqa: E731
    cl = find("E", "6010", "Contract Labor")
    assert len(cl) and cl.driver.iloc[0] == "Rate / efficiency" and "Agency hours" in cl.commentary.iloc[0]
    imp = find("N", "6200", "Supplies")
    assert len(imp) and "price or mix" in imp.commentary.iloc[0]
    ed = find("W", "6100", "Net Patient Revenue")
    assert len(ed) and ed.driver.iloc[0] == "Volume"
    assert len(find("W", "8000", "Drugs"))


def test_business_day_schedule():
    assert nth_business_day(2026, 11, 3) == date(2026, 11, 4)    # Nov 1-2 2026: Sun, Mon
    assert period_due(date(2026, 11, 4)) == "2026-10"
    assert period_due(date(2026, 11, 5)) is None
    assert nth_business_day(2026, 9, 3) == date(2026, 9, 3)      # Sep 7 Labor Day not in play; Sep 1-3 Tue-Thu


def test_generated_packs(tmp_path, s):
    from close import pipeline as pl
    import close
    pl.SITE = close.SITE = tmp_path
    assert pl.run(P) == 0
    out = tmp_path / "packs" / P
    assert len(list((out / "departments").glob("*.xlsx"))) == 32
    assert (out / f"Executive_Summary_{P}.pdf").stat().st_size > 20_000
    wb = load_workbook(out / f"Close_Package_{P}.xlsx")
    ws = wb["P&L by Facility"]
    formulas = sum(1 for row in ws.iter_rows() for c in row if isinstance(c.value, str) and c.value.startswith("="))
    assert formulas == 4 * 36 + 14 * 4                            # facility subtotals + variances, and system totals, are live formulas
    pr = (tmp_path / "packs" / "2026-09" / "parallel_run.json").read_text()
    assert '"unexplained": 0' in pr and '"differences": 3' in pr
