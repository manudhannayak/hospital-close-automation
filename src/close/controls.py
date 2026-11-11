"""Close controls run before any report is built.

FAIL stops the close (reports are not distributed). WARN is listed in the controls tab and the
exec summary. Where the fix is mechanical, the control proposes the correcting entry (for example
a reclass journal) instead of just raising a flag.
"""
from __future__ import annotations

from dataclasses import asdict, dataclass

import pandas as pd

from .simulate import ACCOUNTS

# Natural accounts that only belong in certain departments
EXPECTED_HOME = {"6200": ("6500", "Drugs belong in Pharmacy"), "6100": ("6200", "Implants belong in Surgical Services")}
REVENUE_ACCTS = {"4000", "4100", "4800", "4900"}


@dataclass
class Control:
    control: str
    status: str
    detail: str
    dollars: float = 0.0


def run_controls(gl: pd.DataFrame, stats: pd.DataFrame, period: str, prior: pd.DataFrame,
                 settings: dict) -> tuple[pd.DataFrame, list[Control], pd.DataFrame]:
    out: list[Control] = []
    reclass = []
    gl = gl.copy()

    # 1. Exact duplicate lines (extract glitch): remove, report the dollars
    dup = gl.duplicated(subset=["je_id", "line"], keep="first")
    if dup.any():
        amt = float((gl.loc[dup, "debit"] + gl.loc[dup, "credit"]).sum())
        out.append(Control("duplicate_lines", "WARN", f"{int(dup.sum())} duplicated journal line(s) removed: "
                           f"{', '.join(gl.loc[dup, 'je_id'].unique())}", amt))
        gl = gl[~dup]
    else:
        out.append(Control("duplicate_lines", "PASS", "no duplicated journal lines"))

    # 2. Every journal entry balances
    je = gl.groupby("je_id")[["debit", "credit"]].sum()
    unbalanced = je[(je.debit - je.credit).abs() > .005]
    out.append(Control("journal_entries_balance", "FAIL" if len(unbalanced) else "PASS",
                       f"{len(unbalanced)} unbalanced of {len(je)} journal entries"
                       + (f": {', '.join(unbalanced.index[:5])}" if len(unbalanced) else ""),
                       float((unbalanced.debit - unbalanced.credit).abs().sum())))

    # 3. Trial balance: total debits = total credits
    diff = round(gl.debit.sum() - gl.credit.sum(), 2)
    out.append(Control("trial_balance", "PASS" if abs(diff) < .01 else "FAIL",
                       f"debits ${gl.debit.sum():,.2f} vs credits ${gl.credit.sum():,.2f} (difference ${diff:,.2f})", abs(diff)))

    # 4. Mapping: every account and cost center is known
    bad_acct = set(gl.account) - set(ACCOUNTS)
    valid_cc = {f"{f}-{c}" for f in settings["facilities"] for c in settings["departments"]}
    bad_cc = set(gl.cost_center) - valid_cc
    out.append(Control("chart_of_accounts_mapping", "FAIL" if bad_acct or bad_cc else "PASS",
                       f"{len(bad_acct)} unmapped accounts, {len(bad_cc)} unmapped cost centers"))

    # 5. Postings dated inside the period
    outside = (~gl.posting_date.astype(str).str.startswith(period)).sum()
    out.append(Control("period_cutoff", "FAIL" if outside else "PASS", f"{int(outside)} lines dated outside {period}"))

    # 6. Account-to-department reasonableness -> proposed reclass entries
    found = []
    for acct, (home_cc, why) in EXPECTED_HOME.items():
        bad = gl[(gl.account == acct) & (gl.debit > 0) & ~gl.cost_center.str.endswith(home_cc)]
        for r in bad.itertuples():
            target = f"{r.cost_center.split('-')[0]}-{home_cc}"
            found.append(f"{r.je_id} ${r.debit:,.0f} in {r.cost_center}")
            reclass += [{"period": period, "account": acct, "cost_center": r.cost_center, "debit": 0.0,
                         "credit": r.debit, "memo": f"Reclass {r.je_id}: {why}"},
                        {"period": period, "account": acct, "cost_center": target, "debit": r.debit,
                         "credit": 0.0, "memo": f"Reclass {r.je_id}: {why}"}]
    rev_admin = gl[gl.account.isin(REVENUE_ACCTS) & gl.cost_center.str.endswith("8000")]
    found += [f"revenue posted to {cc}" for cc in rev_admin.cost_center.unique()]
    out.append(Control("account_department_fit", "WARN" if found else "PASS",
                       ("possible misposting, reclass proposed: " + "; ".join(found)) if found else "all postings in expected departments",
                       float(sum(r["debit"] for r in reclass))))

    # 7. Statistics complete for every revenue department
    rev_cc = {f"{f}-{c}" for f in settings["facilities"] for c, d in settings["departments"].items() if d["type"] == "revenue"}
    missing = rev_cc - set(stats.cost_center)
    out.append(Control("statistics_complete", "FAIL" if missing else "PASS", f"{len(missing)} departments missing volume statistics"))

    # 8. Analytical review: big month-over-month swings by facility x account
    if prior is not None and len(prior):
        def net(df):
            return df[df.account.str[0].isin(list("4567"))].assign(v=lambda d: d.debit - d.credit) \
                .groupby(["facility", "account"]).v.sum()
        cur, prv = net(gl), net(prior)
        j = pd.concat([cur.rename("cur"), prv.rename("prv")], axis=1).fillna(0)
        j["chg"] = j.cur - j.prv
        big = j[(j.chg.abs() > 250_000) & (j.chg.abs() > .4 * j.prv.abs())]
        out.append(Control("month_over_month_fluctuation", "WARN" if len(big) else "PASS",
                           "; ".join(f"{f} {ACCOUNTS[a][0]} {r.chg:+,.0f}" for (f, a), r in big.iterrows())
                           or "no swings over $250K and 40%", float(big.chg.abs().sum())))
    return gl, out, pd.DataFrame(reclass)


def summarize(results: list[Control]) -> dict:
    counts = {s: sum(r.status == s for r in results) for s in ("PASS", "WARN", "FAIL")}
    return {"counts": counts, "passed": counts["FAIL"] == 0, "results": [asdict(r) for r in results]}
