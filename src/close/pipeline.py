"""Monthly close run.

    python -m close.pipeline                    # close last month (scheduled on the 3rd business day)
    python -m close.pipeline --period 2026-10   # re-run any period
"""
from __future__ import annotations

import argparse
import csv
import hashlib
import json
import logging
import os
import sys
from datetime import date, datetime, timezone
from zoneinfo import ZoneInfo

import pandas as pd

from . import EXTRACTS, SITE, load_settings
from .commentary import exec_narrative, flag_variances
from .controls import run_controls, summarize
from .model import PL_LINES, REVENUE_LINES, actuals, budget_for, kpis, load_period, roll_up, variance_table
from .reconcile import reconcile
from .reports import department_pack, pdf_summary, system_pack
from .simulate import periods, write_extracts

log = logging.getLogger("close")


def last_closed_month(settings: dict) -> str:
    today = datetime.now(ZoneInfo(settings["timezone"])).date()
    return str(pd.Period(today, freq="M") - 1)


def label(period: str) -> str:
    return pd.Period(period, freq="M").strftime("%B %Y")


def _clean(o):
    """JSON for browsers: NaN -> null, numpy types -> Python."""
    if isinstance(o, dict):
        return {k: _clean(v) for k, v in o.items()}
    if isinstance(o, (list, tuple)):
        return [_clean(v) for v in o]
    if hasattr(o, "item"):
        o = o.item()
    if isinstance(o, float) and o != o:
        return None
    return o


def clean(gl: pd.DataFrame) -> pd.DataFrame:
    return gl.drop_duplicates(subset=["je_id", "line"])


def bridge(sys_v: pd.DataFrame) -> list[dict]:
    """Budget operating income -> actual, in Fav/(Unfav) steps: volume, then rate/efficiency by line."""
    d = sys_v.set_index("line")
    steps = [{"label": "Budget op. income", "value": float(d.loc["Operating Income", "budget"]), "kind": "total"}]
    detail = ["Gross Patient Revenue", "Deductions from Revenue", "Salaries & Wages", "Contract Labor", "Benefits",
              "Supplies", "Drugs", "Purchased Services", "Other Operating Expense", "Depreciation & Interest"]
    sign = {l: (1 if l in REVENUE_LINES else -1) for l in detail}
    vol = sum(sign[l] * d.loc[l, "volume_variance"] for l in detail)
    steps.append({"label": "Patient volume", "value": float(vol), "kind": "delta"})
    groups = {"Price & payer mix": ["Gross Patient Revenue", "Deductions from Revenue"],
              "Employed labor": ["Salaries & Wages", "Benefits"], "Contract labor": ["Contract Labor"],
              "Supplies & implants": ["Supplies"], "Drugs": ["Drugs"],
              "Services & other": ["Purchased Services", "Other Operating Expense", "Depreciation & Interest"]}
    for g, ls in groups.items():
        steps.append({"label": g, "value": float(sum(sign[l] * d.loc[l, "rate_variance"] for l in ls)), "kind": "delta"})
    steps.append({"label": "Actual op. income", "value": float(d.loc["Operating Income", "actual"]), "kind": "total"})
    return steps


def run(period: str) -> int:
    s = load_settings()
    fac_names = s["facilities"]
    org = s["organization"]
    write_extracts(period, s)            # simulated ERP drop; in production this is the ERP extract job
    gl_raw, stats = load_period(period)
    prev = str(pd.Period(period, freq="M") - 1)
    prior_gl = clean(load_period(prev)[0]) if (EXTRACTS / f"gl_{prev}.csv").exists() else None

    gl, results, reclass = run_controls(gl_raw, stats, period, prior_gl, s)
    ctrl = summarize(results)
    log.info("Controls %s", ctrl["counts"])
    out = SITE / "packs" / period
    out.mkdir(parents=True, exist_ok=True)
    if not ctrl["passed"]:
        (out / "controls_FAILED.json").write_text(json.dumps(ctrl, indent=2))
        log.error("Close controls FAILED; package not distributed")
        return 1

    act = actuals(gl)
    py_period = str(pd.Period(period, freq="M") - 12)
    py = actuals(clean(load_period(py_period)[0])) if (EXTRACTS / f"gl_{py_period}.csv").exists() else None
    v = variance_table(period, act, stats, py, s)
    v["facility"] = v.cost_center.str[0]
    fac_pl = roll_up(v, ["facility"])
    sys_pl = roll_up(v.assign(system="All"), ["system"])
    _, bstats = budget_for(period)
    flags = flag_variances(v, stats, bstats, s, fac_names, gl)

    hist = periods(str(pd.Period(period, freq="M") - 12), period)
    all_act = pd.concat([actuals(clean(load_period(p)[0])) for p in hist])
    all_stats = pd.concat([load_period(p)[1] for p in hist])
    all_gl = pd.concat([clean(load_period(p)[0]) for p in hist])
    kpi = kpis(all_act, all_stats, all_gl)

    narrative = exec_narrative(label(period), sys_pl, flags, kpi, ctrl)
    files = sorted(EXTRACTS.glob(f"*{period}*.csv"))
    run_info = {"period": period, "generated_at": datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC"),
                "gl_rows": int(len(gl_raw)), "stat_rows": int(len(stats)),
                "input_sha256": hashlib.sha256(b"".join(f.read_bytes() for f in files)).hexdigest()}

    sys_file = out / f"Close_Package_{period}.xlsx"
    system_pack(sys_file, label(period), org, fac_names, fac_pl, flags, v, kpi, ctrl, reclass, narrative, s, run_info)
    pdf_summary(out / f"Executive_Summary_{period}.pdf", label(period), org, narrative, sys_pl, flags, kpi, fac_names, ctrl, out)
    (out / "margin_trend.png").unlink(missing_ok=True)

    dept_dir = out / "departments"
    dept_dir.mkdir(exist_ok=True)
    manifest = []
    dom = s["distribution"]["email_domain"]
    for f, fname in fac_names.items():
        for cc, d in s["departments"].items():
            ccid = f"{f}-{cc}"
            owner = f"Director, {d['name']} - {fname}"
            email = f"{d['name'].lower().split()[0].replace('/', '')}.{fname.split()[0].lower()}@{dom}"
            fn = f"{ccid}_{d['name'].replace('/', '-').replace(' ', '_')}_{period}.xlsx"
            department_pack(dept_dir / fn, label(period), org, ccid, d["name"], fname, owner, v, stats, bstats, flags)
            manifest.append({"recipient": owner, "email": email, "file": f"departments/{fn}",
                             "material_variances": int((flags.cost_center == ccid).sum()),
                             "status": "queued (dry run)" if s["distribution"]["send_mode"] == "dry_run" else "sent"})
    for f, fname in fac_names.items():
        manifest.append({"recipient": f"CFO, {fname}", "email": f"cfo.{fname.split()[0].lower()}@{dom}",
                         "file": sys_file.name, "material_variances": int((flags.facility == f).sum()), "status": "queued (dry run)"})
    manifest.append({"recipient": "System CFO", "email": f"cfo@{dom}", "file": f"Executive_Summary_{period}.pdf",
                     "material_variances": int(len(flags)), "status": "queued (dry run)"})
    with open(out / "distribution_manifest.csv", "w", newline="") as fh:
        w = csv.DictWriter(fh, fieldnames=list(manifest[0]))
        w.writeheader()
        w.writerows(manifest)

    # Parallel run against the legacy manual workbook (September 2026 was run both ways)
    sep = "2026-09"
    if period >= sep:
        sg, sst = load_period(sep)
        sv = variance_table(sep, actuals(clean(sg)), sst, None, s)
        sv["facility"] = sv.cost_center.str[0]
        parallel = reconcile(roll_up(sv, ["facility"]), sv, fac_names, SITE / "packs" / sep)
    else:
        parallel = None

    # Dashboard data
    grid = v[v.line.isin([l for l in PL_LINES if l not in {"Gross Patient Revenue", "Deductions from Revenue",
                                                          "Total Operating Expense", "EBIDA"}])].copy()
    grid["fav"] = grid.apply(lambda r: r.variance if r.line in REVENUE_LINES else -r.variance, axis=1)
    site = {
        "period": period, "label": label(period), "org": org, "generated_at": run_info["generated_at"],
        "narrative": narrative, "controls": ctrl, "reclass": reclass.to_dict("records"),
        "bridge": bridge(sys_pl),
        "system_pl": sys_pl.assign(line=sys_pl.line.astype(str)).drop(columns="system").to_dict("records"),
        "facility_pl": fac_pl.assign(line=fac_pl.line.astype(str)).to_dict("records"),
        "grid": grid[["cost_center", "line", "actual", "budget", "flexed_budget", "volume_variance", "rate_variance", "fav"]].round(0).to_dict("records"),
        "flags": flags.round(1).to_dict("records"), "kpi": kpi.round(2).to_dict("records"),
        "facilities": fac_names, "departments": {k: v_["name"] for k, v_ in s["departments"].items()},
        "parallel_run": parallel, "manifest_count": len(manifest), "run_info": run_info,
    }
    (out / "close.json").write_text(json.dumps(_clean(site), default=float, indent=1, allow_nan=False))
    lib = sorted([p.name for p in (SITE / "packs").iterdir() if (p / "close.json").exists()], reverse=True)
    (SITE / "library.json").write_text(json.dumps({"periods": lib, "latest": lib[0] if lib else None}))
    log.info("\n%s", narrative)
    log.info("Wrote %s department packs, system pack, PDF; %s recipients queued", len(s["departments"]) * len(fac_names), len(manifest))
    return 0


def main() -> None:
    p = argparse.ArgumentParser()
    p.add_argument("--period", default=None, help="YYYY-MM (default: last closed month)")
    a = p.parse_args()
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(name)s: %(message)s")
    logging.getLogger("matplotlib").setLevel(logging.WARNING)
    sys.exit(run(a.period or last_closed_month(load_settings())))


if __name__ == "__main__":
    main()
