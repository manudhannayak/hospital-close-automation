# Hospital Month-End Close Automation

**Live dashboard: https://manudhannayak.github.io/hospital-close-automation/** · **Monthly packages: [Releases](https://github.com/manudhannayak/hospital-close-automation/releases)**

![Month-end close](https://github.com/manudhannayak/hospital-close-automation/actions/workflows/close.yml/badge.svg)

Hospital finance teams spend the first days of every month pulling the general ledger into Excel, checking it,
comparing actuals to budget, writing variance explanations and emailing reports to dozens of department directors.
This project automates that whole cycle.

On the **3rd business day of each month** it:
1. **Extracts** the general ledger, volume statistics and budget.
2. **Runs close controls**: every journal entry balances, the trial balance ties, no duplicate lines, no postings outside the period,
   every account and department is mapped, accounts sit in the expected departments, and large month-over-month swings are flagged.
   When a fix is mechanical it **proposes the correcting journal entry**.
3. **Analyzes variances with a flexible budget**: each variance is split into **volume** (more or fewer patients) and
   **rate/efficiency** (price, staffing, spending), by facility, department and P&L line.
4. **Writes the commentary**: driver-based explanations built from patient days, visits, cases, agency hours and unit costs.
5. **Generates the package**:
   - a formatted Excel pack with live formulas
   - a PDF executive summary
   - one Excel pack for each of the 32 department directors
   - a distribution manifest
6. **Publishes** it as a GitHub Release and refreshes the CFO dashboard.

> **Fictional health system, synthetic ledger.** Hospital books are private, so `simulate.py` produces ERP-style GL extracts
> (balanced journal entries by facility, cost center and natural account), statistics and a budget for "Desert Valley Health",
> 4 hospitals with 8 departments each. Margins, contractual allowances, labor and supply costs are set at realistic hospital levels.

## What the October 2026 close found

| Finding | How the automation explained it |
|---|---|
| Operating income **$3.49M vs $5.27M budget** (margin 5.3% vs 8.0%) | The bridge splits the $1.78M gap into volume -$75K, price/payer mix -$538K, contract labor -$685K, supplies -$747K, partly offset by employed labor +$300K |
| East Campus Med/Surg **contract labor $483K over budget** | Patient days were in line (+4.3%), but agency hours were 4,852 vs 900 budgeted, so contract nurses were covering vacancies (employed salaries $191K under budget) |
| North Campus Surgery **supplies $448K over** | Surgical cases were *down* 3.5%, but cost per case was $5,640 vs $4,820 (+17%): an implant price increase, not volume |
| West Campus ED **net revenue $301K under** | ED visits 3,158 vs 3,538 budget (-10.7%) explain $263K of it: a volume problem |
| **$186,400 drug invoice in Administration** | Controls flagged the account/department mismatch and proposed the reclass to Pharmacy |
| **Duplicated journal line** in the extract | Controls removed it before the trial balance was checked |

## Parallel run: proof that the automation can replace the manual process

Before a manual report is retired, finance runs both versions for the same month. For September 2026, the automated close
was compared with the analyst's hand-built workbook: **48 figures compared, 45 matched, and all 3 differences were traced
to errors in the manual workbook**, diagnosed automatically:
- **East Campus contract labor:** a transposition error. The difference ($36,000) is divisible by 9 and both numbers use the same digits.
- **Central Campus supplies:** the manual SUM range skipped the Laboratory department. The difference equals that department to the dollar.
- **West Campus net revenue budget:** pasted from a prior budget version, exactly 97.0% of the approved figure.

See `docs/packs/2026-09/Parallel_Run_Reconciliation_2026-09.xlsx`.

## Outputs per close

| File | Contents |
|---|---|
| `Close_Package_YYYY-MM.xlsx` | Cover, Exec Summary, P&L by Facility, Variance Analysis, Department Detail, 13-month KPI Trend with chart, Controls with proposed entries, Notes. Subtotals, Fav/(Unfav) variances and system totals are **live formulas**; blue = loaded values, black = formulas. |
| `Executive_Summary_YYYY-MM.pdf` | Narrative, system P&L, margin trend, top variances with commentary, controls |
| `departments/*.xlsx` (32) | Each director's P&L vs budget, flexed budget and prior year, volume and labor statistics, and "what needs your attention" |
| `distribution_manifest.csv` | Who receives which file. Sending runs in dry-run mode; set `send_mode: smtp` in `config/settings.yaml` |

## Code

| Module | Role |
|---|---|
| `simulate.py` | ERP-style GL, statistics and budget extracts (planted events and data issues are documented in the docstring) |
| `controls.py` | Close controls and proposed correcting entries |
| `model.py` | P&L with subtotals, flexible budget, volume vs rate/efficiency variances, KPIs (adjusted patient days, labor %, contract labor %, supply cost per case...) |
| `commentary.py` | Materiality flags and driver-based explanations |
| `reports.py` | Excel packs (openpyxl) and the PDF (reportlab + matplotlib) |
| `reconcile.py` | Parallel-run reconciliation with automatic diagnosis of differences |
| `schedule.py` | "3rd business day" logic (cron alone can't express it) |
| `pipeline.py` | Orchestration, budget-to-actual bridge, dashboard data |

## Run it

```bash
pip install -r requirements.txt
export PYTHONPATH=src
python -m close.pipeline --period 2026-10     # writes docs/packs/2026-10/
python -m http.server -d docs 8000            # dashboard at localhost:8000
pytest -q                                     # 11 tests
```

To run it on real data, replace `simulate.write_extracts` with your ERP extract (Workday, Oracle or Lawson GL detail with the same
columns: je_id, line, period, posting_date, facility, cost_center, account, debit, credit) and point `config/settings.yaml` at your
chart of accounts and departments.

## Limitations
- The ledger is synthetic and simplified: one entry per module per department, no allocations or accrual reversals.
- Department packs are distributed in dry-run mode (a manifest, no emails sent).
- Flexible-budget variable shares are planning assumptions, documented in the Notes tab.
