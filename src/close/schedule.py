"""Business-day gate for the scheduler: cron can't express "3rd business day", so the workflow runs
on days 1-7 and this decides whether today is the day.

    python -m close.schedule        # prints the period to close, or nothing
"""
from __future__ import annotations

import sys
from datetime import date, datetime, timedelta
from zoneinfo import ZoneInfo

import pandas as pd

from . import load_settings

# US federal holidays that can fall in the first week of a month
HOLIDAYS = {(1, 1), (7, 4)}


def nth_business_day(year: int, month: int, n: int) -> date:
    d, count = date(year, month, 1), 0
    while True:
        if d.weekday() < 5 and (d.month, d.day) not in HOLIDAYS and not (d.month == 9 and d.weekday() == 0 and d.day <= 7):
            count += 1                       # (first Monday of September = Labor Day)
            if count == n:
                return d
        d += timedelta(days=1)


def period_due(today: date) -> str | None:
    if today == nth_business_day(today.year, today.month, 3):
        return str(pd.Period(today, freq="M") - 1)
    return None


if __name__ == "__main__":
    s = load_settings()
    p = period_due(datetime.now(ZoneInfo(s["timezone"])).date())
    if p:
        print(p)
    sys.exit(0)
