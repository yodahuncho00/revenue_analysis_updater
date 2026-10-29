import pytest
from datetime import date
from common.dates import DateManager

def test_regular_date_range():
    dm = DateManager()
    # If today is Oct 28, 2026
    D = date(2026, 10, 28)
    from_date, to_date, rpt_month_end = dm.get_dashboard_date_range(D)
    assert from_date == date(2026, 10, 1)
    assert to_date == date(2026, 10, 27)
    assert rpt_month_end == date(2026, 10, 31)

def test_first_of_month_behavior():
    dm = DateManager(first_of_month_behavior="process_prior_month")
    # If today is Nov 1, 2026
    D = date(2026, 11, 1)
    from_date, to_date, rpt_month_end = dm.get_dashboard_date_range(D)
    assert from_date == date(2026, 10, 1)
    assert to_date == date(2026, 10, 31)
    assert rpt_month_end == date(2026, 10, 31)

def test_working_days_skips_weekends():
    dm = DateManager()
    # Friday Oct 23 to Monday Oct 26
    friday = date(2026, 10, 23)
    monday = date(2026, 10, 26)
    days = dm.get_append_working_days(friday, monday)
    assert days == [monday]  # Skips Sat Oct 24, Sun Oct 25
