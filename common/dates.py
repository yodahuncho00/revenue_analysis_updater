from datetime import datetime, date, timedelta, timezone
from typing import List, Tuple, Optional
import calendar

# Business timezone: UTC+3
BUSINESS_TIMEZONE = timezone(timedelta(hours=3))

class DateManager:
    def __init__(self, tz: timezone = BUSINESS_TIMEZONE, first_of_month_behavior: str = "process_prior_month"):
        self.tz = tz
        self.first_of_month_behavior = first_of_month_behavior
        self.holidays = set()  # Extensible holiday calendar set of date objects

    def add_holidays(self, holiday_dates: List[date]):
        self.holidays.update(holiday_dates)

    def is_working_day(self, d: date) -> bool:
        """Working days are Monday (0) to Friday (4), excluding holidays."""
        if d.weekday() >= 5:  # Saturday or Sunday
            return False
        if d in self.holidays:
            return False
        return True

    def get_effective_today(self, override_date: Optional[date] = None) -> date:
        """Returns today in the business timezone, or the override date."""
        if override_date is not None:
            return override_date
        return datetime.now(self.tz).date()

    def get_dashboard_date_range(self, effective_today: Optional[date] = None) -> Tuple[date, date, date]:
        """
        Calculates (from_date, to_date, rpt_month_end).
        - D = effective_today
        - to_date = D - 1 day
        - If D is the 1st day of the month:
            from_date = 1st of prior month
            to_date = last day of prior month
            rpt_month_end = last day of prior month
        - Otherwise:
            from_date = 1st of D's month
            to_date = D - 1 day
            rpt_month_end = last day of D's month
        """
        D = self.get_effective_today(effective_today)
        
        if D.day == 1:
            # On the 1st of the month, process the full prior month
            last_day_prior_month = D - timedelta(days=1)
            from_date = date(last_day_prior_month.year, last_day_prior_month.month, 1)
            to_date = last_day_prior_month
            rpt_month_end = last_day_prior_month
        else:
            from_date = date(D.year, D.month, 1)
            to_date = D - timedelta(days=1)
            # Find last day of current month
            _, last_day = calendar.monthrange(D.year, D.month)
            rpt_month_end = date(D.year, D.month, last_day)

        return from_date, to_date, rpt_month_end

    def get_append_working_days(self, latest_date: date, effective_today: Optional[date] = None) -> List[date]:
        """
        Returns all working days in range [latest_date + 1 day, D].
        Skips Saturdays, Sundays, and configured holidays.
        """
        D = self.get_effective_today(effective_today)
        start_date = latest_date + timedelta(days=1)
        
        if start_date > D:
            return []

        working_days = []
        cur = start_date
        while cur <= D:
            if self.is_working_day(cur):
                working_days.append(cur)
            cur += timedelta(days=1)
            
        return working_days
