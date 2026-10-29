import pytest
import openpyxl
from datetime import datetime, date

from main import (
    ALLOWED_STEPS,
    _earliest_missing_transaction_detail_date,
    run_pipeline,
)


@pytest.mark.parametrize("removed_step", ["3", "providers_balance"])
def test_provider_balance_step_is_rejected(removed_step):
    assert removed_step not in ALLOWED_STEPS
    with pytest.raises(ValueError, match="Provider Balance is not part of the ETL"):
        run_pipeline(step=removed_step)


def test_earliest_missing_transaction_detail_date_ignores_complete_and_future_rows():
    workbook = openpyxl.Workbook()
    worksheet = workbook.active
    worksheet.append(["Id", "Created", None, None, None, None, None, None, None, "Order", "Note"])
    worksheet.append([1, datetime(2025, 6, 1), None, None, None, None, None, None, None, None, "Note"])
    worksheet.append([2, datetime(2025, 5, 1), None, None, None, None, None, None, None, "Order", None])
    worksheet.append([3, datetime(2026, 10, 29), None, None, None, None, None, None, None, None, None])

    assert _earliest_missing_transaction_detail_date(
        worksheet, date(2026, 10, 28)
    ) == date(2025, 5, 1)

    workbook.close()
