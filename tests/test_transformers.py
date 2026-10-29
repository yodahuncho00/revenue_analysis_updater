import pytest
from datetime import date, datetime
from transformers.bo_extract import BOExtractTransformer
from transformers.data_sync import DataSheetSyncTransformer
from transformers.append_only import AppendOnlyTransformer

def test_bo_extract_transformer_types_and_filter():
    raw_row = {
        "country": "DRC", "currency": "CDF", "merchant": "Lag DRC", "payment_method": "Airtel", "provider": "PayGo (93)",
        "deposits": "1,500.50", "deposits_fee": "45.00", "fee_pct_dep": "3.0%",
        "deposits_cost": "30.00", "cost_pct_dep": "2.0%",
        "withdrawals": "500", "withdrawals_fee": "10", "fee_pct_wd": "2%",
        "withdrawals_cost": "5", "cost_pct_wd": "1%",
        "settlement": "0", "topup": "0", "merchant_delta": "1445.50",
        "provider_delta": "1465.50", "partner_revenue": "20.00"
    }
    rpt_end = date(2026, 10, 31)
    transformed = BOExtractTransformer.transform_rows([raw_row], rpt_end)
    assert len(transformed) == 1
    row = transformed[0]
    assert row[0] == datetime(2026, 10, 31, 0, 0)
    assert row[1] == "M"
    assert row[2] == "DRC"
    assert row[7] == 1500.50
    assert row[23] == '=IF(G{row}="",0,1)'

def test_data_sync_keeps_only_leaf_rows():
    parent_row = [datetime(2026, 10, 31), "M", "DRC", "CDF", "Lag DRC", "Airtel", "", 100, 5, 5, 2, 2, 0, 0, 0, 0, 0, 0, 0, 0, 0, 3]
    leaf_row = [datetime(2026, 10, 31), "M", "DRC", "CDF", "Lag DRC", "Airtel", "PayGo (93)", 100, 5, 5, 2, 2, 0, 0, 0, 0, 0, 0, 0, 0, 0, 3]
    
    data_rows = DataSheetSyncTransformer.filter_and_map_rows([parent_row, leaf_row])
    assert len(data_rows) == 1
    assert data_rows[0][6] == "PayGo (93)"
    assert data_rows[0][9] == '=IFERROR(I{row}/H{row}*100,0)'

def test_append_only_sorting_and_idempotency():
    records = [
        {"id": 200, "created": "2026-10-25 10:00:00", "amount": 500},
        {"id": 100, "created": "2026-10-24 10:00:00", "amount": 250}
    ]
    existing_ids = {50}
    transformed = AppendOnlyTransformer.transform_records(records, existing_ids, "Settlements")
    assert len(transformed) == 2
    assert transformed[0][0] == 100  # Sorted ascending
    assert transformed[1][0] == 200
    assert transformed[0][11] == '=YEAR(B{row})'
