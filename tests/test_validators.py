import pytest
from validators.raw_validator import RawDataValidator, RawValidationError

def test_hierarchy_validation_passes():
    # Construct a valid parent and children
    parent = {
        "country": "DRC", "currency": "CDF", "merchant": "", "payment_method": "", "provider": "",
        "deposits": 100.0, "deposits_fee": 5.0, "deposits_cost": 2.0,
        "withdrawals": 50.0, "withdrawals_fee": 1.0, "withdrawals_cost": 0.5,
        "settlement": 0.0, "topup": 0.0, "partner_revenue": 3.5
    }
    child1 = {
        "country": "DRC", "currency": "CDF", "merchant": "M1", "payment_method": "", "provider": "",
        "deposits": 60.0, "deposits_fee": 3.0, "deposits_cost": 1.0,
        "withdrawals": 30.0, "withdrawals_fee": 0.6, "withdrawals_cost": 0.3,
        "settlement": 0.0, "topup": 0.0, "partner_revenue": 2.3
    }
    child2 = {
        "country": "DRC", "currency": "CDF", "merchant": "M2", "payment_method": "", "provider": "",
        "deposits": 40.0, "deposits_fee": 2.0, "deposits_cost": 1.0,
        "withdrawals": 20.0, "withdrawals_fee": 0.4, "withdrawals_cost": 0.2,
        "settlement": 0.0, "topup": 0.0, "partner_revenue": 1.2
    }
    RawDataValidator.validate_dashboard_hierarchy([parent, child1, child2])

def test_hierarchy_validation_fails_on_discrepancy():
    parent = {
        "country": "DRC", "currency": "CDF", "merchant": "", "payment_method": "", "provider": "",
        "deposits": 100.0, "deposits_fee": 5.0, "deposits_cost": 2.0,
        "withdrawals": 50.0, "withdrawals_fee": 1.0, "withdrawals_cost": 0.5,
        "settlement": 0.0, "topup": 0.0, "partner_revenue": 3.5
    }
    child1 = {
        "country": "DRC", "currency": "CDF", "merchant": "M1", "payment_method": "", "provider": "",
        "deposits": 50.0, "deposits_fee": 3.0, "deposits_cost": 1.0,
        "withdrawals": 30.0, "withdrawals_fee": 0.6, "withdrawals_cost": 0.3,
        "settlement": 0.0, "topup": 0.0, "partner_revenue": 2.3
    }
    # 50 != 100, should raise
    with pytest.raises(RawValidationError):
        RawDataValidator.validate_dashboard_hierarchy([parent, child1])

def test_duplicate_key_guard():
    leaf1 = {"country": "ZMB", "currency": "ZMW", "merchant": "M1", "payment_method": "MTN", "provider": "P1"}
    leaf2 = {"country": "ZMB", "currency": "ZMW", "merchant": "M1", "payment_method": "MTN", "provider": "P1"}
    with pytest.raises(RawValidationError):
        RawDataValidator.validate_no_duplicate_keys([leaf1, leaf2])

def test_append_only_duplicate_guard():
    existing = {1001, 1002}
    new_records = [{"id": 1002}, {"id": 1003}]
    with pytest.raises(RawValidationError):
        RawDataValidator.validate_append_only_ids(existing, new_records, "Settlements")
