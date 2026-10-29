import openpyxl
from datetime import datetime, date
from typing import Dict, Any, List

class ReconciliationError(Exception):
    pass

class PostWriteReconciler:
    """Performs rigorous post-write reconciliation against raw captures and workbook integrity gates."""

    # Rows on Control tab that are strict math/schema integrity gates
    STRICT_CONTROL_GATES = {
        68: "Rows with missing FX rate",
        74: "Years carrying both FY and monthly rows",
        75: "Revenue identity: fees - costs - Partner Revenue",
        76: "Reporting month present in Data",
        77: "Zoyk + Unipesa - Partner Revenue (USD)",
        78: "Direction split: deposit + withdrawal net revenue - Partner Revenue"
    }

    # Rows that are informational master-data mappings
    INFORMATIONAL_MAPPING_ROWS = {
        69: "Merchants in Data missing from Merchant_master",
        70: "Providers in Data missing from Provider_master",
        71: "Channels in Data missing from Channel_master",
        72: "Merchant types / countries / methods missing from Lists",
        73: "Country_merchant and route keys missing from Lists"
    }

    @classmethod
    def reconcile(
        cls,
        workbook_path: str,
        raw_captures: Dict[str, Any],
        to_date: date,
        rpt_month_end: date
    ) -> Dict[str, Any]:
        wb = openpyxl.load_workbook(workbook_path, data_only=True)
        results = {"checks_passed": [], "checks_failed": [], "warnings": []}

        try:
            # 1. Reconcile Control As-Of date
            ws_ctrl = wb["Control"]
            as_of_val = None
            for r in range(1, 20):
                if ws_ctrl.cell(r, 1).value and "Data as-of date (last day in the export)" in str(ws_ctrl.cell(r, 1).value):
                    as_of_val = ws_ctrl.cell(r, 2).value
                    break

            expected_as_of = datetime(to_date.year, to_date.month, to_date.day, 0, 0)
            if isinstance(as_of_val, date) and not isinstance(as_of_val, datetime):
                as_of_val = datetime(as_of_val.year, as_of_val.month, as_of_val.day, 0, 0)

            if as_of_val != expected_as_of:
                raise ReconciliationError(
                    f"Control as-of date mismatch: expected {expected_as_of}, got {as_of_val}"
                )
            results["checks_passed"].append(f"control_as_of_date_match ({to_date})")

            # 2. Reconcile BO_extract vs Data current month rows
            ws_bo = wb["BO_extract"]
            ws_data = wb["Data"]

            bo_leaf_rows = []
            for r in range(2, ws_bo.max_row + 1):
                p_val = ws_bo.cell(r, 1).value
                is_curr = False
                if isinstance(p_val, (date, datetime)):
                    is_curr = (p_val.year == rpt_month_end.year and p_val.month == rpt_month_end.month)
                elif str(p_val).startswith(rpt_month_end.strftime("%Y-%m")):
                    is_curr = True

                if is_curr:
                    prov = ws_bo.cell(r, 7).value
                    if prov and str(prov).strip():
                        row_vals = [ws_bo.cell(r, c).value for c in range(1, 23)]
                        bo_leaf_rows.append(row_vals)

            data_curr_rows = []
            for r in range(5, ws_data.max_row + 1):
                p_val = ws_data.cell(r, 1).value
                is_curr = False
                if isinstance(p_val, (date, datetime)):
                    is_curr = (p_val.year == rpt_month_end.year and p_val.month == rpt_month_end.month)
                elif str(p_val).startswith(rpt_month_end.strftime("%Y-%m")):
                    is_curr = True

                if is_curr:
                    row_vals = [ws_data.cell(r, c).value for c in range(1, 23)]
                    data_curr_rows.append(row_vals)

            if len(bo_leaf_rows) != len(data_curr_rows):
                raise ReconciliationError(
                    f"Row count mismatch between BO_extract (Filter=1) ({len(bo_leaf_rows)}) "
                    f"and Data sheet ({len(data_curr_rows)}) for {rpt_month_end}"
                )
            results["checks_passed"].append(f"data_row_count_exact_match ({len(data_curr_rows)} rows)")

            # Check exact value parity on key metrics
            for i, (bo_row, d_row) in enumerate(zip(bo_leaf_rows, data_curr_rows)):
                for metric_idx in [7, 8, 10, 12, 13, 15, 21]:
                    v_bo = float(bo_row[metric_idx] or 0)
                    v_d = float(d_row[metric_idx] or 0)
                    if abs(v_bo - v_d) > 0.001:
                        raise ReconciliationError(
                            f"Discrepancy at leaf row {i} for metric index {metric_idx}: BO={v_bo}, Data={v_d}"
                        )
            results["checks_passed"].append("bo_extract_data_values_exact_equality")

            # 3. Check Control sheet strict math gates
            for r, chk_name in cls.STRICT_CONTROL_GATES.items():
                status_val = ws_ctrl.cell(r, 3).value
                if status_val and str(status_val).strip() != "OK":
                    raise ReconciliationError(
                        f"Control strict gate failed at row {r} '{chk_name}': status = '{status_val}'"
                    )
                results["checks_passed"].append(f"control_gate_passed: {chk_name}")

            # Informational mapping checks
            for r, chk_name in cls.INFORMATIONAL_MAPPING_ROWS.items():
                status_val = ws_ctrl.cell(r, 3).value
                count_val = ws_ctrl.cell(r, 2).value
                if status_val and str(status_val).strip() != "OK":
                    results["warnings"].append(f"Mapping notice at row {r} '{chk_name}': {status_val} (count={count_val})")
                else:
                    results["checks_passed"].append(f"mapping_check_ok: {chk_name}")

            # 4. Check Settlements and Topups max ID
            if "settlements" in raw_captures and raw_captures["settlements"]:
                ws_settle = wb["BO_settlements"]
                max_fetched = max(int(r["id"]) for r in raw_captures["settlements"])
                sheet_ids = [int(ws_settle.cell(r, 1).value) for r in range(2, ws_settle.max_row + 1) if ws_settle.cell(r, 1).value is not None]
                if sheet_ids and max_fetched > max(sheet_ids):
                    raise ReconciliationError(f"Settlements max ID in workbook ({max(sheet_ids)}) < max fetched ({max_fetched})")
                results["checks_passed"].append("settlements_max_id_reconciled")

            if "topups" in raw_captures and raw_captures["topups"]:
                ws_topup = wb["BO_topups"]
                max_fetched = max(int(r["id"]) for r in raw_captures["topups"])
                sheet_ids = [int(ws_topup.cell(r, 1).value) for r in range(2, ws_topup.max_row + 1) if ws_topup.cell(r, 1).value is not None]
                if sheet_ids and max_fetched > max(sheet_ids):
                    raise ReconciliationError(f"Topups max ID in workbook ({max(sheet_ids)}) < max fetched ({max_fetched})")
                results["checks_passed"].append("topups_max_id_reconciled")

        finally:
            wb.close()

        return results
