import os
import sys
import time
import argparse
import yaml
from datetime import datetime, date

from common.dates import DateManager, BUSINESS_TIMEZONE
from common.logger import setup_logger
from common.csv_export import export_raw_captures
from validators.structure_validator import WorkbookStructureValidator, StructureValidationError
from validators.raw_validator import RawDataValidator, RawValidationError
from transformers.bo_extract import BOExtractTransformer
from transformers.data_sync import DataSheetSyncTransformer
from transformers.append_only import AppendOnlyTransformer
from writers.copy_manager import CopyManager, CopyManagerError
from writers.excel_writer import ExcelWriter
from writers.excel_table_paster import ExcelTablePaster
from writers.calculator import ExcelCalculator
from reconcile.reconciler import PostWriteReconciler, ReconciliationError
from notify.report import RunReportGenerator
from notify.alerter import Alerter
from extractors.fixture_extractor import FixtureExtractor
from extractors.playwright_extractor import PlaywrightExtractor

ALLOWED_STEPS = (
    "dashboard", "1",
    "merchants_balance", "2",
    "settlements", "4",
    "topups", "5",
    "control", "6",
)


def _filter_to_working_days(records, working_days, entity_name):
    working_day_set = set(working_days)
    if not working_days:
        if records:
            raise ValueError(f"{entity_name} returned records although there are no working days to fetch.")
        return []

    start_date = working_days[0]
    end_date = working_days[-1]
    filtered = []
    for record in records:
        created_value = record.get("created", record.get("created_at"))
        created_date = AppendOnlyTransformer.parse_datetime(created_value).date()
        if not start_date <= created_date <= end_date:
            raise ValueError(
                f"{entity_name} returned a record outside the requested date range: "
                f"{created_value!r}"
            )
        if created_date in working_day_set:
            filtered.append(record)
    return filtered


def _earliest_missing_transaction_detail_date(worksheet, through_date: date):
    missing_dates = []
    for row in range(2, worksheet.max_row + 1):
        if worksheet.cell(row, 1).value is None:
            continue
        status = worksheet.cell(row, 5).value
        status_is_numeric_code = (
            isinstance(status, (int, float))
            or (isinstance(status, str) and status.strip().isdigit())
        )
        country_is_name = str(worksheet.cell(row, 8).value or "").strip().upper() == "ZAMBIA"
        if all(
            worksheet.cell(row, column).value is not None
            and str(worksheet.cell(row, column).value).strip()
            for column in (10, 11)
        ) and not status_is_numeric_code and not country_is_name:
            continue
        created = worksheet.cell(row, 2).value
        if created is None:
            raise ValueError(
                f"{worksheet.title} row {row} has missing transaction details "
                "and no Created date."
            )
        created_date = AppendOnlyTransformer.parse_datetime(created).date()
        if created_date <= through_date:
            missing_dates.append(created_date)
    return min(missing_dates, default=None)


def load_config(config_path: str = "config.yaml") -> dict:
    if not os.path.exists(config_path):
        raise FileNotFoundError(f"Configuration file not found: {config_path}")
    with open(config_path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)

def run_pipeline(
    config_path: str = "config.yaml",
    dry_run: bool = False,
    override_date: str = None,
    step: str = None,
    extractor_type: str = "playwright"
) -> int:
    if step is not None and step not in ALLOWED_STEPS:
        raise ValueError(
            f"Unsupported ETL step {step!r}. Allowed steps: "
            f"{', '.join(ALLOWED_STEPS)}. Provider Balance is not part of the ETL."
        )

    start_time = time.time()
    config = load_config(config_path)
    
    # 1. Setup logging
    logger, run_id = setup_logger(log_dir="data/logs")
    alerter = Alerter(logger)
    logger.info(f"Starting Revenue Analysis ETL Job (run_id={run_id}, dry_run={dry_run}, extractor={extractor_type})")

    # 2. Date calculations
    date_mgr = DateManager(
        tz=BUSINESS_TIMEZONE,
        first_of_month_behavior=config["app"].get("first_of_month_behavior", "process_prior_month")
    )
    
    eff_today = None
    if override_date:
        eff_today = datetime.strptime(override_date, "%Y-%m-%d").date()
        logger.info(f"Overriding effective today to {eff_today}")

    from_date, to_date, rpt_month_end = date_mgr.get_dashboard_date_range(eff_today)
    date_range_info = {
        "effective_today": str(date_mgr.get_effective_today(eff_today)),
        "from_date": str(from_date),
        "to_date": str(to_date),
        "rpt_month_end": str(rpt_month_end)
    }
    logger.info(f"Calculated date range: from={from_date}, to={to_date}, rpt_month_end={rpt_month_end}")

    # 3. Pre-flight checks & Copy Manager
    prod_path = config["paths"]["workbook_path"]
    copy_mgr = CopyManager(
        prod_path=prod_path,
        backup_dir=config["paths"]["backup_dir"],
        backups_to_retain=config["app"].get("backups_to_retain", 30)
    )

    row_counts = {}
    checks_passed = []
    checks_failed = []
    balance_tables = {}
    working_copy_path = None

    try:
        logger.info("Executing Pre-Flight structural checks on target workbook...")
        WorkbookStructureValidator.validate(prod_path)
        checks_passed.append("preflight_workbook_structure_valid")

        logger.info("Acquiring run lock and creating timestamped backup...")
        copy_mgr.acquire_lock()
        backup_path, backup_hash = copy_mgr.create_backup()
        logger.info(f"Backup created: {backup_path} (SHA-256: {backup_hash[:16]}...)")
        checks_passed.append(f"preflight_backup_created ({os.path.basename(backup_path)})")

        working_copy_path = copy_mgr.create_working_copy()
        logger.info(f"Working copy ready: {working_copy_path}")

        # 4. Instantiate Extractor
        raw_storage_dir = os.path.join(config["paths"]["raw_dir"], run_id)
        if extractor_type == "playwright":
            extractor = PlaywrightExtractor(config, raw_storage_dir=raw_storage_dir)
        else:
            extractor = FixtureExtractor(fixtures_dir=config["paths"]["fixtures_dir"])

        # 5. Extraction & Raw Validation
        raw_captures = {}
        downloaded_transaction_captures = {}

        # STEP 1: Dashboard -> BO_extract -> Data
        if step in (None, "dashboard", "1"):
            logger.info("--- STEP 1: Dashboard Table Extraction & Validation ---")
            raw_dashboard = extractor.extract_dashboard(from_date, to_date)
            RawDataValidator.validate_non_empty(raw_dashboard, "Dashboard")
            RawDataValidator.validate_dashboard_hierarchy(raw_dashboard, tolerance=config["tolerances"]["hierarchy_additive"])
            checks_passed.append("raw_dashboard_hierarchy_validated")

            raw_captures["dashboard"] = raw_dashboard
            bo_rows = BOExtractTransformer.transform_rows(raw_dashboard, rpt_month_end)
            data_rows = DataSheetSyncTransformer.filter_and_map_rows(bo_rows)

            RawDataValidator.validate_no_duplicate_keys([r for r in raw_dashboard if r.get("provider")])
            checks_passed.append("raw_dashboard_no_duplicate_leaf_keys")

        # STEP 2: Merchants Balance
        if step in (None, "merchants_balance", "2"):
            logger.info("--- STEP 2: Merchants Balance Extraction ---")
            raw_mb = extractor.extract_merchant_balance()
            RawDataValidator.validate_non_empty(raw_mb.rows, "Merchants Balance")
            raw_captures["merchant_balance"] = raw_mb.rows
            balance_tables["BO_merchant_balance"] = (raw_mb, "A4")
            checks_passed.append("raw_merchant_balance_validated")

        # Read existing workbook state for append-only guards
        writer = ExcelWriter(working_copy_path)

        # STEP 4: Settlements (append-only)
        if step in (None, "settlements", "4"):
            logger.info("--- STEP 4: Settlements Extraction & Duplicate Guard ---")
            ws_settle = writer.wb["BO_settlements"]
            removed_future_rows = writer.remove_future_transaction_rows(
                "BO_settlements", to_date
            )
            if removed_future_rows:
                logger.warning(
                    f"Cleared {removed_future_rows} trailing BO_settlements rows "
                    f"created after reporting cutoff {to_date}."
                )
            existing_settle_ids = {
                int(ws_settle.cell(r, 1).value)
                for r in range(2, ws_settle.max_row + 1)
                if ws_settle.cell(r, 1).value is not None
            }
            # Column B contains Created; columns L and M are derived Year/Month formulas.
            latest_dt = max(
                (ws_settle.cell(r, 2).value for r in range(2, ws_settle.max_row + 1) if ws_settle.cell(r, 2).value is not None),
                default=datetime(2024, 1, 1)
            )
            latest_date = latest_dt.date() if isinstance(latest_dt, datetime) else latest_dt
            working_days = date_mgr.get_append_working_days(latest_date, eff_today)

            detail_backfill_start = _earliest_missing_transaction_detail_date(
                ws_settle, to_date
            )
            fetch_starts = [
                start for start in (
                    working_days[0] if working_days else None,
                    detail_backfill_start,
                )
                if start is not None
            ]
            fetch_start = min(fetch_starts, default=None)
            fetched_settlements = (
                extractor.extract_settlements(fetch_start, to_date)
                if fetch_start is not None and fetch_start <= to_date else []
            )
            downloaded_transaction_captures["settlements"] = fetched_settlements
            if detail_backfill_start is not None:
                updated_details = writer.backfill_transaction_details(
                    "BO_settlements", fetched_settlements
                )
                logger.info(
                    "Backfilled Settlements CSV details: "
                    f"Country={updated_details['country']}, "
                    f"Status={updated_details['status']}, "
                    f"Merchant Order Id={updated_details['merchant_order_id']}, "
                    f"Note={updated_details['note']}"
                )
            working_day_set = set(working_days)
            new_settlements = [
                record for record in fetched_settlements
                if int(record["id"]) not in existing_settle_ids
                and AppendOnlyTransformer.parse_datetime(
                    record.get("created", record.get("created_at", ""))
                ).date() in working_day_set
            ]
            raw_settlements = _filter_to_working_days(
                new_settlements, working_days, "Settlements"
            )
            raw_captures["settlements"] = raw_settlements

            settle_rows = AppendOnlyTransformer.transform_records(raw_settlements, existing_settle_ids, "Settlements")
            checks_passed.append(f"settlements_duplicate_guard_passed ({len(raw_settlements)} fetched records)")

        # STEP 5: Top-ups (append-only)
        if step in (None, "topups", "5"):
            logger.info("--- STEP 5: Top-ups Extraction & Duplicate Guard ---")
            ws_topup = writer.wb["BO_topups"]
            removed_future_rows = writer.remove_future_transaction_rows(
                "BO_topups", to_date
            )
            if removed_future_rows:
                logger.warning(
                    f"Cleared {removed_future_rows} trailing BO_topups rows "
                    f"created after reporting cutoff {to_date}."
                )
            existing_topup_ids = {
                int(ws_topup.cell(r, 1).value)
                for r in range(2, ws_topup.max_row + 1)
                if ws_topup.cell(r, 1).value is not None
            }
            latest_dt = max(
                (ws_topup.cell(r, 2).value for r in range(2, ws_topup.max_row + 1) if ws_topup.cell(r, 2).value is not None),
                default=datetime(2024, 1, 1)
            )
            latest_date = latest_dt.date() if isinstance(latest_dt, datetime) else latest_dt
            working_days = date_mgr.get_append_working_days(latest_date, eff_today)
            detail_backfill_start = _earliest_missing_transaction_detail_date(
                ws_topup, to_date
            )
            fetch_starts = [
                start for start in (
                    working_days[0] if working_days else None,
                    detail_backfill_start,
                )
                if start is not None
            ]
            fetch_start = min(fetch_starts, default=None)
            fetched_topups = (
                extractor.extract_topups(fetch_start, to_date)
                if fetch_start is not None and fetch_start <= to_date else []
            )
            downloaded_transaction_captures["topups"] = fetched_topups
            if detail_backfill_start is not None:
                updated_details = writer.backfill_transaction_details(
                    "BO_topups", fetched_topups
                )
                logger.info(
                    "Backfilled Top-ups CSV details: "
                    f"Country={updated_details['country']}, "
                    f"Status={updated_details['status']}, "
                    f"Merchant Order Id={updated_details['merchant_order_id']}, "
                    f"Note={updated_details['note']}"
                )
            working_day_set = set(working_days)
            new_topups = [
                record for record in fetched_topups
                if int(record["id"]) not in existing_topup_ids
                and AppendOnlyTransformer.parse_datetime(
                    record.get("created", record.get("created_at", ""))
                ).date() in working_day_set
            ]
            raw_topups = _filter_to_working_days(
                new_topups, working_days, "Top-ups"
            )
            raw_captures["topups"] = raw_topups

            topup_rows = AppendOnlyTransformer.transform_records(raw_topups, existing_topup_ids, "Topups")
            checks_passed.append(f"topups_duplicate_guard_passed ({len(raw_topups)} fetched records)")

        export_captures = dict(raw_captures)
        export_captures.update(downloaded_transaction_captures)
        export_paths = export_raw_captures(
            export_captures,
            config["paths"].get("exports_dir", "data/exports"),
            run_id,
        )
        for name, path in export_paths.items():
            logger.info(f"Exported pipeline raw capture: {name} -> {path}")

        # 6. Execute Writes on Working Copy
        logger.info("Executing precision writes to working copy...")
        if step in (None, "dashboard", "1"):
            res_bo = writer.update_bo_extract(bo_rows, rpt_month_end)
            row_counts["bo_extract"] = res_bo["rows_written"]
            logger.info(f"Updated BO_extract: wrote {res_bo['rows_written']} rows (rows {res_bo['start_row']}..{res_bo['end_row']})")

            res_data = writer.update_data_sheet(data_rows, rpt_month_end)
            row_counts["data"] = res_data["rows_written"]
            logger.info(f"Updated Data sheet: wrote {res_data['rows_written']} rows (rows {res_data['start_row']}..{res_data['end_row']})")

        if step in (None, "merchants_balance", "2"):
            logger.info("Queued complete Merchant Balance table for Excel paste")

        if step in (None, "settlements", "4"):
            res_st = writer.append_settlements(settle_rows)
            row_counts["bo_settlements"] = res_st["rows_appended"]
            logger.info(f"Appended BO_settlements: added {res_st['rows_appended']} rows")

        if step in (None, "topups", "5"):
            res_tp = writer.append_topups(topup_rows)
            row_counts["bo_topups"] = res_tp["rows_appended"]
            logger.info(f"Appended BO_topups: added {res_tp['rows_appended']} rows")

        if step in (None, "control", "6"):
            res_ctrl = writer.update_control_as_of(to_date)
            logger.info(f"Updated Control sheet as-of date: {res_ctrl['cell']} = {to_date}")
            checks_passed.append("control_as_of_date_updated")

        for sheet_name, (capture, start_cell) in balance_tables.items():
            sheet_cfg = config["sheets"][sheet_name.lower()]
            end_row = writer.update_balance_summary_formulas(
                sheet_name=sheet_name,
                start_cell=start_cell,
                row_count=capture.row_count,
                summary_row=sheet_cfg["summary_row"],
                summary_columns=sheet_cfg["summary_cols"],
                filter_column=sheet_cfg["filter_col"],
                formula_start_row=sheet_cfg["summary_formula_start_row"],
                helper_columns=sheet_cfg["helper_cols"],
                helper_number_format=sheet_cfg["helper_number_format"],
            )
            logger.info(
                f"Updated {sheet_name} summary formulas through row {end_row}"
            )

        writer.save()
        writer.close()

        if balance_tables:
            logger.info("Pasting complete balance tables through Excel and unmerging pasted cells...")
            ExcelTablePaster.paste_tables(working_copy_path, balance_tables)
            for sheet_name, (capture, start_cell) in balance_tables.items():
                row_counts[sheet_name.lower()] = capture.row_count
                logger.info(
                    f"Pasted {capture.row_count} rows x {capture.column_count} columns "
                    f"into {sheet_name}!{start_cell}"
                )

        # 7. Excel COM Recalculate
        logger.info("Triggering Excel COM recalculation on Windows...")
        try:
            ExcelCalculator.recalculate(working_copy_path)
            checks_passed.append("excel_com_full_recalculate_success")
        except Exception as e:
            logger.warning(f"Excel COM recalculation skipped or failed: {e}")

        # 8. Post-Write Reconciliation
        logger.info("Executing Post-Write Reconciliation against raw captures and integrity gates...")
        recon_res = PostWriteReconciler.reconcile(working_copy_path, raw_captures, to_date, rpt_month_end)
        checks_passed.extend(recon_res.get("checks_passed", []))
        logger.info(f"All reconciliation checks passed successfully ({len(recon_res['checks_passed'])} gates)")

        # 9. Atomic Commit
        if dry_run:
            logger.info("[DRY RUN] All checks passed! Skipping atomic commit to production.")
            checks_passed.append("dry_run_commit_bypassed")
            # Clean up working copy since all checks passed
            if working_copy_path and os.path.exists(working_copy_path):
                os.remove(working_copy_path)
                logger.info(f"Deleted working copy: {working_copy_path}")
        else:
            logger.info("Executing atomic commit to replace production workbook...")
            copy_mgr.commit()
            checks_passed.append("atomic_commit_successful")
            logger.info("Production workbook updated successfully.")

        # Release lock
        copy_mgr.release_lock()

        duration = time.time() - start_time
        RunReportGenerator.generate_report(
            run_id=run_id,
            success=True,
            date_range=date_range_info,
            row_counts=row_counts,
            checks_passed=checks_passed,
            checks_failed=[],
            duration_seconds=duration,
            output_dir=config["paths"]["reports_dir"]
        )
        logger.info(f"ETL Run Completed Successfully in {duration:.2f}s")
        return 0

    except Exception as e:
        duration = time.time() - start_time
        logger.error(f"ETL Run FAILED: {e}", exc_info=True)
        checks_failed.append(str(e))
        alerter.alert_failure(run_id, "pipeline", e)

        copy_mgr.cleanup_on_failure()

        RunReportGenerator.generate_report(
            run_id=run_id,
            success=False,
            date_range=date_range_info,
            row_counts=row_counts,
            checks_passed=checks_passed,
            checks_failed=checks_failed,
            duration_seconds=duration,
            error_message=str(e),
            output_dir=config["paths"]["reports_dir"]
        )
        return 1

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Autonomous Daily Revenue Analysis ETL Job")
    parser.add_argument("--config", default="config.yaml", help="Path to config.yaml")
    parser.add_argument("--dry-run", action="store_true", help="Execute full pipeline without committing")
    parser.add_argument("--date", help="Override effective date (YYYY-MM-DD)")
    parser.add_argument(
        "--step",
        choices=ALLOWED_STEPS,
        help="Run one ETL step (Provider Balance is excluded)",
    )
    parser.add_argument(
        "--extractor",
        default="playwright",
        choices=["fixture", "playwright"],
        help="Extractor type (fixture is intended for tests)",
    )

    args = parser.parse_args()
    code = run_pipeline(
        config_path=args.config,
        dry_run=args.dry_run,
        override_date=args.date,
        step=args.step,
        extractor_type=args.extractor
    )
    sys.exit(code)
