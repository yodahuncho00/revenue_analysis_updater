"""Independent full-refresh pipeline for the POS Portal workbook."""

import os
import time
from datetime import date, datetime, timedelta

import openpyxl

from common.checkpoint import StageCheckpoint
from common.logger import setup_logger
from extractors.playwright_extractor import PlaywrightExtractor
from notify.report import RunReportGenerator
from scripts.update_pos_dashboard import update_pos_dashboard
from writers.calculator import ExcelCalculator
from writers.copy_manager import CopyManager


POS_HEADERS = [
    "Payment ID", "Datetime", "Type", "Amount", "Currency", "Merchant fee",
    "Provider cost", "Partner net revenue", "Account", "Receipt Name", "Agent",
    "Terminal", "Parent agent", "Provider", "Status", "Customer fee",
    "Customer fee mode", "Agent rebate",
]


def run_pos_pipeline(
    config,
    dry_run=False,
    logger=None,
    resume=False,
    override_date=None,
):
    """Refresh the POS workbook and its dashboard without depending on Gateway ETL."""
    paths = config["paths"]
    if logger is None:
        logger, run_id = setup_logger(
            log_dir=paths.get("logs_dir", "data/logs"),
            source="pos",
        )
    else:
        run_id = "shared"

    workbook_path = paths.get("pos_workbook_path", "05. Revenue_analysis_POS.xlsx")
    dashboard_path = paths.get("pos_dashboard_path", "05. Revenue_dashboard_POS.html")
    dashboard_template = paths.get(
        "pos_dashboard_template_path",
        "templates/05. Revenue_dashboard_POS.template.html",
    )
    if not os.path.isfile(workbook_path):
        raise FileNotFoundError(f"POS workbook not found: {os.path.abspath(workbook_path)}")
    if not os.path.isfile(dashboard_template):
        raise FileNotFoundError(
            f"POS dashboard template not found: {os.path.abspath(dashboard_template)}"
        )

    effective_today = (
        datetime.strptime(override_date, "%Y-%m-%d").date()
        if override_date
        else date.today()
    )
    as_of_date = effective_today - timedelta(days=1)
    date_range_info = {
        "effective_today": str(effective_today),
        "from_date": "all_time",
        "to_date": str(as_of_date),
        "rpt_month_end": str(as_of_date),
    }
    checkpoint = StageCheckpoint(
        source="pos",
        business_date=str(effective_today),
        base_dir=paths.get("checkpoints_dir"),
        scope_key=workbook_path,
    )

    manager = CopyManager(
        prod_path=workbook_path,
        backup_dir=paths.get("backup_dir", "data/backups"),
        backups_to_retain=config.get("app", {}).get("backups_to_retain", 30),
        run_id=f"{run_id}_pos",
    )
    reports_dir = paths.get("reports_dir", "data/reports")
    checks_passed = []
    checks_failed = []
    row_counts = {}
    working_path = None
    started = time.time()
    rows = None
    dashboard_rows = None
    workbook_already_committed = (
        resume and not dry_run and checkpoint.is_stage_done("workbook_commit")
    )

    try:
        logger.info("--- POS PORTAL: Full Payments CSV refresh (resume=%s) ---", resume)
        if workbook_already_committed:
            logger.info(
                "Resume mode: POS workbook commit already completed for %s; "
                "resuming post-commit publishing.",
                effective_today,
            )
            rows = checkpoint.load_pos_rows()
            checks_passed.append("atomic_commit_successful")
            checks_passed.append("resumed_after_workbook_commit")
        else:
            # Fail fast before browser launch if Excel or OneDrive persistently locks the workbook.
            if not dry_run:
                manager.verify_target_writable()
                checks_passed.append("preflight_pos_workbook_writable")

            manager.acquire_lock()
            backup_path, backup_hash = manager.create_backup()
            logger.info("POS workbook backup created: %s (%s...)", backup_path, backup_hash[:16])
            checks_passed.append(f"preflight_backup_created ({os.path.basename(backup_path)})")
            working_path = manager.create_working_copy()

            cached_rows = (
                checkpoint.load_pos_rows()
                if resume and checkpoint.is_stage_done("extract_pos_payments")
                else None
            )
            if cached_rows is not None:
                logger.info(
                    "Reusing staged POS payments from checkpoint (%d rows).",
                    len(cached_rows),
                )
                rows = cached_rows
            else:
                raw_storage_dir = os.path.join(paths.get("raw_dir", "data/raw"), f"{run_id}_pos")
                extractor = PlaywrightExtractor(
                    config, raw_storage_dir=raw_storage_dir, logger=logger
                )
                rows = extractor.extract_pos_payments()
                if not dry_run:
                    checkpoint.save_pos_rows(rows)
                    checkpoint.mark_stage("extract_pos_payments", details={"rows": len(rows)})
            checks_passed.append(f"pos_payments_extracted ({len(rows)} rows)")

            workbook = openpyxl.load_workbook(working_path, data_only=False)
            try:
                if "BO_extract" not in workbook.sheetnames:
                    raise ValueError("POS workbook has no BO_extract sheet.")
                sheet = workbook["BO_extract"]
                actual_headers = [sheet.cell(3, column).value for column in range(1, 19)]
                if actual_headers != POS_HEADERS:
                    raise ValueError(
                        "POS workbook BO_extract!A3:R3 headers changed. "
                        f"Expected {POS_HEADERS!r}; received {actual_headers!r}."
                    )

                previous_last_row = max(sheet.max_row, 3)
                for row_number in range(4, previous_last_row + 1):
                    for column in range(1, 19):
                        sheet.cell(row_number, column).value = None
                for row_number, values in enumerate(rows, start=4):
                    if len(values) != 18:
                        raise ValueError(f"POS payment row has {len(values)} columns; expected 18.")
                    for column, value in enumerate(values, start=1):
                        cell = sheet.cell(row_number, column)
                        cell.value = value
                        if column == 2 and isinstance(value, (datetime, date)):
                            # Display as dd.mm.yyyy (date only, no time component shown)
                            cell.number_format = "DD.MM.YYYY"

                # Stamp DD.MM.YYYY on every data row in col AA (formula column - value preserved).
                last_data_row = max(sheet.max_row, 3 + len(rows))
                for row_number in range(4, last_data_row + 1):
                    sheet.cell(row_number, 27).number_format = "DD.MM.YYYY"

                if "Control" not in workbook.sheetnames:
                    raise ValueError("POS workbook has no Control sheet.")
                control_cell = workbook["Control"]["B6"]
                control_cell.value = datetime.combine(as_of_date, datetime.min.time())
                control_cell.number_format = "DD.MM.YYYY"

                workbook.calculation.fullCalcOnLoad = True
                workbook.calculation.forceFullCalc = True
                workbook.calculation.calcMode = "auto"
                workbook.save(working_path)
            finally:
                workbook.close()

            row_counts["bo_extract"] = len(rows)
            logger.info("POS BO_extract refreshed with %s rows at A4:R%s.", len(rows), len(rows) + 3)
            logger.info("Updated POS Control!B6 to yesterday (%s).", as_of_date)
            ExcelCalculator.recalculate(working_path)
            checks_passed.append("excel_com_full_recalculate_success")
            logger.info("POS workbook formulas recalculated by Excel.")

            if dry_run:
                manager.cleanup_working_copy()
                checks_passed.append("dry_run_commit_bypassed")
                logger.info("POS dry run completed; workbook/dashboard were not committed.")
                RunReportGenerator.generate_report(
                    run_id=run_id,
                    success=True,
                    date_range=date_range_info,
                    row_counts=row_counts,
                    checks_passed=checks_passed,
                    checks_failed=[],
                    duration_seconds=time.time() - started,
                    output_dir=reports_dir,
                )
                return {"success": True, "rows": len(rows), "committed": False}

            manager.commit()
            manager.cleanup_working_copy()
            manager.release_lock()
            checkpoint.mark_stage("workbook_commit", details={"run_id": run_id, "rows": len(rows)})
            checks_passed.append("atomic_commit_successful")

        # Decoupled post-commit sinks
        if not (resume and checkpoint.is_stage_done("publish_dashboard")):
            dashboard_rows = update_pos_dashboard(workbook_path, dashboard_template, dashboard_path)
            checkpoint.mark_stage("publish_dashboard", details={"rows": dashboard_rows})
            checks_passed.append("html_dashboard_refreshed")
            logger.info("POS dashboard refreshed with %s transactions.", dashboard_rows)
        else:
            dashboard_rows = checkpoint.get_stage_details("publish_dashboard").get("rows", 0)
            checks_passed.append("html_dashboard_refreshed")
            logger.info("Resume mode: POS dashboard already refreshed; skipping.")

        google_settings = config.get("google_sheets", {})
        if google_settings.get("enabled", False):
            if not (resume and checkpoint.is_stage_done("publish_google_sheets")):
                if rows is None:
                    # Re-read committed BO_extract rows if resuming without staged rows
                    wb_committed = openpyxl.load_workbook(
                        workbook_path, read_only=True, data_only=True
                    )
                    try:
                        ws_committed = wb_committed["BO_extract"]
                        rows = [
                            list(r[:18])
                            for r in ws_committed.iter_rows(min_row=4, max_col=18, values_only=True)
                            if r[0] is not None
                        ]
                    finally:
                        wb_committed.close()

                from integrations.google_sheets_sync import GoogleSheetsSync

                logger.info("Syncing committed POS payment data to Google Sheets...")
                GoogleSheetsSync(config).sync_pos_payments(
                    os.path.abspath(workbook_path),
                    google_settings.get("pos_spreadsheet_name", "Backoffice_POS"),
                    POS_HEADERS,
                    rows,
                    logger=logger,
                )
                checkpoint.mark_stage("publish_google_sheets", details={"run_id": run_id})
                checks_passed.append("google_sheets_pos_sync_successful")
            else:
                checks_passed.append("google_sheets_pos_sync_successful")
                logger.info("Resume mode: POS Google Sheets already synced; skipping.")

        duration = time.time() - started
        row_counts["bo_extract"] = len(rows) if rows is not None else 0
        RunReportGenerator.generate_report(
            run_id=run_id,
            success=True,
            date_range=date_range_info,
            row_counts=row_counts,
            checks_passed=checks_passed,
            checks_failed=[],
            duration_seconds=duration,
            output_dir=reports_dir,
        )
        logger.info("POS full refresh completed in %.2f seconds.", duration)
        return {
            "success": True,
            "rows": row_counts["bo_extract"],
            "dashboard_rows": dashboard_rows,
            "committed": True,
        }
    except Exception as exc:
        manager.cleanup_on_failure()
        checks_failed.append(str(exc))
        logger.exception("POS Portal ETL failed.")
        try:
            RunReportGenerator.generate_report(
                run_id=run_id,
                success=False,
                date_range=date_range_info,
                row_counts=row_counts,
                checks_passed=checks_passed,
                checks_failed=checks_failed,
                duration_seconds=time.time() - started,
                error_message=str(exc),
                output_dir=reports_dir,
            )
        except Exception:
            pass
        raise
    finally:
        manager.release_lock()
