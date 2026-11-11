"""Independent full-refresh pipeline for the POS Portal workbook."""

import os
import time
from datetime import date, datetime, timedelta

import openpyxl

from common.logger import setup_logger
from extractors.playwright_extractor import PlaywrightExtractor
from writers.calculator import ExcelCalculator
from writers.copy_manager import CopyManager
from scripts.update_pos_dashboard import update_pos_dashboard


POS_HEADERS = [
    "Payment ID", "Datetime", "Type", "Amount", "Currency", "Merchant fee",
    "Provider cost", "Partner net revenue", "Account", "Receipt Name", "Agent",
    "Terminal", "Parent agent", "Provider", "Status", "Customer fee",
    "Customer fee mode", "Agent rebate",
]


def run_pos_pipeline(config, dry_run=False, logger=None):
    """Refresh the POS workbook and its dashboard without depending on Gateway ETL."""
    if logger is None:
        logger, run_id = setup_logger(log_dir=config["paths"].get("logs_dir", "data/logs"))
    else:
        run_id = "shared"

    paths = config["paths"]
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

    manager = CopyManager(
        prod_path=workbook_path,
        backup_dir=paths.get("backup_dir", "data/backups"),
        backups_to_retain=config.get("app", {}).get("backups_to_retain", 30),
        run_id=f"{run_id}_pos",
    )
    working_path = None
    started = time.time()
    try:
        logger.info("--- POS PORTAL: Full Payments CSV refresh ---")
        manager.acquire_lock()
        backup_path, backup_hash = manager.create_backup()
        logger.info("POS workbook backup created: %s (%s...)", backup_path, backup_hash[:16])
        working_path = manager.create_working_copy()

        raw_storage_dir = os.path.join(paths.get("raw_dir", "data/raw"), f"{run_id}_pos")
        extractor = PlaywrightExtractor(config, raw_storage_dir=raw_storage_dir)
        rows = extractor.extract_pos_payments()
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
                        cell.number_format = 'DD.MM.YYYY'

            # Stamp DD.MM.YYYY on every data row in col AA (formula column – value preserved).
            last_data_row = max(sheet.max_row, 3 + len(rows))
            for row_number in range(4, last_data_row + 1):
                sheet.cell(row_number, 27).number_format = 'DD.MM.YYYY'

            if "Control" not in workbook.sheetnames:
                raise ValueError("POS workbook has no Control sheet.")
            control_cell = workbook["Control"]["B6"]
            control_cell.value = datetime.combine(date.today() - timedelta(days=1), datetime.min.time())
            control_cell.number_format = "DD.MM.YYYY"

            workbook.calculation.fullCalcOnLoad = True
            workbook.calculation.forceFullCalc = True
            workbook.calculation.calcMode = "auto"
            workbook.save(working_path)
        finally:
            workbook.close()

        logger.info("POS BO_extract refreshed with %s rows at A4:R%s.", len(rows), len(rows) + 3)
        logger.info("Updated POS Control!B6 to yesterday (%s).", control_cell.value.date())
        ExcelCalculator.recalculate(working_path)
        logger.info("POS workbook formulas recalculated by Excel.")

        if dry_run:
            manager.cleanup_working_copy()
            logger.info("POS dry run completed; workbook/dashboard were not committed.")
            return {"success": True, "rows": len(rows), "committed": False}

        manager.commit()
        manager.cleanup_working_copy()
        dashboard_rows = update_pos_dashboard(workbook_path, dashboard_template, dashboard_path)
        logger.info("POS dashboard refreshed with %s transactions.", dashboard_rows)
        logger.info("POS full refresh completed in %.2f seconds.", time.time() - started)
        return {"success": True, "rows": len(rows), "dashboard_rows": dashboard_rows, "committed": True}
    except Exception:
        manager.cleanup_on_failure()
        logger.exception("POS Portal ETL failed.")
        raise
    finally:
        manager.release_lock()
