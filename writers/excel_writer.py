from copy import copy
import openpyxl
from datetime import datetime, date
from typing import List, Any, Dict

class ExcelWriterError(Exception):
    pass


TRANSACTION_DATE_NUMBER_FORMAT = "dd.mm.yyyy"


class ExcelWriter:
    """Precision writer applying deterministic transformations to the working workbook copy."""

    def __init__(self, workbook_path: str):
        self.workbook_path = workbook_path
        self.wb = openpyxl.load_workbook(workbook_path, data_only=False)

    def save(self) -> None:
        self.wb.save(self.workbook_path)

    def close(self) -> None:
        self.wb.close()

    def update_balance_summary_formulas(
        self,
        sheet_name: str,
        start_cell: str,
        row_count: int,
        summary_row: int,
        summary_columns: List[str],
        filter_column: str,
        formula_start_row: int,
        helper_columns: List[str],
        helper_number_format: str,
    ) -> int:
        from openpyxl.utils.cell import (
            column_index_from_string,
            coordinate_to_tuple,
        )

        start_row, _ = coordinate_to_tuple(start_cell)
        end_row = start_row + row_count - 1
        worksheet = self.wb[sheet_name]
        for column in summary_columns:
            worksheet[f"{column}{summary_row}"] = (
                f"=SUMIFS({column}${formula_start_row}:{column}${end_row},"
                f"${filter_column}${formula_start_row}:${filter_column}${end_row},1)"
            )
        helper_start = column_index_from_string(helper_columns[0])
        helper_end = column_index_from_string(helper_columns[1])
        for row in range(5, formula_start_row):
            for column in range(helper_start, helper_end + 1):
                worksheet.cell(row, column).value = None
        last_helper_row = max(
            (
                row
                for row in range(formula_start_row, worksheet.max_row + 1)
                if any(
                    worksheet.cell(row, column).value is not None
                    for column in range(helper_start, helper_end + 1)
                )
            ),
            default=end_row,
        )
        for row in range(end_row + 1, last_helper_row + 1):
            for column in range(helper_start, helper_end + 1):
                worksheet.cell(row, column).value = None
        if formula_start_row > 5:
            for column in range(helper_start, helper_end + 1):
                worksheet.cell(formula_start_row - 1, column).fill = copy(
                    worksheet.cell(formula_start_row, column).fill
                )
        for row in range(max(5, formula_start_row - 1), end_row + 1):
            for column in range(helper_start, helper_end + 1):
                worksheet.cell(row, column).number_format = helper_number_format
        return end_row

    @staticmethod
    def _bo_extract_signature(values: List[Any]) -> tuple:
        return tuple("" if value is None else value for value in values[3:22])

    # -------------------------------------------------------------
    # Step 1: BO_extract & Data Sheet
    # -------------------------------------------------------------
    def update_bo_extract(self, new_rows: List[List[Any]], rpt_month_end: date) -> Dict[str, Any]:
        ws = self.wb["BO_extract"]
        target_dt = datetime(rpt_month_end.year, rpt_month_end.month, rpt_month_end.day, 0, 0)

        # Identify existing rows matching target_dt
        match_rows = []
        for r in range(2, ws.max_row + 1):
            val = ws.cell(r, 1).value
            if isinstance(val, (datetime, date)):
                if val.year == target_dt.year and val.month == target_dt.month:
                    match_rows.append(r)
            elif str(val).startswith(target_dt.strftime("%Y-%m")):
                match_rows.append(r)

        if match_rows:
            start_row = min(match_rows)
            end_row = max(match_rows)
        else:
            start_row = ws.max_row + 1
            end_row = start_row - 1

        # Delete existing current-month range
        # To avoid index shifting issues, if replacing with exact or new size:
        # We can clear the rows and rewrite, or delete and insert.
        # But clearing and adjusting rows is cleaner for openpyxl:
        num_existing = len(match_rows)
        num_new = len(new_rows)
        orphan_rows = []
        if match_rows:
            new_signatures = {self._bo_extract_signature(row) for row in new_rows}
            scan_row = max(match_rows) + 1
            while scan_row <= ws.max_row:
                row_values = [ws.cell(scan_row, col).value for col in range(1, 23)]
                if row_values[0] is not None or row_values[1] is not None:
                    break
                if not any(value is not None for value in row_values[2:]):
                    break
                signature = self._bo_extract_signature(row_values)
                if not any(value != "" for value in signature) or signature not in new_signatures:
                    raise ExcelWriterError(
                        f"BO_extract row {scan_row} has data but no Period or Type and does not "
                        "match the replacement capture; refusing to discard it."
                    )
                orphan_rows.append(scan_row)
                scan_row += 1

        style_width = max((len(row) for row in new_rows), default=0)
        template_styles = []
        if style_width:
            candidate_rows = (
                match_rows
                + list(range(start_row - 1, 1, -1))
                + list(range(start_row + num_existing, ws.max_row + 1))
                + list(range(2, ws.max_row + 1))
            )
            checked_rows = set()
            template_row = None
            for candidate_row in candidate_rows:
                if candidate_row in checked_rows:
                    continue
                checked_rows.add(candidate_row)
                if (
                    ws.cell(candidate_row, 1).number_format != "General"
                    and ws.cell(candidate_row, 8).number_format != "General"
                ):
                    template_row = candidate_row
                    break

            if template_row is None:
                raise ExcelWriterError(
                    "Could not find a BO_extract row with both date and numeric cell formatting "
                    "to use as the replacement style template."
                )
            template_styles = [
                copy(ws.cell(template_row, col)._style)
                for col in range(1, style_width + 1)
            ]

        if num_existing > 0:
            ws.delete_rows(start_row, num_existing)

        # Insert rows if needed or write directly
        ws.insert_rows(start_row, num_new)

        for offset, row_vals in enumerate(new_rows):
            curr_r = start_row + offset
            for col_idx, style in enumerate(template_styles, 1):
                ws.cell(curr_r, col_idx)._style = copy(style)
            for col_idx, val in enumerate(row_vals, 1):
                if isinstance(val, str) and "{row}" in val:
                    val = val.format(row=curr_r)
                ws.cell(curr_r, col_idx, val)

        for old_row in orphan_rows:
            new_row = old_row - num_existing + num_new
            for col_idx in range(1, ws.max_column + 1):
                ws.cell(new_row, col_idx).value = None

        return {
            "start_row": start_row,
            "end_row": start_row + num_new - 1,
            "rows_written": num_new,
            "previous_rows": num_existing,
            "orphan_rows_cleared": len(orphan_rows)
        }

    def update_data_sheet(self, new_data_rows: List[List[Any]], rpt_month_end: date) -> Dict[str, Any]:
        ws = self.wb["Data"]
        target_dt = datetime(rpt_month_end.year, rpt_month_end.month, rpt_month_end.day, 0, 0)

        match_rows = []
        for r in range(5, ws.max_row + 1):
            val = ws.cell(r, 1).value
            if isinstance(val, (datetime, date)):
                if val.year == target_dt.year and val.month == target_dt.month:
                    match_rows.append(r)
            elif str(val).startswith(target_dt.strftime("%Y-%m")):
                match_rows.append(r)

        if match_rows:
            start_row = min(match_rows)
            old_count = len(match_rows)
        else:
            # Find first empty row in col A starting at row 5
            start_row = 5
            while ws.cell(start_row, 1).value is not None:
                start_row += 1
            old_count = 0

        num_new = len(new_data_rows)

        # Overwrite rows starting at start_row
        for offset, row_vals in enumerate(new_data_rows):
            curr_r = start_row + offset
            for col_idx, val in enumerate(row_vals, 1):
                if isinstance(val, str) and "{row}" in val:
                    val = val.format(row=curr_r)
                ws.cell(curr_r, col_idx, val)

        # If old_count > num_new, clear vacated rows in cols A:V and J, L, O, Q
        if old_count > num_new:
            for r in range(start_row + num_new, start_row + old_count):
                for c in range(1, 23):
                    ws.cell(r, c, None)

        # Ensure formula columns X:BI cover up to max data row
        final_data_row = start_row + num_new - 1
        if final_data_row > 2000:
            # Propagate formulas from row 2000 downwards
            for r in range(2001, final_data_row + 1):
                for c in range(24, 62):
                    src_f = ws.cell(2000, c).value
                    if isinstance(src_f, str) and src_f.startswith("="):
                        ws.cell(r, c, src_f.replace("2000", str(r)))

        return {
            "start_row": start_row,
            "end_row": final_data_row,
            "rows_written": num_new,
            "previous_rows": old_count
        }

    # -------------------------------------------------------------
    # Step 4: BO_settlements (Append-only)
    # -------------------------------------------------------------
    def append_settlements(self, new_rows: List[List[Any]]) -> Dict[str, Any]:
        ws = self.wb["BO_settlements"]
        last_data_row = max(
            (row for row in range(2, ws.max_row + 1) if ws.cell(row, 1).value is not None),
            default=1,
        )
        start_row = last_data_row + 1

        for offset, row_vals in enumerate(new_rows):
            curr_r = start_row + offset
            for col_idx, val in enumerate(row_vals, 1):
                if isinstance(val, str) and "{row}" in val:
                    val = val.format(row=curr_r)
                ws.cell(curr_r, col_idx, val)

        self._format_transaction_dates(ws)
        end_row = start_row + len(new_rows) - 1 if new_rows else start_row - 1
        return {
            "start_row": start_row,
            "end_row": end_row,
            "rows_appended": len(new_rows)
        }

    # -------------------------------------------------------------
    # Step 5: BO_topups (Append-only)
    # -------------------------------------------------------------
    def append_topups(self, new_rows: List[List[Any]]) -> Dict[str, Any]:
        ws = self.wb["BO_topups"]

        # Clean trailing empty rows that only have formulas
        last_data_row = 1
        for r in range(ws.max_row, 1, -1):
            if ws.cell(r, 1).value is not None:
                last_data_row = r
                break

        # Clear leftover formula rows past last_data_row
        for r in range(last_data_row + 1, ws.max_row + 1):
            for c in range(1, ws.max_column + 1):
                ws.cell(r, c, None)

        start_row = last_data_row + 1
        for offset, row_vals in enumerate(new_rows):
            curr_r = start_row + offset
            for col_idx, val in enumerate(row_vals, 1):
                if isinstance(val, str) and "{row}" in val:
                    val = val.format(row=curr_r)
                ws.cell(curr_r, col_idx, val)

        self._format_transaction_dates(ws)
        end_row = start_row + len(new_rows) - 1 if new_rows else last_data_row
        return {
            "start_row": start_row,
            "end_row": end_row,
            "rows_appended": len(new_rows)
        }

    def backfill_transaction_details(
        self, sheet_name: str, records: List[Dict[str, Any]]
    ) -> Dict[str, int]:
        if sheet_name not in {"BO_settlements", "BO_topups"}:
            raise ExcelWriterError(
                f"Transaction detail backfill is not supported for {sheet_name!r}."
            )

        worksheet = self.wb[sheet_name]
        rows_by_id = {}
        for row in range(2, worksheet.max_row + 1):
            value = worksheet.cell(row, 1).value
            if value is None:
                continue
            transaction_id = int(value)
            if transaction_id in rows_by_id:
                raise ExcelWriterError(
                    f"Duplicate transaction ID {transaction_id} in {sheet_name}."
                )
            rows_by_id[transaction_id] = row

        updated = {"country": 0, "status": 0, "merchant_order_id": 0, "note": 0}
        for record in records:
            transaction_id = int(record["id"])
            row = rows_by_id.get(transaction_id)
            if row is None:
                continue
            # The CSV export is authoritative for the displayed status. Older
            # workbook rows may contain the API state code (for example, "2")
            # instead of the human-readable CSV status (for example, "Success").
            for column, field in (
                (8, "country"),
                (5, "status"),
                (10, "merchant_order_id"),
                (11, "note"),
            ):
                current = worksheet.cell(row, column)
                source_value = str(record.get(field) or "").strip()
                if field == "country" and source_value.upper() == "ZAMBIA":
                    source_value = "ZMB"
                is_missing = current.value is None or not str(current.value).strip()
                is_zambia_name = field == "country" and str(
                    current.value or ""
                ).strip().upper() == "ZAMBIA"
                is_numeric_status = field == "status" and (
                    isinstance(current.value, (int, float))
                    or (isinstance(current.value, str) and current.value.strip().isdigit())
                )
                should_update = is_numeric_status or (
                    field != "status" and is_missing
                ) or is_zambia_name
                if should_update and source_value:
                    current.value = source_value
                    updated[field] += 1
        return updated

    def remove_future_transaction_rows(
        self, sheet_name: str, through_date: date
    ) -> int:
        """Clear trailing transactions later than the current reporting cutoff."""
        if sheet_name not in {"BO_settlements", "BO_topups"}:
            raise ExcelWriterError(
                f"Future transaction cleanup is not supported for {sheet_name!r}."
            )

        worksheet = self.wb[sheet_name]
        populated_rows = [
            row for row in range(2, worksheet.max_row + 1)
            if worksheet.cell(row, 1).value is not None
        ]
        future_rows = []
        future_seen = False
        for row in populated_rows:
            created = worksheet.cell(row, 2).value
            if isinstance(created, datetime):
                created_date = created.date()
            elif isinstance(created, date):
                created_date = created
            else:
                raise ExcelWriterError(
                    f"{sheet_name} row {row} has an invalid Created value: {created!r}."
                )

            if created_date > through_date:
                future_seen = True
                future_rows.append(row)
            else:
                if future_seen:
                    raise ExcelWriterError(
                        f"{sheet_name} contains an in-range transaction at row {row} "
                        "after a future-dated row; refusing to remove a non-trailing row."
                    )

        if not future_rows:
            return 0

        # Future rows must be a suffix. The rows are cleared rather than deleted
        # so workbook formatting and neighboring formula references stay stable.
        for row in future_rows:
            for column in range(1, max(worksheet.max_column, 13) + 1):
                worksheet.cell(row, column).value = None
        return len(future_rows)

    @staticmethod
    def _format_transaction_dates(worksheet: Any) -> None:
        for row in range(2, worksheet.max_row + 1):
            cell = worksheet.cell(row, 2)
            if isinstance(cell.value, (datetime, date)):
                cell.number_format = TRANSACTION_DATE_NUMBER_FORMAT

    # -------------------------------------------------------------
    # Step 6: Control sheet as-of date
    # -------------------------------------------------------------
    def update_control_as_of(self, to_date: date) -> Dict[str, Any]:
        ws = self.wb["Control"]
        as_of_dt = datetime(to_date.year, to_date.month, to_date.day, 0, 0)
        
        target_cell = None
        for r in range(1, 25):
            val = ws.cell(r, 1).value
            if val and "Data as-of date (last day in the export)" in str(val):
                ws.cell(r, 2, as_of_dt)
                target_cell = f"B{r}"
                break

        if not target_cell:
            raise ExcelWriterError("Could not locate 'Data as-of date (last day in the export)' on Control sheet")

        return {
            "cell": target_cell,
            "value": as_of_dt
        }
