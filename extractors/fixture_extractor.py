import os
import json
from html import escape
from datetime import date
from typing import List, Dict, Any
from extractors.base import BaseExtractor, BalanceTableCapture

class FixtureExtractor(BaseExtractor):
    """Fixture extractor for testing, shadow runs, and dry-run validation."""

    def __init__(self, fixtures_dir: str = "data/fixtures"):
        self.fixtures_dir = os.path.abspath(fixtures_dir)

    def _load_fixture(self, filename: str) -> List[Dict[str, Any]]:
        path = os.path.join(self.fixtures_dir, filename)
        if not os.path.exists(path):
            raise FileNotFoundError(f"Fixture file not found: {path}")
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)

    def extract_dashboard(self, from_date: date, to_date: date) -> List[Dict[str, Any]]:
        return self._load_fixture("dashboard.json")

    @staticmethod
    def _balance_table(rows: List[Dict[str, Any]], provider: bool) -> BalanceTableCapture:
        table_rows = []
        width = 8 if provider else 6
        for record in rows:
            if provider:
                values = [
                    record.get("col_a"), record.get("col_b"), record.get("col_c"), "",
                    record.get("on_date"), record.get("provider_balance"),
                    record.get("calculated_balance"), record.get("merchant_balance"),
                ]
            else:
                values = [
                    record.get("col_a"), record.get("col_b"), record.get("col_c"),
                    record.get("total"), record.get("pending"), record.get("available"),
                ]
            cells = ["" if value is None else str(value) for value in values]
            first_cell = cells[0].strip()
            second_cell = cells[1].strip()
            is_country = bool(first_cell) and (
                second_cell.lower().endswith(" balance")
                if not provider
                else (
                    not cells[1].strip()
                    and not cells[2].strip()
                    and len(first_cell) <= 4
                    and first_cell.isupper()
                )
            )
            table_rows.append({
                "row_class": "country-row" if is_country else "",
                "cells": cells,
            })

        while table_rows and not any(
            value.strip() for value in table_rows[-1]["cells"]
        ):
            table_rows.pop()

        html_rows = "".join(
            f'<tr class="{escape(row["row_class"])}">'
            + "".join(f"<td>{escape(value)}</td>" for value in row["cells"])
            + "</tr>"
            for row in table_rows
        )
        table_html = f"<table><tbody>{html_rows}</tbody></table>"
        return BalanceTableCapture(table_rows, table_html, len(table_rows), width)

    def extract_merchant_balance(self) -> BalanceTableCapture:
        return self._balance_table(self._load_fixture("merchant_balance.json"), provider=False)

    def extract_provider_balance(self) -> BalanceTableCapture:
        return self._balance_table(self._load_fixture("provider_balance.json"), provider=True)

    def extract_settlements(self, from_date: date, to_date: date) -> List[Dict[str, Any]]:
        records = self._load_fixture("settlements.json")
        # Filter within range if needed
        return records

    def extract_topups(self, from_date: date, to_date: date) -> List[Dict[str, Any]]:
        records = self._load_fixture("topups.json")
        return records
