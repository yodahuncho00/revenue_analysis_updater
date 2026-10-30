import os
import json
import time
import hashlib
import json
import re
import csv
import io
from contextlib import contextmanager
from datetime import date
from pathlib import Path
from typing import List, Dict, Any, Optional, Iterator
from urllib.parse import parse_qs, urlparse
from extractors.base import BaseExtractor, BalanceTableCapture

class PlaywrightExtractionError(Exception):
    pass

class PlaywrightExtractor(BaseExtractor):
    """Playwright-based live web extractor for Zoyktech Backoffice."""

    TRANSACTION_CSV_COLUMNS = (
        "Id",
        "Created",
        "Method",
        "Provider Id",
        "Status",
        "Amount",
        "Currency",
        "Country",
        "Merchant",
        "Merchant Order Id",
        "Note",
    )

    DASHBOARD_METRICS = (
        "deposits",
        "deposits_fee",
        "deposits_fee_pct",
        "deposits_cost",
        "deposits_cost_pct",
        "withdrawals",
        "withdrawals_fee",
        "withdrawals_fee_pct",
        "withdrawals_cost",
        "withdrawals_cost_pct",
        "settlement",
        "topup",
        "merchant_delta",
        "provider_delta",
        "partner_revenue",
    )

    def __init__(self, config: Dict[str, Any], raw_storage_dir: Optional[str] = None):
        self.config = config
        self.bo_cfg = config.get("backoffice", {})
        self.routes = self.bo_cfg.get("routes", {})
        self.storage_state_path = config.get("paths", {}).get("storage_state_path", "data/storage_state.json")
        self.browser_profile_dir = config.get("paths", {}).get(
            "browser_profile_dir",
            "%LOCALAPPDATA%\\RevenueAnalysisUpdater\\BackofficeProfile",
        )
        self.raw_storage_dir = raw_storage_dir

    def launch_persistent_context(
        self, playwright: Any, headless: Optional[bool] = None
    ) -> Any:
        profile_path = Path(
            os.path.expandvars(os.path.expanduser(self.browser_profile_dir))
        )
        if "%" in str(profile_path):
            raise PlaywrightExtractionError(
                "Could not resolve paths.browser_profile_dir; verify LOCALAPPDATA "
                "or configure an absolute browser profile path."
            )
        profile_path.mkdir(parents=True, exist_ok=True)
        marker_path = profile_path / ".storage_state_imported"
        browser_cfg = self.bo_cfg.get("browser", {})
        context = playwright.chromium.launch_persistent_context(
            str(profile_path),
            headless=browser_cfg.get("headless", True) if headless is None else headless,
            viewport=browser_cfg.get("viewport", {"width": 1280, "height": 720}),
        )
        try:
            if not marker_path.exists() and os.path.isfile(self.storage_state_path):
                with open(self.storage_state_path, "r", encoding="utf-8") as state_file:
                    state = json.load(state_file)
                cookies = state.get("cookies", [])
                if cookies:
                    context.add_cookies(cookies)
                origins = state.get("origins", [])
                if origins:
                    context.add_init_script(
                        script=(
                            "const savedOrigins = "
                            f"{json.dumps(origins)};"
                            "const savedOrigin = savedOrigins.find("
                            "entry => entry.origin === window.location.origin);"
                            "if (savedOrigin) {"
                            "for (const item of savedOrigin.localStorage) {"
                            "window.localStorage.setItem(item.name, item.value);"
                            "}"
                            "}"
                        )
                    )
                marker_path.write_text("Imported saved Playwright storage state.\n", encoding="utf-8")
        except Exception:
            context.close()
            raise
        return context

    @staticmethod
    def _assert_authenticated(page: Any) -> None:
        page_url = page.url
        path = urlparse(page_url).path
        if re.search(r"/(login|sign[-_]?in|auth|otp|verification)(/|$)", path, re.IGNORECASE):
            raise PlaywrightExtractionError(
                f"Backoffice authentication is required (current page: {page_url}). "
                "Refresh the saved session with "
                "'python scripts\\refresh_backoffice_session.py', then retry."
            )

        auth_inputs = page.locator(
            'input[type="password"], input[autocomplete="one-time-code"], '
            'input[name*="otp" i], input[name*="verification" i]'
        )
        for index in range(auth_inputs.count()):
            if auth_inputs.nth(index).is_visible():
                raise PlaywrightExtractionError(
                    f"Backoffice is showing a login or verification form "
                    f"(current page: {page_url}). Refresh the saved session with "
                    "'python scripts\\refresh_backoffice_session.py', then retry."
                )

    def _wait_for_backoffice_ready(self, page: Any, locator: Any, page_name: str) -> None:
        from playwright.sync_api import TimeoutError as PlaywrightTimeoutError

        self._assert_authenticated(page)
        try:
            locator.wait_for(
                state="visible",
                timeout=self.bo_cfg.get("browser", {}).get("timeout_ms", 30000),
            )
        except PlaywrightTimeoutError as exc:
            self._assert_authenticated(page)
            raise PlaywrightExtractionError(
                f"{page_name} did not become ready: the expected Backoffice controls "
                f"were not visible at {page.url}. Check network access and whether "
                "the Backoffice page layout has changed."
            ) from exc

    def _save_raw(self, name: str, data: Any, metadata: Dict[str, Any]) -> str:
        if not self.raw_storage_dir:
            return ""
        os.makedirs(self.raw_storage_dir, exist_ok=True)
        raw_path = os.path.join(self.raw_storage_dir, f"{name}.json")
        meta_path = os.path.join(self.raw_storage_dir, f"{name}_meta.json")

        content = json.dumps(data, indent=2, ensure_ascii=False)
        sha256 = hashlib.sha256(content.encode("utf-8")).hexdigest()

        with open(raw_path, "w", encoding="utf-8") as f:
            f.write(content)

        metadata["sha256"] = sha256
        metadata["timestamp"] = time.time()
        metadata["row_count"] = len(data) if isinstance(data, list) else 1

        with open(meta_path, "w", encoding="utf-8") as f:
            json.dump(metadata, f, indent=2, ensure_ascii=False)

        return raw_path

    @contextmanager
    def _authenticated_page(
        self, playwright: Any, url: str, response_handler: Any = None
    ) -> Iterator[Any]:
        context = self.launch_persistent_context(playwright)
        try:
            page = context.new_page()
            if response_handler is not None:
                page.on("response", response_handler)
            page.goto(
                url,
                wait_until="commit",
                timeout=self.bo_cfg.get("browser", {}).get(
                    "navigation_timeout_ms", 60000
                ),
            )
            self._assert_authenticated(page)
            yield page
        finally:
            context.close()

    @staticmethod
    def _read_table_rows(page: Any) -> List[Dict[str, Any]]:
        return page.locator("tbody tr").evaluate_all(
            "rows => rows.map(row => ({"
            "class: row.getAttribute('class') || '',"
            "cells: Array.from(row.querySelectorAll('td')).map(cell => cell.innerText.trim())"
            "}))"
        )

    @staticmethod
    def _capture_balance_table(
        page: Any, *, omit_header_rows: bool = False
    ) -> BalanceTableCapture:
        table = page.evaluate(
            """(omitHeaderRows) => {
                const table = Array.from(document.querySelectorAll("table"))
                    .find(candidate => candidate.querySelector("tbody tr.account-row"));
                if (!table) return null;

                const body = Array.from(table.tBodies)
                    .find(candidate => candidate.querySelector("tr.account-row"));
                if (!body) return null;

                const styleProperties = [
                    "border-collapse", "border-spacing", "table-layout", "width", "height",
                    "font-family", "font-size", "font-weight", "font-style", "color",
                    "background-color", "text-align", "vertical-align", "white-space",
                    "padding", "border-top", "border-right", "border-bottom", "border-left"
                ];
                const cloneStyled = source => {
                    const clone = source.cloneNode(false);
                    if (source.nodeType === Node.ELEMENT_NODE) {
                        const style = getComputedStyle(source);
                        for (const property of styleProperties) {
                            const value = style.getPropertyValue(property);
                            if (value) clone.style.setProperty(property, value);
                        }
                    }
                    for (const child of source.childNodes) {
                        if (child.nodeType === Node.ELEMENT_NODE) {
                            if (source === table && child !== body && child.tagName !== "COLGROUP") {
                                continue;
                            }
                            clone.appendChild(cloneStyled(child));
                        } else {
                            clone.appendChild(child.cloneNode(true));
                        }
                    }
                    return clone;
                };
                const isHeaderRow = row =>
                    row.classList.contains("headers-row");
                const removeScheduleIndicators = node => {
                    node.querySelectorAll("i.q-icon").forEach(icon => {
                        if (icon.textContent.trim().toLowerCase() === "schedule") {
                            icon.remove();
                        }
                    });
                };
                const copiedTable = cloneStyled(table);
                removeScheduleIndicators(copiedTable);
                copiedTable.querySelectorAll("button").forEach(button => {
                    if (button.innerText.trim().toLowerCase() === "refresh") {
                        button.remove();
                    }
                });
                if (omitHeaderRows) {
                    copiedTable.querySelectorAll("tbody tr.headers-row").forEach(row =>
                        row.remove()
                    );
                }
                const rows = Array.from(body.rows)
                    .filter(row => !omitHeaderRows || !isHeaderRow(row))
                    .map(row => ({
                    row_class: row.getAttribute("class") || "",
                    height: row.getBoundingClientRect().height,
                    cells: Array.from(row.cells).map(cell => {
                        const content = cell.cloneNode(true);
                        removeScheduleIndicators(content);
                        content.querySelectorAll("button").forEach(button => {
                            if (button.innerText.trim().toLowerCase() === "refresh") {
                                button.remove();
                            }
                        });
                        return content.innerText.trim();
                    }),
                    spans: Array.from(row.cells).map(cell => ({
                        col_span: cell.colSpan,
                        row_span: cell.rowSpan
                    }))
                }));
                while (rows.length && !rows[rows.length - 1].cells.some(
                    cell => cell.trim()
                )) {
                    rows.pop();
                }
                const copiedBody = Array.from(copiedTable.tBodies)
                    .find(candidate => candidate.querySelector("tr"));
                while (copiedBody && copiedBody.rows.length &&
                    !copiedBody.rows[copiedBody.rows.length - 1].innerText.trim()) {
                    copiedBody.deleteRow(copiedBody.rows.length - 1);
                }
                const columnCount = Math.max(0, ...rows.map(row =>
                    row.spans.reduce((total, span) => total + span.col_span, 0)
                ));
                return {
                    html: copiedTable.outerHTML,
                    rows,
                    row_count: rows.length,
                    column_count: columnCount
                };
            }""",
            omit_header_rows,
        )
        if not table or not table["rows"] or not table["column_count"]:
            raise PlaywrightExtractionError("Could not capture the complete balance table.")
        return BalanceTableCapture(
            rows=table["rows"],
            html=table["html"],
            row_count=table["row_count"],
            column_count=table["column_count"],
        )

    @staticmethod
    def _scroll_balance_table_to_end(page: Any) -> None:
        container = page.locator(".providers-balance-table")
        if container.count() != 1:
            raise PlaywrightExtractionError(
                "Provider Balance did not expose exactly one scrollable table."
            )

        stable_observations = 0
        previous_state = None
        for _ in range(10):
            container.evaluate("element => element.scrollTo(0, element.scrollHeight)")
            page.wait_for_timeout(500)
            state = container.evaluate(
                """element => ({
                    row_count: element.querySelectorAll("tbody tr").length,
                    scroll_height: element.scrollHeight,
                    scroll_top: element.scrollTop,
                    client_height: element.clientHeight
                })"""
            )
            at_bottom = state["scroll_top"] + state["client_height"] >= state["scroll_height"] - 1
            if at_bottom and state == previous_state:
                stable_observations += 1
            else:
                stable_observations = 0
            if stable_observations >= 3:
                container.evaluate("element => element.scrollTo(0, 0)")
                return
            previous_state = state

        raise PlaywrightExtractionError(
            "Provider Balance table did not finish loading after scrolling to the end."
        )

    def _save_balance_capture(
        self, name: str, capture: BalanceTableCapture, metadata: Dict[str, Any]
    ) -> None:
        self._save_raw(name, capture.rows, metadata)
        if self.raw_storage_dir:
            html_path = os.path.join(self.raw_storage_dir, f"{name}.html")
            with open(html_path, "w", encoding="utf-8") as html_file:
                html_file.write(capture.html)

    @staticmethod
    def _transaction_payload(response: Any, expected_page: int) -> Dict[str, Any]:
        if not response.ok:
            raise PlaywrightExtractionError(
                f"Transaction request failed with HTTP {response.status}."
            )
        content_type = response.headers.get("content-type", "")
        if "json" not in content_type.lower():
            raise PlaywrightExtractionError(
                "Transaction request returned an unexpected content type "
                f"({content_type or 'missing'}), not JSON."
            )
        body = response.json()
        transactions = body.get("props", {}).get("settlements")
        if not isinstance(transactions, dict) or not isinstance(transactions.get("data"), list):
            raise PlaywrightExtractionError("Transaction response has no paginated record list.")
        if transactions.get("current_page") != expected_page:
            raise PlaywrightExtractionError(
                f"Transaction response page mismatch: expected {expected_page}, "
                f"received {transactions.get('current_page')!r}."
            )
        return transactions

    @classmethod
    def _parse_transaction_csv(
        cls,
        content: bytes,
        api_records: List[Dict[str, Any]],
        capture_name: str,
    ) -> List[Dict[str, Any]]:
        try:
            text = content.decode("utf-8-sig")
        except UnicodeDecodeError as exc:
            raise PlaywrightExtractionError(
                f"{capture_name.title()} CSV export is not valid UTF-8."
            ) from exc

        reader = csv.DictReader(io.StringIO(text, newline=""))
        headers = reader.fieldnames or []
        missing_columns = [column for column in cls.TRANSACTION_CSV_COLUMNS if column not in headers]
        if missing_columns:
            raise PlaywrightExtractionError(
                f"{capture_name.title()} CSV export is missing columns: "
                f"{', '.join(missing_columns)}."
            )
        if len(headers) != len(set(headers)):
            raise PlaywrightExtractionError(
                f"{capture_name.title()} CSV export contains duplicate column names."
            )

        csv_records = list(reader)
        if any(None in record for record in csv_records):
            raise PlaywrightExtractionError(
                f"{capture_name.title()} CSV export contains rows with unexpected extra fields."
            )
        if any(
            record.get(column) is None
            for record in csv_records
            for column in cls.TRANSACTION_CSV_COLUMNS
        ):
            raise PlaywrightExtractionError(
                f"{capture_name.title()} CSV export contains an incomplete row."
            )

        api_records_by_id = {}
        for record in api_records:
            try:
                transaction_id = int(record["id"])
            except (KeyError, TypeError, ValueError) as exc:
                raise PlaywrightExtractionError(
                    f"{capture_name.title()} API result contains an invalid transaction ID."
                ) from exc
            if transaction_id in api_records_by_id:
                raise PlaywrightExtractionError(
                    f"{capture_name.title()} API result contains duplicate transaction IDs."
                )
            api_records_by_id[transaction_id] = record

        csv_ids = []
        for record in csv_records:
            try:
                csv_ids.append(int(record["Id"]))
            except (TypeError, ValueError) as exc:
                raise PlaywrightExtractionError(
                    f"{capture_name.title()} CSV export contains a non-numeric transaction ID."
                ) from exc

        if len(csv_ids) != len(set(csv_ids)):
            raise PlaywrightExtractionError(
                f"{capture_name.title()} CSV export contains duplicate transaction IDs."
            )
        csv_id_set = set(csv_ids)
        unexpected_csv_ids = csv_id_set - api_records_by_id.keys()
        if unexpected_csv_ids:
            raise PlaywrightExtractionError(
                f"{capture_name.title()} CSV export contains transaction IDs that are "
                "absent from the filtered API result."
            )
        excluded_api_ids = api_records_by_id.keys() - csv_id_set
        unexpected_exclusions = [
            transaction_id
            for transaction_id in excluded_api_ids
            if str(api_records_by_id[transaction_id].get("state")) != "3"
        ]
        if unexpected_exclusions:
            raise PlaywrightExtractionError(
                f"{capture_name.title()} CSV export omitted {len(unexpected_exclusions)} "
                "API transactions that are not in excluded state 3."
            )

        normalized = []
        for record in csv_records:
            normalized.append({
                "id": record["Id"],
                "created": record["Created"],
                "created_at": record["Created"],
                "method": record["Method"],
                "provider_id": record["Provider Id"],
                "status": record["Status"],
                "amount": record["Amount"],
                "currency": record["Currency"],
                "country": record["Country"],
                "merchant": record["Merchant"],
                "merchant_order_id": record["Merchant Order Id"],
                "note": record["Note"],
            })
        return normalized

    @classmethod
    def _parse_dashboard_rows(cls, rows: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """Convert rendered country/client/method/provider rows to pipeline records."""
        records = []
        country = currency = merchant = payment_method = None

        for row in rows:
            row_class = row["class"]
            if not any(
                marker in row_class
                for marker in ("currency-row", "client-row", "payment_method-row", "gate-row")
            ):
                continue
            cells = row["cells"]
            if len(cells) != 20:
                raise PlaywrightExtractionError(
                    f"Unexpected Dashboard row shape ({len(cells)} cells) for {row_class!r}."
                )

            if "currency-row" in row_class:
                country, currency = cells[2].strip(), cells[3].strip()
                merchant = payment_method = None
                level = "country"
            elif "client-row" in row_class:
                if not country or not currency:
                    raise PlaywrightExtractionError("Dashboard merchant row has no country/currency parent.")
                merchant = cells[2].strip()
                payment_method = None
                level = "merchant"
            elif "payment_method-row" in row_class:
                if not country or not currency or not merchant:
                    raise PlaywrightExtractionError("Dashboard payment-method row has no merchant parent.")
                payment_method = cells[3].strip()
                level = "payment_method"
            elif "gate-row" in row_class:
                if not country or not currency or not merchant or not payment_method:
                    raise PlaywrightExtractionError("Dashboard provider row has an incomplete hierarchy.")
                provider = cells[4].strip()
                if not provider:
                    raise PlaywrightExtractionError("Dashboard provider row has an empty provider name.")
                level = "provider"

            record = {
                "country": country,
                "currency": currency,
                "merchant": merchant or "",
                "payment_method": payment_method or "",
                "provider": provider if level == "provider" else "",
                "hierarchy_level": level,
            }
            for name, value in zip(cls.DASHBOARD_METRICS, cells[5:]):
                try:
                    record[name] = float(value.replace(",", "").replace(" ", ""))
                except ValueError as exc:
                    raise PlaywrightExtractionError(
                        f"Dashboard {level} row has an invalid {name} value: {value!r}."
                    ) from exc
            records.append(record)

        return records

    def extract_dashboard(self, from_date: date, to_date: date) -> List[Dict[str, Any]]:
        """
        Navigates to Dashboard, sets From/To date filter, ensures 'Show in USD' is OFF,
        intercepts JSON response or scrapes table.
        """
        from playwright.sync_api import sync_playwright

        url = self.routes.get("dashboard")
        if not url:
            raise PlaywrightExtractionError("Backoffice dashboard route is not configured.")
        from_str = f"{from_date.strftime('%Y-%m-%d')} 00:00:00"
        to_str = f"{to_date.strftime('%Y-%m-%d')} 23:59:59"

        with sync_playwright() as p:
            context = self.launch_persistent_context(p)
            page = context.new_page()
            try:
                page.goto(
                    url,
                    wait_until="commit",
                    timeout=self.bo_cfg.get("browser", {}).get(
                        "navigation_timeout_ms", 60000
                    ),
                )
                reset_filters = page.get_by_role("button", name="Reset Filters", exact=True)
                self._wait_for_backoffice_ready(page, reset_filters, "Dashboard")

                reset_filters.click()
                usd_toggle = page.get_by_role("switch", name="Show in USD", exact=True)
                if usd_toggle.get_attribute("aria-checked") == "true":
                    usd_toggle.click()
                if usd_toggle.get_attribute("aria-checked") != "false":
                    raise PlaywrightExtractionError("Could not verify that Show in USD is off.")

                date_filter = page.locator("button.u-filter").filter(has_text="Date").first
                date_filter.click()
                from_input = page.get_by_role("textbox", name="From", exact=True)
                to_input = page.get_by_role("textbox", name="To", exact=True)
                if from_input.count() != 1 or to_input.count() != 1:
                    raise PlaywrightExtractionError(
                        "Dashboard date picker did not expose exactly one From and To field."
                    )
                from_input.fill(from_str)
                to_input.fill(to_str)
                with page.expect_response(
                    lambda response: "/summary?" in response.url and from_date.isoformat() in response.url,
                    timeout=self.bo_cfg.get("browser", {}).get("timeout_ms", 30000),
                ) as response_info:
                    page.get_by_role("button", name="Apply", exact=True).click()

                response = response_info.value
                if not response.ok:
                    raise PlaywrightExtractionError(
                        f"Dashboard summary request failed with HTTP {response.status}."
                    )
                body = response.json()
                summary_rows = body.get("summary")
                filters = body.get("filters", {})
                expected_from = f"{from_date:%Y-%m-%d} 00:00:00.000"
                expected_to = f"{to_date:%Y-%m-%d} 23:59:59.999"
                if not isinstance(summary_rows, list):
                    raise PlaywrightExtractionError("Dashboard summary response has no summary row list.")
                if filters.get("date", {}).get("gte") != expected_from or filters.get("date", {}).get("lt") != expected_to:
                    raise PlaywrightExtractionError(
                        f"Dashboard response filter mismatch: expected {expected_from} through {expected_to}, "
                        f"received {filters!r}."
                    )
                if set(filters) - {"date"}:
                    raise PlaywrightExtractionError(
                        f"Unexpected active Dashboard filters were returned: {filters!r}."
                    )

                applied_filter_text = " ".join(date_filter.inner_text().split())
                if from_str not in applied_filter_text or to_str not in applied_filter_text:
                    raise PlaywrightExtractionError(
                        f"Dashboard did not confirm the requested date range: {applied_filter_text!r}."
                    )

                page.wait_for_function(
                    "(expected) => document.querySelectorAll('tbody tr.gate-row').length === expected",
                    arg=len(summary_rows),
                    timeout=self.bo_cfg.get("browser", {}).get("timeout_ms", 30000),
                )
                dom_rows = page.locator("tbody tr").evaluate_all(
                    "rows => rows.map(row => ({"
                    "class: row.getAttribute('class') || '',"
                    "cells: Array.from(row.querySelectorAll('td')).map(cell => cell.innerText.trim())"
                    "}))"
                )
                captured_data = self._parse_dashboard_rows(dom_rows)
                if sum(row["hierarchy_level"] == "provider" for row in captured_data) != len(summary_rows):
                    raise PlaywrightExtractionError(
                        "Rendered Dashboard provider row count does not match the filtered summary response."
                    )

                self._save_raw("dashboard", captured_data, {
                    "url": response.url,
                    "filters_read_back": filters,
                    "show_in_usd": False,
                    "row_count": len(captured_data),
                })
                return captured_data
            finally:
                context.close()

    def extract_merchant_balance(self) -> BalanceTableCapture:
        from playwright.sync_api import sync_playwright

        url = self.routes.get("merchant_balance")
        if not url:
            raise PlaywrightExtractionError("Merchant Balance route is not configured.")

        with sync_playwright() as playwright:
            with self._authenticated_page(playwright, url) as page:
                self._wait_for_backoffice_ready(
                    page,
                    page.get_by_role("button", name="Reset Filters", exact=True),
                    "Merchant Balance",
                )
                page.wait_for_selector("tbody tr.account-row")
                first_read = self._read_table_rows(page)
                page.wait_for_timeout(300)
                second_read = self._read_table_rows(page)
                if first_read != second_read:
                    raise PlaywrightExtractionError("Merchant Balance table changed during capture.")

                capture = self._capture_balance_table(page, omit_header_rows=True)
                self._save_balance_capture("merchant_balance", capture, {
                    "url": page.url,
                    "filters_read_back": [
                        " ".join(text.split())
                        for text in page.locator("button.u-filter").all_inner_texts()
                    ],
                    "snapshot": True,
                    "row_count": capture.row_count,
                    "column_count": capture.column_count,
                })
                return capture

    def extract_provider_balance(self) -> BalanceTableCapture:
        from playwright.sync_api import sync_playwright

        url = self.routes.get("provider_balance")
        if not url:
            raise PlaywrightExtractionError("Provider Balance route is not configured.")

        api_responses = []

        def capture_provider_response(response: Any) -> None:
            if urlparse(response.url).path.rstrip("/").endswith(
                "/api/v1/providers-balance/accounts"
            ):
                if not response.ok:
                    raise PlaywrightExtractionError(
                        f"Provider Balance data request failed with HTTP {response.status}."
                    )
                api_responses.append(response.json())

        with sync_playwright() as playwright:
            with self._authenticated_page(
                playwright, url, response_handler=capture_provider_response
            ) as page:
                self._wait_for_backoffice_ready(
                    page,
                    page.get_by_role("button", name="Reset Filters", exact=True),
                    "Provider Balance",
                )
                page.wait_for_selector("tbody tr.account-row")
                self._scroll_balance_table_to_end(page)
                if len(api_responses) != 1:
                    raise PlaywrightExtractionError(
                        "Provider Balance did not return exactly one complete account data response."
                    )
                accounts = api_responses[0].get("accounts")
                if not isinstance(accounts, list) or not accounts:
                    raise PlaywrightExtractionError(
                        "Provider Balance data response has no account list."
                    )
                expanded_rows = page.evaluate(
                    """accounts => {
                        const table = Array.from(document.querySelectorAll("table"))
                            .find(candidate => candidate.querySelector("tbody tr.account-row"));
                        const body = table && Array.from(table.tBodies)
                            .find(candidate => candidate.querySelector("tr.account-row"));
                        if (!body) return {expanded: 0, rendered_ids: []};

                        const used = new Set();
                        const mappedIds = [];
                        let expanded = 0;
                        for (const row of Array.from(body.rows)) {
                            if (!row.classList.contains("account-row")) continue;
                            const cells = Array.from(row.cells);
                            const accountName = row.querySelector(".account-name-link")
                                ?.innerText.trim();
                            const gateCell = row.querySelector("td.gate-name");
                            if (!accountName || !gateCell) continue;

                            const displayedIds = new Set(
                                (gateCell.innerText.match(/\\d+/g) || [])
                            );
                            const account = accounts.find(candidate =>
                                !used.has(candidate) && candidate.name === accountName &&
                                (candidate.provider_ids || []).some(provider =>
                                    displayedIds.has(String(provider.id))
                                )
                            );
                            if (!account) continue;
                            used.add(account);
                            const providerIds = account.provider_ids || [];
                            if (!providerIds.length) continue;

                            const replacements = providerIds.map((provider, index) => {
                                const copy = row.cloneNode(true);
                                const copyCells = Array.from(copy.cells);
                                const copiedGateCell = copy.querySelector("td.gate-name");
                                copiedGateCell.textContent = `${provider.id} ${provider.country || ""}`.trim();
                                // The Backoffice shows account-level balance and calculated
                                // values once, while liabilities are specific to each provider ID.
                                if (index > 0) {
                                    copyCells[5].textContent = "";
                                    copyCells[6].textContent = "";
                                }
                                const liability = provider.liabilities?.amount;
                                copyCells[7].textContent = liability == null
                                    ? ""
                                    : Number(liability).toLocaleString("en-US", {
                                        minimumFractionDigits: 2,
                                        maximumFractionDigits: 2
                                    });
                                copy.style.height = "";
                                for (const cell of copyCells) cell.style.height = "";
                                return copy;
                            });
                            row.replaceWith(...replacements);
                            providerIds.forEach(provider => mappedIds.push(String(provider.id)));
                            expanded += providerIds.length;
                        }

                        const allIds = accounts.flatMap(account =>
                            (account.provider_ids || []).map(provider => String(provider.id))
                        );
                        const renderedIds = Array.from(body.querySelectorAll("tr.account-row td.gate-name"))
                            .flatMap(cell => (cell.innerText.match(/^\\s*(\\d+)/) || []).slice(1));
                        return {expanded, provider_ids: mappedIds, rendered_ids: renderedIds,
                            expected_ids: allIds,
                            matched_accounts: used.size};
                    }""",
                    accounts,
                )
                if (
                    expanded_rows["matched_accounts"] == 0
                    or sorted(expanded_rows["rendered_ids"])
                    != sorted(expanded_rows["expected_ids"])
                ):
                    raise PlaywrightExtractionError(
                        "Provider Balance account IDs could not be mapped completely to the page."
                    )
                first_read = self._read_table_rows(page)
                page.wait_for_timeout(300)
                second_read = self._read_table_rows(page)
                if first_read != second_read:
                    raise PlaywrightExtractionError("Provider Balance table changed during capture.")

                capture = self._capture_balance_table(page, omit_header_rows=True)
                self._save_balance_capture("provider_balance", capture, {
                    "url": page.url,
                    "filters_read_back": [
                        " ".join(text.split())
                        for text in page.locator("button.u-filter").all_inner_texts()
                    ],
                    "snapshot": True,
                    "row_count": capture.row_count,
                    "column_count": capture.column_count,
                    "provider_id_count": len(expanded_rows["rendered_ids"]),
                    "source": "Backoffice provider balance accounts API",
                })
                return capture

    def extract_settlements(self, from_date: date, to_date: date) -> List[Dict[str, Any]]:
        return self._extract_transactions("settlements", from_date, to_date, "settlements")

    def extract_topups(self, from_date: date, to_date: date) -> List[Dict[str, Any]]:
        return self._extract_transactions("top_ups", from_date, to_date, "topups")

    def _extract_transactions(
        self,
        route_name: str,
        from_date: date,
        to_date: date,
        capture_name: str,
    ) -> List[Dict[str, Any]]:
        from playwright.sync_api import sync_playwright

        url = self.routes.get(route_name)
        if not url:
            raise PlaywrightExtractionError(f"{capture_name.title()} route is not configured.")
        if from_date > to_date:
            raise PlaywrightExtractionError(
                f"Invalid {capture_name} date range: {from_date} is after {to_date}."
            )

        route_path = urlparse(url).path
        from_str = f"{from_date:%Y-%m-%d} 00:00:00"
        to_str = f"{to_date:%Y-%m-%d} 23:59:59"
        expected_filters = {
            "date": {
                "gte": f"{from_date:%Y-%m-%d} 00:00:00.000",
                "lt": f"{to_date:%Y-%m-%d} 23:59:59.999",
            }
        }

        with sync_playwright() as playwright:
            with self._authenticated_page(playwright, url) as page:
                reset = page.get_by_role("button", name="Reset Filters", exact=True)
                self._wait_for_backoffice_ready(page, reset, capture_name.title())
                reset.click()
                date_filter = page.locator("button.u-filter").filter(has_text="Date").first
                date_filter.click()
                from_input = page.get_by_role("textbox", name="From", exact=True)
                to_input = page.get_by_role("textbox", name="To", exact=True)
                if from_input.count() != 1 or to_input.count() != 1:
                    raise PlaywrightExtractionError(
                        f"{capture_name.title()} date picker did not expose exactly one From and To field."
                    )
                from_input.fill(from_str)
                to_input.fill(to_str)

                timeout = self.bo_cfg.get("browser", {}).get(
                    "transaction_response_timeout_ms", 120000
                )
                try:
                    with page.expect_response(
                        lambda response: urlparse(response.url).path == route_path,
                        timeout=timeout,
                    ) as response_info:
                        page.get_by_role("button", name="Apply", exact=True).click()
                except Exception as exc:
                    from playwright.sync_api import TimeoutError as PlaywrightTimeoutError

                    if not isinstance(exc, PlaywrightTimeoutError):
                        raise
                    raise PlaywrightExtractionError(
                        f"{capture_name.title()} filter request returned no response "
                        f"within {timeout} ms (page: {page.url}). Retry the run; "
                        "if it repeats, refresh the Backoffice session and check connectivity."
                    ) from exc

                response = response_info.value
                first_page = self._transaction_payload(response, 1)
                response_body = response.json()
                filters = response_body.get("props", {}).get("filters", {})
                if filters != expected_filters:
                    raise PlaywrightExtractionError(
                        f"{capture_name.title()} response filter mismatch: expected "
                        f"{expected_filters!r}, received {filters!r}."
                    )
                applied_filter_text = " ".join(date_filter.inner_text().split())
                if from_str not in applied_filter_text or to_str not in applied_filter_text:
                    raise PlaywrightExtractionError(
                        f"{capture_name.title()} did not confirm the requested date range: "
                        f"{applied_filter_text!r}."
                    )

                expected_total = first_page.get("total")
                last_page = first_page.get("last_page")
                if not isinstance(expected_total, int) or not isinstance(last_page, int) or last_page < 1:
                    raise PlaywrightExtractionError(
                        f"{capture_name.title()} response has invalid pagination metadata."
                    )
                records = list(first_page["data"])

                for page_number in range(2, last_page + 1):
                    try:
                        with page.expect_response(
                            lambda next_response, page_number=page_number:
                            urlparse(next_response.url).path == route_path
                            and parse_qs(urlparse(next_response.url).query).get("page") == [str(page_number)],
                            timeout=timeout,
                        ) as next_response_info:
                            page.get_by_role("button", name="Next page", exact=True).click()
                    except Exception as exc:
                        from playwright.sync_api import TimeoutError as PlaywrightTimeoutError

                        if not isinstance(exc, PlaywrightTimeoutError):
                            raise
                        raise PlaywrightExtractionError(
                            f"{capture_name.title()} page {page_number} returned no response "
                            f"within {timeout} ms. Retry the run; if it repeats, refresh "
                            "the Backoffice session and check connectivity."
                        ) from exc
                    transactions = self._transaction_payload(next_response_info.value, page_number)
                    next_body = next_response_info.value.json()
                    if next_body.get("props", {}).get("filters", {}) != expected_filters:
                        raise PlaywrightExtractionError(
                            f"{capture_name.title()} page {page_number} returned different filters."
                        )
                    if transactions.get("total") != expected_total or transactions.get("last_page") != last_page:
                        raise PlaywrightExtractionError(
                            f"{capture_name.title()} pagination changed while the capture was in progress."
                        )
                    records.extend(transactions["data"])

                if len(records) != expected_total:
                    raise PlaywrightExtractionError(
                        f"{capture_name.title()} row count mismatch: expected {expected_total}, "
                        f"captured {len(records)}."
                    )

                export_button = page.get_by_role("button", name=re.compile("Export to csv", re.IGNORECASE))
                if export_button.count() != 1:
                    raise PlaywrightExtractionError(
                        f"{capture_name.title()} page did not expose exactly one CSV export button."
                    )
                if expected_total and export_button.is_disabled():
                    raise PlaywrightExtractionError(
                        f"{capture_name.title()} CSV export is disabled despite matching records."
                    )
                if not expected_total and export_button.is_disabled():
                    normalized = []
                else:
                    export_timeout = self.bo_cfg.get("browser", {}).get(
                        "transaction_export_timeout_ms", 120000
                    )
                    try:
                        with page.expect_download(timeout=export_timeout) as download_info:
                            export_button.click()
                    except Exception as exc:
                        from playwright.sync_api import TimeoutError as PlaywrightTimeoutError

                        if not isinstance(exc, PlaywrightTimeoutError):
                            raise
                        raise PlaywrightExtractionError(
                            f"{capture_name.title()} Export to csv button did not produce "
                            f"a download within {export_timeout} ms. The CSV export is "
                            "required; no API fallback was used."
                        ) from exc
                    download = download_info.value
                    download_path = download.path()
                    if not download_path:
                        raise PlaywrightExtractionError(
                            f"{capture_name.title()} CSV export did not produce a downloadable file."
                        )
                    csv_content = Path(download_path).read_bytes()
                    normalized = self._parse_transaction_csv(csv_content, records, capture_name)

                self._save_raw(capture_name, normalized, {
                    "url": response.url,
                    "filters_read_back": filters,
                    "pages_read": last_page,
                    "source": "filtered_csv_export",
                    "api_row_count": len(records),
                    "row_count": len(normalized),
                    "api_state_3_excluded_count": len(records) - len(normalized),
                    "show_in_usd": None,
                })
                return normalized
