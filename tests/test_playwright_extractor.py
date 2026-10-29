import json

import pytest

from extractors.playwright_extractor import PlaywrightExtractionError, PlaywrightExtractor


def test_persistent_context_reuses_profile_and_imports_saved_state_once(tmp_path):
    profile_dir = tmp_path / "profile"
    state_path = tmp_path / "storage_state.json"
    state_path.write_text(
        json.dumps(
            {
                "cookies": [{"name": "trusted", "value": "test-cookie"}],
                "origins": [
                    {
                        "origin": "https://backoffice.example",
                        "localStorage": [{"name": "device", "value": "test-value"}],
                    }
                ],
            }
        ),
        encoding="utf-8",
    )

    class Context:
        def __init__(self):
            self.cookies = []
            self.scripts = []
            self.closed = False

        def add_cookies(self, cookies):
            self.cookies.extend(cookies)

        def add_init_script(self, *, script):
            self.scripts.append(script)

        def close(self):
            self.closed = True

    class Chromium:
        def __init__(self):
            self.calls = []
            self.contexts = []

        def launch_persistent_context(self, user_data_dir, **options):
            self.calls.append((user_data_dir, options))
            context = Context()
            self.contexts.append(context)
            return context

    chromium = Chromium()
    extractor = PlaywrightExtractor(
        {
            "paths": {
                "storage_state_path": str(state_path),
                "browser_profile_dir": str(profile_dir),
            },
            "backoffice": {
                "browser": {
                    "headless": True,
                    "viewport": {"width": 1920, "height": 1080},
                }
            },
        }
    )

    first_context = extractor.launch_persistent_context(
        type("Playwright", (), {"chromium": chromium})()
    )
    first_context.close()
    second_context = extractor.launch_persistent_context(
        type("Playwright", (), {"chromium": chromium})()
    )
    second_context.close()

    assert chromium.calls[0][0] == chromium.calls[1][0] == str(profile_dir)
    assert chromium.calls[0][1] == {
        "headless": True,
        "viewport": {"width": 1920, "height": 1080},
    }
    assert chromium.contexts[0].cookies == [
        {"name": "trusted", "value": "test-cookie"}
    ]
    assert "https://backoffice.example" in chromium.contexts[0].scripts[0]
    assert chromium.contexts[1].cookies == []
    assert chromium.contexts[1].scripts == []


def _dashboard_row(row_class, *, country="", currency="", merchant="", method="", provider=""):
    cells = ["0"] * 20
    cells[2] = country
    cells[3] = method or currency
    cells[4] = provider
    if "client-row" in row_class:
        cells[2] = merchant
    cells[5] = "1,234.50"
    return {"class": row_class, "cells": cells}


def test_parse_dashboard_rows_preserves_hierarchy_and_numeric_metrics():
    rows = [
        {"class": "headers-row", "cells": []},
        _dashboard_row("currency-row", country="Kazakhstan", currency="KZT"),
        _dashboard_row("client-row", merchant="Merchant A"),
        _dashboard_row("payment_method-row", method="Card"),
        _dashboard_row("gate-row", provider="Provider A"),
    ]

    parsed = PlaywrightExtractor._parse_dashboard_rows(rows)

    assert [row["hierarchy_level"] for row in parsed] == [
        "country",
        "merchant",
        "payment_method",
        "provider",
    ]
    assert all(row["deposits"] == 1234.5 for row in parsed)
    assert parsed[-1]["country"] == "Kazakhstan"
    assert parsed[-1]["currency"] == "KZT"
    assert parsed[-1]["merchant"] == "Merchant A"
    assert parsed[-1]["payment_method"] == "Card"
    assert parsed[-1]["provider"] == "Provider A"


def test_parse_dashboard_rows_rejects_unexpected_metric_values():
    row = _dashboard_row("currency-row", country="Kazakhstan", currency="KZT")
    row["cells"][5] = "not-a-number"

    with pytest.raises(PlaywrightExtractionError, match="invalid deposits value"):
        PlaywrightExtractor._parse_dashboard_rows([row])

def test_authentication_redirect_fails_immediately_with_refresh_instructions():
    class Page:
        url = "https://backoffice.example/login"

    with pytest.raises(PlaywrightExtractionError, match="refresh_backoffice_session.py"):
        PlaywrightExtractor._assert_authenticated(Page())


def test_visible_otp_form_fails_with_refresh_instructions():
    from playwright.sync_api import sync_playwright

    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(headless=True)
        try:
            page = browser.new_page()
            page.set_content(
                '<form><input autocomplete="one-time-code" '
                'aria-label="Verification code"></form>'
            )
            with pytest.raises(
                PlaywrightExtractionError, match="login or verification form"
            ):
                PlaywrightExtractor._assert_authenticated(page)
        finally:
            browser.close()

def test_missing_dashboard_controls_reports_readiness_error():
    from playwright.sync_api import sync_playwright

    extractor = PlaywrightExtractor(
        {"backoffice": {"browser": {"timeout_ms": 50}}}
    )
    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(headless=True)
        try:
            page = browser.new_page()
            page.set_content("<main>Unexpected page</main>")
            with pytest.raises(
                PlaywrightExtractionError, match="Dashboard did not become ready"
            ):
                extractor._wait_for_backoffice_ready(
                    page,
                    page.get_by_role("button", name="Reset Filters", exact=True),
                    "Dashboard",
                )
        finally:
            browser.close()


def test_capture_balance_table_keeps_blank_rows_cells_and_formatting():
    from playwright.sync_api import sync_playwright

    html = """
    <table style="border-collapse: collapse; font-size: 12px">
      <tbody>
        <tr class="country-row"><td style="color: blue">Country</td><td>Currency</td></tr>
        <tr class="client-row"><td>Merchant</td><td></td></tr>
        <tr class="account-row"><td></td><td rowspan="2">Method</td></tr>
        <tr class="spacer-row"><td></td></tr>
        <tr class="account-row"><td>Account</td><td>12.5</td></tr>
        <tr class="account-row"><td colspan="2"></td></tr>
      </tbody>
    </table>
    """
    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(headless=True)
        try:
            page = browser.new_page()
            page.set_content(html)
            capture = PlaywrightExtractor._capture_balance_table(page)
        finally:
            browser.close()

    assert capture.row_count == 5
    assert capture.column_count == 2
    assert capture.rows[1]["cells"] == ["Merchant", ""]
    assert capture.rows[3]["row_class"] == "spacer-row"
    assert capture.rows[3]["cells"] == [""]
    assert capture.rows[2]["spans"][1]["row_span"] == 2
    assert capture.rows[0]["height"] > 0
    assert capture.rows[2]["height"] > 0
    assert 'color: blue' in capture.html or 'color: rgb(0, 0, 255)' in capture.html
    assert "rowspan=\"2\"" in capture.html
    assert capture.html.count("<tr") == 5


def test_balance_capture_omits_header_and_refresh_button_but_keeps_balance_button():
    from playwright.sync_api import sync_playwright

    html = """
    <table><tbody>
      <tr class="headers-row"><td>Total</td><td>Hold</td><td>Available</td></tr>
      <tr class="account-row">
        <td><button type="button">refresh</button></td>
        <td><span class="balance-with-icon">1,234.56<i class="q-icon">schedule</i></span></td>
        <td><button type="button">123,456.78</button></td>
      </tr>
    </tbody></table>
    """
    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(headless=True)
        try:
            page = browser.new_page()
            page.set_content(html)
            capture = PlaywrightExtractor._capture_balance_table(
                page, omit_header_rows=True
            )
        finally:
            browser.close()

    assert capture.row_count == 1
    assert capture.rows[0]["cells"] == ["", "1,234.56", "123,456.78"]
    assert "Total" not in capture.html
    assert "refresh" not in capture.html
    assert "schedule" not in capture.html
    assert "123,456.78" in capture.html


def test_provider_table_scroll_waits_for_lazy_rows():
    from playwright.sync_api import sync_playwright

    html = """
    <div class="providers-balance-table" style="height: 40px; overflow-y: auto">
      <table><tbody>
        <tr class="account-row"><td>1</td></tr>
        <tr class="account-row"><td>2</td></tr>
        <tr class="account-row"><td>3</td></tr>
        <tr class="account-row"><td>4</td></tr>
      </tbody></table>
    </div>
    <script>
      const table = document.querySelector(".providers-balance-table");
      table.addEventListener("scroll", () => {
        if (!table.dataset.loaded && table.scrollTop > 0) {
          table.dataset.loaded = "true";
          table.querySelector("tbody").insertAdjacentHTML(
            "beforeend",
            "<tr class='account-row'><td>5</td></tr>" +
            "<tr class='account-row'><td>6</td></tr>"
          );
        }
      });
    </script>
    """
    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(headless=True)
        try:
            page = browser.new_page()
            page.set_content(html)
            PlaywrightExtractor._scroll_balance_table_to_end(page)
            assert page.locator("tbody tr").count() == 6
        finally:
            browser.close()
