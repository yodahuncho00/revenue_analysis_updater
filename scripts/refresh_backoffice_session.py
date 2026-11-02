import os
import sys
import argparse
from pathlib import Path

import yaml
from playwright.sync_api import sync_playwright

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT))

from extractors.playwright_extractor import (
    PlaywrightExtractionError,
    PlaywrightExtractor,
)


def main(preflight: bool = False) -> None:
    config_path = PROJECT_ROOT / "config.yaml"
    with config_path.open("r", encoding="utf-8") as config_file:
        config = yaml.safe_load(config_file)

    backoffice = config.get("backoffice", {})
    login_url = backoffice.get("login_url")
    dashboard_url = backoffice.get("routes", {}).get("dashboard")
    if not login_url or not dashboard_url:
        raise ValueError("Configure backoffice.login_url and backoffice.routes.dashboard.")

    configured_state_path = Path(
        config.get("paths", {}).get("storage_state_path", "data/storage_state.json")
    )
    state_path = (
        configured_state_path
        if configured_state_path.is_absolute()
        else PROJECT_ROOT / configured_state_path
    )
    state_path = str(state_path.resolve())
    os.makedirs(os.path.dirname(state_path), exist_ok=True)

    with sync_playwright() as playwright:
        extractor = PlaywrightExtractor(config)

        if preflight:
            context = extractor.launch_persistent_context(playwright, headless=True)
            needs_login = False
            try:
                page = context.new_page()
                page.goto(
                    dashboard_url,
                    wait_until="domcontentloaded",
                    timeout=backoffice.get("browser", {}).get("timeout_ms", 30000),
                )
                try:
                    extractor._assert_authenticated(page)
                except PlaywrightExtractionError:
                    needs_login = True
                if not needs_login:
                    ready = page.get_by_role("button", name="Reset Filters", exact=True)
                    extractor._wait_for_backoffice_ready(page, ready, "Dashboard")
                    context.storage_state(path=state_path)
                    print("Backoffice session is valid; starting the daily ETL.")
                    return
            finally:
                context.close()

            if needs_login:
                print("Backoffice login is required before today's ETL can start.")

        context = extractor.launch_persistent_context(playwright, headless=False)
        try:
            page = context.new_page()
            page.goto(
                dashboard_url,
                wait_until="domcontentloaded",
                timeout=backoffice.get("browser", {}).get("timeout_ms", 30000),
            )
            ready = page.get_by_role("button", name="Reset Filters", exact=True)
            try:
                extractor._assert_authenticated(page)
            except PlaywrightExtractionError:
                page.goto(
                    login_url,
                    wait_until="domcontentloaded",
                    timeout=backoffice.get("browser", {}).get("timeout_ms", 30000),
                )
                print(
                    "Today's ETL is paused. Sign in and complete OTP in the visible "
                    "browser; if needed, open the Backoffice Dashboard. The ETL will "
                    "continue once its controls are ready."
                )
                ready.wait_for(
                    state="visible",
                    timeout=backoffice.get("browser", {}).get(
                        "login_timeout_ms", 600000
                    ),
                )
            else:
                print("Existing Playwright authentication found; reusing the profile.")

            extractor._wait_for_backoffice_ready(page, ready, "Dashboard")
            context.storage_state(path=state_path)
            print(f"Authenticated session saved to {state_path}")
        finally:
            context.close()


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Check or refresh the Backoffice session.")
    parser.add_argument(
        "--preflight",
        action="store_true",
        help="Check authentication before the daily ETL; prompt for login if expired.",
    )
    main(preflight=parser.parse_args().preflight)
