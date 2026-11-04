import os
import sys
import argparse
import time
from pathlib import Path
from urllib.parse import urlparse

import yaml
from playwright.sync_api import sync_playwright

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT))

from extractors.playwright_extractor import (
    PlaywrightExtractionError,
    PlaywrightExtractor,
)


def wait_for_dashboard(extractor, context, dashboard_url: str, timeout_ms: int):
    """Wait until an authenticated tab reaches the configured Dashboard route."""
    dashboard_path = urlparse(dashboard_url).path.rstrip("/").lower()
    deadline = time.monotonic() + timeout_ms / 1000
    last_url = "(no open page)"
    while time.monotonic() < deadline:
        for page in context.pages:
            if page.is_closed():
                continue
            last_url = page.url
            try:
                current_path = urlparse(page.url).path.rstrip("/").lower()
                if current_path != dashboard_path:
                    continue
                extractor._assert_authenticated(page)
                return page
            except PlaywrightExtractionError:
                continue
        time.sleep(0.25)
    raise PlaywrightExtractionError(
        "Backoffice login was not confirmed: no open tab reached the authenticated "
        "Dashboard route within "
        f"{timeout_ms // 1000} seconds. "
        f"Last observed page: {last_url}. Keep the browser open until the Dashboard "
        "is verified; see data/logs/scheduler.log for details."
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
            print("Checking Backoffice in Chromium before starting today's ETL.", flush=True)
        context = extractor.launch_persistent_context(playwright, headless=False)
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
                page.goto(
                    login_url,
                    wait_until="domcontentloaded",
                    timeout=backoffice.get("browser", {}).get("timeout_ms", 30000),
                )
                print(
                    "Today's ETL is paused. Sign in and complete OTP in the visible "
                    "browser; if needed, open the Backoffice Dashboard. The ETL will "
                    "continue once the authenticated Dashboard opens.",
                    flush=True,
                )
            else:
                print(
                    "Existing Playwright authentication found; reusing the profile.",
                    flush=True,
                )

            page = wait_for_dashboard(
                extractor,
                context,
                dashboard_url,
                backoffice.get("browser", {}).get("login_timeout_ms", 600000),
            )
            context.storage_state(path=state_path)
            print(
                f"Authenticated session saved to {state_path}; continuing to ETL.",
                flush=True,
            )
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
