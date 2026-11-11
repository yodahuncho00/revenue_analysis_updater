"""Authenticate Backoffice and POS in one visible browser session."""

import argparse
import os
import re
import subprocess
import sys
import tempfile
import time
from pathlib import Path
from urllib.parse import urlparse

import yaml
from playwright.sync_api import sync_playwright

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT))

from extractors.playwright_extractor import PlaywrightExtractionError, PlaywrightExtractor


def _backoffice_ready(extractor, page, dashboard_url):
    try:
        if page.is_closed():
            return False
        extractor._assert_authenticated(page)
        target = urlparse(dashboard_url).path.rstrip("/").lower()
        if urlparse(page.url).path.rstrip("/").lower() != target:
            page.goto(dashboard_url, wait_until="domcontentloaded", timeout=60000)
            extractor._assert_authenticated(page)
        return True
    except Exception:
        return False


def _pos_ready(extractor, page, pos_url="https://pos-portal.unipesa.tech/payments"):
    try:
        if page.is_closed():
            return False
        if urlparse(page.url).hostname != urlparse(pos_url).hostname:
            return False
        curr_path = urlparse(page.url).path.rstrip("/").lower()
        target_path = urlparse(pos_url).path.rstrip("/").lower()
        if curr_path != target_path:
            if not re.search(r"/(login|auth|sign[-_]?in)", curr_path):
                page.goto(pos_url, wait_until="domcontentloaded", timeout=60000)
        save_csv = page.get_by_role("button", name=re.compile(r"Save\s*CSV", re.I)).first
        if save_csv.is_visible():
            return True
        if page.locator("button, [role='button'], a").filter(has_text=re.compile(r"Save\s*CSV", re.I)).first.is_visible():
            return True
        return False
    except Exception:
        return False


def main(preflight: bool = False) -> int:
    config_path = PROJECT_ROOT / "config.yaml"
    config = yaml.safe_load(config_path.read_text(encoding="utf-8"))
    backoffice = config.get("backoffice", {})
    dashboard_url = backoffice.get("routes", {}).get("dashboard")
    pos_url = config.get("pos_portal", {}).get(
        "payments_url", "https://pos-portal.unipesa.tech/payments"
    )
    if not dashboard_url or not pos_url:
        raise ValueError("Configure the Backoffice Dashboard and POS Payments routes.")

    state_setting = config.get("paths", {}).get("storage_state_path", "data/storage_state.json")
    state_path = Path(state_setting)
    if not state_path.is_absolute():
        state_path = PROJECT_ROOT / state_path
    state_path.parent.mkdir(parents=True, exist_ok=True)
    login_timeout = backoffice.get("browser", {}).get("login_timeout_ms", 600000)
    all_ok = True
    completed = {"backoffice": False, "pos": False}
    etl_processes = {}
    pending_etls = []
    etl_state_dir = tempfile.TemporaryDirectory(prefix="revenue_analysis_sessions_") if preflight else None

    try:
        try:
            with sync_playwright() as playwright:
                extractor = PlaywrightExtractor(config)
                if preflight:
                    print(
                        "Daily ETL is waiting for both logins. Chromium tab 1 is Backoffice; "
                        "tab 2 is POS Payments.", flush=True
                    )
                context = extractor.launch_persistent_context(playwright, headless=False)
                try:
                    pages = list(context.pages)
                    if not pages:
                        backoffice_page = context.new_page()
                        pos_page = context.new_page()
                    elif len(pages) == 1:
                        backoffice_page = pages[0]
                        pos_page = context.new_page()
                    else:
                        backoffice_page = pages[0]
                        pos_page = pages[1]
                        for extra in pages[2:]:
                            try:
                                extra.close()
                            except Exception:
                                pass

                    for page, target, label in (
                        (backoffice_page, dashboard_url, "Backoffice"),
                        (pos_page, pos_url, "POS Payments"),
                    ):
                        try:
                            page.goto(target, wait_until="domcontentloaded", timeout=60000)
                        except Exception as exc:
                            print(f"{label} is still loading ({exc}).", flush=True)

                    try:
                        backoffice_page.bring_to_front()
                    except Exception:
                        pass

                    deadline = time.monotonic() + login_timeout / 1000
                    next_notice = time.monotonic() + 300
                    while not all(completed.values()):
                        if backoffice_page.is_closed() or pos_page.is_closed():
                            missing = [name for name, done in completed.items() if not done]
                            print(
                                f"Browser tab was closed before all logins passed. Missing: {', '.join(missing)}.",
                                flush=True,
                            )
                            all_ok = False
                            break
                        if time.monotonic() >= deadline:
                            missing = [name for name, done in completed.items() if not done]
                            print(
                                f"Login timeout ({login_timeout // 1000}s) reached. Missing: {', '.join(missing)}.",
                                flush=True,
                            )
                            all_ok = False
                            break

                        candidates = {
                            "backoffice": _backoffice_ready(extractor, backoffice_page, dashboard_url),
                            "pos": _pos_ready(extractor, pos_page, pos_url),
                        }
                        for source in ("backoffice", "pos"):
                            if completed[source] or not candidates[source]:
                                continue
                            try:
                                context.storage_state(path=str(state_path))
                            except Exception as exc:
                                print(f"Warning: could not save session state for {source} ({exc}).", flush=True)
                            completed[source] = True
                            print(f"{source.title()} login confirmed.", flush=True)
                            if preflight:
                                source_state_path = Path(etl_state_dir.name) / f"{source}.json"
                                try:
                                    context.storage_state(path=str(source_state_path))
                                    child_env = os.environ.copy()
                                    child_env["REVENUE_ETL_STORAGE_STATE"] = str(source_state_path)
                                    command = [
                                        sys.executable,
                                        str(PROJECT_ROOT / "main.py"),
                                        "--config",
                                        str(config_path),
                                        "--source",
                                        source,
                                    ]
                                    if etl_processes:
                                        print(
                                            f"Queuing {source.title()} ETL until Chromium closes.",
                                            flush=True,
                                        )
                                        pending_etls.append((source, command, child_env))
                                    else:
                                        print(
                                            f"Starting {source.title()} ETL in the background.",
                                            flush=True,
                                        )
                                        etl_processes[source] = subprocess.Popen(
                                            command,
                                            cwd=str(PROJECT_ROOT),
                                            env=child_env,
                                        )
                                except Exception as exc:
                                    all_ok = False
                                    print(f"{source.title()} ETL failed to start: {exc}", flush=True)

                        if all(completed.values()):
                            break
                        if time.monotonic() >= next_notice:
                            missing = [name for name, done in completed.items() if not done]
                            print(
                                f"Still waiting for {', '.join(missing)} login. Keep Chromium open "
                                "and finish the remaining login.",
                                flush=True,
                            )
                            next_notice = time.monotonic() + 300
                        time.sleep(0.5)

                    try:
                        context.storage_state(path=str(state_path))
                    except Exception as exc:
                        print(f"Warning: could not save final session state ({exc}).", flush=True)
                    if all(completed.values()):
                        print("Both logins passed. Closing Chromium.", flush=True)
                    else:
                        print("Closing Chromium. Session state saved for completed logins.", flush=True)
                    for _page in list(context.pages):
                        try:
                            _page.close()
                        except Exception:
                            pass
                finally:
                    context.close()
        finally:
            for source in ("backoffice", "pos"):
                process = etl_processes.get(source)
                if process is None:
                    continue
                result = process.wait()
                if result != 0:
                    all_ok = False
                    print(f"{source.title()} ETL finished with errors (exit {result}).", flush=True)
                else:
                    print(f"{source.title()} ETL completed successfully.", flush=True)
            for source, command, child_env in pending_etls:
                try:
                    result = subprocess.run(
                        command,
                        cwd=str(PROJECT_ROOT),
                        env=child_env,
                    )
                except Exception as exc:
                    all_ok = False
                    print(f"{source.title()} ETL failed: {exc}", flush=True)
                    continue
                if result.returncode != 0:
                    all_ok = False
                    print(
                        f"{source.title()} ETL finished with errors "
                        f"(exit {result.returncode}).",
                        flush=True,
                    )
                else:
                    print(f"{source.title()} ETL completed successfully.", flush=True)
    finally:
        if etl_state_dir is not None:
            etl_state_dir.cleanup()
    return 0 if (all_ok and all(completed.values())) else 1


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--preflight", action="store_true",
        help="Authenticate both portals and run each source after its login passes.",
    )
    sys.exit(main(preflight=parser.parse_args().preflight))
