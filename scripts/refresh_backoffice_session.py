"""Authenticate Backoffice and POS in one browser session, then run ETLs sequentially."""

import argparse
import json
import os
import re
import subprocess
import sys
import tempfile
import time
from datetime import datetime
from pathlib import Path
from urllib.parse import urlparse

import yaml
from playwright.sync_api import sync_playwright

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT))

from extractors.playwright_extractor import PlaywrightExtractor
from main import load_config


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
        if page.locator("button, [role='button'], a").filter(
            has_text=re.compile(r"Save\s*CSV", re.I)
        ).first.is_visible():
            return True
        return False
    except Exception:
        return False


def _probe_existing_sessions_headless(
    extractor, dashboard_url: str, pos_url: str, state_path: Path, snapshot_dir: Path
) -> dict:
    """Quickly verify if saved sessions are already valid without opening a visible browser."""
    confirmed = {"backoffice": False, "pos": False}
    try:
        with sync_playwright() as playwright:
            context = extractor.launch_persistent_context(playwright, headless=True)
            try:
                for source, url, checker in (
                    ("backoffice", dashboard_url, _backoffice_ready),
                    ("pos", pos_url, _pos_ready),
                ):
                    page = context.new_page()
                    try:
                        page.goto(url, wait_until="domcontentloaded", timeout=20000)
                        # Brief settle for client-side redirects / SPA render
                        deadline = time.monotonic() + 8.0
                        while time.monotonic() < deadline:
                            if checker(extractor, page, url):
                                confirmed[source] = True
                                break
                            time.sleep(0.4)
                    except Exception:
                        pass
                    finally:
                        try:
                            if not page.is_closed():
                                page.close()
                        except Exception:
                            pass

                if all(confirmed.values()):
                    try:
                        context.storage_state(path=str(state_path))
                    except Exception:
                        pass
                    if snapshot_dir is not None:
                        for source in ("backoffice", "pos"):
                            try:
                                context.storage_state(path=str(snapshot_dir / f"{source}.json"))
                            except Exception:
                                pass
            finally:
                context.close()
    except Exception:
        return {"backoffice": False, "pos": False}
    return confirmed


def _read_latest_run_id(logs_dir: Path, source: str):
    path = logs_dir / f"latest_run_{source}.txt"
    if not path.is_file():
        return None
    try:
        run_id = path.read_text(encoding="utf-8").strip()
        return run_id or None
    except OSError:
        return None


def _write_preflight_summary(logs_dir: Path, summary: dict) -> None:
    try:
        logs_dir.mkdir(parents=True, exist_ok=True)
        summary_path = logs_dir / "preflight_summary.json"
        fd, tmp_path = tempfile.mkstemp(prefix=".summary_", suffix=".json", dir=str(logs_dir))
        try:
            with os.fdopen(fd, "w", encoding="utf-8") as handle:
                json.dump(summary, handle, indent=2, ensure_ascii=False)
                handle.flush()
                os.fsync(handle.fileno())
            os.replace(tmp_path, summary_path)
        finally:
            if os.path.exists(tmp_path):
                os.remove(tmp_path)
    except Exception:
        pass


def main(preflight: bool = False, resume: bool = False, force_visible: bool = False) -> int:
    config_path = PROJECT_ROOT / "config.yaml"
    config = load_config(str(config_path))
    backoffice = config.get("backoffice", {})
    dashboard_url = backoffice.get("routes", {}).get("dashboard")
    pos_url = config.get("pos_portal", {}).get(
        "payments_url", "https://pos-portal.unipesa.tech/payments"
    )
    if not dashboard_url or not pos_url:
        raise ValueError("Configure the Backoffice Dashboard and POS Payments routes.")

    paths_cfg = config.get("paths", {})
    logs_dir = Path(paths_cfg.get("logs_dir", "data/logs"))
    if not logs_dir.is_absolute():
        logs_dir = PROJECT_ROOT / logs_dir
    reports_dir = Path(paths_cfg.get("reports_dir", "data/reports"))
    if not reports_dir.is_absolute():
        reports_dir = PROJECT_ROOT / reports_dir

    state_setting = paths_cfg.get("storage_state_path", "data/storage_state.json")
    state_path = Path(state_setting)
    if not state_path.is_absolute():
        state_path = PROJECT_ROOT / state_path
    state_path.parent.mkdir(parents=True, exist_ok=True)
    login_timeout = backoffice.get("browser", {}).get("login_timeout_ms", 600000)

    all_ok = True
    completed = {"backoffice": False, "pos": False}
    login_errors = {"backoffice": None, "pos": None}
    source_state_paths = {}
    etl_state_dir = (
        tempfile.TemporaryDirectory(prefix="revenue_analysis_sessions_")
        if preflight
        else None
    )
    summary = {
        "timestamp": datetime.now().isoformat(),
        "preflight": preflight,
        "resume": resume,
        "overall_ok": False,
        "sources": {
            "backoffice": {
                "login_confirmed": False,
                "etl_status": "pending",
                "exit_code": None,
                "run_id": None,
                "log_path": None,
                "report_path": None,
                "error": None,
            },
            "pos": {
                "login_confirmed": False,
                "etl_status": "pending",
                "exit_code": None,
                "run_id": None,
                "log_path": None,
                "report_path": None,
                "error": None,
            },
        },
    }

    try:
        extractor = PlaywrightExtractor(config)

        # Non-blocking preflight health check for Google Sheets OAuth so expired/missing
        # tokens are surfaced immediately without hanging unattended runs.
        google_cfg = config.get("google_sheets", {})
        if preflight and google_cfg.get("enabled", False) and google_cfg.get("auth_type") == "oauth_user":
            try:
                from integrations.google_sheets_sync import _oauth_user_credentials

                _oauth_user_credentials(
                    google_cfg, str(PROJECT_ROOT), allow_interactive=False
                )
                print("Google Sheets OAuth token verified.", flush=True)
            except Exception as exc:
                print(
                    f"Warning: Google Sheets OAuth check failed ({exc}). "
                    "Local workbooks/dashboards will still update; run "
                    "scripts\\authorize_google_sheets.py then re-run with --resume.",
                    flush=True,
                )

        # Fast headless check first during preflight when an existing saved session exists,
        # so unattended scheduled runs succeed without popping up a browser when cookies are valid.
        has_saved_state = state_path.is_file()
        if preflight and not force_visible and has_saved_state:
            print("Checking saved Backoffice and POS sessions headlessly...", flush=True)
            headless_ok = _probe_existing_sessions_headless(
                extractor,
                dashboard_url,
                pos_url,
                state_path,
                Path(etl_state_dir.name) if etl_state_dir is not None else None,
            )
            if all(headless_ok.values()):
                completed["backoffice"] = True
                completed["pos"] = True
                for src in ("backoffice", "pos"):
                    snapshot = Path(etl_state_dir.name) / f"{src}.json"
                    if snapshot.is_file():
                        source_state_paths[src] = snapshot
                print(
                    "Both Backoffice and POS sessions are already valid; skipping interactive login.",
                    flush=True,
                )

        if not all(completed.values()):
            with sync_playwright() as playwright:
                if preflight:
                    print(
                        "Daily ETL is waiting for both logins. Chromium tab 1 is Backoffice; "
                        "tab 2 is POS Payments.",
                        flush=True,
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
                            msg = (
                                "Browser tab was closed before all logins passed. "
                                f"Missing: {', '.join(missing)}."
                            )
                            print(msg, flush=True)
                            for name in missing:
                                login_errors[name] = msg
                            all_ok = False
                            break
                        if time.monotonic() >= deadline:
                            missing = [name for name, done in completed.items() if not done]
                            msg = (
                                f"Login timeout ({login_timeout // 1000}s) reached. "
                                f"Missing: {', '.join(missing)}."
                            )
                            print(msg, flush=True)
                            for name in missing:
                                login_errors[name] = msg
                            all_ok = False
                            break

                        candidates = {
                            "backoffice": (
                                True
                                if completed["backoffice"]
                                else _backoffice_ready(extractor, backoffice_page, dashboard_url)
                            ),
                            "pos": (
                                True
                                if completed["pos"]
                                else _pos_ready(extractor, pos_page, pos_url)
                            ),
                        }
                        for source in ("backoffice", "pos"):
                            if completed[source] or not candidates[source]:
                                continue
                            try:
                                context.storage_state(path=str(state_path))
                            except Exception as exc:
                                print(
                                    f"Warning: could not save session state for {source} ({exc}).",
                                    flush=True,
                                )
                            completed[source] = True
                            print(f"{source.title()} login confirmed.", flush=True)
                            if preflight and etl_state_dir is not None:
                                snap = Path(etl_state_dir.name) / f"{source}.json"
                                try:
                                    context.storage_state(path=str(snap))
                                    source_state_paths[source] = snap
                                except Exception as exc:
                                    all_ok = False
                                    login_errors[source] = f"Could not snapshot session state: {exc}"
                                    print(
                                        f"Warning: could not snapshot session state for {source} ({exc}).",
                                        flush=True,
                                    )

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
                        print(
                            "Closing Chromium. Session state saved for completed logins.",
                            flush=True,
                        )
                    for _page in list(context.pages):
                        try:
                            _page.close()
                        except Exception:
                            pass
                finally:
                    context.close()

        # Run confirmed ETLs sequentially AFTER Chromium is cleanly closed (eliminates
        # concurrent browser/profile/session races while still updating any confirmed source).
        if preflight:
            for source in ("backoffice", "pos"):
                summary["sources"][source]["login_confirmed"] = completed[source]
                if not completed[source]:
                    summary["sources"][source]["etl_status"] = "skipped_login"
                    summary["sources"][source]["error"] = (
                        login_errors[source] or f"{source.title()} login did not complete."
                    )
                    continue

                # Clear stale per-source marker before launching this source
                source_marker = logs_dir / f"latest_run_{source}.txt"
                if source_marker.is_file():
                    try:
                        source_marker.unlink()
                    except OSError:
                        pass

                child_env = os.environ.copy()
                snap = source_state_paths.get(source)
                if snap is not None and snap.is_file():
                    child_env["REVENUE_ETL_STORAGE_STATE"] = str(snap)

                command = [
                    sys.executable,
                    str(PROJECT_ROOT / "main.py"),
                    "--config",
                    str(config_path),
                    "--source",
                    source,
                ]
                if resume:
                    command.append("--resume")

                print(f"Starting {source.title()} ETL...", flush=True)
                try:
                    result = subprocess.run(
                        command,
                        cwd=str(PROJECT_ROOT),
                        env=child_env,
                    )
                except Exception as exc:
                    all_ok = False
                    summary["sources"][source]["etl_status"] = "failed"
                    summary["sources"][source]["error"] = str(exc)
                    print(f"{source.title()} ETL failed to execute: {exc}", flush=True)
                    continue

                run_id = _read_latest_run_id(logs_dir, source)
                summary["sources"][source]["exit_code"] = result.returncode
                summary["sources"][source]["run_id"] = run_id
                if run_id:
                    summary["sources"][source]["log_path"] = str(logs_dir / f"{run_id}.log")
                    report_file = reports_dir / f"{run_id}.json"
                    if report_file.is_file():
                        summary["sources"][source]["report_path"] = str(report_file)

                if result.returncode != 0:
                    all_ok = False
                    summary["sources"][source]["etl_status"] = "failed"
                    print(
                        f"{source.title()} ETL finished with errors (exit {result.returncode}).",
                        flush=True,
                    )
                else:
                    summary["sources"][source]["etl_status"] = "success"
                    print(f"{source.title()} ETL completed successfully.", flush=True)
    finally:
        summary["overall_ok"] = bool(all_ok and all(completed.values()))
        if preflight:
            _write_preflight_summary(logs_dir, summary)
        if etl_state_dir is not None:
            etl_state_dir.cleanup()
    return 0 if (all_ok and all(completed.values())) else 1


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--preflight",
        action="store_true",
        help="Authenticate both portals and run each source sequentially after Chromium closes.",
    )
    parser.add_argument(
        "--resume",
        action="store_true",
        help="Pass --resume to source ETL runs to reuse completed same-day checkpoints.",
    )
    parser.add_argument(
        "--force-visible",
        action="store_true",
        help="Always open the visible Chromium window instead of checking saved sessions headlessly first.",
    )
    args = parser.parse_args()
    sys.exit(
        main(
            preflight=args.preflight,
            resume=args.resume,
            force_visible=args.force_visible,
        )
    )
