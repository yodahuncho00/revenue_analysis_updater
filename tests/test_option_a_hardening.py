"""Tests for Option A reliability hardening: checkpoints, retries, OneDrive commits, and reporting."""

from datetime import date, datetime
import json
from pathlib import Path
import ssl

import openpyxl
import pytest

from common.checkpoint import StageCheckpoint
from extractors.base import BalanceTableCapture
from extractors.playwright_extractor import (
    PlaywrightAuthenticationError,
    PlaywrightExtractionError,
    PlaywrightExtractor,
)
from integrations import google_sheets_sync
from integrations.google_sheets_sync import GoogleSheetsSync
from pos_pipeline import POS_HEADERS, run_pos_pipeline
from scripts import refresh_backoffice_session
from writers.calculator import ExcelCalculator
from writers.copy_manager import CopyManager, CopyManagerError


def test_stage_checkpoint_roundtrips_json_balance_and_pos_rows(tmp_path):
    ckpt = StageCheckpoint("backoffice", "2026-11-26", base_dir=str(tmp_path))
    assert not ckpt.is_stage_done("extract_dashboard")

    ckpt.save_json("dashboard", [{"country": "Zambia", "deposits": 100.5}])
    ckpt.mark_stage("extract_dashboard", details={"rows": 1})

    capture = BalanceTableCapture(
        rows=[{"row_class": "country-row", "cells": ["Zambia", "100.00"]}],
        html="<table><tbody><tr><td>Zambia</td><td>100.00</td></tr></tbody></table>",
        row_count=1,
        column_count=2,
    )
    ckpt.save_balance_capture("merchant_balance", capture)

    pos_rows = [[123, datetime(2026, 11, 25, 14, 30), "Receive", 50, date(2026, 11, 25)]]
    ckpt.save_pos_rows(pos_rows)

    reloaded = StageCheckpoint("backoffice", "2026-11-26", base_dir=str(tmp_path))
    assert reloaded.is_stage_done("extract_dashboard")
    assert reloaded.get_stage_details("extract_dashboard") == {"rows": 1}
    assert reloaded.load_json("dashboard") == [{"country": "Zambia", "deposits": 100.5}]
    assert reloaded.load_balance_capture("merchant_balance") == capture
    assert reloaded.load_pos_rows() == pos_rows


def test_copy_manager_commit_retries_transient_onedrive_permission_error(tmp_path, monkeypatch):
    prod = tmp_path / "prod.xlsx"
    prod.write_bytes(b"original")
    mgr = CopyManager(str(prod), backup_dir=str(tmp_path / "backups"), run_id="test_retry")
    mgr.working_copy_dir = str(tmp_path / "working")
    mgr.working_copy_path = str(tmp_path / "working" / "prod_test_retry.working.xlsx")
    Path(mgr.working_copy_dir).mkdir(parents=True, exist_ok=True)
    Path(mgr.working_copy_path).write_bytes(b"updated")

    monkeypatch.setattr("writers.copy_manager.time.sleep", lambda _: None)
    real_replace = CopyManager.commit.__globals__["os"].replace
    attempts = {"count": 0}

    def flaky_replace(src, dst):
        attempts["count"] += 1
        if attempts["count"] < 3:
            raise PermissionError(13, "Access denied", src, 5, dst)
        return real_replace(src, dst)

    monkeypatch.setattr("writers.copy_manager.os.replace", flaky_replace)

    mgr.commit()
    assert attempts["count"] == 3
    assert prod.read_bytes() == b"updated"


def test_copy_manager_commit_cleans_temp_file_when_retries_exhausted(tmp_path, monkeypatch):
    prod = tmp_path / "prod.xlsx"
    prod.write_bytes(b"original")
    mgr = CopyManager(str(prod), backup_dir=str(tmp_path / "backups"), run_id="test_fail")
    mgr.working_copy_dir = str(tmp_path / "working")
    mgr.working_copy_path = str(tmp_path / "working" / "prod_test_fail.working.xlsx")
    Path(mgr.working_copy_dir).mkdir(parents=True, exist_ok=True)
    Path(mgr.working_copy_path).write_bytes(b"updated")

    monkeypatch.setattr("writers.copy_manager.time.sleep", lambda _: None)
    monkeypatch.setattr(CopyManager, "COMMIT_RETRY_ATTEMPTS", 2)
    monkeypatch.setattr(
        "writers.copy_manager.os.replace",
        lambda *_args: (_ for _ in ()).throw(PermissionError("Locked by OneDrive")),
    )

    with pytest.raises(CopyManagerError, match="Could not replace production workbook"):
        mgr.commit()

    assert prod.read_bytes() == b"original"
    assert not list(tmp_path.glob("*.tmp"))


def test_playwright_retry_retries_transient_but_not_auth_errors(monkeypatch):
    monkeypatch.setattr("extractors.playwright_extractor.time.sleep", lambda _: None)
    extractor = PlaywrightExtractor(
        {"backoffice": {"browser": {"step_retry_attempts": 3, "step_retry_delay_seconds": 0}}}
    )

    calls = {"transient": 0, "auth": 0}

    def flaky():
        calls["transient"] += 1
        if calls["transient"] < 3:
            raise PlaywrightExtractionError("Merchant Balance table changed during capture.")
        return "ok"

    assert extractor._extract_with_retry("Merchant Balance", flaky) == "ok"
    assert calls["transient"] == 3

    def auth_fail():
        calls["auth"] += 1
        raise PlaywrightAuthenticationError("Backoffice authentication is required")

    with pytest.raises(PlaywrightAuthenticationError):
        extractor._extract_with_retry("Dashboard", auth_fail)
    assert calls["auth"] == 1


def test_pos_all_time_url_sets_period_all_cleanly():
    assert (
        PlaywrightExtractor._pos_all_time_url("https://pos-portal.unipesa.tech/payments")
        == "https://pos-portal.unipesa.tech/payments?period=all"
    )
    assert (
        PlaywrightExtractor._pos_all_time_url(
            "https://pos-portal.unipesa.tech/payments?period=today&page=1"
        )
        == "https://pos-portal.unipesa.tech/payments?page=1&period=all"
    )


def test_calculator_retries_servercall_retrylater(monkeypatch):
    calls = []
    monkeypatch.setattr(ExcelCalculator, "COM_RETRY_ATTEMPTS", 3)
    monkeypatch.setattr(ExcelCalculator, "COM_RETRY_DELAY_SECONDS", 0)

    class FakeComError(Exception):
        pass

    def op():
        calls.append(1)
        if len(calls) == 1:
            raise FakeComError(ExcelCalculator.RPC_E_SERVERCALL_RETRYLATER)
        return "recalculated"

    assert ExcelCalculator._retry_excel_call(op, FakeComError) == "recalculated"
    assert len(calls) == 2


def test_google_sheets_retries_503_and_connection_error_but_not_ssl(monkeypatch):
    import httplib2
    from googleapiclient.errors import HttpError

    monkeypatch.setattr(google_sheets_sync.time_module, "sleep", lambda _: None)

    class Flaky503:
        def __init__(self):
            self.calls = 0

        def execute(self):
            self.calls += 1
            if self.calls == 1:
                raise HttpError(httplib2.Response({"status": 503}), b"unavailable")
            if self.calls == 2:
                raise ConnectionError("reset by peer")
            return {"ok": True}

    req = Flaky503()
    assert GoogleSheetsSync._execute_write(req) == {"ok": True}
    assert req.calls == 3

    class SslFail:
        def __init__(self):
            self.calls = 0

        def execute(self):
            self.calls += 1
            raise ssl.SSLError("certificate verify failed")

    ssl_req = SslFail()
    with pytest.raises(ssl.SSLError):
        GoogleSheetsSync._execute_write(ssl_req)
    assert ssl_req.calls == 1


def test_pos_pipeline_resume_skips_extraction_and_commit_when_already_committed(
    tmp_path, monkeypatch
):
    workbook_path = tmp_path / "05. Revenue_analysis_POS.xlsx"
    dashboard_path = tmp_path / "05. Revenue_dashboard_POS.html"
    template_path = Path("templates/05. Revenue_dashboard_POS.template.html")
    checkpoints_dir = tmp_path / "checkpoints"

    wb = openpyxl.Workbook()
    ws_bo = wb.active
    ws_bo.title = "BO_extract"
    ws_bo.append(["Title"])
    ws_bo.append(["Instructions"])
    ws_bo.append(list(POS_HEADERS) + [""] * 36)
    row_data = [
        490580001,
        datetime(2026, 8, 10, 15, 0),
        "Receive",
        50,
        "ZMW",
        1,
        0.5,
        0.5,
        260111,
        "Agent",
        "A",
        "T",
        None,
        "Airtel",
        "Success: ",
        0,
        None,
        0,
    ]
    ws_bo.append(row_data + [None] * 36)
    ws_ctrl = wb.create_sheet("Control")
    ws_ctrl["B5"] = datetime(2026, 11, 25)
    ws_ctrl["B6"] = datetime(2026, 11, 25)
    ws_ctrl["B20"] = "ZMW"
    wb.save(workbook_path)
    wb.close()

    # Pre-seed checkpoint as if workbook_commit succeeded earlier today
    ckpt = StageCheckpoint("pos", "2026-11-26", base_dir=str(checkpoints_dir))
    ckpt.save_pos_rows([row_data])
    ckpt.mark_stage("extract_pos_payments", details={"rows": 1})
    ckpt.mark_stage("workbook_commit", details={"rows": 1})

    # Extractor and ExcelCalculator must NOT be called during post-commit resume
    monkeypatch.setattr(
        PlaywrightExtractor,
        "extract_pos_payments",
        lambda _self: (_ for _ in ()).throw(AssertionError("Should not re-extract")),
    )
    monkeypatch.setattr(
        ExcelCalculator,
        "recalculate",
        lambda _path: (_ for _ in ()).throw(AssertionError("Should not recalculate")),
    )

    cfg = {
        "paths": {
            "pos_workbook_path": str(workbook_path),
            "pos_dashboard_path": str(dashboard_path),
            "pos_dashboard_template_path": str(template_path),
            "backup_dir": str(tmp_path / "backups"),
            "raw_dir": str(tmp_path / "raw"),
            "logs_dir": str(tmp_path / "logs"),
            "reports_dir": str(tmp_path / "reports"),
            "checkpoints_dir": str(checkpoints_dir),
        },
        "app": {"backups_to_retain": 5},
    }

    res = run_pos_pipeline(cfg, dry_run=False, resume=True, override_date="2026-11-26")
    assert res["success"] is True
    assert dashboard_path.is_file()
    assert StageCheckpoint("pos", "2026-11-26", base_dir=str(checkpoints_dir)).is_stage_done(
        "publish_dashboard"
    )


def test_preflight_writes_per_source_summary_when_one_login_fails(tmp_path, monkeypatch):
    config_path = tmp_path / "config.yaml"
    config_path.write_text(
        """
paths:
  storage_state_path: data/storage_state.json
  logs_dir: data/logs
  reports_dir: data/reports
backoffice:
  routes:
    dashboard: https://backoffice.example/dashboard
  browser:
    login_timeout_ms: 1000
pos_portal:
  payments_url: https://pos.example/payments
""",
        encoding="utf-8",
    )
    monkeypatch.setattr(refresh_backoffice_session, "PROJECT_ROOT", tmp_path)
    monkeypatch.setattr(refresh_backoffice_session, "_backoffice_ready", lambda *_args: False)
    monkeypatch.setattr(refresh_backoffice_session, "_pos_ready", lambda *_args: True)

    class Page:
        def __init__(self, close_after=None):
            self.checks = 0
            self.close_after = close_after

        def goto(self, *_args, **_kwargs):
            pass

        def bring_to_front(self):
            pass

        def is_closed(self):
            self.checks += 1
            return self.close_after is not None and self.checks >= self.close_after

        def close(self):
            pass

    # Backoffice tab closes on 2nd poll after POS confirms
    bo_page = Page(close_after=2)
    pos_page = Page()

    class Context:
        pages = [bo_page, pos_page]

        def storage_state(self, *, path):
            Path(path).write_text("{}", encoding="utf-8")

        def close(self):
            pass

    class Extractor:
        def __init__(self, _config):
            pass

        def launch_persistent_context(self, *_args, **_kwargs):
            return Context()

    monkeypatch.setattr(refresh_backoffice_session, "PlaywrightExtractor", Extractor)

    class PlaywrightManager:
        def __enter__(self):
            return object()

        def __exit__(self, *_args):
            pass

    monkeypatch.setattr(
        refresh_backoffice_session, "sync_playwright", lambda: PlaywrightManager()
    )
    monkeypatch.setattr(refresh_backoffice_session.time, "sleep", lambda _: None)

    ran_sources = []

    def fake_run(command, **_kwargs):
        source = command[-1]
        ran_sources.append(source)
        logs_dir = tmp_path / "data" / "logs"
        logs_dir.mkdir(parents=True, exist_ok=True)
        (logs_dir / f"latest_run_{source}.txt").write_text(f"run_{source}_123\n", encoding="utf-8")
        return type("Result", (), {"returncode": 0})()

    monkeypatch.setattr(refresh_backoffice_session.subprocess, "run", fake_run)

    exit_code = refresh_backoffice_session.main(preflight=True)
    assert exit_code == 1
    assert ran_sources == ["pos"]

    summary = json.loads((tmp_path / "data" / "logs" / "preflight_summary.json").read_text(encoding="utf-8"))
    assert summary["overall_ok"] is False
    assert summary["sources"]["pos"]["etl_status"] == "success"
    assert summary["sources"]["pos"]["run_id"] == "run_pos_123"
    assert summary["sources"]["backoffice"]["etl_status"] == "skipped_login"
    assert "Browser tab was closed" in summary["sources"]["backoffice"]["error"]


def test_load_config_deep_merges_local_override(tmp_path):
    from main import load_config

    (tmp_path / "config.yaml").write_text(
        """
paths:
  workbook_path: "default_gw.xlsx"
  pos_workbook_path: "default_pos.xlsx"
google_sheets:
  enabled: true
  folder_id: "base-folder"
""",
        encoding="utf-8",
    )
    (tmp_path / "config.local.yaml").write_text(
        """
paths:
  workbook_path: "C:/Custom/gw.xlsx"
auto_update:
  enabled: true
  branch: "release"
""",
        encoding="utf-8",
    )

    cfg = load_config(str(tmp_path / "config.yaml"))
    assert cfg["paths"]["workbook_path"] == "C:/Custom/gw.xlsx"
    assert cfg["paths"]["pos_workbook_path"] == "default_pos.xlsx"
    assert cfg["google_sheets"]["folder_id"] == "base-folder"
    assert cfg["auto_update"] == {"enabled": True, "branch": "release"}


def test_auto_update_script_updates_guards_dirty_and_rolls_back_syntax_errors(tmp_path):
    import shutil
    import subprocess
    import sys

    if sys.platform != "win32" or not shutil.which("git") or not shutil.which("powershell"):
        pytest.skip("Windows PowerShell + git required for auto_update.ps1 end-to-end test.")

    script_path = Path("scripts/auto_update.ps1").resolve()
    remote_dir = tmp_path / "remote.git"
    dev_dir = tmp_path / "dev"
    client_dir = tmp_path / "client"

    def git(cwd, *args):
        res = subprocess.run(
            ["git", *args],
            cwd=str(cwd),
            capture_output=True,
            text=True,
            check=True,
        )
        return res.stdout.strip()

    subprocess.run(["git", "init", "--bare", str(remote_dir)], check=True, capture_output=True)
    dev_dir.mkdir()
    git(dev_dir, "init")
    git(dev_dir, "config", "user.email", "test@example.com")
    git(dev_dir, "config", "user.name", "Test")
    (dev_dir / "scripts").mkdir()
    (dev_dir / "main.py").write_text("VALUE = 1\n", encoding="utf-8")
    (dev_dir / "scripts" / "run_daily_etl.bat").write_text("@echo v1\n", encoding="utf-8")
    git(dev_dir, "add", ".")
    git(dev_dir, "commit", "-m", "v1")
    git(dev_dir, "branch", "-M", "release")
    git(dev_dir, "remote", "add", "origin", str(remote_dir))
    git(dev_dir, "push", "-u", "origin", "release")

    subprocess.run(
        ["git", "clone", "-b", "release", str(remote_dir), str(client_dir)],
        check=True,
        capture_output=True,
    )

    def run_updater():
        return subprocess.run(
            [
                "powershell",
                "-NoProfile",
                "-ExecutionPolicy",
                "Bypass",
                "-File",
                str(script_path),
                "-ProjectDir",
                str(client_dir),
                "-PythonExe",
                sys.executable,
            ],
            capture_output=True,
            text=True,
        )

    # Push v2 from dev
    (dev_dir / "main.py").write_text("VALUE = 2\n", encoding="utf-8")
    git(dev_dir, "commit", "-am", "v2")
    git(dev_dir, "push", "origin", "release")

    # 1. Opt-in disabled by default -> client stays on v1
    r0 = run_updater()
    assert r0.returncode == 0
    assert (client_dir / "main.py").read_text(encoding="utf-8") == "VALUE = 1\n"

    # Enable auto_update in config.local.yaml
    (client_dir / "config.local.yaml").write_text(
        "auto_update:\n  enabled: true\n  branch: release\n", encoding="utf-8"
    )

    # 2. Dirty tracked file -> refuses to overwrite
    (client_dir / "main.py").write_text("VALUE = 999\n", encoding="utf-8")
    r_dirty = run_updater()
    assert r_dirty.returncode == 0
    assert (client_dir / "main.py").read_text(encoding="utf-8") == "VALUE = 999\n"
    git(client_dir, "checkout", "--", "main.py")

    # 3. Clean client -> updates to v2 cleanly
    r_ok = run_updater()
    assert r_ok.returncode == 0
    assert (client_dir / "main.py").read_text(encoding="utf-8") == "VALUE = 2\n"

    # 4. Broken syntax pushed to release -> auto-rolls back to v2
    (dev_dir / "main.py").write_text("def broken(:\n", encoding="utf-8")
    git(dev_dir, "commit", "-am", "broken")
    git(dev_dir, "push", "origin", "release")
    r_bad = run_updater()
    assert r_bad.returncode == 0
    assert "Rolling back" in r_bad.stdout
    assert (client_dir / "main.py").read_text(encoding="utf-8") == "VALUE = 2\n"

    # 5. Valid update that also changes run_daily_etl.bat -> returns exit code 10 (re-exec)
    (dev_dir / "main.py").write_text("VALUE = 3\n", encoding="utf-8")
    (dev_dir / "scripts" / "run_daily_etl.bat").write_text("@echo v3\n", encoding="utf-8")
    git(dev_dir, "commit", "-am", "v3")
    git(dev_dir, "push", "origin", "release")
    r_bat = run_updater()
    assert r_bat.returncode == 10
    assert (client_dir / "main.py").read_text(encoding="utf-8") == "VALUE = 3\n"


