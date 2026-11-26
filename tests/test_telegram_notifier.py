"""Unit tests for the opt-in Telegram notifier."""

import json

from notify.telegram_notifier import TelegramNotifier


class _FakeResponse:
    def __init__(self, status_code=200):
        self.status_code = status_code
        self.ok = 200 <= status_code < 300


class _FakeSession:
    def __init__(self, status_codes=None):
        self.calls = []
        self._codes = list(status_codes or [200])

    def post(self, url, data=None, files=None, timeout=None):
        self.calls.append({"url": url, "data": data, "files": files, "timeout": timeout})
        code = self._codes.pop(0) if len(self._codes) > 1 else self._codes[0]
        return _FakeResponse(code)

    def close(self):
        pass


def test_telegram_disabled_is_noop():
    notifier = TelegramNotifier.from_config({})
    assert notifier.enabled is False
    assert notifier.notify_preflight_summary({"overall_ok": False, "sources": {}}) is False


def test_telegram_success_respects_notify_on_success_flag():
    session = _FakeSession()
    notifier = TelegramNotifier(
        enabled=True,
        bot_token="tok",
        chat_id="123",
        machine_name="Brother-PC",
        notify_on_success=False,
        session=session,
    )
    summary = {
        "overall_ok": True,
        "timestamp": "2026-11-26T09:15:00",
        "sources": {
            "backoffice": {"etl_status": "success", "run_id": "run_bo"},
            "pos": {"etl_status": "success", "run_id": "run_pos"},
        },
    }
    assert notifier.notify_preflight_summary(summary) is False
    assert session.calls == []

    notifier.notify_on_success = True
    assert notifier.notify_preflight_summary(summary) is True
    assert len(session.calls) == 1
    assert session.calls[0]["url"].endswith("/bottok/sendMessage")
    assert "Revenue ETL OK [Brother-PC]" in session.calls[0]["data"]["text"]


def test_telegram_failure_includes_report_error_and_attaches_log(tmp_path):
    log_file = tmp_path / "run_bo.log"
    log_file.write_text("2026-11-26 ERROR boom\n", encoding="utf-8")
    report_file = tmp_path / "run_bo.json"
    report_file.write_text(
        json.dumps(
            {
                "error_message": "Google OAuth token expired",
                "checks_passed": ["atomic_commit_successful"],
            }
        ),
        encoding="utf-8",
    )

    session = _FakeSession()
    notifier = TelegramNotifier(
        enabled=True,
        bot_token="tok",
        chat_id="123",
        machine_name="Brother-PC",
        attach_logs_on_failure=True,
        session=session,
    )
    summary = {
        "overall_ok": False,
        "timestamp": "2026-11-26T09:15:00",
        "sources": {
            "backoffice": {
                "etl_status": "failed",
                "run_id": "run_bo",
                "log_path": str(log_file),
                "report_path": str(report_file),
            },
            "pos": {"etl_status": "success", "run_id": "run_pos"},
        },
    }

    assert notifier.notify_preflight_summary(summary, logs_dir=tmp_path) is True
    assert len(session.calls) == 2
    msg_text = session.calls[0]["data"]["text"]
    assert "Revenue ETL FAILED [Brother-PC]" in msg_text
    assert "Google OAuth token expired" in msg_text
    assert "retry with --resume" in msg_text

    doc_call = session.calls[1]
    assert doc_call["url"].endswith("/bottok/sendDocument")
    filename, stream, _mime = doc_call["files"]["document"]
    assert filename == "backoffice_run_bo.log"
    assert b"ERROR boom" in stream.read()


def test_telegram_failure_falls_back_to_scheduler_tail_when_no_run_log(tmp_path):
    sched = tmp_path / "scheduler.log"
    sched.write_text("line1\nBrowser tab was closed before all logins passed.\n", encoding="utf-8")
    session = _FakeSession()
    notifier = TelegramNotifier(
        enabled=True,
        bot_token="tok",
        chat_id="123",
        machine_name="Brother-PC",
        session=session,
    )
    summary = {
        "overall_ok": False,
        "timestamp": "2026-11-26T09:15:00",
        "sources": {
            "backoffice": {
                "etl_status": "skipped_login",
                "error": "Browser tab was closed before all logins passed.",
            },
            "pos": {"etl_status": "success", "run_id": "run_pos"},
        },
    }

    assert notifier.notify_preflight_summary(summary, logs_dir=tmp_path) is True
    assert len(session.calls) == 2
    filename, stream, _mime = session.calls[1]["files"]["document"]
    assert filename == "scheduler_tail.log"
    assert b"Browser tab was closed" in stream.read()


def test_telegram_network_error_never_raises(monkeypatch):
    monkeypatch.setattr("notify.telegram_notifier.time.sleep", lambda _: None)

    class BrokenSession:
        def post(self, *_args, **_kwargs):
            raise ConnectionError("offline")

        def close(self):
            pass

    notifier = TelegramNotifier(
        enabled=True,
        bot_token="tok",
        chat_id="123",
        session=BrokenSession(),
    )
    assert notifier.notify_preflight_summary({"overall_ok": False, "sources": {}}) is False
