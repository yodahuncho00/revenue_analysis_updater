"""Send concise ETL status summaries and failure log attachments via Telegram Bot API."""

import io
import json
import os
import socket
import ssl
import tempfile
import time
from pathlib import Path
from typing import Any, Dict, List, Optional


TELEGRAM_API_BASE = "https://api.telegram.org"
REQUEST_TIMEOUT_SECONDS = 15
SEND_RETRIES = 2
MAX_MESSAGE_LEN = 3900
SCHEDULER_TAIL_LINES = 150


def _telegram_ca_bundle() -> Any:
    """Use combined certifi + Windows trusted ROOT store (same pattern as Google APIs)."""
    try:
        import certifi
    except ImportError:
        return True

    certifi_path = certifi.where()
    if not hasattr(ssl, "enum_certificates"):
        return certifi_path

    try:
        server_auth_oid = "1.3.6.1.5.5.7.3.1"
        windows_roots = []
        for certificate, encoding, trust in ssl.enum_certificates("ROOT"):
            if encoding != "x509_asn":
                continue
            if trust is not True and trust and server_auth_oid not in trust:
                continue
            windows_roots.append(ssl.DER_cert_to_PEM_cert(certificate).encode("ascii"))
        if not windows_roots:
            return certifi_path

        with open(certifi_path, "rb") as certifi_file:
            bundle_contents = certifi_file.read() + b"\n" + b"\n".join(windows_roots)

        cache_root = os.environ.get("LOCALAPPDATA") or tempfile.gettempdir()
        cache_dir = os.path.join(cache_root, "RevenueAnalysisUpdater")
        os.makedirs(cache_dir, exist_ok=True)
        bundle_path = os.path.join(cache_dir, "google-api-ca-bundle.pem")
        bundle_is_current = False
        if os.path.isfile(bundle_path):
            with open(bundle_path, "rb") as existing:
                bundle_is_current = existing.read() == bundle_contents
        if not bundle_is_current:
            fd, tmp_path = tempfile.mkstemp(prefix="tg-ca-", suffix=".pem", dir=cache_dir)
            try:
                with os.fdopen(fd, "wb") as handle:
                    handle.write(bundle_contents)
                os.replace(tmp_path, bundle_path)
            finally:
                if os.path.exists(tmp_path):
                    os.remove(tmp_path)
        return bundle_path
    except Exception:
        return certifi_path


def _read_report_error(report_path: Optional[str]) -> Optional[str]:
    if not report_path or not os.path.isfile(report_path):
        return None
    try:
        data = json.loads(Path(report_path).read_text(encoding="utf-8"))
        err = data.get("error_message")
        return str(err).strip() if err else None
    except Exception:
        return None


def _read_report_committed(report_path: Optional[str]) -> bool:
    if not report_path or not os.path.isfile(report_path):
        return False
    try:
        data = json.loads(Path(report_path).read_text(encoding="utf-8"))
        return "atomic_commit_successful" in (data.get("checks_passed") or [])
    except Exception:
        return False


def _tail_text(path: Path, max_lines: int = SCHEDULER_TAIL_LINES) -> bytes:
    text = path.read_text(encoding="utf-8", errors="replace")
    lines = text.splitlines()[-max_lines:]
    return ("\n".join(lines) + "\n").encode("utf-8")


class TelegramNotifier:
    """Opt-in Telegram Bot notifier configured via config.local.yaml."""

    def __init__(
        self,
        *,
        enabled: bool = False,
        bot_token: str = "",
        chat_id: str = "",
        machine_name: str = "",
        notify_on_success: bool = True,
        attach_logs_on_failure: bool = True,
        session: Any = None,
        verify: Any = None,
    ) -> None:
        self.enabled = bool(enabled and bot_token and chat_id)
        self.bot_token = str(bot_token or "").strip()
        self.chat_id = str(chat_id or "").strip()
        self.machine_name = str(machine_name or "").strip() or socket.gethostname()
        self.notify_on_success = bool(notify_on_success)
        self.attach_logs_on_failure = bool(attach_logs_on_failure)
        self._session = session
        self._verify = verify

    @classmethod
    def from_config(cls, config: Dict[str, Any]) -> "TelegramNotifier":
        tg = config.get("telegram") or {}
        if not isinstance(tg, dict):
            return cls(enabled=False)
        return cls(
            enabled=bool(tg.get("enabled", False)),
            bot_token=str(tg.get("bot_token", "") or ""),
            chat_id=str(tg.get("chat_id", "") or ""),
            machine_name=str(tg.get("machine_name", "") or ""),
            notify_on_success=bool(tg.get("notify_on_success", True)),
            attach_logs_on_failure=bool(tg.get("attach_logs_on_failure", True)),
        )

    def _get_session(self):
        if self._session is not None:
            return self._session, False
        import requests

        session = requests.Session()
        session.verify = self._verify if self._verify is not None else _telegram_ca_bundle()
        return session, True

    def _post(self, method: str, *, data: Dict[str, Any], files: Optional[Dict[str, Any]] = None) -> bool:
        if not self.enabled:
            return False
        url = f"{TELEGRAM_API_BASE}/bot{self.bot_token}/{method}"
        session, owns_session = self._get_session()
        try:
            for attempt in range(SEND_RETRIES + 1):
                try:
                    response = session.post(
                        url,
                        data=data,
                        files=files,
                        timeout=REQUEST_TIMEOUT_SECONDS,
                    )
                    if response.ok:
                        return True
                    if response.status_code < 500 and response.status_code != 429:
                        return False
                except Exception:
                    if attempt == SEND_RETRIES:
                        return False
                if attempt < SEND_RETRIES:
                    time.sleep(1.5 * (attempt + 1))
            return False
        finally:
            if owns_session:
                session.close()

    def send_message(self, text: str) -> bool:
        trimmed = text if len(text) <= MAX_MESSAGE_LEN else text[: MAX_MESSAGE_LEN - 20] + "\n...(truncated)"
        return self._post(
            "sendMessage",
            data={
                "chat_id": self.chat_id,
                "text": trimmed,
                "disable_web_page_preview": "true",
            },
        )

    def send_document(self, filename: str, content: bytes, caption: str = "") -> bool:
        if not content:
            return False
        data = {"chat_id": self.chat_id}
        if caption:
            data["caption"] = caption[:950]
        return self._post(
            "sendDocument",
            data=data,
            files={"document": (filename, io.BytesIO(content), "text/plain")},
        )

    def format_preflight_summary(self, summary: Dict[str, Any]) -> str:
        overall_ok = bool(summary.get("overall_ok"))
        badge = "OK" if overall_ok else "FAILED"
        icon = "✅" if overall_ok else "❌"
        lines: List[str] = [
            f"{icon} Revenue ETL {badge} [{self.machine_name}]",
            f"Time: {str(summary.get('timestamp', ''))[:19].replace('T', ' ')}",
        ]
        if summary.get("resume"):
            lines.append("Mode: --resume")

        sources = summary.get("sources") or {}
        for source_name in ("backoffice", "pos"):
            info = sources.get(source_name)
            if not isinstance(info, dict):
                continue
            status = info.get("etl_status", "unknown")
            run_id = info.get("run_id")
            line = f"- {source_name.upper()}: {status}"
            if run_id:
                line += f" ({run_id})"
            lines.append(line)

            error_text = info.get("error") or _read_report_error(info.get("report_path"))
            if error_text and status != "success":
                short_err = " ".join(str(error_text).split())
                if len(short_err) > 400:
                    short_err = short_err[:397] + "..."
                lines.append(f"  Error: {short_err}")
            if status != "success" and _read_report_committed(info.get("report_path")):
                lines.append("  Note: Local workbook committed; retry with --resume.")

        return "\n".join(lines)

    def notify_preflight_summary(
        self, summary: Dict[str, Any], logs_dir: Optional[Path] = None
    ) -> bool:
        """Send summary text and (on failure) attach failing log files. Never raises."""
        if not self.enabled:
            return False
        overall_ok = bool(summary.get("overall_ok"))
        if overall_ok and not self.notify_on_success:
            return False

        try:
            sent = self.send_message(self.format_preflight_summary(summary))
            if not overall_ok and self.attach_logs_on_failure:
                attached_any = False
                sources = summary.get("sources") or {}
                for source_name in ("backoffice", "pos"):
                    info = sources.get(source_name) or {}
                    if info.get("etl_status") == "success":
                        continue
                    log_path = info.get("log_path")
                    if log_path and os.path.isfile(log_path):
                        raw_bytes = Path(log_path).read_bytes()
                        if len(raw_bytes) > 512 * 1024:
                            raw_bytes = raw_bytes[-512 * 1024 :]
                        self.send_document(
                            f"{source_name}_{Path(log_path).name}",
                            raw_bytes,
                            caption=f"{self.machine_name} — {source_name.upper()} failure log",
                        )
                        attached_any = True

                if not attached_any and logs_dir is not None:
                    sched = Path(logs_dir) / "scheduler.log"
                    if sched.is_file():
                        self.send_document(
                            "scheduler_tail.log",
                            _tail_text(sched),
                            caption=f"{self.machine_name} — scheduler.log tail",
                        )
            return sent
        except Exception:
            return False
