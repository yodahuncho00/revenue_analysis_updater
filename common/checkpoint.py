"""Per-date stage checkpointing stored outside OneDrive for fast, safe ETL resume."""

from datetime import date, datetime
import json
import os
import tempfile
import time
from typing import Any, Dict, List, Optional

from extractors.base import BalanceTableCapture


def _default_checkpoint_root() -> str:
    local_data = os.environ.get("LOCALAPPDATA") or tempfile.gettempdir()
    return os.path.join(local_data, "RevenueAnalysisUpdater", "Checkpoints")


def _atomic_write_json(path: str, payload: Any) -> None:
    os.makedirs(os.path.dirname(path), exist_ok=True)
    fd, tmp_path = tempfile.mkstemp(prefix=".ckpt_", suffix=".json", dir=os.path.dirname(path))
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as handle:
            json.dump(payload, handle, indent=2, ensure_ascii=False)
            handle.flush()
            os.fsync(handle.fileno())
        os.replace(tmp_path, path)
    finally:
        if os.path.exists(tmp_path):
            try:
                os.remove(tmp_path)
            except OSError:
                pass


def _encode_pos_value(value: Any) -> Any:
    if isinstance(value, datetime):
        return {"__dt__": value.isoformat()}
    if isinstance(value, date):
        return {"__d__": value.isoformat()}
    return value


def _decode_pos_value(value: Any) -> Any:
    if isinstance(value, dict):
        if "__dt__" in value and len(value) == 1:
            return datetime.fromisoformat(value["__dt__"])
        if "__d__" in value and len(value) == 1:
            return date.fromisoformat(value["__d__"])
    return value


class StageCheckpoint:
    """Persists per-stage outputs and completion status for a (source, business_date) run."""

    def __init__(
        self,
        source: str,
        business_date: str,
        base_dir: Optional[str] = None,
    ) -> None:
        self.source = source
        self.business_date = business_date
        root = base_dir or _default_checkpoint_root()
        safe_date = business_date.replace("/", "-").replace(":", "-")
        self.directory = os.path.join(root, f"{source}_{safe_date}")
        self.manifest_path = os.path.join(self.directory, "manifest.json")
        self._manifest: Dict[str, Any] = self._load_manifest()

    def _load_manifest(self) -> Dict[str, Any]:
        if os.path.isfile(self.manifest_path):
            try:
                with open(self.manifest_path, "r", encoding="utf-8") as handle:
                    data = json.load(handle)
                if isinstance(data, dict) and isinstance(data.get("stages"), dict):
                    return data
            except (OSError, ValueError):
                pass
        return {
            "source": self.source,
            "business_date": self.business_date,
            "stages": {},
        }

    def _save_manifest(self) -> None:
        self._manifest["updated_at"] = time.time()
        _atomic_write_json(self.manifest_path, self._manifest)

    def is_stage_done(self, stage: str) -> bool:
        entry = self._manifest.get("stages", {}).get(stage)
        return isinstance(entry, dict) and entry.get("status") == "done"

    def mark_stage(
        self,
        stage: str,
        status: str = "done",
        details: Optional[Dict[str, Any]] = None,
    ) -> None:
        self._manifest.setdefault("stages", {})[stage] = {
            "status": status,
            "timestamp": time.time(),
            "details": details or {},
        }
        self._save_manifest()

    def get_stage_details(self, stage: str) -> Dict[str, Any]:
        entry = self._manifest.get("stages", {}).get(stage)
        if isinstance(entry, dict) and isinstance(entry.get("details"), dict):
            return entry["details"]
        return {}

    def save_json(self, name: str, data: Any) -> str:
        path = os.path.join(self.directory, f"{name}.json")
        _atomic_write_json(path, data)
        return path

    def load_json(self, name: str) -> Optional[Any]:
        path = os.path.join(self.directory, f"{name}.json")
        if not os.path.isfile(path):
            return None
        try:
            with open(path, "r", encoding="utf-8") as handle:
                return json.load(handle)
        except (OSError, ValueError):
            return None

    def save_balance_capture(self, name: str, capture: BalanceTableCapture) -> str:
        payload = {
            "rows": capture.rows,
            "html": capture.html,
            "row_count": capture.row_count,
            "column_count": capture.column_count,
        }
        return self.save_json(name, payload)

    def load_balance_capture(self, name: str) -> Optional[BalanceTableCapture]:
        data = self.load_json(name)
        if not isinstance(data, dict):
            return None
        try:
            return BalanceTableCapture(
                rows=data["rows"],
                html=data["html"],
                row_count=int(data["row_count"]),
                column_count=int(data["column_count"]),
            )
        except (KeyError, TypeError, ValueError):
            return None

    def save_pos_rows(self, rows: List[List[Any]]) -> str:
        encoded = [[_encode_pos_value(val) for val in row] for row in rows]
        return self.save_json("pos_payments", encoded)

    def load_pos_rows(self) -> Optional[List[List[Any]]]:
        data = self.load_json("pos_payments")
        if not isinstance(data, list):
            return None
        return [[_decode_pos_value(val) for val in row] for row in data]

