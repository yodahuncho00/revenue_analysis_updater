import os
import shutil
import hashlib
from datetime import datetime
from typing import Optional, Tuple

class CopyManagerError(Exception):
    pass

class CopyManager:
    """Manages run lock, working copy isolation, timestamped backups, and atomic commit."""

    def __init__(self, prod_path: str, backup_dir: str = "data/backups", backups_to_retain: int = 30):
        self.prod_path = os.path.abspath(prod_path)
        self.backup_dir = os.path.abspath(backup_dir)
        self.backups_to_retain = backups_to_retain
        self.lock_file = f"{self.prod_path}.lock"
        self.working_copy_path = f"{self.prod_path}.working.xlsx"
        self.is_locked = False
        os.makedirs(self.backup_dir, exist_ok=True)

    @staticmethod
    def compute_sha256(filepath: str) -> str:
        h = hashlib.sha256()
        with open(filepath, "rb") as f:
            while chunk := f.read(65536):
                h.update(chunk)
        return h.hexdigest()

    def acquire_lock(self) -> None:
        if os.path.exists(self.lock_file):
            raise CopyManagerError(
                f"Another ETL run is in progress or lock file was abandoned: {self.lock_file}. "
                "Confirm no other process is running before removing the lock."
            )
        try:
            with open(self.lock_file, "x") as f:
                f.write(f"pid={os.getpid()}, timestamp={datetime.now().isoformat()}\n")
            self.is_locked = True
        except Exception as e:
            raise CopyManagerError(f"Failed to acquire single-run lock: {e}")

    def release_lock(self) -> None:
        if self.is_locked and os.path.exists(self.lock_file):
            try:
                os.remove(self.lock_file)
            except Exception:
                pass
            self.is_locked = False

    def create_backup(self) -> Tuple[str, str]:
        """Creates a timestamped backup before any change. Returns (backup_path, sha256)."""
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        base_name = os.path.splitext(os.path.basename(self.prod_path))[0]
        backup_name = f"{base_name}_{ts}.xlsx"
        backup_path = os.path.join(self.backup_dir, backup_name)

        shutil.copy2(self.prod_path, backup_path)
        sha256_hash = self.compute_sha256(backup_path)

        # Cleanup old backups
        self._prune_old_backups()
        return backup_path, sha256_hash

    def _prune_old_backups(self) -> None:
        try:
            files = [
                os.path.join(self.backup_dir, f)
                for f in os.listdir(self.backup_dir)
                if f.endswith(".xlsx")
            ]
            files.sort(key=lambda x: os.path.getmtime(x), reverse=True)
            for old_file in files[self.backups_to_retain:]:
                try:
                    os.remove(old_file)
                except Exception:
                    pass
        except Exception:
            pass

    def create_working_copy(self) -> str:
        """Copies the production workbook to a clean isolated working copy."""
        if os.path.exists(self.working_copy_path):
            try:
                os.remove(self.working_copy_path)
            except Exception:
                pass
        shutil.copy2(self.prod_path, self.working_copy_path)
        return self.working_copy_path

    def commit(self) -> None:
        """
        Atomically commits working copy to production:
        1. Verifies working copy exists and is readable.
        2. Renames working copy to replace production workbook via os.replace (atomic on Windows/POSIX).
        """
        if not os.path.exists(self.working_copy_path):
            raise CopyManagerError(f"Working copy not found for commit: {self.working_copy_path}")

        # Flush buffers
        with open(self.working_copy_path, "r+b") as f:
            f.flush()
            os.fsync(f.fileno())

        temp_target = f"{self.prod_path}.tmp"
        shutil.copy2(self.working_copy_path, temp_target)
        
        # Atomic replace
        os.replace(temp_target, self.prod_path)

        # Clean working copy
        if os.path.exists(self.working_copy_path):
            try:
                os.remove(self.working_copy_path)
            except Exception:
                pass

    def cleanup_on_failure(self) -> None:
        """On failure, leaves the working copy intact for debugging, but releases the run lock."""
        self.release_lock()
