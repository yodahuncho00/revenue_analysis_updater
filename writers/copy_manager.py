import os
import shutil
import hashlib
import sys
import ctypes
import re
import tempfile
from datetime import datetime
from typing import Optional, Tuple

class CopyManagerError(Exception):
    pass

class CopyManager:
    """Manages run lock, working copy isolation, timestamped backups, and atomic commit."""

    def __init__(
        self,
        prod_path: str,
        backup_dir: str = "data/backups",
        backups_to_retain: int = 30,
        run_id: Optional[str] = None,
    ):
        self.prod_path = os.path.abspath(prod_path)
        self.backup_dir = os.path.abspath(backup_dir)
        self.backups_to_retain = backups_to_retain
        self.lock_file = f"{self.prod_path}.lock"
        local_data = os.environ.get("LOCALAPPDATA") or tempfile.gettempdir()
        run_suffix = run_id or f"pid_{os.getpid()}"
        self.working_copy_dir = os.path.join(
            local_data, "RevenueAnalysisUpdater", "WorkingCopies"
        )
        self.working_copy_path = os.path.join(
            self.working_copy_dir,
            f"{os.path.splitext(os.path.basename(self.prod_path))[0]}"
            f"_{run_suffix}.working.xlsx",
        )
        self.legacy_working_copy_path = f"{self.prod_path}.working.xlsx"
        self.is_locked = False
        self._mutex_handle = None
        self._lock_fd = None
        os.makedirs(self.backup_dir, exist_ok=True)

    @staticmethod
    def compute_sha256(filepath: str) -> str:
        h = hashlib.sha256()
        with open(filepath, "rb") as f:
            while chunk := f.read(65536):
                h.update(chunk)
        return h.hexdigest()

    def acquire_lock(self) -> None:
        if sys.platform == "win32":
            self._acquire_windows_mutex()
        else:
            self._acquire_posix_file_lock()
            return

        try:
            # The OS lock is authoritative. A leftover marker is safe to remove
            # only after acquiring it, which also recovers after a hard stop.
            if os.path.exists(self.lock_file):
                old_pid = self._read_lock_pid()
                if old_pid is None or self._is_pid_running(old_pid):
                    raise CopyManagerError(
                        f"A previous ETL process may still be running (lock: {self.lock_file})."
                    )
                os.remove(self.lock_file)

            with open(self.lock_file, "x") as f:
                f.write(f"pid={os.getpid()}, timestamp={datetime.now().isoformat()}\n")
            self.is_locked = True
        except Exception:
            self._release_os_lock()
            raise

    def _acquire_windows_mutex(self) -> None:
        """Acquire a kernel mutex that Windows releases automatically on exit."""
        digest = hashlib.sha256(os.path.normcase(self.prod_path).encode("utf-8")).hexdigest()[:24]
        name = f"Local\\RevenueAnalysisUpdater_{digest}"
        kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
        kernel32.CreateMutexW.argtypes = (ctypes.c_void_p, ctypes.c_int, ctypes.c_wchar_p)
        kernel32.CreateMutexW.restype = ctypes.c_void_p
        kernel32.WaitForSingleObject.argtypes = (ctypes.c_void_p, ctypes.c_ulong)
        kernel32.WaitForSingleObject.restype = ctypes.c_ulong
        kernel32.CloseHandle.argtypes = (ctypes.c_void_p,)
        handle = kernel32.CreateMutexW(None, False, name)
        if not handle:
            raise CopyManagerError(f"Could not create ETL run mutex (Windows error {ctypes.get_last_error()}).")

        wait_result = kernel32.WaitForSingleObject(handle, 0)
        if wait_result == 0x00000102:  # WAIT_TIMEOUT: another run owns the mutex
            kernel32.CloseHandle(handle)
            raise CopyManagerError("Another ETL run is already active for this workbook.")
        if wait_result not in (0x00000000, 0x00000080):  # WAIT_OBJECT_0 / WAIT_ABANDONED
            error = ctypes.get_last_error()
            kernel32.CloseHandle(handle)
            raise CopyManagerError(f"Could not acquire ETL run mutex (Windows error {error}).")
        self._mutex_handle = handle

    def _acquire_posix_file_lock(self) -> None:
        import fcntl

        self._lock_fd = open(self.lock_file, "a+")
        try:
            fcntl.flock(self._lock_fd.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
        except BlockingIOError as exc:
            self._lock_fd.close()
            self._lock_fd = None
            raise CopyManagerError("Another ETL run is already active for this workbook.") from exc
        self._lock_fd.seek(0)
        self._lock_fd.truncate()
        self._lock_fd.write(f"pid={os.getpid()}, timestamp={datetime.now().isoformat()}\n")
        self._lock_fd.flush()
        self.is_locked = True

    def _read_lock_pid(self) -> Optional[int]:
        try:
            with open(self.lock_file, "r", encoding="utf-8") as lock:
                match = re.search(r"\bpid=(\d+)\b", lock.read())
            return int(match.group(1)) if match else None
        except OSError:
            return None

    @staticmethod
    def _is_pid_running(pid: int) -> bool:
        if sys.platform == "win32":
            kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
            kernel32.OpenProcess.argtypes = (ctypes.c_ulong, ctypes.c_int, ctypes.c_ulong)
            kernel32.OpenProcess.restype = ctypes.c_void_p
            kernel32.WaitForSingleObject.argtypes = (ctypes.c_void_p, ctypes.c_ulong)
            kernel32.WaitForSingleObject.restype = ctypes.c_ulong
            kernel32.CloseHandle.argtypes = (ctypes.c_void_p,)
            handle = kernel32.OpenProcess(0x00100000, False, pid)  # SYNCHRONIZE
            if not handle:
                return ctypes.get_last_error() != 87  # ERROR_INVALID_PARAMETER means no such PID
            try:
                return kernel32.WaitForSingleObject(handle, 0) == 0x00000102  # WAIT_TIMEOUT
            finally:
                kernel32.CloseHandle(handle)

        try:
            os.kill(pid, 0)
            return True
        except ProcessLookupError:
            return False
        except PermissionError:
            return True

    def _release_os_lock(self) -> None:
        if self._mutex_handle:
            kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
            kernel32.ReleaseMutex.argtypes = (ctypes.c_void_p,)
            kernel32.CloseHandle.argtypes = (ctypes.c_void_p,)
            kernel32.ReleaseMutex(self._mutex_handle)
            kernel32.CloseHandle(self._mutex_handle)
            self._mutex_handle = None
        if self._lock_fd:
            import fcntl

            fcntl.flock(self._lock_fd.fileno(), fcntl.LOCK_UN)
            self._lock_fd.close()
            self._lock_fd = None

    def release_lock(self) -> None:
        if self.is_locked:
            try:
                if sys.platform == "win32" and os.path.exists(self.lock_file):
                    os.remove(self.lock_file)
            finally:
                self._release_os_lock()
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
        """Copies production to a per-run working copy outside synced folders."""
        os.makedirs(self.working_copy_dir, exist_ok=True)
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

    def cleanup_working_copy(self) -> list[str]:
        """Delete this run's and any legacy project-folder working copies."""
        deleted_paths = []
        paths = dict.fromkeys(
            (self.working_copy_path, self.legacy_working_copy_path)
        )
        failures = []
        for path in paths:
            if not os.path.exists(path):
                continue
            last_error = None
            for attempt in range(8):
                try:
                    os.remove(path)
                    deleted_paths.append(path)
                    break
                except FileNotFoundError:
                    break
                except OSError as exc:
                    last_error = exc
                    if attempt < 7:
                        import time

                        time.sleep(0.5)
            else:
                failures.append((path, last_error))
        if failures:
            details = "; ".join(f"{path} ({error})" for path, error in failures)
            raise CopyManagerError(
                f"Could not delete successful-run working workbook(s): {details}"
            ) from failures[0][1]
        return deleted_paths

    def cleanup_on_failure(self) -> None:
        """On failure, leaves the working copy intact for debugging, but releases the run lock."""
        self.release_lock()
