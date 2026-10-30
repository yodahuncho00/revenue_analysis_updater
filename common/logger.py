import os
import sys
import uuid
import logging
from datetime import datetime
from typing import Optional, Tuple

class SecretFilter(logging.Filter):
    def __init__(self, secrets=None):
        super().__init__()
        self.secrets = [s for s in (secrets or []) if s]

    def filter(self, record):
        if isinstance(record.msg, str):
            for s in self.secrets:
                record.msg = record.msg.replace(s, "******")
        return True

def setup_logger(run_id: Optional[str] = None, log_dir: str = "data/logs", secrets=None) -> Tuple[logging.Logger, str]:
    if run_id is None:
        run_id = f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_{uuid.uuid4().hex[:6]}"

    os.makedirs(log_dir, exist_ok=True)
    log_file = os.path.join(log_dir, f"{run_id}.log")
    with open(os.path.join(log_dir, "latest_run.txt"), "w", encoding="utf-8") as run_file:
        run_file.write(f"{run_id}\n")

    logger = logging.getLogger(run_id)
    logger.setLevel(logging.DEBUG)
    logger.propagate = False

    # Clear existing handlers
    if logger.hasHandlers():
        logger.handlers.clear()

    # Secret filter
    sec_filter = SecretFilter(secrets)

    # Formatter
    fmt = logging.Formatter(
        fmt="%(asctime)s [%(levelname)s] [run_id=%(name)s] %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S"
    )

    # Console Handler (INFO+)
    ch = logging.StreamHandler(sys.stdout)
    ch.setLevel(logging.INFO)
    ch.setFormatter(fmt)
    ch.addFilter(sec_filter)
    logger.addHandler(ch)

    # File Handler (DEBUG+)
    fh = logging.FileHandler(log_file, encoding="utf-8")
    fh.setLevel(logging.DEBUG)
    fh.setFormatter(fmt)
    fh.addFilter(sec_filter)
    logger.addHandler(fh)

    logger.info(f"Initialized logger for run_id={run_id}, logging to {log_file}")
    return logger, run_id
