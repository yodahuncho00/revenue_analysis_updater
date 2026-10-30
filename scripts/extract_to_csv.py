"""Run the live ETL in dry-run mode and export its exact raw captures as CSVs."""

import argparse
import os
import sys

PROJECT_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
sys.path.insert(0, PROJECT_ROOT)
os.chdir(PROJECT_ROOT)

from main import ALLOWED_STEPS, run_pipeline


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Validate live Backoffice data through the ETL and export its raw captures."
    )
    parser.add_argument("--config", default="config.yaml", help="Path to config.yaml")
    parser.add_argument("--date", help="Override effective date (YYYY-MM-DD)")
    parser.add_argument(
        "--step",
        choices=ALLOWED_STEPS,
        help="Run one ETL step",
    )
    args = parser.parse_args()

    return run_pipeline(
        config_path=args.config,
        dry_run=True,
        override_date=args.date,
        step=args.step,
        extractor_type="playwright",
    )


if __name__ == "__main__":
    sys.exit(main())
