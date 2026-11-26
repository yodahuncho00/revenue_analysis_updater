"""Authorize the configured Google user interactively without running the ETL."""

import argparse
import os
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
os.chdir(PROJECT_ROOT)
sys.path.insert(0, str(PROJECT_ROOT))

from integrations.google_sheets_sync import GoogleSheetsSync
from main import load_config


def main(force: bool = False, verify_folder: bool = True) -> int:
    config = load_config(str(PROJECT_ROOT / "config.yaml"))
    settings = config.get("google_sheets", {})
    if settings.get("auth_type") != "oauth_user":
        raise ValueError("Set google_sheets.auth_type to 'oauth_user' before authorizing.")
    sync = GoogleSheetsSync(config, allow_interactive=True, force_reauth=force)
    if verify_folder:
        for key, default_name in (
            ("gateway_spreadsheet_name", "Backoffice_Gateway"),
            ("pos_spreadsheet_name", "Backoffice_POS"),
        ):
            title = settings.get(key, default_name)
            sheet_id = sync._find_spreadsheet(title)
            print(f"Verified access to Google Sheet {title!r} (id={sheet_id}).")
    print("Google Sheets OAuth is authorized and the local refresh token is ready.")
    return 0


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--force",
        action="store_true",
        help="Force a fresh interactive OAuth consent flow and replace any existing token.",
    )
    parser.add_argument(
        "--skip-verify",
        action="store_true",
        help="Skip verifying access to the configured Drive spreadsheets after authorization.",
    )
    args = parser.parse_args()
    sys.exit(main(force=args.force, verify_folder=not args.skip_verify))
