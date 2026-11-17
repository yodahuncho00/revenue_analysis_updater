from datetime import datetime

import certifi
import openpyxl
from openpyxl.utils.datetime import CALENDAR_WINDOWS_1900, to_excel

from integrations import google_sheets_sync
from integrations.google_sheets_sync import GoogleSheetsSync


class _Request:
    def __init__(self, response=None):
        self.response = response or {}

    def execute(self):
        return self.response


class _Drive:
    def __init__(self):
        self.query = None

    def files(self):
        return self

    def list(self, **kwargs):
        self.query = kwargs
        return _Request({"files": [{"id": "spreadsheet-id", "name": "Backoffice_POS"}]})


class _Sheets:
    def __init__(self, sheet_title="BO_extract"):
        self.sheet_title = sheet_title
        self.cleared = []
        self.updated = []
        self.batch_updates = []

    def spreadsheets(self):
        return self

    def get(self, **kwargs):
        return _Request(
            {
                "sheets": [
                    {
                        "properties": {
                            "sheetId": 12,
                            "title": self.sheet_title,
                            "gridProperties": {"rowCount": 10, "columnCount": 26},
                        }
                    }
                ]
            }
        )

    def batchUpdate(self, **kwargs):
        self.batch_updates.append(kwargs)
        return _Request()

    def values(self):
        return self

    def clear(self, **kwargs):
        self.cleared.append(kwargs)
        return _Request()

    def update(self, **kwargs):
        self.updated.append(kwargs)
        return _Request()


def _sync(tmp_path, drive, sheets):
    key_path = tmp_path / "service-account.json"
    key_path.write_text("test-only", encoding="utf-8")
    return GoogleSheetsSync(
        {
            "_project_root": str(tmp_path),
            "google_sheets": {
                "folder_id": "drive-folder",
                "service_account_file": str(key_path),
            },
        },
        drive_service=drive,
        sheets_service=sheets,
    )


def test_pos_sync_replaces_a3_r_and_copies_cell_format(tmp_path):
    workbook_path = tmp_path / "pos.xlsx"
    workbook = openpyxl.Workbook()
    worksheet = workbook.active
    worksheet.title = "BO_extract"
    worksheet.append(["Title"])
    worksheet.append(["Instructions"])
    worksheet.append(["Payment ID", "Datetime"] + [""] * 16)
    worksheet["A3"].font = openpyxl.styles.Font(bold=True)
    worksheet.column_dimensions["A"].width = 22
    worksheet.row_dimensions[3].height = 24
    worksheet["B4"] = datetime(2026, 11, 12, 9, 30)
    worksheet["B4"].number_format = "DD.MM.YYYY"
    workbook.save(workbook_path)
    workbook.close()

    drive = _Drive()
    sheets = _Sheets()
    sync = _sync(tmp_path, drive, sheets)
    rows = [[123, datetime(2026, 11, 12, 9, 30)] + [None] * 16]
    sync.sync_pos_payments(
        str(workbook_path),
        "Backoffice_POS",
        ["Payment ID", "Datetime"] + [""] * 16,
        rows,
    )

    assert drive.query["q"].find("'drive-folder' in parents") >= 0
    assert sheets.cleared[0]["range"] == "'BO_extract'!A3:R10"
    assert sheets.updated[0]["range"] == "'BO_extract'!A3"
    assert sheets.updated[0]["body"]["values"][0][0:2] == ["Payment ID", "Datetime"]
    assert sheets.updated[0]["body"]["values"][1][0] == 123
    assert sheets.updated[0]["body"]["values"][1][1] == to_excel(
        datetime(2026, 11, 12, 9, 30), CALENDAR_WINDOWS_1900
    )
    assert any(
        request.get("repeatCell", {})
        .get("cell", {})
        .get("userEnteredFormat", {})
        .get("numberFormat", {})
        .get("pattern")
        == "DD.MM.YYYY"
        for batch in sheets.batch_updates
        for request in batch["body"]["requests"]
    )
    assert any(
        request.get("updateDimensionProperties", {})
        .get("range", {})
        .get("dimension")
        == "COLUMNS"
        for batch in sheets.batch_updates
        for request in batch["body"]["requests"]
    )


def test_sync_fails_before_writing_if_google_tab_is_missing(tmp_path):
    workbook_path = tmp_path / "pos.xlsx"
    workbook = openpyxl.Workbook()
    workbook.active.title = "BO_extract"
    workbook.save(workbook_path)
    workbook.close()

    sheets = _Sheets(sheet_title="DifferentTab")
    sync = _sync(tmp_path, _Drive(), sheets)

    try:
        sync.sync_pos_payments(
            str(workbook_path), "Backoffice_POS", ["Payment ID"], [[123]]
        )
    except ValueError as exc:
        assert "has no tab 'BO_extract'" in str(exc)
    else:
        raise AssertionError("Expected missing Google tab to stop the sync.")

    assert sheets.cleared == []
    assert sheets.updated == []


def test_gateway_sync_validates_all_tabs_before_any_remote_write(tmp_path):
    workbook_path = tmp_path / "gateway.xlsx"
    workbook = openpyxl.Workbook()
    workbook.active.title = "BO_extract"
    workbook.create_sheet("BO_provider_balance")
    workbook.save(workbook_path)
    workbook.close()

    sheets = _Sheets(sheet_title="BO_extract")
    sync = _sync(tmp_path, _Drive(), sheets)

    try:
        sync.sync_workbook_tabs(
            str(workbook_path),
            "Backoffice_Gateway",
            {
                "BO_extract": "BO_extract",
                "BO_provider_balance": "BO_provider_balance",
            },
        )
    except ValueError as exc:
        assert "missing tab(s): BO_provider_balance" in str(exc)
    else:
        raise AssertionError("Expected missing Google tab to stop the sync.")

    assert sheets.cleared == []
    assert sheets.updated == []


def test_gateway_sync_maps_local_and_remote_tabs_by_config(tmp_path):
    workbook_path = tmp_path / "gateway.xlsx"
    workbook = openpyxl.Workbook()
    worksheet = workbook.active
    worksheet.title = "BO_merchant_balance"
    worksheet["A1"] = "merchant data"
    workbook.save(workbook_path)
    workbook.close()

    sheets = _Sheets(sheet_title="BO_merchants_balance")
    sync = _sync(tmp_path, _Drive(), sheets)
    sync.sync_workbook_tabs(
        str(workbook_path),
        "Backoffice_Gateway",
        {"BO_merchant_balance": "BO_merchants_balance"},
    )

    assert sheets.cleared[0]["range"].startswith("'BO_merchants_balance'!")
    assert sheets.updated[0]["range"] == "'BO_merchants_balance'!A1"
    assert sheets.updated[0]["body"]["values"][0][0] == "merchant data"


def test_default_google_ca_bundle_combines_certifi_and_windows_roots(
    tmp_path, monkeypatch
):
    certifi_path = tmp_path / "certifi.pem"
    certifi_path.write_bytes(b"certifi-root\n")
    monkeypatch.setattr(certifi, "where", lambda: str(certifi_path))
    monkeypatch.setattr(
        google_sheets_sync.ssl,
        "enum_certificates",
        lambda store: [
            (b"trusted", "x509_asn", {"1.3.6.1.5.5.7.3.1"}),
            (b"other-usage", "x509_asn", {"1.2.3.4"}),
            (b"not-x509", "pkcs_7_asn", True),
        ],
        raising=False,
    )
    monkeypatch.setattr(
        google_sheets_sync.ssl,
        "DER_cert_to_PEM_cert",
        lambda certificate: "PEM:{}".format(certificate.decode()),
    )
    monkeypatch.setenv("LOCALAPPDATA", str(tmp_path))

    bundle_path = google_sheets_sync._google_ca_bundle({}, str(tmp_path))
    with open(bundle_path, "rb") as bundle_file:
        bundle = bundle_file.read().decode("ascii")

    assert "certifi-root" in bundle
    assert "PEM:trusted" in bundle
    assert "other-usage" not in bundle
    assert "not-x509" not in bundle


def test_sheets_write_retries_rate_limit_once(monkeypatch):
    import httplib2
    from googleapiclient.errors import HttpError

    class _RetryOnce:
        def __init__(self):
            self.calls = 0

        def execute(self):
            self.calls += 1
            if self.calls == 1:
                raise HttpError(
                    httplib2.Response({"status": 429, "retry-after": "2"}),
                    b"rate limited",
                )
            return {"success": True}

    sleeps = []
    request = _RetryOnce()
    monkeypatch.setattr(google_sheets_sync.time_module, "sleep", sleeps.append)

    assert GoogleSheetsSync._execute_write(request) == {"success": True}
    assert request.calls == 2
    assert sleeps == [2.0]
