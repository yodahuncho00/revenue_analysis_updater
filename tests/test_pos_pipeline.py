import csv
import io
from datetime import date, datetime, timedelta
from pathlib import Path
import tempfile
import pytest
import openpyxl

from extractors.playwright_extractor import PlaywrightExtractor, PlaywrightExtractionError
from scripts.update_pos_dashboard import update_pos_dashboard, _transaction_columns
from pos_pipeline import run_pos_pipeline, POS_HEADERS


def test_pos_headers_match_extractor_columns():
    assert list(POS_HEADERS) == list(PlaywrightExtractor.POS_PAYMENT_COLUMNS)
    assert len(POS_HEADERS) == 18


def test_pos_payments_csv_parsing_and_sorting():
    csv_data = (
        "Payment ID,Datetime,Type,Amount,Currency,Merchant fee,Provider cost,Partner net revenue,"
        "Account,Receipt Name,Agent,Terminal,Parent agent,Provider,Status,Customer fee,Customer fee mode,Agent rebate\n"
        "#490590002,2026-09-02 10:00:00,Receive,50,ZMW,1,0.5,0.5,260111,Agent1,Ag1,T1,,Airtel,Success: ,0,,0\n"
        "#490580005,2026-08-15 12:00:00,Receive,100,ZMW,2,1.0,1.0,260222,Agent2,Ag2,T2,,MTN,Success: ,0,,0\n"
        "#490580001,2026-08-10 09:00:00,Receive,10,ZMW,,0,,260333,Agent3,Ag3,T3,,Airtel,Rejected: ,0,,\n"
        "#490590001,2026-09-01 08:00:00,Receive,20,ZMW,0.5,0.2,0.3,260444,Agent4,Ag4,T4,,MTN,Success: ,0,,0\n"
    )

    extractor = PlaywrightExtractor({"backoffice": {}, "pos_portal": {}, "paths": {}})
    reader = csv.DictReader(io.StringIO(csv_data, newline=""))
    parsed = []
    for line_number, record in enumerate(reader, start=2):
        values = [record[column] for column in extractor.POS_PAYMENT_COLUMNS]
        created = datetime.strptime(values[1].strip(), "%Y-%m-%d %H:%M:%S")
        payment_id = int(values[0].strip().lstrip("#"))
        typed = []
        for index, value in enumerate(values):
            raw_value = value
            value = value.strip() if isinstance(value, str) else value
            if not value or value in ("-", "n/a", "N/A", "null", "None"):
                typed.append(None)
            elif index == 0:
                typed.append(payment_id)
            elif index == 1:
                typed.append(created)
            elif index == 8 and isinstance(value, str) and value.isdigit():
                typed.append(int(value))
            elif index in (3, 5, 6, 7, 15, 17):
                num = float(str(value).replace(",", ""))
                typed.append(int(num) if num.is_integer() else num)
            else:
                typed.append(raw_value)
        parsed.append((created.year, created.month, payment_id, typed))

    parsed.sort(key=lambda item: (item[0], item[1], item[2]))
    rows = [item[3] for item in parsed]

    assert len(rows) == 4
    # Aug payment 490580001 first
    assert rows[0][0] == 490580001
    assert rows[0][1] == datetime(2026, 8, 10, 9, 0)
    assert rows[0][5] is None  # empty fee

    # Aug payment 490580005 second
    assert rows[1][0] == 490580005
    assert rows[1][1] == datetime(2026, 8, 15, 12, 0)

    # Sep payment 490590001 third
    assert rows[2][0] == 490590001
    assert rows[2][1] == datetime(2026, 9, 1, 8, 0)

    # Sep payment 490590002 fourth
    assert rows[3][0] == 490590002
    assert rows[3][1] == datetime(2026, 9, 2, 10, 0)


def test_pos_dashboard_handles_float_datetime(tmp_path):
    workbook_path = tmp_path / "test_pos.xlsx"
    template_path = Path("templates/05. Revenue_dashboard_POS.template.html")
    dashboard_path = tmp_path / "test_dashboard.html"

    wb = openpyxl.Workbook()
    ws_bo = wb.active
    ws_bo.title = "BO_extract"
    ws_bo.append([""] * 18)
    ws_bo.append([""] * 18)
    ws_bo.append(list(POS_HEADERS) + [""] * 40)

    # Add row with float Excel serial datetime (46334.69996527778 is 2026-11-08 16:47:57)
    row_data = [490590957, 46334.69996527778, "Receive", 100, "ZMW", 2, 1, 1, 260999, "Agent", "A", "T", None, "Airtel", "Success: ", 0, None, 0]
    # pad up to column BB (col 54)
    padded = row_data + [None] * (54 - len(row_data))
    padded[29] = 1  # Success (AD)
    padded[30] = "Success"  # AE
    padded[34] = "ZMW"  # AI
    padded[38] = "QR"  # AM
    padded[39] = "Airtel"  # AN
    padded[43] = "Parent"  # AR
    padded[44] = "Merchant"  # AS
    padded[45] = "Terminal"  # AT
    padded[53] = "CustKey"  # BB
    ws_bo.append(padded)

    ws_mm = wb.create_sheet("Merchant_master")
    ws_mm.append([""] * 10)
    ws_mm.append([""] * 10)
    ws_mm.append([""] * 10)
    ws_mm.append([""] * 10)
    ws_mm.append(["SysName", "Merchant", "", "", "Retail", "", 0])

    ws_pm = wb.create_sheet("Parent_master")
    ws_pm.append([""] * 10)
    ws_pm.append([""] * 10)
    ws_pm.append([""] * 10)
    ws_pm.append([""] * 10)
    ws_pm.append(["SysName", "Parent", "", "Independent"])

    ws_ctrl = wb.create_sheet("Control")
    ws_ctrl["B5"] = datetime(2026, 11, 8)
    ws_ctrl["B6"] = datetime(2026, 11, 8)
    ws_ctrl["B20"] = "ZMW"

    ws_lists = wb.create_sheet("Lists")
    ws_lists.cell(5, 1).value = "Retail"
    ws_lists.cell(5, 18).value = 0

    ws_fx = wb.create_sheet("FX")
    ws_fx.cell(5, 5).value = "USD"
    ws_fx.cell(5, 6).value = "ZMW"
    ws_fx.cell(5, 16).value = "USD"
    ws_fx.cell(5, 17).value = "ZMW"
    ws_fx.append(["", "", 2026, 11, 1.0, 20.0] + [None] * 10 + [1.0, 20.0])

    ws_ch = wb.create_sheet("Channel_master")
    ws_ch.append([""] * 12)
    ws_ch.append([""] * 12)
    ws_ch.append([""] * 12)
    ws_ch.append([""] * 12)
    ws_ch.append(["Airtel QR", "", "", "Airtel", "QR", "C2B", "", "", 0.02, 0, 0.5, 0])

    wb.save(workbook_path)
    wb.close()

    tx_count = update_pos_dashboard(workbook_path, template_path, dashboard_path)
    assert tx_count == 1
    assert dashboard_path.exists()
    content = dashboard_path.read_text(encoding="utf-8")
    assert "490590957" not in content  # Payment ID is indexed/omitted in tx, but timestamp should be there
    assert "2026-11-08T16:47" in content


def test_pos_pipeline_row_replacement_preserves_formulas(tmp_path, monkeypatch):
    workbook_path = tmp_path / "05. Revenue_analysis_POS.xlsx"
    dashboard_path = tmp_path / "05. Revenue_dashboard_POS.html"
    template_path = Path("templates/05. Revenue_dashboard_POS.template.html")

    wb = openpyxl.Workbook()
    ws_bo = wb.active
    ws_bo.title = "BO_extract"
    ws_bo.append(["Title"])
    ws_bo.append(["Instructions"])
    ws_bo.append(list(POS_HEADERS) + [""] * 8 + ["FormulaColAA", "FormulaColAB"])

    # Add old rows (rows 4, 5, 6) with formula in col AA (col 27)
    # 18 data cols + 8 empty cols (S..Z) = 26 cols before col 27 (AA)
    ws_bo.append([111, datetime(2026, 8, 1), "Receive", 10, "ZMW"] + [None] * 13 + [None] * 8 + ["=YEAR(B4)", "=MONTH(B4)"])
    ws_bo.append([222, datetime(2026, 8, 2), "Receive", 20, "ZMW"] + [None] * 13 + [None] * 8 + ["=YEAR(B5)", "=MONTH(B5)"])
    ws_bo.append([333, datetime(2026, 8, 3), "Receive", 30, "ZMW"] + [None] * 13 + [None] * 8 + ["=YEAR(B6)", "=MONTH(B6)"])

    ws_ctrl = wb.create_sheet("Control")
    ws_ctrl.cell(6, 1).value = "Data as-of date (last day in the export)"
    ws_ctrl.cell(6, 2).value = datetime(2026, 8, 3)

    wb.save(workbook_path)
    wb.close()

    new_rows = [
        [490580001, datetime(2026, 8, 10, 15, 0), "Receive", 50, "ZMW", 1, 0.5, 0.5, 260111, "Agent", "A", "T", None, "Airtel", "Success: ", 0, None, 0],
    ]

    # Monkeypatch PlaywrightExtractor.extract_pos_payments
    monkeypatch.setattr(PlaywrightExtractor, "extract_pos_payments", lambda self: new_rows)
    # Monkeypatch ExcelCalculator.recalculate
    from writers.calculator import ExcelCalculator
    monkeypatch.setattr(ExcelCalculator, "recalculate", lambda path: None)

    cfg = {
        "paths": {
            "pos_workbook_path": str(workbook_path),
            "pos_dashboard_path": str(dashboard_path),
            "pos_dashboard_template_path": str(template_path),
            "backup_dir": str(tmp_path / "backups"),
            "raw_dir": str(tmp_path / "raw"),
        },
        "app": {"backups_to_retain": 5},
    }

    result = run_pos_pipeline(cfg, dry_run=False)
    assert result["success"] is True
    assert result["rows"] == 1

    # Verify workbook contents
    wb_updated = openpyxl.load_workbook(workbook_path, data_only=False)
    ws = wb_updated["BO_extract"]
    # Row 4 is new data
    assert ws.cell(4, 1).value == 490580001
    assert ws.cell(4, 2).value == datetime(2026, 8, 10, 15, 0)
    assert ws.cell(4, 2).number_format == "DD.MM.YYYY"
    # Formula column AA (col 27) was preserved!
    assert ws.cell(4, 27).value == "=YEAR(B4)"

    # Rows 5 and 6: data cols 1-18 should be cleared!
    assert ws.cell(5, 1).value is None
    assert ws.cell(6, 1).value is None
    # Formula column AA on rows 5 and 6 should still be there!
    assert ws.cell(5, 27).value == "=YEAR(B5)"
    assert ws.cell(6, 27).value == "=YEAR(B6)"

    # Control!B6 always reflects yesterday, independent of transaction dates.
    ctrl = wb_updated["Control"]
    assert ctrl["B6"].value == datetime.combine(
        date.today() - timedelta(days=1), datetime.min.time()
    )
    assert ctrl["B6"].number_format == "DD.MM.YYYY"
    wb_updated.close()


def test_pos_ready_helper():
    from scripts.refresh_backoffice_session import _pos_ready

    class MockLocator:
        def __init__(self, visible=True):
            self._visible = visible

        def is_visible(self):
            return self._visible

        @property
        def first(self):
            return self

        def filter(self, **kwargs):
            return self

    class MockPage:
        def __init__(self, url="https://pos-portal.unipesa.tech/payments", closed=False, save_visible=True):
            self.url = url
            self._closed = closed
            self._save_visible = save_visible

        def is_closed(self):
            return self._closed

        def get_by_role(self, role, name=None, exact=False):
            return MockLocator(visible=self._save_visible)

        def locator(self, selector):
            return MockLocator(visible=self._save_visible)

        def goto(self, url, **kwargs):
            self.url = url

    # Test 1: When on payments page and Save CSV is visible -> True
    p1 = MockPage("https://pos-portal.unipesa.tech/payments", save_visible=True)
    assert _pos_ready(None, p1) is True

    # Test 2: When on login page and Save CSV is not visible -> False
    p2 = MockPage("https://pos-portal.unipesa.tech/login", save_visible=False)
    assert _pos_ready(None, p2) is False

    # Test 3: When page is closed -> False
    p3 = MockPage(closed=True)
    assert _pos_ready(None, p3) is False

    # Test 4: When on another domain -> False
    p4 = MockPage("https://example.com/payments")
    assert _pos_ready(None, p4) is False
