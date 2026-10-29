"""Refresh the embedded data in the standalone revenue dashboard from Excel."""

import argparse
import json
import os
import tempfile
from datetime import date, datetime
from pathlib import Path

import openpyxl


PROJECT_ROOT = Path(__file__).resolve().parents[1]


def _json_value(value):
    if isinstance(value, (datetime, date)):
        return value.date().isoformat() if isinstance(value, datetime) else value.isoformat()
    return value


def _rows(sheet, start_row, key_column=1):
    for values in sheet.iter_rows(min_row=start_row, values_only=True):
        if key_column <= len(values) and values[key_column - 1] is not None:
            yield values


def _build_data(workbook):
    data_sheet = workbook["Data"]
    facts = []
    for row in _rows(data_sheet, 5):
        if not row[0]:
            continue
        facts.append({
            "period": _json_value(row[0]), "type": row[1], "year": row[23], "month": row[24],
            "country": row[2], "ccy": row[3], "merchant": row[25], "provider": row[26],
            "channel": row[27], "method": row[28], "merchant_type": row[29], "test": row[30],
            "fx": row[32], "turnover": row[33], "deposits": row[34], "withdrawals": row[35],
            "dep_fees": row[36], "wd_fees": row[37], "fees": row[38], "dep_cost": row[39],
            "wd_cost": row[40], "cost": row[41], "revenue": row[44], "settlement": row[47],
            "topup": row[48], "vintage": row[53], "channel_type": row[59], "risk": row[60],
        })

    control_sheet = workbook["Control"]
    control = {
        "rpt_month": _json_value(control_sheet["B5"].value),
        "as_of": _json_value(control_sheet["B6"].value),
        "cy": control_sheet["B7"].value,
        "cm": control_sheet["B8"].value,
        "pmy": control_sheet["B10"].value,
        "pmm": control_sheet["B11"].value,
        "py": control_sheet["B12"].value,
        "days_in": control_sheet["B13"].value,
        "days_el": control_sheet["B14"].value,
        "days_pm": control_sheet["B16"].value,
        "months_el": control_sheet["B17"].value,
        "zoyk_share": control_sheet["B18"].value,
        "include_test": control_sheet["B20"].value,
    }

    merchant_sheet = workbook["Merchant_master"]
    merchant_master = [{
        "merchant": row[1], "system_name": row[0], "merchant_type": row[3],
        "registered": _json_value(row[4]), "test": row[5], "country": row[6],
        "currency": row[7], "risk": row[8], "first_period": _json_value(row[9]),
        "vintage": row[11], "status": row[14],
    } for row in _rows(merchant_sheet, 5) if row[1]]

    provider_master = [{
        "provider": row[0], "provider_type": row[1], "country": row[2], "status": row[6],
    } for row in _rows(workbook["Provider_master"], 5)]

    channel_master = [{
        "channel": row[0], "country": row[1], "currency": row[2], "provider": row[3],
        "method": row[4], "channel_type": row[5], "high_risk_allowed": row[6],
        "cost_c2b": row[13], "cost_b2c": row[14], "status": row[17],
    } for row in _rows(workbook["Channel_master"], 5)]

    def recon_rows(sheet_name, third_key):
        result = []
        for row in _rows(workbook[sheet_name], 5):
            result.append({
                "country": row[0], "currency": row[1], third_key: row[2],
                "dep": row[3], "dfee": row[4], "wd": row[5], "wfee": row[6],
                "settle": row[7], "topup": row[8], "calc": row[9], "cumonly": row[11],
                "delta": row[12], "bal": row[13], "bal2": row[14], "bal3": row[15],
                "var": row[16], "status": row[17],
            })
        return result

    offer_card = [list(row[:6]) for row in _rows(workbook["Offer_card"], 5)]

    return {
        "facts": facts,
        "control": [control],
        "merchant_master": merchant_master,
        "provider_master": provider_master,
        "channel_master": channel_master,
        "merchant_recon": recon_rows("Merchant_Balance", "merchant"),
        "provider_recon": recon_rows("Provider_Balance", "provider"),
        "offer_card": offer_card,
    }


def update_dashboard(workbook_path, template_path, dashboard_path):
    workbook = openpyxl.load_workbook(workbook_path, read_only=True, data_only=True)
    try:
        data = _build_data(workbook)
    finally:
        workbook.close()

    template_path = Path(template_path)
    dashboard_path = Path(dashboard_path)
    html = template_path.read_text(encoding="utf-8")
    marker = "const D = "
    start = html.find(marker)
    if start < 0:
        raise ValueError(f"Could not find embedded dashboard data in {dashboard_path}.")
    end = html.find("\n", start + len(marker))
    if end < 0:
        raise ValueError(f"Could not find the end of embedded dashboard data in {dashboard_path}.")
    serialized = json.dumps(data, ensure_ascii=False, separators=(",", ":"), allow_nan=False)
    semicolon = html.rfind(";", start, end)
    if semicolon < 0:
        raise ValueError(f"Could not find the end of the dashboard data statement in {dashboard_path}.")
    updated_html = html[:start + len(marker)] + serialized + html[semicolon:]

    dashboard_path.parent.mkdir(parents=True, exist_ok=True)
    descriptor, temp_name = tempfile.mkstemp(
        prefix=f".{dashboard_path.stem}.", suffix=".tmp", dir=dashboard_path.parent
    )
    try:
        with os.fdopen(descriptor, "w", encoding="utf-8", newline="") as output:
            output.write(updated_html)
            output.flush()
            os.fsync(output.fileno())
        os.replace(temp_name, dashboard_path)
    except Exception:
        if os.path.exists(temp_name):
            os.remove(temp_name)
        raise
    return len(data["facts"])


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--workbook", default=str(PROJECT_ROOT / "05. Revenue_analysis.xlsx"))
    parser.add_argument(
        "--template",
        default=str(PROJECT_ROOT / "templates" / "05. Revenue_dashboard.template.html"),
    )
    parser.add_argument("--dashboard", default=str(PROJECT_ROOT / "05. Revenue_dashboard.html"))
    args = parser.parse_args()
    count = update_dashboard(args.workbook, args.template, args.dashboard)
    print(f"Updated {args.dashboard} from {args.workbook} ({count} data rows).")


if __name__ == "__main__":
    main()
