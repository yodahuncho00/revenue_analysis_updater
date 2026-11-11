"""Refresh the transaction data embedded in the local POS dashboard."""

import json
import os
import re
import tempfile
from datetime import date, datetime
from pathlib import Path

import openpyxl


from openpyxl.utils.datetime import from_excel


RAW_BLOCK = re.compile(
    r'(<script id="rawdata" type="application/json">)(.*?)(</script>)', re.DOTALL
)
TX_KEYS = (
    "t", "d", "ok", "g", "rs", "a", "mf", "cf", "f", "c", "b", "n",
    "cu", "p", "m", "tr", "me", "pr", "ccy", "mt", "pt", "test",
)
DICT_KEYS = {"d", "g", "rs", "cu", "p", "m", "tr", "me", "pr", "ccy", "mt", "pt"}


def _json_scalar(value):
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%dT%H:%M")
    if isinstance(value, date):
        return value.isoformat()
    if value is None:
        return ""
    if isinstance(value, float) and (value != value or abs(value) == float("inf")):
        return None
    return value


def _json_date(value):
    if isinstance(value, (datetime, date)):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, (int, float)):
        try:
            return from_excel(value).strftime("%Y-%m-%d")
        except Exception:
            pass
    return _json_scalar(value)


def _transaction_columns(workbook):
    sheet = workbook["BO_extract"]
    merchant_types = {}
    merchant_tests = {}
    if "Merchant_master" in workbook.sheetnames:
        for row in workbook["Merchant_master"].iter_rows(min_row=5, values_only=True):
            if row[1]:
                merchant_types[str(row[1])] = row[4] or "TBD"
                merchant_tests[str(row[1])] = row[6] or 0
    parent_types = {}
    if "Parent_master" in workbook.sheetnames:
        for row in workbook["Parent_master"].iter_rows(min_row=5, values_only=True):
            if row[1]:
                parent_types[str(row[1])] = row[3] or "Independent (single merchant)"

    cols = {key: [] for key in TX_KEYS}
    for row in sheet.iter_rows(min_row=4, values_only=True):
        if not row[0]:
            continue
        created = row[1]
        if isinstance(created, (int, float)):
            try:
                created = from_excel(created)
            except Exception:
                pass
        elif isinstance(created, str):
            try:
                created = datetime.fromisoformat(created.strip().replace(" ", "T"))
            except Exception:
                pass
        if not isinstance(created, (datetime, date)):
            raise ValueError(f"BO_extract payment {row[0]!r} has no calculated Datetime value ({created!r}).")
        values = {
            "t": created,
            "d": "D" if str(row[2]).strip().lower() == "receive" else "W",
            "ok": (row[29] if len(row) > 29 else 0) or 0,
            "g": (row[30] if len(row) > 30 else "") or "",
            "rs": (row[31] if len(row) > 31 else "") or "",
            "a": (row[3] if len(row) > 3 else 0) or 0,
            "mf": (row[5] if len(row) > 5 else 0) or 0,
            "cf": (row[15] if len(row) > 15 else 0) or 0,
            "f": ((row[5] if len(row) > 5 else 0) or 0) + ((row[15] if len(row) > 15 else 0) or 0),
            "c": (row[6] if len(row) > 6 else 0) or 0,
            "b": (row[17] if len(row) > 17 else 0) or 0,
            "n": (row[7] if len(row) > 7 else 0) or 0,
            "cu": (row[53] if len(row) > 53 else "") or "",
            "p": (row[43] if len(row) > 43 else "") or "",
            "m": (row[44] if len(row) > 44 else "") or "",
            "tr": (row[45] if len(row) > 45 else "") or "",
            "me": (row[38] if len(row) > 38 else "") or "",
            "pr": (row[39] if len(row) > 39 else "") or "",
            "ccy": (row[34] if len(row) > 34 else "") or "",
            "mt": merchant_types.get(str((row[44] if len(row) > 44 else "") or ""), "TBD"),
            "pt": parent_types.get(str((row[43] if len(row) > 43 else "") or ""), "Independent (single merchant)"),
            "test": merchant_tests.get(str((row[44] if len(row) > 44 else "") or ""), 0),
        }
        for key in TX_KEYS:
            cols[key].append(_json_scalar(values[key]))

    dictionaries = {}
    for key in DICT_KEYS:
        mapping = {}
        encoded = []
        for value in cols[key]:
            if value not in mapping:
                mapping[value] = len(mapping)
            encoded.append(mapping[value])
        dictionaries[key] = list(mapping)
        cols[key] = encoded
    return {"n": len(cols["t"]), "cols": cols, "dicts": dictionaries}


def update_pos_dashboard(workbook_path, template_path, dashboard_path):
    workbook_path = Path(workbook_path)
    dashboard_path = Path(dashboard_path)
    workbook = openpyxl.load_workbook(workbook_path, read_only=True, data_only=True)
    try:
        tx = _transaction_columns(workbook)
        last_date = max(
            (value for value in tx["cols"]["t"]),
            default="",
        )
        control = workbook["Control"] if "Control" in workbook.sheetnames else None
        asof = control["B6"].value if control else None
        rpt = control["B5"].value if control else None
        lists = workbook["Lists"] if "Lists" in workbook.sheetnames else None
        merchant_types = [
            lists.cell(row, 1).value
            for row in range(5, 65)
            if lists.cell(row, 1).value
        ] if lists else []
        bands = [
            lists.cell(row, 18).value
            for row in range(5, 13)
            if lists.cell(row, 18).value is not None
        ] if lists else []
        fxall, fxclose = {}, {}
        if "FX" in workbook.sheetnames:
            fx_sheet = workbook["FX"]
            avg_ccys = [
                str(fx_sheet.cell(5, column).value)
                for column in range(5, 15)
                if fx_sheet.cell(5, column).value
            ]
            close_ccys = [
                str(fx_sheet.cell(5, column).value)
                for column in range(16, 26)
                if fx_sheet.cell(5, column).value
            ]
            for row in fx_sheet.iter_rows(min_row=6, values_only=True):
                year, month = row[2], row[3]
                if year is None or month is None:
                    continue
                month_index = int(year) * 12 + int(month)
                fxall[str(month_index)] = {
                    ccy: _json_scalar(row[4 + index])
                    for index, ccy in enumerate(avg_ccys)
                    if row[4 + index] is not None
                }
                fxclose[str(month_index)] = {
                    ccy: _json_scalar(row[15 + index])
                    for index, ccy in enumerate(close_ccys)
                    if row[15 + index] is not None
                }
        channels = []
        if "Channel_master" in workbook.sheetnames:
            for row in workbook["Channel_master"].iter_rows(min_row=5, values_only=True):
                if not row[0]:
                    continue
                channels.append({
                    "ch": row[0], "prov": row[3], "method": row[4], "dir": row[5],
                    "fee": _json_scalar(row[8]) or 0,
                    "bonus": _json_scalar(row[9]) or 0,
                    "cost": _json_scalar(row[10]) or 0,
                    "costp": _json_scalar(row[11]) or 0,
                })
        rc = (control["B20"].value if control else None) or "ZMW"
    finally:
        workbook.close()

    html = Path(template_path).read_text(encoding="utf-8")
    match = RAW_BLOCK.search(html)
    if not match:
        raise ValueError(f"Could not find embedded POS dashboard data in {dashboard_path}.")
    raw = json.loads(match.group(2))
    raw["tx"] = tx
    raw["fxall"] = fxall
    raw["fxclose"] = fxclose
    raw["chans"] = channels
    raw["rc"] = rc
    raw.setdefault("meta", {})
    raw["meta"].update({
        "asof": _json_date(asof),
        "rpt": _json_date(rpt),
        "lastTxn": last_date[:10] if last_date else "",
        "built": datetime.now().strftime("%Y-%m-%d %H:%M"),
    })
    if merchant_types:
        raw["types"] = merchant_types
    if bands:
        raw["bands"] = bands
    serialized = json.dumps(raw, ensure_ascii=False, separators=(",", ":"), allow_nan=False)
    updated = html[:match.start(2)] + serialized + html[match.end(2):]

    descriptor, temp_name = tempfile.mkstemp(
        prefix=f".{dashboard_path.stem}.", suffix=".tmp", dir=dashboard_path.parent
    )
    try:
        with os.fdopen(descriptor, "w", encoding="utf-8", newline="") as output:
            output.write(updated)
            output.flush()
            os.fsync(output.fileno())
        os.replace(temp_name, dashboard_path)
    except Exception:
        if os.path.exists(temp_name):
            os.remove(temp_name)
        raise
    return tx["n"]
