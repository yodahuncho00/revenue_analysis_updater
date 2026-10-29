import csv
from datetime import date

from common.csv_export import export_raw_captures
from main import _filter_to_working_days


def test_exports_exact_raw_records_to_unique_run_directory(tmp_path):
    records = [
        {"id": 12, "amount": "1,234.50", "metadata": {"source": "backoffice"}},
        {"id": 13, "amount": 25, "created": "2026-10-28 12:00:00"},
    ]

    paths = export_raw_captures({"settlements": records}, str(tmp_path), "run-1")

    with open(paths["settlements"], encoding="utf-8-sig", newline="") as source:
        rows = list(csv.DictReader(source))

    assert rows == [
        {"id": "12", "amount": "1,234.50", "metadata": '{"source": "backoffice"}', "created": ""},
        {"id": "13", "amount": "25", "metadata": "", "created": "2026-10-28 12:00:00"},
    ]


def test_exports_append_only_records_sorted_by_numeric_id(tmp_path):
    records = [
        {"id": 100, "amount": 1},
        {"id": 20, "amount": 2},
        {"id": 3, "amount": 3},
    ]

    paths = export_raw_captures({"settlements": records}, str(tmp_path), "run-1")

    with open(paths["settlements"], encoding="utf-8-sig", newline="") as source:
        rows = list(csv.DictReader(source))

    assert [row["id"] for row in rows] == ["3", "20", "100"]
    assert [record["id"] for record in records] == [100, 20, 3]


def test_does_not_overwrite_an_existing_run_directory(tmp_path):
    export_raw_captures({"dashboard": [{"country": "Zambia"}]}, str(tmp_path), "run-1")

    try:
        export_raw_captures({"dashboard": [{"country": "DR Congo"}]}, str(tmp_path), "run-1")
    except FileExistsError:
        pass
    else:
        raise AssertionError("Expected an existing run export to be preserved.")


def test_filters_append_only_records_to_requested_working_days():
    records = [
        {"id": 1, "created": "2026-10-26 10:00:00"},
        {"id": 2, "created": "2026-10-27 10:00:00"},
        {"id": 3, "created": "2026-10-28 10:00:00"},
    ]

    filtered = _filter_to_working_days(
        records,
        [date(2026, 10, 26), date(2026, 10, 28)],
        "Settlements",
    )

    assert [record["id"] for record in filtered] == [1, 3]
