import csv
import json
import os
import tempfile
from typing import Any, Dict, List


def export_raw_captures(
    captures: Dict[str, List[Dict[str, Any]]],
    output_dir: str,
    run_id: str,
) -> Dict[str, str]:
    """Write the pipeline's exact raw records to one immutable CSV set per run."""
    run_dir = os.path.join(output_dir, run_id)
    os.makedirs(run_dir, exist_ok=False)
    paths = {}

    try:
        for name, records in captures.items():
            if any(not isinstance(record, dict) for record in records):
                raise TypeError(f"Raw capture '{name}' contains a non-object record.")

            export_records = records
            if name in {"settlements", "topups"}:
                export_records = sorted(records, key=lambda record: int(record["id"]))

            headers = list(dict.fromkeys(key for record in export_records for key in record))
            destination = os.path.join(run_dir, f"{name}.csv")
            with tempfile.NamedTemporaryFile(
                mode="w",
                encoding="utf-8-sig",
                newline="",
                dir=run_dir,
                suffix=".tmp",
                delete=False,
            ) as output:
                temp_path = output.name
                writer = csv.DictWriter(output, fieldnames=headers, extrasaction="raise")
                if headers:
                    writer.writeheader()
                    for record in export_records:
                        writer.writerow({
                            key: json.dumps(value, ensure_ascii=False) if isinstance(value, (dict, list)) else value
                            for key, value in record.items()
                        })

            os.replace(temp_path, destination)
            paths[name] = destination
    except Exception:
        for filename in os.listdir(run_dir):
            os.remove(os.path.join(run_dir, filename))
        os.rmdir(run_dir)
        raise

    return paths
