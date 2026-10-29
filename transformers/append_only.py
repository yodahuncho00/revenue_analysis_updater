from typing import List, Dict, Any, Tuple
from datetime import datetime

class AppendOnlyTransformer:
    """Transforms raw settlements and top-ups records into append-ready row vectors."""

    @staticmethod
    def parse_datetime(val: Any) -> datetime:
        if isinstance(val, datetime):
            return val
        s = str(val).strip()
        # Common formats: "YYYY-MM-DD HH:MM:SS", "YYYY-MM-DDTHH:MM:SS", "YYYY-MM-DD HH:MM:SS.ffffff"
        for fmt in (
            "%Y-%m-%d %H:%M:%S.%f",
            "%Y-%m-%d %H:%M:%S",
            "%Y-%m-%dT%H:%M:%S.%f",
            "%Y-%m-%dT%H:%M:%S",
            "%Y-%m-%d"
        ):
            try:
                return datetime.strptime(s, fmt)
            except ValueError:
                pass
        raise ValueError(f"Cannot parse datetime string: {val}")

    @classmethod
    def transform_records(
        cls,
        raw_records: List[Dict[str, Any]],
        existing_ids: set,
        entity_name: str
    ) -> List[List[Any]]:
        """
        Sorts raw records by numeric Id ascending, checks for duplicate IDs against existing_ids,
        and generates 13-column vectors for BO_settlements or BO_topups.
        """
        if not raw_records:
            return []

        # Sort by numeric ID ascending
        sorted_records = sorted(raw_records, key=lambda r: int(r["id"]))

        output_rows = []
        for r in sorted_records:
            rid = int(r["id"])
            if rid in existing_ids:
                raise ValueError(f"Duplicate {entity_name} ID: {rid} already exists in the workbook!")

            created_dt = cls.parse_datetime(r.get("created", r.get("created_at", "")))
            method = str(r.get("method", "") or "").strip()
            provider_id = int(r.get("provider_id", 0)) if str(r.get("provider_id", "")).strip().isdigit() else r.get("provider_id")
            status = str(r.get("status", "") or "").strip()
            
            amount_val = r.get("amount", 0)
            amount = float(str(amount_val).replace(",", "").strip()) if amount_val is not None else 0.0

            currency = str(r.get("currency", "") or "").strip().upper()
            country = str(r.get("country", "") or "").strip().upper()
            if country == "ZAMBIA":
                country = "ZMB"
            merchant = str(r.get("merchant", "") or "").strip()
            merchant_order_id = str(r.get("merchant_order_id", "") or "").strip()
            note = str(r.get("note", "") or "").strip()

            row_vector = [
                rid,               # Col A (1) Id
                created_dt,        # Col B (2) Created
                method,            # Col C (3) Method
                provider_id,       # Col D (4) Provider Id
                status,            # Col E (5) Status
                amount,            # Col F (6) Amount
                currency,          # Col G (7) Currency
                country,           # Col H (8) Country
                merchant,          # Col I (9) Merchant
                merchant_order_id, # Col J (10) Merchant Order Id
                note if note else None, # Col K (11) Note
                '=YEAR(B{row})',   # Col L (12) Year formula
                '=MONTH(B{row})'   # Col M (13) Month formula
            ]
            output_rows.append(row_vector)

        return output_rows
