import re
from typing import List, Dict, Any, Tuple
from datetime import datetime, date

class RawValidationError(Exception):
    pass

class RawDataValidator:
    """Hard gates for validating raw Backoffice captures before transformation."""

    ADDITIVE_METRICS = [
        "deposits",
        "deposits_fee",
        "deposits_cost",
        "withdrawals",
        "withdrawals_fee",
        "withdrawals_cost",
        "settlement",
        "topup",
        "partner_revenue"
    ]

    @classmethod
    def validate_non_empty(cls, records: List[Dict[str, Any]], context_name: str) -> None:
        if not records:
            raise RawValidationError(f"Raw data capture for {context_name} returned 0 rows for a non-empty range.")

    @classmethod
    def validate_dashboard_hierarchy(cls, rows: List[Dict[str, Any]], tolerance: float = 0.10) -> None:
        """
        Validates the hierarchical tree consistency in the Dashboard table.
        Tree structure:
        - Level 0: Country/Currency block header
        - Level 1: Merchant row
        - Level 2: Payment Method row
        - Level 3: Provider row (Leaf)
        
        For every parent node with children, sum(children) must equal the parent's additive columns within tolerance.
        """
        if not rows:
            raise RawValidationError("Cannot validate hierarchy on empty dashboard data.")

        cur_cc = None
        roots = []
        curr_root = None
        curr_merchant = None
        curr_method = None

        for idx, r in enumerate(rows):
            cc = (
                str(r.get("country", "") or "").strip().lower(),
                str(r.get("currency", "") or "").strip().upper()
            )
            prov = str(r.get("provider", "") or "").strip()
            meth = str(r.get("payment_method", "") or "").strip()
            merch = str(r.get("merchant", "") or "").strip()

            if cc != cur_cc:
                # Level 0: New Country/Currency root
                cur_cc = cc
                curr_root = {"idx": idx, "cc": cc, "row": r, "children": []}
                roots.append(curr_root)
                curr_merchant = None
                curr_method = None
            elif not prov and not meth:
                # Level 1: Merchant under current Country/Currency
                curr_merchant = {"idx": idx, "merch": merch, "row": r, "children": []}
                if curr_root:
                    curr_root["children"].append(curr_merchant)
                curr_method = None
            elif not prov and meth:
                # Level 2: Payment Method under current Merchant
                curr_method = {"idx": idx, "meth": meth, "row": r, "children": []}
                if curr_merchant:
                    curr_merchant["children"].append(curr_method)
            elif prov:
                # Level 3: Provider (Leaf) under current Method
                leaf = {"idx": idx, "prov": prov, "row": r}
                if curr_method:
                    curr_method["children"].append(leaf)

        # Validate sums
        for root in roots:
            if root["children"]:
                for metric in cls.ADDITIVE_METRICS:
                    p_val = float(root["row"].get(metric, 0) or 0)
                    c_sum = sum(float(m["row"].get(metric, 0) or 0) for m in root["children"])
                    diff = abs(p_val - c_sum)
                    if diff > tolerance:
                        raise RawValidationError(
                            f"Hierarchy sum mismatch for metric '{metric}' at Level 0 Root {root['cc']}: "
                            f"parent={p_val}, sum(merchants)={c_sum}, diff={diff:.4f} > {tolerance}"
                        )

            for m in root["children"]:
                if m["children"]:
                    for metric in cls.ADDITIVE_METRICS:
                        p_val = float(m["row"].get(metric, 0) or 0)
                        c_sum = sum(float(meth["row"].get(metric, 0) or 0) for meth in m["children"])
                        diff = abs(p_val - c_sum)
                        if diff > tolerance:
                            raise RawValidationError(
                                f"Hierarchy sum mismatch for metric '{metric}' at Level 1 Merchant '{m['merch']}': "
                                f"parent={p_val}, sum(methods)={c_sum}, diff={diff:.4f} > {tolerance}"
                            )

                for meth in m["children"]:
                    if meth["children"]:
                        for metric in cls.ADDITIVE_METRICS:
                            p_val = float(meth["row"].get(metric, 0) or 0)
                            c_sum = sum(float(p["row"].get(metric, 0) or 0) for p in meth["children"])
                            diff = abs(p_val - c_sum)
                            if diff > tolerance:
                                raise RawValidationError(
                                    f"Hierarchy sum mismatch for metric '{metric}' at Level 2 Method '{meth['meth']}': "
                                    f"parent={p_val}, sum(providers)={c_sum}, diff={diff:.4f} > {tolerance}"
                                )

    @classmethod
    def validate_no_duplicate_keys(cls, leaf_rows: List[Dict[str, Any]]) -> None:
        """Confirms no duplicate (country, currency, merchant, payment_method, provider) keys exist."""
        seen = set()
        for r in leaf_rows:
            key = (
                str(r.get("country", "")).strip().lower(),
                str(r.get("currency", "")).strip().upper(),
                str(r.get("merchant", "")).strip().lower(),
                str(r.get("payment_method", "")).strip().lower(),
                str(r.get("provider", "")).strip().lower()
            )
            if key in seen:
                raise RawValidationError(f"Duplicate key found in raw leaf rows: {key}")
            seen.add(key)

    @classmethod
    def validate_append_only_ids(cls, existing_ids: set, new_records: List[Dict[str, Any]], entity_name: str) -> None:
        """Confirms no newly fetched settlement/topup ID already exists in the workbook."""
        for r in new_records:
            rid = int(r["id"])
            if rid in existing_ids:
                raise RawValidationError(f"Duplicate {entity_name} ID: {rid} already exists in the workbook!")
