from datetime import date, datetime
from typing import List, Dict, Any

class BOExtractTransformer:
    """Transforms cleaned dashboard hierarchy rows into BO_extract row vectors."""

    @staticmethod
    def clean_text(val: Any) -> str:
        if val is None:
            return ""
        s = str(val).strip()
        # Remove chevrons or expand/collapse icons
        s = s.replace("▼", "").replace("▲", "").replace("▶", "").replace("▼", "")
        return s.strip()

    @staticmethod
    def clean_number(val: Any) -> float:
        if val is None:
            return 0.0
        if isinstance(val, (int, float)):
            return float(val)
        s = str(val).replace(",", "").replace(" ", "").replace("%", "").strip()
        if not s or s == "-" or s == "—":
            return 0.0
        try:
            return float(s)
        except ValueError:
            return 0.0

    @classmethod
    def transform_rows(cls, raw_rows: List[Dict[str, Any]], rpt_month_end: date) -> List[List[Any]]:
        """
        Converts raw dashboard rows into exact 27-column BO_extract row lists.
        Output order per row:
        0: Period (datetime)
        1: Type ('M')
        2: Country
        3: Currency
        4: Merchant
        5: Payment Method
        6: Provider
        7: Deposits
        8: Deposits Fee
        9: Fee%
        10: Deposits Cost
        11: Cost%
        12: Withdrawals
        13: Withdrawals Fee
        14: Fee%
        15: Withdrawals Cost
        16: Cost%
        17: Settlement
        18: Topup
        19: Merchant Delta
        20: Provider Delta
        21: Partner Revenue
        22: None (Col W spacer)
        23: Filter formula (Col X) -> string placeholder, substituted with real row number during write
        24: None (Col Y)
        25: None (Col Z)
        26: None (Col AA)
        """
        transformed = []
        period_dt = datetime(rpt_month_end.year, rpt_month_end.month, rpt_month_end.day, 0, 0)

        for r in raw_rows:
            country = cls.clean_text(r.get("country", ""))
            currency = cls.clean_text(r.get("currency", "")).upper()
            merchant = cls.clean_text(r.get("merchant", ""))
            payment_method = cls.clean_text(r.get("payment_method", ""))
            provider = cls.clean_text(r.get("provider", ""))

            deposits = cls.clean_number(r.get("deposits", 0))
            deposits_fee = cls.clean_number(r.get("deposits_fee", 0))
            dep_fee_pct = cls.clean_number(r.get("deposits_fee_pct", r.get("fee_pct_dep", 0)))
            deposits_cost = cls.clean_number(r.get("deposits_cost", 0))
            dep_cost_pct = cls.clean_number(r.get("deposits_cost_pct", r.get("cost_pct_dep", 0)))

            withdrawals = cls.clean_number(r.get("withdrawals", 0))
            withdrawals_fee = cls.clean_number(r.get("withdrawals_fee", 0))
            wd_fee_pct = cls.clean_number(r.get("withdrawals_fee_pct", r.get("fee_pct_wd", 0)))
            withdrawals_cost = cls.clean_number(r.get("withdrawals_cost", 0))
            wd_cost_pct = cls.clean_number(r.get("withdrawals_cost_pct", r.get("cost_pct_wd", 0)))

            settlement = cls.clean_number(r.get("settlement", 0))
            topup = cls.clean_number(r.get("topup", 0))
            merchant_delta = cls.clean_number(r.get("merchant_delta", 0))
            provider_delta = cls.clean_number(r.get("provider_delta", 0))
            partner_revenue = cls.clean_number(r.get("partner_revenue", 0))

            row_vector = [
                period_dt,        # Col A (1)
                "M",              # Col B (2)
                country,          # Col C (3)
                currency,         # Col D (4)
                merchant,         # Col E (5)
                payment_method,   # Col F (6)
                provider,         # Col G (7)
                deposits,         # Col H (8)
                deposits_fee,     # Col I (9)
                dep_fee_pct,      # Col J (10)
                deposits_cost,    # Col K (11)
                dep_cost_pct,     # Col L (12)
                withdrawals,      # Col M (13)
                withdrawals_fee,  # Col N (14)
                wd_fee_pct,       # Col O (15)
                withdrawals_cost, # Col P (16)
                wd_cost_pct,      # Col Q (17)
                settlement,       # Col R (18)
                topup,            # Col S (19)
                merchant_delta,   # Col T (20)
                provider_delta,   # Col U (21)
                partner_revenue,  # Col V (22)
                None,             # Col W (23)
                '=IF(G{row}="",0,1)', # Col X (24) Filter
                None,             # Col Y (25)
                None,             # Col Z (26)
                None              # Col AA (27)
            ]
            transformed.append(row_vector)

        return transformed
