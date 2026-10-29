from typing import List, Any

class DataSheetSyncTransformer:
    """Transforms Filter=1 BO_extract rows into the exact Data sheet format."""

    @classmethod
    def filter_and_map_rows(cls, bo_extract_rows: List[List[Any]]) -> List[List[Any]]:
        """
        Takes transformed 27-col BO_extract rows.
        Keeps only leaf rows (where Provider at index 6 is non-empty).
        Produces 22-column value/formula blocks for Data sheet cols A:V.
        """
        data_rows = []
        for r in bo_extract_rows:
            provider = str(r[6] or "").strip()
            if not provider:
                # Parent row, skip
                continue

            # r indices:
            # 0: Period, 1: Type, 2: Country, 3: Currency, 4: Merchant, 5: Payment Method, 6: Provider
            # 7: Deposits, 8: Deposits Fee, 9: Fee%, 10: Deposits Cost, 11: Cost%
            # 12: Withdrawals, 13: Withdrawals Fee, 14: Fee%, 15: Withdrawals Cost, 16: Cost%
            # 17: Settlement, 18: Topup, 19: Merchant Delta, 20: Provider Delta, 21: Partner Revenue

            data_row = [
                r[0], # A (1): Period
                r[1], # B (2): Type
                r[2], # C (3): Country
                r[3], # D (4): Currency
                r[4], # E (5): Merchant
                r[5], # F (6): Payment Method
                r[6], # G (7): Provider account
                r[7], # H (8): Deposits
                r[8], # I (9): Deposits Fee
                '=IFERROR(I{row}/H{row}*100,0)', # J (10): Fee%
                r[10], # K (11): Deposits Cost
                '=IFERROR(K{row}/H{row}*100,0)', # L (12): Cost%
                r[12], # M (13): Withdrawals
                r[13], # N (14): Withdrawals Fee
                '=IFERROR(N{row}/M{row}*100,0)', # O (15): Fee%
                r[15], # P (16): Withdrawals Cost
                '=IFERROR(P{row}/M{row}*100,0)', # Q (17): Cost%
                r[17], # R (18): Settlement
                r[18], # S (19): Topup
                r[19], # T (20): Merchant Delta
                r[20], # U (21): Provider Delta
                r[21]  # V (22): Partner Revenue
            ]
            data_rows.append(data_row)

        return data_rows
