import os
import sys
import time

class ExcelCalculator:
    """Automates Excel via COM on Windows to recalculate full workbook and refresh cached formulas."""

    COM_RETRY_ATTEMPTS = 40
    COM_RETRY_DELAY_SECONDS = 0.25
    RPC_E_CALL_REJECTED = -2147418111

    @classmethod
    def _retry_excel_call(cls, operation, com_error_type):
        for attempt in range(cls.COM_RETRY_ATTEMPTS):
            try:
                return operation()
            except com_error_type as exc:
                if not exc.args or exc.args[0] != cls.RPC_E_CALL_REJECTED:
                    raise
                if attempt == cls.COM_RETRY_ATTEMPTS - 1:
                    raise
                time.sleep(cls.COM_RETRY_DELAY_SECONDS)

    @staticmethod
    def recalculate(workbook_path: str) -> None:
        abs_path = os.path.abspath(workbook_path)
        try:
            import pywintypes
            import win32com.client
        except ImportError:
            raise RuntimeError("pywin32 is required for Excel calculation on Windows.")

        excel = None
        try:
            excel = ExcelCalculator._retry_excel_call(
                lambda: win32com.client.DispatchEx("Excel.Application"),
                pywintypes.com_error,
            )
            com_error_type = pywintypes.com_error
            ExcelCalculator._retry_excel_call(
                lambda: setattr(excel, "Visible", False), com_error_type
            )
            ExcelCalculator._retry_excel_call(
                lambda: setattr(excel, "DisplayAlerts", False), com_error_type
            )
            ExcelCalculator._retry_excel_call(
                lambda: setattr(excel, "ScreenUpdating", False), com_error_type
            )

            wb = ExcelCalculator._retry_excel_call(
                lambda: excel.Workbooks.Open(abs_path), com_error_type
            )
            # Full recalculation across all sheets
            ExcelCalculator._retry_excel_call(
                lambda: excel.CalculateFullRebuild(), com_error_type
            )
            ExcelCalculator._retry_excel_call(lambda: wb.Save(), com_error_type)
            ExcelCalculator._retry_excel_call(
                lambda: wb.Close(SaveChanges=True), com_error_type
            )
        except Exception as e:
            raise RuntimeError(f"Excel COM recalculation failed: {e}")
        finally:
            if excel is not None:
                try:
                    ExcelCalculator._retry_excel_call(
                        lambda: excel.Quit(), pywintypes.com_error
                    )
                except Exception:
                    pass
