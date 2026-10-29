import os
import sys

class ExcelCalculator:
    """Automates Excel via COM on Windows to recalculate full workbook and refresh cached formulas."""

    @staticmethod
    def recalculate(workbook_path: str) -> None:
        abs_path = os.path.abspath(workbook_path)
        try:
            import win32com.client
        except ImportError:
            raise RuntimeError("pywin32 is required for Excel calculation on Windows.")

        excel = None
        try:
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.ScreenUpdating = False

            wb = excel.Workbooks.Open(abs_path)
            # Full recalculation across all sheets
            excel.CalculateFullRebuild()
            wb.Save()
            wb.Close(SaveChanges=True)
        except Exception as e:
            raise RuntimeError(f"Excel COM recalculation failed: {e}")
        finally:
            if excel is not None:
                try:
                    excel.Quit()
                except Exception:
                    pass
