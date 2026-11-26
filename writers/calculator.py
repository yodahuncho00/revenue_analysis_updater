"""Automates Excel via COM on Windows to recalculate workbooks and refresh cached formulas."""

import gc
import os
import time


class ExcelCalculator:
    """Automates Excel via COM on Windows to recalculate full workbook and refresh cached formulas."""

    COM_RETRY_ATTEMPTS = 40
    COM_RETRY_DELAY_SECONDS = 0.25
    COM_MAX_RETRY_DELAY_SECONDS = 1.0
    INSTANCE_RETRY_ATTEMPTS = 3
    CALC_WAIT_TIMEOUT_SECONDS = 120.0

    RPC_E_CALL_REJECTED = -2147418111
    RPC_E_SERVERCALL_RETRYLATER = -2147417846
    COM_RETRYABLE_ERRORS = {RPC_E_CALL_REJECTED, RPC_E_SERVERCALL_RETRYLATER}
    COM_DISCONNECTED_ERRORS = {-2147417848, -2147023174}

    @classmethod
    def _retry_excel_call(cls, operation, com_error_type):
        for attempt in range(cls.COM_RETRY_ATTEMPTS):
            try:
                return operation()
            except com_error_type as exc:
                code = exc.args[0] if exc.args else None
                if code not in cls.COM_RETRYABLE_ERRORS:
                    raise
                if attempt == cls.COM_RETRY_ATTEMPTS - 1:
                    raise
                delay = min(
                    cls.COM_RETRY_DELAY_SECONDS * (1 + attempt // 8),
                    cls.COM_MAX_RETRY_DELAY_SECONDS,
                )
                time.sleep(delay)

    @classmethod
    def _recalculate_once(cls, abs_path: str) -> None:
        import pythoncom
        import pywintypes
        import win32com.client

        com_error_type = pywintypes.com_error
        pythoncom.CoInitialize()
        excel = None
        wb = None
        try:
            excel = cls._retry_excel_call(
                lambda: win32com.client.DispatchEx("Excel.Application"),
                com_error_type,
            )
            cls._retry_excel_call(lambda: setattr(excel, "Visible", False), com_error_type)
            cls._retry_excel_call(lambda: setattr(excel, "DisplayAlerts", False), com_error_type)
            cls._retry_excel_call(lambda: setattr(excel, "ScreenUpdating", False), com_error_type)
            # Prevent hidden modal prompts (macros, link updates, workbook events, overwrite alerts)
            for attr, val in (
                ("AskToUpdateLinks", False),
                ("EnableEvents", False),
                ("AlertBeforeOverwriting", False),
                ("AutomationSecurity", 3),  # msoAutomationSecurityForceDisable
            ):
                try:
                    cls._retry_excel_call(
                        lambda a=attr, v=val: setattr(excel, a, v), com_error_type
                    )
                except Exception:
                    pass

            wb = cls._retry_excel_call(
                lambda: excel.Workbooks.Open(
                    abs_path,
                    UpdateLinks=0,
                    ReadOnly=False,
                    IgnoreReadOnlyRecommended=True,
                ),
                com_error_type,
            )
            cls._retry_excel_call(lambda: excel.CalculateFullRebuild(), com_error_type)

            # Wait until Excel finishes asynchronous calculation (xlDone == 0)
            deadline = time.monotonic() + cls.CALC_WAIT_TIMEOUT_SECONDS
            while time.monotonic() < deadline:
                try:
                    state = cls._retry_excel_call(
                        lambda: excel.CalculationState, com_error_type
                    )
                    if state == 0:
                        break
                except Exception:
                    break
                time.sleep(0.1)

            cls._retry_excel_call(lambda: wb.Save(), com_error_type)
            cls._retry_excel_call(lambda: wb.Close(SaveChanges=True), com_error_type)
            wb = None
        finally:
            if wb is not None:
                try:
                    cls._retry_excel_call(
                        lambda: wb.Close(SaveChanges=False), com_error_type
                    )
                except Exception:
                    pass
                wb = None
            if excel is not None:
                try:
                    cls._retry_excel_call(lambda: excel.Quit(), com_error_type)
                except Exception:
                    pass
                excel = None
            gc.collect()
            pythoncom.CoUninitialize()

    @classmethod
    def recalculate(cls, workbook_path: str) -> None:
        abs_path = os.path.abspath(workbook_path)
        try:
            import pywintypes
            import win32com.client  # noqa: F401
        except ImportError:
            raise RuntimeError("pywin32 is required for Excel calculation on Windows.")

        last_error = None
        for instance_attempt in range(1, cls.INSTANCE_RETRY_ATTEMPTS + 1):
            try:
                cls._recalculate_once(abs_path)
                return
            except pywintypes.com_error as exc:
                code = exc.args[0] if exc.args else None
                last_error = exc
                is_recoverable = (
                    code in cls.COM_RETRYABLE_ERRORS
                    or code in cls.COM_DISCONNECTED_ERRORS
                )
                if not is_recoverable or instance_attempt >= cls.INSTANCE_RETRY_ATTEMPTS:
                    raise RuntimeError(f"Excel COM recalculation failed: {exc}") from exc
                time.sleep(1.5 * instance_attempt)
            except Exception as exc:
                raise RuntimeError(f"Excel COM recalculation failed: {exc}") from exc

        raise RuntimeError(f"Excel COM recalculation failed: {last_error}")
