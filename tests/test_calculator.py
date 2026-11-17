import pytest

from writers.calculator import ExcelCalculator


class ComError(Exception):
    pass


def test_retry_excel_call_retries_rejected_com_call(monkeypatch):
    calls = []
    monkeypatch.setattr(ExcelCalculator, "COM_RETRY_ATTEMPTS", 3)
    monkeypatch.setattr(ExcelCalculator, "COM_RETRY_DELAY_SECONDS", 0)

    def operation():
        calls.append(None)
        if len(calls) == 1:
            raise ComError(ExcelCalculator.RPC_E_CALL_REJECTED)
        return "done"

    assert ExcelCalculator._retry_excel_call(operation, ComError) == "done"
    assert len(calls) == 2


def test_retry_excel_call_does_not_retry_other_com_errors(monkeypatch):
    calls = []
    monkeypatch.setattr(ExcelCalculator, "COM_RETRY_ATTEMPTS", 3)

    def operation():
        calls.append(None)
        raise ComError(-1)

    with pytest.raises(ComError):
        ExcelCalculator._retry_excel_call(operation, ComError)
    assert len(calls) == 1


def test_retry_excel_call_raises_after_retries_are_exhausted(monkeypatch):
    calls = []
    monkeypatch.setattr(ExcelCalculator, "COM_RETRY_ATTEMPTS", 2)
    monkeypatch.setattr(ExcelCalculator, "COM_RETRY_DELAY_SECONDS", 0)

    def operation():
        calls.append(None)
        raise ComError(ExcelCalculator.RPC_E_CALL_REJECTED)

    with pytest.raises(ComError):
        ExcelCalculator._retry_excel_call(operation, ComError)
    assert len(calls) == 2
