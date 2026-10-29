from abc import ABC, abstractmethod
from datetime import date
from dataclasses import dataclass
from typing import List, Dict, Any


@dataclass(frozen=True)
class BalanceTableCapture:
    rows: List[Dict[str, Any]]
    html: str
    row_count: int
    column_count: int

class BaseExtractor(ABC):
    """Abstract interface decoupling data extraction from downstream validation and writing."""

    @abstractmethod
    def extract_dashboard(self, from_date: date, to_date: date) -> List[Dict[str, Any]]:
        """Extracts hierarchical Dashboard summary table rows."""
        pass

    @abstractmethod
    def extract_merchant_balance(self) -> BalanceTableCapture:
        """Captures the complete rendered Merchant Balance table."""
        pass

    @abstractmethod
    def extract_provider_balance(self) -> BalanceTableCapture:
        """Captures the complete rendered Provider Balance table."""
        pass

    @abstractmethod
    def extract_settlements(self, from_date: date, to_date: date) -> List[Dict[str, Any]]:
        """Extracts Settlements records for the given date range."""
        pass

    @abstractmethod
    def extract_topups(self, from_date: date, to_date: date) -> List[Dict[str, Any]]:
        """Extracts Top-ups records for the given date range."""
        pass
