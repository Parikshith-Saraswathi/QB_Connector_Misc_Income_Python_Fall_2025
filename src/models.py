"""Domain models for misc income synchronisation."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Literal


SourceLiteral = Literal["excel", "quickbooks"]
ConflictReason = Literal["name_mismatch", "missing_in_excel", "missing_in_quickbooks"]


@dataclass(slots=True)
class MiscIncome:
    """Represents a Misc Income synchronised between Excel and QuickBooks."""

    amount: float
    chart_of_account1: str
    chart_of_account2: str
    memo: str
    source: SourceLiteral
    customer_name: str = "Default Customer"

    # Define self
    def __str__(self):
        return (
            f"MiscIncome(amount={self.amount}, customer_name='{self.customer_name}', "
            f"chart_of_account1='{self.chart_of_account1}', chart_of_account2='{self.chart_of_account2}', "
            f"memo='{self.memo}', source='{self.source}')"
        )


@dataclass(slots=True)
class Conflict:
    """Describes a discrepancy between Excel and QuickBooks for misc income."""

    record_id: str
    excel_name: str | None
    qb_name: str | None
    reason: ConflictReason


@dataclass(slots=True)
class ComparisonReport:
    """Groups comparison outcomes for later processing."""

    excel_only: list[MiscIncome] = field(default_factory=list)
    qb_only: list[MiscIncome] = field(default_factory=list)
    conflicts: list[Conflict] = field(default_factory=list)


__all__ = [
    "MiscIncome",
    "Conflict",
    "ComparisonReport",
    "ConflictReason",
    "SourceLiteral",
]
