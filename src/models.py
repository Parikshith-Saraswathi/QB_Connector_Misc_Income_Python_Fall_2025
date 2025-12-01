from __future__ import annotations
from dataclasses import dataclass, field
from typing import Literal

SourceLiteral = Literal["excel", "quickbooks"]
ConflictReason = Literal["name_mismatch", "missing_in_excel", "missing_in_quickbooks"]


@dataclass(slots=True)
class MiscIncome:
    amount: float
    chart_of_account: str
    memo: str
    source: SourceLiteral
    customer_name: str = "Default Customer"

    def __str__(self):
        return (
            f"MiscIncome(amount={self.amount}, customer_name='{self.customer_name}', "
            f"chart_of_account='{self.chart_of_account}', "
            f"memo='{self.memo}', source='{self.source}')"
        )


@dataclass(slots=True)
class Conflict:
    chart_of_account: str
    excel_name: str | None
    qb_name: str | None
    reason: ConflictReason

    def __str__(self):
        return (
            f"Conflict(chart_of_account='{self.chart_of_account}', "
            f"excel_name='{self.excel_name}', qb_name='{self.qb_name}', reason='{self.reason}')"
        )


@dataclass(slots=True)
class ComparisonReport:
    excel_only: list[MiscIncome] = field(default_factory=list)
    qb_only: list[MiscIncome] = field(default_factory=list)
    conflicts: list[Conflict] = field(default_factory=list)
    match_count: int = 0  # <── NEW FIELD
