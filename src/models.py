from __future__ import annotations
from dataclasses import dataclass, field
from typing import Literal

SourceLiteral = Literal["excel", "quickbooks"]
ConflictReason = Literal["name_mismatch", "missing_in_excel", "missing_in_quickbooks"]


@dataclass(slots=True)
class MiscIncome:
    record_id: str
    amount: float
    chart_of_account: str
    source: SourceLiteral
    customer_name: str = "Default Customer"

    def __str__(self):
        return (
            f"MiscIncome(record_id='{self.record_id}', amount={self.amount}, customer_name='{self.customer_name}', "
            f"chart_of_account='{self.chart_of_account}', "
            f"source='{self.source}')"
        )


@dataclass(slots=True)
class Conflict:
    record_id: str
    qb_chart_of_account: str
    excel_chart_of_account: str | None
    qb_amount: float | None
    excel_amount: float | None
    reason: ConflictReason

    def __str__(self):
        return (
            f"Conflict(record_id='{self.record_id}' qb_chart_of_account='{self.qb_chart_of_account}', excel_chart_of_account='{self.excel_chart_of_account}', "
            f"qb_amount={self.qb_amount}, excel_amount={self.excel_amount}, "
            f"reason='{self.reason}')"
        )


@dataclass(slots=True)
class ComparisonReport:
    excel_only: list[MiscIncome] = field(default_factory=list)
    qb_only: list[MiscIncome] = field(default_factory=list)
    conflicts: list[Conflict] = field(default_factory=list)
    match_count: int = 0
