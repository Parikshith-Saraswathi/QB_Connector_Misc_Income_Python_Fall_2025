"""Domain models for misc income synchronisation."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Literal


SourceLiteral = Literal["excel", "quickbooks"]
ConflictReason = Literal["name_mismatch", "missing_in_excel", "missing_in_quickbooks"]


@dataclass(slots=True)
class MiscIncome:
    """Represents a misc income synchronised between Excel and QuickBooks."""

    record_id: str
    name: str
    source: SourceLiteral


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
