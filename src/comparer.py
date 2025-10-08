"""Comparison helpers for misc amount."""

"""Has to be implemented"""

from __future__ import annotations

from typing import Dict, Iterable

from .models import ComparisonReport, Conflict, MiscIncome


def compare_misc_income(
    excel_terms: Iterable[MiscIncome],
    qb_terms: Iterable[MiscIncome],
) -> ComparisonReport:
    """Compare misc income from Excel and QuickBooks."""
    raise NotImplementedError("To be implemented")

__all__ = ["compare_misc_income"]