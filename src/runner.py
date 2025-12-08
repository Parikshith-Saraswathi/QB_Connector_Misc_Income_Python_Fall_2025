from __future__ import annotations
import dataclasses
from pathlib import Path
import sys
from typing import Dict, List

from .excel_reader import extract_deposits
from .qb_reader import fetch_deposit_lines
from .qb_adder import add_misc_income
from .comparer import compare_excel_qb
from .models import Conflict, MiscIncome
from .reporting import iso_timestamp, write_report

DEFAULT_REPORT_NAME = "misc_income_report.json"


def _term_to_dict(term: MiscIncome) -> Dict[str, str]:
    return {
        "id": term.record_id,
        "chart_of_account": term.chart_of_account,
        "source": term.source,
    }


def _conflict_to_dict(conflict: Conflict) -> Dict[str, object]:
    return {
        "record_id": conflict.record_id,
        "qb_chart_of_account": conflict.qb_chart_of_account,
        "excel_chart_of_account": conflict.excel_chart_of_account,
        "qb_amount": conflict.qb_amount,
        "excel_amount": conflict.excel_amount,
        "reason": conflict.reason,
    }


def _missing_in_excel_conflict(term: MiscIncome) -> Dict[str, object]:
    return {
        "record_id": term.record_id,
        "qb_chart_of_account": term.chart_of_account,
        "excel_chart_of_account": None,
        "qb_amount": term.amount,
        "excel_amount": None,
        "reason": "missing_in_excel",
    }


def run_misc_income(
    workbook_path: Path,
    *,
    bank_account_json: Path | str,
    output_path: str | None = None,
) -> Path:
    """Contract entry point for synchronising misc income."""

    report_path = Path(output_path) if output_path else Path(DEFAULT_REPORT_NAME)
    report_payload: Dict[str, object] = {
        "status": "success",
        "generated_at": iso_timestamp(),
        "added_misc_income": [],
        "conflicts": [],
        "same_misc_income": 0,
        "error": None,
    }

    from .input_settings import InputSettings

    try:
        # If running as a frozen exe, prefer treating the argument as the
        # bank account name. Otherwise, accept either a JSON path or a
        # direct bank account name.
        if getattr(sys, "frozen", False) and isinstance(bank_account_json, str):
            settings = InputSettings(bank_account=bank_account_json)
        elif (
            isinstance(bank_account_json, str) and not Path(bank_account_json).exists()
        ):
            settings = InputSettings(bank_account=bank_account_json)
        else:
            settings = InputSettings.load(Path(bank_account_json))

        excel_terms = extract_deposits(workbook_path)
        qb_terms = fetch_deposit_lines()
        comparison = compare_excel_qb(excel_terms, qb_terms)

        # Add Excel-only misc income into QB first
        add_misc_income(comparison.excel_only, settings)

        # Re-fetch QB terms so that Excel-only items are included
        updated_qb_terms = fetch_deposit_lines()
        updated_comparison = compare_excel_qb(excel_terms, updated_qb_terms)

        # Conflict collection (after adding Excel-only items)
        conflicts: List[Dict[str, object]] = []
        conflicts.extend(_conflict_to_dict(c) for c in updated_comparison.conflicts)
        conflicts.extend(
            _missing_in_excel_conflict(t) for t in updated_comparison.qb_only
        )

        # Add sections to report
        report_payload["added_misc_income"] = [
            dataclasses.asdict(item) for item in comparison.excel_only
        ]
        report_payload["conflicts"] = conflicts

        # Count matched data after adding Excel-only terms
        total_excel = len(excel_terms)
        unmatched = len(updated_comparison.excel_only) + len(
            updated_comparison.conflicts
        )
        matched = total_excel - unmatched
        report_payload["same_misc_income"] = max(matched, 0)

    except Exception as exc:
        report_payload["status"] = "error"
        report_payload["error"] = str(exc)

    write_report(report_payload, report_path)
    return report_path


__all__ = ["run_misc_income", "DEFAULT_REPORT_NAME"]


if __name__ == "__main__":
    excel_file = Path("company_data.xlsx")
    bank_account_json = Path("src/input_settings.json")
    run_misc_income(
        workbook_path=excel_file,
        bank_account_json=bank_account_json,
    )
