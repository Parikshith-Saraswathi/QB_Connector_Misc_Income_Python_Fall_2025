"""High-level orchestration for the Misc income CLI."""

from __future__ import annotations
import dataclasses
from pathlib import Path
from typing import Dict, List

from excel_reader import extract_deposits
from qb_reader import fetch_deposit_lines  # your QB code
from qb_adder import add_misc_income
from comparer import compare_excel_qb
from models import Conflict, MiscIncome
from reporting import iso_timestamp, write_report

DEFAULT_REPORT_NAME = "misc_income_report.json"


def _term_to_dict(term: MiscIncome) -> Dict[str, str]:
    return {
        "id": term.memo,
        "chart_of_account": term.chart_of_account,
        "source": term.source,
    }


def _conflict_to_dict(conflict: Conflict) -> Dict[str, object]:
    return {
        "chart_of_account": conflict.chart_of_account,
        "excel_name": conflict.excel_name,
        "qb_name": conflict.qb_name,
        "reason": conflict.reason,
    }


def _missing_in_excel_conflict(term: MiscIncome) -> Dict[str, object]:
    return {
        "id": term.memo,
        "excel_name": None,
        "qb_name": term.memo,
        "chart_of_account": term.chart_of_account,
        "reason": "missing_in_excel",
    }


def run_misc_income(
    workbook_path: str,
    *,
    output_path: str | None = None,
) -> Path:
    """Contract entry point for synchronising misc income.

    Args:
        company_file_path: Path to the QuickBooks company file. Use an empty
            string to reuse the currently open company file.
        workbook_path: Path to the Excel workbook containing the
            misc income worksheet.
        output_path: Optional JSON output path. Defaults to
            misc_income_report.json in the current working directory.

    Returns:
        Path to the generated JSON report.
    """

    report_path = Path(output_path) if output_path else Path(DEFAULT_REPORT_NAME)
    report_payload: Dict[str, object] = {
        "status": "success",
        "generated_at": iso_timestamp(),
        "added_misc_income": [],
        "conflicts": [],
        "error": None,
    }

    try:
        excel_terms = extract_deposits(workbook_path)
        qb_terms = fetch_deposit_lines()
        comparison = compare_excel_qb(excel_terms, qb_terms)

        add_misc_income(comparison.excel_only)

        conflicts: List[Dict[str, object]] = []
        conflicts.extend(
            _conflict_to_dict(conflict) for conflict in comparison.conflicts
        )
        conflicts.extend(
            _missing_in_excel_conflict(term) for term in comparison.qb_only
        )

        report_payload["added_misc_income"] = [
            dataclasses.asdict(item) for item in comparison.excel_only
        ]
        report_payload["conflicts"] = conflicts

    except Exception as exc:  # pragma: no cover - behaviour verified via tests
        report_payload["status"] = "error"
        report_payload["error"] = str(exc)

    write_report(report_payload, report_path)
    return report_path


__all__ = ["run_misc_income", "DEFAULT_REPORT_NAME"]


if __name__ == "__main__":
    excel_file = Path("company_data.xlsx")
    run_misc_income(str(excel_file))
