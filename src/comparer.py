from pathlib import Path
from typing import List
import dataclasses
import json
from datetime import datetime, timezone

from src.models import MiscIncome, ComparisonReport, Conflict
from src.excel_reader import extract_deposits
from src.qb_reader import fetch_deposit_lines
from src.qb_adder import add_misc_income


def compare_excel_qb(excel_data, qb_data) -> ComparisonReport:
    """Compare Excel data with QuickBooks data."""

    report = ComparisonReport()
    report.match_count = 0

    def _key(item: MiscIncome) -> str:
        return f"{item.record_id}"

    excel_dict = {_key(item): item for item in excel_data}
    qb_dict = {_key(item): item for item in qb_data}

    # Excel-only and conflicts
    for key, excel_item in excel_dict.items():
        qb_item = qb_dict.get(key)

        if not qb_item:
            report.excel_only.append(excel_item)
        else:
            # Check for perfect match
            if (
                abs(float(excel_item.amount) - float(qb_item.amount)) < 0.001
                and excel_item.chart_of_account == qb_item.chart_of_account
                and excel_item.record_id == qb_item.record_id
            ):
                report.match_count += 1
            else:
                report.conflicts.append(
                    Conflict(
                        record_id=qb_item.record_id,
                        qb_chart_of_account=qb_item.chart_of_account,
                        excel_chart_of_account=excel_item.chart_of_account,
                        qb_amount=qb_item.amount,
                        excel_amount=excel_item.amount,
                        reason="Data_mismatch",
                    )
                )

    # QuickBooks-only
    for key, qb_item in qb_dict.items():
        if key not in excel_dict:
            report.qb_only.append(qb_item)

    return report


if __name__ == "__main__":
    excel_file = Path("company_data.xlsx")
    excel_data: List[MiscIncome] = extract_deposits(excel_file)
    qb_data: List[MiscIncome] = fetch_deposit_lines()
    report = compare_excel_qb(excel_data, qb_data)

    print("Excel Only:")
    for item in report.excel_only:
        print(item)

    print("\nQuickBooks Only:")
    for item in report.qb_only:
        print(item)

    print("\nConflicts:")
    for conflict in report.conflicts:
        print(conflict)

    print("\nMatched Count:", report.match_count)

    # Save JSON report
    report_json = {
        "excel_only": [dataclasses.asdict(item) for item in report.excel_only],
        "qb_only": [dataclasses.asdict(item) for item in report.qb_only],
        "conflicts": [dataclasses.asdict(c) for c in report.conflicts],
        "match_count": report.match_count,
    }
    with open("comparison_report.json", "w") as f:
        json.dump(report_json, f, indent=4)

    print("\nComparison report saved to comparison_report.json")

    add_misc_income(report.excel_only)
