from pathlib import Path
from typing import List
import dataclasses
import json

from models import MiscIncome, ComparisonReport, Conflict
from excel_reader import extract_deposits
from qb_reader import fetch_deposit_lines, _qb_session, _send_qbxml  # your QB code


def compare_excel_qb(excel_path: Path) -> ComparisonReport:
    """Compare Excel data with QuickBooks data."""
    excel_data: List[MiscIncome] = extract_deposits(excel_path)
    qb_data: List[MiscIncome] = fetch_deposit_lines()

    report = ComparisonReport()

    def _key(item: MiscIncome) -> str:
        return f"{item.chart_of_account}|{item.memo}"

    excel_dict = {_key(item): item for item in excel_data}
    qb_dict = {_key(item): item for item in qb_data}

    # Excel-only and conflicts
    for key, excel_item in excel_dict.items():
        qb_item = qb_dict.get(key)
        if not qb_item:
            report.excel_only.append(excel_item)
        else:
            if abs(excel_item.amount - qb_item.amount) > 0.001 or excel_item.memo != qb_item.memo:
                report.conflicts.append(
                    Conflict(
                        chart_of_account=excel_item.chart_of_account,
                        excel_name=f"amount: {excel_item.amount}, memo: {excel_item.memo}",
                        qb_name=f"amount: {qb_item.amount}, memo: {qb_item.memo}",
                        reason="name_mismatch",
                    )
                )

    # QuickBooks-only
    for key, qb_item in qb_dict.items():
        if key not in excel_dict:
            report.qb_only.append(qb_item)

    return report


def add_excel_only_to_qb(excel_only: List[MiscIncome], bank_account: str = "Default Bank") -> None:
    """Add Excel-only deposits to QuickBooks using your existing QB logic."""
    if not excel_only:
        print("No Excel-only items to add to QuickBooks.")
        return

    requests = []
    for income in excel_only:
        requests.append(
            f"<DepositAddRq>\n"
            f"  <DepositAdd>\n"
            f"    <DepositToAccountRef>\n"
            f"      <FullName>{bank_account}</FullName>\n"
            f"    </DepositToAccountRef>\n"
            f"    <DepositLineAdd>\n"
            f"      <AccountRef>\n"
            f"        <FullName>{income.chart_of_account}</FullName>\n"
            f"      </AccountRef>\n"
            f"      <Memo>{income.memo}</Memo>\n"
            f"      <Amount>{income.amount:.2f}</Amount>\n"
            f"    </DepositLineAdd>\n"
            f"  </DepositAdd>\n"
            f"</DepositAddRq>"
        )

    qbxml = (
        '<?xml version="1.0"?>\n'
        '<?qbxml version="13.0"?>\n'
        "<QBXML>\n"
        '  <QBXMLMsgsRq onError="continueOnError">\n' +
        "\n".join(requests) +
        "\n  </QBXMLMsgsRq>\n"
        "</QBXML>"
    )

    try:
        _send_qbxml(qbxml)
        print(f"Added {len(excel_only)} Excel-only items to QuickBooks.")
    except Exception as e:
        print(f"Failed to add Excel-only items: {e}")


if __name__ == "__main__":
    excel_file = Path("company_data.xlsx")
    report = compare_excel_qb(excel_file)

    print("Excel Only:")
    for item in report.excel_only:
        print(item)

    print("\nQuickBooks Only:")
    for item in report.qb_only:
        print(item)

    print("\nConflicts:")
    for conflict in report.conflicts:
        print(conflict)

    # Save JSON report
    report_json = {
        "excel_only": [dataclasses.asdict(item) for item in report.excel_only],
        "qb_only": [dataclasses.asdict(item) for item in report.qb_only],
        "conflicts": [dataclasses.asdict(c) for c in report.conflicts],
    }
    with open("comparison_report.json", "w") as f:
        json.dump(report_json, f, indent=4)
    print("\nComparison report saved to comparison_report.json")

    # Add Excel-only items to QuickBooks
    add_excel_only_to_qb(report.excel_only, bank_account="Chase")  # replace with your bank account
