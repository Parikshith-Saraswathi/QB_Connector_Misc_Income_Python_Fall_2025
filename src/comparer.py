from pathlib import Path
from typing import List
import dataclasses
import json
from datetime import datetime, timezone

from models import MiscIncome, ComparisonReport, Conflict
from excel_reader import extract_deposits
from qb_reader import fetch_deposit_lines, _send_qbxml


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


def add_excel_only_to_qb(excel_only: List[MiscIncome], bank_account: str = "Default Bank") -> List[MiscIncome]:
    """Add Excel-only deposits to QuickBooks and return the list of added items."""
    if not excel_only:
        print("No Excel-only items to add to QuickBooks.")
        return []

    added_items = []
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
        added_items.append(income)

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
        added_items = []  # if error occurs, consider none added

    return added_items


def generate_report(excel_file: Path, output_path: Path = Path("comparison_report.json"), bank_account: str = "Chase"):
    """Generate JSON report in the requested format."""
    try:
        report = compare_excel_qb(excel_file)
        added_items = add_excel_only_to_qb(report.excel_only, bank_account=bank_account)

        payload = {
            "status": "success",
            "generated_at": datetime.now(timezone.utc).isoformat(),
            "added_terms": [
                {
                    "record_id": str(i + 1),
                    "name": f"{item.chart_of_account} | {item.memo}",
                    "source": item.source,
                }
                for i, item in enumerate(added_items)
            ],
            "conflicts": [
                {
                    "record_id": str(i + 1),
                    "excel_name": c.excel_name,
                    "qb_name": c.qb_name,
                    "reason": c.reason,
                }
                for i, c in enumerate(report.conflicts)
            ] + [
                {
                    "record_id": str(i + 1),
                    "excel_name": None,
                    "qb_name": f"{item.chart_of_account} | {item.memo}",
                    "reason": "missing_in_excel",
                }
                for i, item in enumerate(report.qb_only)
            ],
            "error": None
        }

    except Exception as e:
        payload = {
            "status": "error",
            "generated_at": datetime.now(timezone.utc).isoformat(),
            "added_terms": [],
            "conflicts": [],
            "error": str(e)
        }

    # Write JSON file
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(payload, f, indent=2)
    print(f"JSON report saved to {output_path}")


if __name__ == "__main__":
    excel_path = Path("company_data.xlsx")
    generate_report(excel_path, bank_account="Chase")
