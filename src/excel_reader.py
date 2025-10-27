from __future__ import annotations
from pathlib import Path
from typing import List
from openpyxl import load_workbook
from model import MiscIncome


class AccountCreditRecord:
    """Represents one 'Account Credit Nonvendor' row from Excel."""
    def __init__(self,
                 customer_name: str,
                 memo: str,
                 amount: str,
                 chart_of_account2: str,
                 chart_of_account1: str,
                 source: str = "excel"):
        self.customer_name = customer_name
        self.memo = memo
        self.amount = amount
        self.chart_of_account2 = chart_of_account2
        self.chart_of_account1 = chart_of_account1
        self.source = source

    def to_dict(self):
        return {
            "Customer Name": self.customer_name,
            "Memo": self.memo,
            "Amount": self.amount,
            "chart_of_account2": self.chart_of_account2,
            "chart_of_account1": self.chart_of_account1,
            "Source": self.source
        }


def extract_account_credit_records(workbook_path: Path) -> List[AccountCreditRecord]:
    """Extract rows from the 'Account Credit Nonvendor' sheet of the Excel file."""
    workbook_path = Path(workbook_path)
    if not workbook_path.exists():
        raise FileNotFoundError(f"Workbook not found: {workbook_path}")

    wb = load_workbook(filename=workbook_path, read_only=True, data_only=True)
    sheet = wb["account credit nonvendor"]  # change if needed

    rows = sheet.iter_rows(values_only=True)
    headers = [str(h).strip() if h else "" for h in next(rows, [])]
    header_index = {h: i for i, h in enumerate(headers)}

    def _value(row, column):
        idx = header_index.get(column)
        if idx is None or idx >= len(row):
            return None
        return row[idx]

    records: List[AccountCreditRecord] = []
    for row in rows:
        customer_name = _value(row, "Customer") or _value(row, "Parent ID")
        if not customer_name:
            continue

        record = AccountCreditRecord(
            str(customer_name),
            str(_value(row, "Child ID") or ""),
            str(_value(row, "Check Amount") or ""),
            str(_value(row, "Tier 2 - Chart of Account") or ""),
            str(_value(row, "Tier 1 - Type") or "")
        )
        records.append(record)

    wb.close()
    return records


__all__ = ["extract_account_credit_records", "AccountCreditRecord"]
