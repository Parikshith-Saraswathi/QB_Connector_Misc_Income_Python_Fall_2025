from __future__ import annotations
from pathlib import Path
from typing import List
from openpyxl import load_workbook

class DepositRecord:
    """Represents one deposit or account row from Excel."""
    def __init__(self, record_id: str, type_: str, number: str, name: str, source: str = "excel"):
        self.record_id = record_id
        self.type = type_
        self.number = number
        self.name = name
        self.source = source

    def to_dict(self):
        return {
            "ID": self.record_id,
            "Type": self.type,
            "Number": self.number,
            "Name": self.name,
            "Source": self.source,
        }

def extract_deposits(workbook_path: Path) -> List[DepositRecord]:
    """Extract deposit-related data from company_data.xlsx"""
    workbook_path = Path(workbook_path)
    if not workbook_path.exists():
        raise FileNotFoundError(f"Workbook not found: {workbook_path}")

    wb = load_workbook(filename=workbook_path, read_only=True, data_only=True)
    sheet = wb.active

    rows = sheet.iter_rows(values_only=True)
    headers = [str(h).strip() if h else "" for h in next(rows, [])]
    header_index = {h: i for i, h in enumerate(headers)}

    def _value(row, column):
        idx = header_index.get(column)
        if idx is None or idx >= len(row):
            return None
        return row[idx]

    records = []
    for row in rows:
        record_id = _value(row, "ID")
        if not record_id:
            continue
        type_ = _value(row, "Type") or ""
        number = _value(row, "Number") or ""
        name = _value(row, "Name") or ""

        records.append(DepositRecord(str(record_id), str(type_), str(number), str(name)))

    wb.close()
    return records

__all__ = ["extract_deposits", "DepositRecord"]
