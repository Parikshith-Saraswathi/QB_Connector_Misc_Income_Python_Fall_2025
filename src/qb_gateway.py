from __future__ import annotations
import xml.etree.ElementTree as ET
from typing import List
from pathlib import Path

from .models import MiscIncome
from .excel_reader import extract_deposits
from .qb_reader import _send_qbxml  # Assuming your provided QB functions are in qb_reader.py

APP_NAME = "Quickbooks Connector"  # must stay the same


def add_misc_incomes_to_qb(misc_incomes: List[MiscIncome]) -> None:
    """
    Add a list of MiscIncome objects to QuickBooks as DepositAddRq requests.
    Each MiscIncome will be sent as a separate deposit line.
    """
    if not misc_incomes:
        print("No income records to add.")
        return

    for income in misc_incomes:
        qbxml = (
            '<?xml version="1.0"?>\n'
            '<?qbxml version="16.0"?>\n'
            "<QBXML>\n"
            '  <QBXMLMsgsRq onError="stopOnError">\n'
            "    <DepositAddRq requestID=\"1\">\n"
            "      <DepositAdd>\n"
            f"        <DepositToAccountRef>\n"
            f"          <FullName>{income.chart_of_account}</FullName>\n"
            f"        </DepositToAccountRef>\n"
            f"        <Memo>{income.memo}</Memo>\n"
            f"        <DepositLineAdd>\n"
            f"          <DepositLineType>CashBack</DepositLineType>\n"
            f"          <Amount>{income.amount}</Amount>\n"
            f"          <Memo>{income.memo}</Memo>\n"
            f"        </DepositLineAdd>\n"
            "      </DepositAdd>\n"
            "    </DepositAddRq>\n"
            "  </QBXMLMsgsRq>\n"
            "</QBXML>"
        )

        print(f"Sending new deposit to QuickBooks:\n{qbxml}\n")
        response = _send_qbxml(qbxml)

        # Print confirmation from QuickBooks
        status = response.find(".//*[@statusMessage]")
        msg = status.get("statusMessage") if status is not None else "Unknown response"
        print(f"QuickBooks response: {msg}\n")


# --- MAIN GUARD FOR MANUAL EXECUTION ---
if __name__ == "__main__":
    print("Reading Excel and pushing deposits to QuickBooks...\n")
    try:
        # Example: manually test with Excel file
        excel_path = Path("company_data.xlsx")
        misc_incomes = extract_deposits(excel_path)

        print(f"Loaded {len(misc_incomes)} records from Excel.")
        add_misc_incomes_to_qb(misc_incomes)

        print("All deposits processed successfully.\n")
    except Exception as e:
        print(f"Error during QuickBooks upload: {e}")
