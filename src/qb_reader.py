
import xml.etree.ElementTree as ET


#from .model import MiscIncome

from dataclasses import dataclass
from typing import Literal
SourceLiteral = Literal["excel", "quickbooks"]
@dataclass(slots=True)  
class MiscIncome:

    """Represents a payment term synchronised between Excel and QuickBooks."""
    amount :float
    customer_name: str  # Unique identifier (typically numeric days, e.g., "30")
    chart_of_account1: str
    chart_of_account2: str  # Human-readable name (e.g., "Net 30")
    memo : str
    source : SourceLiteral


import xml.etree.ElementTree as ET
from contextlib import contextmanager
from typing import Iterator, List

try:
    import win32com.client  # type: ignore
except ImportError:  # pragma: no cover
    win32com = None  # type: ignore

#from .models import MiscIncome


APP_NAME = "Quickbooks Connector"  # do not chanege this


def _require_win32com() -> None:
    if win32com is None:  # pragma: no cover - exercised via tests
        raise RuntimeError("pywin32 is required to communicate with QuickBooks")


@contextmanager
def _qb_session() -> Iterator[tuple[object, object]]:
    _require_win32com()
    session = win32com.client.Dispatch("QBXMLRP2.RequestProcessor")
    session.OpenConnection2("", APP_NAME, 1)
    ticket = session.BeginSession("", 0)
    try:
        yield session, ticket
    finally:
        try:
            session.EndSession(ticket)
        finally:
            session.CloseConnection()


def _send_qbxml(qbxml: str) -> ET.Element:
    with _qb_session() as (session, ticket):
        print(f"Sending QBXML:\n{qbxml}")  # Debug output
        raw_response = session.ProcessRequest(ticket, qbxml)  # type: ignore[attr-defined]
        print(f"Received response:\n{raw_response}")  # Debug output
    return _parse_response(raw_response)

def _parse_response(raw_xml: str) -> ET.Element:
    root = ET.fromstring(raw_xml)
    response = root.find(".//*[@statusCode]")
    if response is None:
        raise RuntimeError("QuickBooks response missing status information")

    status_code = int(response.get("statusCode", "0"))
    status_message = response.get("statusMessage", "")
    # Status code 1 means "no matching objects found" - this is OK for queries
    if status_code != 0 and status_code != 1:
        print(f"QuickBooks error ({status_code}): {status_message}")
        raise RuntimeError(status_message)
    return root

def fetch_deposit_lines()->list[MiscIncome]:
    qbxml = (
        '<?xml version="1.0"?>\n'
        '<?qbxml version="16.0"?>\n'
        "<QBXML>\n"
        '  <QBXMLMsgsRq onError="stopOnError">\n'
        "    <DepositQueryRq>\n"
        "      <IncludeLineItems >true</IncludeLineItems >\n"
        "    </DepositQueryRq>\n"
        "  </QBXMLMsgsRq>\n"
        "</QBXML>"
    )
    root = _send_qbxml(qbxml)
    deposit: list[MiscIncome] = []
    for detail in root.findall('.//DepositQueryRs/DepositRet'):
        amount = detail.findtext('DepositTotal')
        customer_Name = detail.findtext('DepositLineRet/EntityRef/FullName')
        listID = detail.findtext('DepositLineRet/AccountRef/ListID')
        account_type = fetch_account_types(listID)
        chart_of_accounts = detail.findtext('DepositLineRet/AccountRef/FullName')
        if detail.find('DepositLineRet/Memo') is not None:
            memo = detail.findtext('DepositLineRet/Memo')
        else:
            memo = "No Memo"
        misc_income = MiscIncome(
            amount=float(amount),
            customer_name=customer_Name,
            chart_of_account1=account_type,
            chart_of_account2=chart_of_accounts,
            memo = memo,
            source = "quickbooks"
        )
        deposit.append(misc_income)
    return deposit

def fetch_account_types(id: str) -> str:
    qbxml = (
        '<?xml version="1.0"?>\n'
        '<?qbxml version="16.0"?>\n'
        "<QBXML>\n"
        '  <QBXMLMsgsRq onError="stopOnError">\n'
        "    <AccountQueryRq>\n"
        f"      <ListID>{id}</ListID>\n"
        "    </AccountQueryRq>\n"
        "  </QBXMLMsgsRq>\n"
        "</QBXML>"
    )
    print(qbxml)
    root = _send_qbxml(qbxml)
    account_types: str = ""
    account_types = root.find('.//AccountQueryRs/AccountRet/AccountType')
    return account_types.text if account_types is not None else ""





#def fetch_deposit_lines()->list[MiscIncome]:
    deposit: list[MiscIncome] = []
    for detail in root.findall('.//DepositQueryRs/DepositRet'):
        amount = detail.findtext('DepositTotal')
        customer_Name = detail.findtext('DepositLineRet/EntityRef/FullName')
        listID = detail.findtext('DepositLineRet/AccountRef/ListID')
        account_type = account(listID)
        chart_of_accounts = detail.findtext('DepositLineRet/AccountRef/FullName')
        if detail.find('DepositLineRet/Memo') is not None:
            memo = detail.findtext('DepositLineRet/Memo')
        else:
            memo = "No Memo"
        misc_income = MiscIncome(
            amount=float(amount),
            customer_name=customer_Name,
            chart_of_account1=account_type,
            chart_of_account2=chart_of_accounts,
            memo = memo,
            source = "quickbooks"
        )
        deposit.append(misc_income)
    return deposit

# for entity in root.findall('.//DepositQueryRs/DepositRet/DepositLineRet'):
#     customer_Name = entity.findtext('Name')
#     print(f'Name: {name}')

# print(root.attrib)

#print all the elements in the class MiscIncome

#Reading excel file using pandas
#def fetch_details_from_excel(excel_file: str) -> list[MiscIncome]:
    wb = load_workbook(excel_file,read_only=True,values_only=True)

    #sheet = wb.active

    if "account credit nonvendor" not in wb.sheetnames:
        raise ValueError("Sheet 'account credit nonvendor' not found in the Excel file.")
    sheet = wb["account credit nonvendor"]


    misc_incomes: list[MiscIncome] = []
    for row in sheet.iter_rows(min_row=2, values_only=True):
        misc_income = MiscIncome(
            amount=row[0],
            customer_name=row[1],
            chart_of_account1=row[2],
            chart_of_account2=row[3],
            memo=row[4],
            source="excel"
        )
        misc_incomes.append(misc_income)
    return misc_incomes


if __name__ == "__main__":
    deposits = fetch_deposit_lines()
    for deposit in deposits:
        print(deposit)