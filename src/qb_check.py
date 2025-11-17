import xml.etree.ElementTree as ET
from .models import MiscIncome
from contextlib import contextmanager
from typing import Iterator

try:
    import win32com.client  # type: ignore
except ImportError:  # pragma: no cover
    win32com = None  # type: ignore

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


def fetch_deposit_lines() -> list[MiscIncome]:
    qbxml = f"""<?xml version="1.0" encoding="utf-8"?>
<?qbxml version="16.0"?>
<QBXML>
  <QBXMLMsgsRq onError="stopOnError">
    <DepositAddRq>
      <DepositAdd>
        <DepositToAccountRef>
          <FullName>Chase</FullName>
        </DepositToAccountRef>
        <Memo>{_escape_xml("Rent")}</Memo>
        <DepositLineAdd>
          <AccountRef>
            <FullName>{_escape_xml("Rent")}</FullName>
          </AccountRef>
          <Memo>{_escape_xml("2")}</Memo>
          <Amount>5.00</Amount>
        </DepositLineAdd>
      </DepositAdd>
    </DepositAddRq>
  </QBXMLMsgsRq>
</QBXML>
"""
    root = _send_qbxml(qbxml)
    print(root)
    # deposit: list[MiscIncome] = []
    # for detail in root.findall(".//DepositQueryRs/DepositRet"):
    #     amount = detail.findtext("DepositTotal") or "0.0"
    #     customer_Name = detail.findtext("DepositLineRet/EntityRef/FullName") or ""
    #     listID = detail.findtext("DepositLineRet/AccountRef/ListID")
    #     # Only call fetch_account_types if listID exists
    #     chart_of_accounts = detail.findtext("DepositLineRet/AccountRef/FullName") or ""
    #     memo = detail.findtext("DepositLineRet/Memo") or ""
    #     # if detail.find("DepositLineRet/Memo") is not None:
    #     #     memo = detail.findtext("DepositLineRet/Memo") or ""
    #     try:
    #         misc_income = MiscIncome(
    #             amount=float(amount),
    #             customer_name=customer_Name,
    #             chart_of_account=chart_of_accounts,
    #             memo=memo,
    #             source="quickbooks",
    #         )
    #     except ValueError:
    #         # Skip if amount cannot be converted to float
    #         continue
    #     deposit.append(misc_income)
    return []


def _escape_xml(value: str) -> str:
    """Escape XML special characters for safe QBXML construction."""
    return (
        value.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
        .replace("'", "&apos;")
    )


if __name__ == "__main__":
    fetch_deposit_lines()
