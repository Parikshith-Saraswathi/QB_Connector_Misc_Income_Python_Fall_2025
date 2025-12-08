import xml.etree.ElementTree as ET
from src.models import MiscIncome
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


def fetch_deposit_lines(bank_account: str) -> list[MiscIncome]:
    qbxml = (
        '<?xml version="1.0"?>\n'
        '<?qbxml version="16.0"?>\n'
        "<QBXML>\n"
        '  <QBXMLMsgsRq onError="stopOnError">\n'
        "    <DepositQueryRq>\n"
        "      <AccountFilter >\n"
        f"           <FullName>{_escape_xml(str(bank_account))}</FullName>\n"
        "      </AccountFilter>\n"
        "    </DepositQueryRq>\n"
        "  </QBXMLMsgsRq>\n"
        "</QBXML>"
    )
    root = _send_qbxml(qbxml)
    deposit: list[MiscIncome] = []
    for detail in root.findall(".//DepositQueryRs/DepositRet"):
        amount = detail.findtext("DepositTotal") or "0.0"
        deposit_to_account = detail.findtext("DepositToAccountRef/FullName") or ""
        customer_name = detail.findtext("DepositLineRet/AccountRef/FullName") or ""
        memo = detail.findtext("DepositLineRet/Memo") or ""
        try:
            misc_income = MiscIncome(
                amount=float(amount),
                customer_name=deposit_to_account,
                chart_of_account=customer_name,
                record_id=memo,
                source="quickbooks",
            )
        except ValueError:
            # Skip if amount cannot be converted to float
            continue
        deposit.append(misc_income)
    return deposit


def _escape_xml(value: str) -> str:
    """Escape XML special characters for safe QBXML construction."""
    return (
        value.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
        .replace("'", "&apos;")
    )


# def fetch_account_types(id: str) -> str:
#     qbxml = (
#         '<?xml version="1.0"?>\n'
#         '<?qbxml version="16.0"?>\n'
#         "<QBXML>\n"
#         '  <QBXMLMsgsRq onError="stopOnError">\n'
#         "    <AccountQueryRq>\n"
#         f"      <ListID>{id}</ListID>\n"
#         "    </AccountQueryRq>\n"
#         "  </QBXMLMsgsRq>\n"
#         "</QBXML>"
#     )
#     # print(qbxml)
#     root = _send_qbxml(qbxml)
#     account_types = root.find(".//AccountQueryRs/AccountRet/AccountType")
#     if account_types is not None and account_types.text is not None:
#         return account_types.text
#     return "Unknown Account Type"


__all__ = ["fetch_deposit_lines", "MiscIncome"]

if __name__ == "__main__":
    # Example usage: pass a bank account name
    deposits = fetch_deposit_lines("Chase")
    for deposit in deposits:
        print(deposit)
