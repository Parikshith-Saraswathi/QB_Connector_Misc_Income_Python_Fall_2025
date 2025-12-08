import xml.etree.ElementTree as ET
from src.models import MiscIncome
from contextlib import contextmanager
from typing import Iterator, List
from src.input_settings import InputSettings

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


def add_misc_income(
    miscIncome: list[MiscIncome], settings: InputSettings
) -> list[MiscIncome]:
    """Create multiple Misc Income in QuickBooks in a single batch request."""

    if not miscIncome:
        return []  # Nothing to add; return early

    # Build the QBXML with multiple DepositAddRq entries
    requests = []  # Collect individual add requests to embed in one batch
    for income in miscIncome:
        try:
            # Validate that the amount is numeric
            miscAmount = float(income.amount)  # QuickBooks expects a numeric amount
        except ValueError as exc:
            raise ValueError(
                f"amount must be numeric for QuickBooks payment terms: {income.amount}"
            ) from exc

        # Build the QBXML snippet for misc income addition
        requests.append(
            f"    <DepositAddRq>\n"
            f"      <DepositAdd>\n"
            f"        <DepositToAccountRef>\n"
            f"          <FullName>{_escape_xml(str(settings.bank_account))}</FullName>\n"
            f"        </DepositToAccountRef>\n"
            f"        <DepositLineAdd>\n"
            f"          <AccountRef>\n"
            f"            <FullName>{_escape_xml(str(income.chart_of_account))}</FullName>\n"
            f"          </AccountRef>\n"
            f"          <Memo>{_escape_xml(str(income.record_id))}</Memo>\n"
            f"          <Amount>{miscAmount:.2f}</Amount>\n"
            f"        </DepositLineAdd>\n"
            f"      </DepositAdd>\n"
            f"    </DepositAddRq>"
        )

    qbxml = (
        '<?xml version="1.0"?>\n'
        '<?qbxml version="13.0"?>\n'
        "<QBXML>\n"
        '  <QBXMLMsgsRq onError="continueOnError">\n' + "\n".join(requests) + "\n"
        "  </QBXMLMsgsRq>\n"
        "</QBXML>"
    )  # Batch request enabling partial success on errors

    try:
        root = _send_qbxml(qbxml)  # Submit the batch to QuickBooks
    except RuntimeError as exc:
        # If the entire batch fails, return empty list
        print(f"Batch add failed: {exc}")
        return []

    # Parse all responses
    deposit: List[MiscIncome] = []  # Deposits confirmed/returned by QuickBooks
    for detail in root.findall(".//DepositQueryRs/DepositRet"):
        amount = detail.findtext("DepositTotal")  # Extract the amount
        memo = detail.findtext("DepositLineRet/Memo")  # Extract the memo
        customer_name = detail.findtext(
            "DepositLineRet/AccountRef/FullName"
        )  # Extract and customer name
        depositToAccount = detail.findtext(
            "DepositToAccountRef/FullName"
        )  # Extract and deposit account name
        assert depositToAccount == settings.bank_account, (
            f"DepositToAccount mismatch: expected {settings.bank_account}, got {depositToAccount}"
        )
        # assert customer_name == income.chart_of_account, f"Customer Name mismatch: expected {income.chart_of_account}, got {customer_name}"
        deposit.append(
            MiscIncome(
                amount=float(amount) if amount else 0.0,
                chart_of_account=str(customer_name),
                record_id=str(memo),
                source="quickbooks",
            )
        )

    return deposit  # Return all terms that were added/acknowledged


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
    # # Example usage: add a misc income
    # test_income = MiscIncome(
    #     amount=35.63,
    #     chart_of_account="new Income 35.63",
    #     memo="51",
    #     source="quickbooks"
    # )
    # added_incomes = add_misc_income([test_income])
    # #asserting the values which came back from QuickBooks
    # for income in added_incomes:
    #     assert income.amount == test_income.amount, f"Amount mismatch: expected {test_income.amount}, got {income.amount}"
    #     assert income.chart_of_account == test_income.chart_of_account, f"Chart of Account mismatch: expected {test_income.chart_of_account}, got {income.chart_of_account}"
    #     assert income.memo == test_income.memo, f"Memo mismatch: expected {test_income.memo}, got {income.memo}"
    #     assert income.source == test_income.source, f"Source mismatch: expected {test_income.source}, got {income.source}"
    # #print a confirmation message
    # print("Misc Income added successfully for single income entry")

    # Example usage: add multiple misc incomes
    test_incomes = [
        MiscIncome(
            amount=50.23, chart_of_account="Sales", record_id="65", source="quickbooks"
        ),
        MiscIncome(
            amount=40.01,
            chart_of_account="Misc Credits",
            record_id="66",
            source="quickbooks",
        ),
    ]
    # Example InputSettings for testing
    from src.input_settings import InputSettings

    settings = InputSettings()
    added_incomes = add_misc_income(test_incomes, settings)
    # asserting the values which came back from QuickBooks
    for income, test_income in zip(added_incomes, test_incomes):
        assert income.amount == test_income.amount, (
            f"Amount mismatch: expected {test_income.amount}, got {income.amount}"
        )
        assert income.chart_of_account == test_income.chart_of_account, (
            f"Chart of Account mismatch: expected {test_income.chart_of_account}, got {income.chart_of_account}"
        )
        assert income.record_id == test_income.record_id, (
            f"Memo(Record Id) mismatch: expected {test_income.record_id}, got {income.record_id}"
        )
        assert income.source == test_income.source, (
            f"Source mismatch: expected {test_income.source}, got {income.source}"
        )
    # print a confirmation message
    print("Misc Income added successfully for multiple income entries")
