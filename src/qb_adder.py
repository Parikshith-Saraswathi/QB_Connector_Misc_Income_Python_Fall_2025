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

def add_misc_income(company_file:str |None, incomes: list[MiscIncome]) -> list[MiscIncome]:
    """Create multiple Misc Income in QuickBooks in a single batch request."""

    if not incomes:
        return []  # Nothing to add; return early

    # Build the QBXML with multiple StandardTermsAddRq entries
    requests = []  # Collect individual add requests to embed in one batch
    for income in incomes:
        try:
            days_value = int(term.record_id)  # QuickBooks expects a numeric days value
        except ValueError as exc:
            raise ValueError(
                f"record_id must be numeric for QuickBooks payment terms: {term.record_id}"
            ) from exc

        # Build the QBXML snippet for this term creation
        requests.append(
            f"    <StandardTermsAddRq>\n"
            f"      <StandardTermsAdd>\n"
            f"        <Name>{_escape_xml(term.name)}</Name>\n"
            f"        <StdDiscountDays>{days_value}</StdDiscountDays>\n"
            f"        <DiscountPct>0</DiscountPct>\n"
            f"      </StandardTermsAdd>\n"
            f"    </StandardTermsAddRq>"
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
    added_terms: List[PaymentTerm] = []  # Terms confirmed/returned by QuickBooks
    for term_ret in root.findall(".//StandardTermsRet"):
        record_id = term_ret.findtext("StdDiscountDays")  # Extract the ID
        if not record_id:
            continue
        try:
            record_id = str(int(record_id))  # Normalise numeric string
        except ValueError:
            record_id = record_id.strip()
        name = (term_ret.findtext("Name") or "").strip()  # Extract and trim name
        added_terms.append(
            PaymentTerm(record_id=record_id, name=name, source="quickbooks")
        )

    return added_terms  # Return all terms that were added/acknowledged


def add_payment_term(company_file: str | None, term: PaymentTerm) -> PaymentTerm:
    """Create a single payment term in QuickBooks and return the stored record."""

    try:
        days_value = int(term.record_id)  # Validate that the ID is numeric
    except ValueError as exc:
        raise ValueError(
            "record_id must be numeric for QuickBooks payment terms"
        ) from exc

    qbxml = (
        '<?xml version="1.0"?>\n'
        '<?qbxml version="13.0"?>\n'
        "<QBXML>\n"
        '  <QBXMLMsgsRq onError="stopOnError">\n'
        "    <StandardTermsAddRq>\n"
        "      <StandardTermsAdd>\n"
        f"        <Name>{_escape_xml(term.name)}</Name>\n"
        f"        <StdDiscountDays>{days_value}</StdDiscountDays>\n"
        "        <DiscountPct>0</DiscountPct>\n"
        "      </StandardTermsAdd>\n"
        "    </StandardTermsAddRq>\n"
        "  </QBXMLMsgsRq>\n"
        "</QBXML>"
    )  # Single add request with stop-on-error behavior

    try:
        root = _send_qbxml(qbxml)  # Send request and parse response
    except RuntimeError as exc:
        # Check if error is "name already in use" (error code 3100)
        if "already in use" in str(exc):
            # Return the term as-is since it already exists
            return PaymentTerm(
                record_id=term.record_id, name=term.name, source="quickbooks"
            )
        raise  # Re-raise other errors

    term_ret = root.find(".//StandardTermsRet")  # Extract the returned record, if any
    if term_ret is None:
        # Some responses may omit the created object; fall back to input values
        return PaymentTerm(
            record_id=term.record_id, name=term.name, source="quickbooks"
        )

    record_id = term_ret.findtext("StdDiscountDays") or term.record_id  # Prefer QB's ID
    try:
        record_id = str(int(record_id))  # Normalise to a clean numeric string
    except ValueError:
        record_id = record_id.strip()
    name = (term_ret.findtext("Name") or term.name).strip()  # Prefer QB's name

    return PaymentTerm(record_id=record_id, name=name, source="quickbooks")

def _escape_xml(value: str) -> str:
    """Escape XML special characters for safe QBXML construction."""
    return (
        value.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
        .replace("'", "&apos;")
    )