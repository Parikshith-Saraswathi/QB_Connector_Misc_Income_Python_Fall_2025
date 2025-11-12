import xml.etree.ElementTree as ET
from contextlib import contextmanager
from typing import Iterator
from .models import MiscIncome

try:
    import win32com.client  # type: ignore
except ImportError:  # pragma: no cover
    win32com = None  # type: ignore

APP_NAME = "Quickbooks Connector"  # ⚠️ do not modify


# ---------- QuickBooks Connection Utilities ---------- #

def _ensure_win32com() -> None:
    """Verify that pywin32 is installed before connecting to QuickBooks."""
    if win32com is None:
        raise RuntimeError("pywin32 must be installed to connect with QuickBooks.")


@contextmanager
def _open_qb_session() -> Iterator[tuple[object, object]]:
    """Context manager for opening and closing QuickBooks sessions."""
    _ensure_win32com()
    conn = win32com.client.Dispatch("QBXMLRP2.RequestProcessor")
    conn.OpenConnection2("", APP_NAME, 1)
    token = conn.BeginSession("", 0)
    try:
        yield conn, token
    finally:
        try:
            conn.EndSession(token)
        finally:
            conn.CloseConnection()


# ---------- Request Handling ---------- #

def _dispatch_qbxml(request_xml: str) -> ET.Element:
    """Send QBXML to QuickBooks and return parsed response."""
    with _open_qb_session() as (conn, token):
        print(f"QBXML Request:\n{request_xml}")
        response_data = conn.ProcessRequest(token, request_xml)
        print(f"QBXML Response:\n{response_data}")
    return _decode_qb_response(response_data)


def _decode_qb_response(raw_xml: str) -> ET.Element:
    """Validate and parse the QuickBooks XML response."""
    xml_root = ET.fromstring(raw_xml)
    resp_node = xml_root.find(".//*[@statusCode]")
    if resp_node is None:
        raise RuntimeError("Invalid QuickBooks response (no status code found).")

    code = int(resp_node.get("statusCode", "0"))
    message = resp_node.get("statusMessage", "")

    # Codes 0 (success) and 1 (no results) are acceptable
    if code not in (0, 1):
        print(f"QuickBooks Error [{code}]: {message}")
        raise RuntimeError(message)

    return xml_root


# ---------- Main Functionality ---------- #

def add_misc_income(company_file: str | None, incomes: list[MiscIncome]) -> list[MiscIncome]:
    """Add multiple MiscIncome entries to QuickBooks in one batch operation."""
    if not incomes:
        return []

    xml_segments = []
    for item in incomes:
        try:
            days = int(item.record_id)
        except ValueError as e:
            raise ValueError(
                f"record_id must be numeric (invalid: {item.record_id})"
            ) from e

        xml_segments.append(
            f"    <StandardTermsAddRq>\n"
            f"      <StandardTermsAdd>\n"
            f"        <Name>{_xml_escape(item.name)}</Name>\n"
            f"        <StdDiscountDays>{days}</StdDiscountDays>\n"
            f"        <DiscountPct>0</DiscountPct>\n"
            f"      </StandardTermsAdd>\n"
            f"    </StandardTermsAddRq>"
        )

    qbxml_body = (
        '<?xml version="1.0"?>\n'
        '<?qbxml version="13.0"?>\n'
        "<QBXML>\n"
        '  <QBXMLMsgsRq onError="continueOnError">\n'
        + "\n".join(xml_segments)
        + "\n  </QBXMLMsgsRq>\n</QBXML>"
    )

    try:
        parsed = _dispatch_qbxml(qbxml_body)
    except RuntimeError as err:
        print(f"QuickBooks batch add failed: {err}")
        return []

    added_entries: list[MiscIncome] = []
    for ret in parsed.findall(".//StandardTermsRet"):
        record_id = ret.findtext("StdDiscountDays")
        if not record_id:
            continue
        try:
            record_id = str(int(record_id))
        except ValueError:
            record_id = record_id.strip()
        name = (ret.findtext("Name") or "").strip()
        added_entries.append(MiscIncome(record_id=record_id, name=name, source="quickbooks"))

    return added_entries


# ---------- Helpers ---------- #

def _xml_escape(value: str) -> str:
    """Escape special characters for valid XML."""
    return (
        value.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
        .replace("'", "&apos;")
    )
