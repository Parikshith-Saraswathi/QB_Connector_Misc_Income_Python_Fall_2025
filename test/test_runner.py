import pytest

from src.models import MiscIncome


@pytest.fixture
def mock_excel_terms():
    """Mock Excel payment terms."""
    return [
        MiscIncome(
            amount=100.0,
            record_id="Term 30",
            chart_of_account="Rental",
            source="excel",
        ),
        MiscIncome(
            amount=200.0,
            record_id="Term 45",
            chart_of_account="Net 45",
            source="excel",
        ),
    ]


def test_misc_income_attributes(mock_excel_terms):
    """Test that MiscIncome objects have correct attributes and values."""
    term = mock_excel_terms[0]

    assert term.amount == 100.0
    assert term.record_id == "Term 30"
    assert term.chart_of_account == "Rental"
    assert term.source == "excel"
    assert term.customer_name == "Default Customer"  # Default value from the class

    # Test string representation
    expected_str = (
        "MiscIncome(record_id='Term 30', amount=100.0, customer_name='Default Customer', "
        "chart_of_account='Rental', "
        "source='excel')"
    )
    assert str(term) == expected_str
