import pytest

from src.models import MiscIncome


@pytest.fixture
def mock_excel_terms():
    """Mock Excel payment terms."""
    return [
        MiscIncome(
            amount=100.0,
            memo="Term 30",
            chart_of_account1="30",
            chart_of_account2="Net 30",
            source="excel",
        ),
        MiscIncome(
            amount=200.0,
            memo="Term 45",
            chart_of_account1="45",
            chart_of_account2="Net 45",
            source="excel",
        ),
    ]


def test_misc_income_attributes(mock_excel_terms):
    """Test that MiscIncome objects have correct attributes and values."""
    term = mock_excel_terms[0]

    assert term.amount == 100.0
    assert term.memo == "Term 30"
    assert term.chart_of_account1 == "30"
    assert term.chart_of_account2 == "Net 30"
    assert term.source == "excel"
    assert term.customer_name == "Default Customer"  # Default value from the class

    # Test string representation
    expected_str = (
        "MiscIncome(amount=100.0, customer_name='Default Customer', "
        "chart_of_account1='30', chart_of_account2='Net 30', "
        "memo='Term 30', source='excel')"
    )
    assert str(term) == expected_str
