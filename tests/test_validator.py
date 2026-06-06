"""Tests for Validator module."""

import pytest
from datetime import date

from audit_framework.models import DataType, SchemaField
from audit_framework.validator import Validator


def test_validate_required_field():
    """Test required field validation."""
    # Valid
    is_valid, _ = Validator.validate_required_field("value", "field1")
    assert is_valid

    # Empty string
    is_valid, error = Validator.validate_required_field("", "field1")
    assert not is_valid
    assert "empty" in error.lower()

    # None
    is_valid, error = Validator.validate_required_field(None, "field1")
    assert not is_valid


def test_validate_data_type_text():
    """Test text data type validation."""
    is_valid, _ = Validator.validate_data_type("hello", DataType.TEXT, "field1")
    assert is_valid


def test_validate_data_type_number():
    """Test number data type validation."""
    is_valid, _ = Validator.validate_data_type("123.45", DataType.NUMBER, "field1")
    assert is_valid

    is_valid, error = Validator.validate_data_type("abc", DataType.NUMBER, "field1")
    assert not is_valid
    assert "Invalid" in error


def test_validate_data_type_currency():
    """Test currency data type validation."""
    is_valid, _ = Validator.validate_data_type("$1,234.56", DataType.CURRENCY, "field1")
    assert is_valid

    is_valid, _ = Validator.validate_data_type("1234.56", DataType.CURRENCY, "field1")
    assert is_valid


def test_validate_data_type_date():
    """Test date data type validation."""
    is_valid, _ = Validator.validate_data_type("2026-06-04", DataType.DATE, "field1")
    assert is_valid

    is_valid, _ = Validator.validate_data_type("06/04/2026", DataType.DATE, "field1")
    assert is_valid

    is_valid, error = Validator.validate_data_type("invalid", DataType.DATE, "field1")
    assert not is_valid


def test_validate_number_range():
    """Test number range validation."""
    is_valid, _ = Validator.validate_number_range(100, "amount", min_val=0, max_val=1000)
    assert is_valid

    is_valid, error = Validator.validate_number_range(2000, "amount", min_val=0, max_val=1000)
    assert not is_valid
    assert "exceeds maximum" in error


def test_normalize_value_text():
    """Test value normalization for text."""
    result = Validator.normalize_value("  hello  ", DataType.TEXT)
    assert result == "  hello  "


def test_normalize_value_number():
    """Test value normalization for numbers."""
    result = Validator.normalize_value("$1,234.56", DataType.CURRENCY)
    assert result == 1234.56


def test_validate_row():
    """Test row validation."""
    schema = [
        SchemaField(field_name="Name", data_type=DataType.TEXT, required=True),
        SchemaField(field_name="Amount", data_type=DataType.CURRENCY, required=True),
    ]

    # Valid row
    row = {"Name": "Test", "Amount": "$100.00"}
    is_valid, errors = Validator.validate_row(row, schema)
    assert is_valid
    assert len(errors) == 0

    # Invalid row (missing required field)
    invalid_row = {"Name": "", "Amount": "$100.00"}
    is_valid, errors = Validator.validate_row(invalid_row, schema)
    assert not is_valid
    assert len(errors) > 0
