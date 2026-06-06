"""Tests for TagBuilder module."""

import pytest

from audit_framework.exceptions import TagBuildError
from audit_framework.models import (
    DataType,
    ExtractionMethod,
    TagDefinition,
)
from audit_framework.tag_builder import TagBuilder


def test_build_search_tag():
    """Test building DS_SEARCH tag."""
    tag_def = TagDefinition(
        tag_id="TAG_001",
        extraction_method=ExtractionMethod.DS_SEARCH,
        field_name="InvoiceNumber",
        source_document="Invoice.pdf",
        field_type=DataType.TEXT,
        search_keywords="Invoice #",
        start_anchor="Invoice #",
        end_anchor="Date:",
    )

    tag = TagBuilder.build_search_tag(tag_def)

    assert tag.startswith("DS_SEARCH:")
    assert "InvoiceNumber" in tag
    assert "Invoice #" in tag


def test_build_coords_tag():
    """Test building DS_COORDS tag."""
    tag_def = TagDefinition(
        tag_id="TAG_002",
        extraction_method=ExtractionMethod.DS_COORDS,
        field_name="GrossAmount",
        source_document="Invoice.pdf",
        field_type=DataType.CURRENCY,
        coord_page=1,
        coord_x=150,
        coord_y=320,
        coord_width=60,
        coord_height=20,
    )

    tag = TagBuilder.build_coords_tag(tag_def)

    assert tag.startswith("DS_COORDS:")
    assert "GrossAmount" in tag
    assert "page=1" in tag
    assert "x=150" in tag


def test_build_search_tag_missing_keywords():
    """Test building DS_SEARCH tag without keywords raises error."""
    tag_def = TagDefinition(
        tag_id="TAG_003",
        extraction_method=ExtractionMethod.DS_SEARCH,
        field_name="Amount",
        source_document="Invoice.pdf",
        # Missing search_keywords
    )

    with pytest.raises(TagBuildError):
        TagBuilder.build_search_tag(tag_def)


def test_build_tag_by_method():
    """Test building tag using generic build_tag method."""
    tag_def = TagDefinition(
        tag_id="TAG_004",
        extraction_method=ExtractionMethod.DS_SEARCH,
        field_name="VendorName",
        source_document="Invoice.pdf",
        search_keywords="Vendor:",
    )

    tag = TagBuilder.build_tag(tag_def)

    assert "DS_SEARCH:" in tag
    assert "VendorName" in tag


def test_validate_tag_syntax():
    """Test tag syntax validation."""
    # Valid tag
    valid_tag = "DS_SEARCH:FieldName:keywords:(start=Start|end=End|type=text)"
    is_valid, _ = TagBuilder.validate_tag_syntax(valid_tag)
    assert is_valid

    # Invalid tag
    invalid_tag = "INVALID_TAG:something"
    is_valid, error = TagBuilder.validate_tag_syntax(invalid_tag)
    assert not is_valid
    assert error
