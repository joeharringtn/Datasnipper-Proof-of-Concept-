"""
DataMapper - Transform and normalize extracted data.

Responsibility:
- Map raw extracted values to output schema
- Apply data transformations (type conversions, cleanups)
- Normalize formatting
- Handle lookups and cross-field calculations
"""

from datetime import datetime
from typing import Any, Optional

from audit_framework.exceptions import DataMapperError
from audit_framework.models import SchemaField, TagDefinition
from audit_framework.validator import Validator


class DataMapper:
    """Maps and transforms extracted data to schema."""

    def __init__(self) -> None:
        """Initialize DataMapper."""
        self.mappings: dict[str, str] = {}  # field_name -> tag_id mappings
        self.lookups: dict[str, dict[str, Any]] = {}  # field_name -> lookup table

    def set_field_mapping(self, field_name: str, tag_id: str) -> None:
        """
        Set mapping between field and tag.

        Args:
            field_name: Output field name
            tag_id: Tag ID from TAG_ENGINE
        """
        self.mappings[field_name] = tag_id

    def set_lookup_table(self, field_name: str, lookup_dict: dict[str, Any]) -> None:
        """
        Set lookup table for field validation.

        Args:
            field_name: Field name
            lookup_dict: Dictionary of valid values
        """
        self.lookups[field_name] = lookup_dict

    def map_row(
        self,
        raw_row: dict[str, Any],
        schema: list[SchemaField],
        tag_definitions: Optional[list[TagDefinition]] = None,
    ) -> dict[str, Any]:
        """
        Map raw extracted row to output schema.

        Args:
            raw_row: Raw extracted data (tag_id -> value)
            schema: Output schema fields
            tag_definitions: Optional tag definitions for context

        Returns:
            Mapped row (field_name -> normalized value)

        Raises:
            DataMapperError: If mapping fails
        """
        try:
            mapped_row: dict[str, Any] = {}

            # Build tag_id to tag_definition map for quick lookup
            tag_lookup = {}
            if tag_definitions:
                tag_lookup = {tag.tag_id: tag for tag in tag_definitions}

            for field in schema:
                # Find source tag for this field
                tag_id = self.mappings.get(field.field_name)

                if not tag_id:
                    # If no explicit mapping, look for tag with matching field_name
                    matching_tags = [
                        tag
                        for tag in (tag_definitions or [])
                        if tag.field_name == field.field_name
                    ]
                    if matching_tags:
                        tag_id = matching_tags[0].tag_id

                # Get raw value
                raw_value = raw_row.get(tag_id) if tag_id else None

                # Normalize value
                if raw_value is not None:
                    try:
                        normalized = Validator.normalize_value(raw_value, field.data_type)
                        mapped_row[field.field_name] = normalized
                    except Exception as e:
                        raise DataMapperError(
                            f"Failed to normalize {field.field_name}: {e}"
                        ) from e
                else:
                    mapped_row[field.field_name] = None

            return mapped_row

        except DataMapperError:
            raise
        except Exception as e:
            raise DataMapperError(f"Failed to map row: {e}") from e

    def map_batch(
        self,
        raw_rows: list[dict[str, Any]],
        schema: list[SchemaField],
        tag_definitions: Optional[list[TagDefinition]] = None,
    ) -> list[dict[str, Any]]:
        """
        Map multiple rows.

        Args:
            raw_rows: List of raw extracted rows
            schema: Output schema
            tag_definitions: Optional tag definitions

        Returns:
            List of mapped rows

        Raises:
            DataMapperError: If mapping fails
        """
        mapped_rows = []
        errors = []

        for idx, row in enumerate(raw_rows):
            try:
                mapped_row = self.map_row(row, schema, tag_definitions)
                mapped_rows.append(mapped_row)
            except DataMapperError as e:
                errors.append(f"Row {idx}: {e}")

        if errors:
            raise DataMapperError(f"Failed to map batch:\n" + "\n".join(errors))

        return mapped_rows

    def apply_lookup(
        self, value: Any, field_name: str, lookup_dict: Optional[dict[str, Any]] = None
    ) -> tuple[bool, Optional[str]]:
        """
        Validate value against lookup table.

        Args:
            value: Value to validate
            field_name: Field name
            lookup_dict: Optional lookup dict (uses stored if not provided)

        Returns:
            Tuple of (is_valid, error_message)
        """
        lookup = lookup_dict or self.lookups.get(field_name)

        if not lookup:
            return True, None

        if value not in lookup:
            valid_values = list(lookup.keys())
            return False, f"'{value}' not in valid values for {field_name}: {valid_values}"

        return True, None

    def apply_cross_field_validation(
        self,
        row: dict[str, Any],
        field1: str,
        field2: str,
        validation_fn,
    ) -> tuple[bool, Optional[str]]:
        """
        Validate cross-field relationship.

        Args:
            row: Data row
            field1: First field name
            field2: Second field name
            validation_fn: Function that validates relationship

        Returns:
            Tuple of (is_valid, error_message)
        """
        value1 = row.get(field1)
        value2 = row.get(field2)

        try:
            is_valid = validation_fn(value1, value2)
            if not is_valid:
                return False, f"Cross-field validation failed between {field1} and {field2}"
            return True, None
        except Exception as e:
            return False, f"Error validating {field1} vs {field2}: {e}"

    def denormalize_row(self, mapped_row: dict[str, Any], schema: list[SchemaField]) -> dict[str, str]:
        """
        Convert mapped row to string representation for Excel export.

        Args:
            mapped_row: Mapped data row
            schema: Output schema

        Returns:
            Row with all values as strings
        """
        output_row: dict[str, str] = {}

        for field in schema:
            value = mapped_row.get(field.field_name)

            if value is None:
                output_row[field.field_name] = ""
            elif isinstance(value, datetime):
                output_row[field.field_name] = value.strftime("%Y-%m-%d")
            else:
                output_row[field.field_name] = str(value)

        return output_row
