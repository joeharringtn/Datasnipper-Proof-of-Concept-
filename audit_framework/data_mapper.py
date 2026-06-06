"""
DataMapper — bridges raw DataSnipper extraction output and the output schema.

After DataSnipper extracts values from source documents, those values are keyed
by tag_id (e.g. "TAG_001").  The output schema expects field_name keys
(e.g. "InvoiceAmount").  DataMapper resolves this mapping and normalizes each
value to its declared DataType before QAEngine sees it.

Mapping resolution order
-------------------------
1. Explicit mappings registered via set_field_mapping(field_name, tag_id).
2. Auto-match: find a TagDefinition whose field_name equals the schema field name.
3. If neither resolves, the output field is set to None.

This design lets you either pre-configure explicit mappings (useful when
tag_ids don't match field names) or rely on convention (tag field_name == output
field name) with no extra setup.

Lookup tables
-------------
set_lookup_table() registers a dict of valid values for a field.  apply_lookup()
checks a value against that table.  This is separate from LOOKUP-type QA rules
so that DataMapper can coerce / remap values (e.g. "APPROVED" → True) during
normalization before QAEngine runs.
"""

from datetime import datetime
from typing import Any, Optional

from audit_framework.exceptions import DataMapperError
from audit_framework.models import SchemaField, TagDefinition
from audit_framework.validator import Validator


class DataMapper:
    """Maps and normalizes raw extracted values into the output schema."""

    def __init__(self) -> None:
        """
        Initialize with empty mapping and lookup tables.

        Call set_field_mapping() and/or set_lookup_table() before map_row()
        if the default auto-match behavior is insufficient.
        """
        # Explicit field → tag_id overrides (set by caller)
        self.mappings: dict[str, str] = {}
        # Per-field allowed-value tables for lookup validation
        self.lookups: dict[str, dict[str, Any]] = {}

    def set_field_mapping(self, field_name: str, tag_id: str) -> None:
        """
        Register an explicit mapping from output field name to source tag ID.

        Use when the tag's field_name in TAG_ENGINE doesn't match the output
        schema column name, or when multiple tags feed into one output field.

        Args:
            field_name: Output schema column name.
            tag_id:     TAG_ENGINE tag ID that supplies the value.
        """
        self.mappings[field_name] = tag_id

    def set_lookup_table(self, field_name: str, lookup_dict: dict[str, Any]) -> None:
        """
        Register a lookup table for a field.

        The lookup table maps raw extracted values to canonical output values,
        enabling coercion during normalization (e.g. "Yes" → True).

        Args:
            field_name:  Output field this table applies to.
            lookup_dict: Dict whose keys are valid raw values.
        """
        self.lookups[field_name] = lookup_dict

    def map_row(
        self,
        raw_row: dict[str, Any],
        schema: list[SchemaField],
        tag_definitions: Optional[list[TagDefinition]] = None,
    ) -> dict[str, Any]:
        """
        Map one raw extracted row to the output schema.

        For each SchemaField:
          1. Determine the source tag_id (explicit mapping or auto-match).
          2. Read the raw value from raw_row using that tag_id.
          3. Normalize the value to the field's declared DataType.

        Args:
            raw_row:         Dict of {tag_id: raw_value} from DataSnipper output.
            schema:          Target output columns and their types.
            tag_definitions: Optional tag list for auto-match resolution.

        Returns:
            Dict of {field_name: normalized_value}.  Fields that have no
            corresponding raw value are included with value None.

        Raises:
            DataMapperError: If normalization of any field fails.
        """
        try:
            mapped_row: dict[str, Any] = {}

            # Build a tag_id → TagDefinition index once per row for auto-matching
            tag_lookup: dict[str, TagDefinition] = (
                {tag.tag_id: tag for tag in tag_definitions} if tag_definitions else {}
            )

            for field in schema:
                # Step 1: resolve which tag supplies this field's value
                tag_id = self.mappings.get(field.field_name)

                if not tag_id and tag_definitions:
                    # Auto-match: find a tag whose field_name equals the schema column name
                    matching = [t for t in tag_definitions if t.field_name == field.field_name]
                    if matching:
                        tag_id = matching[0].tag_id

                # Step 2: fetch raw value
                raw_value = raw_row.get(tag_id) if tag_id else None

                # Step 3: normalize
                if raw_value is not None:
                    try:
                        mapped_row[field.field_name] = Validator.normalize_value(
                            raw_value, field.data_type
                        )
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
        Map a list of raw rows, collecting errors across the entire batch.

        Rather than short-circuiting on the first failure, all rows are
        attempted and a single aggregated error is raised if any failed.

        Args:
            raw_rows:        List of raw {tag_id: value} dicts.
            schema:          Output schema.
            tag_definitions: Optional tag list for auto-match resolution.

        Returns:
            List of mapped {field_name: normalized_value} dicts.

        Raises:
            DataMapperError: Aggregated message listing every failed row.
        """
        mapped_rows: list[dict[str, Any]] = []
        errors: list[str] = []

        for idx, row in enumerate(raw_rows):
            try:
                mapped_rows.append(self.map_row(row, schema, tag_definitions))
            except DataMapperError as e:
                errors.append(f"Row {idx}: {e}")

        if errors:
            raise DataMapperError("Failed to map batch:\n" + "\n".join(errors))

        return mapped_rows

    def apply_lookup(
        self, value: Any, field_name: str, lookup_dict: Optional[dict[str, Any]] = None
    ) -> tuple[bool, Optional[str]]:
        """
        Check whether a value is present in a lookup table.

        Uses the provided lookup_dict if given; falls back to the table
        registered via set_lookup_table() for the same field.  Returns valid
        if no lookup table is configured (lookup is optional, not mandatory).

        Args:
            value:       The value to check.
            field_name:  The field name (used both for table lookup and error msg).
            lookup_dict: Optional explicit lookup dict; overrides stored table.

        Returns:
            (True, None)         if valid or no lookup table found.
            (False, error_msg)   if value is not in the table.
        """
        lookup = lookup_dict or self.lookups.get(field_name)

        if not lookup:
            return True, None

        if value not in lookup:
            return False, f"'{value}' not in valid values for {field_name}: {list(lookup.keys())}"

        return True, None

    def apply_cross_field_validation(
        self,
        row: dict[str, Any],
        field1: str,
        field2: str,
        validation_fn,
    ) -> tuple[bool, Optional[str]]:
        """
        Run a custom cross-field validation function against two fields in a row.

        validation_fn receives (value1, value2) and should return True if the
        relationship is valid.  Any exception raised inside validation_fn is
        caught and reported as a validation failure.

        Args:
            row:           Mapped data row.
            field1:        Name of the first field.
            field2:        Name of the second field.
            validation_fn: Callable(val1, val2) → bool.

        Returns:
            (True, None)         if the relationship is valid.
            (False, error_msg)   if invalid or if validation_fn raised.
        """
        value1 = row.get(field1)
        value2 = row.get(field2)

        try:
            if not validation_fn(value1, value2):
                return False, f"Cross-field validation failed between {field1} and {field2}"
            return True, None
        except Exception as e:
            return False, f"Error validating {field1} vs {field2}: {e}"

    def denormalize_row(
        self, mapped_row: dict[str, Any], schema: list[SchemaField]
    ) -> dict[str, str]:
        """
        Convert a mapped row to string values for Excel export.

        All typed values (float, date, bool) are converted to strings in a
        format suitable for writing back to an Excel cell.  datetime objects
        use ISO date format; everything else uses str().

        Args:
            mapped_row: Normalized {field_name: typed_value} dict.
            schema:     Output schema (defines column order).

        Returns:
            {field_name: string_value} dict; None values become empty strings.
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
