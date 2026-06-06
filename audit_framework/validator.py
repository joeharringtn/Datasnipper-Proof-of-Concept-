"""
Validator — structural type and format checking for extracted data.

This module sits between DataSnipper extraction and QA rule evaluation.
Its job is to confirm that raw extracted strings can be interpreted as
the types the output schema declares — before business rules run against
them.

Distinction from QAEngine
--------------------------
Validator   → "Is this value a valid number?" (structural)
QAEngine    → "Is this number within the allowed range?" (business rule)

All methods are static — Validator holds no state and is safe to call
from any context without instantiation.

Typical call sequence
---------------------
    # 1. Validate an entire batch of rows against the output schema
    errors_by_row = Validator.validate_batch(rows, schema)

    # 2. Normalize values before writing to the output sheet
    normalized = Validator.normalize_value(raw_value, DataType.CURRENCY)
"""

import re
from datetime import datetime
from typing import Any, Optional

from audit_framework.exceptions import ValidationError
from audit_framework.models import DataType, SchemaField, TagDefinition


class Validator:
    """Static validation and normalization methods for extracted data."""

    @staticmethod
    def validate_required_field(value: Any, field_name: str) -> tuple[bool, str]:
        """
        Check that a required field is not empty.

        Both None and a whitespace-only string are treated as empty — DataSnipper
        sometimes extracts a cell containing only spaces.

        Returns:
            (True, "")           if the field has content.
            (False, error_msg)   if the field is None or blank.
        """
        if value is None or (isinstance(value, str) and not value.strip()):
            return False, f"Required field '{field_name}' is empty"
        return True, ""

    @staticmethod
    def validate_data_type(
        value: Any, data_type: DataType, field_name: str
    ) -> tuple[bool, str]:
        """
        Verify that value can be interpreted as the declared data_type.

        None values are accepted for any type — required-field enforcement is
        handled separately by validate_required_field().

        Supported types and their parsing rules:
            TEXT     — always valid; no parsing needed.
            NUMBER   — must be parseable as float.
            CURRENCY — like NUMBER but strips leading $ and thousands commas.
            DATE     — tried against four common date formats; see code below.
            BOOLEAN  — must be one of yes/no/true/false/1/0 (case-insensitive).

        Returns:
            (True, "")           on success or if value is None.
            (False, error_msg)   if parsing fails.
        """
        if value is None:
            return True, ""

        value_str = str(value).strip()

        try:
            if data_type == DataType.TEXT:
                return True, ""

            elif data_type == DataType.NUMBER:
                float(value_str)
                return True, ""

            elif data_type == DataType.CURRENCY:
                # Strip common formatting before numeric parse
                float(re.sub(r"[$,]", "", value_str))
                return True, ""

            elif data_type == DataType.DATE:
                # Try formats in order of most-to-least common in audit docs
                for fmt in ["%Y-%m-%d", "%m/%d/%Y", "%d/%m/%Y", "%Y/%m/%d"]:
                    try:
                        datetime.strptime(value_str, fmt)
                        return True, ""
                    except ValueError:
                        continue
                raise ValueError("Unrecognized date format")

            elif data_type == DataType.BOOLEAN:
                if value_str.lower() in ("yes", "no", "true", "false", "1", "0"):
                    return True, ""
                raise ValueError("Must be Yes/No or True/False")

            else:
                return True, ""

        except Exception as e:
            return (
                False,
                f"Invalid {data_type.value} value '{value}' in field '{field_name}': {e}",
            )

    @staticmethod
    def validate_number_range(
        value: Any,
        field_name: str,
        min_val: Optional[float] = None,
        max_val: Optional[float] = None,
    ) -> tuple[bool, str]:
        """
        Check that a numeric value falls within [min_val, max_val] (inclusive).

        Accepts currency-formatted strings (strips $ and ,) as well as plain
        numerics.  None values pass without error.

        Args:
            value:      The value to check.
            field_name: Used in error messages.
            min_val:    Inclusive lower bound; None means no lower bound.
            max_val:    Inclusive upper bound; None means no upper bound.

        Returns:
            (True, "")           if within bounds or value is None.
            (False, error_msg)   if out of bounds or not parseable as float.
        """
        if value is None:
            return True, ""

        try:
            num_value = float(re.sub(r"[$,]", "", str(value).strip()))

            if min_val is not None and num_value < min_val:
                return False, f"Value {num_value} in '{field_name}' is below minimum {min_val}"

            if max_val is not None and num_value > max_val:
                return False, f"Value {num_value} in '{field_name}' exceeds maximum {max_val}"

            return True, ""

        except Exception as e:
            return False, f"Failed to validate range for '{field_name}': {e}"

    @staticmethod
    def validate_string_length(
        value: Any,
        field_name: str,
        min_length: Optional[int] = None,
        max_length: Optional[int] = None,
    ) -> tuple[bool, str]:
        """
        Verify that a string value falls within a length range (inclusive).

        The value is converted to string and stripped before measuring — this
        matches how DataSnipper returns extracted text.

        Args:
            value:      The value to check.
            field_name: Used in error messages.
            min_length: Inclusive minimum length; None means no minimum.
            max_length: Inclusive maximum length; None means no maximum.

        Returns:
            (True, "")           if within bounds or value is None.
            (False, error_msg)   if out of bounds.
        """
        if value is None:
            return True, ""

        length = len(str(value).strip())

        if min_length is not None and length < min_length:
            return False, f"Value in '{field_name}' is shorter than minimum {min_length}"

        if max_length is not None and length > max_length:
            return False, f"Value in '{field_name}' exceeds maximum {max_length}"

        return True, ""

    @staticmethod
    def validate_pattern(value: Any, field_name: str, pattern: str) -> tuple[bool, str]:
        """
        Validate that a value matches a regex pattern (full-string match via re.match).

        Used by QAEngine for FORMAT rules, where the QA rule definition IS the
        regex pattern string.

        Args:
            value:    The value to check.
            field_name: Used in error messages.
            pattern:  A Python-compatible regex string.

        Returns:
            (True, "")           if matches or value is None.
            (False, error_msg)   if no match or invalid regex.
        """
        if value is None:
            return True, ""

        try:
            if not re.match(pattern, str(value).strip()):
                return (
                    False,
                    f"Value '{value}' in field '{field_name}' doesn't match pattern {pattern}",
                )
            return True, ""
        except Exception as e:
            return False, f"Failed to validate pattern for '{field_name}': {e}"

    @staticmethod
    def validate_email(value: Any, field_name: str) -> tuple[bool, str]:
        """
        Validate basic email address format.

        Uses a simple regex rather than a full RFC-5322 parser — sufficient
        for the approver email fields in this framework.

        Returns:
            (True, "")           if valid email or value is None.
            (False, error_msg)   if format is invalid.
        """
        if value is None:
            return True, ""

        email_pattern = r"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$"
        value_str = str(value).strip()

        if not re.match(email_pattern, value_str):
            return False, f"Invalid email format in field '{field_name}': {value}"

        return True, ""

    @staticmethod
    def validate_row(
        row_data: dict[str, Any], schema: list[SchemaField]
    ) -> tuple[bool, list[str]]:
        """
        Validate a single data row against every field in the output schema.

        For each SchemaField:
          1. If required=True, check the field is non-empty.
          2. Check the value matches the declared data_type.

        A required-field failure skips the type check for that field (the value
        is None/empty, so type-checking it is meaningless).

        Args:
            row_data: Dict of {field_name: value} for one data row.
            schema:   List of SchemaField objects defining expected columns.

        Returns:
            (True, [])             if all fields pass.
            (False, [error, ...])  listing every validation failure.
        """
        errors: list[str] = []

        for field in schema:
            value = row_data.get(field.field_name)

            if field.required:
                is_valid, error = Validator.validate_required_field(value, field.field_name)
                if not is_valid:
                    errors.append(error)
                    continue  # no point type-checking a missing required value

            is_valid, error = Validator.validate_data_type(value, field.data_type, field.field_name)
            if not is_valid:
                errors.append(error)

        return len(errors) == 0, errors

    @staticmethod
    def validate_batch(
        rows: list[dict[str, Any]], schema: list[SchemaField]
    ) -> dict[int, list[str]]:
        """
        Validate multiple rows, returning only those with errors.

        Args:
            rows:   List of row dicts to validate.
            schema: Output schema to validate against.

        Returns:
            Dict mapping row index (0-based) → list of error strings.
            Rows that pass are absent from the returned dict, so an empty
            dict means all rows are valid.
        """
        results: dict[int, list[str]] = {}

        for idx, row in enumerate(rows):
            is_valid, errors = Validator.validate_row(row, schema)
            if not is_valid:
                results[idx] = errors

        return results

    @staticmethod
    def normalize_value(value: Any, data_type: DataType) -> Any:
        """
        Convert a raw extracted string to its canonical Python type.

        Called by DataMapper before writing values to the output schema.
        Returns None unchanged.  Raises ValidationError if the value cannot
        be parsed into the target type.

        Normalization rules:
            TEXT     → stripped str
            NUMBER   → float (strips $ and ,)
            CURRENCY → float (strips $ and ,)
            DATE     → datetime.date object (tries 4 common formats)
            BOOLEAN  → bool (yes/true/1 → True, everything else → False)

        Raises:
            ValidationError: If a DATE value cannot be parsed.
        """
        if value is None:
            return None

        value_str = str(value).strip()

        if data_type == DataType.TEXT:
            return value_str

        elif data_type in (DataType.NUMBER, DataType.CURRENCY):
            return float(re.sub(r"[$,]", "", value_str))

        elif data_type == DataType.DATE:
            for fmt in ["%Y-%m-%d", "%m/%d/%Y", "%d/%m/%Y", "%Y/%m/%d"]:
                try:
                    return datetime.strptime(value_str, fmt).date()
                except ValueError:
                    continue
            raise ValidationError(f"Could not parse date: {value}")

        elif data_type == DataType.BOOLEAN:
            return value_str.lower() in ("yes", "true", "1")

        else:
            return value_str
