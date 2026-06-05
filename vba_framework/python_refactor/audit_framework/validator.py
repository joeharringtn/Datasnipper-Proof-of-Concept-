"""
Validator - Input validation and format checking.

Responsibility:
- Validate extracted data completeness
- Check data type and format compliance
- Flag missing or malformed values
- Generate validation reports
"""

import re
from datetime import datetime
from typing import Any, Optional

from audit_framework.exceptions import ValidationError
from audit_framework.models import DataType, SchemaField, TagDefinition


class Validator:
    """Validates extracted data against schema and rules."""

    @staticmethod
    def validate_required_field(value: Any, field_name: str) -> tuple[bool, str]:
        """
        Check if required field has a value.

        Args:
            value: Field value
            field_name: Field name

        Returns:
            Tuple of (is_valid, error_message)
        """
        if value is None or (isinstance(value, str) and not value.strip()):
            return False, f"Required field '{field_name}' is empty"
        return True, ""

    @staticmethod
    def validate_data_type(
        value: Any, data_type: DataType, field_name: str
    ) -> tuple[bool, str]:
        """
        Validate value matches expected data type.

        Args:
            value: Field value
            data_type: Expected data type
            field_name: Field name

        Returns:
            Tuple of (is_valid, error_message)
        """
        if value is None:
            return True, ""  # Null is acceptable for optional fields

        value_str = str(value).strip()

        try:
            if data_type == DataType.TEXT:
                # Text accepts anything
                return True, ""

            elif data_type == DataType.NUMBER:
                float(value_str)
                return True, ""

            elif data_type == DataType.CURRENCY:
                # Remove common currency symbols
                cleaned = re.sub(r"[$,]", "", value_str)
                float(cleaned)
                return True, ""

            elif data_type == DataType.DATE:
                # Try to parse common date formats
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
            return False, f"Invalid {data_type.value} value '{value}' in field '{field_name}': {e}"

    @staticmethod
    def validate_number_range(
        value: Any, field_name: str, min_val: Optional[float] = None, max_val: Optional[float] = None
    ) -> tuple[bool, str]:
        """
        Validate numeric value is within range.

        Args:
            value: Numeric value
            field_name: Field name
            min_val: Minimum allowed value (inclusive)
            max_val: Maximum allowed value (inclusive)

        Returns:
            Tuple of (is_valid, error_message)
        """
        if value is None:
            return True, ""

        try:
            value_str = str(value).strip()
            cleaned = re.sub(r"[$,]", "", value_str)
            num_value = float(cleaned)

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
        Validate string length.

        Args:
            value: String value
            field_name: Field name
            min_length: Minimum length (inclusive)
            max_length: Maximum length (inclusive)

        Returns:
            Tuple of (is_valid, error_message)
        """
        if value is None:
            return True, ""

        value_str = str(value).strip()
        length = len(value_str)

        if min_length is not None and length < min_length:
            return False, f"Value in '{field_name}' is shorter than minimum {min_length}"

        if max_length is not None and length > max_length:
            return False, f"Value in '{field_name}' exceeds maximum {max_length}"

        return True, ""

    @staticmethod
    def validate_pattern(
        value: Any, field_name: str, pattern: str
    ) -> tuple[bool, str]:
        """
        Validate value matches regex pattern.

        Args:
            value: Value to validate
            field_name: Field name
            pattern: Regex pattern to match

        Returns:
            Tuple of (is_valid, error_message)
        """
        if value is None:
            return True, ""

        try:
            value_str = str(value).strip()
            if not re.match(pattern, value_str):
                return False, f"Value '{value}' in field '{field_name}' doesn't match pattern {pattern}"
            return True, ""
        except Exception as e:
            return False, f"Failed to validate pattern for '{field_name}': {e}"

    @staticmethod
    def validate_email(value: Any, field_name: str) -> tuple[bool, str]:
        """
        Validate email address format.

        Args:
            value: Email value
            field_name: Field name

        Returns:
            Tuple of (is_valid, error_message)
        """
        if value is None:
            return True, ""

        value_str = str(value).strip()
        email_pattern = r"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$"

        if not re.match(email_pattern, value_str):
            return False, f"Invalid email format in field '{field_name}': {value}"

        return True, ""

    @staticmethod
    def validate_row(
        row_data: dict[str, Any], schema: list[SchemaField]
    ) -> tuple[bool, list[str]]:
        """
        Validate entire row against schema.

        Args:
            row_data: Dictionary of field values
            schema: List of schema fields

        Returns:
            Tuple of (is_valid, list of error messages)
        """
        errors = []

        for field in schema:
            value = row_data.get(field.field_name)

            # Check required fields
            if field.required:
                is_valid, error = Validator.validate_required_field(value, field.field_name)
                if not is_valid:
                    errors.append(error)
                    continue

            # Check data type
            is_valid, error = Validator.validate_data_type(value, field.data_type, field.field_name)
            if not is_valid:
                errors.append(error)

        return len(errors) == 0, errors

    @staticmethod
    def validate_batch(
        rows: list[dict[str, Any]], schema: list[SchemaField]
    ) -> dict[int, list[str]]:
        """
        Validate multiple rows.

        Args:
            rows: List of row dictionaries
            schema: Schema fields

        Returns:
            Dictionary mapping row index to error lists (only rows with errors)
        """
        validation_results = {}

        for idx, row in enumerate(rows):
            is_valid, errors = Validator.validate_row(row, schema)
            if not is_valid:
                validation_results[idx] = errors

        return validation_results

    @staticmethod
    def normalize_value(value: Any, data_type: DataType) -> Any:
        """
        Normalize value to standard format.

        Args:
            value: Original value
            data_type: Target data type

        Returns:
            Normalized value
        """
        if value is None:
            return None

        value_str = str(value).strip()

        if data_type == DataType.TEXT:
            return value_str

        elif data_type == DataType.NUMBER:
            cleaned = re.sub(r"[$,]", "", value_str)
            return float(cleaned)

        elif data_type == DataType.CURRENCY:
            cleaned = re.sub(r"[$,]", "", value_str)
            return float(cleaned)

        elif data_type == DataType.DATE:
            for fmt in ["%Y-%m-%d", "%m/%d/%Y", "%d/%m/%Y", "%Y/%m/%d"]:
                try:
                    dt = datetime.strptime(value_str, fmt)
                    return dt.date()
                except ValueError:
                    continue
            raise ValidationError(f"Could not parse date: {value}")

        elif data_type == DataType.BOOLEAN:
            return value_str.lower() in ("yes", "true", "1")

        else:
            return value_str
