"""
QAEngine - Quality assurance rule application and exception management.

Responsibility:
- Apply QA rules to extracted data
- Flag exceptions and violations
- Track exception status (pending, accepted, overridden, rejected)
- Generate QA reports
"""

import re
from typing import Any, Optional

from audit_framework.exceptions import QARuleError
from audit_framework.models import FailAction, QARule, RuleType, SchemaField, Severity
from audit_framework.validator import Validator


class QAException:
    """Represents a QA exception (rule violation)."""

    def __init__(
        self,
        exception_id: str,
        rule_id: str,
        field_name: str,
        extracted_value: Any,
        reason: str,
        severity: Severity,
        fail_action: FailAction,
    ) -> None:
        """
        Initialize QA exception.

        Args:
            exception_id: Unique exception identifier
            rule_id: Rule that triggered exception
            field_name: Field with exception
            extracted_value: Original extracted value
            reason: Description of exception
            severity: Exception severity
            fail_action: Action on failure (FLAG, BLOCK, WARN)
        """
        self.exception_id = exception_id
        self.rule_id = rule_id
        self.field_name = field_name
        self.extracted_value = extracted_value
        self.reason = reason
        self.severity = severity
        self.fail_action = fail_action
        self.status = "PENDING"  # PENDING, ACCEPTED, OVERRIDDEN, REJECTED
        self.override_value: Optional[Any] = None
        self.override_justification: str = ""
        self.reviewed_by: Optional[str] = None
        self.reviewed_at: Optional[str] = None

    def to_dict(self) -> dict[str, Any]:
        """Convert exception to dictionary."""
        return {
            "ExceptionID": self.exception_id,
            "RuleID": self.rule_id,
            "FieldName": self.field_name,
            "ExtractedValue": self.extracted_value,
            "Reason": self.reason,
            "Severity": self.severity.value,
            "FailAction": self.fail_action.value,
            "Status": self.status,
            "OverrideValue": self.override_value,
            "OverrideJustification": self.override_justification,
            "ReviewedBy": self.reviewed_by,
            "ReviewedAt": self.reviewed_at,
        }


class QAEngine:
    """Applies QA rules and manages exceptions."""

    def __init__(self) -> None:
        """Initialize QA engine."""
        self.exceptions: list[QAException] = []
        self.exception_counter = 0

    def apply_rule(
        self, rule: QARule, value: Any, row_data: dict[str, Any]
    ) -> Optional[QAException]:
        """
        Apply single QA rule to a value.

        Args:
            rule: QA rule to apply
            value: Value to validate
            row_data: Complete row for cross-field rules

        Returns:
            QAException if rule violated, None if passed

        Raises:
            QARuleError: If rule application fails
        """
        try:
            if value is None:
                return None

            is_violated = False
            reason = ""

            if rule.rule_type == RuleType.RANGE:
                # Rule definition format: "min:max"
                try:
                    parts = rule.rule_definition.split(":")
                    min_val = float(parts[0]) if len(parts) > 0 else None
                    max_val = float(parts[1]) if len(parts) > 1 else None

                    is_valid, error = Validator.validate_number_range(
                        value, rule.field_name, min_val, max_val
                    )
                    is_violated = not is_valid
                    reason = error or ""
                except Exception as e:
                    raise QARuleError(f"Invalid RANGE rule definition: {e}") from e

            elif rule.rule_type == RuleType.FORMAT:
                # Rule definition is a regex pattern
                is_valid, error = Validator.validate_pattern(
                    value, rule.field_name, rule.rule_definition
                )
                is_violated = not is_valid
                reason = error or ""

            elif rule.rule_type == RuleType.LOOKUP:
                # Rule definition format: "value1,value2,value3"
                valid_values = [v.strip() for v in rule.rule_definition.split(",")]
                if str(value).strip() not in valid_values:
                    is_violated = True
                    reason = f"Value '{value}' not in allowed list: {valid_values}"

            elif rule.rule_type == RuleType.DUPLICATE:
                # Rule definition format: "field_name"
                # Check if this value already exists in row_data
                dup_field = rule.rule_definition.strip()
                if dup_field in row_data and row_data[dup_field] == value:
                    is_violated = True
                    reason = f"Duplicate value found in field {dup_field}"

            elif rule.rule_type == RuleType.CROSS_FIELD:
                # Rule definition format: "field1:operator:field2"
                # e.g., "StartDate:before:EndDate"
                try:
                    parts = rule.rule_definition.split(":")
                    if len(parts) >= 3:
                        field1 = parts[0].strip()
                        operator = parts[1].strip().lower()
                        field2 = parts[2].strip()

                        val1 = row_data.get(field1)
                        val2 = row_data.get(field2)

                        if val1 and val2:
                            if operator == "before" and val1 >= val2:
                                is_violated = True
                                reason = f"{field1} must be before {field2}"
                            elif operator == "after" and val1 <= val2:
                                is_violated = True
                                reason = f"{field1} must be after {field2}"
                except Exception as e:
                    raise QARuleError(f"Invalid CROSS_FIELD rule: {e}") from e

            elif rule.rule_type == RuleType.CUSTOM:
                # Custom rules would require custom logic
                # For now, skip with warning
                return None

            # Create exception if rule violated
            if is_violated:
                self.exception_counter += 1
                exception = QAException(
                    exception_id=f"EXC_{self.exception_counter:06d}",
                    rule_id=rule.rule_id,
                    field_name=rule.field_name,
                    extracted_value=value,
                    reason=reason,
                    severity=rule.severity,
                    fail_action=rule.fail_action,
                )
                self.exceptions.append(exception)
                return exception

            return None

        except QARuleError:
            raise
        except Exception as e:
            raise QARuleError(f"Failed to apply rule {rule.rule_id}: {e}") from e

    def apply_rules_to_row(
        self, rules: list[QARule], row_data: dict[str, Any]
    ) -> list[QAException]:
        """
        Apply all applicable rules to a row.

        Args:
            rules: List of QA rules
            row_data: Data row

        Returns:
            List of exceptions found
        """
        row_exceptions = []

        for rule in rules:
            if rule.field_name in row_data:
                exception = self.apply_rule(rule, row_data[rule.field_name], row_data)
                if exception:
                    row_exceptions.append(exception)

        return row_exceptions

    def apply_rules_to_batch(
        self, rules: list[QARule], rows: list[dict[str, Any]]
    ) -> dict[int, list[QAException]]:
        """
        Apply rules to multiple rows.

        Args:
            rules: List of QA rules
            rows: List of data rows

        Returns:
            Dictionary mapping row index to exceptions
        """
        batch_exceptions: dict[int, list[QAException]] = {}

        for idx, row in enumerate(rows):
            exceptions = self.apply_rules_to_row(rules, row)
            if exceptions:
                batch_exceptions[idx] = exceptions

        return batch_exceptions

    def accept_exception(self, exception_id: str, reviewed_by: str) -> None:
        """
        Accept exception as-is.

        Args:
            exception_id: Exception identifier
            reviewed_by: User accepting exception
        """
        exc = self._find_exception(exception_id)
        if exc:
            exc.status = "ACCEPTED"
            exc.reviewed_by = reviewed_by
            exc.reviewed_at = str(self._get_current_timestamp())

    def override_exception(
        self, exception_id: str, override_value: Any, justification: str, reviewed_by: str
    ) -> None:
        """
        Override exception with corrected value.

        Args:
            exception_id: Exception identifier
            override_value: Corrected value
            justification: Reason for override
            reviewed_by: User making override
        """
        exc = self._find_exception(exception_id)
        if exc:
            exc.status = "OVERRIDDEN"
            exc.override_value = override_value
            exc.override_justification = justification
            exc.reviewed_by = reviewed_by
            exc.reviewed_at = str(self._get_current_timestamp())

    def reject_exception(self, exception_id: str, reason: str, reviewed_by: str) -> None:
        """
        Reject exception (data unusable).

        Args:
            exception_id: Exception identifier
            reason: Rejection reason
            reviewed_by: User rejecting exception
        """
        exc = self._find_exception(exception_id)
        if exc:
            exc.status = "REJECTED"
            exc.override_justification = reason
            exc.reviewed_by = reviewed_by
            exc.reviewed_at = str(self._get_current_timestamp())

    def get_pending_exceptions(self) -> list[QAException]:
        """Get all pending exceptions."""
        return [e for e in self.exceptions if e.status == "PENDING"]

    def get_reviewed_exceptions(self) -> list[QAException]:
        """Get all reviewed exceptions."""
        return [e for e in self.exceptions if e.status in ("ACCEPTED", "OVERRIDDEN", "REJECTED")]

    def get_exceptions_by_severity(self, severity: Severity) -> list[QAException]:
        """Get exceptions by severity level."""
        return [e for e in self.exceptions if e.severity == severity]

    def get_blocking_exceptions(self) -> list[QAException]:
        """Get exceptions that should block processing."""
        return [
            e
            for e in self.exceptions
            if e.fail_action == FailAction.BLOCK and e.status == "PENDING"
        ]

    def can_proceed_to_approval(self) -> bool:
        """Check if all exceptions have been reviewed."""
        return len(self.get_pending_exceptions()) == 0

    def to_dict(self) -> dict[str, Any]:
        """Convert all exceptions to dictionary format."""
        return {
            "exceptions": [e.to_dict() for e in self.exceptions],
            "summary": {
                "total": len(self.exceptions),
                "pending": len(self.get_pending_exceptions()),
                "accepted": len([e for e in self.exceptions if e.status == "ACCEPTED"]),
                "overridden": len([e for e in self.exceptions if e.status == "OVERRIDDEN"]),
                "rejected": len([e for e in self.exceptions if e.status == "REJECTED"]),
            },
        }

    # ========== Private Methods ==========

    def _find_exception(self, exception_id: str) -> Optional[QAException]:
        """Find exception by ID."""
        for exc in self.exceptions:
            if exc.exception_id == exception_id:
                return exc
        return None

    @staticmethod
    def _get_current_timestamp() -> str:
        """Get current timestamp."""
        from datetime import datetime

        return datetime.now().isoformat()
