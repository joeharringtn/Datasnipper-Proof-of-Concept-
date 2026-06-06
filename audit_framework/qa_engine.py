"""
QAEngine — applies business-rule QA checks to extracted data and manages exceptions.

After DataMapper normalizes raw extracted values into the output schema, QAEngine
runs the configured QARules against those values.  A rule violation creates a
QAException that must be reviewed (accepted, overridden, or rejected) before the
engagement can proceed to the approval stage.

Rule evaluation flow
---------------------
    for each row:
        for each QARule whose field_name is in the row:
            apply_rule(rule, value, full_row)
                → None           (rule passed)
                → QAException    (rule violated; stored in engine state)

Exception lifecycle
-------------------
    PENDING   → ACCEPTED   (reviewer accepts the anomaly as-is)
    PENDING   → OVERRIDDEN (reviewer provides a corrected value)
    PENDING   → REJECTED   (reviewer marks the row unusable)

    can_proceed_to_approval() returns True only when no PENDING exceptions remain.
    get_blocking_exceptions() returns PENDING exceptions with fail_action=BLOCK,
    which should prevent output from being exported regardless of approval status.

Performance note
-----------------
Exceptions are stored in both a list (for ordered iteration) and a dict
(_exception_index) for O(1) lookup by exception_id.  For engagements with
thousands of exceptions, accept/override/reject calls no longer do a full scan.
"""

import re
from typing import Any, Optional

from audit_framework.exceptions import QARuleError
from audit_framework.models import FailAction, QARule, RuleType, SchemaField, Severity
from audit_framework.validator import Validator


class QAException:
    """
    A single QA rule violation detected during processing.

    Instances are created by QAEngine.apply_rule() and should not be
    constructed directly.  The exception_id format is "EXC_NNNNNN" where N
    is a zero-padded counter maintained by the engine.
    """

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
        self.exception_id = exception_id
        self.rule_id = rule_id
        self.field_name = field_name
        self.extracted_value = extracted_value
        self.reason = reason
        self.severity = severity
        self.fail_action = fail_action
        self.status = "PENDING"   # PENDING | ACCEPTED | OVERRIDDEN | REJECTED
        self.override_value: Optional[Any] = None
        self.override_justification: str = ""
        self.reviewed_by: Optional[str] = None
        self.reviewed_at: Optional[str] = None

    def to_dict(self) -> dict[str, Any]:
        """Serialize to a flat dict suitable for writing to an Excel sheet."""
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
    """Applies QA rules to extracted data and manages the resulting exceptions."""

    def __init__(self) -> None:
        """
        Initialise a fresh QA engine with no exceptions.

        One QAEngine instance per engagement run; do not reuse across
        different workbooks or datasets.
        """
        self.exceptions: list[QAException] = []
        # Parallel dict for O(1) lookup by exception_id (see module docstring)
        self._exception_index: dict[str, QAException] = {}
        self.exception_counter = 0

    def apply_rule(
        self, rule: QARule, value: Any, row_data: dict[str, Any]
    ) -> Optional[QAException]:
        """
        Evaluate a single QA rule against one field value.

        The rule_definition string is interpreted according to rule.rule_type —
        see RuleType docstring in models.py for the encoding of each type.

        Args:
            rule:      The rule to evaluate.
            value:     The field value extracted from the document.
            row_data:  The complete row dict; required for CROSS_FIELD rules
                       which compare two fields against each other.

        Returns:
            A QAException if the rule is violated, None if it passes.
            CUSTOM rules always return None (not yet implemented).

        Raises:
            QARuleError: If the rule definition is malformed (e.g. non-numeric
                         RANGE boundary) — distinct from a rule *violation*.
        """
        try:
            # Skip None values — missing data is a Validator concern, not QA
            if value is None:
                return None

            is_violated = False
            reason = ""

            if rule.rule_type == RuleType.RANGE:
                # rule_definition format: "min:max"  (e.g. "0:1000000")
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
                # rule_definition is the full regex pattern
                is_valid, error = Validator.validate_pattern(
                    value, rule.field_name, rule.rule_definition
                )
                is_violated = not is_valid
                reason = error or ""

            elif rule.rule_type == RuleType.LOOKUP:
                # rule_definition format: "val1,val2,val3"
                valid_values = [v.strip() for v in rule.rule_definition.split(",")]
                if str(value).strip() not in valid_values:
                    is_violated = True
                    reason = f"Value '{value}' not in allowed list: {valid_values}"

            elif rule.rule_type == RuleType.DUPLICATE:
                # rule_definition: name of a second field to compare against.
                # Flags when the target field holds the same value as this one —
                # useful for catching copy-paste errors between related fields.
                dup_field = rule.rule_definition.strip()
                if dup_field in row_data and row_data[dup_field] == value:
                    is_violated = True
                    reason = f"Duplicate value found in field {dup_field}"

            elif rule.rule_type == RuleType.CROSS_FIELD:
                # rule_definition format: "field1:operator:field2"
                # Supported operators: "before" (field1 < field2), "after" (field1 > field2)
                # Example: "InvoiceDate:before:DueDate"
                try:
                    parts = rule.rule_definition.split(":")
                    if len(parts) >= 3:
                        field1, operator, field2 = parts[0].strip(), parts[1].strip().lower(), parts[2].strip()

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
                # Custom rules require bespoke logic not expressible in the
                # rule_definition string.  Skip silently — callers can extend
                # QAEngine to handle custom types before calling apply_rule().
                return None

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
                # Index by ID for O(1) lookup in accept/override/reject methods
                self._exception_index[exception.exception_id] = exception
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
        Apply all rules whose field_name is present in the row.

        Rules targeting fields not in row_data are silently skipped rather
        than erroring — this handles schema mismatches gracefully.

        Args:
            rules:    Full list of QARule objects for this engagement.
            row_data: One mapped data row (field_name → value).

        Returns:
            List of QAException objects generated for this row (may be empty).
        """
        return [
            exc
            for rule in rules
            if rule.field_name in row_data
            for exc in [self.apply_rule(rule, row_data[rule.field_name], row_data)]
            if exc is not None
        ]

    def apply_rules_to_batch(
        self, rules: list[QARule], rows: list[dict[str, Any]]
    ) -> dict[int, list[QAException]]:
        """
        Apply rules to every row in a batch.

        Args:
            rules: QA rules to evaluate.
            rows:  List of mapped data rows.

        Returns:
            Dict mapping row index (0-based) → list of exceptions.
            Rows with no violations are absent from the dict.
        """
        return {
            idx: exceptions
            for idx, row in enumerate(rows)
            for exceptions in [self.apply_rules_to_row(rules, row)]
            if exceptions
        }

    def accept_exception(self, exception_id: str, reviewed_by: str) -> None:
        """
        Mark an exception as accepted (anomaly acknowledged, data used as-is).

        Args:
            exception_id: The EXC_NNNNNN identifier.
            reviewed_by:  Username of the reviewer.
        """
        exc = self._find_exception(exception_id)
        if exc:
            exc.status = "ACCEPTED"
            exc.reviewed_by = reviewed_by
            exc.reviewed_at = self._get_current_timestamp()

    def override_exception(
        self, exception_id: str, override_value: Any, justification: str, reviewed_by: str
    ) -> None:
        """
        Override an exception with a corrected value and justification.

        The original extracted_value is preserved for audit purposes; the
        override_value is what flows into the final output.

        Args:
            exception_id:   The EXC_NNNNNN identifier.
            override_value: The corrected replacement value.
            justification:  Business reason for the override.
            reviewed_by:    Username of the reviewer.
        """
        exc = self._find_exception(exception_id)
        if exc:
            exc.status = "OVERRIDDEN"
            exc.override_value = override_value
            exc.override_justification = justification
            exc.reviewed_by = reviewed_by
            exc.reviewed_at = self._get_current_timestamp()

    def reject_exception(self, exception_id: str, reason: str, reviewed_by: str) -> None:
        """
        Reject an exception (the associated data row is unusable).

        Args:
            exception_id: The EXC_NNNNNN identifier.
            reason:       Explanation of why the data is being rejected.
            reviewed_by:  Username of the reviewer.
        """
        exc = self._find_exception(exception_id)
        if exc:
            exc.status = "REJECTED"
            exc.override_justification = reason
            exc.reviewed_by = reviewed_by
            exc.reviewed_at = self._get_current_timestamp()

    def get_pending_exceptions(self) -> list[QAException]:
        """Return all exceptions that have not yet been reviewed."""
        return [e for e in self.exceptions if e.status == "PENDING"]

    def get_reviewed_exceptions(self) -> list[QAException]:
        """Return all exceptions that have been accepted, overridden, or rejected."""
        return [e for e in self.exceptions if e.status in ("ACCEPTED", "OVERRIDDEN", "REJECTED")]

    def get_exceptions_by_severity(self, severity: Severity) -> list[QAException]:
        """Return all exceptions matching the given severity level."""
        return [e for e in self.exceptions if e.severity == severity]

    def get_blocking_exceptions(self) -> list[QAException]:
        """
        Return PENDING exceptions with fail_action=BLOCK.

        These must be resolved before the engagement can be exported,
        regardless of whether the approval workflow has been completed.
        """
        return [
            e for e in self.exceptions
            if e.fail_action == FailAction.BLOCK and e.status == "PENDING"
        ]

    def can_proceed_to_approval(self) -> bool:
        """Return True only when every exception has been reviewed."""
        return len(self.get_pending_exceptions()) == 0

    def to_dict(self) -> dict[str, Any]:
        """
        Serialize all exceptions and a summary to a dict.

        Suitable for JSON export or writing to an Excel audit sheet.
        """
        return {
            "exceptions": [e.to_dict() for e in self.exceptions],
            "summary": {
                "total": len(self.exceptions),
                "pending": len(self.get_pending_exceptions()),
                "accepted": sum(1 for e in self.exceptions if e.status == "ACCEPTED"),
                "overridden": sum(1 for e in self.exceptions if e.status == "OVERRIDDEN"),
                "rejected": sum(1 for e in self.exceptions if e.status == "REJECTED"),
            },
        }

    # ========== Private Helpers ==========

    def _find_exception(self, exception_id: str) -> Optional[QAException]:
        """
        Look up an exception by ID using the O(1) index.

        Returns None (rather than raising) if the ID doesn't exist — callers
        that need to distinguish "found but pending" from "not found" should
        check the return value.
        """
        return self._exception_index.get(exception_id)

    @staticmethod
    def _get_current_timestamp() -> str:
        """Return the current UTC time as an ISO-8601 string."""
        from datetime import datetime
        return datetime.now().isoformat()
