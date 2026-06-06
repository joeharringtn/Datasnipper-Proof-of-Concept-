"""
Custom exception hierarchy for the audit framework.

All framework exceptions inherit from AuditFrameworkError, which lets callers
catch the entire framework's error surface with a single clause, or be
selective by catching a specific subtype.

Hierarchy:
    AuditFrameworkError
    ├── ConfigError               — bad/missing engagement config
    │   └── WorksheetMissingError — a required Excel sheet doesn't exist
    ├── ValidationError           — extracted data fails type/format checks
    ├── ExtractionError           — DataSnipper snipping operation failed
    ├── TagBuildError             — tag string construction failed
    ├── DataMapperError           — raw-value → schema-field mapping failed
    ├── QARuleError               — QA rule could not be evaluated
    ├── ApprovalError             — approval workflow action failed
    └── WorkflowStateError        — illegal state transition attempted
"""


class AuditFrameworkError(Exception):
    """Base exception for all audit framework errors."""
    pass


class ConfigError(AuditFrameworkError):
    """
    Raised when engagement configuration is invalid or incomplete.

    Covers missing required keys in the CONFIG sheet, invalid enum values,
    failed Pydantic model construction, and date-range violations.
    """
    pass


class WorksheetMissingError(ConfigError):
    """
    Raised when a required worksheet is absent from the workbook.

    Subclasses ConfigError so callers that catch ConfigError also catch this.
    Thrown early in ConfigManager.__init__ to fail fast before any I/O.
    """
    pass


class ValidationError(AuditFrameworkError):
    """
    Raised when extracted data fails schema-level type or format checks.

    Distinct from QARuleError: Validator enforces structural correctness
    (e.g., "this field must be a number"), while QAEngine enforces business
    rules (e.g., "amount must be > 0").
    """
    pass


class ExtractionError(AuditFrameworkError):
    """
    Raised when a DataSnipper extraction operation fails.

    Typically wraps lower-level errors from the DataSnipper API or file I/O
    when reading source documents.
    """
    pass


class TagBuildError(AuditFrameworkError):
    """
    Raised when a DataSnipper tag string cannot be constructed.

    Common causes: missing required fields (e.g., search_keywords for a
    DS_SEARCH tag), or an unknown ExtractionMethod enum value.
    """
    pass


class DataMapperError(AuditFrameworkError):
    """
    Raised when raw extracted values cannot be mapped to output schema fields.

    Wraps normalization failures (e.g., a currency string that can't be
    parsed) and reports the offending field name.
    """
    pass


class QARuleError(AuditFrameworkError):
    """
    Raised when a QA rule cannot be evaluated against a data value.

    Distinct from a rule *violation* (which produces a QAException, not an
    error). This is thrown when the rule definition itself is malformed —
    e.g., a RANGE rule with a non-numeric boundary.
    """
    pass


class ApprovalError(AuditFrameworkError):
    """
    Raised when an approval workflow action cannot be completed.

    Examples: approver not configured, required fields missing, attempting
    to approve when blocking exceptions are still pending.
    """
    pass


class WorkflowStateError(AuditFrameworkError):
    """
    Raised when an illegal workflow state transition is attempted.

    The engagement lifecycle follows a strict sequence (config → extract →
    validate → QA → approve → export). This exception fires when a caller
    tries to skip a step or move backwards.
    """
    pass
