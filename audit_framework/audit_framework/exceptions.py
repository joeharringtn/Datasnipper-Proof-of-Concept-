"""
Custom exceptions for audit framework.
"""


class AuditFrameworkError(Exception):
    """Base exception for all audit framework errors."""

    pass


class ConfigError(AuditFrameworkError):
    """Raised when configuration is invalid or incomplete."""

    pass


class WorksheetMissingError(ConfigError):
    """Raised when a required worksheet is missing from the workbook."""

    pass


class ValidationError(AuditFrameworkError):
    """Raised when data validation fails."""

    pass


class ExtractionError(AuditFrameworkError):
    """Raised when extraction/snipping fails."""

    pass


class TagBuildError(AuditFrameworkError):
    """Raised when tag building fails."""

    pass


class DataMapperError(AuditFrameworkError):
    """Raised when data mapping/transformation fails."""

    pass


class QARuleError(AuditFrameworkError):
    """Raised when QA rule application fails."""

    pass


class ApprovalError(AuditFrameworkError):
    """Raised when approval workflow fails."""

    pass


class WorkflowStateError(AuditFrameworkError):
    """Raised when workflow state transition is invalid."""

    pass
