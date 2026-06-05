"""
AuditLog - Audit trail and event logging.

Responsibility:
- Record all workflow events and decisions
- Track data transformations and changes
- Maintain audit compliance
- Generate audit reports
"""

from datetime import datetime
from enum import Enum
from pathlib import Path
from typing import Any, Optional

from audit_framework.excel_utils import ExcelWorkbookManager


class EventType(str, Enum):
    """Types of audit events."""

    CONFIG_LOADED = "CONFIG_LOADED"
    TAGS_BUILT = "TAGS_BUILT"
    DATA_EXTRACTED = "DATA_EXTRACTED"
    VALIDATION_STARTED = "VALIDATION_STARTED"
    VALIDATION_PASSED = "VALIDATION_PASSED"
    VALIDATION_FAILED = "VALIDATION_FAILED"
    TRANSFORMATION_APPLIED = "TRANSFORMATION_APPLIED"
    QA_RULE_APPLIED = "QA_RULE_APPLIED"
    EXCEPTION_CREATED = "EXCEPTION_CREATED"
    EXCEPTION_ACCEPTED = "EXCEPTION_ACCEPTED"
    EXCEPTION_OVERRIDDEN = "EXCEPTION_OVERRIDDEN"
    EXCEPTION_REJECTED = "EXCEPTION_REJECTED"
    APPROVAL_REQUESTED = "APPROVAL_REQUESTED"
    APPROVAL_GRANTED = "APPROVAL_GRANTED"
    APPROVAL_DENIED = "APPROVAL_DENIED"
    DATA_EXPORTED = "DATA_EXPORTED"
    ERROR_OCCURRED = "ERROR_OCCURRED"
    OTHER = "OTHER"


class AuditLogEntry:
    """Single audit log entry."""

    def __init__(
        self,
        event_type: EventType,
        message: str,
        user: Optional[str] = None,
        module: Optional[str] = None,
        row_id: Optional[int] = None,
        field_name: Optional[str] = None,
        old_value: Optional[Any] = None,
        new_value: Optional[Any] = None,
        details: Optional[dict[str, Any]] = None,
    ) -> None:
        """
        Initialize audit log entry.

        Args:
            event_type: Type of event
            message: Human-readable description
            user: User who triggered event
            module: Module that generated event
            row_id: Row number (for data changes)
            field_name: Field affected
            old_value: Previous value
            new_value: New value
            details: Additional details
        """
        self.timestamp = datetime.now()
        self.event_type = event_type
        self.message = message
        self.user = user
        self.module = module
        self.row_id = row_id
        self.field_name = field_name
        self.old_value = old_value
        self.new_value = new_value
        self.details = details or {}

    def to_dict(self) -> dict[str, Any]:
        """Convert entry to dictionary."""
        return {
            "Timestamp": self.timestamp.isoformat(),
            "EventType": self.event_type.value,
            "Message": self.message,
            "User": self.user or "",
            "Module": self.module or "",
            "RowID": self.row_id or "",
            "FieldName": self.field_name or "",
            "OldValue": str(self.old_value) if self.old_value is not None else "",
            "NewValue": str(self.new_value) if self.new_value is not None else "",
            "Details": str(self.details) if self.details else "",
        }


class AuditLog:
    """Manages audit logging for engagement."""

    def __init__(self, workbook_path: str | Path) -> None:
        """
        Initialize AuditLog.

        Args:
            workbook_path: Path to engagement workbook
        """
        self.workbook_path = Path(workbook_path)
        self.excel_mgr = ExcelWorkbookManager(self.workbook_path)
        self.entries: list[AuditLogEntry] = []

    def log_event(
        self,
        event_type: EventType,
        message: str,
        user: Optional[str] = None,
        module: Optional[str] = None,
        **kwargs,
    ) -> None:
        """
        Log an event.

        Args:
            event_type: Type of event
            message: Event description
            user: User (optional)
            module: Module name (optional)
            **kwargs: Additional fields (row_id, field_name, old_value, new_value, etc)
        """
        entry = AuditLogEntry(
            event_type=event_type,
            message=message,
            user=user,
            module=module,
            **kwargs,
        )
        self.entries.append(entry)

    def log_config_loaded(self, engagement_id: str, user: Optional[str] = None) -> None:
        """Log configuration load."""
        self.log_event(
            EventType.CONFIG_LOADED,
            f"Configuration loaded for engagement {engagement_id}",
            user=user,
            module="ConfigManager",
        )

    def log_tags_built(self, tag_count: int, user: Optional[str] = None) -> None:
        """Log tag building."""
        self.log_event(
            EventType.TAGS_BUILT,
            f"Generated {tag_count} extraction tags",
            user=user,
            module="TagBuilder",
        )

    def log_data_extracted(self, row_count: int, user: Optional[str] = None) -> None:
        """Log data extraction."""
        self.log_event(
            EventType.DATA_EXTRACTED,
            f"Extracted data for {row_count} rows",
            user=user,
            module="Main",
        )

    def log_validation_passed(self, row_count: int, user: Optional[str] = None) -> None:
        """Log validation passed."""
        self.log_event(
            EventType.VALIDATION_PASSED,
            f"Validation passed for {row_count} rows",
            user=user,
            module="Validator",
        )

    def log_validation_failed(
        self, error_count: int, user: Optional[str] = None, details: Optional[dict] = None
    ) -> None:
        """Log validation failure."""
        self.log_event(
            EventType.VALIDATION_FAILED,
            f"Validation failed with {error_count} errors",
            user=user,
            module="Validator",
            details=details,
        )

    def log_qa_exception(
        self,
        exception_id: str,
        rule_id: str,
        field_name: str,
        value: Any,
        reason: str,
        user: Optional[str] = None,
    ) -> None:
        """Log QA exception."""
        self.log_event(
            EventType.EXCEPTION_CREATED,
            f"Exception {exception_id} created by rule {rule_id}",
            user=user,
            module="QAEngine",
            row_id=None,
            field_name=field_name,
            old_value=value,
            details={"exception_id": exception_id, "rule_id": rule_id, "reason": reason},
        )

    def log_exception_decision(
        self,
        exception_id: str,
        decision: str,  # ACCEPTED, OVERRIDDEN, REJECTED
        user: str,
        override_value: Optional[Any] = None,
        justification: Optional[str] = None,
    ) -> None:
        """Log QA exception decision."""
        event_type = {
            "ACCEPTED": EventType.EXCEPTION_ACCEPTED,
            "OVERRIDDEN": EventType.EXCEPTION_OVERRIDDEN,
            "REJECTED": EventType.EXCEPTION_REJECTED,
        }.get(decision, EventType.OTHER)

        self.log_event(
            event_type,
            f"Exception {exception_id} marked as {decision}",
            user=user,
            module="QAEngine",
            details={
                "exception_id": exception_id,
                "decision": decision,
                "override_value": override_value,
                "justification": justification,
            },
        )

    def log_approval_granted(
        self, approver: str, approval_level: int, comment: Optional[str] = None
    ) -> None:
        """Log approval granted."""
        self.log_event(
            EventType.APPROVAL_GRANTED,
            f"Approval granted by {approver} (Level {approval_level})",
            user=approver,
            module="Main",
            details={"approval_level": approval_level, "comment": comment},
        )

    def log_approval_denied(
        self, approver: str, approval_level: int, reason: str
    ) -> None:
        """Log approval denied."""
        self.log_event(
            EventType.APPROVAL_DENIED,
            f"Approval denied by {approver} (Level {approval_level}): {reason}",
            user=approver,
            module="Main",
            details={"approval_level": approval_level, "reason": reason},
        )

    def log_error(
        self,
        module: str,
        error_message: str,
        error_code: Optional[int] = None,
        user: Optional[str] = None,
    ) -> None:
        """Log error."""
        self.log_event(
            EventType.ERROR_OCCURRED,
            f"{module}: {error_message}",
            user=user,
            module=module,
            details={"error_code": error_code},
        )

    def write_to_sheet(self, sheet_name: str = "AUDIT_LOG") -> None:
        """
        Write audit log entries to Excel sheet.

        Args:
            sheet_name: Name of sheet to write to
        """
        if not self.excel_mgr.sheet_exists(sheet_name):
            self.excel_mgr.create_sheet(sheet_name)

        # Convert entries to dictionaries
        rows = [entry.to_dict() for entry in self.entries]

        # Get headers from first entry
        headers = list(rows[0].keys()) if rows else []

        # Write to sheet
        self.excel_mgr.write_rows_from_dicts(sheet_name, rows, headers, start_row=1)

    def get_entries(self, event_type: Optional[EventType] = None) -> list[AuditLogEntry]:
        """Get audit log entries (optionally filtered by type)."""
        if event_type:
            return [e for e in self.entries if e.event_type == event_type]
        return self.entries

    def get_summary(self) -> dict[str, Any]:
        """Get audit log summary."""
        summary: dict[str, int] = {}

        for entry in self.entries:
            event_key = entry.event_type.value
            summary[event_key] = summary.get(event_key, 0) + 1

        return {
            "total_events": len(self.entries),
            "event_types": summary,
            "first_event": self.entries[0].timestamp.isoformat() if self.entries else None,
            "last_event": self.entries[-1].timestamp.isoformat() if self.entries else None,
        }

    def save(self) -> None:
        """Save workbook with audit log."""
        self.write_to_sheet()
        self.excel_mgr.save()

    def close(self) -> None:
        """Close workbook."""
        self.excel_mgr.close()

    def __enter__(self) -> "AuditLog":
        """Context manager entry."""
        return self

    def __exit__(self, exc_type, exc_val, exc_tb) -> None:
        """Context manager exit."""
        self.close()
