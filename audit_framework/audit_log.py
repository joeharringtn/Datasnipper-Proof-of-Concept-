"""
AuditLog — immutable event trail for the entire engagement lifecycle.

Every significant action in the framework is recorded here: config loads,
tag builds, validation passes/failures, QA exception decisions, approval
grants/denials, and export events.  The log is persisted to an AUDIT_LOG
sheet in the engagement workbook so auditors have a full chain of custody.

Design decisions
-----------------
- In-memory list of AuditLogEntry objects; written to Excel only on save().
- No deduplication — every call to log_event() appends a new entry.
- Thread safety: not thread-safe; use one AuditLog per thread.
- The class is a context manager to ensure the workbook is closed on exit.

Usage
-----
    with AuditLog("engagement.xlsx") as log:
        log.log_config_loaded("ENG-001", user="j.smith")
        log.log_tags_built(42, user="j.smith")
        log.save()   # writes AUDIT_LOG sheet and saves the workbook
"""

from datetime import datetime
from enum import Enum
from pathlib import Path
from typing import Any, Optional

from audit_framework.excel_utils import ExcelWorkbookManager


class EventType(str, Enum):
    """
    Enum of all auditable event types in the engagement lifecycle.

    Using str as the base class ensures .value serializes to a plain string,
    making Excel/JSON output human-readable without extra conversion.
    """

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
    """
    A single timestamped event in the audit log.

    Instances are created by AuditLog.log_event() and should not be
    constructed directly by other modules.
    """

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
        Args:
            event_type:  Enum value classifying this event.
            message:     Human-readable description for audit reviewers.
            user:        Username who triggered the event (None for system events).
            module:      Framework module that generated the event.
            row_id:      Data row number, when the event relates to specific row.
            field_name:  Field name, when the event relates to a specific field.
            old_value:   Previous value, for change-tracking events.
            new_value:   New value, for change-tracking events.
            details:     Arbitrary extra context as a dict.
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
        """
        Serialize to a flat dict matching the AUDIT_LOG sheet column layout.

        None values are converted to empty strings so Excel cells are never
        written with the literal text "None".
        """
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
    """Manages the engagement audit trail and writes it to the workbook."""

    def __init__(self, workbook_path: str | Path) -> None:
        """
        Open the engagement workbook for audit log writing.

        The AUDIT_LOG sheet is created automatically by write_to_sheet() if
        it doesn't exist yet — no need to pre-create it.

        Args:
            workbook_path: Path to the engagement .xlsx / .xlsm file.
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
        Append a new event to the in-memory log.

        This is the primary entry point for all logging.  The convenience
        methods below (log_config_loaded, log_tags_built, etc.) are thin
        wrappers around this method that enforce a consistent message format.

        Args:
            event_type: Enum value classifying this event.
            message:    Human-readable description.
            user:       Username (optional; system events may omit this).
            module:     Originating module name.
            **kwargs:   Forwarded to AuditLogEntry (row_id, field_name,
                        old_value, new_value, details).
        """
        self.entries.append(
            AuditLogEntry(event_type=event_type, message=message, user=user, module=module, **kwargs)
        )

    # ---- Convenience wrappers ----

    def log_config_loaded(self, engagement_id: str, user: Optional[str] = None) -> None:
        """Log that the engagement configuration was successfully loaded."""
        self.log_event(
            EventType.CONFIG_LOADED,
            f"Configuration loaded for engagement {engagement_id}",
            user=user,
            module="ConfigManager",
        )

    def log_tags_built(self, tag_count: int, user: Optional[str] = None) -> None:
        """Log that DataSnipper extraction tags were generated."""
        self.log_event(
            EventType.TAGS_BUILT,
            f"Generated {tag_count} extraction tags",
            user=user,
            module="TagBuilder",
        )

    def log_data_extracted(self, row_count: int, user: Optional[str] = None) -> None:
        """Log that DataSnipper completed extraction for a set of rows."""
        self.log_event(
            EventType.DATA_EXTRACTED,
            f"Extracted data for {row_count} rows",
            user=user,
            module="Main",
        )

    def log_validation_passed(self, row_count: int, user: Optional[str] = None) -> None:
        """Log that schema validation passed for all rows."""
        self.log_event(
            EventType.VALIDATION_PASSED,
            f"Validation passed for {row_count} rows",
            user=user,
            module="Validator",
        )

    def log_validation_failed(
        self, error_count: int, user: Optional[str] = None, details: Optional[dict] = None
    ) -> None:
        """Log that schema validation found errors."""
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
        """Log creation of a new QA exception."""
        self.log_event(
            EventType.EXCEPTION_CREATED,
            f"Exception {exception_id} created by rule {rule_id}",
            user=user,
            module="QAEngine",
            field_name=field_name,
            old_value=value,
            details={"exception_id": exception_id, "rule_id": rule_id, "reason": reason},
        )

    def log_exception_decision(
        self,
        exception_id: str,
        decision: str,
        user: str,
        override_value: Optional[Any] = None,
        justification: Optional[str] = None,
    ) -> None:
        """
        Log a reviewer's decision on a QA exception.

        Args:
            decision: One of "ACCEPTED", "OVERRIDDEN", or "REJECTED".
        """
        # Map the decision string to the appropriate enum value
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
        """Log that an approver granted sign-off at the given workflow level."""
        self.log_event(
            EventType.APPROVAL_GRANTED,
            f"Approval granted by {approver} (Level {approval_level})",
            user=approver,
            module="Main",
            details={"approval_level": approval_level, "comment": comment},
        )

    def log_approval_denied(self, approver: str, approval_level: int, reason: str) -> None:
        """Log that an approver denied sign-off and why."""
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
        """Log an unexpected error from any module."""
        self.log_event(
            EventType.ERROR_OCCURRED,
            f"{module}: {error_message}",
            user=user,
            module=module,
            details={"error_code": error_code},
        )

    # ---- Persistence ----

    def write_to_sheet(self, sheet_name: str = "AUDIT_LOG") -> None:
        """
        Write all in-memory entries to an Excel sheet.

        Creates the sheet if it doesn't exist.  Overwrites from row 1 on
        each call — this is intentional; the authoritative log is the
        in-memory entries list, not whatever was previously in the sheet.

        Args:
            sheet_name: Target sheet name (default "AUDIT_LOG").
        """
        if not self.excel_mgr.sheet_exists(sheet_name):
            self.excel_mgr.create_sheet(sheet_name)

        rows = [entry.to_dict() for entry in self.entries]
        headers = list(rows[0].keys()) if rows else []
        self.excel_mgr.write_rows_from_dicts(sheet_name, rows, headers, start_row=1)

    def get_entries(self, event_type: Optional[EventType] = None) -> list[AuditLogEntry]:
        """
        Return log entries, optionally filtered by event type.

        Args:
            event_type: If provided, return only entries of this type.
                        If None, return all entries.
        """
        if event_type:
            return [e for e in self.entries if e.event_type == event_type]
        return self.entries

    def get_summary(self) -> dict[str, Any]:
        """
        Return a count-by-event-type summary of the log.

        Also includes ISO timestamps of the first and last events, useful
        for reporting engagement start/end times.
        """
        summary: dict[str, int] = {}
        for entry in self.entries:
            key = entry.event_type.value
            summary[key] = summary.get(key, 0) + 1

        return {
            "total_events": len(self.entries),
            "event_types": summary,
            "first_event": self.entries[0].timestamp.isoformat() if self.entries else None,
            "last_event": self.entries[-1].timestamp.isoformat() if self.entries else None,
        }

    def save(self) -> None:
        """Write the audit log to the AUDIT_LOG sheet and save the workbook."""
        self.write_to_sheet()
        self.excel_mgr.save()

    def close(self) -> None:
        """Close the workbook without saving."""
        self.excel_mgr.close()

    def __enter__(self) -> "AuditLog":
        return self

    def __exit__(self, exc_type, exc_val, exc_tb) -> None:
        self.close()
