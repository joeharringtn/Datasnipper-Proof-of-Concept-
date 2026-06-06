"""
ConfigManager — loads, validates, and vends the engagement configuration.

This is the first module invoked in every workflow step.  It reads the
CONFIG sheet for top-level metadata and then reads the TAG_ENGINE, QA, and
OUTPUT sheets to populate the full ConfigObject used downstream.

Expected Excel workbook structure
----------------------------------
CONFIG sheet  — key/value pairs (col A = key, col B = value). Required keys:
                  EngagementID, EngagementType, PeriodStartDate, PeriodEndDate,
                  LeadAuditor, ClientName.
                Optional keys (None if absent, never raises):
                  ClientContact, FrameworkVersion, WorkflowType,
                  Level1Approver, Level1Email, Level2Approver, Level2Email.

TAG_ENGINE    — tabular; headers in row 1, one tag definition per data row.
QA            — tabular; headers in row 1, one rule per data row.
OUTPUT        — headers in row 1 define the output schema column names.

Usage example
-------------
    with ConfigManager("engagement.xlsx") as mgr:
        config = mgr.load_config()
        is_valid, errors = mgr.validate_config()
        if not is_valid:
            raise SystemExit(errors)
        tags = mgr.get_tags()
"""

from datetime import datetime
from pathlib import Path
from typing import Optional

from audit_framework.exceptions import ConfigError, WorksheetMissingError
from audit_framework.excel_utils import ExcelWorkbookManager
from audit_framework.models import (
    ApprovalWorkflow,
    ConfigObject,
    DataType,
    DocumentSpec,
    DocumentType,
    EngagementType,
    ExtractionMethod,
    QARule,
    SchemaField,
    TagDefinition,
    WorkflowType,
)


class ConfigManager:
    """Manages engagement configuration loading and validation."""

    def __init__(self, workbook_path: str | Path) -> None:
        """
        Open the workbook and verify the CONFIG sheet is present.

        Fails fast here rather than midway through load_config() — if the
        workbook doesn't have a CONFIG sheet, nothing else can work.

        Args:
            workbook_path: Path to the engagement .xlsx / .xlsm file.

        Raises:
            FileNotFoundError:      If the workbook file doesn't exist.
            WorksheetMissingError:  If the CONFIG sheet is absent.
        """
        self.workbook_path = Path(workbook_path)
        self.excel_mgr = ExcelWorkbookManager(self.workbook_path)
        self.config: Optional[ConfigObject] = None
        self.load_timestamp: Optional[datetime] = None

        if not self.excel_mgr.sheet_exists("CONFIG"):
            raise WorksheetMissingError(
                "CONFIG sheet is missing. Create a sheet named CONFIG with engagement settings."
            )

    def load_config(self) -> ConfigObject:
        """
        Read the workbook and construct a validated ConfigObject.

        Required CONFIG keys are fetched via _read_config_value(), which
        raises ConfigError on a missing key.  Optional keys are fetched
        directly via excel_mgr.read_config_value(), which returns None.

        Returns:
            A fully validated ConfigObject ready for downstream use.

        Raises:
            ConfigError: On missing required keys or Pydantic validation failure.
        """
        try:
            self.load_timestamp = datetime.now()

            # --- Required fields (raise ConfigError if absent) ---
            engagement_id = self._read_required_config_value("EngagementID")
            engagement_type = self._read_required_config_value("EngagementType")
            period_start_date = self._read_required_config_value("PeriodStartDate")
            period_end_date = self._read_required_config_value("PeriodEndDate")
            lead_auditor = self._read_required_config_value("LeadAuditor")
            client_name = self._read_required_config_value("ClientName")

            # --- Optional fields (None if absent, never raise) ---
            client_contact = self.excel_mgr.read_config_value("ClientContact")
            framework_version = self.excel_mgr.read_config_value("FrameworkVersion") or "1.0"

            # --- Sheet-based sub-configs ---
            documents = self._load_documents()
            tags = self._load_tags()
            qa_rules = self._load_qa_rules()
            output_schema = self._load_output_schema()
            approval_workflow = self._load_approval_workflow()

            # Pydantic validates types, date ordering, etc.
            self.config = ConfigObject(
                engagement_id=engagement_id,
                engagement_type=EngagementType(engagement_type),
                period_start_date=period_start_date,
                period_end_date=period_end_date,
                lead_auditor=lead_auditor,
                client_name=client_name,
                client_contact=client_contact,
                framework_version=framework_version,
                documents=documents,
                tags=tags,
                qa_rules=qa_rules,
                output_schema=output_schema,
                approval_workflow=approval_workflow,
            )

            return self.config

        except ValueError as e:
            # Pydantic raises ValueError; wrap with context
            raise ConfigError(f"Configuration validation failed: {e}") from e
        except ConfigError:
            raise
        except Exception as e:
            raise ConfigError(f"Failed to load configuration: {e}") from e

    def validate_config(self) -> tuple[bool, list[str]]:
        """
        Run post-load business-rule checks on the loaded ConfigObject.

        Pydantic already catches type errors during load_config(); this method
        enforces higher-level rules (e.g. approval workflow completeness) that
        don't fit neatly into field validators.

        Returns:
            (True, [])             if all checks pass.
            (False, [error, ...])  with a human-readable list of failures.

        Note:
            Document presence is NOT checked here because _load_documents() is
            not yet implemented — see that method's docstring.
        """
        if not self.config:
            return False, ["Configuration not loaded"]

        errors = []

        if not self.config.engagement_id:
            errors.append("EngagementID is required")

        if not self.config.engagement_type:
            errors.append("EngagementType is required")

        # BUG FIX: the previous expression was a Python chained comparison
        #   `start <= end is False`  →  `(start <= end) and (end is False)`
        # which is always False (a date is never `is False`).  The intent
        # is to flag configs where the period is inverted.
        if not (self.config.period_start_date <= self.config.period_end_date):
            errors.append(
                f"PeriodStartDate ({self.config.period_start_date}) must be before "
                f"PeriodEndDate ({self.config.period_end_date})"
            )

        if not self.config.lead_auditor:
            errors.append("LeadAuditor is recommended for audit traceability")

        if not self.config.client_name:
            errors.append("ClientName is required")

        if not self.config.tags:
            errors.append("At least one tag definition is required")

        # MULTI_LEVEL workflow requires at least one named approver
        if (
            self.config.approval_workflow.workflow_type == WorkflowType.MULTI_LEVEL
            and not self.config.approval_workflow.level1_approver
        ):
            errors.append("Approval workflow requires Level1Approver")

        return len(errors) == 0, errors

    def is_loaded(self) -> bool:
        """Return True if load_config() has been successfully called."""
        return self.config is not None

    def get_config(self) -> ConfigObject:
        """
        Return the loaded ConfigObject.

        Raises:
            ConfigError: If load_config() has not been called yet.
        """
        if not self.config:
            raise ConfigError("Configuration not loaded. Call load_config() first.")
        return self.config

    def get_engagement_type(self) -> EngagementType:
        """Return the engagement type enum value."""
        if not self.config:
            raise ConfigError("Configuration not loaded")
        return self.config.engagement_type

    def get_documents(self) -> list[DocumentSpec]:
        """Return the list of document specifications."""
        if not self.config:
            raise ConfigError("Configuration not loaded")
        return self.config.documents

    def get_tags(self) -> list[TagDefinition]:
        """Return the list of tag definitions from the TAG_ENGINE sheet."""
        if not self.config:
            raise ConfigError("Configuration not loaded")
        return self.config.tags

    def get_qa_rules(self) -> list[QARule]:
        """Return the list of QA rules from the QA sheet."""
        if not self.config:
            raise ConfigError("Configuration not loaded")
        return self.config.qa_rules

    def get_output_schema(self) -> list[SchemaField]:
        """Return the output schema inferred from the OUTPUT sheet headers."""
        if not self.config:
            raise ConfigError("Configuration not loaded")
        return self.config.output_schema

    def get_approval_workflow(self) -> ApprovalWorkflow:
        """Return the approval workflow configuration."""
        if not self.config:
            raise ConfigError("Configuration not loaded")
        return self.config.approval_workflow

    # ========== Private Helpers ==========

    def _read_required_config_value(self, key: str) -> str:
        """
        Read a required key from the CONFIG sheet.

        Args:
            key: CONFIG sheet key name (case-insensitive).

        Returns:
            The value as a stripped string.

        Raises:
            ConfigError: If the key is not present in the CONFIG sheet.
        """
        value = self.excel_mgr.read_config_value(key)
        if value is None:
            raise ConfigError(f"Required configuration '{key}' not found in CONFIG sheet")
        return str(value).strip()

    def _load_documents(self) -> list[DocumentSpec]:
        """
        Load document specifications.

        TODO: Document specs are not yet stored in a dedicated sheet.  The
        VBA version embedded them in the CONFIG sheet as a multi-row block,
        but no equivalent Python reader has been implemented.  For now this
        returns an empty list, and validate_config() does not enforce a
        non-empty documents list so that the rest of the workflow can proceed.

        When implemented, this should read from a DOCUMENTS sheet with headers:
            DocID, FilePath, PageCount, DocType, ExtractionMethod
        """
        return []

    def _load_tags(self) -> list[TagDefinition]:
        """
        Read tag definitions from the TAG_ENGINE sheet.

        Each row in TAG_ENGINE corresponds to one TagDefinition.  Rows where
        TagID is empty are skipped (handles trailing blank rows).

        Returns:
            List of TagDefinition objects; empty list if TAG_ENGINE doesn't exist.

        Raises:
            ConfigError: If a row is malformed (e.g. invalid ExtractionMethod).
        """
        tags: list[TagDefinition] = []

        if not self.excel_mgr.sheet_exists("TAG_ENGINE"):
            return tags

        try:
            tag_rows = self.excel_mgr.read_rows_as_dicts("TAG_ENGINE")

            for row in tag_rows:
                if not row.get("TagID"):
                    continue

                extraction_method = ExtractionMethod(
                    row.get("ExtractionMethod", "DS_SEARCH")
                )

                # Gracefully fall back to TEXT for unrecognised type strings
                field_type_str = row.get("FieldType", "Text")
                try:
                    field_type = DataType(field_type_str)
                except ValueError:
                    field_type = DataType.TEXT

                # Accept common truthy strings from Excel cells
                required_str = str(row.get("Required", "No")).strip().upper()
                required = required_str in ("YES", "TRUE", "1")

                # `or 0` converts None (empty cell) to 0 before int()
                coord_page = int(row.get("CoordPage", 0) or 0)
                coord_x = int(row.get("CoordX", 0) or 0)
                coord_y = int(row.get("CoordY", 0) or 0)
                coord_width = int(row.get("CoordWidth", 0) or 0)
                coord_height = int(row.get("CoordHeight", 0) or 0)
                tolerance = int(row.get("Tolerance", 0) or 0)

                tag = TagDefinition(
                    tag_id=row.get("TagID", ""),
                    tag_type=row.get("TagType", "extraction"),
                    extraction_method=extraction_method,
                    field_name=row.get("FieldName", ""),
                    source_document=row.get("SourceDocument", ""),
                    field_type=field_type,
                    required=required,
                    search_keywords=row.get("SearchKeywords"),
                    start_anchor=row.get("StartAnchor"),
                    end_anchor=row.get("EndAnchor"),
                    coord_page=coord_page,
                    coord_x=coord_x,
                    coord_y=coord_y,
                    coord_width=coord_width,
                    coord_height=coord_height,
                    tolerance=tolerance,
                    fallback_keywords=row.get("FallbackKeywords"),
                    notes=row.get("Notes"),
                )
                tags.append(tag)

        except Exception as e:
            raise ConfigError(f"Failed to load TAG_ENGINE: {e}") from e

        return tags

    def _load_qa_rules(self) -> list[QARule]:
        """
        Read QA rules from the QA sheet.

        Rows where RuleID is empty are skipped.

        Returns:
            List of QARule objects; empty list if QA sheet doesn't exist.

        Raises:
            ConfigError: If a row is malformed.
        """
        rules: list[QARule] = []

        if not self.excel_mgr.sheet_exists("QA"):
            return rules

        try:
            qa_rows = self.excel_mgr.read_rows_as_dicts("QA")

            for row in qa_rows:
                if not row.get("RuleID"):
                    continue

                rule = QARule(
                    rule_id=row.get("RuleID", ""),
                    rule_name=row.get("RuleName", ""),
                    field_name=row.get("FieldName", ""),
                    rule_type=row.get("RuleType", "CUSTOM"),
                    rule_definition=row.get("RuleDefinition", ""),
                    fail_action=row.get("FailAction", "FLAG"),
                    severity=row.get("Severity", "MEDIUM"),
                    priority=int(row.get("Priority", 0) or 0),
                )
                rules.append(rule)

        except Exception as e:
            raise ConfigError(f"Failed to load QA rules: {e}") from e

        return rules

    def _load_output_schema(self) -> list[SchemaField]:
        """
        Infer the output schema from the column headers of the OUTPUT sheet.

        All columns default to DataType.TEXT and required=False.  If richer
        schema information is needed, a dedicated schema config sheet should
        be added in a future iteration.

        Returns:
            List of SchemaField objects; empty list if OUTPUT sheet doesn't exist.
        """
        schema: list[SchemaField] = []

        if not self.excel_mgr.sheet_exists("OUTPUT"):
            return schema

        try:
            ws = self.excel_mgr.get_sheet("OUTPUT")

            for col_idx in range(1, ws.max_column + 1):
                header = ws.cell(1, col_idx).value
                if header:
                    schema.append(
                        SchemaField(
                            field_name=str(header).strip(),
                            data_type=DataType.TEXT,
                            required=False,
                        )
                    )

        except Exception as e:
            raise ConfigError(f"Failed to load output schema: {e}") from e

        return schema

    def _load_approval_workflow(self) -> ApprovalWorkflow:
        """
        Read approval workflow settings from the CONFIG sheet.

        All fields are optional in CONFIG; missing keys produce None rather
        than raising an error, so a minimal workbook still loads cleanly.

        Returns:
            ApprovalWorkflow with whatever fields are present in CONFIG.

        Raises:
            ConfigError: If the WorkflowType value is not a valid enum member.
        """
        try:
            workflow_type_str = self.excel_mgr.read_config_value("WorkflowType")
            workflow_type = (
                WorkflowType(workflow_type_str) if workflow_type_str else WorkflowType.SIMPLE
            )

            return ApprovalWorkflow(
                workflow_type=workflow_type,
                level1_approver=self.excel_mgr.read_config_value("Level1Approver"),
                level1_email=self.excel_mgr.read_config_value("Level1Email"),
                level2_approver=self.excel_mgr.read_config_value("Level2Approver"),
                level2_email=self.excel_mgr.read_config_value("Level2Email"),
            )
        except Exception as e:
            raise ConfigError(f"Failed to load approval workflow: {e}") from e

    def close(self) -> None:
        """Close the underlying workbook and release the file handle."""
        self.excel_mgr.close()

    def __enter__(self) -> "ConfigManager":
        return self

    def __exit__(self, exc_type, exc_val, exc_tb) -> None:
        self.close()
