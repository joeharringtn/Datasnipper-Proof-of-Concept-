"""
ConfigManager - Load, validate, and manage engagement configuration.

Responsibility:
- Load configuration from CONFIG sheet
- Validate completeness and correctness
- Provide configuration data to other modules
- Track configuration version and changes
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
        Initialize ConfigManager.

        Args:
            workbook_path: Path to engagement workbook

        Raises:
            FileNotFoundError: If workbook doesn't exist
            WorksheetMissingError: If CONFIG sheet is missing
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
        """Load complete engagement configuration."""
        try:
            self.load_timestamp = datetime.now()

            engagement_id = self._read_config_value("EngagementID")
            engagement_type = self._read_config_value("EngagementType")
            period_start_date = self._read_config_value("PeriodStartDate")
            period_end_date = self._read_config_value("PeriodEndDate")
            lead_auditor = self._read_config_value("LeadAuditor")
            client_name = self._read_config_value("ClientName")
            client_contact = self._read_config_value("ClientContact")
            framework_version = self._read_config_value("FrameworkVersion") or "1.0"

            documents = self._load_documents()
            tags = self._load_tags()
            qa_rules = self._load_qa_rules()
            output_schema = self._load_output_schema()
            approval_workflow = self._load_approval_workflow()

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
            raise ConfigError(f"Configuration validation failed: {e}") from e
        except Exception as e:
            raise ConfigError(f"Failed to load configuration: {e}") from e

    def validate_config(self) -> tuple[bool, list[str]]:
        """Validate loaded configuration."""
        if not self.config:
            return False, ["Configuration not loaded"]

        errors = []

        if not self.config.engagement_id:
            errors.append("EngagementID is required")

        if not self.config.engagement_type:
            errors.append("EngagementType is required")

        if self.config.period_start_date <= self.config.period_end_date is False:
            errors.append(
                f"PeriodStartDate ({self.config.period_start_date}) must be before "
                f"PeriodEndDate ({self.config.period_end_date})"
            )

        if not self.config.lead_auditor:
            errors.append("LeadAuditor is recommended for audit traceability")

        if not self.config.client_name:
            errors.append("ClientName is required")

        if not self.config.documents:
            errors.append("At least one document specification is required")

        if not self.config.tags:
            errors.append("At least one tag definition is required")

        if (
            self.config.approval_workflow.workflow_type == WorkflowType.MULTI_LEVEL
            and not self.config.approval_workflow.level1_approver
        ):
            errors.append("Approval workflow requires Level1Approver")

        return len(errors) == 0, errors

    def is_loaded(self) -> bool:
        """Check if configuration is loaded."""
        return self.config is not None

    def get_config(self) -> ConfigObject:
        """Get loaded configuration."""
        if not self.config:
            raise ConfigError("Configuration not loaded. Call load_config() first.")
        return self.config

    def get_engagement_type(self) -> EngagementType:
        """Get engagement type."""
        if not self.config:
            raise ConfigError("Configuration not loaded")
        return self.config.engagement_type

    def get_documents(self) -> list[DocumentSpec]:
        """Get document specifications."""
        if not self.config:
            raise ConfigError("Configuration not loaded")
        return self.config.documents

    def get_tags(self) -> list[TagDefinition]:
        """Get tag definitions."""
        if not self.config:
            raise ConfigError("Configuration not loaded")
        return self.config.tags

    def get_qa_rules(self) -> list[QARule]:
        """Get QA rules."""
        if not self.config:
            raise ConfigError("Configuration not loaded")
        return self.config.qa_rules

    def get_output_schema(self) -> list[SchemaField]:
        """Get output schema."""
        if not self.config:
            raise ConfigError("Configuration not loaded")
        return self.config.output_schema

    def get_approval_workflow(self) -> ApprovalWorkflow:
        """Get approval workflow."""
        if not self.config:
            raise ConfigError("Configuration not loaded")
        return self.config.approval_workflow

    def _read_config_value(self, key: str) -> str:
        """Read configuration value from CONFIG sheet."""
        value = self.excel_mgr.read_config_value(key)
        if value is None:
            raise ConfigError(f"Required configuration '{key}' not found in CONFIG sheet")
        return str(value).strip()

    def _load_documents(self) -> list[DocumentSpec]:
        """Load document specifications from CONFIG or DOCUMENTS sheet."""
        return []

    def _load_tags(self) -> list[TagDefinition]:
        """Load tag definitions from TAG_ENGINE sheet."""
        tags = []

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

                field_type_str = row.get("FieldType", "Text")
                try:
                    field_type = DataType(field_type_str)
                except ValueError:
                    field_type = DataType.TEXT

                required_str = str(row.get("Required", "No")).strip().upper()
                required = required_str in ("YES", "TRUE", "1")

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
        """Load QA rules from QA sheet."""
        rules = []

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
        """Load output schema from OUTPUT sheet."""
        schema = []

        if not self.excel_mgr.sheet_exists("OUTPUT"):
            return schema

        try:
            ws = self.excel_mgr.get_sheet("OUTPUT")

            for col_idx in range(1, ws.max_column + 1):
                header = ws.cell(1, col_idx).value
                if header:
                    field = SchemaField(
                        field_name=str(header).strip(),
                        data_type=DataType.TEXT,
                        required=False,
                    )
                    schema.append(field)

        except Exception as e:
            raise ConfigError(f"Failed to load output schema: {e}") from e

        return schema

    def _load_approval_workflow(self) -> ApprovalWorkflow:
        """Load approval workflow from CONFIG sheet."""
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
        """Close workbook and cleanup."""
        self.excel_mgr.close()

    def __enter__(self) -> "ConfigManager":
        """Context manager entry."""
        return self

    def __exit__(self, exc_type, exc_val, exc_tb) -> None:
        """Context manager exit."""
        self.close()
