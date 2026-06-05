"""
Data models for audit framework.

Defines all core data structures using Pydantic v2 for validation,
serialization, and type safety.
"""

from datetime import date
from enum import Enum
from typing import Optional

from pydantic import BaseModel, EmailStr, Field, field_validator


class ExtractionMethod(str, Enum):
    """DataSnipper extraction methods."""

    DS_SEARCH = "DS_SEARCH"
    DS_COORDS = "DS_COORDS"
    HYBRID = "HYBRID"


class DocumentType(str, Enum):
    """Types of source documents."""

    INVOICE = "Invoice"
    PO = "PO"
    RECEIPT = "Receipt"
    STATEMENT = "Statement"
    CONTRACT = "Contract"
    OTHER = "Other"


class RuleType(str, Enum):
    """QA rule validation types."""

    RANGE = "RANGE"
    LOOKUP = "LOOKUP"
    FORMAT = "FORMAT"
    CROSS_FIELD = "CROSS_FIELD"
    DUPLICATE = "DUPLICATE"
    CUSTOM = "CUSTOM"


class FailAction(str, Enum):
    """Action when QA rule fails."""

    FLAG = "FLAG"
    BLOCK = "BLOCK"
    WARN = "WARN"


class Severity(str, Enum):
    """QA exception severity levels."""

    LOW = "LOW"
    MEDIUM = "MEDIUM"
    HIGH = "HIGH"
    CRITICAL = "CRITICAL"


class EngagementType(str, Enum):
    """Types of audit engagements."""

    CASH = "Cash"
    AR = "AR"
    AP = "AP"
    CONTRACTS = "Contracts"
    INVENTORY = "Inventory"


class DataType(str, Enum):
    """Output field data types."""

    TEXT = "Text"
    NUMBER = "Number"
    DATE = "Date"
    CURRENCY = "Currency"
    BOOLEAN = "Boolean"


class WorkflowType(str, Enum):
    """Approval workflow types."""

    SIMPLE = "SIMPLE"
    MULTI_LEVEL = "MULTI_LEVEL"
    NONE = "NONE"


class DocumentSpec(BaseModel):
    """Specification for a source document."""

    doc_id: str = Field(..., min_length=1)
    file_path: str = Field(..., min_length=1)
    page_count: int = Field(..., gt=0)
    doc_type: DocumentType = DocumentType.OTHER
    extraction_method: ExtractionMethod = ExtractionMethod.DS_SEARCH

    class Config:
        """Pydantic configuration."""

        use_enum_values = False


class TagDefinition(BaseModel):
    """Definition of a data extraction tag."""

    tag_id: str = Field(..., min_length=1)
    tag_type: str = Field(default="extraction")
    extraction_method: ExtractionMethod = Field(...)
    field_name: str = Field(..., min_length=1)
    source_document: str = Field(..., min_length=1)
    field_type: DataType = DataType.TEXT
    required: bool = False
    search_keywords: Optional[str] = None
    start_anchor: Optional[str] = None
    end_anchor: Optional[str] = None
    coord_page: int = Field(default=0, ge=0)
    coord_x: int = Field(default=0, ge=0)
    coord_y: int = Field(default=0, ge=0)
    coord_width: int = Field(default=0, ge=0)
    coord_height: int = Field(default=0, ge=0)
    tolerance: int = Field(default=0, ge=0)
    fallback_keywords: Optional[str] = None
    notes: Optional[str] = None

    @field_validator("tag_id", "field_name", mode="before")
    @classmethod
    def strip_strings(cls, v: str) -> str:
        """Strip whitespace from string fields."""
        return v.strip() if isinstance(v, str) else v


class QARule(BaseModel):
    """Quality assurance validation rule."""

    rule_id: str = Field(..., min_length=1)
    rule_name: str = Field(..., min_length=1)
    field_name: str = Field(..., min_length=1)
    rule_type: RuleType = Field(...)
    rule_definition: str = Field(..., min_length=1)
    fail_action: FailAction = FailAction.FLAG
    severity: Severity = Severity.MEDIUM
    priority: int = Field(default=0, ge=0)

    @field_validator("rule_id", "rule_name", "field_name", mode="before")
    @classmethod
    def strip_strings(cls, v: str) -> str:
        """Strip whitespace from string fields."""
        return v.strip() if isinstance(v, str) else v


class SchemaField(BaseModel):
    """Output schema field definition."""

    field_name: str = Field(..., min_length=1)
    data_type: DataType = DataType.TEXT
    required: bool = False
    format_spec: Optional[str] = None

    @field_validator("field_name", mode="before")
    @classmethod
    def strip_string(cls, v: str) -> str:
        """Strip whitespace from field name."""
        return v.strip() if isinstance(v, str) else v


class ApprovalWorkflow(BaseModel):
    """Approval workflow configuration."""

    workflow_type: WorkflowType = WorkflowType.SIMPLE
    level1_approver: Optional[str] = None
    level1_email: Optional[EmailStr] = None
    level2_approver: Optional[str] = None
    level2_email: Optional[EmailStr] = None

    @field_validator("level1_approver", "level2_approver", mode="before")
    @classmethod
    def strip_strings(cls, v: Optional[str]) -> Optional[str]:
        """Strip whitespace from approver names."""
        return v.strip() if isinstance(v, str) else v


class ConfigObject(BaseModel):
    """Master configuration object for an engagement."""

    engagement_id: str = Field(..., min_length=1)
    engagement_type: EngagementType = Field(...)
    period_start_date: date = Field(...)
    period_end_date: date = Field(...)
    lead_auditor: str = Field(..., min_length=1)
    client_name: str = Field(..., min_length=1)
    client_contact: Optional[str] = None
    framework_version: str = Field(default="1.0")
    documents: list[DocumentSpec] = Field(default_factory=list)
    tags: list[TagDefinition] = Field(default_factory=list)
    qa_rules: list[QARule] = Field(default_factory=list)
    output_schema: list[SchemaField] = Field(default_factory=list)
    approval_workflow: ApprovalWorkflow = Field(default_factory=ApprovalWorkflow)

    @field_validator("engagement_id", "lead_auditor", "client_name", mode="before")
    @classmethod
    def strip_strings(cls, v: str) -> str:
        """Strip whitespace from string fields."""
        return v.strip() if isinstance(v, str) else v

    @field_validator("period_end_date")
    @classmethod
    def validate_date_range(cls, v: date, info) -> date:
        """Ensure period_end_date is after period_start_date."""
        start_date = info.data.get("period_start_date")
        if start_date and v < start_date:
            raise ValueError("period_end_date must be after period_start_date")
        return v
