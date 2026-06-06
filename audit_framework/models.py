"""
Pydantic data models for the audit framework.

These models define every structured object passed between modules.
Using Pydantic v2 gives us free validation, serialization, and clear
type contracts at every boundary.

Model relationships (ownership flows top → bottom):

    ConfigObject
    ├── DocumentSpec[]     — source PDF/file references for this engagement
    ├── TagDefinition[]    — how to extract each field from those documents
    ├── QARule[]           — business rules applied after extraction
    ├── SchemaField[]      — expected output columns and their types
    └── ApprovalWorkflow   — who must sign off before the output is final

Enum types are all str-based so their .value serializes cleanly to Excel
cells and JSON without extra conversion.
"""

from datetime import date
from enum import Enum
from typing import Optional

from pydantic import BaseModel, EmailStr, Field, field_validator


class ExtractionMethod(str, Enum):
    """
    DataSnipper extraction strategy for a tag.

    DS_SEARCH  — keyword-based: DataSnipper scans the document text for
                 anchor phrases and extracts the value near them.
    DS_COORDS  — coordinate-based: DataSnipper reads from a fixed (x, y,
                 width, height) bounding box on a specific page.
    HYBRID     — DS_SEARCH is tried first; DS_COORDS is the fallback if
                 the keyword search yields no result.
    """

    DS_SEARCH = "DS_SEARCH"
    DS_COORDS = "DS_COORDS"
    HYBRID = "HYBRID"


class DocumentType(str, Enum):
    """
    Classification of a source document used in this engagement.

    Used for display and filtering only — does not affect extraction logic.
    """

    INVOICE = "Invoice"
    PO = "PO"
    RECEIPT = "Receipt"
    STATEMENT = "Statement"
    CONTRACT = "Contract"
    OTHER = "Other"


class RuleType(str, Enum):
    """
    Determines how QAEngine evaluates a QARule's rule_definition string.

    RANGE       — rule_definition = "min:max"  (numeric bounds)
    LOOKUP      — rule_definition = "val1,val2,..."  (allowlist)
    FORMAT      — rule_definition = regex pattern
    CROSS_FIELD — rule_definition = "field1:operator:field2"
    DUPLICATE   — rule_definition = field name to compare against
    CUSTOM      — rule_definition is ignored; skipped by the engine
    """

    RANGE = "RANGE"
    LOOKUP = "LOOKUP"
    FORMAT = "FORMAT"
    CROSS_FIELD = "CROSS_FIELD"
    DUPLICATE = "DUPLICATE"
    CUSTOM = "CUSTOM"


class FailAction(str, Enum):
    """
    What QAEngine does when a rule violation is detected.

    FLAG   — record the exception but allow processing to continue.
    BLOCK  — record the exception AND prevent the row from reaching the
             approval stage until the exception is resolved.
    WARN   — log a warning; does not create a formal QAException.
    """

    FLAG = "FLAG"
    BLOCK = "BLOCK"
    WARN = "WARN"


class Severity(str, Enum):
    """Severity level attached to a QA exception for triage and reporting."""

    LOW = "LOW"
    MEDIUM = "MEDIUM"
    HIGH = "HIGH"
    CRITICAL = "CRITICAL"


class EngagementType(str, Enum):
    """
    Category of audit engagement.

    Determines which document types and QA rules are typically relevant.
    """

    CASH = "Cash"
    AR = "AR"
    AP = "AP"
    CONTRACTS = "Contracts"
    INVENTORY = "Inventory"


class DataType(str, Enum):
    """
    Output field data type, used by Validator for type-checking and by
    DataMapper for normalization.

    TEXT     — no parsing; stored as-is.
    NUMBER   — parsed to float.
    DATE     — parsed via multiple common formats, stored as date object.
    CURRENCY — like NUMBER but strips $ and , before parsing.
    BOOLEAN  — yes/no/true/false/1/0 → Python bool.
    """

    TEXT = "Text"
    NUMBER = "Number"
    DATE = "Date"
    CURRENCY = "Currency"
    BOOLEAN = "Boolean"


class WorkflowType(str, Enum):
    """
    Approval workflow variant for this engagement.

    SIMPLE      — one approver required.
    MULTI_LEVEL — two levels of approval (level1_approver then level2_approver).
    NONE        — no approval step; output goes straight to export.
    """

    SIMPLE = "SIMPLE"
    MULTI_LEVEL = "MULTI_LEVEL"
    NONE = "NONE"


class DocumentSpec(BaseModel):
    """
    Reference to a source document used in this engagement.

    doc_id must uniquely identify the file within the engagement —
    it is also used as the file= value inside DS_COORDS tags.
    """

    doc_id: str = Field(..., min_length=1)
    file_path: str = Field(..., min_length=1)
    page_count: int = Field(..., gt=0)
    doc_type: DocumentType = DocumentType.OTHER
    extraction_method: ExtractionMethod = ExtractionMethod.DS_SEARCH

    class Config:
        """Keep enum objects (not raw strings) when accessing fields."""
        use_enum_values = False


class TagDefinition(BaseModel):
    """
    Specification for a single DataSnipper extraction tag.

    A tag definition is the input to TagBuilder; the output is a formatted
    tag string that DataSnipper can interpret.  Fields used depend on the
    extraction_method:

    DS_SEARCH   — search_keywords required; start_anchor/end_anchor optional.
    DS_COORDS   — source_document + coord_* fields required.
    HYBRID      — all of the above.

    Coordinate values are in DataSnipper's internal pixel units.
    """

    tag_id: str = Field(..., min_length=1)
    tag_type: str = Field(default="extraction")
    extraction_method: ExtractionMethod = Field(...)
    field_name: str = Field(..., min_length=1)
    source_document: str = Field(..., min_length=1)
    field_type: DataType = DataType.TEXT
    required: bool = False
    search_keywords: Optional[str] = None
    start_anchor: Optional[str] = None   # text that precedes the target value
    end_anchor: Optional[str] = None     # text that follows the target value
    coord_page: int = Field(default=0, ge=0)
    coord_x: int = Field(default=0, ge=0)
    coord_y: int = Field(default=0, ge=0)
    coord_width: int = Field(default=0, ge=0)
    coord_height: int = Field(default=0, ge=0)
    tolerance: int = Field(default=0, ge=0)  # pixel tolerance for coord matching
    fallback_keywords: Optional[str] = None
    notes: Optional[str] = None

    @field_validator("tag_id", "field_name", mode="before")
    @classmethod
    def strip_strings(cls, v: str) -> str:
        return v.strip() if isinstance(v, str) else v


class QARule(BaseModel):
    """
    A single QA validation rule applied by QAEngine to extracted data.

    rule_definition encoding depends on rule_type — see RuleType docstring
    for the expected format of each variant.

    Rules are applied in ascending priority order (lower number = higher
    priority, applied first).
    """

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
        return v.strip() if isinstance(v, str) else v


class SchemaField(BaseModel):
    """
    A single column in the output schema.

    Drives both Validator (type-checking raw input) and DataMapper
    (normalizing values into their target types before export).
    """

    field_name: str = Field(..., min_length=1)
    data_type: DataType = DataType.TEXT
    required: bool = False
    format_spec: Optional[str] = None  # future: format string for display

    @field_validator("field_name", mode="before")
    @classmethod
    def strip_string(cls, v: str) -> str:
        return v.strip() if isinstance(v, str) else v


class ApprovalWorkflow(BaseModel):
    """
    Approval routing configuration for the engagement.

    For MULTI_LEVEL workflows, level1_approver must review first.
    level2_approver is optional for SIMPLE; required for MULTI_LEVEL.
    Email fields are validated as proper email addresses by Pydantic.
    """

    workflow_type: WorkflowType = WorkflowType.SIMPLE
    level1_approver: Optional[str] = None
    level1_email: Optional[EmailStr] = None
    level2_approver: Optional[str] = None
    level2_email: Optional[EmailStr] = None

    @field_validator("level1_approver", "level2_approver", mode="before")
    @classmethod
    def strip_strings(cls, v: Optional[str]) -> Optional[str]:
        return v.strip() if isinstance(v, str) else v


class ConfigObject(BaseModel):
    """
    Master configuration object for a single audit engagement.

    Constructed by ConfigManager from the engagement workbook and passed to
    every other module as the single source of truth for this run.

    period_end_date is validated to be strictly after period_start_date.
    """

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
        return v.strip() if isinstance(v, str) else v

    @field_validator("period_end_date")
    @classmethod
    def validate_date_range(cls, v: date, info) -> date:
        """Reject configs where the audit period ends before it starts."""
        start_date = info.data.get("period_start_date")
        if start_date and v < start_date:
            raise ValueError("period_end_date must be after period_start_date")
        return v
