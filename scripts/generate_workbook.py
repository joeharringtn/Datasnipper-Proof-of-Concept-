"""
Generate a DataSnipper-ready Excel engagement workbook.

This script creates a fully structured .xlsx file that can be opened on a
machine with DataSnipper installed and used to run extractions against PDF
source documents.

Generated sheets
----------------
  CONFIG      — key/value engagement metadata
  TAG_ENGINE  — tag definitions with SourceTag column pre-filled
  QA          — quality assurance business rules
  OUTPUT      — blank output area with column headers (DataSnipper fills this)
  AUDIT_LOG   — empty log sheet, populated by the VBA framework or AuditLog module

Usage
-----
  # With defaults (AP sample engagement, outputs to engagement.xlsx)
  python scripts/generate_workbook.py

  # Custom output path
  python scripts/generate_workbook.py --output my_engagement.xlsx

  # Override engagement metadata
  python scripts/generate_workbook.py --engagement-id ENG-2026-001 --client "Acme Corp"

After running
-------------
  1. Push to GitHub
  2. Pull on your work machine (the one with DataSnipper)
  3. Open the generated .xlsx in Excel
  4. Attach the source PDF(s) via the DataSnipper ribbon
  5. Run DataSnipper extractions — results populate the OUTPUT sheet
  6. Run the VBA framework (or Python CLI) to validate and QA the results
"""

import argparse
from datetime import date, timedelta
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# Allow running from the project root without installing the package
import sys
sys.path.insert(0, str(Path(__file__).parent.parent))

from audit_framework.tag_builder import TagBuilder
from audit_framework.models import (
    DataType,
    EngagementType,
    ExtractionMethod,
    TagDefinition,
    QARule,
    RuleType,
    FailAction,
    Severity,
)


# ── Styling constants ────────────────────────────────────────────────────────

HEADER_FILL = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
HEADER_FONT = Font(color="FFFFFF", bold=True)
SUBHEADER_FILL = PatternFill(start_color="D6E4F0", end_color="D6E4F0", fill_type="solid")
SUBHEADER_FONT = Font(bold=True)


def _style_header_row(ws, row: int, num_cols: int) -> None:
    """Apply dark blue header styling to a row."""
    for col in range(1, num_cols + 1):
        cell = ws.cell(row, col)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")


def _auto_fit_columns(ws) -> None:
    """Set column widths based on max content length (capped at 60)."""
    for col_cells in ws.columns:
        max_len = max(
            (len(str(cell.value)) for cell in col_cells if cell.value is not None),
            default=10,
        )
        ws.column_dimensions[get_column_letter(col_cells[0].column)].width = min(max_len + 4, 60)


# ── Sheet builders ───────────────────────────────────────────────────────────

def build_config_sheet(wb: Workbook, config: dict) -> None:
    """
    Write the CONFIG sheet.

    Layout: column A = key name, column B = value.
    This matches what ExcelWorkbookManager.read_config_value() expects.
    """
    ws = wb.create_sheet("CONFIG")

    # Title row
    ws["A1"] = "Engagement Configuration"
    ws["A1"].font = Font(bold=True, size=14)
    ws.merge_cells("A1:F1")

    # Column headers
    headers = ["Key", "Value", "DataType", "Scope", "Description", "Required"]
    for col, h in enumerate(headers, 1):
        ws.cell(3, col, h)
    _style_header_row(ws, 3, len(headers))

    # Config rows — order matches config_template.csv
    rows = [
        ("EngagementID",    config.get("engagement_id", ""),         "String", "Global", "Unique engagement identifier",              "Yes"),
        ("EngagementType",  config.get("engagement_type", "AP"),     "String", "Global", "Audit area (Cash/AR/AP/Contracts/Inventory)","Yes"),
        ("PeriodStartDate", str(config.get("period_start", "")),     "Date",   "Global", "Start date for the engagement period",       "Yes"),
        ("PeriodEndDate",   str(config.get("period_end", "")),       "Date",   "Global", "End date for the engagement period",         "Yes"),
        ("LeadAuditor",     config.get("lead_auditor", ""),          "String", "Global", "Primary audit lead",                        "Yes"),
        ("ClientName",      config.get("client_name", ""),           "String", "Global", "Client name",                               "Yes"),
        ("ClientContact",   config.get("client_contact", ""),        "String", "Global", "Client contact name",                       "No"),
        ("FrameworkVersion","1.0",                                   "String", "Global", "Version of the framework",                  "No"),
        ("WorkflowType",    config.get("workflow_type", "SIMPLE"),   "String", "Global", "Approval workflow (SIMPLE/MULTI_LEVEL/NONE)","No"),
        ("Level1Approver",  config.get("level1_approver", ""),       "String", "Global", "Primary approver name",                     "No"),
        ("Level1Email",     config.get("level1_email", ""),          "String", "Global", "Primary approver email",                    "No"),
        ("Level2Approver",  config.get("level2_approver", ""),       "String", "Global", "Secondary approver name",                   "No"),
        ("Level2Email",     config.get("level2_email", ""),          "String", "Global", "Secondary approver email",                  "No"),
    ]

    for row_idx, row_data in enumerate(rows, 4):
        for col, value in enumerate(row_data, 1):
            ws.cell(row_idx, col, value)

    _auto_fit_columns(ws)


def build_tag_engine_sheet(wb: Workbook, tag_definitions: list[TagDefinition]) -> None:
    """
    Write the TAG_ENGINE sheet.

    Columns match tag_engine_template.csv, plus a SourceTag column that holds
    the generated DataSnipper tag string — this is the column DataSnipper (or
    the VBA TagBuilder) reads to know how to extract each field.
    """
    ws = wb.create_sheet("TAG_ENGINE")

    headers = [
        "TagID", "TagType", "ExtractionMethod", "FieldName", "SourceDocument",
        "FieldType", "Required", "SearchKeywords", "StartAnchor", "EndAnchor",
        "CoordPage", "CoordX", "CoordY", "CoordWidth", "CoordHeight",
        "Tolerance", "FallbackKeywords", "Notes", "SourceTag",
    ]
    for col, h in enumerate(headers, 1):
        ws.cell(1, col, h)
    _style_header_row(ws, 1, len(headers))

    # Highlight the SourceTag column differently — it's auto-generated output
    source_tag_col = headers.index("SourceTag") + 1
    ws.cell(1, source_tag_col).fill = PatternFill(
        start_color="375623", end_color="375623", fill_type="solid"
    )

    for row_idx, tag in enumerate(tag_definitions, 2):
        # Build the tag string; write "BUILD_ERROR" if it fails so the sheet
        # remains openable and the error is visible in context.
        try:
            tag_string = TagBuilder.build_tag(tag)
        except Exception as e:
            tag_string = f"BUILD_ERROR: {e}"

        values = [
            tag.tag_id,
            tag.tag_type,
            tag.extraction_method.value,
            tag.field_name,
            tag.source_document,
            tag.field_type.value,
            "Yes" if tag.required else "No",
            tag.search_keywords or "",
            tag.start_anchor or "",
            tag.end_anchor or "",
            tag.coord_page if tag.coord_page else "",
            tag.coord_x if tag.coord_x else "",
            tag.coord_y if tag.coord_y else "",
            tag.coord_width if tag.coord_width else "",
            tag.coord_height if tag.coord_height else "",
            tag.tolerance if tag.tolerance else "",
            tag.fallback_keywords or "",
            tag.notes or "",
            tag_string,
        ]
        for col, value in enumerate(values, 1):
            ws.cell(row_idx, col, value)

    _auto_fit_columns(ws)


def build_qa_sheet(wb: Workbook, rules: list[QARule]) -> None:
    """
    Write the QA sheet with business rule definitions.

    These rules are evaluated by QAEngine after DataSnipper populates
    the OUTPUT sheet.
    """
    ws = wb.create_sheet("QA")

    headers = [
        "RuleID", "RuleName", "FieldName", "RuleType",
        "RuleDefinition", "FailAction", "Severity", "Priority",
    ]
    for col, h in enumerate(headers, 1):
        ws.cell(1, col, h)
    _style_header_row(ws, 1, len(headers))

    for row_idx, rule in enumerate(rules, 2):
        values = [
            rule.rule_id,
            rule.rule_name,
            rule.field_name,
            rule.rule_type.value,
            rule.rule_definition,
            rule.fail_action.value,
            rule.severity.value,
            rule.priority,
        ]
        for col, value in enumerate(values, 1):
            ws.cell(row_idx, col, value)

    _auto_fit_columns(ws)


def build_output_sheet(wb: Workbook, output_columns: list[str]) -> None:
    """
    Write the OUTPUT sheet with column headers only.

    DataSnipper will populate the data rows when it runs extractions.
    The column names here must match the field_name values of the TagDefinitions
    so DataSnipper knows which column to write each extracted value into.
    """
    ws = wb.create_sheet("OUTPUT")

    for col, name in enumerate(output_columns, 1):
        ws.cell(1, col, name)
    _style_header_row(ws, 1, len(output_columns))

    # Freeze the header row so it stays visible when scrolling through results
    ws.freeze_panes = "A2"
    _auto_fit_columns(ws)


def build_audit_log_sheet(wb: Workbook) -> None:
    """
    Write an empty AUDIT_LOG sheet with the expected column headers.

    Rows are appended here by AuditLog.save() (Python) or AuditLog.bas (VBA)
    as the engagement progresses through the workflow.
    """
    ws = wb.create_sheet("AUDIT_LOG")

    headers = ["Timestamp", "EventType", "Message", "User", "Module",
               "RowID", "FieldName", "OldValue", "NewValue", "Details"]
    for col, h in enumerate(headers, 1):
        ws.cell(1, col, h)
    _style_header_row(ws, 1, len(headers))
    _auto_fit_columns(ws)


# ── Sample data ──────────────────────────────────────────────────────────────

def make_sample_ap_tags(source_doc: str = "vendor_invoice.pdf") -> list[TagDefinition]:
    """
    Return a realistic set of A/P tag definitions for a vendor invoice PDF.

    Replace the anchors and coordinates with values that match your actual PDF
    before running DataSnipper extractions.
    """
    return [
        TagDefinition(
            tag_id="AP_PoNumber_S",
            tag_type="extraction",
            extraction_method=ExtractionMethod.DS_SEARCH,
            field_name="po_number",
            source_document=source_doc,
            field_type=DataType.TEXT,
            required=True,
            start_anchor="PO #",
            end_anchor="Date",
            notes="Purchase order number — text after 'PO #' label",
        ),
        TagDefinition(
            tag_id="AP_VendorName_S",
            tag_type="extraction",
            extraction_method=ExtractionMethod.DS_SEARCH,
            field_name="vendor_name",
            source_document=source_doc,
            field_type=DataType.TEXT,
            required=True,
            start_anchor="From",
            end_anchor="Invoice",
            notes="Vendor name — text after 'From' header",
        ),
        TagDefinition(
            tag_id="AP_InvoiceDate_C",
            tag_type="extraction",
            extraction_method=ExtractionMethod.DS_COORDS,
            field_name="invoice_date",
            source_document=source_doc,
            field_type=DataType.DATE,
            required=True,
            coord_page=1,
            coord_x=400,
            coord_y=120,
            coord_width=120,
            coord_height=20,
            tolerance=5,
            notes="Invoice date — top-right header area. Adjust coords to match your PDF.",
        ),
        TagDefinition(
            tag_id="AP_InvoiceTotal_S",
            tag_type="extraction",
            extraction_method=ExtractionMethod.DS_SEARCH,
            field_name="invoice_total",
            source_document=source_doc,
            field_type=DataType.CURRENCY,
            required=True,
            start_anchor="Total Amount",
            end_anchor="Tax",
            fallback_keywords="Grand Total",
            notes="Total invoice amount — 'Grand Total' fallback if 'Total Amount' not found",
        ),
        TagDefinition(
            tag_id="AP_DueDate_C",
            tag_type="extraction",
            extraction_method=ExtractionMethod.DS_COORDS,
            field_name="due_date",
            source_document=source_doc,
            field_type=DataType.DATE,
            required=False,
            coord_page=1,
            coord_x=400,
            coord_y=250,
            coord_width=120,
            coord_height=20,
            notes="Payment due date. Adjust coords to match your PDF.",
        ),
        TagDefinition(
            tag_id="AP_InvoiceNum_S",
            tag_type="extraction",
            extraction_method=ExtractionMethod.DS_SEARCH,
            field_name="invoice_number",
            source_document=source_doc,
            field_type=DataType.TEXT,
            required=True,
            start_anchor="Invoice #",
            end_anchor="Date",
            fallback_keywords="Invoice No",
            notes="Invoice number — tries 'Invoice #' then 'Invoice No'",
        ),
    ]


def make_sample_qa_rules() -> list[QARule]:
    """Return basic QA rules for an A/P engagement."""
    return [
        QARule(
            rule_id="QA_001",
            rule_name="Invoice total must be positive",
            field_name="invoice_total",
            rule_type=RuleType.RANGE,
            rule_definition="0.01:99999999",
            fail_action=FailAction.BLOCK,
            severity=Severity.CRITICAL,
            priority=1,
        ),
        QARule(
            rule_id="QA_002",
            rule_name="PO number format",
            field_name="po_number",
            rule_type=RuleType.FORMAT,
            rule_definition=r"^[A-Z0-9\-]{4,20}$",
            fail_action=FailAction.FLAG,
            severity=Severity.MEDIUM,
            priority=2,
        ),
        QARule(
            rule_id="QA_003",
            rule_name="Invoice date must not be in future",
            field_name="invoice_date",
            rule_type=RuleType.CUSTOM,
            rule_definition="date_not_future",
            fail_action=FailAction.WARN,
            severity=Severity.LOW,
            priority=3,
        ),
    ]


# ── Entry point ──────────────────────────────────────────────────────────────

def generate(output_path: Path, config: dict) -> None:
    """Generate and save the engagement workbook."""
    wb = Workbook()

    # Remove the default empty sheet openpyxl creates
    del wb[wb.sheetnames[0]]

    tags = make_sample_ap_tags(source_doc=config.get("source_pdf", "vendor_invoice.pdf"))
    rules = make_sample_qa_rules()
    output_columns = [t.field_name for t in tags]

    build_config_sheet(wb, config)
    build_tag_engine_sheet(wb, tags)
    build_qa_sheet(wb, rules)
    build_output_sheet(wb, output_columns)
    build_audit_log_sheet(wb)

    wb.save(output_path)
    print(f"Generated: {output_path}")
    print(f"  Sheets : {', '.join(wb.sheetnames)}")
    print(f"  Tags   : {len(tags)}")
    print(f"  Rules  : {len(rules)}")
    print(f"\nNext steps:")
    print(f"  1. Push to GitHub and pull on your DataSnipper machine")
    print(f"  2. Open {output_path.name} in Excel")
    print(f"  3. Attach {config.get('source_pdf', 'vendor_invoice.pdf')} via the DataSnipper ribbon")
    print(f"  4. Use TAG_ENGINE > SourceTag values to configure DataSnipper snippers")
    print(f"  5. Run extractions — results populate the OUTPUT sheet")


def main() -> None:
    parser = argparse.ArgumentParser(description="Generate a DataSnipper engagement workbook")
    parser.add_argument("--output", default="engagement.xlsx", help="Output .xlsx file path")
    parser.add_argument("--engagement-id", default="ENG-2026-001")
    parser.add_argument("--client", default="Sample Client Corp")
    parser.add_argument("--auditor", default="Lead Auditor")
    parser.add_argument("--source-pdf", default="vendor_invoice.pdf",
                        help="Filename of the source PDF (must match what DataSnipper sees)")
    args = parser.parse_args()

    today = date.today()
    config = {
        "engagement_id": args.engagement_id,
        "engagement_type": "AP",
        "period_start": today.replace(day=1, month=1),
        "period_end": today,
        "lead_auditor": args.auditor,
        "client_name": args.client,
        "workflow_type": "SIMPLE",
        "source_pdf": args.source_pdf,
    }

    generate(Path(args.output), config)


if __name__ == "__main__":
    main()
