"""
Generate a DataSnipper-ready Excel engagement workbook.

This script creates a .xlsx file named with the _(ds) suffix that DataSnipper
requires.  When you open the workbook in Excel on a machine with DataSnipper
installed, DataSnipper automatically scans every cell for DS_SEARCH and
DS_COORDS tag strings and creates snips against the referenced PDFs.

Generated sheets
----------------
  CONFIG      — key/value engagement metadata
  TAG_ENGINE  — tag definitions; the SourceTag column holds the strings
                DataSnipper reads (DS_SEARCH[...] / DS_COORDS[...])
  QA          — quality assurance business rules (post-extraction validation)
  OUTPUT      — blank output area with column headers (DataSnipper fills this)
  AUDIT_LOG   — empty log sheet populated by the framework during the workflow

DataSnipper tag format (v4.0+)
-------------------------------
  DS_SEARCH[filename|pageNumber|query]
  DS_COORDS[filename|pageNumber|x1|y1|x2|y2]

  The workbook filename MUST end in _(ds) — e.g. engagement_(ds).xlsx

File path guidance
------------------
  Simplest setup: place the PDF in the same folder as the workbook and use
  just the filename in --source-pdf.  DataSnipper resolves relative paths
  from the workbook's location on disk.

Usage
-----
  # Defaults: AP sample engagement, output to engagement_(ds).xlsx
  python scripts/generate_workbook.py

  # Point at your actual PDF (just the filename — keep PDF next to the xlsx)
  python scripts/generate_workbook.py --source-pdf "my_invoice.pdf"

  # Full override
  python scripts/generate_workbook.py \\
      --output "client_ap_(ds).xlsx" \\
      --engagement-id ENG-2026-001 \\
      --client "Acme Corp" \\
      --auditor "Jane Smith" \\
      --source-pdf "acme_invoice_q1.pdf"

After running
-------------
  1. Push to GitHub and pull on your DataSnipper machine
  2. Copy your PDF to the same folder as the generated workbook
  3. Open the _(ds).xlsx in Excel — DataSnipper auto-processes on open
  4. DataSnipper finds SourceTag values in TAG_ENGINE and creates snips
  5. Extracted values appear linked to the snipped regions in the PDF
  6. Run the Python CLI (audit-cli) or VBA framework to validate / QA results

Customizing for your PDF
------------------------
  Edit make_sample_ap_tags() so the search_keywords match text that actually
  appears in your PDF.  For DS_COORDS, open the PDF in DataSnipper first,
  hover over the region you want, and note the pixel coordinates.
"""

import argparse
from datetime import date
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

import sys
sys.path.insert(0, str(Path(__file__).parent.parent))

from audit_framework.tag_builder import TagBuilder
from audit_framework.models import (
    DataType,
    ExtractionMethod,
    TagDefinition,
    QARule,
    RuleType,
    FailAction,
    Severity,
)


# ── Styling constants ────────────────────────────────────────────────────────

HEADER_FILL   = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
HEADER_FONT   = Font(color="FFFFFF", bold=True)
TAG_COL_FILL  = PatternFill(start_color="375623", end_color="375623", fill_type="solid")


def _style_header_row(ws, row: int, num_cols: int) -> None:
    """Apply dark-blue header styling to a row."""
    for col in range(1, num_cols + 1):
        cell = ws.cell(row, col)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")


def _auto_fit_columns(ws) -> None:
    """Set column widths based on max content length (capped at 60 chars)."""
    for col_cells in ws.columns:
        max_len = max(
            (len(str(cell.value)) for cell in col_cells if cell.value is not None),
            default=10,
        )
        ws.column_dimensions[get_column_letter(col_cells[0].column)].width = min(max_len + 4, 60)


# ── Sheet builders ───────────────────────────────────────────────────────────

def build_config_sheet(wb: Workbook, config: dict) -> None:
    """
    Write the CONFIG sheet (col A = key, col B = value).

    This layout matches what ExcelWorkbookManager.read_config_value() expects.
    """
    ws = wb.create_sheet("CONFIG")

    ws["A1"] = "Engagement Configuration"
    ws["A1"].font = Font(bold=True, size=14)
    ws.merge_cells("A1:F1")

    headers = ["Key", "Value", "DataType", "Scope", "Description", "Required"]
    for col, h in enumerate(headers, 1):
        ws.cell(3, col, h)
    _style_header_row(ws, 3, len(headers))

    rows = [
        ("EngagementID",    config.get("engagement_id", ""),       "String", "Global", "Unique engagement identifier",               "Yes"),
        ("EngagementType",  config.get("engagement_type", "AP"),   "String", "Global", "Audit area (Cash/AR/AP/Contracts/Inventory)", "Yes"),
        ("PeriodStartDate", str(config.get("period_start", "")),   "Date",   "Global", "Start date for the engagement period",        "Yes"),
        ("PeriodEndDate",   str(config.get("period_end", "")),     "Date",   "Global", "End date for the engagement period",          "Yes"),
        ("LeadAuditor",     config.get("lead_auditor", ""),        "String", "Global", "Primary audit lead",                         "Yes"),
        ("ClientName",      config.get("client_name", ""),         "String", "Global", "Client name",                                "Yes"),
        ("ClientContact",   config.get("client_contact", ""),      "String", "Global", "Client contact name",                        "No"),
        ("FrameworkVersion","1.0",                                  "String", "Global", "Version of this framework",                  "No"),
        ("WorkflowType",    config.get("workflow_type", "SIMPLE"), "String", "Global", "Approval workflow (SIMPLE/MULTI_LEVEL/NONE)", "No"),
        ("Level1Approver",  config.get("level1_approver", ""),     "String", "Global", "Primary approver name",                      "No"),
        ("Level1Email",     config.get("level1_email", ""),        "String", "Global", "Primary approver email",                     "No"),
        ("Level2Approver",  config.get("level2_approver", ""),     "String", "Global", "Secondary approver name",                    "No"),
        ("Level2Email",     config.get("level2_email", ""),        "String", "Global", "Secondary approver email",                   "No"),
    ]

    for row_idx, row_data in enumerate(rows, 4):
        for col, value in enumerate(row_data, 1):
            ws.cell(row_idx, col, value)

    _auto_fit_columns(ws)


def build_tag_engine_sheet(wb: Workbook, tag_definitions: list[TagDefinition]) -> None:
    """
    Write the TAG_ENGINE sheet.

    The SourceTag column (last column, green header) holds the actual
    DataSnipper tag strings.  DataSnipper scans the workbook on open and
    processes any cell it finds containing a DS_SEARCH[...] or DS_COORDS[...]
    string.

    DS_SEARCH columns: SourceDocument, SearchPage, Query
    DS_COORDS columns: SourceDocument, CoordPage, X1, Y1, X2, Y2
    (unused columns are left blank for each tag type)
    """
    ws = wb.create_sheet("TAG_ENGINE")

    headers = [
        "TagID", "FieldName", "ExtractionMethod",
        # DS_SEARCH fields
        "SourceDocument", "SearchPage", "Query",
        # DS_COORDS fields
        "CoordPage", "X1", "Y1", "X2", "Y2",
        # Metadata
        "Notes",
        # The generated tag string DataSnipper reads
        "SourceTag",
    ]

    for col, h in enumerate(headers, 1):
        ws.cell(1, col, h)
    _style_header_row(ws, 1, len(headers))

    # Highlight the SourceTag column in green — it's the DataSnipper input
    source_tag_col = headers.index("SourceTag") + 1
    ws.cell(1, source_tag_col).fill = TAG_COL_FILL

    for row_idx, tag in enumerate(tag_definitions, 2):
        try:
            tag_string = TagBuilder.build_tag(tag)
        except Exception as e:
            tag_string = f"BUILD_ERROR: {e}"

        x2 = (tag.coord_x + tag.coord_width) if tag.coord_width else ""
        y2 = (tag.coord_y + tag.coord_height) if tag.coord_height else ""

        values = [
            tag.tag_id,
            tag.field_name,
            tag.extraction_method.value,
            # DS_SEARCH
            tag.source_document,
            tag.search_page if tag.extraction_method == ExtractionMethod.DS_SEARCH else "",
            tag.search_keywords or "" if tag.extraction_method == ExtractionMethod.DS_SEARCH else "",
            # DS_COORDS
            tag.coord_page if tag.extraction_method == ExtractionMethod.DS_COORDS else "",
            tag.coord_x if tag.extraction_method == ExtractionMethod.DS_COORDS else "",
            tag.coord_y if tag.extraction_method == ExtractionMethod.DS_COORDS else "",
            x2 if tag.extraction_method == ExtractionMethod.DS_COORDS else "",
            y2 if tag.extraction_method == ExtractionMethod.DS_COORDS else "",
            tag.notes or "",
            tag_string,
        ]
        for col, value in enumerate(values, 1):
            ws.cell(row_idx, col, value)

    _auto_fit_columns(ws)


def build_qa_sheet(wb: Workbook, rules: list[QARule]) -> None:
    """
    Write the QA sheet with business rule definitions.

    Rules are evaluated by QAEngine after DataSnipper populates the OUTPUT sheet.
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
            rule.rule_id, rule.rule_name, rule.field_name,
            rule.rule_type.value, rule.rule_definition,
            rule.fail_action.value, rule.severity.value, rule.priority,
        ]
        for col, value in enumerate(values, 1):
            ws.cell(row_idx, col, value)

    _auto_fit_columns(ws)


def build_output_sheet(wb: Workbook, output_columns: list[str]) -> None:
    """
    Write the OUTPUT sheet with column headers only.

    DataSnipper writes extracted values into cells in this sheet.
    Column names must match the field_name values in TAG_ENGINE.
    """
    ws = wb.create_sheet("OUTPUT")

    for col, name in enumerate(output_columns, 1):
        ws.cell(1, col, name)
    _style_header_row(ws, 1, len(output_columns))

    ws.freeze_panes = "A2"
    _auto_fit_columns(ws)


def build_audit_log_sheet(wb: Workbook) -> None:
    """Write an empty AUDIT_LOG sheet; rows are appended by AuditLog.save()."""
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
    Return a starter set of A/P tag definitions using the actual DataSnipper format.

    IMPORTANT: The search_keywords values below are placeholder queries.
    Open your PDF and replace them with text that actually appears on the page
    near the value you want to extract.

    For DS_COORDS tags, open the PDF in DataSnipper, hover over the target
    region to get the pixel coordinates, then update coord_x/y/width/height.

    Generated tag strings (SourceTag column):
        DS_SEARCH[vendor_invoice.pdf|1|Invoice Number]
        DS_SEARCH[vendor_invoice.pdf|1|Total Amount]
        DS_COORDS[vendor_invoice.pdf|1|400|100|550|120]
        ... etc.
    """
    return [
        # ── DS_SEARCH examples ───────────────────────────────────────────
        TagDefinition(
            tag_id="AP_InvoiceNum_S",
            extraction_method=ExtractionMethod.DS_SEARCH,
            field_name="invoice_number",
            source_document=source_doc,
            search_page=1,
            search_keywords="Invoice Number",
            notes="Replace 'Invoice Number' with the exact label text in your PDF",
        ),
        TagDefinition(
            tag_id="AP_VendorName_S",
            extraction_method=ExtractionMethod.DS_SEARCH,
            field_name="vendor_name",
            source_document=source_doc,
            search_page=1,
            search_keywords="Bill From",
            notes="Replace 'Bill From' with the vendor name header in your PDF",
        ),
        TagDefinition(
            tag_id="AP_InvoiceDate_S",
            extraction_method=ExtractionMethod.DS_SEARCH,
            field_name="invoice_date",
            source_document=source_doc,
            search_page=1,
            search_keywords="Invoice Date",
            notes="Replace 'Invoice Date' with the date label in your PDF",
        ),
        TagDefinition(
            tag_id="AP_InvoiceTotal_S",
            extraction_method=ExtractionMethod.DS_SEARCH,
            field_name="invoice_total",
            source_document=source_doc,
            search_page=1,
            search_keywords="Total Amount",
            notes="Replace 'Total Amount' with the total label in your PDF (try 'Grand Total')",
        ),
        TagDefinition(
            tag_id="AP_PoNumber_S",
            extraction_method=ExtractionMethod.DS_SEARCH,
            field_name="po_number",
            source_document=source_doc,
            search_page=1,
            search_keywords="PO Number",
            notes="Replace 'PO Number' with the PO label in your PDF",
        ),
        TagDefinition(
            tag_id="AP_DueDate_S",
            extraction_method=ExtractionMethod.DS_SEARCH,
            field_name="due_date",
            source_document=source_doc,
            search_page=1,
            search_keywords="Due Date",
            notes="Replace 'Due Date' with the payment due label in your PDF",
        ),

        # ── DS_COORDS example ────────────────────────────────────────────
        # To use this: open PDF in DataSnipper, hover to get pixel coords,
        # update coord_x/y/width/height, then regenerate the workbook.
        TagDefinition(
            tag_id="AP_InvoiceDate_C",
            extraction_method=ExtractionMethod.DS_COORDS,
            field_name="invoice_date_coords",
            source_document=source_doc,
            coord_page=1,
            coord_x=400,
            coord_y=100,
            coord_width=150,
            coord_height=20,
            notes="PLACEHOLDER coords — update x/y/width/height to match your PDF layout",
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
            rule_name="Invoice date not in future",
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
    del wb[wb.sheetnames[0]]  # remove default empty sheet

    tags = make_sample_ap_tags(source_doc=config.get("source_pdf", "vendor_invoice.pdf"))
    rules = make_sample_qa_rules()
    output_columns = [t.field_name for t in tags]

    build_config_sheet(wb, config)
    build_tag_engine_sheet(wb, tags)
    build_qa_sheet(wb, rules)
    build_output_sheet(wb, output_columns)
    build_audit_log_sheet(wb)

    wb.save(output_path)

    print(f"Generated : {output_path}")
    print(f"Sheets    : {', '.join(wb.sheetnames)}")
    print(f"Tags      : {len(tags)}")
    print(f"Source PDF: {config.get('source_pdf')} (must be in same folder as the workbook)")
    print()
    print("Next steps:")
    print("  1. git push here; git pull on your DataSnipper machine")
    print(f"  2. Copy your PDF to the same folder as {output_path.name}")
    print(f"  3. Open {output_path.name} in Excel")
    print("  4. DataSnipper auto-processes on open — check the TAG_ENGINE SourceTag column")
    print("  5. If tags don't fire, verify the PDF filename in SourceTag matches exactly")
    print("  6. For DS_SEARCH: update search_keywords to match text in your actual PDF")
    print("  7. For DS_COORDS: update coordinates using DataSnipper's hover tool")


def main() -> None:
    parser = argparse.ArgumentParser(description="Generate a DataSnipper_(ds) engagement workbook")
    parser.add_argument(
        "--output", default="engagement_(ds).xlsx",
        help="Output .xlsx filename — must end in _(ds) for DataSnipper to auto-process",
    )
    parser.add_argument("--engagement-id", default="ENG-2026-001")
    parser.add_argument("--client", default="Sample Client Corp")
    parser.add_argument("--auditor", default="Lead Auditor")
    parser.add_argument(
        "--source-pdf", default="vendor_invoice.pdf",
        help="PDF filename as DataSnipper will see it (use just the filename if PDF is in same folder)",
    )
    args = parser.parse_args()

    # Enforce the _(ds) naming requirement
    output_name = args.output
    if not Path(output_name).stem.endswith("_(ds)"):
        stem = Path(output_name).stem
        output_name = stem + "_(ds).xlsx"
        print(f"Note: renamed output to {output_name} (DataSnipper requires _(ds) suffix)")

    today = date.today()
    config = {
        "engagement_id":   args.engagement_id,
        "engagement_type": "AP",
        "period_start":    today.replace(day=1, month=1),
        "period_end":      today,
        "lead_auditor":    args.auditor,
        "client_name":     args.client,
        "workflow_type":   "SIMPLE",
        "source_pdf":      args.source_pdf,
    }

    generate(Path(output_name), config)


if __name__ == "__main__":
    main()
