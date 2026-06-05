# CONFIG_GUIDE.md - Configuration & Customization

## Overview

This guide explains how to configure the framework for specific audit engagements. Configuration is stored in the CONFIG sheet of each engagement workbook and managed by the ConfigManager module.

---

## Configuration Structure

### High-Level Organization

```
CONFIG Sheet Layout:
├─ SECTION A: Engagement Metadata
├─ SECTION B: Document Specifications
├─ SECTION C: Tag Definitions
├─ SECTION D: QA Rules (Engagement-Specific)
├─ SECTION E: Output Schema
├─ SECTION F: Approval Workflow
└─ SECTION G: Contact & Support
```

---

## Section A: Engagement Metadata

### Purpose
Define basic information about the audit engagement.

### Configuration Template

```
╔════════════════════════════════════════════════════════════════╗
║                  ENGAGEMENT METADATA                          ║
╠════════════════════════════════════════════════════════════════╣
║                                                                ║
║  Engagement ID:                2026-CASH-01                   ║
║  [Description: Unique engagement identifier for tracking]     ║
║                                                                ║
║  Engagement Type:              Cash                            ║
║  [Options: Cash | AR | AP | Contracts | Inventory]            ║
║                                                                ║
║  Audit Period:                 01/01/2026 - 06/30/2026        ║
║  [From / To dates]                                            ║
║                                                                ║
║  Lead Auditor:                 John Smith                      ║
║  [Name of audit team lead]                                     ║
║                                                                ║
║  Client Name:                  ABC Manufacturing Corp          ║
║  [Company being audited]                                       ║
║                                                                ║
║  Client Contact:               jane.doe@abccorp.com            ║
║  [For coordination/questions]                                  ║
║                                                                ║
║  Framework Version:            1.0                            ║
║  [VBA Framework version being used]                            ║
║                                                                ║
╚════════════════════════════════════════════════════════════════╝
```

---

## Section B: Document Specifications

### Purpose
Define which source documents will be processed and their locations.

### Configuration Template

```
╔════════════════════════════════════════════════════════════════╗
║              DOCUMENT SPECIFICATIONS                           ║
╠════════════════════════════════════════════════════════════════╣
║                                                                ║
║ Document 1:                                                    ║
║   Doc ID:              BANK_STMT_JUNE                          ║
║   File Path:           C:\Engagements\2026-CASH-01\Bank_...   ║
║   File Name:           Bank_Statement_June_2026.pdf            ║
║   Page Count:          25                                      ║
║   Document Type:       Bank Statement                          ║
║   Extraction Method:   DS_COORDS                               ║
║   [Notes: Standard format; same layout each month]             ║
║                                                                ║
║ Document 2:                                                    ║
║   Doc ID:              DEPOSIT_DETAILS_JUNE                    ║
║   File Path:           C:\Engagements\2026-CASH-01\...         ║
║   File Name:           Deposit_Details_June_2026.pdf           ║
║   Page Count:          12                                      ║
║   Document Type:       Deposit Detail Report                   ║
║   Extraction Method:   DS_SEARCH                               ║
║   [Notes: Format varies; use keyword search]                   ║
║                                                                ║
║ Document 3:                                                    ║
║   Doc ID:              RECONCILIATION                          ║
║   File Path:           C:\Engagements\2026-CASH-01\...         ║
║   File Name:           Bank_Reconciliation_June_2026.xlsx      ║
║   [Manual reference; not extracted]                            ║
║                                                                ║
╚════════════════════════════════════════════════════════════════╝
```

### Best Practices

| Aspect | Best Practice | Example |
|--------|---------------|---------|
| **File Paths** | Use full absolute paths | `C:\Engagements\2026-CASH-01\docs\` |
| **Page Counts** | Verify actual page count before entry | If stated 25, confirm it's exactly 25 |
| **Doc IDs** | Use short, descriptive codes | `BANK_STMT` not `Document 1` |
| **Extraction Method** | Choose based on document consistency | DS_COORDS if layout fixed; DS_SEARCH if variable |
| **Document Types** | Use standard categories | Invoice, Receipt, Statement, Report, etc. |

---

## Section C: Tag Definitions (TAG_ENGINE)

### Purpose
Define which data fields to extract from which documents using which methods.

### Configuration Template

Each engagement manually populates the TAG_ENGINE sheet:

```
┌─────┬──────────────┬────────────┬─────────────┬───────────┐
│Tag #│ Field Name   │ Extraction │ Source Docs │ Required? │
├─────┼──────────────┼────────────┼─────────────┼───────────┤
│  1  │ Deposit Amt  │ DS_COORDS  │ BANK_STMT   │ Yes       │
│  2  │ Deposit Date │ DS_SEARCH  │ DEPOSIT_DTL │ Yes       │
│  3  │ Bank Name    │ DS_SEARCH  │ BANK_STMT   │ Yes       │
│  4  │ Account Num  │ DS_COORDS  │ BANK_STMT   │ No        │
│  5  │ Reconciling  │ Manual     │ RECONCIL    │ No        │
│     │ Items        │            │             │           │
└─────┴──────────────┴────────────┴─────────────┴───────────┘
```

### Configuration Steps

**Step 1: Identify Data Points**
```
Question: What data do we need for this audit objective?

Cash Engagement Example:
  ✓ Daily deposit amounts (to trace to GL)
  ✓ Deposit dates (to verify period-end cutoff)
  ✓ Bank names (to verify authorized banks)
  ✓ Account numbers (for reconciliation reference)
```

**Step 2: Map to Source Documents**
```
Question: Where do we get this data?

deposit_amount → Bank Statement (line item report)
deposit_date → Deposit Details report
bank_name → Bank Statement header
account_number → Bank Statement account field
```

**Step 3: Choose Extraction Method**
```
Decision Matrix:

Field: deposit_amount
  Sources: Bank Statement (25 pages)
  Consistency: Very consistent (always same position)
  → Choose DS_COORDS (faster, more reliable)

Field: bank_name
  Source: Deposit Details (12 pages, variable format)
  Consistency: Varies; header appears differently
  → Choose DS_SEARCH (find "Bank:" keyword)
```

**Step 4: Define Search/Coordinate Parameters**
```
For DS_SEARCH:
  start_keyword: "Bank:"
  end_keyword: "Account"
  expected_type: text

For DS_COORDS:
  page: 1
  x_position: 300
  y_position: 450
  width: 100
  height: 20
  expected_type: currency
```

---

## Section D: QA Rules (Engagement-Specific)

### Purpose
Define validation rules for this engagement's data quality.

### Predefined vs. Custom Rules

```
PREDEFINED RULES (from QA_RULES.md):
  ├─ RANGE validation (amount bounds)
  ├─ LOOKUP validation (reference tables)
  ├─ FORMAT validation (pattern matching)
  ├─ CROSS_FIELD validation (multi-field logic)
  ├─ DUPLICATE detection
  └─ CUSTOM business logic

CONFIGURATION FOR THIS ENGAGEMENT:
  Select which predefined rules apply
  Add engagement-specific custom rules
```

### Configuration Template

```
╔════════════════════════════════════════════════════════════════╗
║                  QA RULES - CASH ENGAGEMENT                   ║
╠════════════════════════════════════════════════════════════════╣
║                                                                ║
║ RULE SET: Standard Cash Rules (apply predefined rules)         ║
║   [ X ] RANGE_CASH_01: Deposit $0 - $1M                       ║
║   [ X ] RANGE_CASH_02: Large deposits $100k+ reviewed          ║
║   [ X ] LOOKUP_CASH_01: Bank name in approved list             ║
║   [ X ] FORMAT_CASH_01: Account number format validation       ║
║   [ X ] CROSS_CASH_01: Deposit date in period                  ║
║   [ X ] DUP_CASH_01: No duplicate deposits same date/amount    ║
║                                                                ║
║ CUSTOM RULES (engagement-specific):                            ║
║   [ X ] No deposits from foreign banks (custom)                ║
║   [ X ] Daily totals must reconcile to GL ±$0.01               ║
║   [ X ] Bank reconciliation clearance within 3 days            ║
║                                                                ║
║ RULE OVERRIDE (if applicable):                                 ║
║   Exception: Allow deposits up to $2M if approved by CFO       ║
║   [Unusual; only add if specifically documented]               ║
║                                                                ║
╚════════════════════════════════════════════════════════════════╝
```

### How to Add Custom Rules

```
PROCESS:

1. Identify audit objective or risk:
   "We're concerned about duplicate deposits being recorded"

2. Define the rule logic:
   "For any deposit amount + date combination, check if
    same amount was deposited within 1 day"

3. Name the rule:
   CUSTOM_CASH_DUPLICATE_FUZZY

4. Specify the fail action:
   FLAG (alert QA team; don't block)

5. Assign severity:
   WARNING (warrants investigation but not critical)

6. Document the business reason:
   "Duplicates can occur if DataSnipper misreads pages
    or if manual re-entry occurs. This rule prevents
    accidental duplicate payments."

7. Enter in CONFIG sheet:
   QA Rules section
```

---

## Section E: Output Schema

### Purpose
Define the required fields and formats for final output data.

### Configuration Template

```
╔════════════════════════════════════════════════════════════════╗
║                      OUTPUT SCHEMA                             ║
╠════════════════════════════════════════════════════════════════╣
║                                                                ║
║ Output Field       │ Data Type    │ Format          │ Required ║
├────────────────────┼──────────────┼─────────────────┼──────────┤
║ deposit_amount     │ Currency     │ $#,##0.00       │ Yes      ║
║ deposit_date       │ Date         │ YYYY-MM-DD      │ Yes      ║
║ bank_name          │ Text         │ (no format)     │ Yes      ║
║ account_number     │ Text         │ Alphanumeric    │ No       ║
║ reconciliation_status│ Text       │ Cleared/Pending │ No       ║
║ extraction_timestamp│ DateTime     │ YYYY-MM-DD HH:MM│ No      ║
║ extracted_by_user  │ Text         │ (no format)     │ No       ║
║ qa_status          │ Text         │ Approved/Flagged│ No       ║
║ override_flag      │ Boolean      │ Yes/No          │ No       ║
║ override_justif.   │ Text         │ (no format)     │ No       ║
║                                                                ║
╚════════════════════════════════════════════════════════════════╝
```

### Data Type Definitions

| Type | Example | Format Rule |
|------|---------|------------|
| **Currency** | $1,250.50 | $#,##0.00 (2 decimals) |
| **Number** | 1250.50 | #,##0.## (up to 2 decimals) |
| **Date** | 2026-06-01 | YYYY-MM-DD |
| **DateTime** | 2026-06-01 14:30:00 | YYYY-MM-DD HH:MM:SS |
| **Text** | Chase Bank | No formatting (as-is) |
| **Boolean** | Yes/No | Yes \| No |
| **Percentage** | 95.5% | #.##% |

---

## Section F: Approval Workflow

### Purpose
Define sign-off requirements for this engagement.

### Configuration Template

```
╔════════════════════════════════════════════════════════════════╗
║              APPROVAL WORKFLOW CONFIGURATION                   ║
╠════════════════════════════════════════════════════════════════╣
║                                                                ║
║ Approval Workflow Type:  MULTI_LEVEL                           ║
║ [Options: SIMPLE | MULTI_LEVEL | NONE]                         ║
║                                                                ║
║ Level 1 Approver (QA Lead):                                    ║
║   Name:               Susan Johnson                            ║
║   Title:              Senior QA Specialist                     ║
║   Email:              susan.johnson@auditfirm.com              ║
║   Responsibility:     Approves QA decisions & exceptions       ║
║                                                                ║
║ Level 2 Approver (Engagement Manager):                         ║
║   Name:               Robert Chen                              ║
║   Title:              Audit Manager                            ║
║   Email:              robert.chen@auditfirm.com                ║
║   Responsibility:     Final sign-off on final output           ║
║                                                                ║
║ Electronic Signature:  Required? [ X ] Yes [ ] No              ║
║ Timestamp Required:    [ X ] Yes (automatic)                   ║
║ Exception Override:    Allowed by whom? Manager (Level 2)      ║
║                                                                ║
╚════════════════════════════════════════════════════════════════╝
```

### Approval Workflow Types

**SIMPLE**: One-level sign-off
```
QA Team → Senior Auditor → OUTPUT
(Single review gate)
```

**MULTI_LEVEL**: Two or more review gates
```
QA Team → QA Lead → Engagement Manager → OUTPUT
(Each level approves previous work)
```

**NONE**: No sign-off required
```
Automated processing → OUTPUT
(For low-risk data; not recommended)
```

---

## Section G: Contact & Support

### Purpose
Reference information for troubleshooting and support.

### Configuration Template

```
╔════════════════════════════════════════════════════════════════╗
║              CONTACT & SUPPORT INFORMATION                     ║
╠════════════════════════════════════════════════════════════════╣
║                                                                ║
║ Framework Technical Support:                                   ║
║   Contact: vba-framework-support@auditfirm.com                 ║
║   Slack Channel: #vba-framework-help                           ║
║   Documentation: [internal wiki link]                          ║
║                                                                ║
║ DataSnipper Support:                                           ║
║   Contact: support@datasnipper.com                             ║
║   Phone: +1-650-555-1234                                       ║
║   Known Issues: [reference internal KB]                        ║
║                                                                ║
║ Client IT Contact (for file access):                           ║
║   Name: IT Support Desk                                        ║
║   Email: itsupport@abccorp.com                                 ║
║   Phone: +1-555-123-4567                                       ║
║                                                                ║
║ Audit Partner (escalations):                                   ║
║   Name: Michael Zhang                                          ║
║   Email: michael.zhang@auditfirm.com                           ║
║   Phone: +1-555-999-8888                                       ║
║                                                                ║
╚════════════════════════════════════════════════════════════════╝
```

---

## Configuration Workflow

### Step-by-Step Setup Process

```
PHASE 1: PRE-ENGAGEMENT SETUP (1-2 hours)
  1. Create new workbook from Master_Template.xlsx
  2. Rename file: [EngagementID].xlsx
  3. Complete Section A (Engagement Metadata)
  4. Verify client documents available
  5. Fill in Section B (Document Specifications)

PHASE 2: TAG & QA DESIGN (2-4 hours)
  6. Complete Section C (TAG_ENGINE) - define extraction rules
  7. Review TAG_SPEC.md for proper tag syntax
  8. Complete Section D (QA Rules) - select/customize rules
  9. Verify against QA_RULES.md

PHASE 3: OUTPUT & WORKFLOW SETUP (1-2 hours)
  10. Complete Section E (Output Schema)
  11. Verify data types match source data
  12. Complete Section F (Approval Workflow)
  13. Verify approvers available

PHASE 4: VALIDATION (30 minutes)
  14. Run ConfigManager.LoadConfig() and ValidateConfig()
  15. Fix any configuration errors
  16. Proceed to tag generation
```

---

## Configuration Reuse & Templates

### Reusing Configuration for Similar Engagements

```
SCENARIO: Running Cash audit for multiple entities

STEP 1: Save as Template
  • Create initial engagement (2026-CASH-01)
  • Configure completely
  • Save as: CashTemplate_2026.xlsx

STEP 2: Use Template for New Entity
  • Copy CashTemplate_2026.xlsx
  • Rename: 2026-CASH-02.xlsx
  • Update Section A (Engagement Metadata)
    - Change Engagement ID
    - Update client name
    - Update lead auditor
  • Verify Section B file paths (point to new entity's docs)
  • Keep Sections C, D, E, F (same logic applies)

TIME SAVED: 3-4 hours per engagement
```

### Engagement Type Templates

```
PREDEFINED TEMPLATES PROVIDED:
  ├─ CashTemplate.xlsx          (cash cycle audits)
  ├─ ARTemplate.xlsx            (accounts receivable)
  ├─ APTemplate.xlsx            (accounts payable)
  ├─ ContractsTemplate.xlsx     (contract audits)
  └─ InventoryTemplate.xlsx     (inventory testing)

USAGE:
  1. Start with engagement type template
  2. Customize for specific client/period
  3. Save as engagement workbook
```

---

## Troubleshooting Configuration Issues

| Issue | Cause | Resolution |
|-------|-------|-----------|
| ConfigManager fails to load | Missing required field | Check CONFIG sheet; verify all mandatory fields have values |
| Tags not generating | Incomplete TAG_ENGINE entries | Verify field_name, extraction_method, and source fields populated |
| QA rules not applying | Rules not checked in CONFIG | Enable rules in Section D; verify syntax matches QA_RULES.md |
| Approval workflow blocked | Approver email invalid | Correct email in Section F; verify approver exists in directory |
| Output format incorrect | Schema mismatch | Review Section E; ensure data types match expected values |

---

**Framework Version**: 1.0  
**Last Updated**: June 2026
