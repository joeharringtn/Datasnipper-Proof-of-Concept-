# DATA_MODEL.md - Excel Schema & Data Structures

## Overview

This document defines the complete data model for all Excel sheets used by the framework. Each sheet has a specific role, schema, and interaction with VBA modules.

---

## Sheet Architecture

```
ENGAGEMENT WORKBOOK SHEETS:

Configuration Sheets (INPUT):
  ├─ CONFIG                    (engagement settings)
  └─ TAG_ENGINE                (extraction rule definitions)

Processing Sheets (WORKING):
  ├─ EXTRACTION_INPUT          (raw DataSnipper results)
  ├─ VALIDATION                (format validation results)
  ├─ QA                        (exceptions & manual overrides)
  └─ AUDIT_LOG                 (complete event trail)

Reference/Output Sheets:
  ├─ COORD_REFERENCE           (coordinate mapping reference)
  └─ OUTPUT                    (final approved results)
```

---

## Sheet 1: CONFIG

### Purpose
Master configuration sheet defining all engagement settings and parameters.

### Responsibility
- Owner: Audit Manager (completed once per engagement)
- Input: Manual configuration
- Output: Read by ConfigManager; used by all modules

### Schema

```
┌────────────────────────────────────────────────────────────────┐
│                      CONFIG SHEET                              │
├────────────────────────────────────────────────────────────────┤
│ Column A │ Column B         │ Type    │ Description            │
├──────────┼──────────────────┼─────────┼────────────────────────┤
│  A1      │ SECTION: ENGAGEMENT METADATA                       │
│  A2      │ Engagement ID    │ Text    │ Unique ID (e.g., 2026- │
│          │                  │         │ CASH-01)               │
│  A3      │ Engagement Type  │ Dropdown│ Cash|AR|AP|Contracts   │
│  A4      │ Period Start     │ Date    │ First day of period    │
│  A5      │ Period End       │ Date    │ Last day of period     │
│  A6      │ Lead Auditor     │ Text    │ Name of audit lead     │
│  A7      │ Client Name      │ Text    │ Company being audited  │
│  A8      │ Client Contact   │ Email   │ For coordination       │
│  A9      │                  │         │                        │
│  A10     │ SECTION: DOCUMENT SPECIFICATIONS                  │
│  A11     │ Doc_ID           │ Text    │ Short code (e.g.,      │
│          │                  │         │ BANK_STMT)             │
│  A12     │ File_Path        │ Text    │ Full path to document  │
│  A13     │ File_Name        │ Text    │ Filename only          │
│  A14     │ Page_Count       │ Number  │ Total pages            │
│  A15     │ Doc_Type         │ Dropdown│ Statement|Invoice|etc   │
│  A16     │ Extraction_Method│ Dropdown│ DS_SEARCH|DS_COORDS    │
│          │                  │         │                        │
│  A20     │ SECTION: TAG DEFINITIONS (see TAG_ENGINE)           │
│  A21     │ [Reference to TAG_ENGINE sheet]                    │
│          │                  │         │                        │
│  A30     │ SECTION: QA RULES                                 │
│  A31     │ QA_Rule_Set      │ Dropdown│ Standard|Custom|Hybrid │
│  A32     │ Rule_ID          │ Text    │ Predefined rule ID     │
│  A33     │ Enabled          │ Boolean │ Yes|No                 │
│  A34     │ Custom_Rules     │ Text    │ Space for custom logic │
│          │                  │         │                        │
│  A40     │ SECTION: OUTPUT SCHEMA                            │
│  A41     │ Field_Name       │ Text    │ Output column name     │
│  A42     │ Data_Type        │ Text    │ Currency|Date|Text     │
│  A43     │ Required         │ Boolean │ Yes|No                 │
│  A44     │ Format_Spec      │ Text    │ $#,##0.00 or pattern   │
│          │                  │         │                        │
│  A50     │ SECTION: APPROVAL WORKFLOW                        │
│  A51     │ Workflow_Type    │ Dropdown│ SIMPLE|MULTI_LEVEL     │
│  A52     │ Level1_Approver  │ Text    │ Approver name          │
│  A53     │ Level1_Email     │ Email   │ Approver email         │
│  A54     │ Level2_Approver  │ Text    │ Second approver (opt)  │
│  A55     │ Level2_Email     │ Email   │ Second approver email  │
│          │                  │         │                        │
└────────────────────────────────────────────────────────────────┘
```

### Key Columns

| Column | Required | Validation | Notes |
|--------|----------|------------|-------|
| Engagement ID | Yes | Unique; no spaces/special chars | Used throughout engagement |
| Period End Date | Yes | Must be valid date | Used by CROSS_FIELD rules |
| Lead Auditor | Yes | Text | Recorded in AUDIT_LOG |
| Doc_Path | Yes | Path must exist | Verified by ConfigManager |
| Extraction_Method | Yes | Must be DS_SEARCH or DS_COORDS | Determines tag format |
| QA_Rule_Set | Yes | Choose Standard or Custom | Defines validation behavior |

### Data Integrity Rules
```
✓ All required fields must have values
✓ Engagement ID must be unique (no duplicates)
✓ Period dates must be chronologically valid (start < end)
✓ Document paths must point to existing files
✓ Approver emails must be in valid format
✓ Data types must match their column definitions
```

---

## Sheet 2: TAG_ENGINE

### Purpose
Define extraction rules for each data field; build and validate DataSnipper tags.

### Responsibility
- Owner: Developer/SME (completed during tag design phase)
- Input: Manual definition of extraction rules
- Output: SourceTag column populated by TagBuilder
- Usage: Read by TagBuilder, QAEngine, DataMapper

### Schema

```
┌───────────────────────────────────────────────────────────────────┐
│                    TAG_ENGINE SHEET                               │
├───────────────────────────────────────────────────────────────────┤
│ Col │ Field Name         │ Type    │ Description                 │
├─────┼────────────────────┼─────────┼─────────────────────────────┤
│ A   │ Tag_ID             │ Text    │ Unique ID (e.g., CASH_DEP_01)
│ B   │ Output_Field       │ Text    │ Final output column name    │
│ C   │ Extraction_Method  │ Dropdown│ DS_SEARCH | DS_COORDS       │
│ D   │ Source_Document    │ Text    │ Doc_ID from CONFIG          │
│ E   │ Field_Type         │ Dropdown│ text | number | currency    │
│ F   │ Required           │ Boolean │ Yes | No                    │
│ G   │ Search_Keywords    │ Text    │ For DS_SEARCH: keywords/    │
│     │ (DS_SEARCH only)   │         │ anchors                     │
│ H   │ Start_Anchor       │ Text    │ For DS_SEARCH: "start="     │
│ I   │ End_Anchor         │ Text    │ For DS_SEARCH: "end="       │
│ J   │ Coord_Page         │ Number  │ For DS_COORDS: page number  │
│ K   │ Coord_X            │ Number  │ For DS_COORDS: x pixel pos  │
│ L   │ Coord_Y            │ Number  │ For DS_COORDS: y pixel pos  │
│ M   │ Coord_Width        │ Number  │ For DS_COORDS: width px     │
│ N   │ Coord_Height       │ Number  │ For DS_COORDS: height px    │
│ O   │ Tolerance          │ Number  │ For DS_COORDS: pixel tol    │
│ P   │ Fallback_Keywords  │ Text    │ For DS_SEARCH: alternatives │
│ Q   │ Notes              │ Text    │ Internal notes/documentation│
│ R   │ SourceTag          │ Text    │ GENERATED by TagBuilder     │
│     │ (OUTPUT)           │         │ (do not edit manually)      │
│ S   │ Tag_Status         │ Text    │ GENERATED: Valid|Error      │
│ T   │ Tag_Error_Message  │ Text    │ GENERATED: Error details    │
│     │                    │         │ (if validation fails)       │
│                                                                    │
└───────────────────────────────────────────────────────────────────┘
```

### Example Data

```
Row 1: Header row
Row 2: Deposit Amount | deposit_amount | DS_SEARCH | BANK_STMT | 
       currency | Yes | "Total|Amount" | "Total Amount:" | "Tax" | 
       (blank) | (blank) | (blank) | (blank) | (blank) | (blank) | 
       "Use bold text near total" | 
       [GENERATED: DS_SEARCH:DepAmount:deposit_amount:(start=Total ...] | 
       Valid | (blank)

Row 3: Deposit Date | deposit_date | DS_COORDS | BANK_STMT | 
       date | Yes | (blank) | (blank) | (blank) | 1 | 300 | 450 | 
       100 | 20 | 5 | (blank) | "Standard bank statement position" | 
       [GENERATED: DS_COORDS:DepDate:deposit_date:(page=1|x=300|...] | 
       Valid | (blank)

Row 4: Bank Name | bank_name | DS_SEARCH | DEPOSIT_DTL | 
       text | Yes | "Bank|Institution" | "Bank:" | "Account" | 
       (blank) | (blank) | (blank) | (blank) | (blank) | (blank) | 
       "May appear as 'Bank:' or 'Financial Institution:'" | 
       [GENERATED: DS_SEARCH:BankName:bank_name:(start=Bank|...] | 
       Valid | (blank)
```

### Module Interactions

| Module | Interaction | Direction |
|--------|-------------|-----------|
| **TagBuilder** | Reads columns A-Q; generates column R-T | Read → Write |
| **Validator** | Reads column E (field types) | Read |
| **QAEngine** | Reads column B (output fields) | Read |
| **DataMapper** | Reads column E (data types for conversion) | Read |

---

## Sheet 3: EXTRACTION_INPUT

### Purpose
Store raw DataSnipper extraction results for processing and validation.

### Responsibility
- Owner: Auditor (pastes DataSnipper results)
- Input: DataSnipper output (pasted from clipboard)
- Output: Used by Validator and DataMapper modules

### Schema

```
┌─────────────────────────────────────────────────────────────────┐
│               EXTRACTION_INPUT SHEET                            │
├─────────────────────────────────────────────────────────────────┤
│ Col │ Field Name         │ Type    │ Description               │
├─────┼────────────────────┼─────────┼───────────────────────────┤
│ A   │ Record_ID          │ Number  │ Unique row identifier     │
│ B   │ Extraction_Date    │ DateTime│ When extracted (auto)     │
│ C   │ Extracted_By       │ Text    │ User who ran extraction   │
│ [D-Z] │ [Output Fields]  │ Varies  │ One column per field      │
│ ZZ  │ Data_Quality_Flag  │ Text    │ Validation status        │
│ AAA │ Raw_Source_Page    │ Number  │ Which page extracted from│
│     │                    │         │ (reference)               │
│     │                    │         │                          │
│ EXAMPLE COLUMNS:                                               │
│ D   │ deposit_amount     │ Text    │ RAW value as extracted   │
│ E   │ deposit_date       │ Text    │ RAW value as extracted   │
│ F   │ bank_name          │ Text    │ RAW value as extracted   │
│ G   │ account_number     │ Text    │ RAW value as extracted   │
│                                                                 │
└─────────────────────────────────────────────────────────────────┘
```

### Key Characteristics

- **Format**: CSV or table format pasted from DataSnipper
- **Data Type**: ALL values initially TEXT (no conversion yet)
- **Nulls**: Empty cells or "N/A" indicate missing values
- **Immutable**: This sheet preserved as-is for audit trail
- **Row Count**: One row per extracted instance (e.g., one row per deposit)

### Example Data

```
Record_ID │ Extraction_Date │ Extracted_By │ deposit_amount │ dep_date │ bank_name
━━━━━━━━━━┿━━━━━━━━━━━━━━━━┿━━━━━━━━━━━━━━┿━━━━━━━━━━━━━━━┿━━━━━━━━━━┿━━━━━━━━━━
1         │ 6/1/26 14:30    │ John Smith   │ $1,250.00      │ 5/31/26  │ Chase
2         │ 6/1/26 14:30    │ John Smith   │ $500.00        │ 6/1/26   │ Chase
3         │ 6/1/26 14:35    │ John Smith   │ (blank)        │ 6/2/26   │ Bank of America
4         │ 6/1/26 14:35    │ John Smith   │ $5,000,000.00  │ 6/3/26   │ Unknown Bank
```

---

## Sheet 4: VALIDATION

### Purpose
Store results of input validation (format, completeness, data quality checks).

### Responsibility
- Owner: Validator module (auto-generated)
- Input: EXTRACTION_INPUT raw data
- Output: Read by DataMapper; used for error reporting

### Schema

```
┌──────────────────────────────────────────────────────────────┐
│                  VALIDATION SHEET                            │
├──────────────────────────────────────────────────────────────┤
│ Col │ Field Name         │ Type    │ Description            │
├─────┼────────────────────┼─────────┼────────────────────────┤
│ A   │ Record_ID          │ Number  │ Reference to EXTRACT   │
│ B   │ Field_Name         │ Text    │ Which field validated  │
│ C   │ Raw_Value          │ Text    │ Value as extracted     │
│ D   │ Validation_Type    │ Text    │ Format|Range|Required  │
│ E   │ Is_Valid           │ Boolean │ Yes | No               │
│ F   │ Error_Message      │ Text    │ Description of error   │
│ G   │ Severity           │ Text    │ INFO|WARNING|CRITICAL  │
│ H   │ Requires_Review    │ Boolean │ Yes | No (QA check?)   │
│ I   │ Timestamp          │ DateTime│ When validated         │
│                                                               │
└──────────────────────────────────────────────────────────────┘
```

### Example Data

```
Record_ID │ Field_Name    │ Raw_Value │ Validation_Type │ Is_Valid │ Error
──────────┼───────────────┼───────────┼─────────────────┼──────────┼─────────────────
1         │ deposit_amount│ $1,250.00 │ Format          │ Yes      │ (none)
1         │ deposit_date  │ 5/31/26   │ Format          │ Yes      │ (none)
2         │ bank_name     │ Chase     │ Format          │ Yes      │ (none)
3         │ deposit_amount│ (blank)   │ Required        │ No       │ Required field blank
4         │ deposit_amount│ $5000000  │ Range           │ No       │ Value exceeds max
```

---

## Sheet 5: QA

### Purpose
Store QA exceptions, manual overrides, and resolution decisions.

### Responsibility
- Owner: QA Team (manual review and decisions)
- Input: Exceptions flagged by QAEngine; manual overrides entered by QA
- Output: Used by Main for approval workflow; archived for audit trail

### Schema

```
┌────────────────────────────────────────────────────────────────┐
│                     QA SHEET                                   │
├────────────────────────────────────────────────────────────────┤
│ Col │ Field Name         │ Type    │ Description             │
├─────┼────────────────────┼─────────┼─────────────────────────┤
│ A   │ Exception_ID       │ Number  │ Auto-generated          │
│ B   │ Record_ID          │ Number  │ Reference to extraction │
│ C   │ Field_Name         │ Text    │ Which field            │
│ D   │ Raw_Value          │ Text    │ Extracted value        │
│ E   │ Exception_Reason   │ Text    │ Why flagged (rule name) │
│ F   │ Severity           │ Text    │ WARNING | CRITICAL      │
│ G   │ QA_Decision        │ Dropdown│ ACCEPT|OVERRIDE|REJECT │
│ H   │ Override_Value     │ Text    │ If OVERRIDE: new value │
│ I   │ Override_Justif.   │ Text    │ Why override was needed│
│ J   │ Reviewed_By        │ Text    │ QA team member name    │
│ K   │ Review_Date        │ DateTime│ When reviewed          │
│ L   │ QA_Lead_SignOff    │ Text    │ QA Lead name (final)   │
│ M   │ SignOff_Date       │ DateTime│ Final approval date    │
│ N   │ Notes              │ Text    │ Additional context     │
│                                                                │
└────────────────────────────────────────────────────────────────┘
```

### Example Data

```
Exc_ID │ Rec_ID │ Field       │ Raw_Val   │ Reason              │ Decision │ Override
───────┼────────┼─────────────┼───────────┼─────────────────────┼──────────┼──────────
1      │ 1      │ Dep_Amt     │ $1,250.00 │ (none - no exception)│ ACCEPT   │ (N/A)
2      │ 3      │ Dep_Amt     │ (blank)   │ Required field      │ OVERRIDE │ $750.00
3      │ 4      │ Dep_Amt     │ $5000000  │ Range exceeded      │ REJECT   │ (N/A)
```

### Decision Options

| Decision | Meaning | Impact | Used When |
|----------|---------|--------|-----------|
| **ACCEPT** | Exception is false positive | Value included in output | Data actually OK |
| **OVERRIDE** | Value is wrong; user enters correct | Override value used in output | QA found correct value |
| **REJECT** | Data unusable; exclude from output | Record excluded entirely | Value cannot be determined |

---

## Sheet 6: AUDIT_LOG

### Purpose
Complete audit trail of all processing events, decisions, and sign-offs.

### Responsibility
- Owner: AuditLog module (auto-generated)
- Input: All modules append events
- Output: Exported for external auditor review; retention per audit standards

### Schema

```
┌────────────────────────────────────────────────────────────────┐
│                  AUDIT_LOG SHEET                               │
├────────────────────────────────────────────────────────────────┤
│ Col │ Field Name         │ Type    │ Description             │
├─────┼────────────────────┼─────────┼─────────────────────────┤
│ A   │ Log_ID             │ Number  │ Auto-sequential         │
│ B   │ Timestamp          │ DateTime│ Date/time of event      │
│ C   │ User_ID            │ Text    │ Who performed action    │
│ D   │ Event_Type         │ Text    │ TRANSFORM|EXCEPTION|    │
│     │                    │         │ OVERRIDE|SIGNOFF|ERROR  │
│ E   │ Record_ID          │ Number  │ Which extraction record │
│ F   │ Field_Name         │ Text    │ Which field affected    │
│ G   │ Old_Value          │ Text    │ Before value/state      │
│ H   │ New_Value          │ Text    │ After value/state       │
│ I   │ Reason             │ Text    │ Why change/decision     │
│ J   │ Severity           │ Text    │ INFO|WARNING|CRITICAL   │
│ K   │ Module_Name        │ Text    │ Which module made entry │
│ L   │ Status             │ Text    │ Success|Error           │
│ M   │ Notes              │ Text    │ Additional context      │
│                                                                │
└────────────────────────────────────────────────────────────────┘
```

### Example Events

```
Log_ID │ Timestamp       │ User        │ Event_Type │ Record │ Field    │ Change
───────┼─────────────────┼─────────────┼────────────┼────────┼──────────┼──────────
1      │ 6/1/26 14:00    │ John Smith  │ TRANSFORM  │ 1      │ Dep_Amt  │ $1,250
2      │ 6/1/26 14:00    │ John Smith  │ TRANSFORM  │ 1      │ Dep_Date │ 5/31/26
3      │ 6/1/26 14:05    │ System      │ EXCEPTION  │ 3      │ Dep_Amt  │ FLAG
4      │ 6/1/26 15:30    │ Susan J     │ OVERRIDE   │ 3      │ Dep_Amt  │ $750.00
5      │ 6/1/26 16:00    │ Susan J     │ SIGNOFF    │ QA     │ (batch)  │ Approved
```

---

## Sheet 7: COORD_REFERENCE

### Purpose
Reference table for coordinate-based extractions; used for training and validation.

### Responsibility
- Owner: Developer (created during tag design)
- Input: Manual coordinate mapping from sample documents
- Output: Read-only reference; used for testing/validation

### Schema

```
┌──────────────────────────────────────────────────────────────┐
│              COORD_REFERENCE SHEET                           │
├──────────────────────────────────────────────────────────────┤
│ Col │ Field Name         │ Type    │ Description            │
├─────┼────────────────────┼─────────┼────────────────────────┤
│ A   │ Document_ID        │ Text    │ Which document (BANK_  │
│     │                    │         │ STMT, etc)             │
│ B   │ Output_Field       │ Text    │ Output column name     │
│ C   │ Page_Number        │ Number  │ Which page (1-based)   │
│ D   │ X_Position         │ Number  │ Pixels from left       │
│ E   │ Y_Position         │ Number  │ Pixels from top        │
│ F   │ Width              │ Number  │ Box width in pixels    │
│ G   │ Height             │ Number  │ Box height in pixels   │
│ H   │ Data_Type          │ Text    │ text|number|currency   │
│ I   │ Expected_Sample    │ Text    │ Sample value for ref   │
│ J   │ Notes              │ Text    │ Documentation notes    │
│                                                               │
└──────────────────────────────────────────────────────────────┘
```

### Purpose & Notes
- Used to document coordinate positions during tag design phase
- Serves as reference if coordinates need adjustment
- Helps new team members understand extraction logic
- Example: Sample value "$1,250.00" helps confirm position is correct

---

## Sheet 8: OUTPUT

### Purpose
Final, approved extraction results ready for downstream use.

### Responsibility
- Owner: Main module (generated from processed/approved data)
- Input: Accepted and overridden values from QA sheet
- Output: Exported as final deliverable; locked from editing

### Schema

```
┌──────────────────────────────────────────────────────────────┐
│                   OUTPUT SHEET                               │
├──────────────────────────────────────────────────────────────┤
│ Col │ Field Name         │ Type    │ Description            │
├─────┼────────────────────┼─────────┼────────────────────────┤
│ A   │ Record_ID          │ Number  │ Reference to extraction│
│ B   │ [Output Fields]    │ Varies  │ Per OUTPUT_SCHEMA      │
│     │                    │         │ (See CONFIG sheet)     │
│     │                    │         │                        │
│ EXAMPLE COLUMNS:                                             │
│ C   │ deposit_amount     │ Currency│ $1,250.00 (normalized) │
│ D   │ deposit_date       │ Date    │ 2026-05-31 (normalized)│
│ E   │ bank_name          │ Text    │ Chase Bank             │
│ F   │ extraction_status  │ Text    │ Approved|Overridden    │
│ G   │ qa_review_date     │ DateTime│ When finalized         │
│ H   │ output_notes       │ Text    │ Any final annotations  │
│                                                               │
└──────────────────────────────────────────────────────────────┘
```

### Characteristics
- **Locked**: Read-only after approval (prevents accidental changes)
- **Complete**: Includes all approved records (rejected records excluded)
- **Clean**: Only final values; no intermediate/working data
- **Audit Trail**: Includes metadata showing approval status

---

## Data Relationships & Flow

```
CONFIGURATION:
  CONFIG sheet
    ├─ Defines documents (file paths, pages)
    ├─ Defines extraction rules (TAG_ENGINE)
    ├─ Defines QA rules
    └─ Defines output schema

TAG GENERATION:
  TAG_ENGINE sheet
    ├─ Input: Manual tag definitions
    ├─ Process: TagBuilder generates syntax
    └─ Output: SourceTag column (DS_SEARCH/DS_COORDS)

DATA EXTRACTION:
  EXTRACTION_INPUT sheet
    ├─ Input: DataSnipper results (user pastes)
    ├─ Format: All text, raw
    └─ Status: Baseline for processing

VALIDATION:
  VALIDATION sheet
    ├─ Input: EXTRACTION_INPUT raw data
    ├─ Process: Format & completeness checks
    └─ Output: Error flags & warnings

PROCESSING:
  AUDIT_LOG sheet
    ├─ Input: All transformations & decisions
    ├─ Process: Continuous logging
    └─ Output: Complete event trail

QA REVIEW:
  QA sheet
    ├─ Input: Exceptions flagged by QAEngine
    ├─ Process: QA team reviews & decides
    └─ Output: ACCEPT | OVERRIDE | REJECT

FINAL OUTPUT:
  OUTPUT sheet
    ├─ Input: Accepted + overridden values
    ├─ Process: Format per schema; lock sheet
    └─ Output: Final deliverable
```

---

## Data Validation Rules (By Sheet)

### CONFIG Sheet
```
✓ Engagement_ID: Not blank; unique
✓ Period dates: start_date < end_date
✓ Document paths: Must exist and be readable
✓ Data types: Match column definitions
```

### TAG_ENGINE Sheet
```
✓ Output_Field: Not blank; matches OUTPUT_SCHEMA
✓ Extraction_Method: DS_SEARCH or DS_COORDS (not mixed in one row)
✓ For DS_SEARCH: start_anchor and end_anchor not both blank
✓ For DS_COORDS: page, x, y, width, height all numeric and positive
✓ Generated SourceTag: Must validate per TAG_SPEC syntax
```

### EXTRACTION_INPUT Sheet
```
✓ Record_ID: Unique identifier
✓ Extracted_By: Not blank
✓ Extraction_Date: Valid datetime
✓ Data values: Empty if not extracted (don't put "N/A" or "NULL")
```

### QA Sheet
```
✓ QA_Decision: ACCEPT, OVERRIDE, or REJECT (no blanks)
✓ If OVERRIDE: Override_Value populated; Justification present
✓ Reviewed_By: Not blank after review
✓ Review_Date: Populated after review
```

---

**Framework Version**: 1.0  
**Last Updated**: June 2026
