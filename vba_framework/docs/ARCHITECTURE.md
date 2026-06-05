# ARCHITECTURE.md - VBA Module Design & Interactions

## Module Overview

```
┌─────────────────────────────────────────────────────────────────┐
│                        MAIN ORCHESTRATOR                        │
│  (Entry point; UI coordination; workflow sequencing)            │
└────────────────────┬────────────────────────────────────────────┘
                     │
        ┌────────────┼────────────┬──────────────┬──────────┐
        ▼            ▼            ▼              ▼          ▼
   ┌─────────┐ ┌──────────┐ ┌──────────┐ ┌──────────┐ ┌──────────┐
   │ConfigMgr│ │TagBuilder│ │Validator │ │DataMapper│ │QAEngine  │
   └─────────┘ └──────────┘ └──────────┘ └──────────┘ └──────────┘
        │            │            │              │          │
        │            │            │              │          │
        └────────────┼────────────┼──────────────┼──────────┘
                     │            │              │
                     └────────────┼──────────────┘
                                  ▼
                          ┌──────────────┐
                          │  AuditLog    │
                          │  (Capture    │
                          │  all events) │
                          └──────────────┘
```

---

## 1. ConfigManager Module

### Responsibility
Load, validate, and expose engagement configuration for all other modules.

### What It Manages
- **Engagement Metadata**: Engagement ID, engagement type (Cash/A/R/A/P), period, auditor
- **Document Specifications**: File paths, page counts, document types
- **Tag Definitions**: What data to extract, from where, using which method
- **QA Rule Sets**: Engagement-specific validation rules
- **Output Schema**: Required fields and formats
- **User/Approver Info**: Who can sign off, escalation paths

### Key Operations
| Operation | Input | Output | Notes |
|-----------|-------|--------|-------|
| `LoadConfig()` | CONFIG sheet | ConfigObject | Loads all engagement settings into memory |
| `ValidateConfig()` | ConfigObject | Boolean + errors | Checks completeness; raises errors if missing required fields |
| `GetEngagementType()` | ConfigObject | String | Returns engagement category (used by other modules) |
| `GetDocSpecs()` | ConfigObject | DocumentSpecs[] | Returns file paths and page numbers |
| `GetTagDefinitions()` | ConfigObject | TagDefinition[] | Returns tag rules for TagBuilder |
| `GetQARules()` | ConfigObject | QARule[] | Returns QA rules for QAEngine |
| `GetOutputSchema()` | ConfigObject | SchemaField[] | Returns required output fields |

### Data Structure (In-Memory)
```
ConfigObject:
  ├─ engagement_id: String
  ├─ engagement_type: String (Cash | AR | AP | Contracts | Inventory)
  ├─ period_end_date: Date
  ├─ auditor: String
  ├─ approval_workflow: Enum (SIMPLE | MULTI_LEVEL | NONE)
  ├─ documents[]:
  │  ├─ doc_id: String
  │  ├─ file_path: String
  │  ├─ page_count: Integer
  │  ├─ doc_type: String (Invoice | PO | Receipt | etc)
  │  └─ extraction_method: Enum (DS_SEARCH | DS_COORDS | HYBRID)
  ├─ tags[]:
  │  ├─ tag_id: String
  │  ├─ field_name: String
  │  ├─ extraction_method: Enum
  │  ├─ search_keywords: String[] (for DS_SEARCH)
  │  ├─ coordinates: String (for DS_COORDS)
  │  └─ required: Boolean
  ├─ qa_rules[]:
  │  ├─ rule_id: String
  │  ├─ field_name: String
  │  ├─ validation_type: String (Range | Format | Lookup | Business Rule)
  │  └─ rule_definition: String (formula or description)
  └─ output_schema[]:
     ├─ field_name: String
     ├─ data_type: String (Text | Number | Date | Currency)
     ├─ required: Boolean
     └─ format_spec: String (yyyy-mm-dd, $#,##0.00, etc)
```

### Module Interactions
- **Called By**: Main, TagBuilder, QAEngine, DataMapper
- **Calls**: None (read-only from CONFIG sheet)
- **Error Handling**: Returns error codes if config validation fails

---

## 2. TagBuilder Module

### Responsibility
Generate DS_SEARCH and DS_COORDS tags for DataSnipper extraction.

### What It Does
- Reads tag definitions from ConfigManager
- Constructs valid DataSnipper tag syntax
- Outputs tags to a dedicated Excel column
- Provides tag preview and validation

### Key Operations
| Operation | Input | Output | Notes |
|-----------|-------|--------|-------|
| `BuildTags()` | TagDefinition[] | String[] | Generates all tags for current config |
| `BuildSearchTag()` | keyword[], operators | String | Builds DS_SEARCH tag with search terms |
| `BuildCoordTag()` | doc_id, page, x, y, width, height | String | Builds DS_COORDS tag for known positions |
| `BuildHybridTag()` | TagDefinition (mixed) | String | Builds combined search+coordinate tag |
| `ValidateTagSyntax()` | String[] | Boolean + errors | Checks tag format compliance |
| `ExportTagsToSheet()` | String[] | N/A | Writes tags to TAG_ENGINE output column |

### Tag Syntax Standards
See TAG_SPEC.md for complete DS_SEARCH and DS_COORDS syntax.

#### DS_SEARCH Example
```
DS_SEARCH:Vendor Invoice Number:InvoiceNum:(start=Invoice #|end=Date|type=text)
```

#### DS_COORDS Example
```
DS_COORDS:GR_Receipt_Qty:file=Receipt_2026.pdf|page=1|x=150|y=320|width=60|height=20|type=number
```

### Data Structures
```
TagDefinition:
  ├─ tag_id: String (unique key)
  ├─ field_name: String (output field name)
  ├─ extraction_method: Enum (DS_SEARCH | DS_COORDS | HYBRID)
  ├─ document_id: String (which source document)
  ├─ search_params: SearchParams (for DS_SEARCH)
  │  ├─ keywords: String[]
  │  ├─ start_anchor: String
  │  ├─ end_anchor: String
  │  └─ expected_type: String (text, number, currency, date)
  ├─ coord_params: CoordinateParams (for DS_COORDS)
  │  ├─ page_number: Integer
  │  ├─ x_position: Integer
  │  ├─ y_position: Integer
  │  ├─ width: Integer
  │  ├─ height: Integer
  │  └─ expected_type: String
  ├─ required: Boolean
  └─ notes: String (for user reference)
```

### Module Interactions
- **Called By**: Main
- **Calls**: ConfigManager.GetTagDefinitions(), ConfigManager.GetDocSpecs()
- **Outputs To**: TAG_ENGINE sheet (SourceTag column)
- **Error Handling**: Validates tag syntax; returns error list if generation fails

---

## 3. Validator Module

### Responsibility
Validate raw DataSnipper extractions against input requirements and data quality standards.

### What It Does
- Checks that required fields are populated
- Validates data formats (text, number, date, currency)
- Checks for outliers and anomalies
- Performs basic consistency checks
- Flags missing or malformed data
- Generates validation report

### Key Operations
| Operation | Input | Output | Notes |
|-----------|-------|--------|-------|
| `ValidateInputComplete()` | ExtractionInput[] | ValidationResult | Check all required fields populated |
| `ValidateDataFormat()` | ExtractionInput[], OutputSchema | ValidationResult | Check data types match schema |
| `ValidateDataRange()` | ExtractionInput[], RangeRules | ValidationResult | Check values within acceptable ranges |
| `CheckForNulls()` | ExtractionInput[] | String[] | List fields with null/blank values |
| `GenerateValidationReport()` | ValidationResult[] | Report | Summarize all validation issues |
| `FlagForReview()` | ValidationResult, issue_reason | N/A | Mark record for QA review |

### Validation Types
```
ValidationResult:
  ├─ record_id: String (which extraction row)
  ├─ field_name: String
  ├─ raw_value: String (as extracted)
  ├─ validation_check: String (type of validation)
  ├─ is_valid: Boolean
  ├─ error_message: String (if invalid)
  ├─ severity: Enum (INFO | WARNING | CRITICAL)
  └─ requires_review: Boolean
```

### Module Interactions
- **Called By**: Main
- **Calls**: ConfigManager.GetOutputSchema(), ConfigManager.GetDocSpecs()
- **Outputs To**: VALIDATION sheet
- **Error Handling**: Returns detailed validation errors; flags records for QA review

---

## 4. DataMapper Module

### Responsibility
Transform raw DataSnipper extractions into normalized, schema-compliant data.

### What It Does
- Maps raw extracted values to final output field names
- Applies data type conversions (text → number, string → date, etc.)
- Normalizes formats (currency symbols, date formats, leading zeros, etc.)
- Handles null/blank values per schema rules
- Creates audit trail of transformations

### Key Operations
| Operation | Input | Output | Notes |
|-----------|-------|--------|-------|
| `MapRawToSchema()` | ExtractionInput[], OutputSchema | MappedData[] | Transform raw extractions to output schema |
| `ConvertDataType()` | RawValue, TargetType | ConvertedValue | Convert single value to target type |
| `NormalizeFormat()` | Value, FormatSpec | NormalizedValue | Apply format spec (currency, date, etc) |
| `HandleNullValues()` | Value, NullHandlingRule | Result | Apply null handling logic per schema |
| `CreateAuditTrail()` | RawValue, TransformedValue | AuditEntry | Record transformation for audit log |

### Data Transformation Rules
```
Example: Vendor Invoice Amount

Raw (from DataSnipper):  "$1,250.00"
Target Schema:           Currency, Format: $#,##0.00
Target Field:            invoice_amount
Transformation:
  1. Detect currency symbol ($)
  2. Remove non-numeric characters (,)
  3. Convert to Number: 1250.00
  4. Apply format spec: $1,250.00
  5. Log transformation in AuditLog
  
Result:                  1250.00 (numeric value)
Display Format:          $1,250.00
```

### Module Interactions
- **Called By**: Main
- **Calls**: ConfigManager.GetOutputSchema(), Validator, AuditLog.RecordTransformation()
- **Outputs To**: Intermediate data structure (fed to QAEngine)
- **Error Handling**: Returns conversion errors; flags non-convertible values for manual review

---

## 5. QAEngine Module

### Responsibility
Apply business logic validation rules and identify exceptions requiring human review.

### What It Does
- Applies engagement-specific QA rules (ranges, lookups, business logic)
- Cross-field validation (e.g., invoice date must be within period)
- Performs duplicate detection
- Checks internal consistency rules
- Flags exceptions for manual review
- Generates QA exception report

### Key Operations
| Operation | Input | Output | Notes |
|-----------|-------|--------|-------|
| `ApplyQARules()` | MappedData[], QARule[] | QAResult[] | Apply all engagement rules |
| `CheckRangeValidation()` | Value, RangeRule | Boolean + error | Value within min/max range |
| `CheckLookupValidation()` | Value, LookupTable | Boolean + error | Value exists in reference table |
| `CheckBusinessLogic()` | Record, LogicRule | Boolean + error | Apply custom business rule |
| `CheckDuplicates()` | MappedData[] | DuplicateRecord[] | Identify duplicate extractions |
| `GenerateQAReport()` | QAResult[] | Report | Summarize all exceptions |
| `FlagExceptionForReview()` | QAResult, reason | N/A | Mark for human review in QA sheet |

### QA Rule Types
```
QARule:
  ├─ rule_id: String
  ├─ rule_name: String
  ├─ field_name: String (or multiple fields for cross-field rules)
  ├─ rule_type: Enum:
  │  ├─ RANGE: (min <= value <= max)
  │  ├─ LOOKUP: (value exists in reference table)
  │  ├─ FORMAT: (value matches regex pattern)
  │  ├─ CROSS_FIELD: (field1 + field2 consistency logic)
  │  ├─ DUPLICATE: (no duplicates on key fields)
  │  └─ CUSTOM: (engagement-specific business logic)
  ├─ rule_definition: String (formula or description)
  ├─ fail_action: Enum (FLAG | BLOCK | WARN)
  └─ exception_reason: String (why user should review)
```

### Example QA Rules (Cash Engagement)
```
Rule 1: Deposit Amount Range
  ├─ field_name: deposit_amount
  ├─ rule_type: RANGE
  ├─ min_value: 0
  ├─ max_value: 1,000,000
  └─ fail_action: FLAG (alert user if outside range)

Rule 2: Deposit Date Validity
  ├─ field_name: deposit_date
  ├─ rule_type: CROSS_FIELD
  ├─ rule: (deposit_date >= period_start AND deposit_date <= period_end)
  └─ fail_action: FLAG

Rule 3: Bank Name Validation
  ├─ field_name: bank_name
  ├─ rule_type: LOOKUP
  ├─ lookup_table: APPROVED_BANKS reference sheet
  └─ fail_action: BLOCK (don't allow unrecognized banks)
```

### Module Interactions
- **Called By**: Main
- **Calls**: ConfigManager.GetQARules(), AuditLog.RecordException()
- **Outputs To**: QA sheet (exception list)
- **Error Handling**: Returns exception list; blocks processing if CRITICAL exceptions exist

---

## 6. AuditLog Module

### Responsibility
Capture all processing events, exceptions, overrides, and decisions for audit trail compliance.

### What It Does
- Records all data transformations
- Logs all QA exceptions and resolutions
- Captures manual overrides with justification
- Records user actions and sign-offs
- Generates audit trail report for external review

### Key Operations
| Operation | Input | Output | Notes |
|-----------|-------|--------|-------|
| `RecordTransformation()` | RawValue, TransformedValue, field, reason | N/A | Log data transformation |
| `RecordException()` | ExceptionDetail, severity | N/A | Log QA exception |
| `RecordManualOverride()` | FieldName, OrigValue, NewValue, Justification, UserId | N/A | Log manual QA override |
| `RecordSignOff()` | Approver, Role, DateTime, Status | N/A | Log approval decision |
| `GenerateAuditTrail()` | DateRange | AuditTrailReport | Produce full audit report |
| `ExportAuditLog()` | N/A | CSV/Excel | Export audit log for external review |

### Audit Log Schema
```
AuditLogEntry:
  ├─ log_id: String (unique key)
  ├─ timestamp: DateTime
  ├─ user_id: String
  ├─ event_type: Enum (TRANSFORM | EXCEPTION | OVERRIDE | SIGNOFF | ERROR)
  ├─ record_id: String (which extraction row)
  ├─ field_name: String
  ├─ old_value: String
  ├─ new_value: String
  ├─ reason: String
  ├─ severity: Enum (INFO | WARNING | CRITICAL)
  └─ notes: String
```

### Module Interactions
- **Called By**: DataMapper, QAEngine, Main (for sign-offs)
- **Calls**: None (write-only to AUDIT_LOG sheet)
- **Outputs To**: AUDIT_LOG sheet + optional external export
- **Error Handling**: All logging errors are caught; processing continues (logging never blocks)

---

## 7. Main Module

### Responsibility
Orchestrate workflow; manage UI; coordinate module execution.

### What It Does
- Provides entry points for user actions (buttons in Excel ribbon)
- Calls modules in correct sequence
- Manages state between module calls
- Provides user feedback and progress indicators
- Handles error recovery and rollback
- Manages approval workflows

### Key Operations (Entry Points / Macros)
| Macro Name | Purpose | Calls | Output |
|-----------|---------|-------|--------|
| `LoadEngagement()` | Load config; initialize workbook | ConfigManager.LoadConfig() | CONFIG sheet populated in memory |
| `BuildTags()` | Generate DataSnipper tags | TagBuilder.BuildTags() | TAG_ENGINE.SourceTag column populated |
| `ValidateInput()` | Validate raw extractions | Validator.ValidateInputComplete() | VALIDATION sheet populated |
| `ProcessExtraction()` | Transform & validate raw data | DataMapper, QAEngine | QA sheet populated with exceptions |
| `ReviewExceptions()` | Manual QA review workflow | (User interactive) | QA sheet marked as reviewed |
| `ApproveOutput()` | Final sign-off | AuditLog.RecordSignOff() | OUTPUT sheet locked; audit trail recorded |
| `ExportResults()` | Export to downstream system | (Write to OUTPUT sheet) | Final data exported |
| `GenerateReport()` | Create audit compliance report | AuditLog.GenerateAuditTrail() | Report generated |

### Workflow State Machine
```
IDLE
  ├─ (LoadEngagement) → CONFIG_LOADED
  ├─ (BuildTags) [requires CONFIG_LOADED] → TAGS_READY
  ├─ (SubmitExtraction) [requires TAGS_READY] → EXTRACTION_SUBMITTED
  ├─ (ValidateInput) [requires EXTRACTION_SUBMITTED] → VALIDATION_COMPLETE
  ├─ (ProcessExtraction) [requires VALIDATION_COMPLETE] → QA_IN_PROGRESS
  ├─ (ReviewExceptions) [requires QA_IN_PROGRESS] → QA_REVIEWED
  ├─ (ApproveOutput) [requires QA_REVIEWED] → COMPLETE
  └─ (GenerateReport) [requires COMPLETE] → ARCHIVED
```

### Error Handling in Main
- **Catch & Log**: All module errors caught and logged
- **User Notification**: Show error dialog with actionable message
- **Rollback**: Revert changes if critical error in pipeline
- **Recovery**: Allow user to fix error and retry

### Module Interactions
- **Called By**: User (via ribbon buttons / menu)
- **Calls**: All other modules in sequence
- **Error Handling**: Central error handler; catches and logs all exceptions

---

## Module Interaction Matrix

```
           │Config│TagBldr│Validtr│DataMpr│QAEng│AuditL│Main
───────────┼──────┼───────┼───────┼───────┼─────┼──────┼─────
ConfigMgr  │  N   │   X   │   X   │   X   │  X  │  -   │  X
TagBuilder │  X   │   N   │   -   │   -   │  -  │  -   │  X
Validator  │  X   │   -   │   N   │   X   │  -  │  X   │  X
DataMapper │  X   │   -   │   X   │   N   │  X  │  X   │  X
QAEngine   │  X   │   -   │   -   │   X   │  N  │  X   │  X
AuditLog   │  -   │   -   │   -   │   X   │  X  │  N   │  X
Main       │  X   │   X   │   X   │   X   │  X  │  X   │  N

X = calls; - = no interaction
```

---

## Data Flow Diagram

```
PHASE 1: CONFIGURATION
┌──────────────┐
│ CONFIG Sheet │ ──┐
└──────────────┘   │
                   ▼
             ┌─────────────┐
             │ ConfigMgr   │ (validates, exposes config)
             └─────────────┘
                   │
      ┌────────────┼────────────────────┐
      │            │                    │
      ▼            ▼                    ▼
   (to)Tag     (to)QAEngine        (to)DataMapper
   Builder

PHASE 2: TAG BUILDING
┌─────────────────────┐
│ TagDefinition[] from │
│   ConfigManager     │
└─────────────────────┘
         │
         ▼
    ┌──────────────┐
    │ TagBuilder   │ (constructs DS_SEARCH/DS_COORDS)
    └──────────────┘
         │
         ▼
   ┌──────────────────┐
   │ TAG_ENGINE Sheet │ (user copies tags to DataSnipper)
   └──────────────────┘

PHASE 3: EXTRACTION INPUT
┌──────────────────────────┐
│ EXTRACTION_INPUT Sheet   │ (user pastes DataSnipper results)
└──────────────────────────┘
         │
         ▼
    ┌──────────────┐
    │  Validator   │ (checks completeness, format)
    └──────────────┘
         │
         ▼
   ┌──────────────────┐
   │ VALIDATION Sheet │ (error/warning flags)
   └──────────────────┘

PHASE 4: TRANSFORMATION & QA
┌──────────────────────────┐
│ Raw Extraction Input     │
└──────────────────────────┘
    │
    ├─────────────────────────────┐
    │                             │
    ▼                             ▼
┌──────────────┐          ┌──────────────┐
│ DataMapper   │          │  Validator   │ (format validation)
│(transform to │          └──────────────┘
│ schema)      │                 │
└──────────────┘                 │
    │                            │
    └────────────┬───────────────┘
                 ▼
          ┌─────────────┐
          │ QAEngine    │ (apply business rules)
          └─────────────┘
                 │
                 ▼
           ┌──────────┐
           │ QA Sheet │ (exceptions for manual review)
           └──────────┘

PHASE 5: QA REVIEW
┌──────────────┐
│ QA Sheet     │ (user reviews exceptions)
└──────────────┘
    │
    ├─────────────────┐
    │                 │
    ▼                 ▼
┌─────────────┐  ┌──────────────┐
│ Manual      │  │  AuditLog    │ (capture override with justification)
│ Override    │  └──────────────┘
└─────────────┘

PHASE 6: FINAL OUTPUT
┌──────────────┐
│ OUTPUT Sheet │ (final, cleansed, approved data)
└──────────────┘
    │
    ├─────────────────────┐
    │                     │
    ▼                     ▼
┌──────────────────┐  ┌──────────────┐
│ Export Results   │  │ AUDIT_LOG    │ (full trail for compliance)
└──────────────────┘  └──────────────┘
```

---

## Module Implementation Guidelines

### Coding Standards
- **Function Naming**: CamelCase for public functions (e.g., `BuildTags()`)
- **Variable Naming**: snake_case for module-level variables; camelCase for local variables
- **Error Codes**: Return codes 0 (success), 1-99 (validation errors), 100+ (system errors)
- **Comments**: Document each function's inputs, outputs, side effects
- **Logging**: Call AuditLog for all significant events

### Dependency Management
- **Avoid Circular Dependencies**: Only downward calls (Main → ConfigManager, not vice versa)
- **Dependency Injection**: Pass configuration objects as parameters, not global state
- **Testability**: Each module should be testable in isolation with mock dependencies

### Error Handling Pattern
```
Each module should implement:
  1. Input validation (check parameters)
  2. Business logic validation
  3. Return status code or error object
  4. Call AuditLog for significant errors
  5. Never throw unhandled errors (catch and log)
```

---

**Framework Version**: 1.0  
**Last Updated**: June 2026
