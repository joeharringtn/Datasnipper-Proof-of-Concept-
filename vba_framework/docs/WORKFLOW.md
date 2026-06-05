# WORKFLOW.md - End-to-End Process Flows

## High-Level Process Map

```
┌────────────────────────────────────────────────────────────────────────┐
│                     DATASNIPPER ENGAGEMENT WORKFLOW                    │
└────────────────────────────────────────────────────────────────────────┘

SETUP PHASE (Pre-Engagement)
┌─────────────────────────────────────────────────────────┐
│ 1. Audit manager selects engagement type (Cash/AR/AP)   │
│ 2. Creates workbook from Master_Template.xlsx           │
│ 3. Completes CONFIG sheet with engagement details       │
│    • Engagement ID, period, document file paths        │
│    • Selects QA rule set (standard or custom)          │
│ 4. Developer or SME populates TAG_ENGINE sheet         │
│    • Defines extraction rules for each data point       │
│    • Decides DS_SEARCH vs DS_COORDS for each tag       │
└─────────────────────────────────────────────────────────┘
                         │
                         ▼
BUILD PHASE (Tag Generation)
┌─────────────────────────────────────────────────────────┐
│ 5. Developer runs "Build Tags" macro                    │
│    • TagBuilder reads CONFIG + TAG_ENGINE              │
│    • Generates DS_SEARCH and DS_COORDS tags            │
│    • Validates tag syntax                              │
│    • Outputs to TAG_ENGINE.SourceTag column            │
│ 6. Developer/SME reviews generated tags for accuracy   │
│    • Adjusts tag parameters if needed                  │
│    • Tests tag logic manually if required              │
└─────────────────────────────────────────────────────────┘
                         │
                         ▼
SNIPPING PHASE (DataSnipper Extraction - EXTERNAL)
┌─────────────────────────────────────────────────────────┐
│ 7. Auditor copies DS_SEARCH and DS_COORDS tags from    │
│    TAG_ENGINE.SourceTag column                         │
│ 8. Auditor opens DataSnipper client                    │
│    • Pastes tags into DataSnipper UI                   │
│    • Loads documents (PDF/Image files)                 │
│    • Runs extraction against documents                 │
│    • DataSnipper returns extracted values              │
│ 9. Auditor copies extracted values                     │
│   (DataSnipper places results in clipboard)            │
│ 10. Auditor pastes results into Excel                  │
│    • Pastes into EXTRACTION_INPUT sheet                │
│    • Records who performed snipping, when             │
└─────────────────────────────────────────────────────────┘
                         │
                         ▼
VALIDATION PHASE (Input Validation)
┌─────────────────────────────────────────────────────────┐
│ 11. User clicks "Validate Input" macro                 │
│     • Validator module runs                            │
│     • Checks all required fields populated             │
│     • Validates data formats (text/number/date/curr)   │
│     • Flags missing or malformed data                  │
│     • Results written to VALIDATION sheet              │
│ 12. If critical errors exist:                          │
│     • Macro stops; user fixes errors in EXTRACTION_IN  │
│     • Re-runs "Validate Input"                         │
│ 13. If all validations pass:                           │
│     • Proceeds to next phase                           │
└─────────────────────────────────────────────────────────┘
                         │
                         ▼
TRANSFORMATION PHASE (Data Normalization)
┌─────────────────────────────────────────────────────────┐
│ 14. User clicks "Process Extraction" macro             │
│     • DataMapper transforms raw values to schema       │
│       - Converts text to numbers (e.g., "$1,250.00")   │
│       - Converts to date format (e.g., "06/01/2026")   │
│       - Applies normalizations per schema spec        │
│     • Validator runs format checks                     │
│     • Results stored in intermediate data structure    │
│ 15. QAEngine applies business logic rules              │
│     • Range checks (deposit $0 - $1M)                  │
│     • Lookup validation (bank name in approved list)   │
│     • Cross-field consistency (date in period)         │
│     • Duplicate detection                              │
│     • Exceptions flagged in QA sheet                   │
│ 16. AuditLog records all transformations & exceptions  │
└─────────────────────────────────────────────────────────┘
                         │
                         ▼
QA REVIEW PHASE (Exception Management)
┌─────────────────────────────────────────────────────────┐
│ 17. QA team reviews exceptions in QA sheet             │
│     • Each exception shows:                            │
│       - Field name and extracted value                 │
│       - Reason for flag (range, lookup, rule, etc)     │
│       - Suggested action (accept/override/reject)      │
│ 18. For each exception, QA team decides:              │
│     OPTION A: Accept (data is correct as-is)          │
│       • Mark as "Accepted" in QA sheet                 │
│       • AuditLog records decision                      │
│     OPTION B: Override (change value with justif.)     │
│       • Enter correct value in QA sheet                │
│       • Enter override justification                   │
│       • AuditLog records old/new value and reason      │
│     OPTION C: Reject (data unusable)                   │
│       • Mark as "Rejected"                             │
│       • Enter rejection reason                         │
│       • Item will not appear in OUTPUT sheet           │
│ 19. QA team leader signs off on all decisions          │
│     • Records name, role, date/time in QA sheet        │
│     • AuditLog records sign-off event                  │
└─────────────────────────────────────────────────────────┘
                         │
                         ▼
APPROVAL PHASE (Final Sign-Off)
┌─────────────────────────────────────────────────────────┐
│ 20. User clicks "Approve & Finalize" macro             │
│     • All exceptions must be reviewed (no blanks)      │
│     • If approval workflow enabled:                    │
│       - Requires sign-off from designated approver     │
│       - Records approver name, title, date/time        │
│     • OUTPUT sheet is generated                        │
│       - Accepted and overridden values copied          │
│       - Rejected values excluded                       │
│       - OUTPUT sheet locked from editing              │
│     • AuditLog records final approval                  │
│ 21. If any required approvals missing:                 │
│     • Macro stops with error message                   │
│     • User obtains required signature/sign-off         │
│     • Reruns approval macro                            │
└─────────────────────────────────────────────────────────┘
                         │
                         ▼
EXPORT PHASE (Downstream Use)
┌─────────────────────────────────────────────────────────┐
│ 22. User clicks "Export Results" macro                 │
│     • OUTPUT sheet exported to CSV/Excel for analysis  │
│     • Optional: Export to downstream system (GL, DB)   │
│ 23. User clicks "Generate Audit Report"                │
│     • AUDIT_LOG sheet exported in full                 │
│     • Report includes:                                 │
│       - Data transformations (raw → final)            │
│       - All exceptions and resolutions                 │
│       - User actions and timestamps                    │
│       - Sign-offs and approvals                        │
│       - Complete chain of custody                      │
│ 24. Engagement archive stored                          │
│     • Workbook + audit trail preserved for review      │
└─────────────────────────────────────────────────────────┘
```

---

## Phase 1: Setup & Configuration

### Step 1.1: Create Engagement Workbook
```
INPUT:
  • Audit manager selects engagement type from dropdown
  • Selection: "Cash", "A/R", "A/P", "Contracts", "Inventory"

PROCESS:
  • Framework copies Master_Template.xlsx
  • Renames with engagement ID (e.g., "2026-CASH-01.xlsx")
  • All standard sheets created (CONFIG, TAG_ENGINE, etc)

OUTPUT:
  • New workbook ready for configuration
```

### Step 1.2: Complete CONFIG Sheet
```
INPUT (Audit Manager):
  ┌────────────────────────────────────────┐
  │ Engagement Metadata Section            │
  ├────────────────────────────────────────┤
  │ Engagement ID:        2026-CASH-01      │
  │ Engagement Type:      Cash              │
  │ Period End Date:      06/30/2026        │
  │ Auditor Name:         John Smith        │
  │ Approval Workflow:    MULTI_LEVEL       │
  │ Approver (Level 1):   Susan Johnson     │
  │ Approver (Level 2):   Robert Chen       │
  └────────────────────────────────────────┘
  
  ┌────────────────────────────────────────┐
  │ Document Specifications Section        │
  ├────────────────────────────────────────┤
  │ File Path 1:  C:\\Docs\\Bank_Stmt_Jun.pdf
  │ Page Count:   25                        │
  │ Doc Type:     Bank Statement            │
  │                                         │
  │ File Path 2:  C:\\Docs\\Deposits.pdf   │
  │ Page Count:   12                        │
  │ Doc Type:     Deposit Details           │
  └────────────────────────────────────────┘

PROCESS:
  • ConfigManager validates completeness
  • Checks all required fields populated
  • Validates date formats, file paths exist
  • Confirms approvers have access

OUTPUT:
  • Configuration validated and ready
  • ERROR: If required field missing, shows error dialog
```

### Step 1.3: Define Extraction Tags (TAG_ENGINE)
```
INPUT (Developer/SME):
  
  Manual entry into TAG_ENGINE sheet:
  
  ┌─────────────────────────────────────────────────────────┐
  │ Row │ Field Name      │ Extraction │ Search Keywords    │
  ├─────┼─────────────────┼────────────┼────────────────────┤
  │  1  │ Deposit Amount  │ DS_SEARCH  │ Total|Amount       │
  │  2  │ Deposit Date    │ DS_COORDS  │ [page 1, x:100...] │
  │  3  │ Bank Name       │ DS_SEARCH  │ Bank|Institution   │
  │  4  │ Account Number  │ DS_COORDS  │ [page 2, x:50...]  │
  └─────┴─────────────────┴────────────┴────────────────────┘

PROCESS:
  • Developer manually populates TAG_ENGINE columns:
    - tag_id (unique identifier)
    - field_name (output field name)
    - extraction_method (DS_SEARCH or DS_COORDS)
    - source_keywords or coordinates
    - required (yes/no)
  • ConfigManager reads and validates entries

OUTPUT:
  • TAG_ENGINE sheet complete with extraction rules
  • Ready for TagBuilder to generate actual tags
```

---

## Phase 2: Tag Generation

### Step 2.1: Run TagBuilder
```
TRIGGER: User clicks "Build Tags" button

PROCESS:
  1. Main.BuildTags() called
  2. ConfigManager.LoadConfig() loads all engagement settings
  3. TagBuilder.BuildTags() executes:
     FOR each tag definition in TAG_ENGINE:
       IF extraction_method = "DS_SEARCH":
         TagBuilder.BuildSearchTag() → generates tag string
       ELSE IF extraction_method = "DS_COORDS":
         TagBuilder.BuildCoordTag() → generates coordinate tag
       ENDIF
       TagBuilder.ValidateTagSyntax() → checks syntax validity
       IF syntax invalid:
         Error record created; logged
       ELSE:
         Tag string stored in output array
  4. TagBuilder.ExportTagsToSheet() writes to TAG_ENGINE.SourceTag
  5. AuditLog records tag generation event

OUTPUT:
  ┌──────────────────────────────────────────────────────┐
  │ TAG_ENGINE Sheet - SourceTag Column (Generated)      │
  ├──────────────────────────────────────────────────────┤
  │ DS_SEARCH:Deposit Amount:DepAmount:(start=Total|    │
  │ end=Date|type=currency)                              │
  │                                                      │
  │ DS_COORDS:Deposit_Date:file=Bank_Stmt.pdf|page=1|   │
  │ x=150|y=320|width=100|height=20|type=date           │
  │                                                      │
  │ DS_SEARCH:Bank Name:BankName:(start=Bank|           │
  │ end=Account|type=text)                               │
  └──────────────────────────────────────────────────────┘

ERRORS DURING TAG GENERATION:
  IF syntax error found:
    - Error logged in VALIDATION sheet
    - User notified; provided corrected tag
    - Retry tag generation
```

### Step 2.2: Manual Review & Adjustment
```
ROLE: Developer/SME

REVIEW CHECKLIST:
  [ ] All required fields have tags
  [ ] DS_SEARCH tags include appropriate keywords/anchors
  [ ] DS_COORDS tags reference correct pages/positions
  [ ] Tag syntax is valid per TAG_SPEC.md
  [ ] Tags tested against sample documents (if available)

ADJUSTMENTS ALLOWED:
  • Edit source keywords in TAG_ENGINE sheet
  • Adjust coordinate positions if needed
  • Change extraction method if warranted
  • Add/remove fields as needed

AFTER ADJUSTMENT:
  • Re-run "Build Tags" to regenerate with new parameters
  • Revalidate tag syntax
  • Confirm tags ready for DataSnipper
```

---

## Phase 3: DataSnipper Extraction (EXTERNAL)

**NOTE: This phase runs OUTSIDE the framework—in DataSnipper UI**

### Step 3.1: Prepare DataSnipper
```
USER: Auditor/Fieldwork team

STEPS:
  1. Open DataSnipper client application
  2. Open "New Extraction" dialog
  3. Load source documents (PDFs, images, scanned documents)
  4. Set document page navigation as needed
```

### Step 3.2: Load Tags & Run Extraction
```
STEPS:
  1. In Excel: Select all generated tags from TAG_ENGINE.SourceTag
  2. Copy tags to clipboard (Ctrl+C)
  3. In DataSnipper: Paste tags into tag field (Ctrl+V)
  4. DataSnipper parses tags and configures extraction
  5. Click "Run Extraction" in DataSnipper
  6. DataSnipper processes each tag:
     FOR each tag:
       IF DS_SEARCH:
         DataSnipper finds keyword, extracts context
       ELSE IF DS_COORDS:
         DataSnipper navigates to page/coordinates, extracts value
       ENDIF
  7. DataSnipper returns extracted values (clipboard or download)
```

### Step 3.3: Capture Results & Return to Excel
```
STEPS:
  1. In DataSnipper: Select "Copy Results" or download CSV
  2. In Excel: Click on EXTRACTION_INPUT sheet
  3. Paste extracted values (Ctrl+V) into input columns
  4. Record metadata:
     - User who ran extraction
     - Date/time of extraction
     - DataSnipper version used (if tracking)
     - Any manual adjustments made
```

---

## Phase 4: Validation

### Step 4.1: Input Validation
```
TRIGGER: User clicks "Validate Input" button

PROCESS:
  1. Main.ValidateInput() called
  2. Validator module executes:
     FOR each extracted value in EXTRACTION_INPUT:
       Validator.CheckForNulls() → flag empty required fields
       Validator.ValidateDataFormat() → type checking
         • Currency formatted as text or number?
         • Dates in expected format?
         • Text values non-empty?
       Validator.ValidateDataRange() → range checks
         • Numbers within reasonable bounds?
       Store results in ValidationResult array
  3. Validator.GenerateValidationReport()
  4. Results written to VALIDATION sheet
  5. AuditLog records validation event

OUTPUT - VALIDATION Sheet:
  ┌────────────────────────────────────────────────────┐
  │ Row │ Field      │ Raw Value │ Error           │ Sev│
  ├────────────────────────────────────────────────────┤
  │  1  │ Dep_Amt    │ $1,250.00 │ (none)          │ OK │
  │  2  │ Dep_Date   │ (blank)   │ Required field  │ CRT│
  │  3  │ Bank_Name  │ Chase     │ (none)          │ OK │
  └────────────────────────────────────────────────────┘

NEXT STEPS:
  IF critical errors:
    • Stop processing
    • User corrects errors in EXTRACTION_INPUT
    • Re-run "Validate Input"
  ELSE:
    • Proceed to Phase 5 (Transformation)
```

---

## Phase 5: Transformation & QA Rules

### Step 5.1: Data Transformation
```
TRIGGER: User clicks "Process Extraction" button

PROCESS:
  1. Main.ProcessExtraction() called
  2. DataMapper module executes:
     FOR each validated value:
       DataMapper.ConvertDataType()
         • Convert "$1,250.00" (text) → 1250.00 (number)
         • Convert "6/1/26" → "2026-06-01" (normalized date)
       DataMapper.NormalizeFormat()
         • Apply format spec: currency → $#,##0.00
       DataMapper.HandleNullValues()
         • Apply null handling per schema
       DataMapper.CreateAuditTrail()
         • Log transformation: raw → transformed
     STORE results in intermediate data structure

OUTPUT - Intermediate Data:
  ┌──────────────────────────────────────┐
  │ Deposit Amount: 1250.00 (normalized) │
  │ Deposit Date: 2026-06-01 (normalized)│
  │ Bank Name: Chase (unchanged)         │
  └──────────────────────────────────────┘

3. Validator.ValidateDataFormat() runs format checks
4. QAEngine.ApplyQARules() applies business logic
```

### Step 5.2: QA Rules Application
```
PROCESS:
  QAEngine reads QA rules for engagement type (Cash)
  
  FOR each extracted record:
    FOR each QA rule:
      
      Rule 1: Deposit Amount Range
        IF deposit_amount < 0 OR deposit_amount > 1,000,000:
          Flag: "Deposit amount outside normal range"
          Severity: WARNING
      
      Rule 2: Deposit Date Validity
        IF deposit_date < period_start OR deposit_date > period_end:
          Flag: "Deposit date outside audit period"
          Severity: CRITICAL
      
      Rule 3: Bank Name Validation
        IF bank_name NOT IN approved_banks_table:
          Flag: "Bank name not in approved vendor list"
          Severity: CRITICAL (requires override)
      
      Rule 4: Duplicate Detection
        IF EXISTS(deposit_amount=X AND deposit_date=Y):
          Flag: "Potential duplicate deposit"
          Severity: WARNING
  
  COLLECT all exceptions → store in exception_list
  
  IF exception exists:
    QAEngine.FlagExceptionForReview() → adds to QA sheet

OUTPUT - QA Sheet (Exceptions):
  ┌────────────────────────────────────────┐
  │ Rec│ Field     │ Value    │ Exception  │
  ├────────────────────────────────────────┤
  │  1 │ Dep_Amt   │ $5,000.00│ (none)     │
  │  2 │ Dep_Date  │ 6/1/26   │ CRIT: Not in│
  │    │           │          │ period     │
  │  3 │ Bank_Name │ Unknown  │ CRIT: Not  │
  │    │           │          │ in list    │
  └────────────────────────────────────────┘

5. AuditLog records all transformations and exceptions
```

---

## Phase 6: QA Review & Exception Management

### Step 6.1: Review Exceptions
```
ROLE: QA Team Member

SHEET: QA sheet contains all flagged exceptions

FOR each exception:
  1. READ exception description and reason
  2. REVIEW raw extracted value vs. expected
  3. DECIDE on one of three options:
     
     OPTION A: ACCEPT
       • Exception is a false positive
       • Value is correct as extracted
       • Action: Enter "ACCEPTED" in QA sheet
     
     OPTION B: OVERRIDE
       • Exception indicates data quality issue
       • Correct value is known
       • Action: 
         - Enter corrected value in override column
         - Enter justification (e.g., "DocId 123, Page 2")
         - AuditLog captures override for trail
     
     OPTION C: REJECT
       • Value is unreliable or missing
       • No correction available
       • Action: Enter "REJECTED" with reason
       • Item will not appear in final OUTPUT sheet

EXAMPLE - QA Sheet After Review:
  ┌──────────────────────────────────────────────────┐
  │ Exc│ Field     │ Raw    │Exception  │Decision │ │
  ├──────────────────────────────────────────────────┤
  │ 1  │ Dep_Amt   │$5,000  │ OK        │ACCEPT  │ │
  │ 2  │ Dep_Date  │6/1/26  │Not in pri │OVERRIDE│✓│
  │    │ (override)│→6/5/26 │od        │        │ │
  │    │ Justif.   │Doc#A1  │          │        │ │
  │ 3  │ Bank_Name │Unknown │Not in list│REJECT  │ │
  │    │ Reason    │Cannot locate     │        │ │
  └──────────────────────────────────────────────────┘
```

### Step 6.2: Sign-Off
```
ROLE: QA Team Lead/Supervisor

AFTER all exceptions reviewed:
  1. Verify NO blank decisions (all exceptions decided)
  2. Review override justifications for completeness
  3. Enter sign-off information:
     - Name: [QA Lead Name]
     - Role: Senior QA Specialist
     - Date/Time: [Timestamp]
  4. AuditLog records sign-off event
  5. Proceed to Approval Phase
```

---

## Phase 7: Approval & Finalization

### Step 7.1: Final Approval
```
TRIGGER: User clicks "Approve & Finalize" button

PRE-APPROVAL CHECKS:
  1. ALL exceptions in QA sheet have decisions (no blanks)
  2. ALL overrides have justifications
  3. QA team sign-off present and valid
  4. If approval workflow enabled:
     • Check required approver signatures present
  
IF any check fails:
  • Error dialog shown; processing stops
  • User resolves issue and retries

IF all checks pass:
  1. OUTPUT sheet generated:
     • Accepted values → copied to OUTPUT
     • Overridden values → copied with override values
     • Rejected values → excluded
  2. OUTPUT sheet locked (read-only)
  3. AuditLog records approval event with:
     • Approver name, role, date/time
     • Number of records approved
     • Number of exceptions/overrides
  4. Framework ready for export
```

---

## Phase 8: Export & Audit Trail

### Step 8.1: Export Final Results
```
TRIGGER: User clicks "Export Results" button

PROCESS:
  1. OUTPUT sheet formatted and exported:
     • Format: CSV and XLSX options
     • Columns: Only approved final values
     • No intermediate data or exceptions
  2. Export location: Configurable per engagement
  3. Optional: Feed to downstream system
  4. AuditLog records export event

OUTPUT FILES:
  • engagement_id_final_results.csv
  • engagement_id_final_results.xlsx
```

### Step 8.2: Generate Audit Report
```
TRIGGER: User clicks "Generate Audit Report" button

PROCESS:
  AuditLog.GenerateAuditTrail() exports full trail:
  
  INCLUDES:
    • Date/time of engagement
    • All data transformations (raw → normalized)
    • All QA exceptions and resolutions
    • Override justifications
    • User actions and timestamps
    • Sign-offs and approvals
    • Exception tracking and resolution
    • Final approval information
  
  FORMAT: PDF or Excel with multiple sheets

OUTPUT:
  • engagement_id_audit_trail_report.pdf
  • Suitable for external auditor review
  • Demonstrates control effectiveness
```

---

**Framework Version**: 1.0  
**Last Updated**: June 2026
