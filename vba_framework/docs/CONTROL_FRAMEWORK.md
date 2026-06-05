# CONTROL_FRAMEWORK.md - Audit Control Standards & Implementation

## Overview

This document defines the audit control framework embedded in the VBA platform. Controls are designed to provide audit evidence, detect and prevent errors, and ensure data quality throughout the DataSnipper extraction workflow.

---

## Control Philosophy

### Three Lines of Defense

```
LINE 1: PREVENTIVE CONTROLS (Built-in Validations)
├─ Input validation (format, completeness)
├─ Tag syntax validation
├─ Range/lookup validation
└─ Duplicate detection
    ↓ Prevents invalid data from being accepted

LINE 2: DETECTIVE CONTROLS (Exception Identification)
├─ QA rules flagging exceptions
├─ Outlier detection
├─ Cross-field consistency checking
└─ Manual review requirements
    ↓ Identifies anomalies for investigation

LINE 3: CORRECTIVE CONTROLS (Resolution & Approval)
├─ Manual override capability
├─ Exception resolution workflow
├─ Multi-level approval
└─ Audit trail documentation
    ↓ Ensures corrections documented & approved
```

---

## Control Categories

### 1. INPUT CONTROLS

**Objective**: Ensure data enters the system completely and in correct format.

#### IC-1.1: Completeness Validation
```
WHAT:     All required fields populated
WHO:      Validator module
WHEN:     When "Validate Input" button clicked
HOW:      Check each required field; flag blanks
EVIDENCE: VALIDATION sheet error log
RESULT:   Block processing if critical fields missing

TEST CASE 1: All fields present
  Input: Deposit amount=$1000, date=6/1, bank=Chase
  Expected: PASS validation; proceed
  Actual: [To be verified during testing]

TEST CASE 2: Required field missing
  Input: Deposit amount=$1000, date=(blank), bank=Chase
  Expected: FLAG validation error; block processing
  Actual: [To be verified during testing]
```

#### IC-1.2: Format Validation
```
WHAT:     Data types match expected formats
WHO:      Validator module
WHEN:     During "Validate Input" process
HOW:      Check text/number/date/currency types
EVIDENCE: VALIDATION sheet format errors
RESULT:   Flag non-conforming formats for review

VALID FORMATS:
  Currency: $1,250.50 or 1250.50 (numeric)
  Date: MM/DD/YY or DD-MON-YYYY format
  Number: Numeric only; no letters
  Text: Any characters allowed
```

#### IC-1.3: Source Document Verification
```
WHAT:     Source documents exist and are accessible
WHO:      ConfigManager (during initialization)
WHEN:     When CONFIG sheet is loaded
HOW:      Check file paths; verify readability
EVIDENCE: CONFIG sheet validation report
RESULT:   Stop processing if documents unavailable

ACTION:   Audit manager verifies:
  [ ] File paths correct
  [ ] Files accessible from audit drive
  [ ] Files contain expected content
  [ ] Page counts match actual documents
```

---

### 2. PROCESSING CONTROLS

**Objective**: Ensure data is transformed correctly and consistently.

#### PC-2.1: Tag Syntax Validation
```
WHAT:     DataSnipper tags follow correct syntax
WHO:      TagBuilder module
WHEN:     During "Build Tags" process
HOW:      Validate each tag per TAG_SPEC.md
EVIDENCE: TAG_ENGINE.Tag_Status column
RESULT:   Flag syntax errors; prevent use of invalid tags

EXAMPLE:
  Valid:    DS_SEARCH:Amount:amount:(start=Total|end=Tax|type=currency)
  Invalid:  DS_SEARCH:Amount:amount                    [missing params]
  Result:   Error flagged in Tag_Status column
```

#### PC-2.2: Data Type Conversion Validation
```
WHAT:     Raw values converted to target data type correctly
WHO:      DataMapper module
WHEN:     During "Process Extraction" phase
HOW:      Attempt conversion; log success/failure
EVIDENCE: AUDIT_LOG transformation events
RESULT:   Flag conversion failures for QA review

EXAMPLES:
  "$1,250.50" → 1250.50 (SUCCESS: currency to number)
  "6/1/26" → "2026-06-01" (SUCCESS: date format conversion)
  "abc" → number (FAILURE: logged for review)
```

#### PC-2.3: Normalization Validation
```
WHAT:     Values normalized to output schema format
WHO:      DataMapper module
WHEN:     During transformation
HOW:      Apply format specifications per schema
EVIDENCE: AUDIT_LOG format application events
RESULT:   Ensure consistency across all records

EXAMPLE:
  Schema spec: $#,##0.00
  Input: 1250.5
  Output: $1,250.50 (normalized to schema)
```

---

### 3. BUSINESS LOGIC CONTROLS

**Objective**: Ensure extracted data makes business sense.

#### BLC-3.1: Range Validation (Quantitative)
```
WHAT:     Values fall within acceptable ranges
WHO:      QAEngine module
WHEN:     During "Process Extraction" phase
HOW:      Compare value to min/max thresholds
EVIDENCE: QA sheet exception flags
RESULT:   FLAG warning if outside normal range;
          BLOCK if critically out of range

EXAMPLE - CASH ENGAGEMENT:
  Rule: Deposit Amount Range
  Min: $0
  Max: $1,000,000
  
  Test Case 1: $1,250.00
    Result: PASS (within range)
  
  Test Case 2: $5,000,000.00
    Result: FLAG WARNING (unusual but possible)
  
  Test Case 3: -$100.00
    Result: BLOCK CRITICAL (negative impossible)
```

#### BLC-3.2: Lookup Validation (Reference Tables)
```
WHAT:     Values exist in approved reference tables
WHO:      QAEngine module
WHEN:     During "Process Extraction" phase
HOW:      Compare value to reference list
EVIDENCE: QA sheet exception flags
RESULT:   BLOCK if value not in approved list

EXAMPLE - CASH ENGAGEMENT:
  Rule: Bank Name Lookup
  Reference Table: APPROVED_BANKS
  Values: Chase, Bank of America, Wells Fargo, US Bank
  
  Test Case 1: "Chase"
    Result: PASS (in list)
  
  Test Case 2: "Unknown Bank"
    Result: BLOCK CRITICAL (not in list)
  
  ACTION: Audit manager investigates; confirms or approves bank
```

#### BLC-3.3: Cross-Field Consistency (Logical)
```
WHAT:     Multiple fields pass logical consistency checks
WHO:      QAEngine module
WHEN:     During "Process Extraction" phase
HOW:      Apply business logic rules across fields
EVIDENCE: QA sheet exception flags
RESULT:   FLAG if logical inconsistency detected

EXAMPLE - A/P ENGAGEMENT:
  Rule: Invoice Date Must Be After PO Date
  Logic: invoice_date > po_date
  
  Test Case 1: PO date=5/1, Invoice date=5/15
    Result: PASS (consistent)
  
  Test Case 2: PO date=5/15, Invoice date=5/1
    Result: FLAG WARNING (invoice before PO?)
  
  ACTION: QA team investigates; confirms legitimate or overrides
```

#### BLC-3.4: Duplicate Detection (Uniqueness)
```
WHAT:     No duplicate records extracted for same transaction
WHO:      QAEngine module
WHEN:     During "Process Extraction" phase
HOW:      Compare key fields across records
EVIDENCE: QA sheet duplicate flags
RESULT:   FLAG suspected duplicates for review

EXAMPLE - CASH ENGAGEMENT:
  Rule: No Duplicate Deposits (Same Amount + Date)
  Key Fields: deposit_amount, deposit_date
  
  Test Case 1: Two records with $1,250.00 on 6/1
    Result: FLAG WARNING (potential duplicate)
  
  Test Case 2: Two records both $1,250 but different dates
    Result: PASS (not duplicates)
  
  ACTION: QA team verifies if duplicate or legitimate multi-payment
```

---

### 4. PREVENTIVE CONTROLS (Exception Management)

**Objective**: Structured process for exception handling.

#### PREV-4.1: Exception Identification & Logging
```
WHAT:     All exceptions automatically captured
WHO:      QAEngine + AuditLog modules
WHEN:     During "Process Extraction" phase
HOW:      Flag triggers AuditLog record; stored in QA sheet
EVIDENCE: QA sheet + AUDIT_LOG sheet
RESULT:   Complete record of what was flagged and why

EXCEPTION RECORD INCLUDES:
  ├─ Exception ID (unique)
  ├─ Record ID (which extraction)
  ├─ Field name (which field)
  ├─ Raw value (what was extracted)
  ├─ Rule violated (which control)
  ├─ Severity (WARNING | CRITICAL)
  ├─ Timestamp (when detected)
  └─ Status (pending review)
```

#### PREV-4.2: Exception Resolution Workflow
```
STAGE 1: DETECTION
  QAEngine identifies exception
  → Exception logged to QA sheet
  → Status: PENDING_REVIEW

STAGE 2: REVIEW
  QA team member examines exception
  → Checks raw value against source document
  → Researches underlying cause
  → Documents finding

STAGE 3: DECISION
  QA team decides one of three options:
  
  OPTION A: ACCEPT
    → Exception is false positive
    → Data accepted as extracted
    → Status: ACCEPTED
  
  OPTION B: OVERRIDE
    → Data quality issue identified
    → Correct value known
    → User enters override + justification
    → Status: OVERRIDDEN
    → AuditLog records: old value, new value, reason
  
  OPTION C: REJECT
    → Data unreliable; cannot be corrected
    → Record excluded from output
    → Status: REJECTED
    → AuditLog records: reason for rejection

STAGE 4: APPROVAL
  QA lead reviews all decisions
  → Verifies decisions are supported by evidence
  → Records sign-off (name, date, title)
  → Status: QA_APPROVED

STAGE 5: FINALIZATION
  Manager approves output
  → Reviews QA decision log
  → Records final approval
  → Locks OUTPUT sheet
  → Status: FINAL_APPROVED
```

---

### 5. DETECTIVE CONTROLS (Monitoring & Review)

**Objective**: Monitor data quality throughout processing.

#### DET-5.1: Outlier Detection
```
WHAT:     Unusual values identified for investigation
WHO:      QAEngine module (via range & lookup rules)
WHEN:     During "Process Extraction" phase
HOW:      Compare to statistical norms or thresholds
EVIDENCE: QA sheet exception flags
RESULT:   FLAG unusual values; require review

EXAMPLE - CASH ENGAGEMENT:
  Normal Range: Deposits $0 - $1,000,000
  Outlier: Deposit of $5,000,000.00
  Result: FLAG "Outlier detected; review required"
```

#### DET-5.2: Manual Verification Checkpoints
```
WHAT:     Key review steps require human sign-off
WHO:      Auditor/QA team/Manager
WHEN:     At critical workflow stages
HOW:      Mandatory review before proceeding
EVIDENCE: Sign-off records in QA sheet + AUDIT_LOG
RESULT:   Ensures human oversight of critical decisions

CHECKPOINTS:
  ├─ CONFIG Review: Audit manager verifies settings
  ├─ Tag Review: Developer verifies tag quality
  ├─ QA Review: QA team reviews all exceptions
  ├─ Sign-Off: QA lead approves decisions
  └─ Final Approval: Manager approves finalization
```

#### DET-5.3: Exception Trending & Reporting
```
WHAT:     Track exception rates and patterns
WHO:      Audit Manager (periodic review)
WHEN:     After engagement completion
HOW:      Export AUDIT_LOG; analyze exception trends
EVIDENCE: Summary report of exceptions by type
RESULT:   Identify systematic issues; improve tags/rules

SAMPLE REPORT:
  Total Records Extracted:      100
  Records with Exceptions:      15 (15%)
  Exception Breakdown:
    ├─ Format errors:           3
    ├─ Range violations:        7
    ├─ Lookup failures:         4
    └─ Duplicates:              1
  
  Highest Risk Area: Range violations (7 of 15)
  Action: Review range thresholds for next engagement
```

---

### 6. CORRECTIVE CONTROLS (Audit Trail & Accountability)

**Objective**: Document all actions for audit compliance.

#### CORR-6.1: Complete Audit Trail (AUDIT_LOG)
```
WHAT:     All processing events logged with user/timestamp
WHO:      AuditLog module (all modules call it)
WHEN:     Continuously during processing
HOW:      Record event type, user, old/new values, reason
EVIDENCE: AUDIT_LOG sheet + exported audit report
RESULT:   Complete chain of custody for all data

LOGGED EVENTS:
  ├─ Data transformations (raw → normalized)
  ├─ QA exceptions (detected)
  ├─ QA decisions (ACCEPT/OVERRIDE/REJECT)
  ├─ Manual overrides (value + justification)
  ├─ User sign-offs (name + timestamp)
  ├─ System errors (with context)
  └─ Final approvals (authorization record)

RETENTION:
  ├─ Kept in workbook for engagement
  ├─ Exported to PDF/Excel for archive
  ├─ Retained per audit firm standards (typically 7 years)
  └─ Available for external auditor review
```

#### CORR-6.2: User Accountability
```
WHAT:     Each action attributed to specific user
WHO:      System (captured from Windows login)
WHEN:     Every action logged
HOW:      User ID + timestamp recorded in AUDIT_LOG
EVIDENCE: AUDIT_LOG user field
RESULT:   Clear accountability for all decisions

RECORDED:
  Who extracted data?        → Extracted_By (EXTRACTION_INPUT)
  Who reviewed exceptions?   → Reviewed_By (QA sheet)
  Who approved output?       → Approved_By (OUTPUT sheet)
  
AUDIT QUESTION: "How did we get this number?"
  ANSWER: "Susan Johnson extracted it on 6/1 at 2:30pm,
           John reviewed it on 6/2, no exceptions,
           Robert approved it on 6/3 at 4pm"
           [Full trail in AUDIT_LOG]
```

#### CORR-6.3: Exception Justification Requirements
```
WHAT:     Every override requires written justification
WHO:      QA team (when entering override)
WHEN:     During QA review phase
HOW:      Enter justification in QA sheet; captured in AUDIT_LOG
EVIDENCE: QA sheet override_justification + AUDIT_LOG
RESULT:   Support for every data correction

EXAMPLES OF ACCEPTABLE JUSTIFICATIONS:
  ✓ "Amount verified against invoice dated 5/31; page 2 of receipt"
  ✓ "Bank name corrected per client's vendor master; Doc ID: VM-2026-AB"
  ✓ "Excluded duplicate deposit; identical amount/date extracted twice on page 3"
  ✓ "Invoice date corrected to 5/28 per original PO; DataSnipper misread '8' as '6'"

EXAMPLES OF UNACCEPTABLE JUSTIFICATIONS:
  ✗ "Best guess"
  ✗ "Seemed wrong"
  ✗ "Changed by John"
  ✗ (blank)
```

---

### 7. GOVERNANCE CONTROLS (Configuration & Change Management)

**Objective**: Ensure consistent, approved processes.

#### GOV-7.1: Configuration Management
```
WHAT:     Configuration locked after QA starts
WHO:      System (prevents accidental changes)
WHEN:     Once QA review begins
HOW:      CONFIG sheet can only be modified by Manager (approval)
EVIDENCE: Change log (if implemented)
RESULT:   Prevents mid-project changes that could invalidate work

WORKFLOW:
  Phase 1: Setup → CONFIG editable (development)
  Phase 2: QA starts → CONFIG locked
  If change needed: Request approval from Engagement Manager
  Change approved → Audit all affected records
  Revalidate output
```

#### GOV-7.2: Version Control
```
WHAT:     VBA modules maintained in version control
WHO:      Development team
WHEN:     All changes to .bas files committed
HOW:      Git commits with change descriptions
EVIDENCE: Git history + VERSION.txt file
RESULT:   Track module changes; enable rollback if needed

VERSION STRUCTURE:
  Framework Version: 1.0 (major.minor)
  Release Date: June 2026
  Modules Versioned: All .bas files
  Change Log: Maintained in VERSION.txt
```

#### GOV-7.3: Testing & Validation (Before Deployment)
```
WHAT:     All modules tested with sample data
WHO:      QA team (before deployment)
WHEN:     Before framework released to audit teams
HOW:      Execute test cases per test_cases/test_cases.md
EVIDENCE: Test execution report
RESULT:   Ensure framework works as designed

TEST CATEGORIES:
  ├─ Unit tests (individual module functions)
  ├─ Integration tests (modules working together)
  ├─ Data validation tests (known-good inputs)
  ├─ Exception tests (error handling)
  └─ Performance tests (large datasets)
```

---

## Control Effectiveness Evaluation

### How We Know Controls Work

#### Method 1: Testing Results
```
For each control, execute test cases:
  Input: Known-good data
  Expected: Control works as designed
  Actual: [Test result]
  Pass/Fail: [Document]

If ALL controls PASS all tests → Control effective
If ANY control FAILS → Fix issue; retest; document resolution
```

#### Method 2: Exception Trending
```
After engagement completion, analyze:
  • What did controls catch?
  • Were any errors missed?
  • False positive rate (exceptions that were OK)?
  • Override rate (how many exceptions needed correction)?

Example Analysis:
  Rule: Duplicate Detection
  Effectiveness: 2 suspected duplicates identified
  Result: Confirmed 1 genuine duplicate (prevented error)
  False Positive Rate: 50% (1 of 2 was legitimate multi-payment)
  Conclusion: Rule EFFECTIVE (caught error; acceptable false positive rate)
```

#### Method 3: Audit Procedures
```
After engagement, audit manager verifies:
  ✓ All exceptions documented in QA sheet
  ✓ All overrides have justifications
  ✓ All sign-offs recorded with names/dates
  ✓ AUDIT_LOG complete and consistent
  ✓ OUTPUT matches approved QA decisions
  
If ALL verified → Controls operating effectively
If GAPS found → Document findings; remediate
```

---

## Risk & Control Matrix

| Risk | Control | Evidence | Frequency |
|------|---------|----------|-----------|
| **Data Extraction Errors** | Input validation; format checks | VALIDATION sheet | Every extraction |
| **Invalid Tag Syntax** | Tag syntax validation | TAG_ENGINE.Tag_Status | Build phase |
| **Out-of-Range Values** | Range validation rules | QA sheet flags | QA phase |
| **Unauthorized Vendors/Banks** | Lookup validation | QA sheet blocks | QA phase |
| **Duplicate Transactions** | Duplicate detection | QA sheet flags | QA phase |
| **Data Transformation Errors** | Data type conversion logs | AUDIT_LOG | Processing phase |
| **Unapproved Changes** | Multi-level approval workflow | Sign-off records | Approval phase |
| **Incomplete Documentation** | Audit trail logging | AUDIT_LOG | Continuous |
| **Accountability Issues** | User tracking; timestamps | AUDIT_LOG user field | Every action |
| **Configuration Changes** | Configuration lock + approval | Change log | Change control |

---

## External Audit Readiness

### Information Available for External Auditor

| Document | Location | Purpose |
|----------|----------|---------|
| **Tag Definitions** | TAG_ENGINE sheet | Shows what was extracted and how |
| **QA Exceptions** | QA sheet | Shows what was reviewed and why |
| **Audit Trail** | AUDIT_LOG sheet | Shows who did what, when, and why |
| **Configuration** | CONFIG sheet | Shows engagement parameters and rules |
| **Final Data** | OUTPUT sheet | Shows approved final results |
| **Validation Report** | VALIDATION sheet | Shows format/completeness checks |
| **Exception Analysis** | Summary report | Shows exception trending & control effectiveness |

### Evidence of Effectiveness

```
External Auditor Question: "How did you ensure data quality?"

Answer Supported By:
  1. Configuration (documented requirements)
  2. Tag Definitions (defined extraction rules)
  3. Validation Results (checked completeness/format)
  4. QA Rules (defined acceptance criteria)
  5. Exception Log (identified anomalies)
  6. QA Decisions (reviewed & approved)
  7. Audit Trail (documented everything)
  8. Sign-offs (authorized by management)
  → Demonstrates effective control over data quality
```

---

**Framework Version**: 1.0  
**Last Updated**: June 2026
