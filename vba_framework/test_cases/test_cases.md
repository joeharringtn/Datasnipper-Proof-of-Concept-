# TEST_CASES.md - Framework Testing Scenarios

## Overview

This document defines test cases for validating the VBA framework. Tests are organized by module and include both unit tests (individual functions) and integration tests (end-to-end workflows).

---

## Test Environment Setup

### Requirements
- Test workbook with sample data
- CONFIG sheet populated with test engagement
- Sample EXTRACTION_INPUT data (valid and invalid cases)
- Reference tables (APPROVED_BANKS, etc) for lookup tests

### Test Data File Location
```
test_cases/
├─ sample_data/
│  ├─ cash_test_data.csv           (sample cash extractions)
│  ├─ ar_test_data.csv             (sample A/R extractions)
│  ├─ ap_test_data.csv             (sample A/P extractions)
│  └─ reference_tables.xlsx        (lookup tables for validation)
│
└─ expected_results/
   ├─ cash_expected_output.csv     (expected OUTPUT sheet)
   ├─ ar_expected_output.csv
   └─ ap_expected_output.csv
```

---

## Module 1: ConfigManager Tests

### TC-1.1: Load Valid Configuration
```
Purpose: Verify ConfigManager correctly loads well-formed CONFIG sheet
Setup:
  - Create CONFIG sheet with all required fields populated
  - Engagement ID: 2026-CASH-TEST-01
  - Engagement Type: Cash
  - Period: 01/01/2026 - 06/30/2026
  - Documents: 2 documents with valid file paths
  - Approvers: Valid email addresses

Execute:
  Call ConfigManager.LoadConfig()

Expected Result:
  - ConfigObject returned (not empty)
  - All fields populated correctly
  - No validation errors

Acceptance Criteria:
  ✓ ConfigObject contains engagement metadata
  ✓ DocumentSpecs array has 2 entries
  ✓ TagDefinitions loaded (if any)
  ✓ QARules loaded (if any)

Status: [To be completed during testing]
```

### TC-1.2: Load Missing Required Field
```
Purpose: Verify ConfigManager detects missing required fields
Setup:
  - Create CONFIG sheet with Engagement ID field BLANK
  - All other fields valid

Execute:
  Call ConfigManager.LoadConfig()
  result = ConfigManager.ValidateConfig()

Expected Result:
  - ValidateConfig returns error array
  - Error includes: "Engagement_ID required"
  - ProcessingBlocked = TRUE

Acceptance Criteria:
  ✓ Error code 101 (missing required field)
  ✓ Error message clear and actionable
  ✓ Framework prevents proceeding

Status: [To be completed during testing]
```

### TC-1.3: Validate Date Range Logic
```
Purpose: Verify period start < period end validation
Setup:
  - CONFIG with period start: 06/30/2026
  - CONFIG with period end: 01/01/2026 (reversed)

Execute:
  result = ConfigManager.ValidateConfig()

Expected Result:
  - Validation fails
  - Error: "Period_End must be after Period_Start"
  - Processing blocked

Status: [To be completed during testing]
```

---

## Module 2: TagBuilder Tests

### TC-2.1: Build DS_SEARCH Tag
```
Purpose: Verify TagBuilder generates valid DS_SEARCH tag
Setup:
  - TagDefinition with:
    ├─ extraction_method: DS_SEARCH
    ├─ field_name: deposit_amount
    ├─ start_anchor: "Total Amount:"
    ├─ end_anchor: "Tax"
    └─ field_type: currency

Execute:
  tag = TagBuilder.BuildSearchTag(tagDef)

Expected Result:
  tag = "DS_SEARCH:DepositAmount:deposit_amount:(start=Total Amount:|end=Tax|type=currency)"

Acceptance Criteria:
  ✓ Tag starts with "DS_SEARCH:"
  ✓ Field ID and output field included
  ✓ Parameters formatted with pipes (|)
  ✓ All expected parameters present

Status: [To be completed during testing]
```

### TC-2.2: Build DS_COORDS Tag
```
Purpose: Verify TagBuilder generates valid DS_COORDS tag
Setup:
  - TagDefinition with:
    ├─ extraction_method: DS_COORDS
    ├─ field_name: deposit_date
    ├─ page: 1
    ├─ x: 300, y: 450
    ├─ width: 100, height: 20
    └─ field_type: date

Execute:
  tag = TagBuilder.BuildCoordTag(tagDef)

Expected Result:
  tag = "DS_COORDS:DepositDate:deposit_date:(page=1|x=300|y=450|width=100|height=20|type=date)"

Acceptance Criteria:
  ✓ Tag starts with "DS_COORDS:"
  ✓ All coordinate parameters included
  ✓ Numeric values correct
  ✓ Type parameter present

Status: [To be completed during testing]
```

### TC-2.3: Tag Syntax Validation - Valid
```
Purpose: Verify valid tag passes syntax validation
Setup:
  - Valid tag: "DS_SEARCH:Amount:amount:(start=Total|end=Tax|type=currency)"

Execute:
  result = TagBuilder.ValidateTagSyntax(tag)

Expected Result:
  - result.IsValid = TRUE
  - result.ErrorMessage = ""

Status: [To be completed during testing]
```

### TC-2.4: Tag Syntax Validation - Missing Parameters
```
Purpose: Verify invalid tag (missing parameters) fails validation
Setup:
  - Invalid tag: "DS_SEARCH:Amount:amount"

Execute:
  result = TagBuilder.ValidateTagSyntax(tag)

Expected Result:
  - result.IsValid = FALSE
  - result.ErrorMessage contains: "missing parameters"

Status: [To be completed during testing]
```

---

## Module 3: Validator Tests

### TC-3.1: Complete Required Field Validation - Pass
```
Purpose: Verify validation passes when all required fields present
Setup:
  - EXTRACTION_INPUT with 5 records
  - Record 1: deposit_amount="$1,250.00", deposit_date="6/1/26", bank_name="Chase"
  - All other required fields populated

Execute:
  results = Validator.ValidateInputComplete(extractionData)

Expected Result:
  - No validation errors for Record 1
  - results array is empty (no exceptions)

Status: [To be completed during testing]
```

### TC-3.2: Complete Required Field Validation - Fail (Blank Field)
```
Purpose: Verify validation fails when required field blank
Setup:
  - EXTRACTION_INPUT Record 2:
    ├─ deposit_amount="$500.00"
    ├─ deposit_date=(blank)
    └─ bank_name="Chase"

Execute:
  results = Validator.ValidateInputComplete(extractionData)

Expected Result:
  - Validation error flagged for Record 2
  - Error message: "deposit_date is required"
  - Severity: CRITICAL

Status: [To be completed during testing]
```

### TC-3.3: Data Format Validation - Currency
```
Purpose: Verify currency format validation
Setup:
  - EXTRACTION_INPUT with deposit_amount values:
    ├─ Row 1: "$1,250.50" (valid)
    ├─ Row 2: "1250.50" (valid - numeric)
    ├─ Row 3: "abc" (invalid - non-numeric)

Execute:
  schema.deposit_amount.DataType = "currency"
  results = Validator.ValidateDataFormat(extractionData, schema)

Expected Result:
  - Row 1: VALID
  - Row 2: VALID
  - Row 3: INVALID - "Cannot convert to currency"

Status: [To be completed during testing]
```

### TC-3.4: Data Format Validation - Date
```
Purpose: Verify date format handling of multiple formats
Setup:
  - EXTRACTION_INPUT with deposit_date values:
    ├─ Row 1: "6/1/26" (MM/DD/YY format)
    ├─ Row 2: "01-JUN-2026" (DD-MON-YYYY format)
    ├─ Row 3: "2026-06-01" (YYYY-MM-DD format)

Execute:
  schema.deposit_date.DataType = "date"
  results = Validator.ValidateDataFormat(extractionData, schema)

Expected Result:
  - All 3 rows: VALID (all date formats accepted)

Status: [To be completed during testing]
```

---

## Module 4: DataMapper Tests

### TC-4.1: Data Type Conversion - Currency to Number
```
Purpose: Verify currency text converted to numeric value
Setup:
  - Raw value: "$1,250.50" (extracted from document)
  - Target type: currency
  - Expected output type: Double

Execute:
  converted = DataMapper.ConvertDataType("$1,250.50", "currency")

Expected Result:
  - converted = 1250.50 (numeric)
  - AuditLog records transformation

Acceptance Criteria:
  ✓ $ symbol removed
  ✓ Comma removed
  ✓ Value numeric (Double)

Status: [To be completed during testing]
```

### TC-4.2: Data Type Conversion - Date Format Normalization
```
Purpose: Verify date converted to YYYY-MM-DD format
Setup:
  - Raw value: "6/1/26" (extracted from document)
  - Target type: date
  - Expected format: YYYY-MM-DD

Execute:
  converted = DataMapper.ConvertDataType("6/1/26", "date")

Expected Result:
  - converted = "2026-06-01"
  - Format: YYYY-MM-DD

Status: [To be completed during testing]
```

### TC-4.3: Format Normalization - Currency Display
```
Purpose: Verify currency formatted per schema spec
Setup:
  - Converted value: 1250.50 (numeric)
  - Format spec: "$#,##0.00"

Execute:
  formatted = DataMapper.NormalizeFormat(1250.50, "$#,##0.00")

Expected Result:
  - formatted = "$1,250.50"

Status: [To be completed during testing]
```

---

## Module 5: QAEngine Tests

### TC-5.1: Range Validation - Within Range
```
Purpose: Verify value within acceptable range passes
Setup:
  - QARule: RANGE_CASH_01
  - Min: $0, Max: $1,000,000
  - Value: $500,000.00

Execute:
  result = QAEngine.ApplyQARules(records, [RANGE_CASH_01])

Expected Result:
  - No exception flagged
  - Record passes QA

Status: [To be completed during testing]
```

### TC-5.2: Range Validation - Above Maximum
```
Purpose: Verify value above max range flagged as exception
Setup:
  - QARule: RANGE_CASH_01
  - Min: $0, Max: $1,000,000
  - Value: $5,000,000.00

Execute:
  results = QAEngine.ApplyQARules(records, [RANGE_CASH_01])

Expected Result:
  - Exception flagged in QA sheet
  - Severity: WARNING
  - Message: "Value exceeds maximum range"

Status: [To be completed during testing]
```

### TC-5.3: Lookup Validation - Value in List
```
Purpose: Verify value in approved list passes
Setup:
  - QARule: LOOKUP_CASH_01
  - Reference table: APPROVED_BANKS = ["Chase", "BOA", "WF", "US Bank"]
  - Value: "Chase"

Execute:
  result = QAEngine.ApplyQARules(records, [LOOKUP_CASH_01])

Expected Result:
  - No exception
  - Record passes QA

Status: [To be completed during testing]
```

### TC-5.4: Lookup Validation - Value NOT in List
```
Purpose: Verify value not in list flagged as exception
Setup:
  - QARule: LOOKUP_CASH_01
  - Reference table: APPROVED_BANKS
  - Value: "Unknown Bank"

Execute:
  results = QAEngine.ApplyQARules(records, [LOOKUP_CASH_01])

Expected Result:
  - Exception flagged in QA sheet
  - Severity: CRITICAL
  - Message: "Bank name not in approved list"
  - Fail action: BLOCK

Status: [To be completed during testing]
```

### TC-5.5: Cross-Field Validation - Deposit Date in Period
```
Purpose: Verify cross-field logic (deposit date within period)
Setup:
  - QARule: CROSS_CASH_01 (deposit_date between period_start and period_end)
  - Period: 01/01/2026 - 06/30/2026
  - Value: 06/15/2026

Execute:
  result = QAEngine.ApplyQARules(records, [CROSS_CASH_01])

Expected Result:
  - No exception
  - Record passes QA

Status: [To be completed during testing]
```

### TC-5.6: Cross-Field Validation - Deposit Date After Period
```
Purpose: Verify cross-field validation catches date outside period
Setup:
  - QARule: CROSS_CASH_01
  - Period: 01/01/2026 - 06/30/2026
  - Value: 07/15/2026 (after period end)

Execute:
  results = QAEngine.ApplyQARules(records, [CROSS_CASH_01])

Expected Result:
  - Exception flagged
  - Severity: CRITICAL
  - Message: "Deposit date outside audit period"

Status: [To be completed during testing]
```

### TC-5.7: Duplicate Detection
```
Purpose: Verify duplicate deposits flagged
Setup:
  - Record 1: deposit_amount=$1,250.00, deposit_date=6/1/26
  - Record 2: deposit_amount=$1,250.00, deposit_date=6/1/26
  - QARule: DUP_CASH_01 (flag same amount + date combo)

Execute:
  results = QAEngine.ApplyQARules(records, [DUP_CASH_01])

Expected Result:
  - Exception flagged on Record 2
  - Severity: WARNING
  - Message: "Potential duplicate deposit"

Status: [To be completed during testing]
```

---

## Integration Tests

### IT-1: End-to-End Cash Workflow
```
Purpose: Verify complete workflow from config to output

WORKFLOW:
1. LoadEngagement()
   ├─ Load CONFIG sheet (Cash engagement)
   ├─ Verify ConfigManager returns valid ConfigObject
   └─ Status: CONFIG_LOADED

2. BuildTags()
   ├─ Generate tags from TAG_ENGINE
   ├─ Write to TAG_ENGINE.SourceTag
   └─ Status: TAGS_READY

3. [External: User runs DataSnipper]
   └─ Extracts data; pastes into EXTRACTION_INPUT

4. ValidateInput()
   ├─ Validate completeness and format
   ├─ Write results to VALIDATION sheet
   └─ Status: VALIDATION_COMPLETE

5. ProcessExtraction()
   ├─ Transform to schema
   ├─ Apply QA rules
   ├─ Flag exceptions in QA sheet
   └─ Status: QA_IN_PROGRESS

6. [User reviews QA sheet]
   ├─ Reviews exceptions
   ├─ Enters decisions (ACCEPT/OVERRIDE/REJECT)
   ├─ QA lead signs off
   └─ Status: QA_REVIEWED

7. ApproveOutput()
   ├─ Generate OUTPUT sheet
   ├─ Record final approval
   ├─ Lock OUTPUT sheet
   └─ Status: COMPLETE

8. ExportResults()
   ├─ Export OUTPUT to CSV/XLSX
   └─ [Ready for downstream use]

Expected Result:
  ✓ All 8 steps complete without errors
  ✓ OUTPUT sheet contains expected number of records
  ✓ AUDIT_LOG contains complete trail
  ✓ Sign-offs recorded

Status: [To be completed during testing]
```

---

## Regression Tests (After Updates)

### RT-1: After Updating QAEngine
```
Purpose: Verify QAEngine update doesn't break existing rules

Execute:
  1. Run all QA rules from previous version
  2. Verify they produce same results as before
  3. Compare new rules (if added)

Expected Result:
  ✓ All existing rules work as before
  ✓ New rules work correctly
  ✓ No performance degradation
```

---

## Performance Tests

### PT-1: Process Large Dataset (1000+ Records)
```
Purpose: Verify framework handles large extractions

Setup:
  - EXTRACTION_INPUT with 1000 records
  - CONFIG with 10 QA rules

Execute:
  1. LoadEngagement()
  2. ProcessExtraction() (validate + QA rules)
  3. Measure execution time

Expected Result:
  - Execution time < 30 seconds
  - No crashes or freezes
  - All records processed
  - Memory usage reasonable
```

---

## Test Results Template

```
TEST EXECUTION REPORT

Date: [Date]
Tester: [Name]
Framework Version: 1.0
Test Environment: [Excel version, OS]

Test Case │ Expected │ Actual │ Status │ Notes
──────────┼──────────┼────────┼────────┼──────
TC-1.1    │ Pass     │ [result]│ [P/F] │ [notes]
TC-1.2    │ Pass     │ [result]│ [P/F] │ [notes]
[etc]

Summary:
  Total Tests: [N]
  Passed: [N]
  Failed: [N]
  Success Rate: [%]
  
Issues Found:
  [List any failures]
  
Recommendations:
  [Suggest fixes or enhancements]
```

---

**Framework Version**: 1.0  
**Last Updated**: June 2026
