# DEVELOPMENT.md - Framework Development & Deployment

## Overview

This document describes how developers maintain the VBA framework, deploy it to engagements, and manage versions.

---

## Development Environment Setup

### Prerequisites
- **Excel** 2016 or later
- **VS Code** (for editing .bas files)
- **Git** (for version control)
- **VBA IDE** access in Excel (Tools → Macro)
- **DataSnipper** client (for tag testing)

### Folder Structure for Development
```
vba_framework/
├─ vba_modules/
│  ├─ Main.bas              (edit in VS Code; import to Excel)
│  ├─ ConfigManager.bas
│  ├─ TagBuilder.bas
│  ├─ Validator.bas
│  ├─ DataMapper.bas
│  ├─ QAEngine.bas
│  ├─ AuditLog.bas
│  └─ MODULE_SPECIFICATIONS.txt (reference)
│
├─ templates/
│  ├─ Master_Template.xlsx          (base template for all engagements)
│  ├─ CashTemplate.xlsx             (cash cycle specialization)
│  ├─ ARTemplate.xlsx               (A/R specialization)
│  └─ [etc]
│
└─ test_cases/
   ├─ test_cases.md                 (test scenarios)
   └─ sample_data/                  (test data files)
```

---

## Module Development Workflow

### Step 1: Create Module as Text File

```
EXAMPLE: Creating TagBuilder.bas

1. Open VS Code
2. Create file: vba_modules/TagBuilder.bas
3. Start with file header:
   
   'TagBuilder Module
   'Purpose: Generate DataSnipper tags
   'Version: 1.0
   'Last Updated: [Date]
   'Developer: [Name]
   
4. Implement functions per MODULE_SPECIFICATIONS.txt
5. Add inline comments explaining logic
6. Save file
7. Commit to Git: git add TagBuilder.bas; git commit -m "TagBuilder module implementation"
```

### Step 2: Import Module into Excel (Testing)

```
PROCESS:

1. Open Excel (Engagement workbook or test workbook)
2. Press Alt+F11 (Open VBA IDE)
3. Right-click in Project Explorer → Import File
4. Select vba_modules/TagBuilder.bas
5. Module imported into Excel
6. Test module functions with sample data
7. If issues: Fix in VS Code; re-import
8. If working: Keep in Excel; commit to Git
```

### Step 3: Unit Testing

```
FOR EACH MODULE:

1. Create test cases per test_cases/test_cases.md
2. Execute test case against module
3. Verify output matches expected result

Example Test Case:
  Module: TagBuilder
  Function: BuildSearchTag()
  Input: TagDef with keywords ["Total", "Amount"]
  Expected Output: "DS_SEARCH:Amount:amount:(start=Total|end=..."
  Actual Output: [Run and record]
  Result: PASS / FAIL
  
4. Document test results
5. If FAIL: Fix module; re-test
6. If PASS: Module ready for integration testing
```

### Step 4: Integration Testing

```
TEST: Modules working together

Example:
  1. ConfigManager loads CONFIG sheet
  2. TagBuilder reads ConfigManager.GetTagDefinitions()
  3. TagBuilder outputs to TAG_ENGINE sheet
  4. Validator reads TAG_ENGINE output
  
Test Flow:
  1. Populate CONFIG sheet with test engagement
  2. Run ConfigManager.LoadConfig() - verify loads correctly
  3. Run TagBuilder.BuildTags() - verify tags generated
  4. Run Validator on output - verify validation works
  5. If all pass: Integration working
  6. If fail: Debug module interaction; fix; retest
```

### Step 5: Commit to Version Control

```
AFTER EACH MODULE COMPLETE:

git add vba_modules/[ModuleName].bas
git commit -m "Implement [ModuleName]: [Brief description]"

Example:
  git commit -m "Implement TagBuilder: DS_SEARCH and DS_COORDS tag generation"
  git commit -m "Implement Validator: Format validation and completeness checks"
  git commit -m "Implement QAEngine: Range, lookup, and business logic rules"

Commit messages should be:
  ├─ Descriptive (what was added/fixed)
  ├─ Include module name
  └─ Reference any bug fixes or enhancements
```

---

## Deployment Process

### Scenario: Deploying Framework to New Engagement Workbook

```
STEP 1: Create Engagement Workbook
  1. Copy templates/Master_Template.xlsx
  2. Rename: [EngagementID].xlsx
  3. Open in Excel

STEP 2: Import VBA Modules
  1. Open VBA IDE (Alt+F11)
  2. Right-click Project Explorer → Import File
  3. Import each .bas file in order:
     - Main.bas
     - ConfigManager.bas
     - TagBuilder.bas
     - Validator.bas
     - DataMapper.bas
     - QAEngine.bas
     - AuditLog.bas
  4. Verify imports successful (no errors)

STEP 3: Verify Ribbon/Buttons
  1. Close VBA IDE
  2. Verify buttons appear in Excel ribbon:
     - Load Engagement
     - Build Tags
     - Validate Input
     - Process Extraction
     - [etc]
  3. If buttons missing: Create macro buttons in ribbon

STEP 4: Populate Configuration
  1. Go to CONFIG sheet
  2. Enter engagement details (Engagement ID, dates, documents, etc)
  3. Populate TAG_ENGINE with extraction rules
  4. Save workbook

STEP 5: Test with Sample Data
  1. Run "Load Engagement" → verify CONFIG loads
  2. Run "Build Tags" → verify tags generated
  3. Run "Validate Input" with sample extraction data
  4. Verify workflow works end-to-end

STEP 6: Deploy to Audit Team
  1. Workbook ready for use by audit team
  2. Document any customizations in engagement notes
  3. Provide CONFIG_GUIDE.md to audit team
```

---

## Version Management

### Version Numbering Scheme

```
Framework Version: MAJOR.MINOR

MAJOR: Significant architecture changes or new modules
  1.0 → 2.0: Major refactor or new feature set
  
MINOR: Bug fixes, enhancements, new rules/templates
  1.0 → 1.1: New QA rules added
  1.1 → 1.2: Bug fix in DataMapper

Example: Framework Version 1.2
  Indicates: Version 1 (current generation), Release 2 (second update)
```

### VERSION.txt File

```
FRAMEWORK VERSION HISTORY

Framework Version: 1.0
Release Date: June 2026
Status: Release (Live)

Modules Included:
  ✓ Main.bas              v1.0
  ✓ ConfigManager.bas     v1.0
  ✓ TagBuilder.bas        v1.0
  ✓ Validator.bas         v1.0
  ✓ DataMapper.bas        v1.0
  ✓ QAEngine.bas          v1.0
  ✓ AuditLog.bas          v1.0

Changes in 1.0:
  • Initial framework release
  • All core modules implemented
  • Standard QA rules for Cash, AR, AP engagements
  • Audit logging and control framework
  • Multi-level approval workflow

Known Issues: None

Future Enhancements (1.1 pipeline):
  • Enhanced duplicate detection logic
  • Additional QA rules for Contracts engagement
  • Performance optimization for large extractions
```

### Change Log (in Git)

```
Each commit documents a change:

Commit 1: Initial framework architecture
Commit 2: Implement ConfigManager module
Commit 3: Implement TagBuilder with DS_SEARCH support
Commit 4: Implement TagBuilder with DS_COORDS support
Commit 5: Implement Validator module
Commit 6: Implement DataMapper data type conversions
Commit 7: Implement QAEngine with range validation
Commit 8: Implement QAEngine with lookup validation
Commit 9: Implement AuditLog and event tracking
Commit 10: Add unit tests for all modules
Commit 11: Framework v1.0 ready for deployment
```

---

## Testing Strategy

### Unit Tests (Individual Modules)

```
TEMPLATE FOR EACH MODULE:

Module: ConfigManager
Test Suite: config_manager_tests.vba

Test 1: LoadConfig - Valid CONFIG
  Input: Well-formed CONFIG sheet
  Expected: ConfigObject returned
  Actual: [test result]
  Status: PASS / FAIL

Test 2: LoadConfig - Missing Required Field
  Input: CONFIG missing Engagement ID
  Expected: Error code 101 (missing field)
  Actual: [test result]
  Status: PASS / FAIL

Test 3: ValidateConfig - Valid Config
  Input: Valid ConfigObject
  Expected: No validation errors
  Actual: [test result]
  Status: PASS / FAIL

... [more tests]
```

### Integration Tests (Modules Working Together)

```
TEST: End-to-End Workflow

1. Setup Test Environment
   - Create test workbook with sample CONFIG
   - Add sample extraction data

2. Execute Workflow
   a. LoadEngagement() → verify CONFIG loaded
   b. BuildTags() → verify tags generated
   c. ValidateInput() → verify validation passes
   d. ProcessExtraction() → verify no exceptions
   e. ApproveOutput() → verify output generated

3. Verify Results
   - Check OUTPUT sheet has expected data
   - Check AUDIT_LOG has all events
   - Check workflow state progressed correctly

4. Document Results
   - Test passed: Framework ready
   - Test failed: Debug module interaction
```

### Regression Tests (After Updates)

```
AFTER EACH FRAMEWORK UPDATE:

1. Run all unit tests → verify no modules broken
2. Run integration test → verify end-to-end still works
3. Run on previous engagement data → verify backward compatibility
4. Check performance (if optimization made)
5. Document any regressions
```

---

## Debugging Guide

### Common Issues & Solutions

| Issue | Cause | Solution |
|-------|-------|----------|
| **Module import fails** | Syntax error in .bas file | Check VBA IDE error message; fix syntax |
| **Function returns unexpected value** | Logic error | Add Debug.Print statements; trace execution |
| **AUDIT_LOG not recording** | AuditLog module not called | Verify all modules call AuditLog.Record*() |
| **QA rules not applying** | Rules not loaded from CONFIG | Check ConfigManager.GetQARules() returns data |
| **OUTPUT sheet empty** | No records passed QA | Check QA sheet decisions; verify ACCEPT/OVERRIDE exist |

### Debug Print Statements

```
RECOMMENDED APPROACH:

Add at key points in code:

Public Sub ProcessExtraction()
  Debug.Print "Starting ProcessExtraction"
  Debug.Print "Records to process: " & UBound(rawData)
  
  FOR i = 1 TO UBound(rawData)
    result = DataMapper.MapRawToSchema(rawData(i))
    Debug.Print "Record " & i & " mapped successfully"
  NEXT i
  
  Debug.Print "ProcessExtraction completed"
End Sub

View output: In VBA IDE, View → Immediate Window (Ctrl+G)
```

### Error Handling in Code

```
RECOMMENDED PATTERN:

Public Function MyFunction() As Boolean
  On Error GoTo ErrorHandler
  
  ' Main logic here
  
  MyFunction = True
  Exit Function
  
ErrorHandler:
  Debug.Print "ERROR in MyFunction: " & Err.Description
  Call AuditLog.RecordError("MyFunction", Err.Number, Err.Description)
  MyFunction = False
End Function
```

---

## Performance Optimization

### Considerations for Large Extractions

```
SCENARIO: 1000+ extracted records

PERFORMANCE RISKS:
  ├─ Excel calculation recalculation overhead
  ├─ VBA array processing time
  ├─ AUDIT_LOG growing very large
  └─ QAEngine checking rules for each record

OPTIMIZATION STRATEGIES:
  1. Disable Excel calculations during processing
     Application.Calculation = xlCalculationManual
     [... processing ...]
     Application.Calculation = xlCalculationAutomatic
  
  2. Use arrays instead of cell-by-cell operations
     ├─ Read data into array once
     ├─ Process in memory
     ├─ Write results to sheet once
  
  3. Batch AUDIT_LOG writes
     ├─ Collect events in memory
     ├─ Write to sheet in bulk
     └─ Don't write every single event to sheet
  
  4. Consider data paging
     ├─ Process in chunks (e.g., 100 records at a time)
     ├─ Show progress bar to user
     ├─ Allow user to stop if needed
```

---

## Documentation Standards

### Code Comments

```
EVERY FUNCTION MUST INCLUDE:

Public Function BuildSearchTag(tagDef As TagDefinition) As String
  ' BuildSearchTag: Generate DS_SEARCH tag from definition
  ' Purpose: Create properly formatted DataSnipper search tag
  ' Inputs:  tagDef - TagDefinition with search parameters
  ' Outputs: String containing DS_SEARCH tag
  ' Calls:   ValidateTagSyntax
  ' Errors:  Returns empty string if validation fails
  
  [Function body]
  
End Function
```

### Module Header

```
'==============================================================
' Module: TagBuilder
' Purpose: Generate DataSnipper tag syntax
' Version: 1.0
' Author: [Developer Name]
' Created: [Date]
' Last Modified: [Date]
'
' Dependencies:
'   - ConfigManager module (for tag definitions)
'   - AuditLog module (for event logging)
'   - TAG_SPEC.md (for syntax standards)
'
' Public Functions:
'   - BuildTags()
'   - BuildSearchTag()
'   - BuildCoordTag()
'   - ValidateTagSyntax()
'   - ExportTagsToSheet()
'
' Revision History:
'   1.0 (June 2026) - Initial release
'==============================================================
```

---

## Deployment Checklist

Before releasing framework to audit teams:

```
[ ] All modules implemented per MODULE_SPECIFICATIONS
[ ] All unit tests passing
[ ] Integration tests passing
[ ] Regression tests passing
[ ] Performance acceptable (< 5 seconds per 100 records)
[ ] Code reviewed by senior developer
[ ] Documentation complete
  [ ] README.md
  [ ] ARCHITECTURE.md
  [ ] WORKFLOW.md
  [ ] TAG_SPEC.md
  [ ] QA_RULES.md
  [ ] CONFIG_GUIDE.md
  [ ] DATA_MODEL.md
  [ ] CONTROL_FRAMEWORK.md
[ ] Templates created for each engagement type
  [ ] Master_Template.xlsx
  [ ] CashTemplate.xlsx
  [ ] ARTemplate.xlsx
  [ ] [etc]
[ ] Sample test data included in test_cases/
[ ] VERSION.txt updated
[ ] Git committed and tagged (v1.0)
[ ] Training materials prepared (if applicable)
[ ] Support contact information provided
```

---

**Framework Version**: 1.0  
**Last Updated**: June 2026
