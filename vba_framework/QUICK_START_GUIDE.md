# QUICK_START_GUIDE.md - Framework Quick Reference

## For Audit Managers (Engagement Setup)

### 5-Minute Setup

```
STEP 1: Create Workbook (5 min)
  1. Copy templates/Master_Template.xlsx
  2. Rename: [EngagementID]_[EngagementType].xlsx
  3. Open in Excel

STEP 2: Fill CONFIG Sheet (10 min)
  Section A: Engagement info (ID, type, period, auditor)
  Section B: Document info (file paths, page counts)
  Section C: Select QA rules for your engagement type
  
STEP 3: Populate TAG_ENGINE Sheet (15 min)
  For each data field needed:
    - Enter field name
    - Choose extraction method (DS_SEARCH or DS_COORDS)
    - Enter search keywords OR coordinate positions
    - Mark as required (Yes/No)

STEP 4: Generate Tags (2 min)
  Click "Build Tags" button
  Review generated tags in SourceTag column
  Adjust if needed; re-run if changes made

STEP 5: Ready to Extract (1 min)
  Copy tags into DataSnipper
  Send workbook to audit team
  
TOTAL TIME: ~30 minutes

→ See CONFIG_GUIDE.md for detailed instructions
```

---

## For Auditors (Data Extraction & QA)

### End-to-End Workflow (1-2 hours per engagement)

```
PHASE 1: EXTRACTION (30 min)
  1. Receive configured workbook from audit manager
  2. Copy tags from TAG_ENGINE.SourceTag column
  3. Open DataSnipper client
  4. Paste tags; load source documents
  5. Run extraction
  6. Copy results; paste into EXTRACTION_INPUT sheet

PHASE 2: VALIDATION (10 min)
  1. Click "Validate Input" button
  2. Review VALIDATION sheet for errors
  3. If errors: Fix in EXTRACTION_INPUT; re-validate
  4. If clean: Proceed to next phase

PHASE 3: PROCESSING & QA (20 min)
  1. Click "Process Extraction" button
  2. Framework applies transformations and QA rules
  3. Exceptions written to QA sheet
  4. Flag count shown; review required exceptions

PHASE 4: QA REVIEW (45 min - varies by exception count)
  1. Open QA sheet
  2. For each exception:
     - Review extracted value vs. source document
     - Decide: ACCEPT | OVERRIDE | REJECT
     - If overriding: Enter correct value + justification
  3. QA Lead: Review all decisions; sign off

PHASE 5: APPROVAL (5 min)
  1. Click "Approve & Finalize" button
  2. Manager approves (if approval workflow enabled)
  3. OUTPUT sheet generated and locked
  4. Ready for export

PHASE 6: EXPORT (5 min)
  1. Click "Export Results" button
  2. OUTPUT sheet exported as CSV/XLSX
  3. Ready for downstream analysis

TOTAL TIME: 1-2 hours (depending on exception count)

→ See WORKFLOW.md for detailed process
```

---

## For Developers (Implementation)

### Module Development Cycle (per module)

```
STEP 1: Design (Review specifications)
  - Read MODULE_SPECIFICATIONS.txt
  - Understand inputs, outputs, interactions
  - Plan implementation approach

STEP 2: Code (Write .bas file)
  - Create file in vba_modules/ folder
  - Implement functions per specification
  - Add comments and error handling
  - Save and commit to Git

STEP 3: Unit Test (Test individual functions)
  - Follow test_cases/test_cases.md
  - Execute each function with sample data
  - Verify outputs match expected results
  - Document test results

STEP 4: Integration Test (Test with other modules)
  - Import module into Excel workbook
  - Test interaction with ConfigManager and other modules
  - Verify data flows correctly between modules
  - Fix issues; re-test

STEP 5: Deploy (Release to engagements)
  - Module complete and tested
  - Ready to be imported into engagement workbooks
  - Deployed alongside other modules

→ See DEVELOPMENT.md for detailed developer guide
```

---

## Document Navigation Map

### By Role

**AUDIT MANAGERS:**
- Start: README.md (overview)
- Then: CONFIG_GUIDE.md (how to configure)
- Reference: WORKFLOW.md (process overview)

**AUDITORS:**
- Start: README.md (overview)
- Then: WORKFLOW.md (step-by-step process)
- Reference: CONFIG_GUIDE.md (if modifying config)

**DEVELOPERS:**
- Start: ARCHITECTURE.md (module design)
- Then: MODULE_SPECIFICATIONS.txt (coding specs)
- Reference: DEVELOPMENT.md (dev workflow)
- Test: test_cases/test_cases.md (validation)

**QA SPECIALISTS:**
- Start: QA_RULES.md (validation rules)
- Reference: DATA_MODEL.md (QA sheet structure)
- Reference: CONTROL_FRAMEWORK.md (control logic)

**EXTERNAL AUDITORS:**
- Start: README.md (overview)
- Then: CONTROL_FRAMEWORK.md (controls overview)
- Reference: QA_RULES.md (testing performed)

### By Topic

**Setup & Configuration:**
- CONFIG_GUIDE.md - How to configure for new engagement
- templates/ - Pre-built templates by engagement type

**Tag Design:**
- TAG_SPEC.md - Tag syntax and standards
- WORKFLOW.md - When to build tags

**Data Quality:**
- QA_RULES.md - Validation rules by engagement type
- CONTROL_FRAMEWORK.md - Audit control framework

**Data Structures:**
- DATA_MODEL.md - Excel sheet schemas
- tag_mappings.json - Tag-to-field mappings

**Development:**
- ARCHITECTURE.md - Module design and interactions
- MODULE_SPECIFICATIONS.txt - Detailed module specs
- DEVELOPMENT.md - Dev workflow and deployment
- test_cases/test_cases.md - Testing procedures

---

## Common Questions & Answers

### "How do I set up an engagement?"
→ Follow CONFIG_GUIDE.md Section A-C (30 min setup)

### "What tags should I use?"
→ See TAG_SPEC.md for tag standards; choose DS_SEARCH vs DS_COORDS based on document consistency

### "My extraction has exceptions in QA. What do I do?"
→ Review QA_RULES.md to understand the rule; investigate the data; decide ACCEPT/OVERRIDE/REJECT per WORKFLOW.md Phase 6

### "How do I know the controls are working?"
→ See CONTROL_FRAMEWORK.md; audit trail is complete in AUDIT_LOG sheet

### "Can I customize QA rules for my engagement?"
→ Yes! See QA_RULES.md for rule types; CONFIG_GUIDE.md Section D for how to add custom rules

### "What if my data extraction has 2000+ records?"
→ Framework handles large volumes; see DEVELOPMENT.md section on Performance Optimization

### "How do I export my final results?"
→ Click "Export Results" in Main module; see WORKFLOW.md Phase 8

---

## Framework File Structure (Quick Reference)

```
vba_framework/
├─ README.md                          # START HERE
├─ docs/
│  ├─ ARCHITECTURE.md                 # Module design
│  ├─ WORKFLOW.md                     # Process flows
│  ├─ TAG_SPEC.md                     # Tag syntax
│  ├─ QA_RULES.md                     # Validation rules
│  ├─ CONFIG_GUIDE.md                 # How to configure
│  ├─ DATA_MODEL.md                   # Excel schemas
│  ├─ CONTROL_FRAMEWORK.md            # Audit controls
│  ├─ DEVELOPMENT.md                  # Dev workflow
│  └─ QUICK_START_GUIDE.md            # This file!
│
├─ vba_modules/
│  ├─ MODULE_SPECIFICATIONS.txt       # Detailed specs (no code yet)
│  ├─ Main.bas                        # [To be implemented]
│  ├─ ConfigManager.bas               # [To be implemented]
│  ├─ TagBuilder.bas                  # [To be implemented]
│  ├─ Validator.bas                   # [To be implemented]
│  ├─ DataMapper.bas                  # [To be implemented]
│  ├─ QAEngine.bas                    # [To be implemented]
│  └─ AuditLog.bas                    # [To be implemented]
│
├─ templates/
│  ├─ Master_Template.xlsx            # Base template
│  ├─ CashTemplate.xlsx               # Cash cycle specialization
│  ├─ ARTemplate.xlsx                 # A/R specialization
│  ├─ APTemplate.xlsx                 # A/P specialization
│  └─ [etc]
│
├─ test_cases/
│  ├─ test_cases.md                   # Test scenarios
│  └─ sample_data/
│     ├─ cash_test_data.csv
│     ├─ ar_test_data.csv
│     └─ ap_test_data.csv
│
├─ config/
│  ├─ config_template.yml             # YAML template
│  ├─ engagement_defaults.yml         # Defaults by type
│  └─ tag_mappings.json               # Field mappings
│
├─ schemas/
│  ├─ tag_engine_schema.json          # TAG_ENGINE structure
│  ├─ coord_ref_schema.json           # COORD_REFERENCE structure
│  ├─ qa_schema.json                  # QA sheet structure
│  └─ output_schema.json              # OUTPUT structure
│
└─ VERSION.txt                        # Framework version
```

---

## Key Concepts Glossary

**DS_SEARCH**: Text-based DataSnipper tag; extracts values by searching for keywords

**DS_COORDS**: Coordinate-based DataSnipper tag; extracts from known pixel positions

**QA Rule**: Validation logic applied during processing (range check, lookup, business rule, etc)

**Exception**: Data that fails a QA rule; flagged for human review

**Override**: Manual correction of extracted value with justification

**Audit Trail**: Complete record of all processing events (AUDIT_LOG sheet)

**Schema**: Defined output format and required fields

**Control**: Audit procedure designed to prevent or detect errors

---

## Support & Contact

**Framework Questions:**
- Documentation: See README.md for document navigation
- Issues: Check DEVELOPMENT.md Debugging Guide

**DataSnipper Questions:**
- Reference: TAG_SPEC.md for tag syntax
- External: Contact DataSnipper support

**Engagement Issues:**
- Config: See CONFIG_GUIDE.md
- Process: See WORKFLOW.md
- Controls: See CONTROL_FRAMEWORK.md

---

## Next Steps

### If You're Starting a New Engagement:
1. Read: README.md (5 min overview)
2. Read: CONFIG_GUIDE.md (understand configuration)
3. Do: Follow setup steps in CONFIG_GUIDE.md
4. Reference: WORKFLOW.md during process

### If You're Developing This Framework:
1. Read: ARCHITECTURE.md (understand design)
2. Read: MODULE_SPECIFICATIONS.txt (understand requirements)
3. Read: DEVELOPMENT.md (understand development process)
4. Do: Implement modules per specs
5. Test: Follow test_cases/test_cases.md

### If You're Auditing the Framework:
1. Read: CONTROL_FRAMEWORK.md (understand controls)
2. Reference: QA_RULES.md (understand validation)
3. Review: AUDIT_LOG sheet (verify compliance)
4. Test: Run through WORKFLOW.md steps manually

---

**Framework Version**: 1.0  
**Last Updated**: June 2026  
**Status**: Architecture Complete - Ready for Development
