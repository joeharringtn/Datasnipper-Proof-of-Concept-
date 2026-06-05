# DataSnipper VBA Automation Framework

## Executive Summary

This is an enterprise-grade, low-code automation framework designed to enable audit teams to extract structured data from documents using DataSnipper's tag-based execution model, with built-in validation, quality assurance, and control workflows.

The framework operates entirely within Excel and DataSnipper constraints—no external APIs, no programmatic DataSnipper triggers, only tag-based extraction.

---

## Framework Goals

✓ **Modularity**: Reusable across audit engagements (Cash, A/R, A/P, Contracts, Inventory, etc.)  
✓ **Low-Code**: Minimal VBA; maximum configuration  
✓ **Enterprise-Ready**: Built-in controls, audit trails, exception handling  
✓ **Scalability**: Supports multi-document, multi-page extractions  
✓ **Quality Assurance**: Post-snipping validation and human review workflows  
✓ **Version Control**: .bas files tracked in Git; easy deployment across teams  

---

## Project Structure

```
vba_framework/
├── vba_modules/               # VBA source code (.bas files)
│   ├── TagBuilder.bas         # Tag construction engine
│   ├── Validator.bas          # Input/output validation logic
│   ├── QAEngine.bas           # Post-snipping QA workflows
│   ├── ConfigManager.bas      # Config sheet management
│   ├── DataMapper.bas         # Map extracted data to schema
│   ├── AuditLog.bas           # Audit trail / exception logging
│   └── Main.bas               # Entry points / UI orchestration
│
├── docs/                      # Documentation
│   ├── README.md              # This file
│   ├── ARCHITECTURE.md        # Module design and interactions
│   ├── WORKFLOW.md            # End-to-end process flows
│   ├── TAG_SPEC.md            # DataSnipper tag standards
│   ├── QA_RULES.md            # Validation rules and QA logic
│   ├── CONFIG_GUIDE.md        # Configuration customization
│   ├── DATA_MODEL.md          # Excel schema definitions
│   ├── CONTROL_FRAMEWORK.md   # Audit control standards
│   └── DEVELOPMENT.md         # Dev workflow and CI/CD
│
├── templates/                 # Excel templates and samples
│   ├── Master_Template.xlsx   # Blank engagement template
│   ├── Sample_Cash.xlsx       # Cash cycle example
│   ├── Sample_AR.xlsx         # Accounts receivable example
│   └── Sample_AP.xlsx         # Accounts payable example
│
├── test_cases/                # Test data and validation cases
│   ├── test_cases.md          # Test scenario definitions
│   ├── sample_data/           # Sample extractions
│   │   ├── cash_test.csv
│   │   ├── ar_test.csv
│   │   └── ap_test.csv
│   └── validation_checks.md   # Expected vs. actual results
│
├── config/                    # Configuration files
│   ├── config_template.yml    # YAML config structure
│   ├── engagement_defaults.yml# Default settings by engagement type
│   └── tag_mappings.json      # Tag-to-field mappings
│
├── schemas/                   # Data model definitions
│   ├── tag_engine_schema.json # TAG_ENGINE sheet structure
│   ├── coord_ref_schema.json  # COORD_REFERENCE structure
│   ├── qa_schema.json         # QA sheet structure
│   └── output_schema.json     # OUTPUT sheet structure
│
└── VERSION.txt                # Framework version
```

---

## How It Works (High Level)

```
1. SETUP PHASE
   ├─ Load engagement configuration (engagement type, document paths, page numbers)
   ├─ Populate TAG_ENGINE sheet with extraction rules
   └─ ConfigManager validates configuration completeness

2. TAG BUILDING PHASE
   ├─ TagBuilder reads TAG_ENGINE sheet
   ├─ Constructs DS_SEARCH and DS_COORDS tags
   ├─ Outputs tags to dedicated column in Excel
   └─ User copies tags into DataSnipper UI

3. SNIPPING PHASE (EXTERNAL)
   ├─ User runs DataSnipper with built tags
   ├─ DataSnipper extracts values and pastes results
   └─ User pastes extracted data back into Excel

4. VALIDATION PHASE
   ├─ Validator checks input completeness and formats
   ├─ DataMapper transforms raw extractions to schema
   ├─ QAEngine applies business logic rules
   └─ Exceptions flagged for human review

5. QA REVIEW PHASE
   ├─ User reviews flagged items in QA sheet
   ├─ Manual overrides recorded with justification
   ├─ AuditLog captures all exceptions and decisions
   └─ Sign-off workflow for control compliance

6. OUTPUT PHASE
   ├─ Cleansed data moved to OUTPUT sheet
   ├─ Full audit trail available for review
   └─ Results exported for downstream analysis
```

---

## Key Concepts

### Tags (DataSnipper Integration)

- **DS_SEARCH**: Text-based tag for finding data by keyword/pattern
- **DS_COORDS**: Coordinate-based tag for structured, known-position extractions
- Tags are built dynamically in Excel and manually executed in DataSnipper

### Modules (VBA Architecture)

| Module | Purpose | Key Responsibilities |
|--------|---------|----------------------|
| **TagBuilder** | Generate DS_SEARCH/DS_COORDS tags | Construct tags from config; validate tag syntax |
| **Validator** | Input validation | Check completeness, formats, required fields |
| **DataMapper** | Schema transformation | Map raw extractions to normalized schema |
| **QAEngine** | Business logic validation | Apply audit-specific QA rules; flag exceptions |
| **ConfigManager** | Configuration management | Load/validate engagement settings |
| **AuditLog** | Exception tracking | Record all exceptions, overrides, decisions |
| **Main** | Orchestration | Call modules in sequence; manage UI flow |

### Configuration (Reusability)

- **Engagement Type**: Cash, A/R, A/P, Contracts, etc.
- **Document Paths**: File paths and page numbers for each source
- **Tag Definitions**: What to extract, from where, using which method
- **QA Rules**: Engagement-specific validation logic
- **Output Schema**: Required fields and formats for downstream use

---

## Excel Data Model (Sheets)

| Sheet Name | Purpose | Status | User Input |
|------------|---------|--------|-----------|
| **CONFIG** | Engagement configuration; reusability settings | Master config | Yes (once per engagement) |
| **TAG_ENGINE** | Tag definitions; DS_SEARCH/DS_COORDS rules | Working sheet | Yes (collaborative) |
| **COORD_REFERENCE** | Coordinate mappings for DS_COORDS tags | Reference | No (read-only) |
| **EXTRACTION_INPUT** | Raw DataSnipper outputs (paste results here) | Working sheet | Yes (from DataSnipper) |
| **VALIDATION** | Intermediate validation results | System-generated | No (read-only) |
| **QA** | Exceptions and manual overrides | Working sheet | Yes (QA team) |
| **OUTPUT** | Final, cleansed extraction results | Final output | No (read-only) |
| **AUDIT_LOG** | Full exception history and decisions | Audit trail | No (read-only) |

---

## Development Workflow

### For Developers

1. **Create .bas files** in `vba_modules/` folder using VS Code or text editor
2. **Import into Excel** using VBA IDE or automated import tool
3. **Test iteratively** in Excel; keep .bas files as source of truth
4. **Version control** .bas files in Git
5. **Deploy** by importing modules into engagement workbooks

### For Audit Teams

1. **Select engagement template** (Cash, A/R, etc.)
2. **Customize CONFIG sheet** for specific engagement (file paths, page numbers, QA rules)
3. **Build tags** using TagBuilder module
4. **Copy tags into DataSnipper** and run extraction
5. **Paste results** back into Excel
6. **Review QA sheet** for exceptions
7. **Approve and export** final OUTPUT sheet

---

## Control Framework

### Built-In Audit Controls

| Control | Responsibility | Implementation |
|---------|-----------------|-----------------|
| **Input Validation** | Ensure completeness before processing | Validator module |
| **Format Validation** | Check data types, ranges, required fields | DataMapper module |
| **Business Logic Validation** | Apply engagement-specific rules | QAEngine module |
| **Exception Tracking** | Record all deviations for audit trail | AuditLog module |
| **Manual Override Log** | Capture justification for exceptions | QA sheet + AuditLog |
| **Sign-Off Workflow** | Required approval before finalization | Main module UI |
| **Version Control** | Track module changes and deployments | Git + VERSION.txt |

---

## Getting Started

### Phase 1: Framework Setup (Developer)
- [ ] Review ARCHITECTURE.md for module design
- [ ] Create VBA modules in `vba_modules/` folder
- [ ] Validate module interactions per ARCHITECTURE.md
- [ ] Create unit tests in `test_cases/`

### Phase 2: Configuration (Architect)
- [ ] Customize config templates in `config/` for target engagements
- [ ] Define engagement-specific QA rules in QA_RULES.md
- [ ] Map data schemas in `schemas/` for each engagement type

### Phase 3: Testing (QA Team)
- [ ] Test with sample data in `test_cases/`
- [ ] Validate tag generation per TAG_SPEC.md
- [ ] Verify QA rule logic per QA_RULES.md
- [ ] Sign off on control effectiveness

### Phase 4: Deployment (Operations)
- [ ] Package framework into templates in `templates/`
- [ ] Train audit teams on workflow per WORKFLOW.md
- [ ] Deploy templates and provide CONFIG_GUIDE.md to teams
- [ ] Monitor AuditLog for operational issues

---

## File Organization Strategy

All VBA source files (`.bas`) are maintained in version control under `vba_modules/`. This allows:
- **Git tracking** of all code changes
- **Code review** workflows before deployment
- **Rollback capability** if issues arise
- **Easy redeployment** across engagements
- **Documentation** of changes per module

Excel workbooks (`.xlsx`) contain:
- Configuration (one per engagement)
- Formulas linking to modules
- Data sheets for input, processing, output
- BUT NOT the VBA code itself (imported from `.bas` files)

---

## Next Steps

1. **Review [ARCHITECTURE.md](docs/ARCHITECTURE.md)** for detailed module design and interactions
2. **Review [WORKFLOW.md](docs/WORKFLOW.md)** for end-to-end process flows
3. **Review [TAG_SPEC.md](docs/TAG_SPEC.md)** for DataSnipper tag standards
4. **Review [CONFIG_GUIDE.md](docs/CONFIG_GUIDE.md)** for customization patterns
5. **Review [DATA_MODEL.md](docs/DATA_MODEL.md)** for Excel schema and sheet definitions
6. **Review [CONTROL_FRAMEWORK.md](docs/CONTROL_FRAMEWORK.md)** for audit control implementation

---

**Framework Version**: 1.0  
**Last Updated**: June 2026  
**Status**: Architecture Phase Complete
