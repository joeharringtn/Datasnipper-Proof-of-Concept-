# Audit Framework - Python Refactor

Enterprise-grade audit automation framework for DataSnipper extraction with validation, QA workflows, and approval processes.

**This is a complete Python rewrite of the VBA framework**, maintaining the same modular architecture while providing better maintainability, testability, and audit compliance.

---

## 🏗 Architecture

The framework is organized into 7 core modules mirroring the VBA design:

```
┌──────────────────────────────────┐
│      CLI / Main Orchestrator     │  - Entry points
│      (audit_framework.cli)       │  - Workflow sequencing
└──────────────────────────────────┘
           │
    ┌──────┼──────┬──────────┬──────────┐
    ▼      ▼      ▼          ▼          ▼
┌────────┐┌──────────┐┌──────────┐┌──────────┐
│Config  ││TagBuilder││Validator ││DataMapper│
│Manager ││          ││          ││          │
└────────┘└──────────┘└──────────┘└──────────┘
    │         │            │          │
    └─────────┼────────────┼──────────┘
              │            │
              ▼            ▼
        ┌──────────┐┌──────────┐
        │QAEngine  ││AuditLog  │
        └──────────┘└──────────┘
```

### Modules

| Module | Purpose | Key Classes |
|--------|---------|------------|
| **ConfigManager** | Load & validate engagement configuration | `ConfigManager` |
| **TagBuilder** | Generate DataSnipper extraction tags | `TagBuilder` |
| **Validator** | Input validation & format checking | `Validator` |
| **DataMapper** | Transform & normalize extracted data | `DataMapper` |
| **QAEngine** | Quality assurance rule application | `QAEngine`, `QAException` |
| **AuditLog** | Audit trail & event logging | `AuditLog`, `AuditLogEntry` |
| **CLI** | Command orchestration & UI | `cli` (Click commands) |

---

## 📦 Installation

### Prerequisites
- Python 3.11+
- `pip` package manager

### Setup

1. **Navigate to the project directory:**
   ```bash
   cd python_refactor
   ```

2. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

   Or with development tools:
   ```bash
   pip install -e ".[dev]"
   ```

3. **Verify installation:**
   ```bash
   python -m audit_framework.cli --help
   ```

---

## 🚀 Quick Start

### 1. Load Engagement Configuration

```bash
python -m audit_framework.cli load-engagement path/to/engagement.xlsx --user "john.doe"
```

This:
- Loads CONFIG sheet
- Validates completeness
- Checks engagement dates and approvers
- Logs to audit trail

### 2. Build Extraction Tags

```bash
python -m audit_framework.cli build-tags path/to/engagement.xlsx --user "john.doe"
```

This:
- Reads TAG_ENGINE sheet
- Generates DS_SEARCH and DS_COORDS tags
- Validates tag syntax
- Displays sample tags

### 3. Validate Extracted Data

```bash
python -m audit_framework.cli validate-input path/to/engagement.xlsx --user "auditor"
```

This:
- Reads EXTRACTION_INPUT sheet
- Checks required fields
- Validates data types and formats
- Reports validation errors

### 4. Process & Apply QA Rules

```bash
python -m audit_framework.cli process-extraction path/to/engagement.xlsx --user "auditor"
```

This:
- Transforms raw data to schema
- Applies QA rules
- Flags exceptions
- Logs transformations

### 5. Check Status

```bash
python -m audit_framework.cli show-status path/to/engagement.xlsx
```

Displays engagement summary, schema info, and audit log overview.

---

## 📊 Data Models

All data structures use **Pydantic v2** for validation and type safety:

### Core Models

```python
from audit_framework.models import (
    ConfigObject,
    TagDefinition,
    QARule,
    SchemaField,
    ApprovalWorkflow,
)

# Create config programmatically
config = ConfigObject(
    engagement_id="2026-CASH-01",
    engagement_type=EngagementType.CASH,
    period_start_date=date(2026, 1, 1),
    period_end_date=date(2026, 12, 31),
    lead_auditor="John Doe",
    client_name="Acme Corp",
)
```

### Enums

```python
class ExtractionMethod(str, Enum):
    DS_SEARCH = "DS_SEARCH"
    DS_COORDS = "DS_COORDS"
    HYBRID = "HYBRID"

class RuleType(str, Enum):
    RANGE = "RANGE"
    LOOKUP = "LOOKUP"
    FORMAT = "FORMAT"
    CROSS_FIELD = "CROSS_FIELD"
    DUPLICATE = "DUPLICATE"
    CUSTOM = "CUSTOM"
```

---

## 💾 Excel Integration

The framework reads/writes Excel files using **openpyxl**:

### Reading Data

```python
from audit_framework.excel_utils import ExcelWorkbookManager

with ExcelWorkbookManager("engagement.xlsx") as mgr:
    # Read config value
    engagement_id = mgr.read_config_value("EngagementID")
    
    # Read rows as dictionaries
    rows = mgr.read_rows_as_dicts("EXTRACTION_INPUT")
```

### Writing Data

```python
with ExcelWorkbookManager("engagement.xlsx") as mgr:
    # Write config value
    mgr.write_cell_by_ref("CONFIG", "B1", "VALUE")
    
    # Write rows
    data = [{"Name": "Test", "Amount": "100"}]
    mgr.write_rows_from_dicts("OUTPUT", data)
    
    mgr.save()
```

---

## 🔍 Usage Examples

### Programmatic API

```python
from audit_framework.config_manager import ConfigManager
from audit_framework.tag_builder import TagBuilder
from audit_framework.validator import Validator
from audit_framework.qa_engine import QAEngine
from audit_framework.audit_log import AuditLog

# Load configuration
config_mgr = ConfigManager("engagement.xlsx")
config = config_mgr.load_config()

# Generate tags
tags = TagBuilder.build_all_tags(config.tags)

# Validate data
rows = read_data_from_excel()
validation_errors = Validator.validate_batch(rows, config.output_schema)

# Apply QA rules
qa_engine = QAEngine()
qa_results = qa_engine.apply_rules_to_batch(config.qa_rules, rows)

# Log events
audit_log = AuditLog("engagement.xlsx")
audit_log.log_config_loaded(config.engagement_id, user="john.doe")
audit_log.log_tags_built(len(tags), user="john.doe")
audit_log.save()
```

### CLI Usage

```bash
# Help for any command
audit-cli load-engagement --help

# Load with custom user
audit-cli load-engagement data/engagement.xlsx --user "audit-team"

# Build tags and write to specific sheet
audit-cli build-tags data/engagement.xlsx --output-sheet "TAGS"

# Validate with error reporting
audit-cli validate-input data/engagement.xlsx --input-sheet "RAW_DATA"

# Full workflow
audit-cli load-engagement engagement.xlsx --user "manager"
audit-cli build-tags engagement.xlsx
audit-cli validate-input engagement.xlsx
audit-cli process-extraction engagement.xlsx --user "qa-team"
audit-cli show-status engagement.xlsx
```

---

## 🧪 Testing

The framework includes comprehensive pytest tests:

```bash
# Run all tests
pytest tests/

# Run specific test file
pytest tests/test_config_manager.py

# Run with coverage
pytest tests/ --cov=audit_framework --cov-report=html

# Run specific test
pytest tests/test_tag_builder.py::test_build_search_tag
```

### Test Files

- `test_config_manager.py` - Configuration loading & validation
- `test_tag_builder.py` - Tag generation
- `test_validator.py` - Data validation
- `test_qa_engine.py` - QA rule application

---

## 📝 Workflow States

The framework tracks engagement state through the workflow:

```
CONFIG_LOADED → TAGS_READY → EXTRACTION_SUBMITTED 
→ VALIDATION_COMPLETE → QA_IN_PROGRESS → QA_REVIEWED 
→ APPROVAL_PENDING → COMPLETE
```

---

## 🔐 Audit Trail

All events are logged to AUDIT_LOG sheet:

| Column | Description |
|--------|-------------|
| Timestamp | When event occurred |
| EventType | CONFIG_LOADED, TAGS_BUILT, etc. |
| Message | Human-readable description |
| User | Who triggered event |
| Module | Which module generated event |
| RowID | Row affected (for data changes) |
| FieldName | Field affected |
| OldValue | Previous value |
| NewValue | New value |
| Details | Additional info (JSON) |

---

## 🔗 Differences from VBA

| Aspect | VBA | Python |
|--------|-----|--------|
| **Language** | VB6 dialect | Python 3.11+ |
| **Data Validation** | Manual | Pydantic v2 |
| **Type Safety** | Implicit | Type hints throughout |
| **Testing** | Manual | pytest suite |
| **CLI** | Excel buttons | Click CLI commands |
| **Error Handling** | Try/Catch | Custom exceptions |
| **Logging** | Cell-based | Structured AuditLog |
| **Dependencies** | VBA built-ins | openpyxl, pydantic, click |

---

## 🛠 Development

### Project Structure

```
python_refactor/
├── audit_framework/           # Main package
│   ├── __init__.py
│   ├── models.py              # Pydantic data models
│   ├── exceptions.py          # Custom exceptions
│   ├── config_manager.py      # Configuration
│   ├── tag_builder.py         # Tag generation
│   ├── validator.py           # Validation logic
│   ├── data_mapper.py         # Data transformation
│   ├── qa_engine.py           # QA workflows
│   ├── audit_log.py           # Event logging
│   ├── excel_utils.py         # Excel I/O
│   └── cli.py                 # CLI commands
├── tests/                     # Test suite
│   ├── test_config_manager.py
│   ├── test_tag_builder.py
│   ├── test_validator.py
│   ├── test_qa_engine.py
│   └── ...
├── scripts/                   # Utility scripts
├── pyproject.toml             # Project config
├── requirements.txt           # Dependencies
└── README.md                  # This file
```

### Adding New Modules

1. Create module in `audit_framework/`
2. Add type hints and docstrings
3. Create tests in `tests/`
4. Add to `__init__.py` exports
5. Add CLI commands in `cli.py`

### Code Style

- **Formatting:** Black (line length: 100)
- **Sorting:** isort (Black profile)
- **Linting:** pylint
- **Type checking:** mypy
- **Testing:** pytest

```bash
# Format code
black audit_framework/ tests/

# Sort imports
isort audit_framework/ tests/

# Lint
pylint audit_framework/

# Type check
mypy audit_framework/
```

---

## 📚 API Reference

### ConfigManager

```python
config_mgr = ConfigManager(workbook_path)
config = config_mgr.load_config()
is_valid, errors = config_mgr.validate_config()
config_mgr.close()
```

### TagBuilder

```python
tag = TagBuilder.build_tag(tag_definition)
tags = TagBuilder.build_all_tags(tag_definitions)
is_valid, error = TagBuilder.validate_tag_syntax(tag_string)
```

### Validator

```python
is_valid, error = Validator.validate_required_field(value, field_name)
is_valid, error = Validator.validate_data_type(value, data_type, field_name)
errors = Validator.validate_batch(rows, schema)
normalized = Validator.normalize_value(value, data_type)
```

### DataMapper

```python
mapper = DataMapper()
mapped_row = mapper.map_row(raw_row, schema, tags)
mapped_rows = mapper.map_batch(raw_rows, schema, tags)
```

### QAEngine

```python
qa_engine = QAEngine()
exception = qa_engine.apply_rule(rule, value, row_data)
results = qa_engine.apply_rules_to_batch(rules, rows)
qa_engine.accept_exception(exception_id, user)
qa_engine.override_exception(exception_id, new_value, justification, user)
summary = qa_engine.to_dict()
```

### AuditLog

```python
audit_log = AuditLog(workbook_path)
audit_log.log_config_loaded(engagement_id, user=user)
audit_log.log_tags_built(tag_count, user=user)
audit_log.log_qa_exception(exc_id, rule_id, field_name, value, reason, user=user)
audit_log.write_to_sheet()
audit_log.save()
```

---

## ⚠️ Known Limitations & Future Work

### Current Limitations
- HYBRID tag building not yet fully tested with real DataSnipper output
- Custom QA rules require manual implementation
- No built-in reporting/PDF generation (yet)
- Approval workflow UI is CLI-based (no web dashboard yet)

### Future Enhancements
- [ ] Web dashboard (Flask/Streamlit)
- [ ] Database backend (SQLite/PostgreSQL)
- [ ] Advanced reporting & PDF export
- [ ] Audit-ready digital signatures
- [ ] Multi-user concurrent workflows
- [ ] DataSnipper API integration (if available)
- [ ] Template library for common engagement types

---

## 🐛 Troubleshooting

### Import Errors

```bash
# Ensure package is installed
pip install -e .

# Check Python path
python -c "import audit_framework; print(audit_framework.__file__)"
```

### Excel File Errors

```python
# Ensure workbook exists
from pathlib import Path
assert Path("engagement.xlsx").exists()

# Check worksheet names
mgr = ExcelWorkbookManager("engagement.xlsx")
print(mgr.workbook.sheetnames)
```

### Validation Failures

```python
# Get detailed errors
config_mgr = ConfigManager("engagement.xlsx")
config = config_mgr.load_config()
is_valid, errors = config_mgr.validate_config()
for error in errors:
    print(f"ERROR: {error}")
```

---

## 📞 Support

For issues or questions:
1. Check test files for usage examples
2. Review docstrings in source code
3. Refer to VBA framework docs (same logic)
4. Run tests: `pytest -v`

---

## 📄 License

MIT License - Same as parent VBA framework

---

## ✅ Implementation Checklist

- [x] Data models (Pydantic)
- [x] Exception handling
- [x] Excel I/O utilities
- [x] ConfigManager module
- [x] TagBuilder module
- [x] Validator module
- [x] DataMapper module
- [x] QAEngine module
- [x] AuditLog module
- [x] CLI interface
- [x] pytest test suite
- [x] Type hints throughout
- [x] Comprehensive docstrings
- [ ] Web dashboard (future)
- [ ] Database backend (future)
- [ ] Advanced reporting (future)

---

**Version:** 1.0.0  
**Last Updated:** 2026-06-04  
**Status:** ✅ Production Ready
