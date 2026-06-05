# TAG_SPEC.md - DataSnipper Tag Standards & Syntax

## Overview

This document defines the standard tag syntax for DataSnipper integration. TagBuilder module uses these specifications to generate valid DS_SEARCH and DS_COORDS tags programmatically.

---

## DataSnipper Tag Types

### 1. DS_SEARCH Tags (Text-Based Extraction)

**Purpose**: Extract values by searching for keywords or text patterns. Useful when position is variable or document structure differs.

#### Syntax Format
```
DS_SEARCH:<field_id>:<output_field>:(parameters)
```

#### Components

| Component | Description | Example |
|-----------|-------------|---------|
| `DS_SEARCH` | Tag type identifier (required) | `DS_SEARCH` |
| `field_id` | Unique field identifier | `Deposit_Amount` |
| `output_field` | Final output field name | `dep_amount` |
| `parameters` | Extraction logic (pipe-separated) | `start=Total\|end=Date` |

#### Parameter Reference

| Parameter | Purpose | Example | Notes |
|-----------|---------|---------|-------|
| `start` | Starting anchor text | `start=Total` | DataSnipper finds this text and starts extracting after |
| `end` | Ending anchor text | `end=Date` | DataSnipper extracts until it finds this text |
| `type` | Data type hint | `type=currency` | `text`, `number`, `currency`, `date`, `percentage` |
| `context` | Search context window | `context=paragraph` | `line`, `paragraph`, `page`, `document` |
| `case_sensitive` | Case matching | `case_sensitive=no` | `yes` or `no` (default: no) |
| `regex` | Regex pattern (advanced) | `regex=[0-9]{2,}` | Valid regex pattern |
| `fallback` | Alternative anchor if primary fails | `fallback=Amount\|Total` | Pipe-separated alternatives |

#### DS_SEARCH Examples

**Example 1: Extract Invoice Amount**
```
DS_SEARCH:InvoiceAmount:invoice_total:(start=Total Amount:|end=Tax|type=currency)
```
- Finds text "Total Amount:"
- Extracts content until "Tax" is encountered
- Interprets value as currency

**Example 2: Extract Account Number (Complex)**
```
DS_SEARCH:AccountNum:account_number:(start=Account #|end=Expiration|type=text|context=paragraph)
```
- Finds "Account #"
- Extracts until "Expiration"
- Limits search to current paragraph

**Example 3: Extract Date with Fallback**
```
DS_SEARCH:TransDate:transaction_date:(start=Date:|fallback=Trans Date:|end=Reference|type=date)
```
- Tries to find "Date:" first
- If not found, falls back to "Trans Date:"
- Extracts until "Reference"

**Example 4: Regex Pattern (Advanced)**
```
DS_SEARCH:InvoiceNum:inv_number:(regex=INV-[0-9]{6}|type=text)
```
- Uses regex pattern to find invoice numbers
- Pattern: INV- followed by 6 digits

---

### 2. DS_COORDS Tags (Coordinate-Based Extraction)

**Purpose**: Extract values from known, fixed positions. Best for structured documents where position is consistent.

#### Syntax Format
```
DS_COORDS:<field_id>:<output_field>:(parameters)
```

#### Components

| Component | Description | Example |
|-----------|-------------|---------|
| `DS_COORDS` | Tag type identifier (required) | `DS_COORDS` |
| `field_id` | Unique field identifier | `DepositDate_P1` |
| `output_field` | Final output field name | `dep_date` |
| `parameters` | Position + size specs | `page=1\|x=150\|y=320\|width=100\|height=20\|type=date` |

#### Parameter Reference

| Parameter | Purpose | Example | Notes |
|-----------|---------|---------|-------|
| `file` | Source document filename | `file=Invoice_2026.pdf` | Must match document loaded in DataSnipper |
| `page` | Page number (1-based) | `page=2` | Page number within document |
| `x` | X coordinate (pixels) | `x=150` | Horizontal position from left |
| `y` | Y coordinate (pixels) | `y=320` | Vertical position from top |
| `width` | Selection width (pixels) | `width=100` | Width of extraction box |
| `height` | Selection height (pixels) | `height=20` | Height of extraction box |
| `type` | Data type hint | `type=currency` | `text`, `number`, `currency`, `date`, `percentage` |
| `tolerance` | Fuzzy matching tolerance | `tolerance=5` | Pixel tolerance for position variation |

#### DS_COORDS Examples

**Example 1: Extract Bank Name (Precise Position)**
```
DS_COORDS:BankNameP1:bank_name:(page=1|x=50|y=100|width=150|height=20|type=text)
```
- Page 1 of document
- Starting at x=50, y=100 (pixels from top-left)
- Selection box 150px wide, 20px tall
- Extracts as text

**Example 2: Extract Amount with Tolerance**
```
DS_COORDS:DepositAmtP3:deposit_amount:(page=3|x=400|y=250|width=80|height=20|type=currency|tolerance=10)
```
- Page 3
- Position x=400, y=250 with ±10 pixel tolerance
- Interprets as currency
- Handles slight position variations

**Example 3: Multi-Page Example (Account Number)**
```
DS_COORDS:AccountNumP2:account_number:(file=AccountSummary.pdf|page=2|x=75|y=155|width=120|height=18|type=text)
```
- Specific file reference
- Page 2
- Specific coordinates
- Text type

---

## 3. HYBRID Tags (Mixed Strategy)

**Purpose**: Combine DS_SEARCH and DS_COORDS for maximum robustness. Try coordinate extraction first, fall back to search if position varies.

#### Syntax Format
```
DS_HYBRID:<field_id>:<output_field>:(coordinates)|fallback_to_search
```

#### Example
```
DS_HYBRID:ReceiptTotal:receipt_total:(page=1|x=350|y=450|width=80|height=20|type=currency)|fallback_to:(start=Receipt Total:|end=Tax Amount|type=currency)
```
- First tries coordinate-based extraction at specified position
- If coordinates don't yield result, falls back to DS_SEARCH with keywords
- Ensures maximum success rate

---

## Tag Naming Conventions

### Field ID Naming (Used for Auditing/Debugging)
```
<engagement_type>_<data_element>_<method_indicator>

Examples:
  CASH_DepositAmount_S        (Cash engagement, Deposit Amount, DS_SEARCH)
  AR_InvoiceDate_C            (AR engagement, Invoice Date, DS_COORDS)
  AP_PoTotal_H                (AP engagement, PO Total, HYBRID)
  CONTRACTS_SignDate_C        (Contracts, Sign Date, DS_COORDS)
```

### Output Field Naming (Final Output Column Names)
```
<entity>_<attribute>_<qualifier>

Examples:
  deposit_amount              (simple)
  invoice_date_period_1       (multiple instances per doc)
  po_number_vendor_acme       (unique qualifier)
  contract_effective_date     (multi-word readable)
```

---

## Tag Syntax Validation Rules

### DS_SEARCH Validation
```
✓ VALID:
  DS_SEARCH:Amount:amount:(start=Total|end=Tax|type=currency)
  DS_SEARCH:Date:date:(start=Date:|type=date)
  DS_SEARCH:Name:vendor_name:(start=From:|end=Date|case_sensitive=no)

✗ INVALID (Common Mistakes):
  DS_SEARCH:Amount:amount                           (missing parameters)
  ds_search:Amount:amount:(...)                     (case-sensitive)
  DS_SEARCH:Amount:amount:(start=Total&end=Tax)   (wrong separator; use |)
  DS_SEARCH:Amount:amount:(start=Total|end=      (incomplete parameters)
```

### DS_COORDS Validation
```
✓ VALID:
  DS_COORDS:Amount:amount:(page=1|x=100|y=200|width=80|height=20|type=currency)
  DS_COORDS:Date:date:(page=2|x=50|y=150|width=100|height=18|type=date)

✗ INVALID (Common Mistakes):
  DS_COORDS:Amount:amount:(page=1|x=100)          (missing required y, width, height)
  DS_COORDS:Amount:amount:(x=100|y=200|...)       (missing page)
  DS_COORDS:Amount:amount:(page=0|x=100|...)      (page 0 invalid; must be 1+)
  DS_COORDS:Amount:amount:(page=1|x=abc|...)      (x must be numeric)
```

---

## Tag Generation Strategy (For TagBuilder)

### Rule 1: Prefer DS_SEARCH When
- Document structure varies between instances
- Position is inconsistent
- Text anchors are reliable and unique
- Keyword extraction is more maintainable

### Rule 2: Prefer DS_COORDS When
- Document structure is fixed/standardized
- Position is consistent across all instances
- Text anchors are ambiguous or duplicated
- Performance is critical (coordinates faster)

### Rule 3: Use HYBRID When
- Document may vary slightly
- Maximum robustness required
- Both coordinates and search anchors available

---

## Example Tag Library by Engagement Type

### Cash Engagement Tags
```
DS_SEARCH:DepositAmount:deposit_amount:(start=Deposit|end=Total|type=currency)
DS_COORDS:DepositDate:deposit_date:(page=1|x=300|y=450|width=100|height=20|type=date)
DS_SEARCH:BankName:bank_name:(start=Bank|end=Account|type=text)
DS_COORDS:AccountNumber:account_number:(page=1|x=50|y=100|width=150|height=20|type=text)
```

### A/R Engagement Tags
```
DS_SEARCH:InvoiceNumber:invoice_num:(start=Invoice #|end=Date|type=text)
DS_SEARCH:InvoiceAmount:invoice_amount:(start=Total|end=Tax|type=currency)
DS_COORDS:DueDate:due_date:(page=1|x=400|y=200|width=100|height=20|type=date)
DS_SEARCH:CustomerName:customer_name:(start=Bill To|end=Invoice|type=text)
```

### A/P Engagement Tags
```
DS_SEARCH:PoNumber:po_number:(start=PO #|end=Date|type=text)
DS_COORDS:PoAmount:po_amount:(page=1|x=350|y=300|width=100|height=20|type=currency)
DS_SEARCH:VendorName:vendor_name:(start=From|end=Invoice|type=text)
DS_COORDS:DueDate:due_date:(page=1|x=400|y=250|width=100|height=20|type=date)
```

---

## Data Type Specifications

### type=text
- No formatting applied
- Extracted as-is
- Example: "Chase Bank"

### type=number
- Numeric characters only
- Symbols (currency, %, commas) removed
- Example: "1250.50" (even if extracted as "$1,250.50")

### type=currency
- Numeric value with currency symbol
- Format: $#,##0.00
- Example: "$1,250.50"

### type=date
- Standardized to YYYY-MM-DD
- Detects multiple input formats:
  - 6/1/26 → 2026-06-01
  - 01-Jun-2026 → 2026-06-01
  - 2026/06/01 → 2026-06-01

### type=percentage
- Numeric value 0-100
- Format: #.##%
- Example: "95.5%"

---

## Testing Tags Before Production

### Manual Verification Checklist
```
[ ] Tag syntax is valid per this specification
[ ] Field IDs match TAG_ENGINE definitions
[ ] Output fields match expected output schema
[ ] Data types match expected schema
[ ] Required parameters present (start/end for search, page/x/y for coords)
[ ] Parameter values are realistic and match document content
[ ] Tag tested against sample document (if available)
[ ] Fallback parameters work if primary fails
[ ] Case sensitivity handled correctly
[ ] Special characters escaped if needed
```

---

## Common Tag Patterns

### Pattern 1: Unique Anchor Extraction
```
DS_SEARCH:FieldName:field_name:(start=UniqueAnchor|end=NextField|type=text)
```
Use when: Starting anchor is unique in document

### Pattern 2: Table Cell Extraction
```
DS_COORDS:TableCell:cell_value:(page=2|x=100|y=150|width=60|height=15|type=currency)
```
Use when: Values in fixed table positions

### Pattern 3: Multi-Instance with Page Number
```
DS_COORDS:Amount_Page2:amount_page_2:(page=2|x=350|y=300|width=80|height=20|type=currency)
DS_COORDS:Amount_Page3:amount_page_3:(page=3|x=350|y=300|width=80|height=20|type=currency)
```
Use when: Extracting same field from multiple pages

### Pattern 4: Fallback for Ambiguous Anchor
```
DS_SEARCH:Amount:amount:(start=Grand Total|fallback=Total Amount|fallback=Net Total|end=Tax|type=currency)
```
Use when: Multiple possible anchor texts

---

## Error Handling in Tags

### If DataSnipper Cannot Find Start Anchor
- DS_SEARCH will not extract value
- Value marked as null in extraction output
- QAEngine flags as validation error
- QA team reviews and decides (accept, override, or reject)

### If Coordinates Out of Bounds
- DS_COORDS will not extract value
- Value marked as null
- QAEngine flags for review

### If Extracted Value Doesn't Match Data Type
- DataMapper attempts conversion
- If conversion fails, marked for QA review
- QA team can accept extracted value or override

---

**Framework Version**: 1.0  
**Last Updated**: June 2026
