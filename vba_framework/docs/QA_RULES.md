# QA_RULES.md - Validation & Audit Control Rules

## Overview

This document defines the Quality Assurance (QA) rules applied by the QAEngine module to flag exceptions and ensure data quality. Rules are configurable per engagement type, allowing customization for specific audit objectives.

---

## QA Rule Framework

### Rule Structure
```
QARule:
  ├─ rule_id: Unique identifier
  ├─ rule_name: Descriptive name
  ├─ engagement_type: Cash | AR | AP | Contracts | Inventory
  ├─ field_name: Target field(s)
  ├─ rule_type: Category of validation
  ├─ rule_definition: Logic or formula
  ├─ fail_action: FLAG | BLOCK | WARN
  ├─ severity: INFO | WARNING | CRITICAL
  ├─ priority: 1 (highest) to 5 (lowest)
  └─ notes: Audit objective or reason for rule
```

---

## Rule Types & Categories

### 1. RANGE Validation (Numeric Bounds)

**Purpose**: Check that numeric values fall within acceptable bounds.

**Examples**

| Rule ID | Field | Min | Max | Fail Action | Severity |
|---------|-------|-----|-----|-------------|----------|
| RANGE_CASH_01 | deposit_amount | 0 | 1,000,000 | FLAG | WARNING |
| RANGE_CASH_02 | deposit_amount | 1,000 | 999,999 | BLOCK | CRITICAL |
| RANGE_AR_01 | invoice_amount | 0 | 500,000 | FLAG | WARNING |
| RANGE_AP_01 | po_amount | 0 | 1,000,000 | FLAG | WARNING |
| RANGE_AP_02 | po_amount | 100,000 | - | BLOCK | CRITICAL |

**Rule Definition Format**
```
value >= MIN AND value <= MAX
```

**Implementation Logic**
```
IF field_value < min_value OR field_value > max_value:
  FLAG exception
  exception_reason = "Value outside acceptable range"
  severity = WARNING
ENDIF

IF field_value < min_value * 0.1 OR field_value > max_value * 10:
  BLOCK processing
  exception_reason = "Value critically out of range"
  severity = CRITICAL
ENDIF
```

---

### 2. LOOKUP Validation (Reference Table)

**Purpose**: Check that extracted values exist in approved reference lists.

**Examples**

| Rule ID | Field | Lookup Table | Fail Action | Severity |
|---------|-------|--------------|-------------|----------|
| LOOKUP_CASH_01 | bank_name | APPROVED_BANKS | BLOCK | CRITICAL |
| LOOKUP_AR_01 | customer_name | CUSTOMER_MASTER | FLAG | WARNING |
| LOOKUP_AP_01 | vendor_name | APPROVED_VENDORS | BLOCK | CRITICAL |
| LOOKUP_AP_02 | payment_method | ALLOWED_METHODS | FLAG | WARNING |

**Reference Tables**

```
APPROVED_BANKS:
  ├─ Chase Bank
  ├─ Bank of America
  ├─ Wells Fargo
  ├─ US Bank
  └─ [additional banks per engagement]

CUSTOMER_MASTER (from A/R subsidiary ledger):
  ├─ Customer ID
  ├─ Customer Name
  └─ [loaded from CONFIG sheet]

APPROVED_VENDORS (from A/P master):
  ├─ Vendor ID
  ├─ Vendor Name
  ├─ Vendor Status (Active/Inactive)
  └─ [loaded from CONFIG sheet]

ALLOWED_METHODS:
  ├─ Check
  ├─ Wire Transfer
  ├─ ACH
  ├─ Credit Card
  └─ Other
```

**Implementation Logic**
```
reference_table = LOOKUP_TABLE_DEFINED_IN_CONFIG

IF extracted_value NOT IN reference_table:
  IF fail_action = BLOCK:
    STOP processing
    exception_reason = "Value not in approved list"
    severity = CRITICAL
  ELSE IF fail_action = FLAG:
    FLAG exception
    exception_reason = "Value not in reference table; may be new vendor/customer"
    severity = WARNING
ENDIF
```

---

### 3. FORMAT Validation (Pattern Matching)

**Purpose**: Validate that text values match expected patterns.

**Examples**

| Rule ID | Field | Pattern | Fail Action | Severity |
|---------|-------|---------|-------------|----------|
| FORMAT_CASH_01 | account_number | `^[0-9]{10,12}$` | FLAG | WARNING |
| FORMAT_AR_01 | invoice_num | `^INV-[0-9]{6}$` | FLAG | WARNING |
| FORMAT_AP_01 | po_number | `^PO-[0-9]{8}$` | FLAG | WARNING |
| FORMAT_GENERAL_01 | email_address | `^[^\s@]+@[^\s@]+\.[^\s@]+$` | FLAG | WARNING |

**Common Patterns**

```
Pattern             Use Case                Example
^[0-9]{10,12}$     Account Number          1234567890
^INV-[0-9]{6}$     Invoice Number          INV-001234
^PO-[0-9]{8}$      Purchase Order Number   PO-20260601
^[A-Z]{2}$         State Code              CA, NY
[0-9]{3}-[0-9]{2}  Social Security Number  123-45-6789
```

**Implementation Logic**
```
IF extracted_value MATCHES regex_pattern:
  ACCEPT value
  exception = None
ELSE:
  IF fail_action = BLOCK:
    STOP processing
    exception_reason = "Value format invalid"
    severity = CRITICAL
  ELSE IF fail_action = FLAG:
    FLAG exception
    exception_reason = "Value format unusual; may require manual verification"
    severity = WARNING
ENDIF
```

---

### 4. CROSS_FIELD Validation (Multi-Field Consistency)

**Purpose**: Validate consistency and logic across multiple fields in same record.

**Examples**

| Rule ID | Fields Involved | Logic | Fail Action | Severity |
|---------|-----------------|-------|-------------|----------|
| CROSS_CASH_01 | deposit_date, period_end | deposit_date <= period_end | FLAG | WARNING |
| CROSS_CASH_02 | deposit_amount, invoice_total | difference <= 1% | FLAG | WARNING |
| CROSS_AR_01 | invoice_date, due_date | due_date > invoice_date | FLAG | WARNING |
| CROSS_AP_01 | po_date, receipt_date | receipt_date >= po_date | FLAG | WARNING |
| CROSS_AP_02 | po_amount, invoice_amount | invoice_amount <= po_amount * 1.1 | FLAG | WARNING |

**Implementation Logic**

```
Rule: CROSS_CASH_01 - Deposit Date Within Period
  IF deposit_date > period_end_date:
    FLAG exception
    exception_reason = "Deposit dated after audit period end"
    severity = CRITICAL

Rule: CROSS_AP_02 - Invoice Within PO Amount + 10% Tolerance
  invoice_variance = (invoice_amount - po_amount) / po_amount
  IF invoice_variance > 0.10:  // More than 10% over PO
    FLAG exception
    exception_reason = "Invoice exceeds PO by more than 10%"
    severity = WARNING
  ENDIF IF invoice_variance < -0.10:  // More than 10% under PO
    FLAG exception
    exception_reason = "Invoice significantly under PO; possible duplicate payment?"
    severity = WARNING
  ENDIF

Rule: CROSS_AR_01 - Due Date After Invoice Date
  IF due_date <= invoice_date:
    FLAG exception
    exception_reason = "Due date not after invoice date"
    severity = CRITICAL
ENDIF
```

---

### 5. DUPLICATE Validation (Uniqueness)

**Purpose**: Identify potential duplicate records based on key fields.

**Examples**

| Rule ID | Key Fields | Engagement | Tolerance | Fail Action |
|---------|-----------|------------|-----------|-------------|
| DUP_CASH_01 | deposit_amount, deposit_date | Cash | Exact match | FLAG |
| DUP_AR_01 | invoice_num, customer_id | A/R | Exact match | FLAG |
| DUP_AP_01 | po_num, vendor_id | A/P | Exact match | FLAG |
| DUP_CASH_02 | deposit_amount, deposit_date | Cash | ±1 day, ±$0.01 | FLAG |

**Implementation Logic**

```
Rule: DUP_CASH_01 - Duplicate Deposit Detection
  FOR each record R in current_extraction:
    FOR each prior_record P in extracted_history:
      IF R.deposit_amount = P.deposit_amount AND
         R.deposit_date = P.deposit_date:
        FLAG exception on R
        exception_reason = "Duplicate deposit detected (matches prior record ID: P.id)"
        severity = WARNING
        suggested_action = "Review prior extraction; determine if same or different deposit"
    ENDFOR
  ENDFOR

Rule: DUP_AP_01 - Fuzzy Duplicate Detection
  FOR each invoice record I in current_extraction:
    FOR each prior_invoice P in extracted_history:
      amount_diff = ABS(I.invoice_amount - P.invoice_amount)
      date_diff = ABS(I.invoice_date - P.invoice_date)  // in days
      IF I.po_number = P.po_number AND
         amount_diff < $0.01 AND
         date_diff <= 1 day:
        FLAG exception on I
        exception_reason = "Potential duplicate invoice (similar amount/date as prior record)"
        severity = WARNING
    ENDFOR
  ENDFOR
```

---

### 6. CUSTOM Rules (Business Logic)

**Purpose**: Engagement-specific validation logic beyond standard categories.

**Examples**

| Rule ID | Logic | Engagement | Fail Action |
|---------|-------|------------|-------------|
| CUSTOM_CASH_01 | Required approval for deposits > $500k | Cash | BLOCK |
| CUSTOM_AR_01 | Invoice must have matching delivery receipt | A/R | FLAG |
| CUSTOM_AP_01 | 3-way match: PO, Receipt, Invoice | A/P | BLOCK |
| CUSTOM_CONTRACTS_01 | Contract signed before effective date | Contracts | FLAG |

**Implementation**

```
Rule: CUSTOM_AP_01 - 3-Way Match (PO + Receipt + Invoice)
  REQUIRED_MATCHES = 3
  matched_count = 0
  
  IF po_exists_for_vendor_and_amount:
    matched_count += 1
  ENDIF
  
  IF receipt_exists_for_vendor_and_date:
    matched_count += 1
  ENDIF
  
  IF invoice_exists_for_po_and_amount:
    matched_count += 1
  ENDIF
  
  IF matched_count < REQUIRED_MATCHES:
    BLOCK processing
    exception_reason = "3-way match incomplete: only " + matched_count + " of 3 documents found"
    severity = CRITICAL
  ENDIF
```

---

## Rule Assignment by Engagement Type

### CASH Engagement Rules
```
RANGE:
  - deposit_amount: $0 - $1,000,000
  - account_balance: $0 - $10,000,000

LOOKUP:
  - bank_name: APPROVED_BANKS

CROSS_FIELD:
  - deposit_date within period
  - account_balance reconciles to GL

DUPLICATE:
  - deposit_amount + deposit_date combinations

CUSTOM:
  - Large deposits require approval
  - Daily totals reconcile to bank statement
```

### A/R Engagement Rules
```
RANGE:
  - invoice_amount: $0 - $500,000

LOOKUP:
  - customer_name: CUSTOMER_MASTER
  - invoice_status: [Open, Paid, Disputed]

CROSS_FIELD:
  - due_date > invoice_date
  - invoice_date within audit period

DUPLICATE:
  - invoice_number + customer_id combinations

CUSTOM:
  - Invoices over $100k require supporting documentation
  - Payment terms must be standard (Net 30, etc)
```

### A/P Engagement Rules
```
RANGE:
  - po_amount: $0 - $1,000,000
  - invoice_amount: $0 - $1,000,000

LOOKUP:
  - vendor_name: APPROVED_VENDORS
  - payment_method: ALLOWED_METHODS

CROSS_FIELD:
  - invoice_date after PO date
  - invoice_amount within 10% of PO
  - payment_date after invoice_date

DUPLICATE:
  - po_number + vendor_id + amount combinations

CUSTOM:
  - 3-way match required (PO + Receipt + Invoice)
  - Invoices > $50k require approval
```

---

## Fail Actions & Severity Levels

### Fail Action Types

| Action | Meaning | Impact | Usage |
|--------|---------|--------|-------|
| **FLAG** | Mark for QA review | Non-blocking; user decides | Default for warnings |
| **BLOCK** | Stop processing | Prevents finalization | Critical issues only |
| **WARN** | Show warning; allow continue | Informational; user acknowledges | Low-priority issues |

### Severity Levels

| Level | Definition | Impact | Resolution |
|-------|-----------|--------|-----------|
| **INFO** | Informational; no action needed | None | Auto-accepted |
| **WARNING** | Worth reviewing but likely acceptable | Flagged for QA | User review |
| **CRITICAL** | Requires resolution before approval | Blocks finalization | Must be resolved |

### Fail Action Matrix
```
Severity  │ RANGE │ LOOKUP │ FORMAT │ CROSS │ DUPLICATE │ CUSTOM
──────────┼───────┼────────┼────────┼───────┼───────────┼────────
INFO      │ WARN  │   -    │   -    │  -    │     -     │   -
WARNING   │ FLAG  │  FLAG  │ FLAG   │ FLAG  │   FLAG    │  FLAG
CRITICAL  │ BLOCK │ BLOCK  │ BLOCK  │ BLOCK │   FLAG    │ BLOCK
```

---

## Control Effectiveness

### Example: 3-Way Match Control for A/P
```
OBJECTIVE: Prevent unauthorized or duplicate payments

CONTROL MECHANISM:
  Rule: CUSTOM_AP_01
  Logic: PO + Receipt + Invoice must all match
  
INPUTS REQUIRED:
  1. PO master file (all active POs)
  2. Goods receipt log (when items received)
  3. Vendor invoices (payments due)
  
VALIDATION LOGIC:
  FOR each invoice extracted:
    Match PO number to vendor + amount
    Match receipt date to invoice date (within tolerance)
    Verify amounts align (invoice within ±10% of PO)
    
TEST CASE 1: Perfect match
  Input: PO #123, Vendor ABC, $1000 | Receipt 6/1 | Invoice $1000
  Result: PASS (all three documents match)
  
TEST CASE 2: Missing receipt
  Input: PO #123, Vendor ABC, $1000 | Receipt X | Invoice $1000
  Result: BLOCK (only 2 of 3 match; missing receipt)
  
TEST CASE 3: Amount discrepancy
  Input: PO #123, Vendor ABC, $1000 | Receipt 6/1 | Invoice $1200
  Result: FLAG (amount > 10% of PO; notify auditor)
```

---

## Exception Resolution Workflow

### Exception Lifecycle
```
1. EXCEPTION DETECTED
   ├─ QAEngine identifies exception per rule
   ├─ Records exception details in QA sheet
   └─ Assigns severity (WARNING or CRITICAL)

2. QA REVIEW
   ├─ QA team member reviews exception
   ├─ Examines raw data vs. expected value
   ├─ Researches if needed (look at original document)
   ├─ Decides:
   │  ├─ ACCEPT: Exception is false positive; data OK
   │  ├─ OVERRIDE: Correct value is different; user enters override
   │  └─ REJECT: Data unusable; exclude from output
   └─ Documents decision and justification

3. SIGN-OFF
   ├─ QA lead verifies all decisions
   ├─ Records sign-off (name, date, role)
   └─ Approves finalization

4. APPROVAL
   ├─ Engagement manager reviews QA decisions
   ├─ Records final approval
   └─ OUTPUT sheet generated

5. AUDIT TRAIL
   ├─ AuditLog captures entire history
   ├─ All decisions documented with justification
   └─ Available for external auditor review
```

---

## Rule Customization for Engagements

### Scenario: Special Controls for High-Risk Vendor

```
ENGAGEMENT: AP Audit 2026, Engagement Type: AP
SITUATION: Client has identified one high-risk vendor (ABC Corp)
CUSTOM RULE TO ADD:

Rule ID: CUSTOM_AP_HIGHRISK_01
Vendor: ABC Corp
Logic: All invoices from ABC Corp require additional review
Implementation:
  IF vendor_name = "ABC Corp":
    FLAG exception (always)
    severity = WARNING
    exception_reason = "High-risk vendor; additional review required"
    required_approver = "Manager"
  ENDIF

Stored in: CONFIG sheet, CUSTOM_RULES section
Applied by: QAEngine when processing extractions
Effect: Any ABC Corp invoices automatically flagged for review
```

---

## Auditor Documentation

All QA rules should be documented and retained as evidence of:
- Testing performed during engagement
- Data quality standards applied
- Exceptions identified and resolved
- Control effectiveness demonstrated

---

**Framework Version**: 1.0  
**Last Updated**: June 2026
