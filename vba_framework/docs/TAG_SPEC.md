# TAG_SPEC.md — DataSnipper Tag Format Reference

**Source of truth**: [DataSnipper official documentation](https://knowledge.datasnipper.com/how-to-automate-your-datasnipper-workflow-using-the-3rd-party-integration-toolset)

**DataSnipper version required**: 4.0.0 and above

---

## How It Works

When you open a workbook named with the `_(ds)` suffix (e.g. `engagement_(ds).xlsx`),
DataSnipper automatically scans every cell in the workbook for tag strings.
For each tag it finds, DataSnipper:

1. Locates the referenced PDF
2. Performs the extraction (search or coordinate-based)
3. Creates a snip linked to the extracted region
4. Writes the extracted value into the workbook

**No manual ribbon interaction needed** — the naming convention triggers the whole process.

---

## Workbook Naming Requirement

The `.xlsx` filename **must end in `_(ds)`**.

```
engagement(ds).xlsx           ✓ DataSnipper auto-processes on open
client_ap(ds).xlsx            ✓
engagement_(ds).xlsx          ✗ underscore before (ds) — DataSnipper ignores this
engagement.xlsx               ✗ DataSnipper ignores this workbook
```

---

## Tag Types

### DS_SEARCH — Text-Based Extraction

Searches a page of a PDF for a text string.

```
DS_SEARCH[filename|pageNumber|query]
```

| Field        | Description                                          | Example          |
|--------------|------------------------------------------------------|------------------|
| `filename`   | PDF filename or path relative to the workbook        | `invoice.pdf`    |
| `pageNumber` | 1-based page number to search                        | `1`              |
| `query`      | Exact text string DataSnipper will search for        | `Invoice Number` |

**Examples**
```
DS_SEARCH[invoice.pdf|1|Invoice Number]
DS_SEARCH[invoice.pdf|2|Total Amount]
DS_SEARCH[PDFS\statement.pdf|1|Account Balance]
```

**When to use DS_SEARCH**
- Document structure varies between instances
- Text labels (anchors) are consistent and unique on the page
- You don't know the exact pixel position of the value

---

### DS_COORDS — Coordinate-Based Extraction

Extracts text from a fixed bounding box on a specific page.

```
DS_COORDS[filename|pageNumber|x1|y1|x2|y2]
```

| Field        | Description                                                | Example |
|--------------|------------------------------------------------------------|---------|
| `filename`   | PDF filename or path relative to the workbook              | `invoice.pdf` |
| `pageNumber` | 1-based page number                                        | `1`     |
| `x1`         | Left edge of the extraction box (pixels from left margin)  | `302`   |
| `y1`         | Top edge of the extraction box (pixels from top margin)    | `81`    |
| `x2`         | Right edge of the extraction box                           | `460`   |
| `y2`         | Bottom edge of the extraction box                          | `101`   |

Note: `x2 = x1 + width`, `y2 = y1 + height`.  Use DataSnipper's hover tool
to read pixel coordinates directly from the PDF.

**Examples**
```
DS_COORDS[invoice.pdf|1|302|81|460|101]
DS_COORDS[2021_regulations.pdf|1|302|81|360|90]
DS_COORDS[statement.pdf|3|50|200|200|220]
```

**When to use DS_COORDS**
- Document layout is fixed and consistent across all instances
- Values appear at the same pixel position every time
- Text anchors are absent, ambiguous, or duplicated

---

## File Path Rules

DataSnipper resolves the `filename` field relative to the **saved location of
the workbook on disk**.

| Path style          | Example                                      | When to use                  |
|---------------------|----------------------------------------------|------------------------------|
| Filename only       | `invoice.pdf`                                | PDF in same folder as xlsx   |
| Relative subfolder  | `PDFS\invoice.pdf`                           | PDF in a subfolder           |
| Parent navigation   | `..\invoices\invoice.pdf`                    | PDF one level up             |
| Absolute path       | `C:\Users\name\Documents\invoice.pdf`        | PDF anywhere on the machine  |

**Recommended POC setup**: keep the PDF in the same folder as the `_(ds).xlsx`
workbook and use just the filename.

---

## Tag Placement in the Workbook

DataSnipper scans **all cells** in the workbook.  The framework writes tag
strings into the **SourceTag column** of the TAG_ENGINE sheet.  Any cell
containing a valid `DS_SEARCH[...]` or `DS_COORDS[...]` string will be
processed.

---

## Python TagBuilder Output

The `TagBuilder` class in `audit_framework/tag_builder.py` generates these
strings from `TagDefinition` model objects:

```python
# DS_SEARCH
TagDefinition(
    tag_id="AP_InvoiceNum_S",
    extraction_method=ExtractionMethod.DS_SEARCH,
    source_document="invoice.pdf",
    search_page=1,
    search_keywords="Invoice Number",
    ...
)
# → DS_SEARCH[invoice.pdf|1|Invoice Number]

# DS_COORDS (x2 = coord_x + coord_width, y2 = coord_y + coord_height)
TagDefinition(
    tag_id="AP_InvoiceDate_C",
    extraction_method=ExtractionMethod.DS_COORDS,
    source_document="invoice.pdf",
    coord_page=1,
    coord_x=302, coord_y=81, coord_width=158, coord_height=20,
    ...
)
# → DS_COORDS[invoice.pdf|1|302|81|460|101]
```

---

## Tag Naming Convention (TagID field)

```
{engagement_type}_{field}_{method}

AP_InvoiceNum_S       A/P engagement, invoice number, DS_SEARCH
AP_InvoiceDate_C      A/P engagement, invoice date, DS_COORDS
AR_CustomerName_S     A/R engagement, customer name, DS_SEARCH
CASH_DepositAmt_S     Cash engagement, deposit amount, DS_SEARCH
```

---

## Common Mistakes

```
engagement.xlsx                   ✗  Missing _(ds) suffix — DataSnipper ignores it
DS_SEARCH(invoice.pdf|1|Total)    ✗  Parentheses instead of square brackets
DS_SEARCH[invoice.pdf|0|Total]    ✗  Page 0 is invalid; pages are 1-based
DS_SEARCH[invoice.pdf|Total]      ✗  Missing page number (only 2 pipe segments)
DS_COORDS[invoice.pdf|1|302|81]   ✗  Missing x2 and y2 (need all 6 fields)
```

---

**Framework Version**: 1.0
**Last Updated**: June 2026
**Spec source**: DataSnipper knowledge base (verified against v4.0 documentation)
