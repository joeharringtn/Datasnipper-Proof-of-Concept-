"""
Peek at the text content of a PDF so you can find the right search terms
to use in DataSnipper DS_SEARCH tags.

Usage
-----
  python scripts/peek_pdf.py "path/to/your.pdf"
  python scripts/peek_pdf.py "path/to/your.pdf" --pages 1-3
  python scripts/peek_pdf.py "path/to/your.pdf" --page 2
  python scripts/peek_pdf.py "path/to/your.pdf" --search "sampling"

Output
------
  Prints the extracted text from each page, then lists unique words/phrases
  that are good candidates for DS_SEARCH query strings.

After running
-------------
  Copy the search terms you want into generate_workbook.py's
  make_sample_ap_tags() function (search_keywords field), then rerun
  generate_workbook.py to regenerate the engagement_(ds).xlsx with
  correct tag strings.
"""

import argparse
import re
import sys
from pathlib import Path

try:
    import pdfplumber
except ImportError:
    print("ERROR: pdfplumber not installed. Run: pip install pdfplumber")
    sys.exit(1)


def extract_text(pdf_path: Path, pages: list[int] | None = None) -> dict[int, str]:
    """Return {page_number: text} for the requested pages (1-based)."""
    results: dict[int, str] = {}
    with pdfplumber.open(pdf_path) as pdf:
        total = len(pdf.pages)
        target_pages = pages if pages else list(range(1, total + 1))
        for page_num in target_pages:
            if page_num < 1 or page_num > total:
                print(f"  Warning: page {page_num} out of range (PDF has {total} pages)")
                continue
            page = pdf.pages[page_num - 1]
            text = page.extract_text() or ""
            results[page_num] = text
    return results


def find_label_candidates(text: str) -> list[str]:
    """
    Extract short label-like phrases from extracted text.

    A good DS_SEARCH query is a label that uniquely identifies a field —
    e.g. "Invoice Number", "Total Amount", "Due Date".  These tend to be
    short, title-case phrases followed by a colon or value.
    """
    # Lines that look like labels: short, end in colon, or are title-cased phrases
    candidates = []
    for line in text.splitlines():
        line = line.strip()
        if not line:
            continue
        # Short lines (likely labels, not paragraph text)
        if len(line) <= 60:
            candidates.append(line)
    return candidates


def parse_page_range(spec: str) -> list[int]:
    """Parse '1-3' or '2' into a list of page numbers."""
    pages = []
    for part in spec.split(","):
        part = part.strip()
        if "-" in part:
            start, end = part.split("-", 1)
            pages.extend(range(int(start), int(end) + 1))
        else:
            pages.append(int(part))
    return pages


def main() -> None:
    parser = argparse.ArgumentParser(description="Extract text from a PDF to find DS_SEARCH query terms")
    parser.add_argument("pdf", help="Path to the PDF file")
    parser.add_argument("--pages", help="Page range to extract, e.g. '1-3' or '1,2,4'")
    parser.add_argument("--page",  type=int, help="Single page number to extract")
    parser.add_argument("--search", help="Search for a specific term in the extracted text")
    args = parser.parse_args()

    pdf_path = Path(args.pdf)
    if not pdf_path.exists():
        print(f"ERROR: File not found: {pdf_path}")
        sys.exit(1)

    # Determine which pages to extract
    if args.page:
        pages = [args.page]
    elif args.pages:
        pages = parse_page_range(args.pages)
    else:
        pages = None  # all pages

    print(f"\nPDF: {pdf_path.name}")
    print("=" * 60)

    page_texts = extract_text(pdf_path, pages)

    for page_num, text in sorted(page_texts.items()):
        print(f"\n--- PAGE {page_num} ---")
        if not text.strip():
            print("  (no extractable text on this page — may be a scanned image)")
            continue

        if args.search:
            # Show only lines containing the search term
            term = args.search.lower()
            matches = [ln for ln in text.splitlines() if term in ln.lower()]
            if matches:
                print(f"  Lines containing '{args.search}':")
                for m in matches:
                    print(f"    {m}")
            else:
                print(f"  '{args.search}' not found on this page")
        else:
            print(text)

    # Suggest candidate search terms from page 1
    if not args.search and 1 in page_texts:
        candidates = find_label_candidates(page_texts[1])
        if candidates:
            print("\n" + "=" * 60)
            print("CANDIDATE DS_SEARCH TERMS from page 1 (short lines / labels):")
            print("Use these as search_keywords in generate_workbook.py")
            print("-" * 60)
            for c in candidates[:30]:
                print(f"  {repr(c)}")

    print("\nDone. Copy a search term above into make_sample_ap_tags() > search_keywords,")
    print("then rerun: python scripts/generate_workbook.py --source-pdf <your_pdf>")


if __name__ == "__main__":
    main()
