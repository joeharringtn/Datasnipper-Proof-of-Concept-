"""
Generate the simplest possible DataSnipper test workbook.

One sheet. One cell. One tag. Nothing else.
If this doesn't snip, the issue is with DataSnipper configuration,
not with our workbook structure.
"""

import sys
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook

def main():
    pdf_filename = sys.argv[1] if len(sys.argv) > 1 else "obs_sampling.pdf"
    search_term  = sys.argv[2] if len(sys.argv) > 2 else "Joe"
    page         = sys.argv[3] if len(sys.argv) > 3 else "1"

    tag = f"DS_SEARCH[{pdf_filename}|{page}|{search_term}]"

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    repo_root = Path(__file__).parent.parent
    out_path  = repo_root / "output" / f"minimal_test_{timestamp}(ds).xlsx"

    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws["A1"] = tag

    wb.save(out_path)

    print(f"Tag    : {tag}")
    print(f"Saved  : {out_path.resolve()}")
    print()
    print("Instructions:")
    print(f"  1. Put {pdf_filename} in the same folder as the xlsx")
    print(f"  2. Open the xlsx in Excel with DataSnipper installed")
    print(f"  3. DataSnipper should snip cell A1 on Sheet1")

if __name__ == "__main__":
    main()
