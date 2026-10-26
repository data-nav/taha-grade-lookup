"""
One-off: build the slim templates.xlsx from a full "Web Host" export.

The render workbook only needs the Marksheet/Transcript sheets, not the data
sheets. Stripping them takes the per-request openpyxl load from ~38s to ~0.5s,
which is where most of the old 10s wait went.

Re-run this whenever a template sheet's layout changes.

    python3 build_templates.py "Web Host.xlsx" templates.xlsx
"""

import sys
import openpyxl


def build(src, dest):
    wb = openpyxl.load_workbook(src)
    keep = [s for s in wb.sheetnames
            if s.endswith("Marksheet") or s.endswith("Transcript") or s.endswith("Transcript Pass")]
    for s in list(wb.sheetnames):
        if s not in keep:
            del wb[s]
    wb.save(dest)
    print("Kept %d template sheets -> %s" % (len(keep), dest))
    for s in keep:
        print("  -", s)


if __name__ == "__main__":
    src = sys.argv[1] if len(sys.argv) > 1 else "Web Host.xlsx"
    dest = sys.argv[2] if len(sys.argv) > 2 else "templates.xlsx"
    build(src, dest)
