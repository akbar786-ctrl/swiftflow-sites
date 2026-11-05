"""
Builds the v1 collection workbook with the original calculation pages:
  source/BLANK_Euro_2028_Resource_NEW.xlsx  (your template, untouched)
  + Demand Sheet, Cost Sheet, MD Demand Template, MD Cost Template, Fleet,
    Cost Calculator, Wider Tournament Template, Command Team & Planning,
    Other Resources, Baseline Costs, GDP Deflators, WT Cost Sheet, WT Demand Sheet
Usage: python3 build.py [output.xlsx]
"""
import os
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import build_calc_pages  # noqa: E402
import inject  # noqa: E402
import swap_headers  # noqa: E402

SOURCE = os.path.join(HERE, "source", "BLANK_Euro_2028_Resource_NEW.xlsx")
OUT = os.path.join(HERE, "..", "output", "Euro_2028_Resource_Collection_v1_with_Calculations.xlsx")

if __name__ == "__main__":
    out = os.path.abspath(sys.argv[1] if len(sys.argv) > 1 else OUT)
    with tempfile.TemporaryDirectory() as tmp:
        pages = build_calc_pages.build(SOURCE, os.path.join(tmp, "pages.xlsx"))
        print(inject.inject(SOURCE, pages, out, {"Match Day Resource": swap_headers.transform},
                                {"xl/workbook.xml": swap_headers.workbook}))
