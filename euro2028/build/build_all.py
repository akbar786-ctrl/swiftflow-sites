"""Regenerate every EURO 2028 output from the Python sources."""
import os
import sys

sys.path.insert(0, os.path.dirname(__file__))
import build_calc_workbook  # noqa: E402
import build_reference  # noqa: E402
import build_sample  # noqa: E402
import build_template  # noqa: E402
import catalogue as C  # noqa: E402

OUT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "output"))

if __name__ == "__main__":
    v = C.TEMPLATE_VERSION
    print(build_template.build(os.path.join(OUT, f"E28_Collection_Template_v{v}.xlsx")))
    print(build_calc_workbook.build(os.path.join(OUT, f"E28_Backup_Calculation_Workbook_v{v}.xlsx")))
    print(build_reference.build(os.path.join(OUT, "E28_Reference_Tables.xlsx")))
    for p in build_sample.build(os.path.join(OUT, "samples")):
        print(p)
