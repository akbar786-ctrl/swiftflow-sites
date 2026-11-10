"""
Builds the template sent to forces: the BLANK v1 template with
  * Match Day Resource rows 3/4 swapped (risk level on top) and the Super High block (AN:AY),
    totals in AZ:BK (same transform as the reviewed lift-and-shift build);
  * two hidden named ranges Power BI reads, so it always gets the same cells whatever Excel
    considers the "used range":  MD_Data = 'Match Day Resource'!$A$1:$AY$345
                                 WT_Data = 'Wider Tournament Resource'!$A$1:$O$105
No calculation pages - forces fill Match Day Resource and Wider Tournament Resource only.

Usage: python3 make_forces_template.py [out.xlsx]
"""
import os
import re
import sys
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
V1 = os.path.join(HERE, "..", "v1_lift_shift")
sys.path.insert(0, V1)
import inject  # noqa: E402
import swap_headers  # noqa: E402

SOURCE = os.path.join(V1, "source", "BLANK_Euro_2028_Resource_NEW.xlsx")
OUT = os.path.join(HERE, "output", "EURO_2028_Resource_Collection_Template.xlsx")
NAMES = {"MD_Data": "'Match Day Resource'!$A$1:$AY$345", "WT_Data": "'Wider Tournament Resource'!$A$1:$O$105"}


def workbook_xml(xml):
    xml = swap_headers.workbook(xml)
    names = "".join(f'<definedName name="{n}" hidden="1">{ref}</definedName>' for n, ref in NAMES.items())
    names = names.replace("'", "&apos;")
    if "<definedNames>" in xml:
        xml = xml.replace("<definedNames>", "<definedNames>" + names, 1)
    else:
        xml = xml.replace("</sheets>", "</sheets><definedNames>" + names + "</definedNames>", 1)
    # recalculate the (rewritten) total formulas when opened
    if "<calcPr" in xml:
        xml = re.sub(r"<calcPr([^>]*?)/>", lambda m: "<calcPr" + re.sub(r'\sfullCalcOnLoad="\d"', "", m.group(1))
                     + ' fullCalcOnLoad="1"/>', xml)
    else:
        xml = xml.replace("</workbook>", '<calcPr fullCalcOnLoad="1"/></workbook>')
    return xml


def build(out):
    tz = zipfile.ZipFile(SOURCE)
    files = {n: tz.read(n) for n in tz.namelist()}
    part = inject.sheet_part(files, "Match Day Resource")
    files[part] = swap_headers.transform(files[part].decode("utf-8")).encode("utf-8")
    files["xl/workbook.xml"] = workbook_xml(files["xl/workbook.xml"].decode("utf-8")).encode("utf-8")
    if "xl/calcChain.xml" in files:
        del files["xl/calcChain.xml"]
        files["[Content_Types].xml"] = re.sub(rb'<Override PartName="/xl/calcChain.xml"[^>]*/>', b"",
                                              files["[Content_Types].xml"])
        files["xl/_rels/workbook.xml.rels"] = re.sub(rb'<Relationship [^>]*calcChain[^>]*/>', b"",
                                                     files["xl/_rels/workbook.xml.rels"])
    os.makedirs(os.path.dirname(out), exist_ok=True)
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for n in tz.namelist():
            if n in files:
                z.writestr(tz.getinfo(n), files[n])
    return out


if __name__ == "__main__":
    print(build(os.path.abspath(sys.argv[1] if len(sys.argv) > 1 else OUT)))
