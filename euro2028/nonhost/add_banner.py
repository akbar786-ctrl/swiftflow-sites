"""
Copies the national team's banner (drawing + images) from their Non-host_collections.xlsx onto
every sheet of the generated Non-Host form, unchanged (the banner is absolutely positioned).

Usage: python3 add_banner.py <team Non-host_collections.xlsx> <generated form.xlsx> <out.xlsx>
"""
import re
import sys
import zipfile

DRAWING_T = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"
ORDER_AFTER = ["legacyDrawing", "legacyDrawingHF", "picture", "oleObjects", "controls", "webPublishItems",
               "tableParts", "extLst"]


def main(team, form, out):
    tz, fz = zipfile.ZipFile(team), zipfile.ZipFile(form)
    files = {n: fz.read(n) for n in fz.namelist()}
    tfiles = {n: tz.read(n) for n in tz.namelist()}
    # banner drawing on the team's data sheet
    trels = tfiles["xl/worksheets/_rels/sheet2.xml.rels"].decode()
    dtarget = re.search(rf'Type="{DRAWING_T}" Target="\.\./drawings/([^"]+)"', trels) or \
        re.search(rf'Target="\.\./drawings/([^"]+)"[^>]*Type="{DRAWING_T}"', trels)
    dname = dtarget.group(1)
    drawing = tfiles[f"xl/drawings/{dname}"]
    drels = tfiles[f"xl/drawings/_rels/{dname}.rels"]
    for n, data in tfiles.items():
        if n.startswith("xl/media/"):
            files[n] = data
    ct = files["[Content_Types].xml"].decode()
    for ext, mime in (("png", "image/png"), ("jpg", "image/jpeg"), ("jpeg", "image/jpeg")):
        if f'Extension="{ext}"' not in ct:
            ct = ct.replace("<Default ", f'<Default Extension="{ext}" ContentType="{mime}"/><Default ', 1)
    sheets = sorted(n for n in files if re.match(r"xl/worksheets/sheet\d+\.xml$", n))[:2]   # Guidance + data tab
    wb = files["xl/workbook.xml"].decode()
    for i, part in enumerate(sheets, 1):
        dp = f"xl/drawings/banner{i}.xml"
        files[dp] = drawing
        files[f"xl/drawings/_rels/banner{i}.xml.rels"] = drels
        ct = ct.replace("</Types>", f'<Override PartName="/{dp}" ContentType="application/vnd.openxmlformats-'
                                    f'officedocument.drawing+xml"/></Types>')
        rels_name = part.replace("worksheets/", "worksheets/_rels/") + ".rels"
        rels = files.get(rels_name, b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships '
                                    b'xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                                    b'</Relationships>').decode()
        rid = "rIdBanner1"
        rels = rels.replace("</Relationships>", f'<Relationship Id="{rid}" Type="{DRAWING_T}" '
                                                f'Target="../drawings/banner{i}.xml"/></Relationships>')
        files[rels_name] = rels.encode()
        xml = files[part].decode()
        if "xmlns:r=" not in xml[:500]:
            xml = xml.replace("<worksheet ", '<worksheet xmlns:r="http://schemas.openxmlformats.org/'
                                            'officeDocument/2006/relationships" ', 1)
        tag = f'<drawing r:id="{rid}"/>'
        for name in ORDER_AFTER:
            m = re.search(rf"<{name}\b", xml)
            if m:
                xml = xml[:m.start()] + tag + xml[m.start():]
                break
        else:
            xml = xml.replace("</worksheet>", tag + "</worksheet>")
        files[part] = xml.encode()
    files["[Content_Types].xml"] = ct.encode()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for n, data in files.items():
            z.writestr(n, data)
    return out


if __name__ == "__main__":
    print(main(*sys.argv[1:4]))
