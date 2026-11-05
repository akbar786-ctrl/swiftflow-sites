"""
Adds the generated calculation pages to the v1 template WITHOUT re-saving the
template through a spreadsheet library, so its pages keep everything exactly
as designed (logos, banner shapes, cross-sheet drop-downs, conditional
formatting, hatching).

Method: copy the template package byte for byte, then add each calculation
worksheet part, appending its styles to the template's styles.xml (indices
remapped) and converting shared strings to inline strings. Excel is told to
recalculate everything when the file is opened (fullCalcOnLoad).
"""
import re
import sys
import zipfile
from xml.sax.saxutils import escape

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


def _section(xml, tag):
    m = re.search(rf"<{tag}\b[^>]*?(?:/>|>.*?</{tag}>)", xml, re.S)
    return m


def _items(block, item):
    return re.findall(rf"<{item}\b[^>]*?(?:/>|>.*?</{item}>)", block, re.S) if block else []


def _count(xml, tag):
    m = _section(xml, tag)
    return len(_items(m.group(0), tag[:-1] if tag.endswith("s") else tag)) if m else 0


def merge_styles(target, source):
    """Append source fonts/fills/borders/numFmts/cellXfs to target. Returns (new target xml, xf offset)."""
    out = target
    offsets = {}
    for tag, item in (("fonts", "font"), ("fills", "fill"), ("borders", "border")):
        tm = _section(out, tag)
        sm = _section(source, tag)
        t_items, s_items = _items(tm.group(0), item), _items(sm.group(0), item)
        offsets[item] = len(t_items)
        block = tm.group(0)
        block = re.sub(r'count="\d+"', f'count="{len(t_items) + len(s_items)}"', block, count=1)
        block = block.replace(f"</{tag}>", "".join(s_items) + f"</{tag}>")
        out = out[:tm.start()] + block + out[tm.end():]
    # number formats
    s_nf = _section(source, "numFmts")
    nf_map = {}
    if s_nf:
        t_nf = _section(out, "numFmts")
        t_ids = [int(x) for x in re.findall(r'numFmtId="(\d+)"', t_nf.group(0))] if t_nf else []
        nxt = max([163] + t_ids) + 1
        new = []
        for it in _items(s_nf.group(0), "numFmt"):
            old = int(re.search(r'numFmtId="(\d+)"', it).group(1))
            nf_map[old] = nxt
            new.append(re.sub(r'numFmtId="\d+"', f'numFmtId="{nxt}"', it))
            nxt += 1
        if t_nf:
            block = t_nf.group(0)
            total = len(_items(block, "numFmt")) + len(new)
            block = re.sub(r'count="\d+"', f'count="{total}"', block, count=1).replace("</numFmts>", "".join(new) + "</numFmts>")
            out = out[:t_nf.start()] + block + out[t_nf.end():]
        else:
            pos = out.index("<fonts")
            out = out[:pos] + f'<numFmts count="{len(new)}">' + "".join(new) + "</numFmts>" + out[pos:]
    # cell formats
    tx = _section(out, "cellXfs")
    sx = _section(source, "cellXfs")
    t_xfs, s_xfs = _items(tx.group(0), "xf"), _items(sx.group(0), "xf")
    xf_off = len(t_xfs)
    new = []
    for xf in s_xfs:
        def rep(attr, off):
            nonlocal xf
            xf = re.sub(rf'{attr}="(\d+)"', lambda m: f'{attr}="{int(m.group(1)) + off}"', xf, count=1)
        rep("fontId", offsets["font"])
        rep("fillId", offsets["fill"])
        rep("borderId", offsets["border"])
        xf = re.sub(r'numFmtId="(\d+)"', lambda m: f'numFmtId="{nf_map.get(int(m.group(1)), int(m.group(1)))}"', xf, count=1)
        xf = re.sub(r'xfId="\d+"', 'xfId="0"', xf)
        new.append(xf)
    block = tx.group(0)
    block = re.sub(r'count="\d+"', f'count="{len(t_xfs) + len(new)}"', block, count=1).replace("</cellXfs>", "".join(new) + "</cellXfs>")
    out = out[:tx.start()] + block + out[tx.end():]
    return out, xf_off


def shared_strings(xml):
    out = []
    for si in re.findall(r"<si>(.*?)</si>", xml, re.S):
        texts = re.findall(r"<t(?:\s[^>]*)?>(.*?)</t>", si, re.S)
        out.append("".join(texts))
    return out


def convert_sheet(xml, sst, xf_off):
    def cell(m):
        tag, body = m.group(1), m.group(2)
        tag = re.sub(r'\bs="(\d+)"', lambda x: f's="{int(x.group(1)) + xf_off}"', tag)
        if 't="s"' in tag:
            idx = int(re.search(r"<v>(\d+)</v>", body).group(1))
            tag = tag.replace('t="s"', 't="inlineStr"')
            body = f'<is><t xml:space="preserve">{sst[idx]}</t></is>'
        return f"<c{tag}>{body}</c>"
    xml = re.sub(r"<c(\b[^>]*?)>(.*?)</c>", cell, xml, flags=re.S)
    xml = re.sub(r"<c(\b[^>]*?)/>", lambda m: "<c" + re.sub(r'\bs="(\d+)"', lambda x: f's="{int(x.group(1)) + xf_off}"', m.group(1)) + "/>", xml)
    xml = re.sub(r"<row(\b[^>]*?)>", lambda m: "<row" + re.sub(r'\bs="(\d+)"', lambda x: f's="{int(x.group(1)) + xf_off}"', m.group(1)) + ">", xml)
    xml = re.sub(r"<col(\b[^>]*?)/>", lambda m: "<col" + re.sub(r'\bstyle="(\d+)"', lambda x: f'style="{int(x.group(1)) + xf_off}"', m.group(1)) + "/>", xml)
    xml = xml.replace(' tabSelected="1"', "")
    return xml


def inject(template, calc, out):
    tz, cz = zipfile.ZipFile(template), zipfile.ZipFile(calc)
    files = {n: tz.read(n) for n in tz.namelist()}
    cfiles = {n: cz.read(n) for n in cz.namelist()}
    styles, xf_off = merge_styles(files["xl/styles.xml"].decode("utf-8"), cfiles["xl/styles.xml"].decode("utf-8"))
    files["xl/styles.xml"] = styles.encode("utf-8")
    sst = shared_strings(cfiles.get("xl/sharedStrings.xml", b"<sst/>").decode("utf-8"))
    sst = [s for s in sst]
    cwb = cfiles["xl/workbook.xml"].decode("utf-8")
    crels = cfiles["xl/_rels/workbook.xml.rels"].decode("utf-8")
    csheets = []
    for tag in re.findall(r"<sheet\b[^>]*/>", cwb):
        attr = dict(re.findall(r'([\w:]+)="([^"]*)"', tag))
        csheets.append((attr["name"], attr.get("state", "visible"), attr["r:id"]))
    wb = files["xl/workbook.xml"].decode("utf-8")
    rels = files["xl/_rels/workbook.xml.rels"].decode("utf-8")
    ct = files["[Content_Types].xml"].decode("utf-8")
    existing = [int(x) for x in re.findall(r"worksheets/sheet(\d+)\.xml", " ".join(files))]
    next_part = max(existing) + 1
    next_id = max(int(x) for x in re.findall(r'sheetId="(\d+)"', wb)) + 1
    next_rid = max(int(x) for x in re.findall(r'Id="rId(\d+)"', rels)) + 1
    new_sheet_tags = []
    for name, state, rid in csheets:
        target = re.search(rf'Id="{rid}"[^>]*Target="([^"]+)"', crels) or re.search(rf'Target="([^"]+)"[^>]*Id="{rid}"', crels)
        part = "xl/" + target.group(1).lstrip("/").replace("xl/", "")
        xml = convert_sheet(cfiles[part].decode("utf-8"), [escape(s) for s in sst], xf_off)
        new_part = f"xl/worksheets/sheet{next_part}.xml"
        files[new_part] = xml.encode("utf-8")
        rels = rels.replace("</Relationships>",
                            f'<Relationship Id="rId{next_rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
                            f'Target="worksheets/sheet{next_part}.xml"/></Relationships>')
        ct = ct.replace("</Types>", f'<Override PartName="/{new_part}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>')
        st = ' state="hidden"' if state == "hidden" else ""
        new_sheet_tags.append(f'<sheet name="{name}" sheetId="{next_id}"{st} r:id="rId{next_rid}"/>')   # name already XML-escaped
        next_part, next_id, next_rid = next_part + 1, next_id + 1, next_rid + 1
    wb = wb.replace("</sheets>", "".join(new_sheet_tags) + "</sheets>")
    if "<calcPr" in wb:
        wb = re.sub(r"<calcPr([^>]*?)/>", lambda m: "<calcPr" + re.sub(r'\sfullCalcOnLoad="\d"', "", m.group(1)) + ' fullCalcOnLoad="1"/>', wb)
    else:
        wb = wb.replace("</workbook>", '<calcPr fullCalcOnLoad="1"/></workbook>')
    files["xl/workbook.xml"] = wb.encode("utf-8")
    files["xl/_rels/workbook.xml.rels"] = rels.encode("utf-8")
    files["[Content_Types].xml"] = ct.encode("utf-8")
    # calcChain lists only existing formula cells; drop it so Excel rebuilds it cleanly
    if "xl/calcChain.xml" in files:
        del files["xl/calcChain.xml"]
        files["[Content_Types].xml"] = re.sub(rb'<Override PartName="/xl/calcChain.xml"[^>]*/>', b"", files["[Content_Types].xml"])
        files["xl/_rels/workbook.xml.rels"] = re.sub(rb'<Relationship [^>]*calcChain[^>]*/>', b"", files["xl/_rels/workbook.xml.rels"])
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for n in tz.namelist():
            if n in files:
                z.writestr(tz.getinfo(n), files[n])
        for n, data in files.items():
            if n not in tz.namelist():
                z.writestr(n, data)
    return out


if __name__ == "__main__":
    print(inject(*sys.argv[1:4]))
