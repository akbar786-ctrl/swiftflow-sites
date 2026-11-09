"""
Writes the Excel VBA macro that turns the user's BLANK v1 template into the
full collection workbook inside Excel (same result as build.py):

  1. Match Day Resource: risk level on row 3 (Low | Medium | High | Super High),
     day and measure on row 4, Super High block AN:AY, totals AZ:BK.
  2. Adds every calculation page with the original workbook's logic.

The macro holds the row layout as data lines generated here from the same
Python definitions build.py uses, and builds every formula with loops that
mirror build_calc_pages.py, so the two routes give the same formulas.

Usage: python3 make_vba.py [out.bas] [--test]
  --test adds a no-prompt entry point used by test_vba_libreoffice.py.
"""
import datetime as dt
import json
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import build_calc_pages as B  # noqa: E402
import layout_v1 as LY  # noqa: E402

SOURCE = os.path.join(HERE, "source", "BLANK_Euro_2028_Resource_NEW.xlsx")
OUT = os.path.join(HERE, "..", "output", "Euro_2028_Build_Macro.bas")
CHUNK = 150                     # data lines per VBA procedure


def enc(s):
    """Encode text for a VBA string literal: ASCII only, no line breaks, quotes doubled."""
    s = "" if s is None else str(s)
    assert "|" not in s and "~" not in s, s
    out = []
    for ch in s:
        if ch == "\n":
            out.append("~n~")
        elif ch == "\r":
            continue
        elif ord(ch) > 126 or ord(ch) < 32:
            out.append(f"~u{ord(ch):X}~")
        else:
            out.append(ch)
    return "".join(out).replace('"', '""')


def rgb(hexstr):
    h = hexstr[-6:]
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    return r + g * 256 + b * 65536


# ------------------------------------------------------------------ data lines
def md_data(t):
    out = []
    for kind, p in B.md_lines(t):
        if kind == "section":
            out.append("S|" + enc(p))
        else:
            out.append("|".join(["L", str(p["sr"]), str(p["cfg"] or ""), p["group"], p["kind"], str(p["cc"]),
                                 enc(p["b"]), enc(p["c"])]))
    return out


def wt_data(t):
    out = []
    for kind, p in B.wt_lines(t):
        if kind == "section":
            out.append("S|" + enc(p))
        else:
            out.append("|".join(["L", str(p["sr"]), str(p["cfg"] or ""), p["kind"], str(p["cc"]),
                                 enc(p["b"]), enc(p["c"])]))
    return out


def cc_data(t):
    """Cost Calculator rows, mirroring build_calc_pages.cost_calculator/put_row."""
    out = []

    def row(sheet, r, out_r, sumif):
        cfg_is = sheet == LY.MD_SHEET and any(r in (c, c + 1, c + 2) for c in LY.MD_IN_STADIUM_CONFIG)
        survey = B.MDS if sheet == LY.MD_SHEET else B.WTS
        label = t.label(sheet, r) or ""
        cform = f'=IF({survey}!$C${r}="","",{survey}!$C${r})'
        counts = LY.MD_COUNTS if sheet == LY.MD_SHEET else LY.WT_COUNTS
        if r in counts:
            mode, rate = ("PKG", "") if counts[r][0] == "MA" else ("SKIP", "")
        else:
            rate = B.rate_expr(t, sheet, r, cfg_is)
            mode, rate = ("RATE", rate) if rate else ("ZERO", "")
        out.append("|".join(["R", "MD" if sheet == LY.MD_SHEET else "WT", str(r), str(out_r), enc(label),
                             enc(cform), mode, enc(rate), "1" if sumif else "0"]))

    for r in sorted(LY.MD_CONFIG):
        for i in range(3):
            row(LY.MD_SHEET, r + i, r + i, False)
    for r in sorted(LY.MD_COUNTS) + LY.md_officer_rows():
        row(LY.MD_SHEET, r, r, r not in LY.MD_COUNTS or LY.MD_COUNTS[r][0] == "MA")
    for r in sorted(LY.WT_CONFIG):
        for i in range(3):
            row(LY.WT_SHEET, r + i, B.WT_CC_OFFSET + r + i, False)
    for r in sorted(LY.WT_COUNTS) + LY.wt_officer_rows():
        row(LY.WT_SHEET, r, B.WT_CC_OFFSET + r, False)
    return out


def static_data():
    data = json.load(open(os.path.join(HERE, "legacy_static_pages.json")))
    out = {}
    for name in ["Command Team & Planning", "Other Resources", "Baseline Costs", "GDP Deflators"]:
        d, lines = data[name], []
        for c in d["cells"]:
            v = c["v"]
            if isinstance(v, dict) and "date" in v:
                typ, val = "d", v["date"][:10]
            elif isinstance(v, bool):
                typ, val = "n", "1" if v else "0"
            elif isinstance(v, (int, float)):
                typ, val = "n", repr(v)
            elif isinstance(v, str) and v.startswith("="):
                typ, val = "f", v.replace("_xlfn.", "")
            elif v is None:
                typ, val = "e", ""
            else:
                typ, val = "s", v
            fc = str(rgb(c["fc"])) if c["fc"] and not str(c["fc"]).startswith("theme") else ""
            if c["fill"] and not str(c["fill"]).startswith("theme"):
                fill = str(rgb(c["fill"]))
            elif c["fill"]:
                fill = str(rgb(B.LIGHT))
            else:
                fill = ""
            lines.append("|".join(["C", c["ref"], typ, enc(val), enc(c["nf"] or "General"), "1" if c["b"] else "0",
                                   "1" if c["i"] else "0", str(c["sz"] or 11), fc, fill, c["h"] or "",
                                   "1" if c["wrap"] else "0", "1" if c["bd"] else "0"]))
        for col, w in d["widths"].items():
            lines.append(f"W|{col}|{w}")
        for m in d["merges"]:
            lines.append(f"M|{m}")
        out[name] = lines
    return out


# ------------------------------------------------------------------ VBA text
def add_line(fn, payload):
    """One VBA statement adding payload (split if the line would be too long)."""
    if len(payload) <= 700:
        return [f'    {fn} "{payload}"']
    parts, out = [payload[i:i + 600] for i in range(0, len(payload), 600)], []
    # never split inside a doubled quote
    fixed, carry = [], ""
    for p in parts:
        p = carry + p
        carry = ""
        while p.endswith('"') and (len(p) - len(p.rstrip('"'))) % 2 == 1:
            carry = p[-1] + carry
            p = p[:-1]
        fixed.append(p)
    if carry:
        fixed[-1] += carry
    out.append(f'    s = "{fixed[0]}"')
    for p in fixed[1:]:
        out.append(f'    s = s & "{p}"')
    out.append(f"    {fn} s")
    return out


def data_procs(prefix, fn, lines):
    procs, names = [], []
    for n, i in enumerate(range(0, len(lines), CHUNK)):
        name = f"{prefix}{n + 1}"
        names.append(name)
        body = [f"Private Sub {name}()", "    Dim s As String"]
        for ln in lines[i:i + CHUNK]:
            body += add_line(fn, ln)
        body.append("End Sub")
        procs.append("\n".join(body))
    return procs, names


HEAD = r'''' =====================================================================================
'  EURO 2028 - National Policing Resource Collection
'  Builds the full collection workbook from the BLANK v1 template, inside Excel.
'
'  HOW TO USE (desktop Excel for Windows or Mac)
'   1. Open BLANK_Euro_2028_Resource_NEW.xlsx (the blank v1 template).
'   2. Press Alt+F11 (Mac: Tools > Macro > Visual Basic Editor).
'   3. Insert > Module, then paste ALL of this text into the empty module.
'      (Or: File > Import File... and choose this .bas file.)
'   4. Close the editor, press Alt+F8, choose BuildEuro2028Workbook, click Run.
'      It takes about one to three minutes.
'   5. File > Save As > "Excel Workbook (*.xlsx)". Saving as .xlsx removes this
'      macro, so the file you send to forces contains no code.
'
'  WHAT IT DOES
'   - Match Day Resource: row 3 = Low | Medium | High | Super High (risk level),
'     row 4 = Match Day -2 / -1 / Match Day / +1 x Number | Normal ToD | Extra Hours.
'     Super High is columns AN:AY; Automatic Total Figures move to AZ:BK.
'   - Adds the calculation pages of the original collection workbook, with the
'     original logic: Demand Sheet, Cost Sheet, MD Demand Template, MD Cost
'     Template, Fleet, Cost Calculator, Wider Tournament Template, Command Team &
'     Planning, Other Resources, Baseline Costs, GDP Deflators, WT Cost Sheet,
'     WT Demand Sheet. Working pages are hidden as in the original.
'   - Super High is new (not in the original) and uses the same logic and rates.
'
'  Run it once, on the blank template. It stops if the pages already exist.
'  Classification: OFFICIAL
' =====================================================================================
Option Explicit

Private Const NR As Long = 4
Private Const CAL_DAYS As Long = 34
Private Const MDS As String = "'Match Day Resource'"
Private Const WTS As String = "'Wider Tournament Resource'"
Private Const BCS As String = "'Baseline Costs'"
Private Const WT_CC_OFFSET As Long = 345
Private Const C_NAVY As Long = %NAVY%
Private Const C_BLUE As Long = %BLUE%
Private Const C_LIGHT As Long = %LIGHT%
Private Const C_BORDER As Long = %BORDER%
Private Const C_INPUT As Long = %INPUT%
Private Const C_WHITE As Long = 16777215
Private Const X_TORIGHT As Long = -4161
Private Const X_PASTEFORMATS As Long = -4122
Private Const X_CENTER As Long = -4108

Private MDL As Collection
Private WTL As Collection
Private CCL As Collection
Private STL As Collection
Private CurStatic As String
Private Quiet As Boolean
Private StepName As String

Private Sub AddMD(ByVal s As String)
    MDL.Add Dec(s)
End Sub

Private Sub AddWT(ByVal s As String)
    WTL.Add Dec(s)
End Sub

Private Sub AddCC(ByVal s As String)
    CCL.Add Dec(s)
End Sub

Private Sub AddST(ByVal s As String)
    STL.Add Dec(s)
End Sub

Private Function Dec(ByVal s As String) As String
    Dim p As Long, q As Long
    s = Replace(s, "~n~", vbLf)
    p = InStr(1, s, "~u")
    Do While p > 0
        q = InStr(p + 2, s, "~")
        s = Left$(s, p - 1) & ChrW(CLng("&H" & Mid$(s, p + 2, q - p - 2))) & Mid$(s, q + 1)
        p = InStr(1, s, "~u")
    Loop
    Dec = s
End Function

Private Function ColL(ByVal n As Long) As String
    Dim s As String
    Do While n > 0
        s = Chr$(65 + ((n - 1) Mod 26)) & s
        n = (n - 1) \ 26
    Loop
    ColL = s
End Function

Private Function DayIdx(ByVal d As String) As Long
    Select Case d
        Case "MD-2": DayIdx = 0
        Case "MD-1": DayIdx = 1
        Case "MD": DayIdx = 2
        Case Else: DayIdx = 3
    End Select
End Function

Private Function MdCol(ByVal d As String, ByVal s As Long, ByVal m As Long) As Long
    MdCol = 4 + s * 12 + DayIdx(d) * 3 + m
End Function

Private Function DemandCol(ByVal k As Long, ByVal s As Long) As Long
    DemandCol = 4 + (NR + 1) * k + s
End Function

Private Function CostCol(ByVal k As Long, ByVal s As Long) As Long
    CostCol = 4 + NR * k + s
End Function

Private Function SurveyNum(ByVal d As String, ByVal s As Long, ByVal r As Long) As String
    SurveyNum = MDS & "!$" & ColL(MdCol(d, s, 0)) & "$" & r
End Function

Private Function CCCol(ByVal d As String, ByVal s As Long) As String
    CCCol = ColL(4 + NR * DayIdx(d) + s)
End Function

Private Function Scen(ByVal s As Long) As String
    Scen = Choose(s + 1, "Low", "Medium", "High", "Super High")
End Function

Private Sub Hdr(ByVal rng As Range, Optional ByVal v As Variant)
    If Not IsMissing(v) Then rng.Cells(1, 1).Value = v
    With rng
        .Font.Name = "Aptos Narrow": .Font.Size = 10: .Font.Bold = True: .Font.Color = C_WHITE
        .Interior.Color = C_NAVY
        .Borders.LineStyle = 1: .Borders.Color = C_BORDER
        .HorizontalAlignment = X_CENTER: .VerticalAlignment = X_CENTER: .WrapText = True
    End With
End Sub

Private Sub Sec(ByVal ws As Worksheet, ByVal r As Long, ByVal txt As String, ByVal lastCol As Long)
    ws.Range(ws.Cells(r, 1), ws.Cells(r, lastCol)).Interior.Color = C_BLUE
    With ws.Cells(r, 1)
        .Value = txt
        .Font.Name = "Aptos Narrow": .Font.Size = 10: .Font.Bold = True: .Font.Color = C_WHITE
    End With
End Sub

Private Sub Lab(ByVal ws As Worksheet, ByVal r As Long, ByVal b As String, ByVal c As String)
    If Len(b) > 0 Then ws.Cells(r, 2).Formula = b
    If Len(c) > 0 Then ws.Cells(r, 3).Formula = c
    ws.Range(ws.Cells(r, 2), ws.Cells(r, 3)).Font.Name = "Aptos Narrow"
    ws.Range(ws.Cells(r, 2), ws.Cells(r, 3)).Font.Size = 9
    With ws.Range(ws.Cells(r, 1), ws.Cells(r, 3)).Borders
        .LineStyle = 1: .Color = C_BORDER
    End With
End Sub

Private Sub Bold9(ByVal c As Range)
    c.Font.Name = "Aptos Narrow": c.Font.Size = 9: c.Font.Bold = True
End Sub

Private Function ChainFormula(conds() As String, vals() As String, ByVal n As Long) As String
    Dim f As String, i As Long
    f = "0"
    For i = n - 1 To 0 Step -1
        f = "IF(" & conds(i) & "," & vals(i) & "," & f & ")"
    Next i
    ChainFormula = "=" & f
End Function

Private Sub Chain(ByVal grp As String, offs() As Long, days() As String, n As Long)
    Select Case grp
        Case "STD"
            n = 4: offs(0) = 0: days(0) = "MD": offs(1) = 1: days(1) = "MD-1"
            offs(2) = 2: days(2) = "MD-2": offs(3) = -1: days(3) = "MD+1"
        Case "NONMD"
            n = 3: offs(0) = 1: days(0) = "MD-1": offs(1) = 2: days(1) = "MD-2": offs(2) = -1: days(2) = "MD+1"
        Case Else
            n = 1: offs(0) = 0: days(0) = "MD"
    End Select
End Sub

Private Sub FreezeAt(ByVal ws As Worksheet, ByVal addr As String)
    ws.Activate
    ActiveWindow.FreezePanes = False
    ActiveWindow.DisplayGridlines = False
    If Len(addr) > 0 Then
        ws.Range(addr).Select
        ActiveWindow.FreezePanes = True
    End If
    ws.Range("A1").Select
End Sub
'''

MAIN = r'''
' ======================================================================== ENTRY POINT
Public Sub BuildEuro2028Workbook()
    RunBuild
End Sub

Private Sub RunBuild()
    Dim wb As Workbook, nm As Variant, t0 As Double
    On Error GoTo Failed
    StepName = "checking the workbook"
    Set wb = ActiveWorkbook
    If Not SheetExists(wb, "Match Day Resource") Then Set wb = ThisWorkbook
    For Each nm In Array("Guidance", "Match Day Resource", "Wider Tournament Resource")
        If Not SheetExists(wb, CStr(nm)) Then
            If Not Quiet Then MsgBox "Open the BLANK v1 template first - sheet '" & nm & "' was not found.", vbExclamation
            Exit Sub
        End If
    Next nm
    If SheetExists(wb, "Demand Sheet") Then
        If Not Quiet Then MsgBox "This workbook already has the calculation pages. Run the macro on the blank template.", vbExclamation
        Exit Sub
    End If
    If Not Quiet Then
        If MsgBox("Build the EURO 2028 collection workbook in '" & wb.Name & "'?" & vbLf & vbLf & _
                  "This changes Match Day Resource (risk level headers, Super High) and adds the " & _
                  "calculation pages. It takes about one to three minutes.", vbOKCancel + vbQuestion) <> vbOK Then Exit Sub
    End If
    t0 = Timer
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = -4135
    StepName = "loading data": LoadData
    StepName = "Match Day Resource layout": MatchDayLayout wb.Worksheets("Match Day Resource")
    StepName = "adding pages": AddPages wb
    StepName = "Demand / Cost Sheet": CalendarPages wb
    StepName = "MD templates": MdTemplates wb
    StepName = "Fleet": FleetPage wb
    StepName = "Cost Calculator": CostCalculatorPage wb
    StepName = "Wider Tournament pages": WtPages wb
    StepName = "static pages": StaticPages wb
    StepName = "finishing": Finish wb
    StepName = "done"
    Application.Calculation = -4105
    Application.Calculate
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    If Not Quiet Then MsgBox "Done in " & Format(Timer - t0, "0") & " seconds." & vbLf & vbLf & _
        "Now use File > Save As and choose 'Excel Workbook (*.xlsx)'.", vbInformation
    Exit Sub
Failed:
    Dim msg As String
    msg = "Stopped while " & StepName & ": " & Err.Description & " (error " & Err.Number & ")"
    Application.Calculation = -4105
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.CutCopyMode = False
    StepName = msg
    If Not Quiet Then MsgBox msg & vbLf & vbLf & "Close the workbook WITHOUT saving, reopen the blank template and try again.", vbCritical
End Sub

Private Function SheetExists(ByVal wb As Workbook, ByVal nm As String) As Boolean
    Dim ws As Object
    For Each ws In wb.Sheets
        If ws.Name = nm Then SheetExists = True: Exit Function
    Next ws
End Function

Private Sub LoadData()
    Set MDL = New Collection: Set WTL = New Collection: Set CCL = New Collection
%LOADCALLS%
End Sub

' ===================================================== 1. Match Day Resource layout
Private Sub MatchDayLayout(ByVal ws As Worksheet)
    Const LR As Long = 400
    Dim targ(0 To 35) As Long
    Dim d As Long, s As Long, m As Long, p As Long, q As Long, i As Long, r As Long
    Dim refs As String, bannerColour As Long
    bannerColour = ws.Range("D3").Interior.Color
    For d = 0 To 3: For s = 0 To 2: For m = 0 To 2
        targ(s * 12 + d * 3 + m) = 4 + d * 9 + s * 3 + m
    Next m: Next s: Next d
    ws.Range("D3:AV3").UnMerge
    ws.AutoFilterMode = False
    ' risk level first: rows 4 onwards of D:AM are moved (cut and paste, so formats move with them and
    ' nothing is duplicated) to a spare area, then back one column at a time in risk-first order
    Const SPARE As Long = 1000
    ws.Range(ws.Cells(4, 4), ws.Cells(LR, 39)).Cut Destination:=ws.Cells(4, SPARE)
    For p = 0 To 35
        q = targ(p) - 4
        ws.Range(ws.Cells(4, SPARE + q), ws.Cells(LR, SPARE + q)).Cut Destination:=ws.Cells(4, 4 + p)
    Next p
    Application.CutCopyMode = False
    ' row 3 banners all styled like D3 (text is rewritten below)
    ws.Range("D3").Copy Destination:=ws.Range("D3:AM3")
    ' Super High (AN:AY) styled like High, and three more total columns (BI:BK) styled like the High totals
    ws.Columns("AN:AY").Insert Shift:=X_TORIGHT
    ws.Columns("BI:BK").Insert Shift:=X_TORIGHT
    ws.Range("AB3:AM" & LR).Copy Destination:=ws.Range("AN3")
    ws.Range("AN5:AY" & LR).ClearContents
    ws.Range("BF3:BH" & LR).Copy Destination:=ws.Range("BI3")
    Application.CutCopyMode = False
    ws.Range("E:BK").ColumnWidth = 26.85546875
    ' headers
    ws.Range("D3:AY3").ClearContents
    ws.Range("BA3:BK3").ClearContents
    For s = 0 To 3
        ws.Cells(3, 4 + s * 12).Value = Scen(s)
        ws.Range(ws.Cells(3, 4 + s * 12), ws.Cells(3, 15 + s * 12)).Merge
        For d = 0 To 3: For m = 0 To 2
            ws.Cells(4, 4 + s * 12 + d * 3 + m).Value = _
                Choose(d + 1, "Match Day -2", "Match Day -1", "Match Day", "Match Day +1") & " - " & _
                Choose(m + 1, "Number required", "Normal ToD", "Extra Hours")
        Next m: Next d
        For m = 0 To 2
            ws.Cells(4, 52 + s * 3 + m).Value = Scen(s) & " - " & Choose(m + 1, "Number required", "Normal ToD", "Extra Hours")
        Next m
    Next s
    ws.Range("AZ3:BK3").UnMerge
    ws.Range("AZ3:BK3").Merge
    ws.Range("D3:AY3").Interior.Color = bannerColour
    ' Automatic Total Figures: direct sums of the four day columns
    For r = 5 To LR
        If ws.Cells(r, 52).HasFormula Then
            For i = 0 To 11
                s = i \ 3: m = i Mod 3
                refs = ""
                For d = 0 To 3
                    refs = refs & IIf(d > 0, ",", "") & ColL(4 + s * 12 + d * 3 + m) & r
                Next d
                ws.Cells(r, 52 + i).Formula = "=SUM(" & refs & ")"
            Next i
        End If
    Next r
    ' filter on the header row, drop-down buttons hidden on the risk/total banners (as the template)
    ws.Range("A3:BK328").AutoFilter
    For i = 4 To 62
        ws.Range("A3:BK328").AutoFilter Field:=i, VisibleDropDown:=False
    Next i
End Sub

' ===================================================== 2. calculation pages
Private Sub AddPages(ByVal wb As Workbook)
    Dim nm As Variant, ws As Worksheet
    For Each nm In Array("Demand Sheet", "Cost Sheet", "MD Demand Template", "MD Cost Template", "Fleet", _
                         "Cost Calculator", "Wider Tournament Template", "Command Team & Planning", _
                         "Other Resources", "Baseline Costs", "GDP Deflators", "WT Cost Sheet", "WT Demand Sheet")
        Set ws = wb.Worksheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        ws.Name = CStr(nm)
    Next nm
End Sub

Private Sub CalendarPages(ByVal wb As Workbook)
    Dim dem As Worksheet, cst As Worksheet, ws As Worksheet
    Dim k As Long, s As Long, c0 As Long, wd As Long, w As Long, r As Long, i As Long, n As Long, kk As Long
    Dim parts() As String, it As Variant, own As Long, sr As Long, cfg As Long, cc As Long
    Dim grp As String, pk As String
    Dim offs(0 To 3) As Long, days(0 To 3) As String
    Dim cd(0 To 3) As String, vd(0 To 3) As String, ccd(0 To 3) As String, vc(0 To 3) As String, nb As Long
    Dim fd() As Variant, fc() As Variant
    Set dem = wb.Worksheets("Demand Sheet"): Set cst = wb.Worksheets("Cost Sheet")
    For w = 0 To 1
        If w = 0 Then
            Set ws = dem: wd = NR + 1
        Else
            Set ws = cst: wd = NR
        End If
        Hdr ws.Range("A3"), "Cell Input Instructions": Hdr ws.Range("B3"), "Resource ": Hdr ws.Range("C3"), "Rank"
        ws.Columns(1).ColumnWidth = 14: ws.Columns(2).ColumnWidth = 50: ws.Columns(3).ColumnWidth = 26
        Hdr ws.Range(ws.Cells(3, 4), ws.Cells(3, 3 + wd * CAL_DAYS))
        Hdr ws.Range(ws.Cells(5, 4), ws.Cells(5, 3 + wd * CAL_DAYS))
        ws.Range(ws.Columns(4), ws.Columns(3 + wd * CAL_DAYS)).ColumnWidth = 7
        For k = 0 To CAL_DAYS - 1
            c0 = 4 + wd * k
            ws.Cells(3, c0).Value = DateSerial(2028, 6, 7) + k
            ws.Cells(3, c0).NumberFormat = "ddd dd mmm yyyy"
            For s = 0 To wd - 1
                ws.Cells(5, c0 + s).Value = Choose(s + 1, "L", "M", "H", "SH", "Total")
            Next s
            For s = 0 To NR - 1
                If w = 0 Then
                    ws.Cells(4, c0 + s).Interior.Color = C_INPUT
                    ws.Cells(4, c0 + s).Borders.LineStyle = 1: ws.Cells(4, c0 + s).Borders.Color = C_BORDER
                Else
                    ws.Cells(4, c0 + s).Formula = "='Demand Sheet'!" & ColL(DemandCol(k, s)) & "4"
                    ws.Cells(4, c0 + s).NumberFormat = "0;-0;;@"
                End If
            Next s
        Next k
    Next w
    dem.Range("A4").Value = "Fixture code >": Bold9 dem.Range("A4")
    dem.Range("A1").Value = "Demand Sheet - type the fixture code (e.g. A1) in row 4 under Low, Medium, High or Super High on each match date. The risk column chosen sets the fixture's risk level."
    cst.Range("A1").Value = "Cost Sheet - fixture codes are copied from the Demand Sheet."
    Bold9 dem.Range("A1"): Bold9 cst.Range("A1")
    ReDim fd(1 To 1, 1 To (NR + 1) * CAL_DAYS)
    ReDim fc(1 To 1, 1 To NR * CAL_DAYS)
    r = 7
    For Each it In MDL
        parts = Split(it, "|")
        If parts(0) = "S" Then
            Sec dem, r, parts(1), 4 + (NR + 1) * CAL_DAYS - 1
            Sec cst, r, parts(1), 4 + NR * CAL_DAYS - 1
        Else
            sr = CLng(parts(1)): grp = parts(3): pk = parts(4): cc = CLng(parts(5))
            If Len(parts(2)) > 0 Then cfg = CLng(parts(2)) Else cfg = 0
            Lab dem, r, parts(6), parts(7): Lab cst, r, parts(6), parts(7)
            Chain grp, offs, days, n
            For k = 0 To CAL_DAYS - 1
                For s = 0 To NR - 1
                    own = DemandCol(k, s)
                    nb = 0
                    For i = 0 To n - 1
                        kk = k + offs(i)
                        If kk >= 0 Then
                            cd(nb) = "'Demand Sheet'!" & ColL(DemandCol(kk, s)) & "4>0"
                            If pk = "PSU" Then
                                vd(nb) = "SUM(" & SurveyNum(days(i), s, cfg) & "*" & SurveyNum(days(i), s, sr) & ")"
                            Else
                                vd(nb) = SurveyNum(days(i), s, sr)
                            End If
                            ccd(nb) = "'Cost Sheet'!$" & ColL(CostCol(kk, s)) & "$4>0"
                            vc(nb) = "SUM('Demand Sheet'!" & ColL(own) & r & "*'Cost Calculator'!" & CCCol(days(i), s) & cc & ")"
                            nb = nb + 1
                        End If
                    Next i
                    fd(1, own - 3) = ChainFormula(cd, vd, nb)
                    fc(1, CostCol(k, s) - 3) = ChainFormula(ccd, vc, nb)
                Next s
                c0 = DemandCol(k, 0)
                fd(1, c0 + NR - 3) = "=AGGREGATE(9,6," & ColL(c0) & r & ":" & ColL(c0 + NR - 1) & r & ")"
            Next k
            dem.Range(dem.Cells(r, 4), dem.Cells(r, 3 + (NR + 1) * CAL_DAYS)).Formula = fd
            cst.Range(cst.Cells(r, 4), cst.Cells(r, 3 + NR * CAL_DAYS)).Formula = fc
            For k = 0 To CAL_DAYS - 1
                dem.Cells(r, DemandCol(k, NR)).Font.Bold = True
            Next k
        End If
        r = r + 1
    Next it
    cst.Range(cst.Cells(7, 4), cst.Cells(r, 3 + NR * CAL_DAYS)).NumberFormat = "#,##0"
    FreezeAt dem, "D6": FreezeAt cst, "D6"
End Sub

Private Sub MdTemplates(ByVal wb As Workbook)
    Dim ws As Worksheet, w As Long, src As String, k As Long, r As Long, lastR As Long, c As Long
    Dim parts() As String, it As Variant, f() As Variant
    For w = 0 To 1
        If w = 0 Then src = "Demand Sheet" Else src = "Cost Sheet"
        Set ws = wb.Worksheets(IIf(w = 0, "MD Demand Template", "MD Cost Template"))
        ws.Range("A1").Value = ws.Name & IIf(w = 0, " - officers per resource per day (all risk levels)", _
                                                  " - cost per resource per day (all risk levels)")
        Bold9 ws.Range("A1")
        Hdr ws.Range("A3"), "Cell Input Instructions": Hdr ws.Range("B3"), "Resource ": Hdr ws.Range("C3"), "Rank"
        Hdr ws.Range(ws.Cells(3, 4), ws.Cells(3, 4 + CAL_DAYS))
        For k = 0 To CAL_DAYS - 1
            ws.Cells(3, 4 + k).Value = DateSerial(2028, 6, 7) + k
            ws.Cells(3, 4 + k).NumberFormat = "dd-mmm"
        Next k
        ws.Range(ws.Columns(4), ws.Columns(3 + CAL_DAYS)).ColumnWidth = 9
        ws.Cells(3, 4 + CAL_DAYS).Value = "TOTAL"
        ws.Columns(1).ColumnWidth = 14: ws.Columns(2).ColumnWidth = 50: ws.Columns(3).ColumnWidth = 26
        ReDim f(1 To 1, 1 To CAL_DAYS + 1)
        r = 7: lastR = 7
        For Each it In MDL
            parts = Split(it, "|")
            If parts(0) = "S" Then
                Sec ws, r, parts(1), 4 + CAL_DAYS
            Else
                Lab ws, r, IIf(Len(parts(6)) > 0, "='" & src & "'!B" & r, ""), "='" & src & "'!C" & r
                For k = 0 To CAL_DAYS - 1
                    If w = 0 Then
                        f(1, k + 1) = "=IFERROR('Demand Sheet'!" & ColL(DemandCol(k, NR)) & r & ","""")"
                    Else
                        f(1, k + 1) = "=IFERROR(AGGREGATE(9,6,'Cost Sheet'!" & ColL(CostCol(k, 0)) & r & ":" & _
                                      ColL(CostCol(k, NR - 1)) & r & "),"""")"
                    End If
                Next k
                f(1, CAL_DAYS + 1) = "=SUM(D" & r & ":" & ColL(3 + CAL_DAYS) & r & ")"
                ws.Range(ws.Cells(r, 4), ws.Cells(r, 4 + CAL_DAYS)).Formula = f
                lastR = r
            End If
            r = r + 1
        Next it
        ws.Range(ws.Cells(7, 4), ws.Cells(lastR, 4 + CAL_DAYS)).NumberFormat = "#,##0"
        ws.Cells(lastR + 2, 2).Value = "TOTAL": Bold9 ws.Cells(lastR + 2, 2)
        For c = 4 To 4 + CAL_DAYS
            ws.Cells(lastR + 2, c).Formula = "=SUM(" & ColL(c) & "7:" & ColL(c) & lastR & ")"
            ws.Cells(lastR + 2, c).NumberFormat = "#,##0"
            Bold9 ws.Cells(lastR + 2, c)
        Next c
        FreezeAt ws, "D4"
    Next w
End Sub

Private Sub FleetPage(ByVal wb As Workbook)
    Dim ws As Worksheet, k As Long, r As Long, parts() As String, it As Variant, f() As Variant
    Set ws = wb.Worksheets("Fleet")
    ws.Range("A1").Value = "Fleet - officers per resource per day (from the Demand Sheet)": Bold9 ws.Range("A1")
    Hdr ws.Range("A3"), "Location": Hdr ws.Range("B3"), "Resource ": Hdr ws.Range("C3"), "Rank"
    Hdr ws.Range(ws.Cells(3, 4), ws.Cells(3, 3 + CAL_DAYS))
    For k = 0 To CAL_DAYS - 1
        ws.Cells(3, 4 + k).Value = DateSerial(2028, 6, 7) + k
        ws.Cells(3, 4 + k).NumberFormat = "dd-mmm"
    Next k
    ReDim f(1 To 1, 1 To CAL_DAYS)
    r = 7
    For Each it In MDL
        parts = Split(it, "|")
        If parts(0) = "S" Then
            ws.Cells(r, 1).Value = parts(1): Bold9 ws.Cells(r, 1)
        Else
            Lab ws, r, IIf(Len(parts(6)) > 0, "='Demand Sheet'!B" & r, ""), "='Demand Sheet'!C" & r
            For k = 0 To CAL_DAYS - 1
                f(1, k + 1) = "='Demand Sheet'!" & ColL(DemandCol(k, NR)) & r
            Next k
            ws.Range(ws.Cells(r, 4), ws.Cells(r, 3 + CAL_DAYS)).Formula = f
        End If
        r = r + 1
    Next it
    FreezeAt ws, ""
End Sub

Private Sub CostCalculatorPage(ByVal wb As Workbook)
    Dim ws As Worksheet, i As Long, s As Long, d As Long, top As Long, tc As Long, lastC As String
    Dim parts() As String, it As Variant, sr As Long, outR As Long, nCols As Long, survey As String
    Dim xh(0 To 15) As String, f() As Variant
    Set ws = wb.Worksheets("Cost Calculator")
    ws.Range("A1").Value = "Cost Calculator - cost per officer (Extra Hours x hourly rate) or per mutual-aid unit (package)"
    Bold9 ws.Range("A1")
    Hdr ws.Range("A3"), "Cell Input Instructions": Hdr ws.Range("B3"), "Resource ": Hdr ws.Range("C3"), "Rank"
    For i = 0 To 3
        Hdr ws.Cells(3, 4 + NR * i), Choose(i + 1, "Match Day -2", "Match Day -1", "Match Day", "Match Day +1")
        For s = 0 To NR - 1
            Hdr ws.Cells(4, 4 + NR * i + s), Scen(s) & "-Number required"
        Next s
    Next i
    Hdr ws.Cells(3, 4 + 4 * NR), "Automatic Total Figures Calculation"
    For s = 0 To NR - 1
        Hdr ws.Cells(4, 4 + 4 * NR + s), Scen(s) & "-Number required"
    Next s
    ws.Columns(1).ColumnWidth = 18: ws.Columns(2).ColumnWidth = 50: ws.Columns(3).ColumnWidth = 24
    ws.Range("A5").Value = "MATCH DAY": Bold9 ws.Range("A5")
    top = WT_CC_OFFSET
    ws.Cells(top - 2, 1).Value = "WIDER TOURNAMENT": Bold9 ws.Cells(top - 2, 1)
    For s = 0 To 3
        Hdr ws.Cells(top - 1, 4 + s), Choose(s + 1, "Low", "Medium", "High", "Super High")
    Next s
    lastC = ColL(3 + 4 * NR)
    For Each it In CCL
        parts = Split(it, "|")
        sr = CLng(parts(2)): outR = CLng(parts(3))
        If parts(1) = "MD" Then
            survey = MDS: nCols = 4 * NR
            For d = 0 To 3: For s = 0 To NR - 1
                xh(d * NR + s) = ColL(4 + s * 12 + d * 3 + 2)
            Next s: Next d
        Else
            survey = WTS: nCols = 4
            For s = 0 To 3: xh(s) = ColL(4 + 3 * s + 2): Next s
        End If
        Lab ws, outR, parts(4), parts(5)
        ReDim f(1 To 1, 1 To nCols)
        For i = 0 To nCols - 1
            Select Case parts(6)
                Case "PKG": f(1, i + 1) = "=IFERROR(SUM(" & BCS & "!$X$17),0)"
                Case "RATE": f(1, i + 1) = "=IFERROR(SUM(" & survey & "!" & xh(i) & sr & "*" & parts(7) & "),0)"
                Case "ZERO": f(1, i + 1) = 0
            End Select
        Next i
        If parts(6) <> "SKIP" Then
            ws.Range(ws.Cells(outR, 4), ws.Cells(outR, 3 + nCols)).Formula = f
            ws.Range(ws.Cells(outR, 4), ws.Cells(outR, 3 + nCols)).NumberFormat = "#,##0.00"
        End If
        If parts(8) = "1" Then
            For s = 0 To NR - 1
                tc = 4 + 4 * NR + s
                ws.Cells(outR, tc).Formula = "=SUMIF($D$4:$" & lastC & "$4," & ColL(tc) & "$4,D" & outR & ":" & lastC & outR & ")"
                ws.Cells(outR, tc).NumberFormat = "#,##0.00"
            Next s
        End If
    Next it
    FreezeAt ws, "D5"
End Sub

Private Sub WtPages(ByVal wb As Workbook)
    Dim dem As Worksheet, cst As Worksheet, tpl As Worksheet, ws As Worksheet, w As Long, c As Long
    Dim heads As Variant, i As Long, r As Long, firstR As Long, lastR As Long, cc As Long, sr As Long, cfg As Long
    Dim parts() As String, it As Variant, nm As String, num As String, cols As Variant
    Set dem = wb.Worksheets("WT Demand Sheet"): Set cst = wb.Worksheets("WT Cost Sheet")
    Set tpl = wb.Worksheets("Wider Tournament Template")
    For w = 0 To 2
        Set ws = Choose(w + 1, dem, cst, tpl)
        If w < 2 Then
            heads = Array("Cell Input Instructions", "Resource ", "Rank", "Low", "Medium", "High", "Super High")
        Else
            heads = Array("Cell Input Instructions", "Resource ", "Rank", "COST ON LOW", "COST ON MEDIUM", "COST ON HIGH", _
                          "COST ON SUPER HIGH", "NUMBER OF LOW", "NUMBER OF MEDIUM", "NUMBER OF HIGH", _
                          "NUMBER OF SUPER HIGH", "TOTAL")
        End If
        For c = 0 To UBound(heads)
            Hdr ws.Cells(3, c + 1), heads(c)
            ws.Columns(c + 1).ColumnWidth = IIf(c < 3, Choose(c + 1, 14, 50, 26), 14)
        Next c
    Next w
    dem.Range("A1").Value = "WT Demand Sheet - officers (or mutual-aid units) per typical wider-tournament day"
    cst.Range("A1").Value = "WT Cost Sheet - cost per typical wider-tournament day"
    tpl.Range("A1").Value = "Wider Tournament Template - cost per day x number of days at each risk level (days set once on Baseline Costs)"
    Bold9 dem.Range("A1"): Bold9 cst.Range("A1"): Bold9 tpl.Range("A1")
    r = 5: firstR = 0
    For Each it In WTL
        parts = Split(it, "|")
        If parts(0) = "S" Then
            Sec dem, r, parts(1), 7: Sec cst, r, parts(1), 7: Sec tpl, r, parts(1), 12
        Else
            If firstR = 0 Then firstR = r
            sr = CLng(parts(1)): cc = WT_CC_OFFSET + CLng(parts(4))
            If Len(parts(2)) > 0 Then cfg = CLng(parts(2)) Else cfg = 0
            Lab dem, r, parts(5), parts(6): Lab cst, r, parts(5), parts(6): Lab tpl, r, parts(5), parts(6)
            For i = 0 To 3
                num = ColL(4 + 3 * i)
                If parts(3) = "PSU" Then
                    dem.Cells(r, 4 + i).Formula = "=IFERROR(SUM(" & WTS & "!" & num & cfg & "*" & WTS & "!" & num & sr & "), 0)"
                Else
                    dem.Cells(r, 4 + i).Formula = "=IFERROR(" & WTS & "!" & num & sr & ",0)"
                End If
                cst.Cells(r, 4 + i).Formula = "=IFERROR(SUM('WT Demand Sheet'!" & ColL(4 + i) & r & "*'Cost Calculator'!" & ColL(4 + i) & cc & "),0)"
                cst.Cells(r, 4 + i).NumberFormat = "#,##0"
                tpl.Cells(r, 4 + i).Formula = "='WT Cost Sheet'!" & ColL(4 + i) & r
                tpl.Cells(r, 4 + i).NumberFormat = "#,##0"
                tpl.Cells(r, 8 + i).Formula = "=" & BCS & "!$C$" & (43 + i)
            Next i
            tpl.Cells(r, 12).Formula = "=IFERROR(SUM(D" & r & "*H" & r & ")+SUM(E" & r & "*I" & r & ")+SUM(F" & r & "*J" & r & ")+SUM(G" & r & "*K" & r & "),0)"
            tpl.Cells(r, 12).NumberFormat = "#,##0"
            lastR = r
        End If
        r = r + 1
    Next it
    For w = 0 To 2
        Set ws = Choose(w + 1, dem, cst, tpl)
        If w < 2 Then cols = Array(4, 5, 6, 7) Else cols = Array(4, 5, 6, 7, 12)
        ws.Cells(lastR + 2, 2).Value = "TOTAL": Bold9 ws.Cells(lastR + 2, 2)
        For c = 0 To UBound(cols)
            ws.Cells(lastR + 2, cols(c)).Formula = "=SUM(" & ColL(cols(c)) & firstR & ":" & ColL(cols(c)) & lastR & ")"
            ws.Cells(lastR + 2, cols(c)).NumberFormat = "#,##0"
            Bold9 ws.Cells(lastR + 2, cols(c))
        Next c
        FreezeAt ws, "D4"
    Next w
End Sub

' ----------------------------------------------- static pages (copied from the original)
Private Sub StaticPages(ByVal wb As Workbook)
    Dim bc As Worksheet, i As Long
%STATICCALLS%
    Set bc = wb.Worksheets("Baseline Costs")
    bc.Range("B35").Value = "RANK RATE LOOKUP - used by Rank (Drop Down) rows (Mounted, Dogs, Roads Policing)"
    Bold9 bc.Range("B35")
    Hdr bc.Range("B36"), "Rank": Hdr bc.Range("C36"), "Force supplied": Hdr bc.Range("D36"), "Mutual aid"
    bc.Range("B37").Value = "Inspector": bc.Range("C37").Value = 0: bc.Range("D37").Formula = "=$X$13"
    bc.Range("B38").Value = "Sergeant": bc.Range("C38").Formula = "=$X$31": bc.Range("D38").Formula = "=$X$14"
    bc.Range("B39").Value = "Constable": bc.Range("C39").Formula = "=$X$32": bc.Range("D39").Formula = "=$X$15"
    bc.Range("C37:D39").NumberFormat = "#,##0.00"
    bc.Range("B41").Value = "WIDER TOURNAMENT - NUMBER OF DAYS AT EACH RISK LEVEL (used by Wider Tournament Template)"
    Bold9 bc.Range("B41")
    Hdr bc.Range("B42"), "Risk level": Hdr bc.Range("C42"), "Days"
    For i = 0 To 3
        bc.Cells(43 + i, 2).Value = Choose(i + 1, "Low", "Medium", "High", "Super High")
        bc.Cells(43 + i, 3).Value = Choose(i + 1, 7, 11, 11, 3)
        bc.Cells(43 + i, 3).Interior.Color = C_INPUT
    Next i
End Sub

Private Sub StaticPage(ByVal ws As Worksheet)
    Dim it As Variant, p() As String, c As Range
    For Each it In STL
        p = Split(it, "|")
        Select Case p(0)
            Case "C"
                Set c = ws.Range(p(1))
                Select Case p(2)
                    Case "n": c.Value = Val(p(3))
                    Case "d": c.Value = DateSerial(CLng(Left$(p(3), 4)), CLng(Mid$(p(3), 6, 2)), CLng(Mid$(p(3), 9, 2)))
                    Case "f": c.Formula = p(3)
                    Case "s"
                        c.Formula = "'" & p(3)              ' leading apostrophe: always stored as text, e.g. "(1)"
                End Select
                c.NumberFormat = p(4)
                c.Font.Name = "Aptos Narrow": c.Font.Size = Val(p(7))
                c.Font.Bold = (p(5) = "1"): c.Font.Italic = (p(6) = "1")
                If Len(p(8)) > 0 Then c.Font.Color = CLng(p(8))
                If Len(p(9)) > 0 Then c.Interior.Color = CLng(p(9))
                Select Case p(10)
                    Case "center": c.HorizontalAlignment = X_CENTER
                    Case "left": c.HorizontalAlignment = -4131
                    Case "right": c.HorizontalAlignment = -4152
                    Case "centerContinuous": c.HorizontalAlignment = 7
                End Select
                If p(11) = "1" Then c.WrapText = True
                If p(12) = "1" Then c.Borders.LineStyle = 1: c.Borders.Color = C_BORDER
            Case "W"
                ws.Columns(p(1)).ColumnWidth = Val(p(2))
            Case "M"
                ws.Range(p(1)).Merge
        End Select
    Next it
End Sub

Private Sub Finish(ByVal wb As Workbook)
    Dim nm As Variant
    For Each nm In Array("Command Team & Planning", "Other Resources", "Baseline Costs", "GDP Deflators", "Wider Tournament Template")
        wb.Worksheets(CStr(nm)).Activate
        ActiveWindow.DisplayGridlines = False
    Next nm
    For Each nm In Array("Demand Sheet", "Cost Sheet", "Fleet", "Cost Calculator", "WT Cost Sheet", "WT Demand Sheet")
        wb.Worksheets(CStr(nm)).Visible = 0
    Next nm
    wb.Worksheets("Match Day Resource").Activate
    wb.Worksheets("Match Day Resource").Range("D5").Select
End Sub
'''

TEST = r'''
Public Function TestBuildNoPrompts() As String
    Quiet = True
    StepName = "starting"
    RunBuild
    TestBuildNoPrompts = StepName
End Function
'''


def build_vba(test=False):
    t = LY.Template(SOURCE)
    probs = LY.verify(SOURCE)
    if probs:
        raise SystemExit("\n".join(probs))
    procs, load_calls = [], []
    for prefix, fn, lines in (("DataMD", "AddMD", md_data(t)), ("DataWT", "AddWT", wt_data(t)),
                              ("DataCC", "AddCC", cc_data(t))):
        p, names = data_procs(prefix, fn, lines)
        procs += p
        load_calls += [f"    {n}" for n in names]
    static_calls = []
    for i, (name, lines) in enumerate(static_data().items()):
        p, names = data_procs(f"DataST{i + 1}_", "AddST", lines)
        procs += p
        static_calls.append("    Set STL = New Collection")
        static_calls += [f"    {n}" for n in names]
        static_calls.append(f'    StaticPage wb.Worksheets("{enc(name)}")')
    head = (HEAD.replace("%NAVY%", str(rgb(B.NAVY))).replace("%BLUE%", str(rgb(B.BLUE)))
            .replace("%LIGHT%", str(rgb(B.LIGHT))).replace("%BORDER%", str(rgb("8EA9DB")))
            .replace("%INPUT%", str(rgb("FFF2CC"))))
    main = MAIN.replace("%LOADCALLS%", "\n".join(load_calls)).replace("%STATICCALLS%", "\n".join(static_calls))
    text = head + main + (TEST if test else "") + "\n' ====================================================== DATA\n" + \
        "\n\n".join(procs) + "\n"
    for ln in text.split("\n"):
        assert len(ln) < 1000, ln[:80]
        assert all(ord(ch) < 127 for ch in ln), ln[:80]
    return text.replace("\n", "\r\n")


if __name__ == "__main__":
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    out = os.path.abspath(args[0] if args else OUT)
    with open(out, "w", newline="") as f:
        f.write(build_vba(test="--test" in sys.argv))
    print(out, sum(1 for _ in open(out)), "lines")
