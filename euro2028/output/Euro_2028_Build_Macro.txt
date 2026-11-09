' =====================================================================================
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
Private Const C_NAVY As Long = 6567967
Private Const C_BLUE As Long = 9917743
Private Const C_LIGHT As Long = 15917529
Private Const C_BORDER As Long = 14395790
Private Const C_INPUT As Long = 13431551
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
    DataMD1
    DataMD2
    DataMD3
    DataWT1
    DataCC1
    DataCC2
    DataCC3
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
    Set STL = New Collection
    DataST1_1
    DataST1_2
    DataST1_3
    DataST1_4
    StaticPage wb.Worksheets("Command Team & Planning")
    Set STL = New Collection
    DataST2_1
    DataST2_2
    StaticPage wb.Worksheets("Other Resources")
    Set STL = New Collection
    DataST3_1
    DataST3_2
    DataST3_3
    DataST3_4
    StaticPage wb.Worksheets("Baseline Costs")
    Set STL = New Collection
    DataST4_1
    DataST4_2
    DataST4_3
    DataST4_4
    DataST4_5
    DataST4_6
    StaticPage wb.Worksheets("GDP Deflators")
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

' ====================================================== DATA
Private Sub DataMD1()
    Dim s As String
    AddMD "S|Public Order / Public Safety"
    AddMD "L|33|7|NONMD|PSU|7|Force Supplied - Public Order (Non-Matchday)|='Match Day Resource'!$C$33&"" - Inspector"""
    AddMD "L|33|8|NONMD|PSU|8||='Match Day Resource'!$C$33&"" - Sergeant"""
    AddMD "L|33|9|NONMD|PSU|9||='Match Day Resource'!$C$33&"" - Constable"""
    AddMD "L|34|10|NONMD|PSU|10||='Match Day Resource'!$C$34&"" - Inspector"""
    AddMD "L|34|11|NONMD|PSU|11||='Match Day Resource'!$C$34&"" - Sergeant"""
    AddMD "L|34|12|NONMD|PSU|12||='Match Day Resource'!$C$34&"" - Constable"""
    AddMD "L|35||NONMD|MAPSU|35|Mutual Aid Requested - Public Order (Non-Matchday)|='Match Day Resource'!$C$35"
    AddMD "L|36||NONMD|MAPSU|36||='Match Day Resource'!$C$36"
    AddMD "L|37|13|MDONLY|PSU|13|Force Supplied - Town Centre Public Order Policing (Matchday Only)|='Match Day Resource'!$C$37&"" - Inspector"""
    AddMD "L|37|14|MDONLY|PSU|14||='Match Day Resource'!$C$37&"" - Sergeant"""
    AddMD "L|37|15|MDONLY|PSU|15||='Match Day Resource'!$C$37&"" - Constable"""
    AddMD "L|38|16|MDONLY|PSU|16||='Match Day Resource'!$C$38&"" - Inspector"""
    AddMD "L|38|17|MDONLY|PSU|17||='Match Day Resource'!$C$38&"" - Sergeant"""
    AddMD "L|38|18|MDONLY|PSU|18||='Match Day Resource'!$C$38&"" - Constable"""
    AddMD "L|39||MDONLY|MAPSU|39|Mutual Aid Requested - Town Centre Public Order Policing (Matchday Only)|='Match Day Resource'!$C$39"
    AddMD "L|40||MDONLY|MAPSU|40||='Match Day Resource'!$C$40"
    AddMD "L|41|19|MDONLY|PSU|19|Force Supplied - Stadium Vicinity Public Order Policing (Matchday Only)|='Match Day Resource'!$C$41&"" - Inspector"""
    AddMD "L|41|20|MDONLY|PSU|20||='Match Day Resource'!$C$41&"" - Sergeant"""
    AddMD "L|41|21|MDONLY|PSU|21||='Match Day Resource'!$C$41&"" - Constable"""
    AddMD "L|42|22|MDONLY|PSU|22||='Match Day Resource'!$C$42&"" - Inspector"""
    AddMD "L|42|23|MDONLY|PSU|23||='Match Day Resource'!$C$42&"" - Sergeant"""
    AddMD "L|42|24|MDONLY|PSU|24||='Match Day Resource'!$C$42&"" - Constable"""
    AddMD "L|43||MDONLY|MAPSU|43|Mutual Aid Requested - Stadium Vicinity Public Order Policing (Matchday Only)|='Match Day Resource'!$C$43"
    AddMD "L|44||MDONLY|MAPSU|44||='Match Day Resource'!$C$44"
    AddMD "L|45|25|MDONLY|PSU|25|Force Supplied - In Stadium Public Order Policing (Matchday Only)|='Match Day Resource'!$C$45&"" - Inspector"""
    AddMD "L|45|26|MDONLY|PSU|26||='Match Day Resource'!$C$45&"" - Sergeant"""
    AddMD "L|45|27|MDONLY|PSU|27||='Match Day Resource'!$C$45&"" - Constable"""
    AddMD "L|46|28|MDONLY|PSU|28||='Match Day Resource'!$C$46&"" - Inspector"""
    AddMD "L|46|29|MDONLY|PSU|29||='Match Day Resource'!$C$46&"" - Sergeant"""
    AddMD "L|46|30|MDONLY|PSU|30||='Match Day Resource'!$C$46&"" - Constable"""
    AddMD "L|47||MDONLY|MAPSU|47|Mutual Aid Requested - In Stadium Public Order Policing (Matchday Only)|='Match Day Resource'!$C$47"
    AddMD "L|48||MDONLY|MAPSU|48||='Match Day Resource'!$C$48"
    AddMD "S|Evidence Gathering Team Officer"
    AddMD "L|50||STD|OFFICER|50|Force Supplied - Evidence Gathering Team Officer|=IF('Match Day Resource'!$C$50="""","""",'Match Day Resource'!$C$50)"
    AddMD "L|51||STD|OFFICER|51||=IF('Match Day Resource'!$C$51="""","""",'Match Day Resource'!$C$51)"
    AddMD "L|52||STD|OFFICER|52|Mutual Aid Requested - Evidence Gathering Team Officer|=IF('Match Day Resource'!$C$52="""","""",'Match Day Resource'!$C$52)"
    AddMD "L|53||STD|OFFICER|53||=IF('Match Day Resource'!$C$53="""","""",'Match Day Resource'!$C$53)"
    AddMD "S|Protest Removal"
    AddMD "L|55||STD|OFFICER|55|Force Supplied - Working at Height Complex Protest Removal|=IF('Match Day Resource'!$C$55="""","""",'Match Day Resource'!$C$55)"
    AddMD "L|56||STD|OFFICER|56||=IF('Match Day Resource'!$C$56="""","""",'Match Day Resource'!$C$56)"
    AddMD "L|57||STD|OFFICER|57||=IF('Match Day Resource'!$C$57="""","""",'Match Day Resource'!$C$57)"
    AddMD "L|58||STD|OFFICER|58|Mutual Aid Requested - Working at Height Complex Protest Removal|=IF('Match Day Resource'!$C$58="""","""",'Match Day Resource'!$C$58)"
    AddMD "L|59||STD|OFFICER|59||=IF('Match Day Resource'!$C$59="""","""",'Match Day Resource'!$C$59)"
    AddMD "L|60||STD|OFFICER|60||=IF('Match Day Resource'!$C$60="""","""",'Match Day Resource'!$C$60)"
    AddMD "L|61||STD|OFFICER|61|Force Supplied - Non-Complex Protest Removal|=IF('Match Day Resource'!$C$61="""","""",'Match Day Resource'!$C$61)"
    AddMD "L|62||STD|OFFICER|62||=IF('Match Day Resource'!$C$62="""","""",'Match Day Resource'!$C$62)"
    AddMD "L|63||STD|OFFICER|63||=IF('Match Day Resource'!$C$63="""","""",'Match Day Resource'!$C$63)"
    AddMD "L|64||STD|OFFICER|64|Mutual Aid Requested - Non-Complex Protest Removal|=IF('Match Day Resource'!$C$64="""","""",'Match Day Resource'!$C$64)"
    AddMD "L|65||STD|OFFICER|65||=IF('Match Day Resource'!$C$65="""","""",'Match Day Resource'!$C$65)"
    AddMD "L|66||STD|OFFICER|66||=IF('Match Day Resource'!$C$66="""","""",'Match Day Resource'!$C$66)"
    AddMD "L|67||STD|OFFICER|67|Force Supplied - Debonding Protest Removal|=IF('Match Day Resource'!$C$67="""","""",'Match Day Resource'!$C$67)"
    AddMD "L|68||STD|OFFICER|68||=IF('Match Day Resource'!$C$68="""","""",'Match Day Resource'!$C$68)"
    AddMD "L|69||STD|OFFICER|69||=IF('Match Day Resource'!$C$69="""","""",'Match Day Resource'!$C$69)"
    AddMD "L|70||STD|OFFICER|70|Mutual Aid Requested - Debonding Protest Removal|=IF('Match Day Resource'!$C$70="""","""",'Match Day Resource'!$C$70)"
    AddMD "L|71||STD|OFFICER|71||=IF('Match Day Resource'!$C$71="""","""",'Match Day Resource'!$C$71)"
    AddMD "L|72||STD|OFFICER|72||=IF('Match Day Resource'!$C$72="""","""",'Match Day Resource'!$C$72)"
    AddMD "S|Additional Public Order / Public Safety Resources"
    AddMD "L|74||STD|OFFICER|74|Force Supplied - Additional Public Order|=IF('Match Day Resource'!$C$74="""","""",'Match Day Resource'!$C$74)"
    AddMD "L|75||STD|OFFICER|75||=IF('Match Day Resource'!$C$75="""","""",'Match Day Resource'!$C$75)"
    AddMD "L|76||STD|OFFICER|76||=IF('Match Day Resource'!$C$76="""","""",'Match Day Resource'!$C$76)"
    AddMD "L|77||STD|OFFICER|77||=IF('Match Day Resource'!$C$77="""","""",'Match Day Resource'!$C$77)"
    AddMD "L|78||STD|OFFICER|78||=IF('Match Day Resource'!$C$78="""","""",'Match Day Resource'!$C$78)"
    AddMD "L|79||STD|OFFICER|79||=IF('Match Day Resource'!$C$79="""","""",'Match Day Resource'!$C$79)"
    AddMD "L|80||STD|OFFICER|80|Mutual Aid Requested -   Additional PO|=IF('Match Day Resource'!$C$80="""","""",'Match Day Resource'!$C$80)"
    AddMD "L|81||STD|OFFICER|81||=IF('Match Day Resource'!$C$81="""","""",'Match Day Resource'!$C$81)"
    AddMD "L|82||STD|OFFICER|82||=IF('Match Day Resource'!$C$82="""","""",'Match Day Resource'!$C$82)"
    AddMD "L|83||STD|OFFICER|83||=IF('Match Day Resource'!$C$83="""","""",'Match Day Resource'!$C$83)"
    AddMD "L|84||STD|OFFICER|84||=IF('Match Day Resource'!$C$84="""","""",'Match Day Resource'!$C$84)"
    AddMD "L|85||STD|OFFICER|85||=IF('Match Day Resource'!$C$85="""","""",'Match Day Resource'!$C$85)"
    AddMD "S|Mounted"
    AddMD "L|88||STD|OFFICER|88|Force Supplied - Mounted Commander|=IF('Match Day Resource'!$C$88="""","""",'Match Day Resource'!$C$88)"
    AddMD "L|89||STD|OFFICER|89|Mutual Aid Requested -   Mounted Commander|=IF('Match Day Resource'!$C$89="""","""",'Match Day Resource'!$C$89)"
    AddMD "L|90||STD|OFFICER|90|Force Supplied - Mounted Officer|=IF('Match Day Resource'!$C$90="""","""",'Match Day Resource'!$C$90)"
    AddMD "L|91||STD|OFFICER|91|Mutual Aid Requested -   Mounted Officer|=IF('Match Day Resource'!$C$91="""","""",'Match Day Resource'!$C$91)"
    AddMD "S|Dogs"
    AddMD "L|94||STD|OFFICER|94|Force Supplied - Dogs GDP|=IF('Match Day Resource'!$C$94="""","""",'Match Day Resource'!$C$94)"
    AddMD "L|95||STD|OFFICER|95||=IF('Match Day Resource'!$C$95="""","""",'Match Day Resource'!$C$95)"
    AddMD "L|96||STD|OFFICER|96||=IF('Match Day Resource'!$C$96="""","""",'Match Day Resource'!$C$96)"
    AddMD "L|97||STD|OFFICER|97|Mutual Aid Requested- Dogs GDP|=IF('Match Day Resource'!$C$97="""","""",'Match Day Resource'!$C$97)"
    AddMD "L|98||STD|OFFICER|98||=IF('Match Day Resource'!$C$98="""","""",'Match Day Resource'!$C$98)"
    AddMD "L|99||STD|OFFICER|99||=IF('Match Day Resource'!$C$99="""","""",'Match Day Resource'!$C$99)"
    AddMD "L|100||STD|OFFICER|100|Force Supplied - Dogs Public Order Serial Commander|=IF('Match Day Resource'!$C$100="""","""",'Match Day Resource'!$C$100)"
    AddMD "L|101||STD|OFFICER|101||=IF('Match Day Resource'!$C$101="""","""",'Match Day Resource'!$C$101)"
    AddMD "L|102||STD|OFFICER|102||=IF('Match Day Resource'!$C$102="""","""",'Match Day Resource'!$C$102)"
    AddMD "L|103||STD|OFFICER|103|Mutual Aid Requested - Dogs Public Order Serial Commander|=IF('Match Day Resource'!$C$103="""","""",'Match Day Resource'!$C$103)"
    AddMD "L|104||STD|OFFICER|104||=IF('Match Day Resource'!$C$104="""","""",'Match Day Resource'!$C$104)"
    AddMD "L|105||STD|OFFICER|105||=IF('Match Day Resource'!$C$105="""","""",'Match Day Resource'!$C$105)"
    AddMD "L|106||STD|OFFICER|106|Force Supplied - Dogs Public Order|=IF('Match Day Resource'!$C$106="""","""",'Match Day Resource'!$C$106)"
    AddMD "L|107||STD|OFFICER|107||=IF('Match Day Resource'!$C$107="""","""",'Match Day Resource'!$C$107)"
    AddMD "L|108||STD|OFFICER|108||=IF('Match Day Resource'!$C$108="""","""",'Match Day Resource'!$C$108)"
    AddMD "L|109||STD|OFFICER|109|Mutual Aid Requested - Dogs Public Order|=IF('Match Day Resource'!$C$109="""","""",'Match Day Resource'!$C$109)"
    AddMD "L|110||STD|OFFICER|110||=IF('Match Day Resource'!$C$110="""","""",'Match Day Resource'!$C$110)"
    AddMD "L|111||STD|OFFICER|111||=IF('Match Day Resource'!$C$111="""","""",'Match Day Resource'!$C$111)"
    AddMD "L|112||STD|OFFICER|112|Force Supplied - Dogs Explo|=IF('Match Day Resource'!$C$112="""","""",'Match Day Resource'!$C$112)"
    AddMD "L|113||STD|OFFICER|113||=IF('Match Day Resource'!$C$113="""","""",'Match Day Resource'!$C$113)"
    AddMD "L|114||STD|OFFICER|114||=IF('Match Day Resource'!$C$114="""","""",'Match Day Resource'!$C$114)"
    AddMD "L|115||STD|OFFICER|115|Mutual Aid Requested -Dogs Explo|=IF('Match Day Resource'!$C$115="""","""",'Match Day Resource'!$C$115)"
    AddMD "L|116||STD|OFFICER|116||=IF('Match Day Resource'!$C$116="""","""",'Match Day Resource'!$C$116)"
    AddMD "L|117||STD|OFFICER|117||=IF('Match Day Resource'!$C$117="""","""",'Match Day Resource'!$C$117)"
    AddMD "S|Football Policing Portfolio"
    AddMD "L|120||STD|OFFICER|120|Force Supplied - Football Spotting DFO|=IF('Match Day Resource'!$C$120="""","""",'Match Day Resource'!$C$120)"
    AddMD "L|121||STD|OFFICER|121||=IF('Match Day Resource'!$C$121="""","""",'Match Day Resource'!$C$121)"
    AddMD "L|122||STD|OFFICER|122|Mutual Aid Requested -  Football Spotting DFO|=IF('Match Day Resource'!$C$122="""","""",'Match Day Resource'!$C$122)"
    AddMD "L|123||STD|OFFICER|123||=IF('Match Day Resource'!$C$123="""","""",'Match Day Resource'!$C$123)"
    AddMD "L|124||STD|OFFICER|124|Force Supplied - Football Spotting OFO|=IF('Match Day Resource'!$C$124="""","""",'Match Day Resource'!$C$124)"
    AddMD "L|125||STD|OFFICER|125||=IF('Match Day Resource'!$C$125="""","""",'Match Day Resource'!$C$125)"
    AddMD "L|126||STD|OFFICER|126|Mutual Aid Requested -Football Spotting OFO|=IF('Match Day Resource'!$C$126="""","""",'Match Day Resource'!$C$126)"
    AddMD "L|127||STD|OFFICER|127||=IF('Match Day Resource'!$C$127="""","""",'Match Day Resource'!$C$127)"
    AddMD "S|Roads Policing"
    AddMD "L|130||STD|OFFICER|130|Force Supplied - Roads Policing Officer|=IF('Match Day Resource'!$C$130="""","""",'Match Day Resource'!$C$130)"
    AddMD "L|131||STD|OFFICER|131||=IF('Match Day Resource'!$C$131="""","""",'Match Day Resource'!$C$131)"
    AddMD "L|132||STD|OFFICER|132|Mutual Aid Requested - Roads Policing Officer|=IF('Match Day Resource'!$C$132="""","""",'Match Day Resource'!$C$132)"
    AddMD "L|133||STD|OFFICER|133||=IF('Match Day Resource'!$C$133="""","""",'Match Day Resource'!$C$133)"
    AddMD "L|134||STD|OFFICER|134|Force Supplied - Roads Policing Officer - Motorcyclists|=IF('Match Day Resource'!$C$134="""","""",'Match Day Resource'!$C$134)"
    AddMD "L|135||STD|OFFICER|135||=IF('Match Day Resource'!$C$135="""","""",'Match Day Resource'!$C$135)"
    AddMD "L|136||STD|OFFICER|136|Mutual Aid Requested - Roads Policing Officer - Motorcyclists|=IF('Match Day Resource'!$C$136="""","""",'Match Day Resource'!$C$136)"
    AddMD "L|137||STD|OFFICER|137||=IF('Match Day Resource'!$C$137="""","""",'Match Day Resource'!$C$137)"
    AddMD "L|138||STD|OFFICER|138|Force Supplied- Roads Policing Officer - VIP Motorcyclist escort|=IF('Match Day Resource'!$C$138="""","""",'Match Day Resource'!$C$138)"
    AddMD "L|139||STD|OFFICER|139||=IF('Match Day Resource'!$C$139="""","""",'Match Day Resource'!$C$139)"
    AddMD "L|140||STD|OFFICER|140|Mutual Aid Requested - Roads Policing Officer - VIP Motorcyclist escort|=IF('Match Day Resource'!$C$140="""","""",'Match Day Resource'!$C$140)"
    AddMD "L|141||STD|OFFICER|141||=IF('Match Day Resource'!$C$141="""","""",'Match Day Resource'!$C$141)"
    AddMD "S|Armed Policing"
    AddMD "L|144||STD|OFFICER|144|Force Supplied- Operational Firearms Commander|=IF('Match Day Resource'!$C$144="""","""",'Match Day Resource'!$C$144)"
    AddMD "L|145||STD|OFFICER|145||=IF('Match Day Resource'!$C$145="""","""",'Match Day Resource'!$C$145)"
    AddMD "L|146||STD|OFFICER|146||=IF('Match Day Resource'!$C$146="""","""",'Match Day Resource'!$C$146)"
    AddMD "L|147||STD|OFFICER|147|Mutual Aid Requested -  Operational Firearms Commander|=IF('Match Day Resource'!$C$147="""","""",'Match Day Resource'!$C$147)"
    AddMD "L|148||STD|OFFICER|148||=IF('Match Day Resource'!$C$148="""","""",'Match Day Resource'!$C$148)"
    AddMD "L|149||STD|OFFICER|149||=IF('Match Day Resource'!$C$149="""","""",'Match Day Resource'!$C$149)"
    AddMD "L|150||STD|OFFICER|150|Force Supplied- Armed Response (ARVs)|=IF('Match Day Resource'!$C$150="""","""",'Match Day Resource'!$C$150)"
    AddMD "L|151||STD|OFFICER|151||=IF('Match Day Resource'!$C$151="""","""",'Match Day Resource'!$C$151)"
    AddMD "L|152||STD|OFFICER|152||=IF('Match Day Resource'!$C$152="""","""",'Match Day Resource'!$C$152)"
    AddMD "L|153||STD|OFFICER|153|Mutual Aid Requested - Armed Response (ARVs)|=IF('Match Day Resource'!$C$153="""","""",'Match Day Resource'!$C$153)"
    AddMD "L|154||STD|OFFICER|154||=IF('Match Day Resource'!$C$154="""","""",'Match Day Resource'!$C$154)"
    AddMD "L|155||STD|OFFICER|155||=IF('Match Day Resource'!$C$155="""","""",'Match Day Resource'!$C$155)"
    AddMD "L|156||STD|OFFICER|156|Force Supplied- CTSFOs|=IF('Match Day Resource'!$C$156="""","""",'Match Day Resource'!$C$156)"
    AddMD "L|157||STD|OFFICER|157||=IF('Match Day Resource'!$C$157="""","""",'Match Day Resource'!$C$157)"
    AddMD "L|158||STD|OFFICER|158||=IF('Match Day Resource'!$C$158="""","""",'Match Day Resource'!$C$158)"
    AddMD "L|159||STD|OFFICER|159|Mutual Aid Requested - CTSFOs|=IF('Match Day Resource'!$C$159="""","""",'Match Day Resource'!$C$159)"
    AddMD "L|160||STD|OFFICER|160||=IF('Match Day Resource'!$C$160="""","""",'Match Day Resource'!$C$160)"
    AddMD "L|161||STD|OFFICER|161||=IF('Match Day Resource'!$C$161="""","""",'Match Day Resource'!$C$161)"
    AddMD "L|162||STD|OFFICER|162|Force Supplied- AFOs|=IF('Match Day Resource'!$C$162="""","""",'Match Day Resource'!$C$162)"
    AddMD "L|163||STD|OFFICER|163||=IF('Match Day Resource'!$C$163="""","""",'Match Day Resource'!$C$163)"
    AddMD "L|164||STD|OFFICER|164||=IF('Match Day Resource'!$C$164="""","""",'Match Day Resource'!$C$164)"
    AddMD "L|165||STD|OFFICER|165|Mutual Aid Requested - AFOs|=IF('Match Day Resource'!$C$165="""","""",'Match Day Resource'!$C$165)"
    AddMD "L|166||STD|OFFICER|166||=IF('Match Day Resource'!$C$166="""","""",'Match Day Resource'!$C$166)"
    AddMD "L|167||STD|OFFICER|167||=IF('Match Day Resource'!$C$167="""","""",'Match Day Resource'!$C$167)"
    AddMD "L|168||STD|OFFICER|168|Force Supplied- Rifle Officer|=IF('Match Day Resource'!$C$168="""","""",'Match Day Resource'!$C$168)"
    AddMD "L|169||STD|OFFICER|169||=IF('Match Day Resource'!$C$169="""","""",'Match Day Resource'!$C$169)"
    AddMD "L|170||STD|OFFICER|170|Mutual Aid Requested - Rifle Officer|=IF('Match Day Resource'!$C$170="""","""",'Match Day Resource'!$C$170)"
End Sub

Private Sub DataMD2()
    Dim s As String
    AddMD "L|171||STD|OFFICER|171||=IF('Match Day Resource'!$C$171="""","""",'Match Day Resource'!$C$171)"
    AddMD "L|172||STD|OFFICER|172|Force Supplied- Protection Officer|=IF('Match Day Resource'!$C$172="""","""",'Match Day Resource'!$C$172)"
    AddMD "L|173||STD|OFFICER|173||=IF('Match Day Resource'!$C$173="""","""",'Match Day Resource'!$C$173)"
    AddMD "L|174||STD|OFFICER|174|Mutual Aid Requested - Protection Officer|=IF('Match Day Resource'!$C$174="""","""",'Match Day Resource'!$C$174)"
    AddMD "L|175||STD|OFFICER|175||=IF('Match Day Resource'!$C$175="""","""",'Match Day Resource'!$C$175)"
    AddMD "S|Additional Armed Policing"
    AddMD "L|177||STD|OFFICER|177|Force Supplied - Armed Response Overlay|=IF('Match Day Resource'!$C$177="""","""",'Match Day Resource'!$C$177)"
    AddMD "L|178||STD|OFFICER|178||=IF('Match Day Resource'!$C$178="""","""",'Match Day Resource'!$C$178)"
    AddMD "L|179||STD|OFFICER|179||=IF('Match Day Resource'!$C$179="""","""",'Match Day Resource'!$C$179)"
    AddMD "L|180||STD|OFFICER|180||=IF('Match Day Resource'!$C$180="""","""",'Match Day Resource'!$C$180)"
    AddMD "L|181||STD|OFFICER|181||=IF('Match Day Resource'!$C$181="""","""",'Match Day Resource'!$C$181)"
    AddMD "L|182||STD|OFFICER|182||=IF('Match Day Resource'!$C$182="""","""",'Match Day Resource'!$C$182)"
    AddMD "L|183||STD|OFFICER|183||=IF('Match Day Resource'!$C$183="""","""",'Match Day Resource'!$C$183)"
    AddMD "L|184||STD|OFFICER|184|Mutual Aid Requested - Armed Response Overlay|=IF('Match Day Resource'!$C$184="""","""",'Match Day Resource'!$C$184)"
    AddMD "L|185||STD|OFFICER|185||=IF('Match Day Resource'!$C$185="""","""",'Match Day Resource'!$C$185)"
    AddMD "L|186||STD|OFFICER|186||=IF('Match Day Resource'!$C$186="""","""",'Match Day Resource'!$C$186)"
    AddMD "L|187||STD|OFFICER|187||=IF('Match Day Resource'!$C$187="""","""",'Match Day Resource'!$C$187)"
    AddMD "L|188||STD|OFFICER|188||=IF('Match Day Resource'!$C$188="""","""",'Match Day Resource'!$C$188)"
    AddMD "L|189||STD|OFFICER|189||=IF('Match Day Resource'!$C$189="""","""",'Match Day Resource'!$C$189)"
    AddMD "L|190||STD|OFFICER|190||=IF('Match Day Resource'!$C$190="""","""",'Match Day Resource'!$C$190)"
    AddMD "S|Disruptive Effects"
    AddMD "L|193||STD|OFFICER|193|Force Supplied - Servator Disruptive Effects|=IF('Match Day Resource'!$C$193="""","""",'Match Day Resource'!$C$193)"
    AddMD "L|194||STD|OFFICER|194||=IF('Match Day Resource'!$C$194="""","""",'Match Day Resource'!$C$194)"
    AddMD "L|195||STD|OFFICER|195|Mutual Aid - Servator Disruptive Effects|=IF('Match Day Resource'!$C$195="""","""",'Match Day Resource'!$C$195)"
    AddMD "L|196||STD|OFFICER|196||=IF('Match Day Resource'!$C$196="""","""",'Match Day Resource'!$C$196)"
    AddMD "L|197||STD|OFFICER|197|Force Supplied - Servator Disruptive Effects Operational Commander|=IF('Match Day Resource'!$C$197="""","""",'Match Day Resource'!$C$197)"
    AddMD "L|198||STD|OFFICER|198||=IF('Match Day Resource'!$C$198="""","""",'Match Day Resource'!$C$198)"
    AddMD "L|199||STD|OFFICER|199|Mutual Aid - Servator Disruptive Effects Operational Commander|=IF('Match Day Resource'!$C$199="""","""",'Match Day Resource'!$C$199)"
    AddMD "L|200||STD|OFFICER|200||=IF('Match Day Resource'!$C$200="""","""",'Match Day Resource'!$C$200)"
    AddMD "L|201||STD|OFFICER|201|Force Supplied - Servator Disruptive Effects Tactical Commander|=IF('Match Day Resource'!$C$201="""","""",'Match Day Resource'!$C$201)"
    AddMD "L|202||STD|OFFICER|202||=IF('Match Day Resource'!$C$202="""","""",'Match Day Resource'!$C$202)"
    AddMD "L|203||STD|OFFICER|203|Mutual Aid - Servator Disruptive Effects Tactical Commander|=IF('Match Day Resource'!$C$203="""","""",'Match Day Resource'!$C$203)"
    AddMD "L|204||STD|OFFICER|204||=IF('Match Day Resource'!$C$204="""","""",'Match Day Resource'!$C$204)"
    AddMD "S|Search"
    AddMD "L|207||STD|OFFICER|207|Force Supplied - LSOs|=IF('Match Day Resource'!$C$207="""","""",'Match Day Resource'!$C$207)"
    AddMD "L|208||STD|OFFICER|208||=IF('Match Day Resource'!$C$208="""","""",'Match Day Resource'!$C$208)"
    AddMD "L|209||STD|OFFICER|209||=IF('Match Day Resource'!$C$209="""","""",'Match Day Resource'!$C$209)"
    AddMD "L|210||STD|OFFICER|210|Mutual Aid Requested - LSOs|=IF('Match Day Resource'!$C$210="""","""",'Match Day Resource'!$C$210)"
    AddMD "L|211||STD|OFFICER|211||=IF('Match Day Resource'!$C$211="""","""",'Match Day Resource'!$C$211)"
    AddMD "L|212||STD|OFFICER|212||=IF('Match Day Resource'!$C$212="""","""",'Match Day Resource'!$C$212)"
    AddMD "L|213||STD|OFFICER|213|Force Supplied - PolSAs|=IF('Match Day Resource'!$C$213="""","""",'Match Day Resource'!$C$213)"
    AddMD "L|214||STD|OFFICER|214||=IF('Match Day Resource'!$C$214="""","""",'Match Day Resource'!$C$214)"
    AddMD "L|215||STD|OFFICER|215||=IF('Match Day Resource'!$C$215="""","""",'Match Day Resource'!$C$215)"
    AddMD "L|216||STD|OFFICER|216|Mutual Aid Requested - PolSAs|=IF('Match Day Resource'!$C$216="""","""",'Match Day Resource'!$C$216)"
    AddMD "L|217||STD|OFFICER|217||=IF('Match Day Resource'!$C$217="""","""",'Match Day Resource'!$C$217)"
    AddMD "L|218||STD|OFFICER|218||=IF('Match Day Resource'!$C$218="""","""",'Match Day Resource'!$C$218)"
    AddMD "L|219||STD|OFFICER|219|Force Supplied - CT Secco|=IF('Match Day Resource'!$C$219="""","""",'Match Day Resource'!$C$219)"
    AddMD "L|220||STD|OFFICER|220||=IF('Match Day Resource'!$C$220="""","""",'Match Day Resource'!$C$220)"
    AddMD "L|221||STD|OFFICER|221||=IF('Match Day Resource'!$C$221="""","""",'Match Day Resource'!$C$221)"
    AddMD "L|222||STD|OFFICER|222|Mutual Aid Requested - CT Secco|=IF('Match Day Resource'!$C$222="""","""",'Match Day Resource'!$C$222)"
    AddMD "L|223||STD|OFFICER|223||=IF('Match Day Resource'!$C$223="""","""",'Match Day Resource'!$C$223)"
    AddMD "L|224||STD|OFFICER|224||=IF('Match Day Resource'!$C$224="""","""",'Match Day Resource'!$C$224)"
    AddMD "S|Unmanned Aerial System (Drones)"
    AddMD "L|227||STD|OFFICER|227|Force Supplied - Counter UAS Bronze Commander|=IF('Match Day Resource'!$C$227="""","""",'Match Day Resource'!$C$227)"
    AddMD "L|228||STD|OFFICER|228||=IF('Match Day Resource'!$C$228="""","""",'Match Day Resource'!$C$228)"
    AddMD "L|229||STD|OFFICER|229||=IF('Match Day Resource'!$C$229="""","""",'Match Day Resource'!$C$229)"
    AddMD "L|230||STD|OFFICER|230|Mutual Aid Requested - Counter UAS Bronze Commander|=IF('Match Day Resource'!$C$230="""","""",'Match Day Resource'!$C$230)"
    AddMD "L|231||STD|OFFICER|231||=IF('Match Day Resource'!$C$231="""","""",'Match Day Resource'!$C$231)"
    AddMD "L|232||STD|OFFICER|232||=IF('Match Day Resource'!$C$232="""","""",'Match Day Resource'!$C$232)"
    AddMD "L|233||STD|OFFICER|233|Force Supplied - Counter UAS Effector Operator|=IF('Match Day Resource'!$C$233="""","""",'Match Day Resource'!$C$233)"
    AddMD "L|234||STD|OFFICER|234||=IF('Match Day Resource'!$C$234="""","""",'Match Day Resource'!$C$234)"
    AddMD "L|235||STD|OFFICER|235||=IF('Match Day Resource'!$C$235="""","""",'Match Day Resource'!$C$235)"
    AddMD "L|236||STD|OFFICER|236|Mutual Aid Requested - Counter UAS Effector Operator|=IF('Match Day Resource'!$C$236="""","""",'Match Day Resource'!$C$236)"
    AddMD "L|237||STD|OFFICER|237||=IF('Match Day Resource'!$C$237="""","""",'Match Day Resource'!$C$237)"
    AddMD "L|238||STD|OFFICER|238||=IF('Match Day Resource'!$C$238="""","""",'Match Day Resource'!$C$238)"
    AddMD "L|239||STD|OFFICER|239|Force Supplied - Counter UAS Observers|=IF('Match Day Resource'!$C$239="""","""",'Match Day Resource'!$C$239)"
    AddMD "L|240||STD|OFFICER|240||=IF('Match Day Resource'!$C$240="""","""",'Match Day Resource'!$C$240)"
    AddMD "L|241||STD|OFFICER|241||=IF('Match Day Resource'!$C$241="""","""",'Match Day Resource'!$C$241)"
    AddMD "L|242||STD|OFFICER|242|Mutual Aid Requested - Counter UAS Observers|=IF('Match Day Resource'!$C$242="""","""",'Match Day Resource'!$C$242)"
    AddMD "L|243||STD|OFFICER|243||=IF('Match Day Resource'!$C$243="""","""",'Match Day Resource'!$C$243)"
    AddMD "L|244||STD|OFFICER|244||=IF('Match Day Resource'!$C$244="""","""",'Match Day Resource'!$C$244)"
    AddMD "L|245||STD|OFFICER|245|Force Supplied - Counter Tactical Advisor|=IF('Match Day Resource'!$C$245="""","""",'Match Day Resource'!$C$245)"
    AddMD "L|246||STD|OFFICER|246||=IF('Match Day Resource'!$C$246="""","""",'Match Day Resource'!$C$246)"
    AddMD "L|247||STD|OFFICER|247||=IF('Match Day Resource'!$C$247="""","""",'Match Day Resource'!$C$247)"
    AddMD "L|248||STD|OFFICER|248|Mutual Aid Requested - Counter Tactical Advisor|=IF('Match Day Resource'!$C$248="""","""",'Match Day Resource'!$C$248)"
    AddMD "L|249||STD|OFFICER|249||=IF('Match Day Resource'!$C$249="""","""",'Match Day Resource'!$C$249)"
    AddMD "L|250||STD|OFFICER|250||=IF('Match Day Resource'!$C$250="""","""",'Match Day Resource'!$C$250)"
    AddMD "L|251||STD|OFFICER|251|Force Supplied - Drone|=IF('Match Day Resource'!$C$251="""","""",'Match Day Resource'!$C$251)"
    AddMD "L|252||STD|OFFICER|252||=IF('Match Day Resource'!$C$252="""","""",'Match Day Resource'!$C$252)"
    AddMD "L|253||STD|OFFICER|253||=IF('Match Day Resource'!$C$253="""","""",'Match Day Resource'!$C$253)"
    AddMD "L|254||STD|OFFICER|254|Mutual Aid Requested - Drone|=IF('Match Day Resource'!$C$254="""","""",'Match Day Resource'!$C$254)"
    AddMD "L|255||STD|OFFICER|255||=IF('Match Day Resource'!$C$255="""","""",'Match Day Resource'!$C$255)"
    AddMD "L|256||STD|OFFICER|256||=IF('Match Day Resource'!$C$256="""","""",'Match Day Resource'!$C$256)"
    AddMD "S|Live Facial Recognition"
    AddMD "L|259||STD|OFFICER|259|Force Supplied - LFR Operational Commander|=IF('Match Day Resource'!$C$259="""","""",'Match Day Resource'!$C$259)"
    AddMD "L|260||STD|OFFICER|260||=IF('Match Day Resource'!$C$260="""","""",'Match Day Resource'!$C$260)"
    AddMD "L|261||STD|OFFICER|261||=IF('Match Day Resource'!$C$261="""","""",'Match Day Resource'!$C$261)"
    AddMD "L|262||STD|OFFICER|262|Mutual Aid Requested - LFR Operational Commander|=IF('Match Day Resource'!$C$262="""","""",'Match Day Resource'!$C$262)"
    AddMD "L|263||STD|OFFICER|263||=IF('Match Day Resource'!$C$263="""","""",'Match Day Resource'!$C$263)"
    AddMD "L|264||STD|OFFICER|264||=IF('Match Day Resource'!$C$264="""","""",'Match Day Resource'!$C$264)"
    AddMD "L|265||STD|OFFICER|265|Force Supplied - LFR Operator|=IF('Match Day Resource'!$C$265="""","""",'Match Day Resource'!$C$265)"
    AddMD "L|266||STD|OFFICER|266||=IF('Match Day Resource'!$C$266="""","""",'Match Day Resource'!$C$266)"
    AddMD "L|267||STD|OFFICER|267||=IF('Match Day Resource'!$C$267="""","""",'Match Day Resource'!$C$267)"
    AddMD "L|268||STD|OFFICER|268|Mutual Aid Requested - LFR Operator|=IF('Match Day Resource'!$C$268="""","""",'Match Day Resource'!$C$268)"
    AddMD "L|269||STD|OFFICER|269||=IF('Match Day Resource'!$C$269="""","""",'Match Day Resource'!$C$269)"
    AddMD "L|270||STD|OFFICER|270||=IF('Match Day Resource'!$C$270="""","""",'Match Day Resource'!$C$270)"
    AddMD "S|CBRN Portfolio"
    AddMD "L|273||STD|OFFICER|273|Force Supplied - General Responder|=IF('Match Day Resource'!$C$273="""","""",'Match Day Resource'!$C$273)"
    AddMD "L|274||STD|OFFICER|274||=IF('Match Day Resource'!$C$274="""","""",'Match Day Resource'!$C$274)"
    AddMD "L|275||STD|OFFICER|275||=IF('Match Day Resource'!$C$275="""","""",'Match Day Resource'!$C$275)"
    AddMD "L|276||STD|OFFICER|276|Mutual Aid Requested - General Responder|=IF('Match Day Resource'!$C$276="""","""",'Match Day Resource'!$C$276)"
    AddMD "L|277||STD|OFFICER|277||=IF('Match Day Resource'!$C$277="""","""",'Match Day Resource'!$C$277)"
    AddMD "L|278||STD|OFFICER|278||=IF('Match Day Resource'!$C$278="""","""",'Match Day Resource'!$C$278)"
    AddMD "L|279||STD|OFFICER|279|Force Supplied - CBRN Officer(s)|=IF('Match Day Resource'!$C$279="""","""",'Match Day Resource'!$C$279)"
    AddMD "L|280||STD|OFFICER|280||=IF('Match Day Resource'!$C$280="""","""",'Match Day Resource'!$C$280)"
    AddMD "L|281||STD|OFFICER|281||=IF('Match Day Resource'!$C$281="""","""",'Match Day Resource'!$C$281)"
    AddMD "L|282||STD|OFFICER|282||=IF('Match Day Resource'!$C$282="""","""",'Match Day Resource'!$C$282)"
    AddMD "L|283||STD|OFFICER|283||=IF('Match Day Resource'!$C$283="""","""",'Match Day Resource'!$C$283)"
    AddMD "L|284||STD|OFFICER|284|Mutual Aid Requested - CBRN Officer(s)|=IF('Match Day Resource'!$C$284="""","""",'Match Day Resource'!$C$284)"
    AddMD "L|285||STD|OFFICER|285||=IF('Match Day Resource'!$C$285="""","""",'Match Day Resource'!$C$285)"
    AddMD "L|286||STD|OFFICER|286||=IF('Match Day Resource'!$C$286="""","""",'Match Day Resource'!$C$286)"
    AddMD "L|287||STD|OFFICER|287||=IF('Match Day Resource'!$C$287="""","""",'Match Day Resource'!$C$287)"
    AddMD "L|288||STD|OFFICER|288||=IF('Match Day Resource'!$C$288="""","""",'Match Day Resource'!$C$288)"
    AddMD "S|Non Public Order Force Command"
    AddMD "L|291||STD|OFFICER|291|Force Supplied - Command|=IF('Match Day Resource'!$C$291="""","""",'Match Day Resource'!$C$291)"
    AddMD "L|292||STD|OFFICER|292||=IF('Match Day Resource'!$C$292="""","""",'Match Day Resource'!$C$292)"
    AddMD "L|293||STD|OFFICER|293||=IF('Match Day Resource'!$C$293="""","""",'Match Day Resource'!$C$293)"
    AddMD "L|294||STD|OFFICER|294||=IF('Match Day Resource'!$C$294="""","""",'Match Day Resource'!$C$294)"
    AddMD "L|295||STD|OFFICER|295||=IF('Match Day Resource'!$C$295="""","""",'Match Day Resource'!$C$295)"
    AddMD "L|296||STD|OFFICER|296||=IF('Match Day Resource'!$C$296="""","""",'Match Day Resource'!$C$296)"
    AddMD "L|297||STD|OFFICER|297||=IF('Match Day Resource'!$C$297="""","""",'Match Day Resource'!$C$297)"
    AddMD "S|Control Room"
    AddMD "L|300||STD|OFFICER|300|Force Supplied-  Control Room|=IF('Match Day Resource'!$C$300="""","""",'Match Day Resource'!$C$300)"
    AddMD "L|301||STD|OFFICER|301||=IF('Match Day Resource'!$C$301="""","""",'Match Day Resource'!$C$301)"
    AddMD "L|302||STD|OFFICER|302||=IF('Match Day Resource'!$C$302="""","""",'Match Day Resource'!$C$302)"
    AddMD "S|Custody"
    AddMD "L|305||STD|OFFICER|305|Force Supplied - Custody|=IF('Match Day Resource'!$C$305="""","""",'Match Day Resource'!$C$305)"
    AddMD "L|306||STD|OFFICER|306||=IF('Match Day Resource'!$C$306="""","""",'Match Day Resource'!$C$306)"
    AddMD "S|Crime"
    AddMD "L|309||STD|OFFICER|309|Force Supplied - Crime|=IF('Match Day Resource'!$C$309="""","""",'Match Day Resource'!$C$309)"
    AddMD "L|310||STD|OFFICER|310||=IF('Match Day Resource'!$C$310="""","""",'Match Day Resource'!$C$310)"
    AddMD "L|311||STD|OFFICER|311||=IF('Match Day Resource'!$C$311="""","""",'Match Day Resource'!$C$311)"
    AddMD "L|312||STD|OFFICER|312|Mutual Aid Requested - Crime|=IF('Match Day Resource'!$C$312="""","""",'Match Day Resource'!$C$312)"
    AddMD "L|313||STD|OFFICER|313||=IF('Match Day Resource'!$C$313="""","""",'Match Day Resource'!$C$313)"
    AddMD "L|314||STD|OFFICER|314||=IF('Match Day Resource'!$C$314="""","""",'Match Day Resource'!$C$314)"
    AddMD "S|Intel"
    AddMD "L|317||STD|OFFICER|317|Force Supplied - Intel|=IF('Match Day Resource'!$C$317="""","""",'Match Day Resource'!$C$317)"
    AddMD "L|318||STD|OFFICER|318||=IF('Match Day Resource'!$C$318="""","""",'Match Day Resource'!$C$318)"
    AddMD "L|319||STD|OFFICER|319||=IF('Match Day Resource'!$C$319="""","""",'Match Day Resource'!$C$319)"
    AddMD "L|320||STD|OFFICER|320|Mutual Aid Requested - Intel|=IF('Match Day Resource'!$C$320="""","""",'Match Day Resource'!$C$320)"
    AddMD "L|321||STD|OFFICER|321||=IF('Match Day Resource'!$C$321="""","""",'Match Day Resource'!$C$321)"
    AddMD "L|322||STD|OFFICER|322||=IF('Match Day Resource'!$C$322="""","""",'Match Day Resource'!$C$322)"
    AddMD "S|Ancillary"
    AddMD "L|325||STD|OFFICER|325|Force Supplied - Ancillary Team|=IF('Match Day Resource'!$C$325="""","""",'Match Day Resource'!$C$325)"
    AddMD "L|326||STD|OFFICER|326||=IF('Match Day Resource'!$C$326="""","""",'Match Day Resource'!$C$326)"
    AddMD "L|327||STD|OFFICER|327|Mutual Aid Requested - Ancillary Team|=IF('Match Day Resource'!$C$327="""","""",'Match Day Resource'!$C$327)"
    AddMD "L|328||STD|OFFICER|328||=IF('Match Day Resource'!$C$328="""","""",'Match Day Resource'!$C$328)"
    AddMD "S|Additional Resources"
    AddMD "L|331||STD|OFFICER|331|Additional Resources|=IF('Match Day Resource'!$C$331="""","""",'Match Day Resource'!$C$331)"
    AddMD "L|332||STD|OFFICER|332||=IF('Match Day Resource'!$C$332="""","""",'Match Day Resource'!$C$332)"
End Sub

Private Sub DataMD3()
    Dim s As String
    AddMD "L|333||STD|OFFICER|333||=IF('Match Day Resource'!$C$333="""","""",'Match Day Resource'!$C$333)"
    AddMD "L|334||STD|OFFICER|334||=IF('Match Day Resource'!$C$334="""","""",'Match Day Resource'!$C$334)"
    AddMD "L|335||STD|OFFICER|335||=IF('Match Day Resource'!$C$335="""","""",'Match Day Resource'!$C$335)"
    AddMD "L|336||STD|OFFICER|336||=IF('Match Day Resource'!$C$336="""","""",'Match Day Resource'!$C$336)"
    AddMD "L|337||STD|OFFICER|337||=IF('Match Day Resource'!$C$337="""","""",'Match Day Resource'!$C$337)"
    AddMD "L|338||STD|OFFICER|338||=IF('Match Day Resource'!$C$338="""","""",'Match Day Resource'!$C$338)"
    AddMD "L|339||STD|OFFICER|339||=IF('Match Day Resource'!$C$339="""","""",'Match Day Resource'!$C$339)"
End Sub

Private Sub DataWT1()
    Dim s As String
    AddWT "S|Public Order / Public Safety"
    AddWT "L|20|6|PSU|6|Force Supplied - Wider Community Public Order Policing|='Wider Tournament Resource'!$C$20&"" - Inspector"""
    AddWT "L|20|7|PSU|7||='Wider Tournament Resource'!$C$20&"" - Sergeant"""
    AddWT "L|20|8|PSU|8||='Wider Tournament Resource'!$C$20&"" - Constable"""
    AddWT "L|21|9|PSU|9||='Wider Tournament Resource'!$C$21&"" - Inspector"""
    AddWT "L|21|10|PSU|10||='Wider Tournament Resource'!$C$21&"" - Sergeant"""
    AddWT "L|21|11|PSU|11||='Wider Tournament Resource'!$C$21&"" - Constable"""
    AddWT "L|22||MAPSU|22|Mutual Aid Requested - Wider Community Public Order Policing|='Wider Tournament Resource'!$C$22"
    AddWT "L|23||MAPSU|23||='Wider Tournament Resource'!$C$23"
    AddWT "L|24|12|PSU|12|Force Supplied - Fan Zone Public Order Policing|='Wider Tournament Resource'!$C$24&"" - Inspector"""
    AddWT "L|24|13|PSU|13||='Wider Tournament Resource'!$C$24&"" - Sergeant"""
    AddWT "L|24|14|PSU|14||='Wider Tournament Resource'!$C$24&"" - Constable"""
    AddWT "L|25|15|PSU|15||='Wider Tournament Resource'!$C$25&"" - Inspector"""
    AddWT "L|25|16|PSU|16||='Wider Tournament Resource'!$C$25&"" - Sergeant"""
    AddWT "L|25|17|PSU|17||='Wider Tournament Resource'!$C$25&"" - Constable"""
    AddWT "L|26||MAPSU|26|Mutual Aid Requested - Fan Zone Public Order Policing|='Wider Tournament Resource'!$C$26"
    AddWT "L|27||MAPSU|27||='Wider Tournament Resource'!$C$27"
    AddWT "L|28||OFFICER|28|Force Supplied - Operationally Deployable non-Public Order Policing|=IF('Wider Tournament Resource'!$C$28="""","""",'Wider Tournament Resource'!$C$28)"
    AddWT "L|29||OFFICER|29||=IF('Wider Tournament Resource'!$C$29="""","""",'Wider Tournament Resource'!$C$29)"
    AddWT "L|30||OFFICER|30||=IF('Wider Tournament Resource'!$C$30="""","""",'Wider Tournament Resource'!$C$30)"
    AddWT "L|31||OFFICER|31|Mutual Aid Requested - Operationally Deployable non-Public Order Policing|=IF('Wider Tournament Resource'!$C$31="""","""",'Wider Tournament Resource'!$C$31)"
    AddWT "L|32||OFFICER|32||=IF('Wider Tournament Resource'!$C$32="""","""",'Wider Tournament Resource'!$C$32)"
    AddWT "L|33||OFFICER|33||=IF('Wider Tournament Resource'!$C$33="""","""",'Wider Tournament Resource'!$C$33)"
    AddWT "L|34||OFFICER|34|Force Supplied - Evidence GT's|=IF('Wider Tournament Resource'!$C$34="""","""",'Wider Tournament Resource'!$C$34)"
    AddWT "L|35||OFFICER|35||=IF('Wider Tournament Resource'!$C$35="""","""",'Wider Tournament Resource'!$C$35)"
    AddWT "L|36||OFFICER|36|Mutual Aid Requested - Evidence Gatherers|=IF('Wider Tournament Resource'!$C$36="""","""",'Wider Tournament Resource'!$C$36)"
    AddWT "L|37||OFFICER|37||=IF('Wider Tournament Resource'!$C$37="""","""",'Wider Tournament Resource'!$C$37)"
    AddWT "S|Additional POPS Resources"
    AddWT "L|39||OFFICER|39|Force Supplied - Additional Public Order|=IF('Wider Tournament Resource'!$C$39="""","""",'Wider Tournament Resource'!$C$39)"
    AddWT "L|40||OFFICER|40||=IF('Wider Tournament Resource'!$C$40="""","""",'Wider Tournament Resource'!$C$40)"
    AddWT "L|41||OFFICER|41||=IF('Wider Tournament Resource'!$C$41="""","""",'Wider Tournament Resource'!$C$41)"
    AddWT "L|42||OFFICER|42||=IF('Wider Tournament Resource'!$C$42="""","""",'Wider Tournament Resource'!$C$42)"
    AddWT "L|43||OFFICER|43||=IF('Wider Tournament Resource'!$C$43="""","""",'Wider Tournament Resource'!$C$43)"
    AddWT "L|44||OFFICER|44||=IF('Wider Tournament Resource'!$C$44="""","""",'Wider Tournament Resource'!$C$44)"
    AddWT "L|45||OFFICER|45||=IF('Wider Tournament Resource'!$C$45="""","""",'Wider Tournament Resource'!$C$45)"
    AddWT "L|46||OFFICER|46||=IF('Wider Tournament Resource'!$C$46="""","""",'Wider Tournament Resource'!$C$46)"
    AddWT "L|47||OFFICER|47||=IF('Wider Tournament Resource'!$C$47="""","""",'Wider Tournament Resource'!$C$47)"
    AddWT "L|48||OFFICER|48|Mutual Aid Requested -   Additional PO|=IF('Wider Tournament Resource'!$C$48="""","""",'Wider Tournament Resource'!$C$48)"
    AddWT "L|49||OFFICER|49||=IF('Wider Tournament Resource'!$C$49="""","""",'Wider Tournament Resource'!$C$49)"
    AddWT "L|50||OFFICER|50||=IF('Wider Tournament Resource'!$C$50="""","""",'Wider Tournament Resource'!$C$50)"
    AddWT "L|51||OFFICER|51||=IF('Wider Tournament Resource'!$C$51="""","""",'Wider Tournament Resource'!$C$51)"
    AddWT "L|52||OFFICER|52||=IF('Wider Tournament Resource'!$C$52="""","""",'Wider Tournament Resource'!$C$52)"
    AddWT "L|53||OFFICER|53||=IF('Wider Tournament Resource'!$C$53="""","""",'Wider Tournament Resource'!$C$53)"
    AddWT "L|54||OFFICER|54||=IF('Wider Tournament Resource'!$C$54="""","""",'Wider Tournament Resource'!$C$54)"
    AddWT "L|55||OFFICER|55||=IF('Wider Tournament Resource'!$C$55="""","""",'Wider Tournament Resource'!$C$55)"
    AddWT "L|56||OFFICER|56||=IF('Wider Tournament Resource'!$C$56="""","""",'Wider Tournament Resource'!$C$56)"
    AddWT "S|Fan Zone - Disruptive Effects"
    AddWT "L|59||OFFICER|59|Force Supplied - Disruptive Effects Officer|=IF('Wider Tournament Resource'!$C$59="""","""",'Wider Tournament Resource'!$C$59)"
    AddWT "L|60||OFFICER|60||=IF('Wider Tournament Resource'!$C$60="""","""",'Wider Tournament Resource'!$C$60)"
    AddWT "L|61||OFFICER|61|Mutual Aid -  Disruptive Effects Officer|=IF('Wider Tournament Resource'!$C$61="""","""",'Wider Tournament Resource'!$C$61)"
    AddWT "L|62||OFFICER|62||=IF('Wider Tournament Resource'!$C$62="""","""",'Wider Tournament Resource'!$C$62)"
    AddWT "L|63||OFFICER|63|Force Supplied - Disruptive Effects Operational Commander|=IF('Wider Tournament Resource'!$C$63="""","""",'Wider Tournament Resource'!$C$63)"
    AddWT "L|64||OFFICER|64||=IF('Wider Tournament Resource'!$C$64="""","""",'Wider Tournament Resource'!$C$64)"
    AddWT "L|65||OFFICER|65|Mutual Aid - Disruptive Effects Operational Commander|=IF('Wider Tournament Resource'!$C$65="""","""",'Wider Tournament Resource'!$C$65)"
    AddWT "L|66||OFFICER|66||=IF('Wider Tournament Resource'!$C$66="""","""",'Wider Tournament Resource'!$C$66)"
    AddWT "L|67||OFFICER|67|Force Supplied - Disruptive Effects Tactical Commander|=IF('Wider Tournament Resource'!$C$67="""","""",'Wider Tournament Resource'!$C$67)"
    AddWT "L|68||OFFICER|68||=IF('Wider Tournament Resource'!$C$68="""","""",'Wider Tournament Resource'!$C$68)"
    AddWT "L|69||OFFICER|69|Mutual Aid - Disruptive Effects Tactical Commander|=IF('Wider Tournament Resource'!$C$69="""","""",'Wider Tournament Resource'!$C$69)"
    AddWT "L|70||OFFICER|70||=IF('Wider Tournament Resource'!$C$70="""","""",'Wider Tournament Resource'!$C$70)"
    AddWT "S|Fan Zone Armed Policing"
    AddWT "L|73||OFFICER|73|Force Supplied- Fan Zone Armed Response (ARVs)|=IF('Wider Tournament Resource'!$C$73="""","""",'Wider Tournament Resource'!$C$73)"
    AddWT "L|74||OFFICER|74||=IF('Wider Tournament Resource'!$C$74="""","""",'Wider Tournament Resource'!$C$74)"
    AddWT "L|75||OFFICER|75|Mutual Aid Requested- Fan Zone Armed Response (ARVs)|=IF('Wider Tournament Resource'!$C$75="""","""",'Wider Tournament Resource'!$C$75)"
    AddWT "L|76||OFFICER|76||=IF('Wider Tournament Resource'!$C$76="""","""",'Wider Tournament Resource'!$C$76)"
    AddWT "S|Wider Tournament Policing"
    AddWT "L|79||OFFICER|79|Force Supplied - Control Room|=IF('Wider Tournament Resource'!$C$79="""","""",'Wider Tournament Resource'!$C$79)"
    AddWT "L|80||OFFICER|80||=IF('Wider Tournament Resource'!$C$80="""","""",'Wider Tournament Resource'!$C$80)"
    AddWT "L|81||OFFICER|81|Force Supplied - Command Support|=IF('Wider Tournament Resource'!$C$81="""","""",'Wider Tournament Resource'!$C$81)"
    AddWT "L|82||OFFICER|82||=IF('Wider Tournament Resource'!$C$82="""","""",'Wider Tournament Resource'!$C$82)"
    AddWT "L|83||OFFICER|83||=IF('Wider Tournament Resource'!$C$83="""","""",'Wider Tournament Resource'!$C$83)"
    AddWT "L|84||OFFICER|84||=IF('Wider Tournament Resource'!$C$84="""","""",'Wider Tournament Resource'!$C$84)"
    AddWT "L|85||OFFICER|85|Mutual Aid Requested - Command Support|=IF('Wider Tournament Resource'!$C$85="""","""",'Wider Tournament Resource'!$C$85)"
    AddWT "L|86||OFFICER|86||=IF('Wider Tournament Resource'!$C$86="""","""",'Wider Tournament Resource'!$C$86)"
    AddWT "L|87||OFFICER|87||=IF('Wider Tournament Resource'!$C$87="""","""",'Wider Tournament Resource'!$C$87)"
    AddWT "L|88||OFFICER|88||=IF('Wider Tournament Resource'!$C$88="""","""",'Wider Tournament Resource'!$C$88)"
    AddWT "L|89||OFFICER|89|Force Supplied - Custody|=IF('Wider Tournament Resource'!$C$89="""","""",'Wider Tournament Resource'!$C$89)"
    AddWT "L|90||OFFICER|90||=IF('Wider Tournament Resource'!$C$90="""","""",'Wider Tournament Resource'!$C$90)"
    AddWT "S|Additional Resources"
    AddWT "L|93||OFFICER|93|Additional Resources|=IF('Wider Tournament Resource'!$C$93="""","""",'Wider Tournament Resource'!$C$93)"
    AddWT "L|94||OFFICER|94||=IF('Wider Tournament Resource'!$C$94="""","""",'Wider Tournament Resource'!$C$94)"
    AddWT "L|95||OFFICER|95||=IF('Wider Tournament Resource'!$C$95="""","""",'Wider Tournament Resource'!$C$95)"
    AddWT "L|96||OFFICER|96||=IF('Wider Tournament Resource'!$C$96="""","""",'Wider Tournament Resource'!$C$96)"
    AddWT "L|97||OFFICER|97||=IF('Wider Tournament Resource'!$C$97="""","""",'Wider Tournament Resource'!$C$97)"
    AddWT "L|98||OFFICER|98||=IF('Wider Tournament Resource'!$C$98="""","""",'Wider Tournament Resource'!$C$98)"
    AddWT "L|99||OFFICER|99||=IF('Wider Tournament Resource'!$C$99="""","""",'Wider Tournament Resource'!$C$99)"
    AddWT "L|100||OFFICER|100||=IF('Wider Tournament Resource'!$C$100="""","""",'Wider Tournament Resource'!$C$100)"
    AddWT "L|101||OFFICER|101||=IF('Wider Tournament Resource'!$C$101="""","""",'Wider Tournament Resource'!$C$101)"
End Sub

Private Sub DataCC1()
    Dim s As String
    AddCC "R|MD|7|7|Force Supplied - Level 2 PSU / Serials (Non-Matchday)|=IF('Match Day Resource'!$C$7="""","""",'Match Day Resource'!$C$7)|ZERO||0"
    AddCC "R|MD|8|8|Force Supplied - Level 2 PSU / Serials (Non-Matchday)|=IF('Match Day Resource'!$C$8="""","""",'Match Day Resource'!$C$8)|RATE|'Baseline Costs'!$X$31|0"
    AddCC "R|MD|9|9|Force Supplied - Level 2 PSU / Serials (Non-Matchday)|=IF('Match Day Resource'!$C$9="""","""",'Match Day Resource'!$C$9)|RATE|'Baseline Costs'!$X$32|0"
    AddCC "R|MD|10|10|Force Supplied - Level 3 BDU / Serials (Non-Matchday)|=IF('Match Day Resource'!$C$10="""","""",'Match Day Resource'!$C$10)|ZERO||0"
    AddCC "R|MD|11|11|Force Supplied - Level 3 BDU / Serials (Non-Matchday)|=IF('Match Day Resource'!$C$11="""","""",'Match Day Resource'!$C$11)|RATE|'Baseline Costs'!$X$31|0"
    AddCC "R|MD|12|12|Force Supplied - Level 3 BDU / Serials (Non-Matchday)|=IF('Match Day Resource'!$C$12="""","""",'Match Day Resource'!$C$12)|RATE|'Baseline Costs'!$X$32|0"
    AddCC "R|MD|13|13|Force Supplied - Level 2 PSU / Serials Town Centre Public Order Policing|=IF('Match Day Resource'!$C$13="""","""",'Match Day Resource'!$C$13)|ZERO||0"
    AddCC "R|MD|14|14|Force Supplied - Level 2 PSU / Serials Town Centre Public Order Policing|=IF('Match Day Resource'!$C$14="""","""",'Match Day Resource'!$C$14)|RATE|'Baseline Costs'!$X$31|0"
    AddCC "R|MD|15|15|Force Supplied - Level 2 PSU / Serials Town Centre Public Order Policing|=IF('Match Day Resource'!$C$15="""","""",'Match Day Resource'!$C$15)|RATE|'Baseline Costs'!$X$32|0"
    AddCC "R|MD|16|16|Force Supplied - Level 3 BDU / Serials Town Centre Public Order Policing|=IF('Match Day Resource'!$C$16="""","""",'Match Day Resource'!$C$16)|ZERO||0"
    AddCC "R|MD|17|17|Force Supplied - Level 3 BDU / Serials Town Centre Public Order Policing|=IF('Match Day Resource'!$C$17="""","""",'Match Day Resource'!$C$17)|RATE|'Baseline Costs'!$X$31|0"
    AddCC "R|MD|18|18|Force Supplied - Level 3 BDU / Serials Town Centre Public Order Policing|=IF('Match Day Resource'!$C$18="""","""",'Match Day Resource'!$C$18)|RATE|'Baseline Costs'!$X$32|0"
    AddCC "R|MD|19|19|Force Supplied - Level 2 PSU / Serials Stadium Vicinity Public Order Policing|=IF('Match Day Resource'!$C$19="""","""",'Match Day Resource'!$C$19)|ZERO||0"
    AddCC "R|MD|20|20|Force Supplied - Level 2 PSU / Serials Stadium Vicinity Public Order Policing|=IF('Match Day Resource'!$C$20="""","""",'Match Day Resource'!$C$20)|RATE|'Baseline Costs'!$X$31|0"
    AddCC "R|MD|21|21|Force Supplied - Level 2 PSU / Serials Stadium Vicinity Public Order Policing|=IF('Match Day Resource'!$C$21="""","""",'Match Day Resource'!$C$21)|RATE|'Baseline Costs'!$X$32|0"
    AddCC "R|MD|22|22|Force Supplied - Level 3 BDU / Serials Stadium Vicinity Public Order Policing|=IF('Match Day Resource'!$C$22="""","""",'Match Day Resource'!$C$22)|ZERO||0"
    AddCC "R|MD|23|23|Force Supplied - Level 3 BDU / Serials Stadium Vicinity Public Order Policing|=IF('Match Day Resource'!$C$23="""","""",'Match Day Resource'!$C$23)|RATE|'Baseline Costs'!$X$31|0"
    AddCC "R|MD|24|24|Force Supplied - Level 3 BDU / Serials Stadium Vicinity Public Order Policing|=IF('Match Day Resource'!$C$24="""","""",'Match Day Resource'!$C$24)|RATE|'Baseline Costs'!$X$32|0"
    AddCC "R|MD|25|25|Force Supplied - Level 2 PSU / Serials In Stadium Public Order Policing|=IF('Match Day Resource'!$C$25="""","""",'Match Day Resource'!$C$25)|RATE|'Baseline Costs'!$X$13|0"
    AddCC "R|MD|26|26|Force Supplied - Level 2 PSU / Serials In Stadium Public Order Policing|=IF('Match Day Resource'!$C$26="""","""",'Match Day Resource'!$C$26)|RATE|'Baseline Costs'!$X$31|0"
    AddCC "R|MD|27|27|Force Supplied - Level 2 PSU / Serials In Stadium Public Order Policing|=IF('Match Day Resource'!$C$27="""","""",'Match Day Resource'!$C$27)|RATE|'Baseline Costs'!$X$32|0"
    AddCC "R|MD|28|28|Force Supplied - Level 3 BDU / Serials In Stadium Public Order Policing|=IF('Match Day Resource'!$C$28="""","""",'Match Day Resource'!$C$28)|RATE|'Baseline Costs'!$X$13|0"
    AddCC "R|MD|29|29|Force Supplied - Level 3 BDU / Serials In Stadium Public Order Policing|=IF('Match Day Resource'!$C$29="""","""",'Match Day Resource'!$C$29)|RATE|'Baseline Costs'!$X$31|0"
    AddCC "R|MD|30|30|Force Supplied - Level 3 BDU / Serials In Stadium Public Order Policing|=IF('Match Day Resource'!$C$30="""","""",'Match Day Resource'!$C$30)|RATE|'Baseline Costs'!$X$32|0"
    AddCC "R|MD|33|33|Force Supplied - Public Order (Non-Matchday)|=IF('Match Day Resource'!$C$33="""","""",'Match Day Resource'!$C$33)|SKIP||0"
    AddCC "R|MD|34|34|Force Supplied - Public Order (Non-Matchday)|=IF('Match Day Resource'!$C$34="""","""",'Match Day Resource'!$C$34)|SKIP||0"
    AddCC "R|MD|35|35|Mutual Aid Requested - Public Order (Non-Matchday)|=IF('Match Day Resource'!$C$35="""","""",'Match Day Resource'!$C$35)|PKG||1"
    AddCC "R|MD|36|36|Mutual Aid Requested - Public Order (Non-Matchday)|=IF('Match Day Resource'!$C$36="""","""",'Match Day Resource'!$C$36)|PKG||1"
    AddCC "R|MD|37|37|Force Supplied - Town Centre Public Order Policing (Matchday Only)|=IF('Match Day Resource'!$C$37="""","""",'Match Day Resource'!$C$37)|SKIP||0"
    AddCC "R|MD|38|38|Force Supplied - Town Centre Public Order Policing (Matchday Only)|=IF('Match Day Resource'!$C$38="""","""",'Match Day Resource'!$C$38)|SKIP||0"
    AddCC "R|MD|39|39|Mutual Aid Requested - Town Centre Public Order Policing (Matchday Only)|=IF('Match Day Resource'!$C$39="""","""",'Match Day Resource'!$C$39)|PKG||1"
    AddCC "R|MD|40|40|Mutual Aid Requested - Town Centre Public Order Policing (Matchday Only)|=IF('Match Day Resource'!$C$40="""","""",'Match Day Resource'!$C$40)|PKG||1"
    AddCC "R|MD|41|41|Force Supplied - Stadium Vicinity Public Order Policing (Matchday Only)|=IF('Match Day Resource'!$C$41="""","""",'Match Day Resource'!$C$41)|SKIP||0"
    AddCC "R|MD|42|42|Force Supplied - Stadium Vicinity Public Order Policing (Matchday Only)|=IF('Match Day Resource'!$C$42="""","""",'Match Day Resource'!$C$42)|SKIP||0"
    AddCC "R|MD|43|43|Mutual Aid Requested - Stadium Vicinity Public Order Policing (Matchday Only)|=IF('Match Day Resource'!$C$43="""","""",'Match Day Resource'!$C$43)|PKG||1"
    AddCC "R|MD|44|44|Mutual Aid Requested - Stadium Vicinity Public Order Policing (Matchday Only)|=IF('Match Day Resource'!$C$44="""","""",'Match Day Resource'!$C$44)|PKG||1"
    AddCC "R|MD|45|45|Force Supplied - In Stadium Public Order Policing (Matchday Only)|=IF('Match Day Resource'!$C$45="""","""",'Match Day Resource'!$C$45)|SKIP||0"
    AddCC "R|MD|46|46|Force Supplied - In Stadium Public Order Policing (Matchday Only)|=IF('Match Day Resource'!$C$46="""","""",'Match Day Resource'!$C$46)|SKIP||0"
    AddCC "R|MD|47|47|Mutual Aid Requested - In Stadium Public Order Policing (Matchday Only)|=IF('Match Day Resource'!$C$47="""","""",'Match Day Resource'!$C$47)|PKG||1"
    AddCC "R|MD|48|48|Mutual Aid Requested - In Stadium Public Order Policing (Matchday Only)|=IF('Match Day Resource'!$C$48="""","""",'Match Day Resource'!$C$48)|PKG||1"
    AddCC "R|MD|50|50|Force Supplied - Evidence Gathering Team Officer|=IF('Match Day Resource'!$C$50="""","""",'Match Day Resource'!$C$50)|RATE|'Baseline Costs'!$X$31|1"
    AddCC "R|MD|51|51|Force Supplied - Evidence Gathering Team Officer|=IF('Match Day Resource'!$C$51="""","""",'Match Day Resource'!$C$51)|RATE|'Baseline Costs'!$X$32|1"
    AddCC "R|MD|52|52|Mutual Aid Requested - Evidence Gathering Team Officer|=IF('Match Day Resource'!$C$52="""","""",'Match Day Resource'!$C$52)|RATE|'Baseline Costs'!$X$14|1"
    AddCC "R|MD|53|53|Mutual Aid Requested - Evidence Gathering Team Officer|=IF('Match Day Resource'!$C$53="""","""",'Match Day Resource'!$C$53)|RATE|'Baseline Costs'!$X$15|1"
    AddCC "R|MD|55|55|Force Supplied - Working at Height Complex Protest Removal|=IF('Match Day Resource'!$C$55="""","""",'Match Day Resource'!$C$55)|ZERO||1"
    AddCC "R|MD|56|56|Force Supplied - Working at Height Complex Protest Removal|=IF('Match Day Resource'!$C$56="""","""",'Match Day Resource'!$C$56)|RATE|'Baseline Costs'!$X$31|1"
    AddCC "R|MD|57|57|Force Supplied - Working at Height Complex Protest Removal|=IF('Match Day Resource'!$C$57="""","""",'Match Day Resource'!$C$57)|RATE|'Baseline Costs'!$X$32|1"
    AddCC "R|MD|58|58|Mutual Aid Requested - Working at Height Complex Protest Removal|=IF('Match Day Resource'!$C$58="""","""",'Match Day Resource'!$C$58)|RATE|'Baseline Costs'!$X$13|1"
    AddCC "R|MD|59|59|Mutual Aid Requested - Working at Height Complex Protest Removal|=IF('Match Day Resource'!$C$59="""","""",'Match Day Resource'!$C$59)|RATE|'Baseline Costs'!$X$14|1"
    AddCC "R|MD|60|60|Mutual Aid Requested - Working at Height Complex Protest Removal|=IF('Match Day Resource'!$C$60="""","""",'Match Day Resource'!$C$60)|RATE|'Baseline Costs'!$X$15|1"
    AddCC "R|MD|61|61|Force Supplied - Non-Complex Protest Removal|=IF('Match Day Resource'!$C$61="""","""",'Match Day Resource'!$C$61)|ZERO||1"
    AddCC "R|MD|62|62|Force Supplied - Non-Complex Protest Removal|=IF('Match Day Resource'!$C$62="""","""",'Match Day Resource'!$C$62)|RATE|'Baseline Costs'!$X$31|1"
    AddCC "R|MD|63|63|Force Supplied - Non-Complex Protest Removal|=IF('Match Day Resource'!$C$63="""","""",'Match Day Resource'!$C$63)|RATE|'Baseline Costs'!$X$32|1"
    AddCC "R|MD|64|64|Mutual Aid Requested - Non-Complex Protest Removal|=IF('Match Day Resource'!$C$64="""","""",'Match Day Resource'!$C$64)|RATE|'Baseline Costs'!$X$13|1"
    AddCC "R|MD|65|65|Mutual Aid Requested - Non-Complex Protest Removal|=IF('Match Day Resource'!$C$65="""","""",'Match Day Resource'!$C$65)|RATE|'Baseline Costs'!$X$14|1"
    AddCC "R|MD|66|66|Mutual Aid Requested - Non-Complex Protest Removal|=IF('Match Day Resource'!$C$66="""","""",'Match Day Resource'!$C$66)|RATE|'Baseline Costs'!$X$15|1"
    AddCC "R|MD|67|67|Force Supplied - Debonding Protest Removal|=IF('Match Day Resource'!$C$67="""","""",'Match Day Resource'!$C$67)|ZERO||1"
    AddCC "R|MD|68|68|Force Supplied - Debonding Protest Removal|=IF('Match Day Resource'!$C$68="""","""",'Match Day Resource'!$C$68)|RATE|'Baseline Costs'!$X$31|1"
    AddCC "R|MD|69|69|Force Supplied - Debonding Protest Removal|=IF('Match Day Resource'!$C$69="""","""",'Match Day Resource'!$C$69)|RATE|'Baseline Costs'!$X$32|1"
    AddCC "R|MD|70|70|Mutual Aid Requested - Debonding Protest Removal|=IF('Match Day Resource'!$C$70="""","""",'Match Day Resource'!$C$70)|RATE|'Baseline Costs'!$X$13|1"
    AddCC "R|MD|71|71|Mutual Aid Requested - Debonding Protest Removal|=IF('Match Day Resource'!$C$71="""","""",'Match Day Resource'!$C$71)|RATE|'Baseline Costs'!$X$14|1"
    AddCC "R|MD|72|72|Mutual Aid Requested - Debonding Protest Removal|=IF('Match Day Resource'!$C$72="""","""",'Match Day Resource'!$C$72)|RATE|'Baseline Costs'!$X$15|1"
    AddCC "R|MD|74|74|Force Supplied - Additional Public Order|=IF('Match Day Resource'!$C$74="""","""",'Match Day Resource'!$C$74)|RATE|'Baseline Costs'!$X$14|1"
    AddCC "R|MD|75|75|Force Supplied - Additional Public Order|=IF('Match Day Resource'!$C$75="""","""",'Match Day Resource'!$C$75)|RATE|'Baseline Costs'!$X$14|1"
    AddCC "R|MD|76|76|Force Supplied - Additional Public Order|=IF('Match Day Resource'!$C$76="""","""",'Match Day Resource'!$C$76)|RATE|'Baseline Costs'!$X$14|1"
    AddCC "R|MD|77|77|Force Supplied - Additional Public Order|=IF('Match Day Resource'!$C$77="""","""",'Match Day Resource'!$C$77)|RATE|'Baseline Costs'!$X$14|1"
    AddCC "R|MD|78|78|Force Supplied - Additional Public Order|=IF('Match Day Resource'!$C$78="""","""",'Match Day Resource'!$C$78)|RATE|'Baseline Costs'!$X$14|1"
    AddCC "R|MD|79|79|Force Supplied - Additional Public Order|=IF('Match Day Resource'!$C$79="""","""",'Match Day Resource'!$C$79)|RATE|'Baseline Costs'!$X$14|1"
    AddCC "R|MD|80|80|Mutual Aid Requested -   Additional PO|=IF('Match Day Resource'!$C$80="""","""",'Match Day Resource'!$C$80)|RATE|'Baseline Costs'!$X$14|1"
    AddCC "R|MD|81|81|Mutual Aid Requested -   Additional PO|=IF('Match Day Resource'!$C$81="""","""",'Match Day Resource'!$C$81)|RATE|'Baseline Costs'!$X$14|1"
    AddCC "R|MD|82|82|Mutual Aid Requested -   Additional PO|=IF('Match Day Resource'!$C$82="""","""",'Match Day Resource'!$C$82)|RATE|'Baseline Costs'!$X$14|1"
    AddCC "R|MD|83|83|Mutual Aid Requested -   Additional PO|=IF('Match Day Resource'!$C$83="""","""",'Match Day Resource'!$C$83)|RATE|'Baseline Costs'!$X$14|1"
    AddCC "R|MD|84|84|Mutual Aid Requested -   Additional PO|=IF('Match Day Resource'!$C$84="""","""",'Match Day Resource'!$C$84)|RATE|'Baseline Costs'!$X$14|1"
    AddCC "R|MD|85|85|Mutual Aid Requested -   Additional PO|=IF('Match Day Resource'!$C$85="""","""",'Match Day Resource'!$C$85)|RATE|'Baseline Costs'!$X$14|1"
    AddCC "R|MD|88|88|Force Supplied - Mounted Commander|=IF('Match Day Resource'!$C$88="""","""",'Match Day Resource'!$C$88)|RATE|IFERROR(VLOOKUP('Match Day Resource'!$C$88,'Baseline Costs'!$B$37:$D$39,3,FALSE),0)|1"
    AddCC "R|MD|89|89|Mutual Aid Requested -   Mounted Commander|=IF('Match Day Resource'!$C$89="""","""",'Match Day Resource'!$C$89)|RATE|IFERROR(VLOOKUP('Match Day Resource'!$C$89,'Baseline Costs'!$B$37:$D$39,4,FALSE),0)|1"
    AddCC "R|MD|90|90|Force Supplied - Mounted Officer|=IF('Match Day Resource'!$C$90="""","""",'Match Day Resource'!$C$90)|RATE|IFERROR(VLOOKUP('Match Day Resource'!$C$90,'Baseline Costs'!$B$37:$D$39,3,FALSE),0)|1"
    AddCC "R|MD|91|91|Mutual Aid Requested -   Mounted Officer|=IF('Match Day Resource'!$C$91="""","""",'Match Day Resource'!$C$91)|RATE|IFERROR(VLOOKUP('Match Day Resource'!$C$91,'Baseline Costs'!$B$37:$D$39,4,FALSE),0)|1"
    AddCC "R|MD|94|94|Force Supplied - Dogs GDP|=IF('Match Day Resource'!$C$94="""","""",'Match Day Resource'!$C$94)|RATE|IFERROR(VLOOKUP('Match Day Resource'!$C$94,'Baseline Costs'!$B$37:$D$39,3,FALSE),0)|1"
    AddCC "R|MD|95|95|Force Supplied - Dogs GDP|=IF('Match Day Resource'!$C$95="""","""",'Match Day Resource'!$C$95)|RATE|IFERROR(VLOOKUP('Match Day Resource'!$C$95,'Baseline Costs'!$B$37:$D$39,3,FALSE),0)|1"
    AddCC "R|MD|96|96|Force Supplied - Dogs GDP|=IF('Match Day Resource'!$C$96="""","""",'Match Day Resource'!$C$96)|RATE|IFERROR(VLOOKUP('Match Day Resource'!$C$96,'Baseline Costs'!$B$37:$D$39,3,FALSE),0)|1"
    AddCC "R|MD|97|97|Mutual Aid Requested- Dogs GDP|=IF('Match Day Resource'!$C$97="""","""",'Match Day Resource'!$C$97)|RATE|IFERROR(VLOOKUP('Match Day Resource'!$C$97,'Baseline Costs'!$B$37:$D$39,4,FALSE),0)|1"
    AddCC "R|MD|98|98|Mutual Aid Requested- Dogs GDP|=IF('Match Day Resource'!$C$98="""","""",'Match Day Resource'!$C$98)|RATE|IFERROR(VLOOKUP('Match Day Resource'!$C$98,'Baseline Costs'!$B$37:$D$39,4,FALSE),0)|1"
    AddCC "R|MD|99|99|Mutual Aid Requested- Dogs GDP|=IF('Match Day Resource'!$C$99="""","""",'Match Day Resource'!$C$99)|RATE|IFERROR(VLOOKUP('Match Day Resource'!$C$99,'Baseline Costs'!$B$37:$D$39,4,FALSE),0)|1"
    AddCC "R|MD|100|100|Force Supplied - Dogs Public Order Serial Commander|=IF('Match Day Resource'!$C$100="""","""",'Match Day Resource'!$C$100)|RATE|IFERROR(VLOOKUP('Match Day Resource'!$C$100,'Baseline Costs'!$B$37:$D$39,3,FALSE),0)|1"
    AddCC "R|MD|101|101|Force Supplied - Dogs Public Order Serial Commander|=IF('Match Day Resource'!$C$101="""","""",'Match Day Resource'!$C$101)|RATE|IFERROR(VLOOKUP('Match Day Resource'!$C$101,'Baseline Costs'!$B$37:$D$39,3,FALSE),0)|1"
    AddCC "R|MD|102|102|Force Supplied - Dogs Public Order Serial Commander|=IF('Match Day Resource'!$C$102="""","""",'Match Day Resource'!$C$102)|RATE|IFERROR(VLOOKUP('Match Day Resource'!$C$102,'Baseline Costs'!$B$37:$D$39,3,FALSE),0)|1"
    AddCC "R|MD|103|103|Mutual Aid Requested - Dogs Public Order Serial Commander|=IF('Match Day Resource'!$C$103="""","""",'Match Day Resource'!$C$103)|RATE|IFERROR(VLOOKUP('Match Day Resource'!$C$103,'Baseline Costs'!$B$37:$D$39,4,FALSE),0)|1"
    AddCC "R|MD|104|104|Mutual Aid Requested - Dogs Public Order Serial Commander|=IF('Match Day Resource'!$C$104="""","""",'Match Day Resource'!$C$104)|RATE|IFERROR(VLOOKUP('Match Day Resource'!$C$104,'Baseline Costs'!$B$37:$D$39,4,FALSE),0)|1"
    AddCC "R|MD|105|105|Mutual Aid Requested - Dogs Public Order Serial Commander|=IF('Match Day Resource'!$C$105="""","""",'Match Day Resource'!$C$105)|RATE|IFERROR(VLOOKUP('Match Day Resource'!$C$105,'Baseline Costs'!$B$37:$D$39,4,FALSE),0)|1"
    AddCC "R|MD|106|106|Force Supplied - Dogs Public Order|=IF('Match Day Resource'!$C$106="""","""",'Match Day Resource'!$C$106)|RATE|IFERROR(VLOOKUP('Match Day Resource'!$C$106,'Baseline Costs'!$B$37:$D$39,3,FALSE),0)|1"
    AddCC "R|MD|107|107|Force Supplied - Dogs Public Order|=IF('Match Day Resource'!$C$107="""","""",'Match Day Resource'!$C$107)|RATE|IFERROR(VLOOKUP('Match Day Resource'!$C$107,'Baseline Costs'!$B$37:$D$39,3,FALSE),0)|1"
    AddCC "R|MD|108|108|Force Supplied - Dogs Public Order|=IF('Match Day Resource'!$C$108="""","""",'Match Day Resource'!$C$108)|RATE|IFERROR(VLOOKUP('Match Day Resource'!$C$108,'Baseline Costs'!$B$37:$D$39,3,FALSE),0)|1"
    AddCC "R|MD|109|109|Mutual Aid Requested - Dogs Public Order|=IF('Match Day Resource'!$C$109="""","""",'Match Day Resource'!$C$109)|RATE|IFERROR(VLOOKUP('Match Day Resource'!$C$109,'Baseline Costs'!$B$37:$D$39,4,FALSE),0)|1"
    AddCC "R|MD|110|110|Mutual Aid Requested - Dogs Public Order|=IF('Match Day Resource'!$C$110="""","""",'Match Day Resource'!$C$110)|RATE|IFERROR(VLOOKUP('Match Day Resource'!$C$110,'Baseline Costs'!$B$37:$D$39,4,FALSE),0)|1"
    AddCC "R|MD|111|111|Mutual Aid Requested - Dogs Public Order|=IF('Match Day Resource'!$C$111="""","""",'Match Day Resource'!$C$111)|RATE|IFERROR(VLOOKUP('Match Day Resource'!$C$111,'Baseline Costs'!$B$37:$D$39,4,FALSE),0)|1"
    AddCC "R|MD|112|112|Force Supplied - Dogs Explo|=IF('Match Day Resource'!$C$112="""","""",'Match Day Resource'!$C$112)|RATE|IFERROR(VLOOKUP('Match Day Resource'!$C$112,'Baseline Costs'!$B$37:$D$39,3,FALSE),0)|1"
    AddCC "R|MD|113|113|Force Supplied - Dogs Explo|=IF('Match Day Resource'!$C$113="""","""",'Match Day Resource'!$C$113)|RATE|IFERROR(VLOOKUP('Match Day Resource'!$C$113,'Baseline Costs'!$B$37:$D$39,3,FALSE),0)|1"
    AddCC "R|MD|114|114|Force Supplied - Dogs Explo|=IF('Match Day Resource'!$C$114="""","""",'Match Day Resource'!$C$114)|RATE|IFERROR(VLOOKUP('Match Day Resource'!$C$114,'Baseline Costs'!$B$37:$D$39,3,FALSE),0)|1"
    AddCC "R|MD|115|115|Mutual Aid Requested -Dogs Explo|=IF('Match Day Resource'!$C$115="""","""",'Match Day Resource'!$C$115)|RATE|IFERROR(VLOOKUP('Match Day Resource'!$C$115,'Baseline Costs'!$B$37:$D$39,4,FALSE),0)|1"
    AddCC "R|MD|116|116|Mutual Aid Requested -Dogs Explo|=IF('Match Day Resource'!$C$116="""","""",'Match Day Resource'!$C$116)|RATE|IFERROR(VLOOKUP('Match Day Resource'!$C$116,'Baseline Costs'!$B$37:$D$39,4,FALSE),0)|1"
    AddCC "R|MD|117|117|Mutual Aid Requested -Dogs Explo|=IF('Match Day Resource'!$C$117="""","""",'Match Day Resource'!$C$117)|RATE|IFERROR(VLOOKUP('Match Day Resource'!$C$117,'Baseline Costs'!$B$37:$D$39,4,FALSE),0)|1"
    AddCC "R|MD|120|120|Force Supplied - Football Spotting DFO|=IF('Match Day Resource'!$C$120="""","""",'Match Day Resource'!$C$120)|RATE|'Baseline Costs'!$X$31|1"
    AddCC "R|MD|121|121|Force Supplied - Football Spotting DFO|=IF('Match Day Resource'!$C$121="""","""",'Match Day Resource'!$C$121)|RATE|'Baseline Costs'!$X$32|1"
    AddCC "R|MD|122|122|Mutual Aid Requested -  Football Spotting DFO|=IF('Match Day Resource'!$C$122="""","""",'Match Day Resource'!$C$122)|RATE|'Baseline Costs'!$X$14|1"
    AddCC "R|MD|123|123|Mutual Aid Requested -  Football Spotting DFO|=IF('Match Day Resource'!$C$123="""","""",'Match Day Resource'!$C$123)|RATE|'Baseline Costs'!$X$15|1"
    AddCC "R|MD|124|124|Force Supplied - Football Spotting OFO|=IF('Match Day Resource'!$C$124="""","""",'Match Day Resource'!$C$124)|RATE|'Baseline Costs'!$X$31|1"
    AddCC "R|MD|125|125|Force Supplied - Football Spotting OFO|=IF('Match Day Resource'!$C$125="""","""",'Match Day Resource'!$C$125)|RATE|'Baseline Costs'!$X$32|1"
    AddCC "R|MD|126|126|Mutual Aid Requested -Football Spotting OFO|=IF('Match Day Resource'!$C$126="""","""",'Match Day Resource'!$C$126)|RATE|'Baseline Costs'!$X$14|1"
    AddCC "R|MD|127|127|Mutual Aid Requested -Football Spotting OFO|=IF('Match Day Resource'!$C$127="""","""",'Match Day Resource'!$C$127)|RATE|'Baseline Costs'!$X$15|1"
    AddCC "R|MD|130|130|Force Supplied - Roads Policing Officer|=IF('Match Day Resource'!$C$130="""","""",'Match Day Resource'!$C$130)|RATE|IFERROR(VLOOKUP('Match Day Resource'!$C$130,'Baseline Costs'!$B$37:$D$39,3,FALSE),0)|1"
    AddCC "R|MD|131|131|Force Supplied - Roads Policing Officer|=IF('Match Day Resource'!$C$131="""","""",'Match Day Resource'!$C$131)|RATE|IFERROR(VLOOKUP('Match Day Resource'!$C$131,'Baseline Costs'!$B$37:$D$39,3,FALSE),0)|1"
    AddCC "R|MD|132|132|Mutual Aid Requested - Roads Policing Officer|=IF('Match Day Resource'!$C$132="""","""",'Match Day Resource'!$C$132)|RATE|IFERROR(VLOOKUP('Match Day Resource'!$C$132,'Baseline Costs'!$B$37:$D$39,4,FALSE),0)|1"
    AddCC "R|MD|133|133|Mutual Aid Requested - Roads Policing Officer|=IF('Match Day Resource'!$C$133="""","""",'Match Day Resource'!$C$133)|RATE|IFERROR(VLOOKUP('Match Day Resource'!$C$133,'Baseline Costs'!$B$37:$D$39,4,FALSE),0)|1"
    AddCC "R|MD|134|134|Force Supplied - Roads Policing Officer - Motorcyclists|=IF('Match Day Resource'!$C$134="""","""",'Match Day Resource'!$C$134)|RATE|IFERROR(VLOOKUP('Match Day Resource'!$C$134,'Baseline Costs'!$B$37:$D$39,3,FALSE),0)|1"
    AddCC "R|MD|135|135|Force Supplied - Roads Policing Officer - Motorcyclists|=IF('Match Day Resource'!$C$135="""","""",'Match Day Resource'!$C$135)|RATE|IFERROR(VLOOKUP('Match Day Resource'!$C$135,'Baseline Costs'!$B$37:$D$39,3,FALSE),0)|1"
    AddCC "R|MD|136|136|Mutual Aid Requested - Roads Policing Officer - Motorcyclists|=IF('Match Day Resource'!$C$136="""","""",'Match Day Resource'!$C$136)|RATE|IFERROR(VLOOKUP('Match Day Resource'!$C$136,'Baseline Costs'!$B$37:$D$39,4,FALSE),0)|1"
    AddCC "R|MD|137|137|Mutual Aid Requested - Roads Policing Officer - Motorcyclists|=IF('Match Day Resource'!$C$137="""","""",'Match Day Resource'!$C$137)|RATE|IFERROR(VLOOKUP('Match Day Resource'!$C$137,'Baseline Costs'!$B$37:$D$39,4,FALSE),0)|1"
    AddCC "R|MD|138|138|Force Supplied- Roads Policing Officer - VIP Motorcyclist escort|=IF('Match Day Resource'!$C$138="""","""",'Match Day Resource'!$C$138)|RATE|IFERROR(VLOOKUP('Match Day Resource'!$C$138,'Baseline Costs'!$B$37:$D$39,3,FALSE),0)|1"
    AddCC "R|MD|139|139|Force Supplied- Roads Policing Officer - VIP Motorcyclist escort|=IF('Match Day Resource'!$C$139="""","""",'Match Day Resource'!$C$139)|RATE|IFERROR(VLOOKUP('Match Day Resource'!$C$139,'Baseline Costs'!$B$37:$D$39,3,FALSE),0)|1"
    AddCC "R|MD|140|140|Mutual Aid Requested - Roads Policing Officer - VIP Motorcyclist escort|=IF('Match Day Resource'!$C$140="""","""",'Match Day Resource'!$C$140)|RATE|IFERROR(VLOOKUP('Match Day Resource'!$C$140,'Baseline Costs'!$B$37:$D$39,4,FALSE),0)|1"
    AddCC "R|MD|141|141|Mutual Aid Requested - Roads Policing Officer - VIP Motorcyclist escort|=IF('Match Day Resource'!$C$141="""","""",'Match Day Resource'!$C$141)|RATE|IFERROR(VLOOKUP('Match Day Resource'!$C$141,'Baseline Costs'!$B$37:$D$39,4,FALSE),0)|1"
    AddCC "R|MD|144|144|Force Supplied- Operational Firearms Commander|=IF('Match Day Resource'!$C$144="""","""",'Match Day Resource'!$C$144)|ZERO||1"
    AddCC "R|MD|145|145|Force Supplied- Operational Firearms Commander|=IF('Match Day Resource'!$C$145="""","""",'Match Day Resource'!$C$145)|RATE|'Baseline Costs'!$X$31|1"
    AddCC "R|MD|146|146|Force Supplied- Operational Firearms Commander|=IF('Match Day Resource'!$C$146="""","""",'Match Day Resource'!$C$146)|RATE|'Baseline Costs'!$X$32|1"
    AddCC "R|MD|147|147|Mutual Aid Requested -  Operational Firearms Commander|=IF('Match Day Resource'!$C$147="""","""",'Match Day Resource'!$C$147)|RATE|'Baseline Costs'!$X$13|1"
    AddCC "R|MD|148|148|Mutual Aid Requested -  Operational Firearms Commander|=IF('Match Day Resource'!$C$148="""","""",'Match Day Resource'!$C$148)|RATE|'Baseline Costs'!$X$14|1"
    AddCC "R|MD|149|149|Mutual Aid Requested -  Operational Firearms Commander|=IF('Match Day Resource'!$C$149="""","""",'Match Day Resource'!$C$149)|RATE|'Baseline Costs'!$X$15|1"
    AddCC "R|MD|150|150|Force Supplied- Armed Response (ARVs)|=IF('Match Day Resource'!$C$150="""","""",'Match Day Resource'!$C$150)|ZERO||1"
    AddCC "R|MD|151|151|Force Supplied- Armed Response (ARVs)|=IF('Match Day Resource'!$C$151="""","""",'Match Day Resource'!$C$151)|RATE|'Baseline Costs'!$X$31|1"
    AddCC "R|MD|152|152|Force Supplied- Armed Response (ARVs)|=IF('Match Day Resource'!$C$152="""","""",'Match Day Resource'!$C$152)|RATE|'Baseline Costs'!$X$32|1"
    AddCC "R|MD|153|153|Mutual Aid Requested - Armed Response (ARVs)|=IF('Match Day Resource'!$C$153="""","""",'Match Day Resource'!$C$153)|RATE|'Baseline Costs'!$X$13|1"
    AddCC "R|MD|154|154|Mutual Aid Requested - Armed Response (ARVs)|=IF('Match Day Resource'!$C$154="""","""",'Match Day Resource'!$C$154)|RATE|'Baseline Costs'!$X$14|1"
    AddCC "R|MD|155|155|Mutual Aid Requested - Armed Response (ARVs)|=IF('Match Day Resource'!$C$155="""","""",'Match Day Resource'!$C$155)|RATE|'Baseline Costs'!$X$15|1"
    AddCC "R|MD|156|156|Force Supplied- CTSFOs|=IF('Match Day Resource'!$C$156="""","""",'Match Day Resource'!$C$156)|ZERO||1"
    AddCC "R|MD|157|157|Force Supplied- CTSFOs|=IF('Match Day Resource'!$C$157="""","""",'Match Day Resource'!$C$157)|RATE|'Baseline Costs'!$X$31|1"
    AddCC "R|MD|158|158|Force Supplied- CTSFOs|=IF('Match Day Resource'!$C$158="""","""",'Match Day Resource'!$C$158)|RATE|'Baseline Costs'!$X$32|1"
    AddCC "R|MD|159|159|Mutual Aid Requested - CTSFOs|=IF('Match Day Resource'!$C$159="""","""",'Match Day Resource'!$C$159)|RATE|'Baseline Costs'!$X$13|1"
    AddCC "R|MD|160|160|Mutual Aid Requested - CTSFOs|=IF('Match Day Resource'!$C$160="""","""",'Match Day Resource'!$C$160)|RATE|'Baseline Costs'!$X$14|1"
    AddCC "R|MD|161|161|Mutual Aid Requested - CTSFOs|=IF('Match Day Resource'!$C$161="""","""",'Match Day Resource'!$C$161)|RATE|'Baseline Costs'!$X$15|1"
    AddCC "R|MD|162|162|Force Supplied- AFOs|=IF('Match Day Resource'!$C$162="""","""",'Match Day Resource'!$C$162)|ZERO||1"
    AddCC "R|MD|163|163|Force Supplied- AFOs|=IF('Match Day Resource'!$C$163="""","""",'Match Day Resource'!$C$163)|RATE|'Baseline Costs'!$X$31|1"
    AddCC "R|MD|164|164|Force Supplied- AFOs|=IF('Match Day Resource'!$C$164="""","""",'Match Day Resource'!$C$164)|RATE|'Baseline Costs'!$X$32|1"
    AddCC "R|MD|165|165|Mutual Aid Requested - AFOs|=IF('Match Day Resource'!$C$165="""","""",'Match Day Resource'!$C$165)|RATE|'Baseline Costs'!$X$13|1"
    AddCC "R|MD|166|166|Mutual Aid Requested - AFOs|=IF('Match Day Resource'!$C$166="""","""",'Match Day Resource'!$C$166)|RATE|'Baseline Costs'!$X$14|1"
    AddCC "R|MD|167|167|Mutual Aid Requested - AFOs|=IF('Match Day Resource'!$C$167="""","""",'Match Day Resource'!$C$167)|RATE|'Baseline Costs'!$X$15|1"
    AddCC "R|MD|168|168|Force Supplied- Rifle Officer|=IF('Match Day Resource'!$C$168="""","""",'Match Day Resource'!$C$168)|RATE|'Baseline Costs'!$X$31|1"
    AddCC "R|MD|169|169|Force Supplied- Rifle Officer|=IF('Match Day Resource'!$C$169="""","""",'Match Day Resource'!$C$169)|RATE|'Baseline Costs'!$X$32|1"
    AddCC "R|MD|170|170|Mutual Aid Requested - Rifle Officer|=IF('Match Day Resource'!$C$170="""","""",'Match Day Resource'!$C$170)|RATE|'Baseline Costs'!$X$14|1"
    AddCC "R|MD|171|171|Mutual Aid Requested - Rifle Officer|=IF('Match Day Resource'!$C$171="""","""",'Match Day Resource'!$C$171)|RATE|'Baseline Costs'!$X$15|1"
End Sub

Private Sub DataCC2()
    Dim s As String
    AddCC "R|MD|172|172|Force Supplied- Protection Officer|=IF('Match Day Resource'!$C$172="""","""",'Match Day Resource'!$C$172)|RATE|'Baseline Costs'!$X$31|1"
    AddCC "R|MD|173|173|Force Supplied- Protection Officer|=IF('Match Day Resource'!$C$173="""","""",'Match Day Resource'!$C$173)|RATE|'Baseline Costs'!$X$32|1"
    AddCC "R|MD|174|174|Mutual Aid Requested - Protection Officer|=IF('Match Day Resource'!$C$174="""","""",'Match Day Resource'!$C$174)|RATE|'Baseline Costs'!$X$14|1"
    AddCC "R|MD|175|175|Mutual Aid Requested - Protection Officer|=IF('Match Day Resource'!$C$175="""","""",'Match Day Resource'!$C$175)|RATE|'Baseline Costs'!$X$15|1"
    AddCC "R|MD|177|177|Force Supplied - Armed Response Overlay|=IF('Match Day Resource'!$C$177="""","""",'Match Day Resource'!$C$177)|RATE|'Baseline Costs'!$X$14|1"
    AddCC "R|MD|178|178|Force Supplied - Armed Response Overlay|=IF('Match Day Resource'!$C$178="""","""",'Match Day Resource'!$C$178)|RATE|'Baseline Costs'!$X$14|1"
    AddCC "R|MD|179|179|Force Supplied - Armed Response Overlay|=IF('Match Day Resource'!$C$179="""","""",'Match Day Resource'!$C$179)|RATE|'Baseline Costs'!$X$14|1"
    AddCC "R|MD|180|180|Force Supplied - Armed Response Overlay|=IF('Match Day Resource'!$C$180="""","""",'Match Day Resource'!$C$180)|RATE|'Baseline Costs'!$X$14|1"
    AddCC "R|MD|181|181|Force Supplied - Armed Response Overlay|=IF('Match Day Resource'!$C$181="""","""",'Match Day Resource'!$C$181)|RATE|'Baseline Costs'!$X$14|1"
    AddCC "R|MD|182|182|Force Supplied - Armed Response Overlay|=IF('Match Day Resource'!$C$182="""","""",'Match Day Resource'!$C$182)|RATE|'Baseline Costs'!$X$14|1"
    AddCC "R|MD|183|183|Force Supplied - Armed Response Overlay|=IF('Match Day Resource'!$C$183="""","""",'Match Day Resource'!$C$183)|RATE|'Baseline Costs'!$X$14|1"
    AddCC "R|MD|184|184|Mutual Aid Requested - Armed Response Overlay|=IF('Match Day Resource'!$C$184="""","""",'Match Day Resource'!$C$184)|RATE|'Baseline Costs'!$X$14|1"
    AddCC "R|MD|185|185|Mutual Aid Requested - Armed Response Overlay|=IF('Match Day Resource'!$C$185="""","""",'Match Day Resource'!$C$185)|RATE|'Baseline Costs'!$X$14|1"
    AddCC "R|MD|186|186|Mutual Aid Requested - Armed Response Overlay|=IF('Match Day Resource'!$C$186="""","""",'Match Day Resource'!$C$186)|RATE|'Baseline Costs'!$X$14|1"
    AddCC "R|MD|187|187|Mutual Aid Requested - Armed Response Overlay|=IF('Match Day Resource'!$C$187="""","""",'Match Day Resource'!$C$187)|RATE|'Baseline Costs'!$X$14|1"
    AddCC "R|MD|188|188|Mutual Aid Requested - Armed Response Overlay|=IF('Match Day Resource'!$C$188="""","""",'Match Day Resource'!$C$188)|RATE|'Baseline Costs'!$X$14|1"
    AddCC "R|MD|189|189|Mutual Aid Requested - Armed Response Overlay|=IF('Match Day Resource'!$C$189="""","""",'Match Day Resource'!$C$189)|RATE|'Baseline Costs'!$X$14|1"
    AddCC "R|MD|190|190|Mutual Aid Requested - Armed Response Overlay|=IF('Match Day Resource'!$C$190="""","""",'Match Day Resource'!$C$190)|RATE|'Baseline Costs'!$X$14|1"
    AddCC "R|MD|193|193|Force Supplied - Servator Disruptive Effects|=IF('Match Day Resource'!$C$193="""","""",'Match Day Resource'!$C$193)|RATE|'Baseline Costs'!$X$31|1"
    AddCC "R|MD|194|194|Force Supplied - Servator Disruptive Effects|=IF('Match Day Resource'!$C$194="""","""",'Match Day Resource'!$C$194)|RATE|'Baseline Costs'!$X$32|1"
    AddCC "R|MD|195|195|Mutual Aid - Servator Disruptive Effects|=IF('Match Day Resource'!$C$195="""","""",'Match Day Resource'!$C$195)|RATE|'Baseline Costs'!$X$14|1"
    AddCC "R|MD|196|196|Mutual Aid - Servator Disruptive Effects|=IF('Match Day Resource'!$C$196="""","""",'Match Day Resource'!$C$196)|RATE|'Baseline Costs'!$X$15|1"
    AddCC "R|MD|197|197|Force Supplied - Servator Disruptive Effects Operational Commander|=IF('Match Day Resource'!$C$197="""","""",'Match Day Resource'!$C$197)|RATE|'Baseline Costs'!$X$31|1"
    AddCC "R|MD|198|198|Force Supplied - Servator Disruptive Effects Operational Commander|=IF('Match Day Resource'!$C$198="""","""",'Match Day Resource'!$C$198)|RATE|'Baseline Costs'!$X$32|1"
    AddCC "R|MD|199|199|Mutual Aid - Servator Disruptive Effects Operational Commander|=IF('Match Day Resource'!$C$199="""","""",'Match Day Resource'!$C$199)|RATE|'Baseline Costs'!$X$14|1"
    AddCC "R|MD|200|200|Mutual Aid - Servator Disruptive Effects Operational Commander|=IF('Match Day Resource'!$C$200="""","""",'Match Day Resource'!$C$200)|RATE|'Baseline Costs'!$X$15|1"
    AddCC "R|MD|201|201|Force Supplied - Servator Disruptive Effects Tactical Commander|=IF('Match Day Resource'!$C$201="""","""",'Match Day Resource'!$C$201)|RATE|'Baseline Costs'!$X$31|1"
    AddCC "R|MD|202|202|Force Supplied - Servator Disruptive Effects Tactical Commander|=IF('Match Day Resource'!$C$202="""","""",'Match Day Resource'!$C$202)|RATE|'Baseline Costs'!$X$32|1"
    AddCC "R|MD|203|203|Mutual Aid - Servator Disruptive Effects Tactical Commander|=IF('Match Day Resource'!$C$203="""","""",'Match Day Resource'!$C$203)|RATE|'Baseline Costs'!$X$14|1"
    AddCC "R|MD|204|204|Mutual Aid - Servator Disruptive Effects Tactical Commander|=IF('Match Day Resource'!$C$204="""","""",'Match Day Resource'!$C$204)|RATE|'Baseline Costs'!$X$15|1"
    AddCC "R|MD|207|207|Force Supplied - LSOs|=IF('Match Day Resource'!$C$207="""","""",'Match Day Resource'!$C$207)|ZERO||1"
    AddCC "R|MD|208|208|Force Supplied - LSOs|=IF('Match Day Resource'!$C$208="""","""",'Match Day Resource'!$C$208)|RATE|'Baseline Costs'!$X$31|1"
    AddCC "R|MD|209|209|Force Supplied - LSOs|=IF('Match Day Resource'!$C$209="""","""",'Match Day Resource'!$C$209)|RATE|'Baseline Costs'!$X$32|1"
    AddCC "R|MD|210|210|Mutual Aid Requested - LSOs|=IF('Match Day Resource'!$C$210="""","""",'Match Day Resource'!$C$210)|RATE|'Baseline Costs'!$X$13|1"
    AddCC "R|MD|211|211|Mutual Aid Requested - LSOs|=IF('Match Day Resource'!$C$211="""","""",'Match Day Resource'!$C$211)|RATE|'Baseline Costs'!$X$14|1"
    AddCC "R|MD|212|212|Mutual Aid Requested - LSOs|=IF('Match Day Resource'!$C$212="""","""",'Match Day Resource'!$C$212)|RATE|'Baseline Costs'!$X$15|1"
    AddCC "R|MD|213|213|Force Supplied - PolSAs|=IF('Match Day Resource'!$C$213="""","""",'Match Day Resource'!$C$213)|ZERO||1"
    AddCC "R|MD|214|214|Force Supplied - PolSAs|=IF('Match Day Resource'!$C$214="""","""",'Match Day Resource'!$C$214)|RATE|'Baseline Costs'!$X$31|1"
    AddCC "R|MD|215|215|Force Supplied - PolSAs|=IF('Match Day Resource'!$C$215="""","""",'Match Day Resource'!$C$215)|RATE|'Baseline Costs'!$X$32|1"
    AddCC "R|MD|216|216|Mutual Aid Requested - PolSAs|=IF('Match Day Resource'!$C$216="""","""",'Match Day Resource'!$C$216)|RATE|'Baseline Costs'!$X$13|1"
    AddCC "R|MD|217|217|Mutual Aid Requested - PolSAs|=IF('Match Day Resource'!$C$217="""","""",'Match Day Resource'!$C$217)|RATE|'Baseline Costs'!$X$14|1"
    AddCC "R|MD|218|218|Mutual Aid Requested - PolSAs|=IF('Match Day Resource'!$C$218="""","""",'Match Day Resource'!$C$218)|RATE|'Baseline Costs'!$X$15|1"
    AddCC "R|MD|219|219|Force Supplied - CT Secco|=IF('Match Day Resource'!$C$219="""","""",'Match Day Resource'!$C$219)|ZERO||1"
    AddCC "R|MD|220|220|Force Supplied - CT Secco|=IF('Match Day Resource'!$C$220="""","""",'Match Day Resource'!$C$220)|RATE|'Baseline Costs'!$X$31|1"
    AddCC "R|MD|221|221|Force Supplied - CT Secco|=IF('Match Day Resource'!$C$221="""","""",'Match Day Resource'!$C$221)|RATE|'Baseline Costs'!$X$32|1"
    AddCC "R|MD|222|222|Mutual Aid Requested - CT Secco|=IF('Match Day Resource'!$C$222="""","""",'Match Day Resource'!$C$222)|RATE|'Baseline Costs'!$X$13|1"
    AddCC "R|MD|223|223|Mutual Aid Requested - CT Secco|=IF('Match Day Resource'!$C$223="""","""",'Match Day Resource'!$C$223)|RATE|'Baseline Costs'!$X$14|1"
    AddCC "R|MD|224|224|Mutual Aid Requested - CT Secco|=IF('Match Day Resource'!$C$224="""","""",'Match Day Resource'!$C$224)|RATE|'Baseline Costs'!$X$15|1"
    AddCC "R|MD|227|227|Force Supplied - Counter UAS Bronze Commander|=IF('Match Day Resource'!$C$227="""","""",'Match Day Resource'!$C$227)|ZERO||1"
    AddCC "R|MD|228|228|Force Supplied - Counter UAS Bronze Commander|=IF('Match Day Resource'!$C$228="""","""",'Match Day Resource'!$C$228)|RATE|'Baseline Costs'!$X$31|1"
    AddCC "R|MD|229|229|Force Supplied - Counter UAS Bronze Commander|=IF('Match Day Resource'!$C$229="""","""",'Match Day Resource'!$C$229)|RATE|'Baseline Costs'!$X$32|1"
    AddCC "R|MD|230|230|Mutual Aid Requested - Counter UAS Bronze Commander|=IF('Match Day Resource'!$C$230="""","""",'Match Day Resource'!$C$230)|RATE|'Baseline Costs'!$X$13|1"
    AddCC "R|MD|231|231|Mutual Aid Requested - Counter UAS Bronze Commander|=IF('Match Day Resource'!$C$231="""","""",'Match Day Resource'!$C$231)|RATE|'Baseline Costs'!$X$14|1"
    AddCC "R|MD|232|232|Mutual Aid Requested - Counter UAS Bronze Commander|=IF('Match Day Resource'!$C$232="""","""",'Match Day Resource'!$C$232)|RATE|'Baseline Costs'!$X$15|1"
    AddCC "R|MD|233|233|Force Supplied - Counter UAS Effector Operator|=IF('Match Day Resource'!$C$233="""","""",'Match Day Resource'!$C$233)|ZERO||1"
    AddCC "R|MD|234|234|Force Supplied - Counter UAS Effector Operator|=IF('Match Day Resource'!$C$234="""","""",'Match Day Resource'!$C$234)|RATE|'Baseline Costs'!$X$31|1"
    AddCC "R|MD|235|235|Force Supplied - Counter UAS Effector Operator|=IF('Match Day Resource'!$C$235="""","""",'Match Day Resource'!$C$235)|RATE|'Baseline Costs'!$X$32|1"
    AddCC "R|MD|236|236|Mutual Aid Requested - Counter UAS Effector Operator|=IF('Match Day Resource'!$C$236="""","""",'Match Day Resource'!$C$236)|RATE|'Baseline Costs'!$X$13|1"
    AddCC "R|MD|237|237|Mutual Aid Requested - Counter UAS Effector Operator|=IF('Match Day Resource'!$C$237="""","""",'Match Day Resource'!$C$237)|RATE|'Baseline Costs'!$X$14|1"
    AddCC "R|MD|238|238|Mutual Aid Requested - Counter UAS Effector Operator|=IF('Match Day Resource'!$C$238="""","""",'Match Day Resource'!$C$238)|RATE|'Baseline Costs'!$X$15|1"
    AddCC "R|MD|239|239|Force Supplied - Counter UAS Observers|=IF('Match Day Resource'!$C$239="""","""",'Match Day Resource'!$C$239)|ZERO||1"
    AddCC "R|MD|240|240|Force Supplied - Counter UAS Observers|=IF('Match Day Resource'!$C$240="""","""",'Match Day Resource'!$C$240)|RATE|'Baseline Costs'!$X$31|1"
    AddCC "R|MD|241|241|Force Supplied - Counter UAS Observers|=IF('Match Day Resource'!$C$241="""","""",'Match Day Resource'!$C$241)|RATE|'Baseline Costs'!$X$32|1"
    AddCC "R|MD|242|242|Mutual Aid Requested - Counter UAS Observers|=IF('Match Day Resource'!$C$242="""","""",'Match Day Resource'!$C$242)|RATE|'Baseline Costs'!$X$13|1"
    AddCC "R|MD|243|243|Mutual Aid Requested - Counter UAS Observers|=IF('Match Day Resource'!$C$243="""","""",'Match Day Resource'!$C$243)|RATE|'Baseline Costs'!$X$14|1"
    AddCC "R|MD|244|244|Mutual Aid Requested - Counter UAS Observers|=IF('Match Day Resource'!$C$244="""","""",'Match Day Resource'!$C$244)|RATE|'Baseline Costs'!$X$15|1"
    AddCC "R|MD|245|245|Force Supplied - Counter Tactical Advisor|=IF('Match Day Resource'!$C$245="""","""",'Match Day Resource'!$C$245)|ZERO||1"
    AddCC "R|MD|246|246|Force Supplied - Counter Tactical Advisor|=IF('Match Day Resource'!$C$246="""","""",'Match Day Resource'!$C$246)|RATE|'Baseline Costs'!$X$31|1"
    AddCC "R|MD|247|247|Force Supplied - Counter Tactical Advisor|=IF('Match Day Resource'!$C$247="""","""",'Match Day Resource'!$C$247)|RATE|'Baseline Costs'!$X$32|1"
    AddCC "R|MD|248|248|Mutual Aid Requested - Counter Tactical Advisor|=IF('Match Day Resource'!$C$248="""","""",'Match Day Resource'!$C$248)|RATE|'Baseline Costs'!$X$13|1"
    AddCC "R|MD|249|249|Mutual Aid Requested - Counter Tactical Advisor|=IF('Match Day Resource'!$C$249="""","""",'Match Day Resource'!$C$249)|RATE|'Baseline Costs'!$X$14|1"
    AddCC "R|MD|250|250|Mutual Aid Requested - Counter Tactical Advisor|=IF('Match Day Resource'!$C$250="""","""",'Match Day Resource'!$C$250)|RATE|'Baseline Costs'!$X$15|1"
    AddCC "R|MD|251|251|Force Supplied - Drone|=IF('Match Day Resource'!$C$251="""","""",'Match Day Resource'!$C$251)|ZERO||1"
    AddCC "R|MD|252|252|Force Supplied - Drone|=IF('Match Day Resource'!$C$252="""","""",'Match Day Resource'!$C$252)|RATE|'Baseline Costs'!$X$31|1"
    AddCC "R|MD|253|253|Force Supplied - Drone|=IF('Match Day Resource'!$C$253="""","""",'Match Day Resource'!$C$253)|RATE|'Baseline Costs'!$X$32|1"
    AddCC "R|MD|254|254|Mutual Aid Requested - Drone|=IF('Match Day Resource'!$C$254="""","""",'Match Day Resource'!$C$254)|RATE|'Baseline Costs'!$X$13|1"
    AddCC "R|MD|255|255|Mutual Aid Requested - Drone|=IF('Match Day Resource'!$C$255="""","""",'Match Day Resource'!$C$255)|RATE|'Baseline Costs'!$X$14|1"
    AddCC "R|MD|256|256|Mutual Aid Requested - Drone|=IF('Match Day Resource'!$C$256="""","""",'Match Day Resource'!$C$256)|RATE|'Baseline Costs'!$X$15|1"
    AddCC "R|MD|259|259|Force Supplied - LFR Operational Commander|=IF('Match Day Resource'!$C$259="""","""",'Match Day Resource'!$C$259)|ZERO||1"
    AddCC "R|MD|260|260|Force Supplied - LFR Operational Commander|=IF('Match Day Resource'!$C$260="""","""",'Match Day Resource'!$C$260)|RATE|'Baseline Costs'!$X$31|1"
    AddCC "R|MD|261|261|Force Supplied - LFR Operational Commander|=IF('Match Day Resource'!$C$261="""","""",'Match Day Resource'!$C$261)|RATE|'Baseline Costs'!$X$32|1"
    AddCC "R|MD|262|262|Mutual Aid Requested - LFR Operational Commander|=IF('Match Day Resource'!$C$262="""","""",'Match Day Resource'!$C$262)|RATE|'Baseline Costs'!$X$13|1"
    AddCC "R|MD|263|263|Mutual Aid Requested - LFR Operational Commander|=IF('Match Day Resource'!$C$263="""","""",'Match Day Resource'!$C$263)|RATE|'Baseline Costs'!$X$14|1"
    AddCC "R|MD|264|264|Mutual Aid Requested - LFR Operational Commander|=IF('Match Day Resource'!$C$264="""","""",'Match Day Resource'!$C$264)|RATE|'Baseline Costs'!$X$15|1"
    AddCC "R|MD|265|265|Force Supplied - LFR Operator|=IF('Match Day Resource'!$C$265="""","""",'Match Day Resource'!$C$265)|ZERO||1"
    AddCC "R|MD|266|266|Force Supplied - LFR Operator|=IF('Match Day Resource'!$C$266="""","""",'Match Day Resource'!$C$266)|RATE|'Baseline Costs'!$X$31|1"
    AddCC "R|MD|267|267|Force Supplied - LFR Operator|=IF('Match Day Resource'!$C$267="""","""",'Match Day Resource'!$C$267)|RATE|'Baseline Costs'!$X$32|1"
    AddCC "R|MD|268|268|Mutual Aid Requested - LFR Operator|=IF('Match Day Resource'!$C$268="""","""",'Match Day Resource'!$C$268)|RATE|'Baseline Costs'!$X$13|1"
    AddCC "R|MD|269|269|Mutual Aid Requested - LFR Operator|=IF('Match Day Resource'!$C$269="""","""",'Match Day Resource'!$C$269)|RATE|'Baseline Costs'!$X$14|1"
    AddCC "R|MD|270|270|Mutual Aid Requested - LFR Operator|=IF('Match Day Resource'!$C$270="""","""",'Match Day Resource'!$C$270)|RATE|'Baseline Costs'!$X$15|1"
    AddCC "R|MD|273|273|Force Supplied - General Responder|=IF('Match Day Resource'!$C$273="""","""",'Match Day Resource'!$C$273)|ZERO||1"
    AddCC "R|MD|274|274|Force Supplied - General Responder|=IF('Match Day Resource'!$C$274="""","""",'Match Day Resource'!$C$274)|RATE|'Baseline Costs'!$X$31|1"
    AddCC "R|MD|275|275|Force Supplied - General Responder|=IF('Match Day Resource'!$C$275="""","""",'Match Day Resource'!$C$275)|RATE|'Baseline Costs'!$X$32|1"
    AddCC "R|MD|276|276|Mutual Aid Requested - General Responder|=IF('Match Day Resource'!$C$276="""","""",'Match Day Resource'!$C$276)|RATE|'Baseline Costs'!$X$13|1"
    AddCC "R|MD|277|277|Mutual Aid Requested - General Responder|=IF('Match Day Resource'!$C$277="""","""",'Match Day Resource'!$C$277)|RATE|'Baseline Costs'!$X$14|1"
    AddCC "R|MD|278|278|Mutual Aid Requested - General Responder|=IF('Match Day Resource'!$C$278="""","""",'Match Day Resource'!$C$278)|RATE|'Baseline Costs'!$X$15|1"
    AddCC "R|MD|279|279|Force Supplied - CBRN Officer(s)|=IF('Match Day Resource'!$C$279="""","""",'Match Day Resource'!$C$279)|RATE|'Baseline Costs'!$X$14|1"
    AddCC "R|MD|280|280|Force Supplied - CBRN Officer(s)|=IF('Match Day Resource'!$C$280="""","""",'Match Day Resource'!$C$280)|RATE|'Baseline Costs'!$X$14|1"
    AddCC "R|MD|281|281|Force Supplied - CBRN Officer(s)|=IF('Match Day Resource'!$C$281="""","""",'Match Day Resource'!$C$281)|RATE|'Baseline Costs'!$X$14|1"
    AddCC "R|MD|282|282|Force Supplied - CBRN Officer(s)|=IF('Match Day Resource'!$C$282="""","""",'Match Day Resource'!$C$282)|RATE|'Baseline Costs'!$X$14|1"
    AddCC "R|MD|283|283|Force Supplied - CBRN Officer(s)|=IF('Match Day Resource'!$C$283="""","""",'Match Day Resource'!$C$283)|RATE|'Baseline Costs'!$X$14|1"
    AddCC "R|MD|284|284|Mutual Aid Requested - CBRN Officer(s)|=IF('Match Day Resource'!$C$284="""","""",'Match Day Resource'!$C$284)|RATE|'Baseline Costs'!$X$14|1"
    AddCC "R|MD|285|285|Mutual Aid Requested - CBRN Officer(s)|=IF('Match Day Resource'!$C$285="""","""",'Match Day Resource'!$C$285)|RATE|'Baseline Costs'!$X$14|1"
    AddCC "R|MD|286|286|Mutual Aid Requested - CBRN Officer(s)|=IF('Match Day Resource'!$C$286="""","""",'Match Day Resource'!$C$286)|RATE|'Baseline Costs'!$X$14|1"
    AddCC "R|MD|287|287|Mutual Aid Requested - CBRN Officer(s)|=IF('Match Day Resource'!$C$287="""","""",'Match Day Resource'!$C$287)|RATE|'Baseline Costs'!$X$14|1"
    AddCC "R|MD|288|288|Mutual Aid Requested - CBRN Officer(s)|=IF('Match Day Resource'!$C$288="""","""",'Match Day Resource'!$C$288)|RATE|'Baseline Costs'!$X$14|1"
    AddCC "R|MD|291|291|Force Supplied - Command|=IF('Match Day Resource'!$C$291="""","""",'Match Day Resource'!$C$291)|ZERO||1"
    AddCC "R|MD|292|292|Force Supplied - Command|=IF('Match Day Resource'!$C$292="""","""",'Match Day Resource'!$C$292)|ZERO||1"
    AddCC "R|MD|293|293|Force Supplied - Command|=IF('Match Day Resource'!$C$293="""","""",'Match Day Resource'!$C$293)|ZERO||1"
    AddCC "R|MD|294|294|Force Supplied - Command|=IF('Match Day Resource'!$C$294="""","""",'Match Day Resource'!$C$294)|ZERO||1"
    AddCC "R|MD|295|295|Force Supplied - Command|=IF('Match Day Resource'!$C$295="""","""",'Match Day Resource'!$C$295)|ZERO||1"
    AddCC "R|MD|296|296|Force Supplied - Command|=IF('Match Day Resource'!$C$296="""","""",'Match Day Resource'!$C$296)|RATE|'Baseline Costs'!$X$31|1"
    AddCC "R|MD|297|297|Force Supplied - Command|=IF('Match Day Resource'!$C$297="""","""",'Match Day Resource'!$C$297)|RATE|'Baseline Costs'!$X$32|1"
    AddCC "R|MD|300|300|Force Supplied-  Control Room|=IF('Match Day Resource'!$C$300="""","""",'Match Day Resource'!$C$300)|ZERO||1"
    AddCC "R|MD|301|301|Force Supplied-  Control Room|=IF('Match Day Resource'!$C$301="""","""",'Match Day Resource'!$C$301)|RATE|'Baseline Costs'!$X$31|1"
    AddCC "R|MD|302|302|Force Supplied-  Control Room|=IF('Match Day Resource'!$C$302="""","""",'Match Day Resource'!$C$302)|RATE|'Baseline Costs'!$X$32|1"
    AddCC "R|MD|305|305|Force Supplied - Custody|=IF('Match Day Resource'!$C$305="""","""",'Match Day Resource'!$C$305)|RATE|'Baseline Costs'!$X$31|1"
    AddCC "R|MD|306|306|Force Supplied - Custody|=IF('Match Day Resource'!$C$306="""","""",'Match Day Resource'!$C$306)|RATE|'Baseline Costs'!$X$32|1"
    AddCC "R|MD|309|309|Force Supplied - Crime|=IF('Match Day Resource'!$C$309="""","""",'Match Day Resource'!$C$309)|ZERO||1"
    AddCC "R|MD|310|310|Force Supplied - Crime|=IF('Match Day Resource'!$C$310="""","""",'Match Day Resource'!$C$310)|RATE|'Baseline Costs'!$X$31|1"
    AddCC "R|MD|311|311|Force Supplied - Crime|=IF('Match Day Resource'!$C$311="""","""",'Match Day Resource'!$C$311)|RATE|'Baseline Costs'!$X$32|1"
    AddCC "R|MD|312|312|Mutual Aid Requested - Crime|=IF('Match Day Resource'!$C$312="""","""",'Match Day Resource'!$C$312)|RATE|'Baseline Costs'!$X$13|1"
    AddCC "R|MD|313|313|Mutual Aid Requested - Crime|=IF('Match Day Resource'!$C$313="""","""",'Match Day Resource'!$C$313)|RATE|'Baseline Costs'!$X$14|1"
    AddCC "R|MD|314|314|Mutual Aid Requested - Crime|=IF('Match Day Resource'!$C$314="""","""",'Match Day Resource'!$C$314)|RATE|'Baseline Costs'!$X$15|1"
    AddCC "R|MD|317|317|Force Supplied - Intel|=IF('Match Day Resource'!$C$317="""","""",'Match Day Resource'!$C$317)|ZERO||1"
    AddCC "R|MD|318|318|Force Supplied - Intel|=IF('Match Day Resource'!$C$318="""","""",'Match Day Resource'!$C$318)|RATE|'Baseline Costs'!$X$31|1"
    AddCC "R|MD|319|319|Force Supplied - Intel|=IF('Match Day Resource'!$C$319="""","""",'Match Day Resource'!$C$319)|RATE|'Baseline Costs'!$X$32|1"
    AddCC "R|MD|320|320|Mutual Aid Requested - Intel|=IF('Match Day Resource'!$C$320="""","""",'Match Day Resource'!$C$320)|RATE|'Baseline Costs'!$X$13|1"
    AddCC "R|MD|321|321|Mutual Aid Requested - Intel|=IF('Match Day Resource'!$C$321="""","""",'Match Day Resource'!$C$321)|RATE|'Baseline Costs'!$X$14|1"
    AddCC "R|MD|322|322|Mutual Aid Requested - Intel|=IF('Match Day Resource'!$C$322="""","""",'Match Day Resource'!$C$322)|RATE|'Baseline Costs'!$X$15|1"
    AddCC "R|MD|325|325|Force Supplied - Ancillary Team|=IF('Match Day Resource'!$C$325="""","""",'Match Day Resource'!$C$325)|RATE|'Baseline Costs'!$X$31|1"
    AddCC "R|MD|326|326|Force Supplied - Ancillary Team|=IF('Match Day Resource'!$C$326="""","""",'Match Day Resource'!$C$326)|RATE|'Baseline Costs'!$X$32|1"
    AddCC "R|MD|327|327|Mutual Aid Requested - Ancillary Team|=IF('Match Day Resource'!$C$327="""","""",'Match Day Resource'!$C$327)|RATE|'Baseline Costs'!$X$14|1"
    AddCC "R|MD|328|328|Mutual Aid Requested - Ancillary Team|=IF('Match Day Resource'!$C$328="""","""",'Match Day Resource'!$C$328)|RATE|'Baseline Costs'!$X$15|1"
    AddCC "R|MD|331|331||=IF('Match Day Resource'!$C$331="""","""",'Match Day Resource'!$C$331)|RATE|'Baseline Costs'!$X$14|1"
    AddCC "R|MD|332|332||=IF('Match Day Resource'!$C$332="""","""",'Match Day Resource'!$C$332)|RATE|'Baseline Costs'!$X$14|1"
    AddCC "R|MD|333|333||=IF('Match Day Resource'!$C$333="""","""",'Match Day Resource'!$C$333)|RATE|'Baseline Costs'!$X$14|1"
    AddCC "R|MD|334|334||=IF('Match Day Resource'!$C$334="""","""",'Match Day Resource'!$C$334)|RATE|'Baseline Costs'!$X$14|1"
    AddCC "R|MD|335|335||=IF('Match Day Resource'!$C$335="""","""",'Match Day Resource'!$C$335)|RATE|'Baseline Costs'!$X$14|1"
    AddCC "R|MD|336|336||=IF('Match Day Resource'!$C$336="""","""",'Match Day Resource'!$C$336)|RATE|'Baseline Costs'!$X$14|1"
    AddCC "R|MD|337|337||=IF('Match Day Resource'!$C$337="""","""",'Match Day Resource'!$C$337)|RATE|'Baseline Costs'!$X$14|1"
    AddCC "R|MD|338|338||=IF('Match Day Resource'!$C$338="""","""",'Match Day Resource'!$C$338)|RATE|'Baseline Costs'!$X$14|1"
    AddCC "R|MD|339|339||=IF('Match Day Resource'!$C$339="""","""",'Match Day Resource'!$C$339)|RATE|'Baseline Costs'!$X$14|1"
    AddCC "R|WT|6|351|Force Supplied - Level 2 PSU / Serials Wider Community Public Order Policing|=IF('Wider Tournament Resource'!$C$6="""","""",'Wider Tournament Resource'!$C$6)|ZERO||0"
    AddCC "R|WT|7|352|Force Supplied - Level 2 PSU / Serials Wider Community Public Order Policing|=IF('Wider Tournament Resource'!$C$7="""","""",'Wider Tournament Resource'!$C$7)|RATE|'Baseline Costs'!$X$31|0"
    AddCC "R|WT|8|353|Force Supplied - Level 2 PSU / Serials Wider Community Public Order Policing|=IF('Wider Tournament Resource'!$C$8="""","""",'Wider Tournament Resource'!$C$8)|RATE|'Baseline Costs'!$X$32|0"
    AddCC "R|WT|9|354|Force Supplied - Level 3 BDU / Serials Wider Community Public Order Policing|=IF('Wider Tournament Resource'!$C$9="""","""",'Wider Tournament Resource'!$C$9)|ZERO||0"
    AddCC "R|WT|10|355|Force Supplied - Level 3 BDU / Serials Wider Community Public Order Policing|=IF('Wider Tournament Resource'!$C$10="""","""",'Wider Tournament Resource'!$C$10)|RATE|'Baseline Costs'!$X$31|0"
    AddCC "R|WT|11|356|Force Supplied - Level 3 BDU / Serials Wider Community Public Order Policing|=IF('Wider Tournament Resource'!$C$11="""","""",'Wider Tournament Resource'!$C$11)|RATE|'Baseline Costs'!$X$32|0"
    AddCC "R|WT|12|357|Force Supplied - Level 2 PSU / Serials Fan Zone Public Order Policing|=IF('Wider Tournament Resource'!$C$12="""","""",'Wider Tournament Resource'!$C$12)|ZERO||0"
End Sub

Private Sub DataCC3()
    Dim s As String
    AddCC "R|WT|13|358|Force Supplied - Level 2 PSU / Serials Fan Zone Public Order Policing|=IF('Wider Tournament Resource'!$C$13="""","""",'Wider Tournament Resource'!$C$13)|RATE|'Baseline Costs'!$X$31|0"
    AddCC "R|WT|14|359|Force Supplied - Level 2 PSU / Serials Fan Zone Public Order Policing|=IF('Wider Tournament Resource'!$C$14="""","""",'Wider Tournament Resource'!$C$14)|RATE|'Baseline Costs'!$X$32|0"
    AddCC "R|WT|15|360|Force Supplied - Level 3 BDU / Serials Fan Zone Public Order Policing|=IF('Wider Tournament Resource'!$C$15="""","""",'Wider Tournament Resource'!$C$15)|ZERO||0"
    AddCC "R|WT|16|361|Force Supplied - Level 3 BDU / Serials Fan Zone Public Order Policing|=IF('Wider Tournament Resource'!$C$16="""","""",'Wider Tournament Resource'!$C$16)|RATE|'Baseline Costs'!$X$31|0"
    AddCC "R|WT|17|362|Force Supplied - Level 3 BDU / Serials Fan Zone Public Order Policing|=IF('Wider Tournament Resource'!$C$17="""","""",'Wider Tournament Resource'!$C$17)|RATE|'Baseline Costs'!$X$32|0"
    AddCC "R|WT|20|365|Force Supplied - Wider Community Public Order Policing|=IF('Wider Tournament Resource'!$C$20="""","""",'Wider Tournament Resource'!$C$20)|SKIP||0"
    AddCC "R|WT|21|366|Force Supplied - Wider Community Public Order Policing|=IF('Wider Tournament Resource'!$C$21="""","""",'Wider Tournament Resource'!$C$21)|SKIP||0"
    AddCC "R|WT|22|367|Mutual Aid Requested - Wider Community Public Order Policing|=IF('Wider Tournament Resource'!$C$22="""","""",'Wider Tournament Resource'!$C$22)|PKG||0"
    AddCC "R|WT|23|368|Mutual Aid Requested - Wider Community Public Order Policing|=IF('Wider Tournament Resource'!$C$23="""","""",'Wider Tournament Resource'!$C$23)|PKG||0"
    AddCC "R|WT|24|369|Force Supplied - Fan Zone Public Order Policing|=IF('Wider Tournament Resource'!$C$24="""","""",'Wider Tournament Resource'!$C$24)|SKIP||0"
    AddCC "R|WT|25|370|Force Supplied - Fan Zone Public Order Policing|=IF('Wider Tournament Resource'!$C$25="""","""",'Wider Tournament Resource'!$C$25)|SKIP||0"
    AddCC "R|WT|26|371|Mutual Aid Requested - Fan Zone Public Order Policing|=IF('Wider Tournament Resource'!$C$26="""","""",'Wider Tournament Resource'!$C$26)|PKG||0"
    AddCC "R|WT|27|372|Mutual Aid Requested - Fan Zone Public Order Policing|=IF('Wider Tournament Resource'!$C$27="""","""",'Wider Tournament Resource'!$C$27)|PKG||0"
    AddCC "R|WT|28|373|Force Supplied - Operationally Deployable non-Public Order Policing|=IF('Wider Tournament Resource'!$C$28="""","""",'Wider Tournament Resource'!$C$28)|ZERO||0"
    AddCC "R|WT|29|374|Force Supplied - Operationally Deployable non-Public Order Policing|=IF('Wider Tournament Resource'!$C$29="""","""",'Wider Tournament Resource'!$C$29)|RATE|'Baseline Costs'!$X$31|0"
    AddCC "R|WT|30|375|Force Supplied - Operationally Deployable non-Public Order Policing|=IF('Wider Tournament Resource'!$C$30="""","""",'Wider Tournament Resource'!$C$30)|RATE|'Baseline Costs'!$X$32|0"
    AddCC "R|WT|31|376|Mutual Aid Requested - Operationally Deployable non-Public Order Policing|=IF('Wider Tournament Resource'!$C$31="""","""",'Wider Tournament Resource'!$C$31)|RATE|'Baseline Costs'!$X$13|0"
    AddCC "R|WT|32|377|Mutual Aid Requested - Operationally Deployable non-Public Order Policing|=IF('Wider Tournament Resource'!$C$32="""","""",'Wider Tournament Resource'!$C$32)|RATE|'Baseline Costs'!$X$14|0"
    AddCC "R|WT|33|378|Mutual Aid Requested - Operationally Deployable non-Public Order Policing|=IF('Wider Tournament Resource'!$C$33="""","""",'Wider Tournament Resource'!$C$33)|RATE|'Baseline Costs'!$X$15|0"
    AddCC "R|WT|34|379|Force Supplied - Evidence GT's|=IF('Wider Tournament Resource'!$C$34="""","""",'Wider Tournament Resource'!$C$34)|RATE|'Baseline Costs'!$X$31|0"
    AddCC "R|WT|35|380|Force Supplied - Evidence GT's|=IF('Wider Tournament Resource'!$C$35="""","""",'Wider Tournament Resource'!$C$35)|RATE|'Baseline Costs'!$X$32|0"
    AddCC "R|WT|36|381|Mutual Aid Requested - Evidence Gatherers|=IF('Wider Tournament Resource'!$C$36="""","""",'Wider Tournament Resource'!$C$36)|RATE|'Baseline Costs'!$X$14|0"
    AddCC "R|WT|37|382|Mutual Aid Requested - Evidence Gatherers|=IF('Wider Tournament Resource'!$C$37="""","""",'Wider Tournament Resource'!$C$37)|RATE|'Baseline Costs'!$X$15|0"
    AddCC "R|WT|39|384|Force Supplied - Additional Public Order|=IF('Wider Tournament Resource'!$C$39="""","""",'Wider Tournament Resource'!$C$39)|RATE|'Baseline Costs'!$X$14|0"
    AddCC "R|WT|40|385|Force Supplied - Additional Public Order|=IF('Wider Tournament Resource'!$C$40="""","""",'Wider Tournament Resource'!$C$40)|RATE|'Baseline Costs'!$X$14|0"
    AddCC "R|WT|41|386|Force Supplied - Additional Public Order|=IF('Wider Tournament Resource'!$C$41="""","""",'Wider Tournament Resource'!$C$41)|RATE|'Baseline Costs'!$X$14|0"
    AddCC "R|WT|42|387|Force Supplied - Additional Public Order|=IF('Wider Tournament Resource'!$C$42="""","""",'Wider Tournament Resource'!$C$42)|RATE|'Baseline Costs'!$X$14|0"
    AddCC "R|WT|43|388|Force Supplied - Additional Public Order|=IF('Wider Tournament Resource'!$C$43="""","""",'Wider Tournament Resource'!$C$43)|RATE|'Baseline Costs'!$X$14|0"
    AddCC "R|WT|44|389|Force Supplied - Additional Public Order|=IF('Wider Tournament Resource'!$C$44="""","""",'Wider Tournament Resource'!$C$44)|RATE|'Baseline Costs'!$X$14|0"
    AddCC "R|WT|45|390|Force Supplied - Additional Public Order|=IF('Wider Tournament Resource'!$C$45="""","""",'Wider Tournament Resource'!$C$45)|RATE|'Baseline Costs'!$X$14|0"
    AddCC "R|WT|46|391|Force Supplied - Additional Public Order|=IF('Wider Tournament Resource'!$C$46="""","""",'Wider Tournament Resource'!$C$46)|RATE|'Baseline Costs'!$X$14|0"
    AddCC "R|WT|47|392|Force Supplied - Additional Public Order|=IF('Wider Tournament Resource'!$C$47="""","""",'Wider Tournament Resource'!$C$47)|RATE|'Baseline Costs'!$X$14|0"
    AddCC "R|WT|48|393|Mutual Aid Requested -   Additional PO|=IF('Wider Tournament Resource'!$C$48="""","""",'Wider Tournament Resource'!$C$48)|RATE|'Baseline Costs'!$X$14|0"
    AddCC "R|WT|49|394|Mutual Aid Requested -   Additional PO|=IF('Wider Tournament Resource'!$C$49="""","""",'Wider Tournament Resource'!$C$49)|RATE|'Baseline Costs'!$X$14|0"
    AddCC "R|WT|50|395|Mutual Aid Requested -   Additional PO|=IF('Wider Tournament Resource'!$C$50="""","""",'Wider Tournament Resource'!$C$50)|RATE|'Baseline Costs'!$X$14|0"
    AddCC "R|WT|51|396|Mutual Aid Requested -   Additional PO|=IF('Wider Tournament Resource'!$C$51="""","""",'Wider Tournament Resource'!$C$51)|RATE|'Baseline Costs'!$X$14|0"
    AddCC "R|WT|52|397|Mutual Aid Requested -   Additional PO|=IF('Wider Tournament Resource'!$C$52="""","""",'Wider Tournament Resource'!$C$52)|RATE|'Baseline Costs'!$X$14|0"
    AddCC "R|WT|53|398|Mutual Aid Requested -   Additional PO|=IF('Wider Tournament Resource'!$C$53="""","""",'Wider Tournament Resource'!$C$53)|RATE|'Baseline Costs'!$X$14|0"
    AddCC "R|WT|54|399|Mutual Aid Requested -   Additional PO|=IF('Wider Tournament Resource'!$C$54="""","""",'Wider Tournament Resource'!$C$54)|RATE|'Baseline Costs'!$X$14|0"
    AddCC "R|WT|55|400|Mutual Aid Requested -   Additional PO|=IF('Wider Tournament Resource'!$C$55="""","""",'Wider Tournament Resource'!$C$55)|RATE|'Baseline Costs'!$X$14|0"
    AddCC "R|WT|56|401|Mutual Aid Requested -   Additional PO|=IF('Wider Tournament Resource'!$C$56="""","""",'Wider Tournament Resource'!$C$56)|RATE|'Baseline Costs'!$X$14|0"
    AddCC "R|WT|59|404|Force Supplied - Disruptive Effects Officer|=IF('Wider Tournament Resource'!$C$59="""","""",'Wider Tournament Resource'!$C$59)|RATE|'Baseline Costs'!$X$31|0"
    AddCC "R|WT|60|405|Force Supplied - Disruptive Effects Officer|=IF('Wider Tournament Resource'!$C$60="""","""",'Wider Tournament Resource'!$C$60)|RATE|'Baseline Costs'!$X$32|0"
    AddCC "R|WT|61|406|Mutual Aid -  Disruptive Effects Officer|=IF('Wider Tournament Resource'!$C$61="""","""",'Wider Tournament Resource'!$C$61)|RATE|'Baseline Costs'!$X$14|0"
    AddCC "R|WT|62|407|Mutual Aid -  Disruptive Effects Officer|=IF('Wider Tournament Resource'!$C$62="""","""",'Wider Tournament Resource'!$C$62)|RATE|'Baseline Costs'!$X$15|0"
    AddCC "R|WT|63|408|Force Supplied - Disruptive Effects Operational Commander|=IF('Wider Tournament Resource'!$C$63="""","""",'Wider Tournament Resource'!$C$63)|RATE|'Baseline Costs'!$X$31|0"
    AddCC "R|WT|64|409|Force Supplied - Disruptive Effects Operational Commander|=IF('Wider Tournament Resource'!$C$64="""","""",'Wider Tournament Resource'!$C$64)|RATE|'Baseline Costs'!$X$32|0"
    AddCC "R|WT|65|410|Mutual Aid - Disruptive Effects Operational Commander|=IF('Wider Tournament Resource'!$C$65="""","""",'Wider Tournament Resource'!$C$65)|RATE|'Baseline Costs'!$X$14|0"
    AddCC "R|WT|66|411|Mutual Aid - Disruptive Effects Operational Commander|=IF('Wider Tournament Resource'!$C$66="""","""",'Wider Tournament Resource'!$C$66)|RATE|'Baseline Costs'!$X$15|0"
    AddCC "R|WT|67|412|Force Supplied - Disruptive Effects Tactical Commander|=IF('Wider Tournament Resource'!$C$67="""","""",'Wider Tournament Resource'!$C$67)|RATE|'Baseline Costs'!$X$31|0"
    AddCC "R|WT|68|413|Force Supplied - Disruptive Effects Tactical Commander|=IF('Wider Tournament Resource'!$C$68="""","""",'Wider Tournament Resource'!$C$68)|RATE|'Baseline Costs'!$X$32|0"
    AddCC "R|WT|69|414|Mutual Aid - Disruptive Effects Tactical Commander|=IF('Wider Tournament Resource'!$C$69="""","""",'Wider Tournament Resource'!$C$69)|RATE|'Baseline Costs'!$X$14|0"
    AddCC "R|WT|70|415|Mutual Aid - Disruptive Effects Tactical Commander|=IF('Wider Tournament Resource'!$C$70="""","""",'Wider Tournament Resource'!$C$70)|RATE|'Baseline Costs'!$X$15|0"
    AddCC "R|WT|73|418|Force Supplied- Fan Zone Armed Response (ARVs)|=IF('Wider Tournament Resource'!$C$73="""","""",'Wider Tournament Resource'!$C$73)|RATE|'Baseline Costs'!$X$31|0"
    AddCC "R|WT|74|419|Force Supplied- Fan Zone Armed Response (ARVs)|=IF('Wider Tournament Resource'!$C$74="""","""",'Wider Tournament Resource'!$C$74)|RATE|'Baseline Costs'!$X$32|0"
    AddCC "R|WT|75|420|Mutual Aid Requested- Fan Zone Armed Response (ARVs)|=IF('Wider Tournament Resource'!$C$75="""","""",'Wider Tournament Resource'!$C$75)|RATE|'Baseline Costs'!$X$14|0"
    AddCC "R|WT|76|421|Mutual Aid Requested- Fan Zone Armed Response (ARVs)|=IF('Wider Tournament Resource'!$C$76="""","""",'Wider Tournament Resource'!$C$76)|RATE|'Baseline Costs'!$X$15|0"
    AddCC "R|WT|79|424|Force Supplied - Control Room|=IF('Wider Tournament Resource'!$C$79="""","""",'Wider Tournament Resource'!$C$79)|RATE|'Baseline Costs'!$X$31|0"
    AddCC "R|WT|80|425|Force Supplied - Control Room|=IF('Wider Tournament Resource'!$C$80="""","""",'Wider Tournament Resource'!$C$80)|RATE|'Baseline Costs'!$X$32|0"
    AddCC "R|WT|81|426|Force Supplied - Command Support|=IF('Wider Tournament Resource'!$C$81="""","""",'Wider Tournament Resource'!$C$81)|ZERO||0"
    AddCC "R|WT|82|427|Force Supplied - Command Support|=IF('Wider Tournament Resource'!$C$82="""","""",'Wider Tournament Resource'!$C$82)|ZERO||0"
    AddCC "R|WT|83|428|Force Supplied - Command Support|=IF('Wider Tournament Resource'!$C$83="""","""",'Wider Tournament Resource'!$C$83)|RATE|'Baseline Costs'!$X$31|0"
    AddCC "R|WT|84|429|Force Supplied - Command Support|=IF('Wider Tournament Resource'!$C$84="""","""",'Wider Tournament Resource'!$C$84)|RATE|'Baseline Costs'!$X$32|0"
    AddCC "R|WT|85|430|Mutual Aid Requested - Command Support|=IF('Wider Tournament Resource'!$C$85="""","""",'Wider Tournament Resource'!$C$85)|RATE|'Baseline Costs'!$X$12|0"
    AddCC "R|WT|86|431|Mutual Aid Requested - Command Support|=IF('Wider Tournament Resource'!$C$86="""","""",'Wider Tournament Resource'!$C$86)|RATE|'Baseline Costs'!$X$13|0"
    AddCC "R|WT|87|432|Mutual Aid Requested - Command Support|=IF('Wider Tournament Resource'!$C$87="""","""",'Wider Tournament Resource'!$C$87)|RATE|'Baseline Costs'!$X$14|0"
    AddCC "R|WT|88|433|Mutual Aid Requested - Command Support|=IF('Wider Tournament Resource'!$C$88="""","""",'Wider Tournament Resource'!$C$88)|RATE|'Baseline Costs'!$X$15|0"
    AddCC "R|WT|89|434|Force Supplied - Custody|=IF('Wider Tournament Resource'!$C$89="""","""",'Wider Tournament Resource'!$C$89)|RATE|'Baseline Costs'!$X$31|0"
    AddCC "R|WT|90|435|Force Supplied - Custody|=IF('Wider Tournament Resource'!$C$90="""","""",'Wider Tournament Resource'!$C$90)|RATE|'Baseline Costs'!$X$32|0"
    AddCC "R|WT|93|438||=IF('Wider Tournament Resource'!$C$93="""","""",'Wider Tournament Resource'!$C$93)|RATE|'Baseline Costs'!$X$14|0"
    AddCC "R|WT|94|439||=IF('Wider Tournament Resource'!$C$94="""","""",'Wider Tournament Resource'!$C$94)|RATE|'Baseline Costs'!$X$14|0"
    AddCC "R|WT|95|440||=IF('Wider Tournament Resource'!$C$95="""","""",'Wider Tournament Resource'!$C$95)|RATE|'Baseline Costs'!$X$14|0"
    AddCC "R|WT|96|441||=IF('Wider Tournament Resource'!$C$96="""","""",'Wider Tournament Resource'!$C$96)|RATE|'Baseline Costs'!$X$14|0"
    AddCC "R|WT|97|442||=IF('Wider Tournament Resource'!$C$97="""","""",'Wider Tournament Resource'!$C$97)|RATE|'Baseline Costs'!$X$14|0"
    AddCC "R|WT|98|443||=IF('Wider Tournament Resource'!$C$98="""","""",'Wider Tournament Resource'!$C$98)|RATE|'Baseline Costs'!$X$14|0"
    AddCC "R|WT|99|444||=IF('Wider Tournament Resource'!$C$99="""","""",'Wider Tournament Resource'!$C$99)|RATE|'Baseline Costs'!$X$14|0"
    AddCC "R|WT|100|445||=IF('Wider Tournament Resource'!$C$100="""","""",'Wider Tournament Resource'!$C$100)|RATE|'Baseline Costs'!$X$14|0"
    AddCC "R|WT|101|446||=IF('Wider Tournament Resource'!$C$101="""","""",'Wider Tournament Resource'!$C$101)|RATE|'Baseline Costs'!$X$14|0"
End Sub

Private Sub DataST1_1()
    Dim s As String
    AddST "C|A4|s|INFLATION DEFLATORS|General|1|0|9.0|192|15917529||0|0"
    AddST "C|D4|s|POLICE PAY AWARD|General|1|0|9.0|192|15917529||0|0"
    AddST "C|G4|s|2025/2026 SALARY|General|1|0|9.0|192|15917529||0|0"
    AddST "C|H4|e||General|1|0|9.0|192|15917529||0|0"
    AddST "C|A5|s|2025/2026|General|0|0|9.0||15917529||0|1"
    AddST "C|B5|n|1|0.000000000|0|0|9.0||15917529||0|1"
    AddST "C|D5|s|2025-2026|General|0|0|9.0||15917529|center|0|1"
    AddST "C|E5|n|0.042|0.00%|0|0|9.0||15917529|center|0|1"
    AddST "C|G5|e||General|0|0|9.0||15917529||0|1"
    AddST "C|H5|s|C/Supt|General|0|0|9.0||15917529|center|0|1"
    AddST "C|I5|s|Supt|General|0|0|9.0||15917529|center|0|1"
    AddST "C|J5|s|C/Insp|General|0|0|9.0||15917529|center|0|1"
    AddST "C|K5|s|Insp|General|0|0|9.0||15917529|center|0|1"
    AddST "C|L5|s|Sgt|General|0|0|9.0||15917529|center|0|1"
    AddST "C|M5|s|PC|General|0|0|9.0||15917529|center|0|1"
    AddST "C|A6|s|2026/2027|General|0|0|9.0||15917529||0|1"
    AddST "C|B6|f|=B5*(1+('GDP Deflators'!D79/100))|0.000000000|0|0|9.0||15917529||0|0"
    AddST "C|D6|s|2026-2027|General|0|0|9.0||15917529|center|0|1"
    AddST "C|E6|n|0.04|0.00%|0|0|9.0||15917529|center|0|0"
    AddST "C|G6|s|Annual|General|0|0|9.0||15917529||0|1"
    AddST "C|H6|n|172900|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|I6|n|143600|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|J6|n|109800|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|K6|n|101900|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|L6|n|83200|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|M6|n|70300|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|A7|s|2027/2028|General|0|0|9.0||15917529||0|1"
    AddST "C|B7|f|=B5*(1+('GDP Deflators'!D80/100))|0.000000000|0|0|9.0||15917529||0|0"
    AddST "C|D7|s|2027-2028|General|0|0|9.0||15917529|center|0|1"
    AddST "C|E7|n|0.035|0.00%|0|0|9.0||15917529|center|0|0"
    AddST "C|G7|s|Monthly|General|0|0|9.0||15917529||0|1"
    AddST "C|H7|f|=H6/12|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|I7|f|=I6/12|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|J7|f|=J6/12|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|K7|f|=K6/12|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|L7|f|=L6/12|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|M7|f|=M6/12|""~uA3~""#,##0.00|0|0|9.0||15917529||0|0"
    AddST "C|A8|s|2028/2029|General|0|0|9.0||15917529||0|1"
    AddST "C|B8|f|=B5*(1+('GDP Deflators'!D81/100))|0.000000000|0|0|9.0||15917529||0|0"
    AddST "C|D8|s|2028-2029|General|0|0|9.0||15917529|center|0|1"
    AddST "C|E8|n|0.03|0.00%|0|0|9.0||15917529|center|0|0"
    AddST "C|G8|s|Hourly|General|0|0|9.0||15917529||0|1"
    AddST "C|H8|n|82.86|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|I8|n|68.82|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|J8|n|52.62|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|K8|n|48.83|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|L8|n|39.87|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|M8|n|33.69|""~uA3~""#,##0.00|0|0|9.0||15917529||0|0"
    AddST "C|B12|d|2026-09-01|mmm-yy|1|0|9.0||7159831||0|1"
    AddST "C|C12|d|2026-10-01|mmm-yy|1|0|9.0||7159831||0|1"
    AddST "C|D12|d|2026-11-01|mmm-yy|1|0|9.0||7159831||0|1"
    AddST "C|E12|d|2026-12-01|mmm-yy|1|0|9.0||7159831||0|1"
    AddST "C|F12|d|2027-01-01|mmm-yy|1|0|9.0||7159831||0|1"
    AddST "C|G12|d|2027-02-01|mmm-yy|1|0|9.0||7159831||0|1"
    AddST "C|H12|d|2027-03-01|mmm-yy|1|0|9.0||7159831||0|1"
    AddST "C|I12|d|2027-04-01|mmm-yy|1|0|9.0||7159831||0|1"
    AddST "C|J12|d|2027-05-01|mmm-yy|1|0|9.0||7159831||0|1"
    AddST "C|K12|d|2027-06-01|mmm-yy|1|0|9.0||7159831||0|1"
    AddST "C|L12|d|2027-07-01|mmm-yy|1|0|9.0||7159831||0|1"
    AddST "C|M12|d|2027-08-01|mmm-yy|1|0|9.0||7159831||0|1"
    AddST "C|N12|d|2027-09-01|mmm-yy|1|0|9.0||7159831||0|1"
    AddST "C|O12|d|2027-10-01|mmm-yy|1|0|9.0||7159831||0|1"
    AddST "C|P12|d|2027-11-01|mmm-yy|1|0|9.0||7159831||0|1"
    AddST "C|Q12|d|2027-12-01|mmm-yy|1|0|9.0||7159831||0|1"
    AddST "C|R12|d|2028-01-01|mmm-yy|1|0|9.0||7159831||0|1"
    AddST "C|S12|d|2028-02-01|mmm-yy|1|0|9.0||7159831||0|1"
    AddST "C|T12|d|2028-03-01|mmm-yy|1|0|9.0||7159831||0|1"
    AddST "C|U12|d|2028-04-01|mmm-yy|1|0|9.0||7159831||0|1"
    AddST "C|V12|d|2028-05-01|mmm-yy|1|0|9.0||7159831||0|1"
    AddST "C|W12|d|2028-06-01|mmm-yy|1|0|9.0||7159831||0|1"
    AddST "C|X12|d|2028-07-01|mmm-yy|1|0|9.0||7159831||0|1"
    AddST "C|Y12|d|2028-08-01|mmm-yy|1|0|9.0||7159831||0|1"
    AddST "C|Z12|d|2028-09-01|mmm-yy|1|0|9.0||7159831||0|1"
    AddST "C|AA12|d|2028-10-01|mmm-yy|1|0|9.0||7159831||0|1"
    AddST "C|AB12|d|2028-11-01|mmm-yy|1|0|9.0||7159831||0|1"
    AddST "C|AC12|d|2028-12-01|mmm-yy|1|0|9.0||7159831||0|1"
    AddST "C|A13|s|C/Superintendent|General|0|0|9.0||15917529||0|1"
    AddST "C|B13|f|=(H7*B6)+(H7*E6)|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|C13|f|=B13|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|D13|f|=C13|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|E13|f|=D13|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|F13|f|=E13|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|G13|f|=F13|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|H13|f|=G13|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|I13|f|=H13|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|J13|f|=I13|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|K13|f|=J13|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|L13|f|=K13|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|M13|f|=L13|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|N13|f|=M13+(M13*$E$7)|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|O13|f|=N13|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|P13|f|=O13|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|Q13|f|=P13|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|R13|f|=Q13|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|S13|f|=R13|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|T13|f|=S13|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|U13|f|=T13|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|V13|f|=U13|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|W13|f|=V13|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|X13|f|=W13|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|Y13|f|=X13|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|Z13|f|=Y13+(Y13*$E$8)|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|AA13|f|=Z13|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|AB13|f|=AA13|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|AC13|f|=AB13|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|A14|s|Superintendent|General|0|0|9.0||15917529||0|1"
    AddST "C|B14|f|=(I7*B6)+(I7*E6)|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|C14|f|=B14|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|D14|f|=C14|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|E14|f|=D14|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|F14|f|=E14|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|G14|f|=F14|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|H14|f|=G14|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|I14|f|=H14|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|J14|f|=I14|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|K14|f|=J14|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|L14|f|=K14|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|M14|f|=L14|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|N14|f|=M14+(M14*$E$7)|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|O14|f|=N14|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|P14|f|=O14|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|Q14|f|=P14|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|R14|f|=Q14|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|S14|f|=R14|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|T14|f|=S14|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|U14|f|=T14|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|V14|f|=U14|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|W14|f|=V14|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|X14|f|=W14|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|Y14|f|=X14|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|Z14|f|=Y14+(Y14*$E$8)|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|AA14|f|=Z14|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|AB14|f|=AA14|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|AC14|f|=AB14|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|A15|s|Chief Inspector|General|0|0|9.0||15917529||0|1"
    AddST "C|B15|f|=(J7*B6)+(J7*E6)|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|C15|f|=B15|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|D15|f|=C15|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|E15|f|=D15|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|F15|f|=E15|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|G15|f|=F15|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|H15|f|=G15|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|I15|f|=H15|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|J15|f|=I15|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|K15|f|=J15|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|L15|f|=K15|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|M15|f|=L15|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|N15|f|=M15+(M15*$E$7)|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|O15|f|=N15|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|P15|f|=O15|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
End Sub

Private Sub DataST1_2()
    Dim s As String
    AddST "C|Q15|f|=P15|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|R15|f|=Q15|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|S15|f|=R15|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|T15|f|=S15|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|U15|f|=T15|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|V15|f|=U15|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|W15|f|=V15|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|X15|f|=W15|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|Y15|f|=X15|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|Z15|f|=Y15+(Y15*$E$8)|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|AA15|f|=Z15|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|AB15|f|=AA15|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|AC15|f|=AB15|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|A16|s|Inspector|General|0|0|9.0||15917529||0|1"
    AddST "C|B16|f|=(K7*B6)+(K7*E6)|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|C16|f|=B16|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|D16|f|=C16|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|E16|f|=D16|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|F16|f|=E16|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|G16|f|=F16|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|H16|f|=G16|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|I16|f|=H16|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|J16|f|=I16|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|K16|f|=J16|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|L16|f|=K16|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|M16|f|=L16|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|N16|f|=M16+(M16*$E$7)|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|O16|f|=N16|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|P16|f|=O16|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|Q16|f|=P16|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|R16|f|=Q16|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|S16|f|=R16|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|T16|f|=S16|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|U16|f|=T16|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|V16|f|=U16|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|W16|f|=V16|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|X16|f|=W16|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|Y16|f|=X16|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|Z16|f|=Y16+(Y16*$E$8)|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|AA16|f|=Z16|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|AB16|f|=AA16|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|AC16|f|=AB16|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|A17|s|Sergeant|General|0|0|9.0||15917529||0|1"
    AddST "C|B17|f|=(L7*B6)+(L7*E6)|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|C17|f|=B17|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|D17|f|=C17|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|E17|f|=D17|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|F17|f|=E17|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|G17|f|=F17|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|H17|f|=G17|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|I17|f|=H17|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|J17|f|=I17|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|K17|f|=J17|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|L17|f|=K17|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|M17|f|=L17|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|N17|f|=M17+(M17*$E$7)|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|O17|f|=N17|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|P17|f|=O17|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|Q17|f|=P17|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|R17|f|=Q17|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|S17|f|=R17|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|T17|f|=S17|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|U17|f|=T17|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|V17|f|=U17|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|W17|f|=V17|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|X17|f|=W17|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|Y17|f|=X17|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|Z17|f|=Y17+(Y17*$E$8)|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|AA17|f|=Z17|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|AB17|f|=AA17|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|AC17|f|=AB17|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|A18|s|Constable|General|0|0|9.0||15917529||0|1"
    AddST "C|B18|f|=(M7*B6)+(M7*E6)|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|C18|f|=B18|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|D18|f|=C18|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|E18|f|=D18|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|F18|f|=E18|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|G18|f|=F18|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|H18|f|=G18|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|I18|f|=H18|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|J18|f|=I18|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|K18|f|=J18|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|L18|f|=K18|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|M18|f|=L18|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|N18|f|=M18+(M18*$E$7)|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|O18|f|=N18|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|P18|f|=O18|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|Q18|f|=P18|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|R18|f|=Q18|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|S18|f|=R18|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|T18|f|=S18|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|U18|f|=T18|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|V18|f|=U18|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|W18|f|=V18|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|X18|f|=W18|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|Y18|f|=X18|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|Z18|f|=Y18+(Y18*$E$8)|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|AA18|f|=Z18|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|AB18|f|=AA18|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|AC18|f|=AB18|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|A21|s|COMMAND TEAM & PLANNING|General|1|0|9.0|192|15917529||0|0"
    AddST "C|B22|s|Start|General|1|0|9.0||7159831|center|0|1"
    AddST "C|C22|s|End|General|1|0|9.0||7159831|center|0|1"
    AddST "C|E22|d|2026-09-01|mmm-yy|1|0|9.0||7159831|center|0|1"
    AddST "C|F22|d|2026-10-01|mmm-yy|1|0|9.0||7159831|center|0|1"
    AddST "C|G22|d|2026-11-01|mmm-yy|1|0|9.0||7159831|center|0|1"
    AddST "C|H22|d|2026-12-01|mmm-yy|1|0|9.0||7159831|center|0|1"
    AddST "C|I22|d|2027-01-01|mmm-yy|1|0|9.0||7159831|center|0|1"
    AddST "C|J22|d|2027-02-01|mmm-yy|1|0|9.0||7159831|center|0|1"
    AddST "C|K22|d|2027-03-01|mmm-yy|1|0|9.0||7159831|center|0|1"
    AddST "C|L22|d|2027-04-01|mmm-yy|1|0|9.0||7159831|center|0|1"
    AddST "C|M22|d|2027-05-01|mmm-yy|1|0|9.0||7159831|center|0|1"
    AddST "C|N22|d|2027-06-01|mmm-yy|1|0|9.0||7159831|center|0|1"
    AddST "C|O22|d|2027-07-01|mmm-yy|1|0|9.0||7159831|center|0|1"
    AddST "C|P22|d|2027-08-01|mmm-yy|1|0|9.0||7159831|center|0|1"
    AddST "C|Q22|d|2027-09-01|mmm-yy|1|0|9.0||7159831|center|0|1"
    AddST "C|R22|d|2027-10-01|mmm-yy|1|0|9.0||7159831|center|0|1"
    AddST "C|S22|d|2027-11-01|mmm-yy|1|0|9.0||7159831|center|0|1"
    AddST "C|T22|d|2027-12-01|mmm-yy|1|0|9.0||7159831|center|0|1"
    AddST "C|U22|d|2028-01-01|mmm-yy|1|0|9.0||7159831|center|0|1"
    AddST "C|V22|d|2028-02-01|mmm-yy|1|0|9.0||7159831|center|0|1"
    AddST "C|W22|d|2028-03-01|mmm-yy|1|0|9.0||7159831|center|0|1"
    AddST "C|X22|d|2028-04-01|mmm-yy|1|0|9.0||7159831|center|0|1"
    AddST "C|Y22|d|2028-05-01|mmm-yy|1|0|9.0||7159831|center|0|1"
    AddST "C|Z22|d|2028-06-01|mmm-yy|1|0|9.0||7159831|center|0|1"
    AddST "C|AA22|d|2028-07-01|mmm-yy|1|0|9.0||7159831|center|0|1"
    AddST "C|AB22|d|2028-08-01|mmm-yy|1|0|9.0||7159831|center|0|1"
    AddST "C|AC22|d|2028-09-01|mmm-yy|1|0|9.0||7159831|center|0|1"
    AddST "C|AD22|d|2028-10-01|mmm-yy|1|0|9.0||7159831|center|0|1"
    AddST "C|AE22|d|2028-11-01|mmm-yy|1|0|9.0||7159831|center|0|1"
    AddST "C|AF22|d|2028-12-01|mmm-yy|1|0|9.0||7159831|center|0|1"
    AddST "C|A23|s|C/Superintendent|General|0|0|9.0||15917529||0|1"
    AddST "C|B23|e||mmm-yy|0|0|9.0||15917529||0|1"
    AddST "C|C23|e||mmm-yy|0|0|9.0||15917529||0|1"
    AddST "C|E23|f|=IF(B23=$E$22, B13, 0)|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|F23|f|=IF(C23=$E$22,0,IF(E23>0,C13,IF(B23=$F$22,C13,IF(C23=$E$22,0,0))))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|G23|f|=IF(C23=$F$22,0,IF(F23>0,D13,IF(B23=$G$22,D13,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|H23|f|=IF(C23=$G$22,0,IF(G23>0,E13,IF(B23=$H$22,E13,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|I23|f|=IF(C23=$H$22,0,IF(H23>0,F13,IF(B23=$I$22,F13,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|J23|f|=IF(C23=$I$22,0,IF(I23>0,G13,IF(B23=$J$22,G13,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|K23|f|=IF(C23=$J$22,0,IF(J23>0,H13,IF(B23=$K$22,H13,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|L23|f|=IF(C23=$K$22,0,IF(K23>0,I13,IF(B23=$L$22,I13,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|M23|f|=IF(C23=$L$22,0,IF(L23>0,L23,IF(B23=$M$22,J13,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|N23|f|=IF(C23=$M$22,0,IF(M23>0,M23,IF(B23=$N$22,K13,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|O23|f|=IF(C23=$N$22,0,IF(N23>0,L13,IF(B23=$O$22,L13,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|P23|f|=IF(C23=$O$22,0,IF(O23>0,M13,IF(B23=$P$22,M13,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|Q23|f|=IF(C23=$P$22,0,IF(P23>0,N13,IF(B23=$Q$22,N13,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|R23|f|=IF(C23=$Q$22,0,IF(Q23>0,O13,IF(B23=$R$22,O13,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|S23|f|=IF(C23=$R$22,0,IF(R23>0,P13,IF(B23=$S$22,P13,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|T23|f|=IF(C23=$S$22,0,IF(S23>0,Q13,IF(B23=$T$22,Q13,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
End Sub

Private Sub DataST1_3()
    Dim s As String
    AddST "C|U23|f|=IF(C23=$T$22,0,IF(T23>0,R13,IF(B23=$U$22,R13,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|V23|f|=IF(C23=$U$22,0,IF(U23>0,S13,IF(B23=$V$22,S13,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|W23|f|=IF(C23=$V$22,0,IF(V23>0,T13,IF(B23=$W$22,T13,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|X23|f|=IF(C23=$W$22,0,IF(W23>0,U13,IF(B23=$X$22,U13,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|Y23|f|=IF(C23=$X$22,0,IF(X23>0,V13,IF(B23=$Y$22,V13,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|Z23|f|=IF(C23=$Y$22,0,IF(Y23>0,W13,IF(B23=$Z$22,W13,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|AA23|f|=IF(C23=$Z$22,0,IF(Z23>0,X13,IF(B23=$AA$22,X13,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|AB23|f|=IF(C23=$AA$22,0,IF(AA23>0,Y13,IF(B23=$AB$22,Y13,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|AC23|f|=IF(C23=$AB$18,0,IF(AB23>0,Z13,IF(B23=$AC$18,Z13,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|AD23|f|=IF(C23=$AC$22,0,IF(AC23>0,AA13,IF(B23=$AD$22,AA13,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|AE23|f|=IF(C23=$AD$22,0,IF(AD23>0,AB13,IF(B23=$AE$22,AB13,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|AF23|f|=IF(C23=$AE$22,0,IF(AE23>0,AC13,IF(B23=$AF$22,AC13,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|A24|s|Superintendent|General|0|0|9.0||15917529||0|1"
    AddST "C|B24|e||mmm-yy|0|0|9.0||15917529||0|1"
    AddST "C|C24|e||mmm-yy|0|0|9.0||15917529||0|1"
    AddST "C|E24|f|=IF(B24=$E$22, B14, 0)|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|F24|f|=IF(C24=$E$22,0,IF(E24>0,C14,IF(B24=$F$22,C14,IF(C24=$E$22,0,0))))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|G24|f|=IF(C24=$F$22,0,IF(F24>0,D14,IF(B24=$G$22,D14,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|H24|f|=IF(C24=$G$22,0,IF(G24>0,E14,IF(B24=$H$22,E14,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|I24|f|=IF(C24=$H$22,0,IF(H24>0,F14,IF(B24=$I$22,F14,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|J24|f|=IF(C24=$I$22,0,IF(I24>0,G14,IF(B24=$J$22,G14,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|K24|f|=IF(C24=$J$22,0,IF(J24>0,H14,IF(B24=$K$22,H14,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|L24|f|=IF(C24=$K$22,0,IF(K24>0,I14,IF(B24=$L$22,I14,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|M24|f|=IF(C24=$L$22,0,IF(L24>0,L24,IF(B24=$M$22,J14,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|N24|f|=IF(C24=$M$22,0,IF(M24>0,M24,IF(B24=$N$22,K14,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|O24|f|=IF(C24=$N$22,0,IF(N24>0,L14,IF(B24=$O$22,L14,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|P24|f|=IF(C24=$O$22,0,IF(O24>0,M14,IF(B24=$P$22,M14,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|Q24|f|=IF(C24=$P$22,0,IF(P24>0,N14,IF(B24=$Q$22,N14,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|R24|f|=IF(C24=$Q$22,0,IF(Q24>0,O14,IF(B24=$R$22,O14,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|S24|f|=IF(C24=$R$22,0,IF(R24>0,P14,IF(B24=$S$22,P14,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|T24|f|=IF(C24=$S$22,0,IF(S24>0,Q14,IF(B24=$T$22,Q14,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|U24|f|=IF(C24=$T$22,0,IF(T24>0,R14,IF(B24=$U$22,R14,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|V24|f|=IF(C24=$U$22,0,IF(U24>0,S14,IF(B24=$V$22,S14,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|W24|f|=IF(C24=$V$22,0,IF(V24>0,T14,IF(B24=$W$22,T14,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|X24|f|=IF(C24=$W$22,0,IF(W24>0,U14,IF(B24=$X$22,U14,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|Y24|f|=IF(C24=$X$22,0,IF(X24>0,V14,IF(B24=$Y$22,V14,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|Z24|f|=IF(C24=$Y$22,0,IF(Y24>0,W14,IF(B24=$Z$22,W14,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|AA24|f|=IF(C24=$Z$22,0,IF(Z24>0,X14,IF(B24=$AA$22,X14,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|AB24|f|=IF(C24=$AA$22,0,IF(AA24>0,Y14,IF(B24=$AB$22,Y14,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|AC24|f|=IF(C24=$AB$18,0,IF(AB24>0,Z14,IF(B24=$AC$18,Z14,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|AD24|f|=IF(C24=$AC$22,0,IF(AC24>0,AA14,IF(B24=$AD$22,AA14,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|AE24|f|=IF(C24=$AD$22,0,IF(AD24>0,AB14,IF(B24=$AE$22,AB14,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|AF24|f|=IF(C24=$AE$22,0,IF(AE24>0,AC14,IF(B24=$AF$22,AC14,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|A25|s|Chief Inspector|General|0|0|9.0||15917529||0|1"
    AddST "C|B25|e||mmm-yy|0|0|9.0||15917529||0|1"
    AddST "C|C25|e||mmm-yy|0|0|9.0||15917529||0|1"
    AddST "C|E25|f|=IF(B25=$E$22, B15, 0)|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|F25|f|=IF(C25=$E$22,0,IF(E25>0,C15,IF(B25=$F$22,C15,IF(C25=$E$22,0,0))))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|G25|f|=IF(C25=$F$22,0,IF(F25>0,D15,IF(B25=$G$22,D15,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|H25|f|=IF(C25=$G$22,0,IF(G25>0,E15,IF(B25=$H$22,E15,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|I25|f|=IF(C25=$H$22,0,IF(H25>0,F15,IF(B25=$I$22,F15,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|J25|f|=IF(C25=$I$22,0,IF(I25>0,G15,IF(B25=$J$22,G15,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|K25|f|=IF(C25=$J$22,0,IF(J25>0,H15,IF(B25=$K$22,H15,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|L25|f|=IF(C25=$K$22,0,IF(K25>0,I15,IF(B25=$L$22,I15,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|M25|f|=IF(C25=$L$22,0,IF(L25>0,L25,IF(B25=$M$22,J15,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|N25|f|=IF(C25=$M$22,0,IF(M25>0,M25,IF(B25=$N$22,K15,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|O25|f|=IF(C25=$N$22,0,IF(N25>0,L15,IF(B25=$O$22,L15,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|P25|f|=IF(C25=$O$22,0,IF(O25>0,M15,IF(B25=$P$22,M15,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|Q25|f|=IF(C25=$P$22,0,IF(P25>0,N15,IF(B25=$Q$22,N15,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|R25|f|=IF(C25=$Q$22,0,IF(Q25>0,O15,IF(B25=$R$22,O15,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|S25|f|=IF(C25=$R$22,0,IF(R25>0,P15,IF(B25=$S$22,P15,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|T25|f|=IF(C25=$S$22,0,IF(S25>0,Q15,IF(B25=$T$22,Q15,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|U25|f|=IF(C25=$T$22,0,IF(T25>0,R15,IF(B25=$U$22,R15,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|V25|f|=IF(C25=$U$22,0,IF(U25>0,S15,IF(B25=$V$22,S15,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|W25|f|=IF(C25=$V$22,0,IF(V25>0,T15,IF(B25=$W$22,T15,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|X25|f|=IF(C25=$W$22,0,IF(W25>0,U15,IF(B25=$X$22,U15,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|Y25|f|=IF(C25=$X$22,0,IF(X25>0,V15,IF(B25=$Y$22,V15,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|Z25|f|=IF(C25=$Y$22,0,IF(Y25>0,W15,IF(B25=$Z$22,W15,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|AA25|f|=IF(C25=$Z$22,0,IF(Z25>0,X15,IF(B25=$AA$22,X15,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|AB25|f|=IF(C25=$AA$22,0,IF(AA25>0,Y15,IF(B25=$AB$22,Y15,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|AC25|f|=IF(C25=$AB$18,0,IF(AB25>0,Z15,IF(B25=$AC$18,Z15,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|AD25|f|=IF(C25=$AC$22,0,IF(AC25>0,AA15,IF(B25=$AD$22,AA15,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|AE25|f|=IF(C25=$AD$22,0,IF(AD25>0,AB15,IF(B25=$AE$22,AB15,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|AF25|f|=IF(C25=$AE$22,0,IF(AE25>0,AC15,IF(B25=$AF$22,AC15,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|A26|s|Inspector|General|0|0|9.0||15917529||0|1"
    AddST "C|B26|e||mmm-yy|0|0|9.0||15917529||0|1"
    AddST "C|C26|e||mmm-yy|0|0|9.0||15917529||0|1"
    AddST "C|E26|f|=IF(B26=$E$22, B16, 0)|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|F26|f|=IF(C26=$E$22,0,IF(E26>0,C16,IF(B26=$F$22,C16,IF(C26=$E$22,0,0))))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|G26|f|=IF(C26=$F$22,0,IF(F26>0,D16,IF(B26=$G$22,D16,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|H26|f|=IF(C26=$G$22,0,IF(G26>0,E16,IF(B26=$H$22,E16,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|I26|f|=IF(C26=$H$22,0,IF(H26>0,F16,IF(B26=$I$22,F16,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|J26|f|=IF(C26=$I$22,0,IF(I26>0,G16,IF(B26=$J$22,G16,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|K26|f|=IF(C26=$J$22,0,IF(J26>0,H16,IF(B26=$K$22,H16,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|L26|f|=IF(C26=$K$22,0,IF(K26>0,I16,IF(B26=$L$22,I16,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|M26|f|=IF(C26=$L$22,0,IF(L26>0,L26,IF(B26=$M$22,J16,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|N26|f|=IF(C26=$M$22,0,IF(M26>0,M26,IF(B26=$N$22,K16,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|O26|f|=IF(C26=$N$22,0,IF(N26>0,L16,IF(B26=$O$22,L16,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|P26|f|=IF(C26=$O$22,0,IF(O26>0,M16,IF(B26=$P$22,M16,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|Q26|f|=IF(C26=$P$22,0,IF(P26>0,N16,IF(B26=$Q$22,N16,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|R26|f|=IF(C26=$Q$22,0,IF(Q26>0,O16,IF(B26=$R$22,O16,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|S26|f|=IF(C26=$R$22,0,IF(R26>0,P16,IF(B26=$S$22,P16,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|T26|f|=IF(C26=$S$22,0,IF(S26>0,Q16,IF(B26=$T$22,Q16,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|U26|f|=IF(C26=$T$22,0,IF(T26>0,R16,IF(B26=$U$22,R16,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|V26|f|=IF(C26=$U$22,0,IF(U26>0,S16,IF(B26=$V$22,S16,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|W26|f|=IF(C26=$V$22,0,IF(V26>0,T16,IF(B26=$W$22,T16,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|X26|f|=IF(C26=$W$22,0,IF(W26>0,U16,IF(B26=$X$22,U16,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|Y26|f|=IF(C26=$X$22,0,IF(X26>0,V16,IF(B26=$Y$22,V16,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|Z26|f|=IF(C26=$Y$22,0,IF(Y26>0,W16,IF(B26=$Z$22,W16,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|AA26|f|=IF(C26=$Z$22,0,IF(Z26>0,X16,IF(B26=$AA$22,X16,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|AB26|f|=IF(C26=$AA$22,0,IF(AA26>0,Y16,IF(B26=$AB$22,Y16,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|AC26|f|=IF(C26=$AB$18,0,IF(AB26>0,Z16,IF(B26=$AC$18,Z16,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|AD26|f|=IF(C26=$AC$22,0,IF(AC26>0,AA16,IF(B26=$AD$22,AA16,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|AE26|f|=IF(C26=$AD$22,0,IF(AD26>0,AB16,IF(B26=$AE$22,AB16,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|AF26|f|=IF(C26=$AE$22,0,IF(AE26>0,AC16,IF(B26=$AF$22,AC16,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|A27|s|Sergeant|General|0|0|9.0||15917529||0|1"
    AddST "C|B27|e||mmm-yy|0|0|9.0||15917529||0|1"
    AddST "C|C27|e||mmm-yy|0|0|9.0||15917529||0|1"
    AddST "C|E27|f|=IF(B27=$E$22, B17, 0)|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|F27|f|=IF(C27=$E$22,0,IF(E27>0,C17,IF(B27=$F$22,C17,IF(C27=$E$22,0,0))))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|G27|f|=IF(C27=$F$22,0,IF(F27>0,D17,IF(B27=$G$22,D17,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|H27|f|=IF(C27=$G$22,0,IF(G27>0,E17,IF(B27=$H$22,E17,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|I27|f|=IF(C27=$H$22,0,IF(H27>0,F17,IF(B27=$I$22,F17,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|J27|f|=IF(C27=$I$22,0,IF(I27>0,G17,IF(B27=$J$22,G17,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|K27|f|=IF(C27=$J$22,0,IF(J27>0,H17,IF(B27=$K$22,H17,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|L27|f|=IF(C27=$K$22,0,IF(K27>0,I17,IF(B27=$L$22,I17,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|M27|f|=IF(C27=$L$22,0,IF(L27>0,L27,IF(B27=$M$22,J17,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|N27|f|=IF(C27=$M$22,0,IF(M27>0,M27,IF(B27=$N$22,K17,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|O27|f|=IF(C27=$N$22,0,IF(N27>0,L17,IF(B27=$O$22,L17,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|P27|f|=IF(C27=$O$22,0,IF(O27>0,M17,IF(B27=$P$22,M17,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|Q27|f|=IF(C27=$P$22,0,IF(P27>0,N17,IF(B27=$Q$22,N17,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|R27|f|=IF(C27=$Q$22,0,IF(Q27>0,O17,IF(B27=$R$22,O17,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|S27|f|=IF(C27=$R$22,0,IF(R27>0,P17,IF(B27=$S$22,P17,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|T27|f|=IF(C27=$S$22,0,IF(S27>0,Q17,IF(B27=$T$22,Q17,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|U27|f|=IF(C27=$T$22,0,IF(T27>0,R17,IF(B27=$U$22,R17,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|V27|f|=IF(C27=$U$22,0,IF(U27>0,S17,IF(B27=$V$22,S17,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|W27|f|=IF(C27=$V$22,0,IF(V27>0,T17,IF(B27=$W$22,T17,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|X27|f|=IF(C27=$W$22,0,IF(W27>0,U17,IF(B27=$X$22,U17,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|Y27|f|=IF(C27=$X$22,0,IF(X27>0,V17,IF(B27=$Y$22,V17,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|Z27|f|=IF(C27=$Y$22,0,IF(Y27>0,W17,IF(B27=$Z$22,W17,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|AA27|f|=IF(C27=$Z$22,0,IF(Z27>0,X17,IF(B27=$AA$22,X17,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|AB27|f|=IF(C27=$AA$22,0,IF(AA27>0,Y17,IF(B27=$AB$22,Y17,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|AC27|f|=IF(C27=$AB$18,0,IF(AB27>0,Z17,IF(B27=$AC$18,Z17,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|AD27|f|=IF(C27=$AC$22,0,IF(AC27>0,AA17,IF(B27=$AD$22,AA17,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|AE27|f|=IF(C27=$AD$22,0,IF(AD27>0,AB17,IF(B27=$AE$22,AB17,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|AF27|f|=IF(C27=$AE$22,0,IF(AE27>0,AC17,IF(B27=$AF$22,AC17,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|A28|s|Constable|General|0|0|9.0||15917529||0|1"
    AddST "C|B28|e||mmm-yy|0|0|9.0||15917529||0|1"
    AddST "C|C28|e||mmm-yy|0|0|9.0||15917529||0|1"
    AddST "C|E28|f|=IF(B28=$E$22, B18, 0)|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|F28|f|=IF(C28=$E$22,0,IF(E28>0,C18,IF(B28=$F$22,C18,IF(C28=$E$22,0,0))))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|G28|f|=IF(C28=$F$22,0,IF(F28>0,D18,IF(B28=$G$22,D18,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|H28|f|=IF(C28=$G$22,0,IF(G28>0,E18,IF(B28=$H$22,E18,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|I28|f|=IF(C28=$H$22,0,IF(H28>0,F18,IF(B28=$I$22,F18,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|J28|f|=IF(C28=$I$22,0,IF(I28>0,G18,IF(B28=$J$22,G18,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|K28|f|=IF(C28=$J$22,0,IF(J28>0,H18,IF(B28=$K$22,H18,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|L28|f|=IF(C28=$K$22,0,IF(K28>0,I18,IF(B28=$L$22,I18,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|M28|f|=IF(C28=$L$22,0,IF(L28>0,L28,IF(B28=$M$22,J18,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|N28|f|=IF(C28=$M$22,0,IF(M28>0,M28,IF(B28=$N$22,K18,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|O28|f|=IF(C28=$N$22,0,IF(N28>0,L18,IF(B28=$O$22,L18,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
End Sub

Private Sub DataST1_4()
    Dim s As String
    AddST "C|P28|f|=IF(C28=$O$22,0,IF(O28>0,M18,IF(B28=$P$22,M18,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|Q28|f|=IF(C28=$P$22,0,IF(P28>0,N18,IF(B28=$Q$22,N18,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|R28|f|=IF(C28=$Q$22,0,IF(Q28>0,O18,IF(B28=$R$22,O18,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|S28|f|=IF(C28=$R$22,0,IF(R28>0,P18,IF(B28=$S$22,P18,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|T28|f|=IF(C28=$S$22,0,IF(S28>0,Q18,IF(B28=$T$22,Q18,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|U28|f|=IF(C28=$T$22,0,IF(T28>0,R18,IF(B28=$U$22,R18,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|V28|f|=IF(C28=$U$22,0,IF(U28>0,S18,IF(B28=$V$22,S18,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|W28|f|=IF(C28=$V$22,0,IF(V28>0,T18,IF(B28=$W$22,T18,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|X28|f|=IF(C28=$W$22,0,IF(W28>0,U18,IF(B28=$X$22,U18,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|Y28|f|=IF(C28=$X$22,0,IF(X28>0,V18,IF(B28=$Y$22,V18,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|Z28|f|=IF(C28=$Y$22,0,IF(Y28>0,W18,IF(B28=$Z$22,W18,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|AA28|f|=IF(C28=$Z$22,0,IF(Z28>0,X18,IF(B28=$AA$22,X18,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|AB28|f|=IF(C28=$AA$22,0,IF(AA28>0,Y18,IF(B28=$AB$22,Y18,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|AC28|f|=IF(C28=$AB$18,0,IF(AB28>0,Z18,IF(B28=$AC$18,Z18,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|AD28|f|=IF(C28=$AC$22,0,IF(AC28>0,AA18,IF(B28=$AD$22,AA18,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|AE28|f|=IF(C28=$AD$22,0,IF(AD28>0,AB18,IF(B28=$AE$22,AB18,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|AF28|f|=IF(C28=$AE$22,0,IF(AE28>0,AC18,IF(B28=$AF$22,AC18,0)))|""~uA3~""#,##0.00|0|0|9.0||15917529||0|1"
    AddST "C|H29|e||0|0|0|9.0||15917529|center|0|0"
    AddST "C|I29|e||0|0|0|9.0||15917529|center|0|0"
    AddST "C|J29|e||0|0|0|9.0||15917529|center|0|0"
    AddST "C|K29|e||0|0|0|9.0||15917529|center|0|0"
    AddST "C|L29|e||0|0|0|9.0||15917529|center|0|0"
    AddST "W|A|14.83203125"
    AddST "W|D|14.83203125"
    AddST "W|M|14.83203125"
    AddST "W|AH|9.33203125"
    AddST "W|EI|9.33203125"
    AddST "M|A2:L2"
End Sub

Private Sub DataST2_1()
    Dim s As String
    AddST "C|A3|e||General|1|0|9.0||7159831|center|0|1"
    AddST "C|C3|s|Rank|General|1|0|9.0||7159831|center|0|1"
    AddST "C|D3|s|Total Hours Required in 2027|d-mmm|1|0|9.0||7159831|center|1|1"
    AddST "C|E3|s|2027 Cost|d-mmm|1|0|9.0||7159831|center|1|1"
    AddST "C|F3|s|Total Hours Required in 2028|d-mmm|1|0|9.0||7159831|center|1|1"
    AddST "C|G3|s|2028 Cost|d-mmm|1|0|9.0||7159831|center|1|1"
    AddST "C|H3|s|Total Cost|d-mmm|1|0|9.0||7159831|center|1|1"
    AddST "C|A6|s|Media Management|General|1|0|9.0||7159831|center|1|1"
    AddST "C|D6|e||General|1|0|9.0||7159831|center|0|0"
    AddST "C|E6|e||General|1|0|9.0||7159831|center|0|0"
    AddST "C|F6|e||General|1|0|9.0||7159831|center|0|0"
    AddST "C|G6|e||General|1|0|9.0||7159831|center|0|0"
    AddST "C|H6|e||General|1|0|9.0||7159831|center|0|0"
    AddST "C|A7|s|Overtime|General|0|0|9.0|||center|0|1"
    AddST "C|C7|s|Sergeant|General|0|0|9.0||||0|1"
    AddST "C|D7|e||General|0|0|9.0|||center|0|1"
    AddST "C|E7|f|=IFERROR(SUM(D7*'Baseline Costs'!$L$31),0)|""~uA3~""#,##0.00|0|0|9.0|||center|0|1"
    AddST "C|F7|e||General|0|0|9.0|||center|0|1"
    AddST "C|G7|f|=IFERROR(SUM(F7*'Baseline Costs'!$X$31),0)|""~uA3~""#,##0.00|0|0|9.0|||center|0|1"
    AddST "C|H7|f|=E7+G7|""~uA3~""#,##0.00|0|0|9.0|||center|0|1"
    AddST "C|C8|s|Constable|General|0|0|9.0||||0|1"
    AddST "C|D8|e||General|0|0|9.0|||center|0|1"
    AddST "C|E8|f|=IFERROR(SUM(D8*'Baseline Costs'!$L$32),0)|""~uA3~""#,##0.00|0|0|9.0|||center|0|1"
    AddST "C|F8|e||General|0|0|9.0|||center|0|1"
    AddST "C|G8|f|=IFERROR(SUM(F8*'Baseline Costs'!$X$32),0)|""~uA3~""#,##0.00|0|0|9.0|||center|0|1"
    AddST "C|H8|f|=E8+G8|""~uA3~""#,##0.00|0|0|9.0|||center|0|1"
    AddST "C|A9|s|Licensing|General|1|0|9.0||7159831|center|1|1"
    AddST "C|D9|e||General|1|0|9.0||7159831|center|0|0"
    AddST "C|E9|e||General|1|0|9.0||7159831|center|0|0"
    AddST "C|F9|e||General|1|0|9.0||7159831|center|0|0"
    AddST "C|G9|e||General|1|0|9.0||7159831|center|0|0"
    AddST "C|H9|e||General|1|0|9.0||7159831|center|0|0"
    AddST "C|A10|s|Overtime|General|0|0|9.0|||center|0|1"
    AddST "C|C10|s|Sergeant|General|0|0|9.0||||0|1"
    AddST "C|D10|e||General|0|0|9.0|||center|0|1"
    AddST "C|E10|f|=IFERROR(SUM(D10*'Baseline Costs'!$L$31),0)|""~uA3~""#,##0.00|0|0|9.0|||center|0|1"
    AddST "C|F10|e||General|0|0|9.0|||center|0|1"
    AddST "C|G10|f|=IFERROR(SUM(F10*'Baseline Costs'!$X$31),0)|""~uA3~""#,##0.00|0|0|9.0|||center|0|1"
    AddST "C|H10|f|=E10+G10|""~uA3~""#,##0.00|0|0|9.0|||center|0|1"
    AddST "C|C11|s|Constable|General|0|0|9.0||||0|1"
    AddST "C|D11|e||General|0|0|9.0|||center|0|1"
    AddST "C|E11|f|=IFERROR(SUM(D11*'Baseline Costs'!$L$32),0)|""~uA3~""#,##0.00|0|0|9.0|||center|0|1"
    AddST "C|F11|e||General|0|0|9.0|||center|0|1"
    AddST "C|G11|f|=IFERROR(SUM(F11*'Baseline Costs'!$X$32),0)|""~uA3~""#,##0.00|0|0|9.0|||center|0|1"
    AddST "C|H11|f|=E11+G11|""~uA3~""#,##0.00|0|0|9.0|||center|0|1"
    AddST "C|A12|s|Crime Investigation|General|1|0|9.0||7159831|center|1|1"
    AddST "C|D12|e||General|1|0|9.0||7159831|center|0|0"
    AddST "C|E12|e||General|1|0|9.0||7159831|center|0|0"
    AddST "C|F12|e||General|1|0|9.0||7159831|center|0|0"
    AddST "C|G12|e||General|1|0|9.0||7159831|center|0|0"
    AddST "C|H12|e||General|1|0|9.0||7159831|center|0|0"
    AddST "C|A13|s|Overtime|General|0|0|9.0|||center|0|1"
    AddST "C|C13|s|Sergeant|General|0|0|9.0||||0|1"
    AddST "C|D13|e||General|0|0|9.0|||center|0|1"
    AddST "C|E13|f|=IFERROR(SUM(D13*'Baseline Costs'!$L$31),0)|""~uA3~""#,##0.00|0|0|9.0|||center|0|1"
    AddST "C|F13|e||General|0|0|9.0|||center|0|1"
    AddST "C|G13|f|=IFERROR(SUM(F13*'Baseline Costs'!$X$31),0)|""~uA3~""#,##0.00|0|0|9.0|||center|0|1"
    AddST "C|H13|f|=E13+G13|""~uA3~""#,##0.00|0|0|9.0|||center|0|1"
    AddST "C|C14|s|Constable|General|0|0|9.0||||0|1"
    AddST "C|D14|e||General|0|0|9.0|||center|0|1"
    AddST "C|E14|f|=IFERROR(SUM(D14*'Baseline Costs'!$L$32),0)|""~uA3~""#,##0.00|0|0|9.0|||center|0|1"
    AddST "C|F14|e||General|0|0|9.0|||center|0|1"
    AddST "C|G14|f|=IFERROR(SUM(F14*'Baseline Costs'!$X$32),0)|""~uA3~""#,##0.00|0|0|9.0|||center|0|1"
    AddST "C|H14|f|=E14+G14|""~uA3~""#,##0.00|0|0|9.0|||center|0|1"
    AddST "C|A15|s|Intelligence Arrangements|General|1|0|9.0||7159831|center|1|1"
    AddST "C|D15|e||General|1|0|9.0||7159831|center|0|0"
    AddST "C|E15|e||General|1|0|9.0||7159831|center|0|0"
    AddST "C|F15|e||General|1|0|9.0||7159831|center|0|0"
    AddST "C|G15|e||General|1|0|9.0||7159831|center|0|0"
    AddST "C|H15|e||General|1|0|9.0||7159831|center|0|0"
    AddST "C|A16|s|Overtime|General|0|0|9.0|||center|0|1"
    AddST "C|C16|s|Sergeant|General|0|0|9.0||||0|1"
    AddST "C|D16|e||General|0|0|9.0|||center|0|1"
    AddST "C|E16|f|=IFERROR(SUM(D16*'Baseline Costs'!$L$31),0)|""~uA3~""#,##0.00|0|0|9.0|||center|0|1"
    AddST "C|F16|e||General|0|0|9.0|||center|0|1"
    AddST "C|G16|f|=IFERROR(SUM(F16*'Baseline Costs'!$X$31),0)|""~uA3~""#,##0.00|0|0|9.0|||center|0|1"
    AddST "C|H16|f|=E16+G16|""~uA3~""#,##0.00|0|0|9.0|||center|0|1"
    AddST "C|C17|s|Constable|General|0|0|9.0||||0|1"
    AddST "C|D17|e||General|0|0|9.0|||center|0|1"
    AddST "C|E17|f|=IFERROR(SUM(D17*'Baseline Costs'!$L$32),0)|""~uA3~""#,##0.00|0|0|9.0|||center|0|1"
    AddST "C|F17|e||General|0|0|9.0|||center|0|1"
    AddST "C|G17|f|=IFERROR(SUM(F17*'Baseline Costs'!$X$32),0)|""~uA3~""#,##0.00|0|0|9.0|||center|0|1"
    AddST "C|H17|f|=E17+G17|""~uA3~""#,##0.00|0|0|9.0|||center|0|1"
    AddST "C|A18|s|Post Event Investigation|General|1|0|9.0||7159831|center|1|1"
    AddST "C|D18|e||General|1|0|9.0||7159831|center|0|0"
    AddST "C|E18|e||General|1|0|9.0||7159831|center|0|0"
    AddST "C|F18|e||General|1|0|9.0||7159831|center|0|0"
    AddST "C|G18|e||General|1|0|9.0||7159831|center|0|0"
    AddST "C|H18|e||General|1|0|9.0||7159831|center|0|0"
    AddST "C|A19|s|Overtime|General|0|0|9.0|||center|0|1"
    AddST "C|C19|s|Sergeant|General|0|0|9.0||||0|1"
    AddST "C|D19|e||General|0|0|9.0|||center|0|1"
    AddST "C|E19|f|=IFERROR(SUM(D19*'Baseline Costs'!$L$31),0)|""~uA3~""#,##0.00|0|0|9.0|||center|0|1"
    AddST "C|F19|e||General|0|0|9.0|||center|0|1"
    AddST "C|G19|f|=IFERROR(SUM(F19*'Baseline Costs'!$X$31),0)|""~uA3~""#,##0.00|0|0|9.0|||center|0|1"
    AddST "C|H19|f|=E19+G19|""~uA3~""#,##0.00|0|0|9.0|||center|0|1"
    AddST "C|C20|s|Constable|General|0|0|9.0||||0|1"
    AddST "C|D20|e||General|0|0|9.0|||center|0|1"
    AddST "C|E20|f|=IFERROR(SUM(D20*'Baseline Costs'!$L$32),0)|""~uA3~""#,##0.00|0|0|9.0|||center|0|1"
    AddST "C|F20|e||General|0|0|9.0|||center|0|1"
    AddST "C|G20|f|=IFERROR(SUM(F20*'Baseline Costs'!$X$32),0)|""~uA3~""#,##0.00|0|0|9.0|||center|0|1"
    AddST "C|H20|f|=E20+G20|""~uA3~""#,##0.00|0|0|9.0|||center|0|1"
    AddST "C|A21|s|Professional Standards & Legal|General|1|0|9.0||7159831|center|1|1"
    AddST "C|D21|e||General|1|0|9.0||7159831|center|0|0"
    AddST "C|E21|e||General|1|0|9.0||7159831|center|0|0"
    AddST "C|F21|e||General|1|0|9.0||7159831|center|0|0"
    AddST "C|G21|e||General|1|0|9.0||7159831|center|0|0"
    AddST "C|H21|e||General|1|0|9.0||7159831|center|0|0"
    AddST "C|A22|s|Overtime|General|0|0|9.0|||center|0|1"
    AddST "C|C22|s|Sergeant|General|0|0|9.0||||0|1"
    AddST "C|D22|e||General|0|0|9.0|||center|0|1"
    AddST "C|E22|f|=IFERROR(SUM(D22*'Baseline Costs'!$L$31),0)|""~uA3~""#,##0.00|0|0|9.0|||center|0|1"
    AddST "C|F22|e||General|0|0|9.0|||center|0|1"
    AddST "C|G22|f|=IFERROR(SUM(F22*'Baseline Costs'!$X$31),0)|""~uA3~""#,##0.00|0|0|9.0|||center|0|1"
    AddST "C|H22|f|=E22+G22|""~uA3~""#,##0.00|0|0|9.0|||center|0|1"
    AddST "C|C23|s|Constable|General|0|0|9.0||||0|1"
    AddST "C|D23|e||General|0|0|9.0|||center|0|1"
    AddST "C|E23|f|=IFERROR(SUM(D23*'Baseline Costs'!$L$32),0)|""~uA3~""#,##0.00|0|0|9.0|||center|0|1"
    AddST "C|F23|e||General|0|0|9.0|||center|0|1"
    AddST "C|G23|f|=IFERROR(SUM(F23*'Baseline Costs'!$X$32),0)|""~uA3~""#,##0.00|0|0|9.0|||center|0|1"
    AddST "C|H23|f|=E23+G23|""~uA3~""#,##0.00|0|0|9.0|||center|0|1"
    AddST "C|A24|s|Command Team & Planning|General|1|0|9.0||7159831|center|1|1"
    AddST "C|D24|e||General|1|0|9.0||7159831|center|0|0"
    AddST "C|E24|e||General|1|0|9.0||7159831|center|0|0"
    AddST "C|F24|e||General|1|0|9.0||7159831|center|0|0"
    AddST "C|G24|e||General|1|0|9.0||7159831|center|0|0"
    AddST "C|H24|e||General|1|0|9.0||7159831|center|0|0"
    AddST "C|A25|s|Overtime|General|0|0|9.0|||center|0|1"
    AddST "C|C25|s|Sergeant|General|0|0|9.0||||0|1"
    AddST "C|D25|e||General|0|0|9.0|||center|0|1"
    AddST "C|E25|f|=IFERROR(SUM(D25*'Baseline Costs'!$L$31),0)|""~uA3~""#,##0.00|0|0|9.0|||center|0|1"
    AddST "C|F25|e||General|0|0|9.0|||center|0|1"
    AddST "C|G25|f|=IFERROR(SUM(F25*'Baseline Costs'!$X$31),0)|""~uA3~""#,##0.00|0|0|9.0|||center|0|1"
    AddST "C|H25|f|=E25+G25|""~uA3~""#,##0.00|0|0|9.0|||center|0|1"
    AddST "C|C26|s|Constable|General|0|0|9.0||||0|1"
    AddST "C|D26|e||General|0|0|9.0|||center|0|1"
    AddST "C|E26|f|=IFERROR(SUM(D26*'Baseline Costs'!$L$32),0)|""~uA3~""#,##0.00|0|0|9.0|||center|0|1"
    AddST "C|F26|e||General|0|0|9.0|||center|0|1"
    AddST "C|G26|f|=IFERROR(SUM(F26*'Baseline Costs'!$X$32),0)|""~uA3~""#,##0.00|0|0|9.0|||center|0|1"
    AddST "C|H26|f|=E26+G26|""~uA3~""#,##0.00|0|0|9.0|||center|0|1"
    AddST "W|A|22.5"
    AddST "W|B|28.5"
    AddST "W|C|20.0"
    AddST "W|D|25.83203125"
    AddST "W|I|9.33203125"
    AddST "W|EE|9.33203125"
    AddST "M|A18:C18"
    AddST "M|F3:F4"
    AddST "M|A21:C21"
    AddST "M|A12:C12"
End Sub

Private Sub DataST2_2()
    Dim s As String
    AddST "M|A22:B23"
    AddST "M|A13:B14"
    AddST "M|G3:G4"
    AddST "M|A19:B20"
    AddST "M|A10:B11"
    AddST "M|D3:D4"
    AddST "M|H3:H4"
    AddST "M|A9:C9"
    AddST "M|A15:C15"
    AddST "M|A24:C24"
    AddST "M|A16:B17"
    AddST "M|A7:B8"
    AddST "M|C3:C4"
    AddST "M|A6:C6"
    AddST "M|E3:E4"
    AddST "M|A2:G2"
    AddST "M|A25:B26"
    AddST "M|A3:B4"
End Sub

Private Sub DataST3_1()
    Dim s As String
    AddST "C|B1|s|MUTUAL AID RATES - DO NOT EDIT|General|1|0|9.0|192|15917529||0|0"
    AddST "C|B2|s|2025-2026|General|1|0|6.0|0|39423|left|0|1"
    AddST "C|C2|s|Constable|General|0|0|6.0||39423||0|1"
    AddST "C|D2|s|Sergeant|General|0|0|6.0||39423||0|1"
    AddST "C|E2|s|Inspector|General|0|0|6.0||39423||0|1"
    AddST "C|F2|s|Chief ~n~Inspector|General|0|0|6.0||39423||0|1"
    AddST "C|G2|s|Supt|General|0|0|6.0||39423||0|1"
    AddST "C|H2|s|Chief Supt|General|0|0|6.0||39423||0|1"
    AddST "C|P2|s|2025-2026|General|1|0|6.0||39423|left|0|1"
    AddST "C|Q2|s|PSU|General|0|0|6.0||39423|center|0|1"
    AddST "C|B3|s|Annual Employable Costs|General|0|0|6.0||39423||0|1"
    AddST "C|C3|n|50053|[$~uA3~-809]#,##0|0|0|6.0||39423||0|0"
    AddST "C|D3|n|63804|[$~uA3~-809]#,##0|0|0|6.0||39423||0|0"
    AddST "C|E3|n|76652|[$~uA3~-809]#,##0|0|0|6.0||39423||0|0"
    AddST "C|F3|n|83959|[$~uA3~-809]#,##0|0|0|6.0||39423||0|0"
    AddST "C|G3|n|103452|[$~uA3~-809]#,##0|0|0|6.0||39423||0|0"
    AddST "C|H3|n|119468|[$~uA3~-809]#,##0|0|0|6.0||39423||0|0"
    AddST "C|J3|s|2025/2026|General|0|0|6.0|0|39423||0|1"
    AddST "C|K3|n|1|0.000000000|0|0|6.0||39423||0|1"
    AddST "C|M3|s|2025-2026|General|0|0|6.0|0|39423||0|1"
    AddST "C|N3|n|0.042|0.00%|0|0|6.0||39423||0|1"
    AddST "C|P3|s|IDR SUB|General|0|0|6.0||39423||0|1"
    AddST "C|Q3|n|31601.07|[$~uA3~-809]#,##0|0|0|6.0||39423||0|0"
    AddST "C|B4|s|Monthly Employable Costs|General|0|0|6.0||39423||0|1"
    AddST "C|C4|n|4171.083333333333|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|D4|n|5317|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|E4|n|6387.666666666667|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|F4|n|6996.583333333333|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|G4|n|8621|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|H4|n|9955.666666666666|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|J4|s|2026/2027|General|0|0|6.0|0|39423||0|1"
    AddST "C|K4|f|=K3*(1+('GDP Deflators'!D79/100))|0.000000000|0|0|6.0||39423||0|0"
    AddST "C|M4|s|2026-2027|General|0|0|6.0|0|39423||0|1"
    AddST "C|N4|n|0.04|0.00%|0|0|6.0||39423||0|0"
    AddST "C|P4|s|IDR/TDR SUB|General|0|0|6.0||39423||0|1"
    AddST "C|Q4|n|31601.07|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|B5|s|Hourly Employable Costs|General|0|0|6.0||39423||0|1"
    AddST "C|C5|n|49.42|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|D5|n|61.49|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|E5|n|61.43|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|F5|n|68.27|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|G5|n|84.2|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|H5|n|97.21|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|J5|s|2027/2028|General|0|0|6.0|0|39423||0|1"
    AddST "C|K5|f|=K3*(1+('GDP Deflators'!D80/100))|0.000000000|0|0|6.0||39423||0|0"
    AddST "C|M5|s|2027-2028|General|0|0|6.0|0|39423||0|1"
    AddST "C|N5|n|0.035|0.00%|0|0|6.0||39423||0|0"
    AddST "C|P5|s|ADMIN|General|0|0|6.0||39423||0|1"
    AddST "C|Q5|f|=Q4*0.05|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|B6|s|Hourly Overtime Costs|General|0|0|6.0||39423||0|1"
    AddST "C|C6|n|49.42|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|D6|n|61.49|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|E6|n|61.43|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|F6|n|68.27|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|G6|n|84.2|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|H6|n|97.21|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|J6|s|2028/2029|General|0|0|6.0|0|39423||0|1"
    AddST "C|K6|f|=K3*(1+('GDP Deflators'!D81/100))|0.000000000|0|0|6.0||39423||0|0"
    AddST "C|M6|s|2028-2029|General|0|0|6.0|0|39423||0|1"
    AddST "C|N6|n|0.03|0.00%|0|0|6.0||39423||0|0"
    AddST "C|P6|s|TOTAL|General|0|0|6.0||39423||0|1"
    AddST "C|Q6|f|=Q4+Q5|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|B9|e||General|1|0|6.0||39423|left|0|1"
    AddST "C|C9|d|2026-09-01|mmm-yy|1|0|6.0||39423|left|0|1"
    AddST "C|D9|d|2026-10-01|mmm-yy|1|0|6.0||39423|left|0|1"
    AddST "C|E9|d|2026-11-01|mmm-yy|1|0|6.0||39423|left|0|1"
    AddST "C|F9|d|2026-12-01|mmm-yy|1|0|6.0||39423|left|0|1"
    AddST "C|G9|d|2027-01-01|mmm-yy|1|0|6.0||39423|left|0|1"
    AddST "C|H9|d|2027-02-01|mmm-yy|1|0|6.0||39423|left|0|1"
    AddST "C|I9|d|2027-03-01|mmm-yy|1|0|6.0||39423|left|0|1"
    AddST "C|J9|d|2027-04-01|mmm-yy|1|0|6.0||39423|left|0|1"
    AddST "C|K9|d|2027-05-01|mmm-yy|1|0|6.0||39423|left|0|1"
    AddST "C|L9|d|2027-06-01|mmm-yy|1|0|6.0||39423|left|0|1"
    AddST "C|M9|d|2027-07-01|mmm-yy|1|0|6.0||39423|left|0|1"
    AddST "C|N9|d|2027-08-01|mmm-yy|1|0|6.0||39423|left|0|1"
    AddST "C|O9|d|2027-09-01|mmm-yy|1|0|6.0||39423|left|0|1"
    AddST "C|P9|d|2027-10-01|mmm-yy|1|0|6.0||39423|left|0|1"
    AddST "C|Q9|d|2027-11-01|mmm-yy|1|0|6.0||39423|left|0|1"
    AddST "C|R9|d|2027-12-01|mmm-yy|1|0|6.0||39423|left|0|1"
    AddST "C|S9|d|2028-01-01|mmm-yy|1|0|6.0||39423|left|0|1"
    AddST "C|T9|d|2028-02-01|mmm-yy|1|0|6.0||39423|left|0|1"
    AddST "C|U9|d|2028-03-01|mmm-yy|1|0|6.0||39423|left|0|1"
    AddST "C|V9|d|2028-04-01|mmm-yy|1|0|6.0||39423|left|0|1"
    AddST "C|W9|d|2028-05-01|mmm-yy|1|0|6.0||39423|left|0|1"
    AddST "C|X9|d|2028-06-01|mmm-yy|1|0|6.0||39423|left|0|1"
    AddST "C|Y9|d|2028-07-01|mmm-yy|1|0|6.0||39423|left|0|1"
    AddST "C|Z9|d|2028-08-01|mmm-yy|1|0|6.0||39423|left|0|1"
    AddST "C|AA9|d|2028-09-01|mmm-yy|1|0|6.0||39423|left|0|1"
    AddST "C|AB9|d|2028-10-01|mmm-yy|1|0|6.0||39423|left|0|1"
    AddST "C|AC9|d|2028-11-01|mmm-yy|1|0|6.0||39423|left|0|1"
    AddST "C|AD9|d|2028-12-01|mmm-yy|1|0|6.0||39423|left|0|1"
    AddST "C|B10|s|Chief Superintendent|General|0|0|6.0||39423||0|1"
    AddST "C|C10|f|=(H6*$K$4)+(H6*$N$4)|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|1"
    AddST "C|D10|f|=C10|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|E10|f|=D10|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|F10|f|=E10|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|G10|f|=F10|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|H10|f|=G10|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|I10|f|=H10|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|J10|f|=I10|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|1"
    AddST "C|K10|f|=J10|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|L10|f|=K10|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|M10|f|=L10|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|N10|f|=M10|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|O10|f|=N10+(N10*$N$5)|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|1"
    AddST "C|P10|f|=O10|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|Q10|f|=P10|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|R10|f|=Q10|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|S10|f|=R10|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|T10|f|=S10|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|U10|f|=T10|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|V10|f|=U10|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|1"
    AddST "C|W10|f|=V10|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|X10|f|=W10|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|Y10|f|=X10|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|Z10|f|=Y10|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|AA10|f|=Z10+(Z10*$N$6)|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|1"
    AddST "C|AB10|f|=AA10|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|AC10|f|=AB10|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|AD10|f|=AC10|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|B11|s|Superintendent|General|0|0|6.0||39423||0|1"
    AddST "C|C11|f|=(G6*K4)+(G6*N4)|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|1"
    AddST "C|D11|f|=C11|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|E11|f|=D11|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|F11|f|=E11|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|G11|f|=F11|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|H11|f|=G11|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|I11|f|=H11|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|J11|f|=I11|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|1"
    AddST "C|K11|f|=J11|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|L11|f|=K11|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|M11|f|=L11|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|N11|f|=M11|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|O11|f|=N11+(N11*$N$5)|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|1"
    AddST "C|P11|f|=O11|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|Q11|f|=P11|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|R11|f|=Q11|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|S11|f|=R11|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|T11|f|=S11|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|U11|f|=T11|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|V11|f|=U11|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|1"
    AddST "C|W11|f|=V11|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|X11|f|=W11|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|Y11|f|=X11|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|Z11|f|=Y11|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|AA11|f|=Z11+(Z11*$N$6)|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|1"
    AddST "C|AB11|f|=AA11|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|AC11|f|=AB11|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|AD11|f|=AC11|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|B12|s|Chief Inspector|General|0|0|6.0||39423||0|1"
End Sub

Private Sub DataST3_2()
    Dim s As String
    AddST "C|C12|f|=(F6*K4)+(F6*N4)|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|1"
    AddST "C|D12|f|=C12|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|E12|f|=D12|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|F12|f|=E12|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|G12|f|=F12|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|H12|f|=G12|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|I12|f|=H12|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|J12|f|=I12|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|1"
    AddST "C|K12|f|=J12|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|L12|f|=K12|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|M12|f|=L12|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|N12|f|=M12|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|O12|f|=N12+(N12*$N$5)|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|1"
    AddST "C|P12|f|=O12|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|Q12|f|=P12|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|R12|f|=Q12|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|S12|f|=R12|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|T12|f|=S12|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|U12|f|=T12|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|V12|f|=U12|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|1"
    AddST "C|W12|f|=V12|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|X12|f|=W12|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|Y12|f|=X12|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|Z12|f|=Y12|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|AA12|f|=Z12+(Z12*$N$6)|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|1"
    AddST "C|AB12|f|=AA12|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|AC12|f|=AB12|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|AD12|f|=AC12|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|B13|s|Inspector|General|0|0|6.0||39423||0|1"
    AddST "C|C13|f|=(E6*K4)+(E6*N4)|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|1"
    AddST "C|D13|f|=C13|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|E13|f|=D13|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|F13|f|=E13|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|G13|f|=F13|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|H13|f|=G13|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|I13|f|=H13|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|J13|f|=I13|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|1"
    AddST "C|K13|f|=J13|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|L13|f|=K13|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|M13|f|=L13|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|N13|f|=M13|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|O13|f|=N13+(N13*$N$5)|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|1"
    AddST "C|P13|f|=O13|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|Q13|f|=P13|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|R13|f|=Q13|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|S13|f|=R13|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|T13|f|=S13|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|U13|f|=T13|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|V13|f|=U13|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|1"
    AddST "C|W13|f|=V13|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|X13|f|=W13|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|Y13|f|=X13|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|Z13|f|=Y13|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|AA13|f|=Z13+(Z13*$N$6)|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|1"
    AddST "C|AB13|f|=AA13|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|AC13|f|=AB13|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|AD13|f|=AC13|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|B14|s|Sergeant|General|0|0|6.0||39423||0|1"
    AddST "C|C14|f|=(D6*K4)+(D6*N4)|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|1"
    AddST "C|D14|f|=C14|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|E14|f|=D14|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|F14|f|=E14|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|G14|f|=F14|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|H14|f|=G14|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|I14|f|=H14|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|J14|f|=I14|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|1"
    AddST "C|K14|f|=J14|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|L14|f|=K14|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|M14|f|=L14|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|N14|f|=M14|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|O14|f|=N14+(N14*$N$5)|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|1"
    AddST "C|P14|f|=O14|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|Q14|f|=P14|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|R14|f|=Q14|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|S14|f|=R14|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|T14|f|=S14|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|U14|f|=T14|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|V14|f|=U14|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|1"
    AddST "C|W14|f|=V14|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|X14|f|=W14|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|Y14|f|=X14|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|Z14|f|=Y14|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|AA14|f|=Z14+(Z14*$N$6)|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|1"
    AddST "C|AB14|f|=AA14|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|AC14|f|=AB14|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|AD14|f|=AC14|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|B15|s|Constable|General|0|0|6.0||39423||0|1"
    AddST "C|C15|f|=(C6*K4)+(C6*N4)|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|1"
    AddST "C|D15|f|=C15|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|E15|f|=D15|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|F15|f|=E15|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|G15|f|=F15|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|H15|f|=G15|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|I15|f|=H15|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|J15|f|=I15|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|1"
    AddST "C|K15|f|=J15|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|L15|f|=K15|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|M15|f|=L15|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|N15|f|=M15|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|O15|f|=N15+(N15*$N$5)|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|1"
    AddST "C|P15|f|=O15|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|Q15|f|=P15|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|R15|f|=Q15|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|S15|f|=R15|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|T15|f|=S15|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|U15|f|=T15|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|V15|f|=U15|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|1"
    AddST "C|W15|f|=V15|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|X15|f|=W15|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|Y15|f|=X15|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|Z15|f|=Y15|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|AA15|f|=Z15+(Z15*$N$6)|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|1"
    AddST "C|AB15|f|=AA15|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|AC15|f|=AB15|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|AD15|f|=AC15|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|B17|s|MA PSU|General|0|0|6.0||39423||0|1"
    AddST "C|C17|f|=Q6+(Q6*N4)|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|1"
    AddST "C|D17|f|=C17|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|1"
    AddST "C|E17|f|=D17|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|1"
    AddST "C|F17|f|=E17|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|1"
    AddST "C|G17|f|=F17|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|1"
    AddST "C|H17|f|=G17|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|1"
    AddST "C|I17|f|=H17|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|1"
    AddST "C|J17|f|=I17|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|1"
    AddST "C|K17|f|=J17|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|1"
    AddST "C|L17|f|=K17|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|1"
    AddST "C|M17|f|=L17|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|1"
    AddST "C|N17|f|=M17|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|1"
    AddST "C|O17|f|=N17+(N17*$N$5)|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|1"
    AddST "C|P17|f|=O17|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|1"
    AddST "C|Q17|f|=P17|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|1"
    AddST "C|R17|f|=Q17|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|1"
    AddST "C|S17|f|=R17|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|1"
    AddST "C|T17|f|=S17|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|1"
    AddST "C|U17|f|=T17|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|1"
    AddST "C|V17|f|=U17|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|1"
    AddST "C|W17|f|=V17|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|1"
    AddST "C|X17|f|=W17|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|1"
    AddST "C|Y17|f|=X17|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|1"
    AddST "C|Z17|f|=Y17|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|1"
    AddST "C|AA17|f|=Z17+(Z17*$N$6)|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|1"
    AddST "C|AB17|f|=AA17|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|1"
    AddST "C|AC17|f|=AB17|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|1"
    AddST "C|AD17|f|=AC17|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|1"
    AddST "C|B19|s|2025-2026|General|1|0|6.0|0|39423|left|0|1"
    AddST "C|C19|s|Constable|General|0|0|6.0||39423||0|1"
    AddST "C|D19|s|Sergeant|General|0|0|6.0||39423||0|1"
    AddST "C|E19|s|Inspector|General|0|0|6.0||39423||0|1"
    AddST "C|F19|s|Chief ~n~Inspector|General|0|0|6.0||39423||0|1"
    AddST "C|G19|s|Supt|General|0|0|6.0||39423||0|1"
End Sub

Private Sub DataST3_3()
    Dim s As String
    AddST "C|H19|s|Chief Supt|General|0|0|6.0||39423||0|1"
    AddST "C|P19|s|2025-2026|General|1|0|6.0||39423|left|0|1"
    AddST "C|Q19|s|PSU|General|0|0|6.0||39423|center|0|1"
    AddST "C|B20|s|Annual Employable Costs|General|0|0|6.0||39423||0|1"
    AddST "C|C20|n|70300|[$~uA3~-809]#,##0|0|0|6.0||39423||0|0"
    AddST "C|D20|n|83200|[$~uA3~-809]#,##0|0|0|6.0||39423||0|0"
    AddST "C|E20|n|101900|[$~uA3~-809]#,##0|0|0|6.0||39423||0|0"
    AddST "C|F20|n|109800|[$~uA3~-809]#,##0|0|0|6.0||39423||0|0"
    AddST "C|G20|n|143600|[$~uA3~-809]#,##0|0|0|6.0||39423||0|0"
    AddST "C|H20|n|172900|[$~uA3~-809]#,##0|0|0|6.0||39423||0|0"
    AddST "C|J20|s|2025/2026|General|0|0|6.0|0|39423||0|1"
    AddST "C|K20|n|1|0.000000000|0|0|6.0||39423||0|1"
    AddST "C|M20|s|2025-2026|General|0|0|6.0|0|39423||0|1"
    AddST "C|N20|n|0.042|0.00%|0|0|6.0||39423||0|1"
    AddST "C|P20|s|IDR SUB|General|0|0|6.0||39423||0|1"
    AddST "C|Q20|n|31601.07|[$~uA3~-809]#,##0|0|0|6.0||39423||0|0"
    AddST "C|B21|s|Monthly Employable Costs|General|0|0|6.0||39423||0|1"
    AddST "C|C21|n|5858|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|D21|n|6933|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|E21|n|8492|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|F21|n|9150|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|G21|n|11967|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|H21|n|14408|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|J21|s|2026/2027|General|0|0|6.0|0|39423||0|1"
    AddST "C|K21|f|=K20*(1+('GDP Deflators'!D79/100))|0.000000000|0|0|6.0||39423||0|0"
    AddST "C|M21|s|2026-2027|General|0|0|6.0|0|39423||0|1"
    AddST "C|N21|n|0.04|0.00%|0|0|6.0||39423||0|0"
    AddST "C|P21|s|IDR/TDR SUB|General|0|0|6.0||39423||0|1"
    AddST "C|Q21|n|31601.07|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|B22|s|Hourly Employable Costs|General|0|0|6.0||39423||0|1"
    AddST "C|C22|n|49.42|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|D22|n|61.49|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|E22|n|61.43|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|F22|n|68.27|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|G22|n|84.2|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|H22|n|97.21|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|J22|s|2027/2028|General|0|0|6.0|0|39423||0|1"
    AddST "C|K22|f|=K20*(1+('GDP Deflators'!D80/100))|0.000000000|0|0|6.0||39423||0|0"
    AddST "C|M22|s|2027-2028|General|0|0|6.0|0|39423||0|1"
    AddST "C|N22|n|0.035|0.00%|0|0|6.0||39423||0|0"
    AddST "C|P22|s|ADMIN|General|0|0|6.0||39423||0|1"
    AddST "C|Q22|f|=Q21*0.05|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|B23|s|Hourly Overtime Costs|General|0|0|6.0||39423||0|1"
    AddST "C|C23|n|44.81|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|D23|n|53.03|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|E23|n|48.83|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|F23|n|52.62|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|G23|n|68.82|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|H23|n|82.86|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|J23|s|2028/2029|General|0|0|6.0|0|39423||0|1"
    AddST "C|K23|f|=K20*(1+('GDP Deflators'!D81/100))|0.000000000|0|0|6.0||39423||0|0"
    AddST "C|M23|s|2028-2029|General|0|0|6.0|0|39423||0|1"
    AddST "C|N23|n|0.03|0.00%|0|0|6.0||39423||0|0"
    AddST "C|P23|s|TOTAL|General|0|0|6.0||39423||0|1"
    AddST "C|Q23|f|=Q21+Q22|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|B25|s|HOST FORCE OVERTIME RATE - TO BE UPDATED|General|1|0|9.0|192|15917529||0|0"
    AddST "C|B26|e||General|1|0|6.0||39423|left|0|1"
    AddST "C|C26|d|2026-09-01|mmm-yy|1|0|6.0||39423|left|0|1"
    AddST "C|D26|d|2026-10-01|mmm-yy|1|0|6.0||39423|left|0|1"
    AddST "C|E26|d|2026-11-01|mmm-yy|1|0|6.0||39423|left|0|1"
    AddST "C|F26|d|2026-12-01|mmm-yy|1|0|6.0||39423|left|0|1"
    AddST "C|G26|d|2027-01-01|mmm-yy|1|0|6.0||39423|left|0|1"
    AddST "C|H26|d|2027-02-01|mmm-yy|1|0|6.0||39423|left|0|1"
    AddST "C|I26|d|2027-03-01|mmm-yy|1|0|6.0||39423|left|0|1"
    AddST "C|J26|d|2027-04-01|mmm-yy|1|0|6.0||39423|left|0|1"
    AddST "C|K26|d|2027-05-01|mmm-yy|1|0|6.0||39423|left|0|1"
    AddST "C|L26|d|2027-06-01|mmm-yy|1|0|6.0||39423|left|0|1"
    AddST "C|M26|d|2027-07-01|mmm-yy|1|0|6.0||39423|left|0|1"
    AddST "C|N26|d|2027-08-01|mmm-yy|1|0|6.0||39423|left|0|1"
    AddST "C|O26|d|2027-09-01|mmm-yy|1|0|6.0||39423|left|0|1"
    AddST "C|P26|d|2027-10-01|mmm-yy|1|0|6.0||39423|left|0|1"
    AddST "C|Q26|d|2027-11-01|mmm-yy|1|0|6.0||39423|left|0|1"
    AddST "C|R26|d|2027-12-01|mmm-yy|1|0|6.0||39423|left|0|1"
    AddST "C|S26|d|2028-01-01|mmm-yy|1|0|6.0||39423|left|0|1"
    AddST "C|T26|d|2028-02-01|mmm-yy|1|0|6.0||39423|left|0|1"
    AddST "C|U26|d|2028-03-01|mmm-yy|1|0|6.0||39423|left|0|1"
    AddST "C|V26|d|2028-04-01|mmm-yy|1|0|6.0||39423|left|0|1"
    AddST "C|W26|d|2028-05-01|mmm-yy|1|0|6.0||39423|left|0|1"
    AddST "C|X26|d|2028-06-01|mmm-yy|1|0|6.0||39423|left|0|1"
    AddST "C|Y26|d|2028-07-01|mmm-yy|1|0|6.0||39423|left|0|1"
    AddST "C|Z26|d|2028-08-01|mmm-yy|1|0|6.0||39423|left|0|1"
    AddST "C|AA26|d|2028-09-01|mmm-yy|1|0|6.0||39423|left|0|1"
    AddST "C|AB26|d|2028-10-01|mmm-yy|1|0|6.0||39423|left|0|1"
    AddST "C|AC26|d|2028-11-01|mmm-yy|1|0|6.0||39423|left|0|1"
    AddST "C|AD26|d|2028-12-01|mmm-yy|1|0|6.0||39423|left|0|1"
    AddST "C|B27|s|Chief Superintendent|General|0|0|6.0||39423||0|1"
    AddST "C|C27|f|=(H23)+(H23*$N$4)|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|1"
    AddST "C|D27|f|=C27|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|E27|f|=D27|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|F27|f|=E27|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|G27|f|=F27|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|H27|f|=G27|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|I27|f|=H27|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|J27|f|=I27|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|1"
    AddST "C|K27|f|=J27|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|L27|f|=K27|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|M27|f|=L27|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|N27|f|=M27|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|O27|f|=N27+(N27*$N$5)|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|1"
    AddST "C|P27|f|=O27|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|Q27|f|=P27|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|R27|f|=Q27|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|S27|f|=R27|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|T27|f|=S27|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|U27|f|=T27|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|V27|f|=U27|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|1"
    AddST "C|W27|f|=V27|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|X27|f|=W27|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|Y27|f|=X27|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|Z27|f|=Y27|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|AA27|f|=Z27+(Z27*$N$6)|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|1"
    AddST "C|AB27|f|=AA27|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|AC27|f|=AB27|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|AD27|f|=AC27|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|B28|s|Superintendent|General|0|0|6.0||39423||0|1"
    AddST "C|C28|f|=(G23)+(G23*N21)|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|1"
    AddST "C|D28|f|=C28|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|E28|f|=D28|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|F28|f|=E28|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|G28|f|=F28|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|H28|f|=G28|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|I28|f|=H28|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|J28|f|=I28|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|1"
    AddST "C|K28|f|=J28|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|L28|f|=K28|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|M28|f|=L28|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|N28|f|=M28|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|O28|f|=N28+(N28*$N$5)|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|1"
    AddST "C|P28|f|=O28|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|Q28|f|=P28|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|R28|f|=Q28|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|S28|f|=R28|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|T28|f|=S28|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|U28|f|=T28|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|V28|f|=U28|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|1"
    AddST "C|W28|f|=V28|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|X28|f|=W28|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|Y28|f|=X28|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|Z28|f|=Y28|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|AA28|f|=Z28+(Z28*$N$6)|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|1"
    AddST "C|AB28|f|=AA28|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|AC28|f|=AB28|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|AD28|f|=AC28|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|B29|s|Chief Inspector|General|0|0|6.0||39423||0|1"
    AddST "C|C29|f|=(F23)+(F23*N21)|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|1"
    AddST "C|D29|f|=C29|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|E29|f|=D29|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|F29|f|=E29|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|G29|f|=F29|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|H29|f|=G29|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
End Sub

Private Sub DataST3_4()
    Dim s As String
    AddST "C|I29|f|=H29|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|J29|f|=I29|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|1"
    AddST "C|K29|f|=J29|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|L29|f|=K29|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|M29|f|=L29|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|N29|f|=M29|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|O29|f|=N29+(N29*$N$5)|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|1"
    AddST "C|P29|f|=O29|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|Q29|f|=P29|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|R29|f|=Q29|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|S29|f|=R29|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|T29|f|=S29|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|U29|f|=T29|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|V29|f|=U29|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|1"
    AddST "C|W29|f|=V29|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|X29|f|=W29|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|Y29|f|=X29|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|Z29|f|=Y29|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|AA29|f|=Z29+(Z29*$N$6)|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|1"
    AddST "C|AB29|f|=AA29|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|AC29|f|=AB29|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|AD29|f|=AC29|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|B30|s|Inspector|General|0|0|6.0||39423||0|1"
    AddST "C|C30|f|=(E23)+(E23*N21)|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|1"
    AddST "C|D30|f|=C30|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|E30|f|=D30|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|F30|f|=E30|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|G30|f|=F30|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|H30|f|=G30|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|I30|f|=H30|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|J30|f|=I30|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|1"
    AddST "C|K30|f|=J30|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|L30|f|=K30|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|M30|f|=L30|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|N30|f|=M30|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|O30|f|=N30+(N30*$N$5)|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|1"
    AddST "C|P30|f|=O30|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|Q30|f|=P30|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|R30|f|=Q30|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|S30|f|=R30|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|T30|f|=S30|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|U30|f|=T30|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|V30|f|=U30|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|1"
    AddST "C|W30|f|=V30|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|X30|f|=W30|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|Y30|f|=X30|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|Z30|f|=Y30|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|AA30|f|=Z30+(Z30*$N$6)|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|1"
    AddST "C|AB30|f|=AA30|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|AC30|f|=AB30|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|AD30|f|=AC30|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|B31|s|Sergeant|General|0|0|6.0||39423||0|1"
    AddST "C|C31|f|=(D23)+(D23*N21)|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|1"
    AddST "C|D31|f|=C31|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|E31|f|=D31|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|F31|f|=E31|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|G31|f|=F31|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|H31|f|=G31|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|I31|f|=H31|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|J31|f|=I31|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|1"
    AddST "C|K31|f|=J31|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|L31|f|=K31|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|M31|f|=L31|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|N31|f|=M31|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|O31|f|=N31+(N31*$N$5)|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|1"
    AddST "C|P31|f|=O31|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|Q31|f|=P31|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|R31|f|=Q31|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|S31|f|=R31|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|T31|f|=S31|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|U31|f|=T31|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|V31|f|=U31|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|1"
    AddST "C|W31|f|=V31|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|X31|f|=W31|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|Y31|f|=X31|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|Z31|f|=Y31|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|AA31|f|=Z31+(Z31*$N$6)|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|1"
    AddST "C|AB31|f|=AA31|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|AC31|f|=AB31|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|AD31|f|=AC31|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|B32|s|Constable|General|0|0|6.0||39423||0|1"
    AddST "C|C32|f|=(C23)+(C23*N21)|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|1"
    AddST "C|D32|f|=C32|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|E32|f|=D32|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|F32|f|=E32|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|G32|f|=F32|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|H32|f|=G32|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|I32|f|=H32|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|J32|f|=I32|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|1"
    AddST "C|K32|f|=J32|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|L32|f|=K32|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|M32|f|=L32|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|N32|f|=M32|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|O32|f|=N32+(N32*$N$5)|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|1"
    AddST "C|P32|f|=O32|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|Q32|f|=P32|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|R32|f|=Q32|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|S32|f|=R32|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|T32|f|=S32|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|U32|f|=T32|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|V32|f|=U32|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|1"
    AddST "C|W32|f|=V32|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|X32|f|=W32|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|Y32|f|=X32|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|Z32|f|=Y32|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|AA32|f|=Z32+(Z32*$N$6)|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|1"
    AddST "C|AB32|f|=AA32|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|AC32|f|=AB32|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "C|AD32|f|=AC32|[$~uA3~-809]#,##0.00|0|0|6.0||39423||0|0"
    AddST "W|B|18.83203125"
    AddST "W|J|9.6640625"
    AddST "W|K|12.6640625"
    AddST "W|AA|8.83203125"
End Sub

Private Sub DataST4_1()
    Dim s As String
    AddST "C|B1|s|GDP DEFLATORS AT MARKET PRICES, AND MONEY GDP|General|1|0|10.0|||left|0|0"
    AddST "C|B3|s|Outturn data are as at the Quarterly National Accounts from ONS - last updated 31 March 2026.|General|0|0|10.0|||left|1|0"
    AddST "C|B4|s|Forecast data are consistent with OBR EFO data as at the Spring Statement 3 March 2026.|General|0|0|10.0|||left|0|0"
    AddST "C|B5|s|Financial year|General|1|0|10.0|||center|0|1"
    AddST "C|H5|s|Calendar year|General|1|0|10.0|||center|1|1"
    AddST "C|B6|e||General|0|0|10.0|||left|0|1"
    AddST "C|C6|s|GDP deflator at market prices |General|0|0|10.0|||center|0|0"
    AddST "C|E6|s|Money GDP (3), (4)|#,##0|0|0|10.0|||centerContinuous|0|0"
    AddST "C|F6|e||#,##0|0|0|10.0|||centerContinuous|0|1"
    AddST "C|H6|e||General|0|0|10.0|||left|0|1"
    AddST "C|I6|s|GDP deflator at market prices |General|0|0|10.0|||center|0|1"
    AddST "C|K6|s|Money GDP (3)|#,##0|0|0|10.0|||centerContinuous|0|1"
    AddST "C|B7|s|Financial year|General|0|0|10.0|||left|1|1"
    AddST "C|C7|s|2024-25 = 100|0.000|0|0|10.0||||1|0"
    AddST "C|D7|s|per cent change on previous year|0.00|0|0|10.0|||left|1|0"
    AddST "C|E7|s|Cash ~uA3~ million~n~Non-Season-ally Adjusted|#,##0|0|0|10.0|||left|1|1"
    AddST "C|F7|s|Cash ~uA3~ million~n~Seasonally Adjusted|#,##0|0|0|10.0|||left|1|1"
    AddST "C|H7|s|Calendar year|General|0|0|10.0|||left|1|1"
    AddST "C|I7|s|2025 = 100|0.000|0|0|10.0|||left|1|0"
    AddST "C|J7|s|per cent change on previous year|0.00|0|0|10.0|||left|1|0"
    AddST "C|K7|s|Cash ~uA3~ million~n~Non-Seasonally Adjusted|#,##0|0|0|10.0|||left|1|1"
    AddST "C|B8|s|1955-56|General|0|0|10.0|||left|1|1"
    AddST "C|C8|n|3.1645|0.000|0|0|10.0|||right|1|0"
    AddST "C|E8|n|19575|#,##0|0|0|10.0|||right|1|0"
    AddST "C|F8|n|19590|#,##0|0|0|10.0|||right|1|1"
    AddST "C|H8|s|1955|General|0|0|10.0|||left|0|1"
    AddST "C|I8|n|3.0344|0.000|0|0|10.0|||right|1|0"
    AddST "C|K8|n|19170|#,##0|0|0|10.0|||right|1|0"
    AddST "C|B9|s|1956-57|General|0|0|10.0|||left|1|1"
    AddST "C|C9|n|3.3593|0.000|0|0|10.0|||right|1|0"
    AddST "C|D9|n|6.155790804234486|0.00|0|0|10.0|||right|1|0"
    AddST "C|E9|n|21164|#,##0|0|0|10.0|||right|1|0"
    AddST "C|F9|n|21140|#,##0|0|0|10.0|||right|1|1"
    AddST "C|H9|s|1956|General|0|0|10.0|||left|0|1"
    AddST "C|I9|n|3.2449|0.000|0|0|10.0|||right|1|0"
    AddST "C|J9|n|6.937121012391237|0.00|0|0|10.0|||right|1|0"
    AddST "C|K9|n|20846|#,##0|0|0|10.0|||right|1|0"
    AddST "C|B10|s|1957-58|General|0|0|10.0|||left|1|1"
    AddST "C|C10|n|3.5142|0.000|0|0|10.0|||right|1|0"
    AddST "C|D10|n|4.611079689220969|0.00|0|0|10.0|||right|1|0"
    AddST "C|E10|n|22512|#,##0|0|0|10.0|||right|1|0"
    AddST "C|F10|n|22527|#,##0|0|0|10.0|||right|1|1"
    AddST "C|H10|s|1957|General|0|0|10.0|||left|0|1"
    AddST "C|I10|n|3.3749|0.000|0|0|10.0|||right|1|0"
    AddST "C|J10|n|4.006286788498873|0.00|0|0|10.0|||right|1|0"
    AddST "C|K10|n|22107|#,##0|0|0|10.0|||right|1|0"
    AddST "C|B11|s|1958-59|General|0|0|10.0|||left|1|1"
    AddST "C|C11|n|3.5995|0.000|0|0|10.0|||right|1|0"
    AddST "C|D11|n|2.4272949746741705|0.00|0|0|10.0|||right|1|0"
    AddST "C|E11|n|23339|#,##0|0|0|10.0|||right|1|0"
    AddST "C|F11|n|23309|#,##0|0|0|10.0|||right|1|1"
    AddST "C|H11|s|1958|General|0|0|10.0|||left|0|1"
    AddST "C|I11|n|3.4983|0.000|0|0|10.0|||right|1|0"
    AddST "C|J11|n|3.6564046342113894|0.00|0|0|10.0|||right|1|0"
    AddST "C|K11|n|23223|#,##0|0|0|10.0|||right|1|0"
    AddST "C|B12|s|1959-60|General|0|0|10.0|||left|1|1"
    AddST "C|C12|n|3.6187|0.000|0|0|10.0|||right|1|0"
    AddST "C|D12|n|0.5334074176969054|0.00|0|0|10.0|||right|1|0"
    AddST "C|E12|n|24878|#,##0|0|0|10.0|||right|1|0"
    AddST "C|F12|n|24933|#,##0|0|0|10.0|||right|1|1"
    AddST "C|H12|s|1959|General|0|0|10.0|||left|0|1"
    AddST "C|I12|n|3.5236|0.000|0|0|10.0|||right|1|0"
    AddST "C|J12|n|0.7232084155161107|0.00|0|0|10.0|||right|1|0"
    AddST "C|K12|n|24392|#,##0|0|0|10.0|||right|1|0"
    AddST "C|B13|s|1960-61|General|0|0|10.0|||left|1|1"
    AddST "C|C13|n|3.6909|0.000|0|0|10.0|||right|1|0"
    AddST "C|D13|n|1.9951916434078547|0.00|0|0|10.0|||right|1|0"
    AddST "C|E13|n|26644|#,##0|0|0|10.0|||right|1|0"
    AddST "C|F13|n|26662|#,##0|0|0|10.0|||right|1|1"
    AddST "C|H13|s|1960|General|0|0|10.0|||left|0|1"
    AddST "C|I13|n|3.5605|0.000|0|0|10.0|||right|1|0"
    AddST "C|J13|n|1.047224429560681|0.00|0|0|10.0|||right|1|0"
    AddST "C|K13|n|26206|#,##0|0|0|10.0|||right|1|0"
    AddST "C|B14|s|1961-62|General|0|0|10.0|||left|1|1"
    AddST "C|C14|n|3.8264|0.000|0|0|10.0|||right|1|0"
    AddST "C|D14|n|3.671191308352975|0.00|0|0|10.0|||right|1|0"
    AddST "C|E14|n|28147|#,##0|0|0|10.0|||right|1|0"
    AddST "C|F14|n|28175|#,##0|0|0|10.0|||right|1|1"
    AddST "C|H14|s|1961|General|0|0|10.0|||left|0|1"
    AddST "C|I14|n|3.6954|0.000|0|0|10.0|||right|1|0"
    AddST "C|J14|n|3.788793708748759|0.00|0|0|10.0|||right|1|0"
    AddST "C|K14|n|27924|#,##0|0|0|10.0|||right|1|0"
    AddST "C|B15|s|1962-63|General|0|0|10.0|||left|1|1"
    AddST "C|C15|n|3.9457|0.000|0|0|10.0|||right|1|0"
    AddST "C|D15|n|3.1178130880200703|0.00|0|0|10.0|||right|1|0"
    AddST "C|E15|n|29457|#,##0|0|0|10.0|||right|1|0"
    AddST "C|F15|n|29475|#,##0|0|0|10.0|||right|1|1"
    AddST "C|H15|s|1962|General|0|0|10.0|||left|0|1"
    AddST "C|I15|n|3.8284|0.000|0|0|10.0|||right|1|0"
    AddST "C|J15|n|3.5990691129512373|0.00|0|0|10.0|||right|1|0"
    AddST "C|K15|n|29242|#,##0|0|0|10.0|||right|1|0"
    AddST "C|B16|s|1963-64|General|0|0|10.0|||left|1|1"
    AddST "C|C16|n|4.0121|0.000|0|0|10.0|||right|1|0"
    AddST "C|D16|n|1.6828446156575572|0.00|0|0|10.0|||right|1|0"
    AddST "C|E16|n|31929|#,##0|0|0|10.0|||right|1|0"
    AddST "C|F16|n|31884|#,##0|0|0|10.0|||right|1|1"
    AddST "C|H16|s|1963|General|0|0|10.0|||left|0|1"
    AddST "C|I16|n|3.8842|0.000|0|0|10.0|||right|1|0"
    AddST "C|J16|n|1.4575279490126443|0.00|0|0|10.0|||right|1|0"
    AddST "C|K16|n|31107|#,##0|0|0|10.0|||right|1|0"
    AddST "C|B17|s|1964-65|General|0|0|10.0||||0|1"
    AddST "C|C17|n|4.201|0.000|0|0|10.0|||right|0|0"
    AddST "C|D17|n|4.708257520998963|0.00|0|0|10.0|||right|1|0"
    AddST "C|E17|n|34864|#,##0|0|0|10.0|||right|0|0"
    AddST "C|F17|n|34933|#,##0|0|0|10.0|||right|0|1"
    AddST "C|H17|s|1964|General|0|0|10.0|||left|0|1"
    AddST "C|I17|n|4.023|0.000|0|0|10.0|||right|0|0"
    AddST "C|J17|n|3.5734514185675255|0.00|0|0|10.0|||right|0|0"
    AddST "C|K17|n|34056|#,##0|0|0|10.0|||right|0|0"
    AddST "C|B18|s|1965-66|General|0|0|10.0||||0|1"
    AddST "C|C18|n|4.4272|0.000|0|0|10.0||||0|0"
    AddST "C|D18|n|5.384432278029051|0.00|0|0|10.0|||right|1|0"
    AddST "C|E18|n|37473|#,##0|0|0|10.0||||0|0"
    AddST "C|F18|n|37427|#,##0|0|0|10.0||||0|1"
    AddST "C|H18|s|1965|General|0|0|10.0|||left|0|1"
    AddST "C|I18|n|4.2621|0.000|0|0|10.0|||right|0|0"
    AddST "C|J18|n|5.943325876211796|0.00|0|0|10.0|||right|0|0"
    AddST "C|K18|n|36839|#,##0|0|0|10.0|||right|0|0"
    AddST "C|B19|s|1966-67|General|0|0|10.0||||0|1"
    AddST "C|C19|n|4.6525|0.000|0|0|10.0||||0|0"
    AddST "C|D19|n|5.088995301770867|0.00|0|0|10.0|||right|1|0"
    AddST "C|E19|n|39962|#,##0|0|0|10.0||||0|0"
    AddST "C|F19|n|39959|#,##0|0|0|10.0||||0|1"
    AddST "C|H19|s|1966|General|0|0|10.0|||left|0|1"
    AddST "C|I19|n|4.4896|0.000|0|0|10.0|||right|0|0"
    AddST "C|J19|n|5.337744304450858|0.00|0|0|10.0|||right|0|0"
    AddST "C|K19|n|39407|#,##0|0|0|10.0|||right|0|0"
    AddST "C|B20|s|1967-68|General|0|0|10.0||||0|1"
    AddST "C|C20|n|4.7796|0.000|0|0|10.0||||0|0"
    AddST "C|D20|n|2.731864588930692|0.00|0|0|10.0|||right|1|0"
    AddST "C|E20|n|42520|#,##0|0|0|10.0||||0|0"
    AddST "C|F20|n|42628|#,##0|0|0|10.0||||0|1"
    AddST "C|H20|s|1967|General|0|0|10.0|||left|0|1"
    AddST "C|I20|n|4.622|0.000|0|0|10.0|||right|0|0"
    AddST "C|J20|n|2.9490377761938618|0.00|0|0|10.0|||right|0|0"
    AddST "C|K20|n|41687|#,##0|0|0|10.0|||right|0|0"
    AddST "C|B21|s|1968-69|General|0|0|10.0||||0|1"
    AddST "C|C21|n|5.0299|0.000|0|0|10.0||||0|0"
    AddST "C|D21|n|5.236839902920732|0.00|0|0|10.0|||right|1|0"
    AddST "C|E21|n|46779|#,##0|0|0|10.0||||0|0"
    AddST "C|F21|n|46800|#,##0|0|0|10.0||||0|1"
    AddST "C|H21|s|1968|General|0|0|10.0|||left|0|1"
    AddST "C|I21|n|4.816|0.000|0|0|10.0|||right|0|0"
    AddST "C|J21|n|4.197317178710514|0.00|0|0|10.0|||right|0|0"
    AddST "C|K21|n|45811|#,##0|0|0|10.0|||right|0|0"
    AddST "C|B22|s|1969-70|General|0|0|10.0||||0|1"
    AddST "C|C22|n|5.3774|0.000|0|0|10.0||||0|0"
    AddST "C|D22|n|6.908686057376889|0.00|0|0|10.0|||right|1|0"
    AddST "C|E22|n|50856|#,##0|0|0|10.0||||0|0"
    AddST "C|F22|n|50908|#,##0|0|0|10.0||||0|1"
End Sub

Private Sub DataST4_2()
    Dim s As String
    AddST "C|H22|s|1969|General|0|0|10.0|||left|0|1"
    AddST "C|I22|n|5.1351|0.000|0|0|10.0|||right|0|0"
    AddST "C|J22|n|6.625830564784066|0.00|0|0|10.0|||right|0|0"
    AddST "C|K22|n|49776|#,##0|0|0|10.0|||right|0|0"
    AddST "C|B23|s|1970-71|General|0|0|10.0||||0|1"
    AddST "C|C23|n|5.9079|0.000|0|0|10.0||||0|0"
    AddST "C|D23|n|9.865362442816231|0.00|0|0|10.0|||right|1|0"
    AddST "C|E23|n|57739|#,##0|0|0|10.0||||0|0"
    AddST "C|F23|n|57823|#,##0|0|0|10.0||||0|1"
    AddST "C|H23|s|1970|General|0|0|10.0|||left|0|1"
    AddST "C|I23|n|5.6294|0.000|0|0|10.0|||right|0|0"
    AddST "C|J23|n|9.625907966738719|0.00|0|0|10.0|||right|0|0"
    AddST "C|K23|n|56056|#,##0|0|0|10.0|||right|0|0"
    AddST "C|B24|s|1971-72|General|0|0|10.0||||0|1"
    AddST "C|C24|n|6.3551|0.000|0|0|10.0||||0|0"
    AddST "C|D24|n|7.56952555053404|0.00|0|0|10.0|||right|1|0"
    AddST "C|E24|n|64591|#,##0|0|0|10.0||||0|0"
    AddST "C|F24|n|64499|#,##0|0|0|10.0||||0|1"
    AddST "C|H24|s|1971|General|0|0|10.0|||left|0|1"
    AddST "C|I24|n|6.0957|0.000|0|0|10.0|||right|0|0"
    AddST "C|J24|n|8.283298397697791|0.00|0|0|10.0|||right|0|0"
    AddST "C|K24|n|62886|#,##0|0|0|10.0|||right|0|0"
    AddST "C|B25|s|1972-73|General|0|0|10.0||||0|1"
    AddST "C|C25|n|6.8955|0.000|0|0|10.0||||0|0"
    AddST "C|D25|n|8.503406712718919|0.00|0|0|10.0|||right|1|0"
    AddST "C|E25|n|73942|#,##0|0|0|10.0||||0|0"
    AddST "C|F25|n|74042|#,##0|0|0|10.0||||0|1"
    AddST "C|H25|s|1972|General|0|0|10.0|||left|0|1"
    AddST "C|I25|n|6.5561|0.000|0|0|10.0|||right|0|0"
    AddST "C|J25|n|7.552865134439029|0.00|0|0|10.0|||right|0|0"
    AddST "C|K25|n|70595|#,##0|0|0|10.0|||right|0|0"
    AddST "C|B26|s|1973-74|General|0|0|10.0||||0|1"
    AddST "C|C26|n|7.5015|0.000|0|0|10.0||||0|0"
    AddST "C|D26|n|8.788340221883836|0.00|0|0|10.0|||right|1|0"
    AddST "C|E26|n|82871|#,##0|0|0|10.0||||0|0"
    AddST "C|F26|n|82895|#,##0|0|0|10.0||||0|1"
    AddST "C|H26|s|1973|General|0|0|10.0|||left|0|1"
    AddST "C|I26|n|7.1384|0.000|0|0|10.0|||right|0|0"
    AddST "C|J26|n|8.881804731471455|0.00|0|0|10.0|||right|0|0"
    AddST "C|K26|n|81809|#,##0|0|0|10.0|||right|0|0"
    AddST "C|B27|s|1974-75|General|0|0|10.0||||0|1"
    AddST "C|C27|n|9.0263|0.000|0|0|10.0||||0|0"
    AddST "C|D27|n|20.3266013463974|0.00|0|0|10.0|||right|1|0"
    AddST "C|E27|n|98192|#,##0|0|0|10.0||||0|0"
    AddST "C|F27|n|98365|#,##0|0|0|10.0||||0|1"
    AddST "C|H27|s|1974|General|0|0|10.0|||left|0|1"
    AddST "C|I27|n|8.2922|0.000|0|0|10.0|||right|0|0"
    AddST "C|J27|n|16.163285890395603|0.00|0|0|10.0|||right|0|0"
    AddST "C|K27|n|92686|#,##0|0|0|10.0|||right|0|0"
    AddST "C|B28|s|1975-76|General|0|0|10.0||||0|1"
    AddST "C|C28|n|11.2333|0.000|0|0|10.0||||0|0"
    AddST "C|D28|n|24.450771634002844|0.00|0|0|10.0|||right|1|0"
    AddST "C|E28|n|120844|#,##0|0|0|10.0||||0|0"
    AddST "C|F28|n|120728|#,##0|0|0|10.0||||0|1"
    AddST "C|H28|s|1975|General|0|0|10.0|||left|0|1"
    AddST "C|I28|n|10.4595|0.000|0|0|10.0|||right|0|0"
    AddST "C|J28|n|26.136610308482684|0.00|0|0|10.0|||right|0|0"
    AddST "C|K28|n|115095|#,##0|0|0|10.0|||right|0|0"
    AddST "C|B29|s|1976-77|General|0|0|10.0||||0|1"
    AddST "C|C29|n|12.8005|0.000|0|0|10.0||||0|0"
    AddST "C|D29|n|13.951376710316646|0.00|0|0|10.0|||right|1|0"
    AddST "C|E29|n|142072|#,##0|0|0|10.0||||0|0"
    AddST "C|F29|n|142583|#,##0|0|0|10.0||||0|1"
    AddST "C|H29|s|1976|General|0|0|10.0|||left|0|1"
    AddST "C|I29|n|12.0781|0.000|0|0|10.0|||right|0|0"
    AddST "C|J29|n|15.474927099765754|0.00|0|0|10.0|||right|0|0"
    AddST "C|K29|n|136995|#,##0|0|0|10.0|||right|0|0"
    AddST "C|B30|s|1977-78|General|0|0|10.0||||0|1"
    AddST "C|C30|n|14.5596|0.000|0|0|10.0||||0|0"
    AddST "C|D30|n|13.742431936252492|0.00|0|0|10.0|||right|1|0"
    AddST "C|E30|n|166087|#,##0|0|0|10.0||||0|0"
    AddST "C|F30|n|165875|#,##0|0|0|10.0||||0|1"
    AddST "C|H30|s|1977|General|0|0|10.0|||left|0|1"
    AddST "C|I30|n|13.752|0.000|0|0|10.0|||right|0|0"
    AddST "C|J30|n|13.858967884021506|0.00|0|0|10.0|||right|0|0"
    AddST "C|K30|n|159775|#,##0|0|0|10.0|||right|0|0"
    AddST "C|B31|s|1978-79|General|0|0|10.0||||0|1"
    AddST "C|C31|n|16.1988|0.000|0|0|10.0||||0|0"
    AddST "C|D31|n|11.258551059095023|0.00|0|0|10.0|||right|1|0"
    AddST "C|E31|n|192341|#,##0|0|0|10.0||||0|0"
    AddST "C|F31|n|192460|#,##0|0|0|10.0||||0|1"
    AddST "C|H31|s|1978|General|0|0|10.0|||left|0|1"
    AddST "C|I31|n|15.3753|0.000|0|0|10.0|||right|0|0"
    AddST "C|J31|n|11.804101221640478|0.00|0|0|10.0|||right|0|0"
    AddST "C|K31|n|186141|#,##0|0|0|10.0|||right|0|0"
    AddST "C|B32|s|1979-80|General|0|0|10.0||||0|1"
    AddST "C|C32|n|18.9324|0.000|0|0|10.0||||0|0"
    AddST "C|D32|n|16.875324098081357|0.00|0|0|10.0|||right|1|0"
    AddST "C|E32|n|232557|#,##0|0|0|10.0||||0|0"
    AddST "C|F32|n|232594|#,##0|0|0|10.0||||0|1"
    AddST "C|H32|s|1979|General|0|0|10.0|||left|0|1"
    AddST "C|I32|n|17.5958|0.000|0|0|10.0|||right|0|0"
    AddST "C|J32|n|14.441994627747109|0.00|0|0|10.0|||right|0|0"
    AddST "C|K32|n|220966|#,##0|0|0|10.0|||right|0|0"
    AddST "C|B33|s|1980-81|General|0|0|10.0||||0|1"
    AddST "C|C33|n|22.5457|0.000|0|0|10.0||||0|0"
    AddST "C|D33|n|19.085271809173683|0.00|0|0|10.0|||right|1|0"
    AddST "C|E33|n|267548|#,##0|0|0|10.0||||0|0"
    AddST "C|F33|n|267830|#,##0|0|0|10.0||||0|1"
    AddST "C|H33|s|1980|General|0|0|10.0|||left|0|1"
    AddST "C|I33|n|21.1501|0.000|0|0|10.0|||right|0|0"
    AddST "C|J33|n|20.19970674820126|0.00|0|0|10.0|||right|0|0"
    AddST "C|K33|n|259956|#,##0|0|0|10.0|||right|0|0"
    AddST "C|B34|s|1981-82|General|0|0|10.0||||0|1"
    AddST "C|C34|n|24.9204|0.000|0|0|10.0||||0|0"
    AddST "C|D34|n|10.532828876459815|0.00|0|0|10.0|||right|1|0"
    AddST "C|E34|n|298739|#,##0|0|0|10.0||||0|0"
    AddST "C|F34|n|298186|#,##0|0|0|10.0||||0|1"
    AddST "C|H34|s|1981|General|0|0|10.0|||left|0|1"
    AddST "C|I34|n|23.7558|0.000|0|0|10.0|||right|0|0"
    AddST "C|J34|n|12.320036311885062|0.00|0|0|10.0|||right|0|0"
    AddST "C|K34|n|290550|#,##0|0|0|10.0|||right|0|0"
    AddST "C|B35|s|1982-83|General|0|0|10.0||||0|1"
    AddST "C|C35|n|26.7569|0.000|0|0|10.0||||0|0"
    AddST "C|D35|n|7.369464374568629|0.00|0|0|10.0|||right|1|0"
    AddST "C|E35|n|327910|#,##0|0|0|10.0||||0|0"
    AddST "C|F35|n|328278|#,##0|0|0|10.0||||0|1"
    AddST "C|H35|s|1982|General|0|0|10.0|||left|0|1"
    AddST "C|I35|n|25.6563|0.000|0|0|10.0|||right|0|0"
    AddST "C|J35|n|8.000151541939235|0.00|0|0|10.0|||right|0|0"
    AddST "C|K35|n|320013|#,##0|0|0|10.0|||right|0|0"
    AddST "C|B36|s|1983-84|General|0|0|10.0||||0|1"
    AddST "C|C36|n|28.0299|0.000|0|0|10.0||||0|0"
    AddST "C|D36|n|4.7576512974223455|0.00|0|0|10.0|||right|1|0"
    AddST "C|E36|n|358644|#,##0|0|0|10.0||||0|0"
    AddST "C|F36|n|357864|#,##0|0|0|10.0||||0|1"
    AddST "C|H36|s|1983|General|0|0|10.0|||left|0|1"
    AddST "C|I36|n|27.0789|0.000|0|0|10.0|||right|0|0"
    AddST "C|J36|n|5.544836940634461|0.00|0|0|10.0|||right|0|0"
    AddST "C|K36|n|351900|#,##0|0|0|10.0|||right|0|0"
    AddST "C|B37|s|1984-85|General|0|0|10.0||||0|1"
    AddST "C|C37|n|29.6689|0.000|0|0|10.0||||0|0"
    AddST "C|D37|n|5.847327318327926|0.00|0|0|10.0|||right|1|0"
    AddST "C|E37|n|386718|#,##0|0|0|10.0||||0|0"
    AddST "C|F37|n|386407|#,##0|0|0|10.0||||0|1"
    AddST "C|H37|s|1984|General|0|0|10.0|||left|0|1"
    AddST "C|I37|n|28.4962|0.000|0|0|10.0|||right|0|0"
    AddST "C|J37|n|5.2339644520272275|0.00|0|0|10.0|||right|0|0"
    AddST "C|K37|n|378608|#,##0|0|0|10.0|||right|0|0"
    AddST "C|B38|s|1985-86|General|0|0|10.0||||0|1"
    AddST "C|C38|n|31.2847|0.000|0|0|10.0||||0|0"
    AddST "C|D38|n|5.446106866112327|0.00|0|0|10.0|||right|1|0"
    AddST "C|E38|n|424559|#,##0|0|0|10.0||||0|0"
    AddST "C|F38|n|424801|#,##0|0|0|10.0||||0|1"
    AddST "C|H38|s|1985|General|0|0|10.0|||left|0|1"
    AddST "C|I38|n|30.0366|0.000|0|0|10.0|||right|0|0"
    AddST "C|J38|n|5.405633031772651|0.00|0|0|10.0|||right|0|0"
    AddST "C|K38|n|415455|#,##0|0|0|10.0|||right|0|0"
    AddST "C|B39|s|1986-87|General|0|0|10.0||||0|1"
    AddST "C|C39|n|32.5556|0.000|0|0|10.0||||0|0"
End Sub

Private Sub DataST4_3()
    Dim s As String
    AddST "C|D39|n|4.06236914530105|0.00|0|0|10.0|||right|1|0"
    AddST "C|E39|n|456205|#,##0|0|0|10.0||||0|0"
    AddST "C|F39|n|457005|#,##0|0|0|10.0||||0|1"
    AddST "C|H39|s|1986|General|0|0|10.0|||left|0|1"
    AddST "C|I39|n|31.3802|0.000|0|0|10.0|||right|0|0"
    AddST "C|J39|n|4.473209351258127|0.00|0|0|10.0|||right|0|0"
    AddST "C|K39|n|447561|#,##0|0|0|10.0|||right|0|0"
    AddST "C|B40|s|1987-88|General|0|0|10.0||||0|1"
    AddST "C|C40|n|34.4507|0.000|0|0|10.0||||0|0"
    AddST "C|D40|n|5.821118332944254|0.00|0|0|10.0|||right|1|0"
    AddST "C|E40|n|512832|#,##0|0|0|10.0||||0|0"
    AddST "C|F40|n|513277|#,##0|0|0|10.0||||0|1"
    AddST "C|H40|s|1987|General|0|0|10.0|||left|0|1"
    AddST "C|I40|n|33.1066|0.000|0|0|10.0|||right|0|0"
    AddST "C|J40|n|5.501558307467772|0.00|0|0|10.0|||right|0|0"
    AddST "C|K40|n|498040|#,##0|0|0|10.0|||right|0|0"
    AddST "C|B41|s|1988-89|General|0|0|10.0||||0|1"
    AddST "C|C41|n|36.7965|0.000|0|0|10.0||||0|0"
    AddST "C|D41|n|6.809150467189358|0.00|0|0|10.0|||right|1|0"
    AddST "C|E41|n|572031|#,##0|0|0|10.0||||0|0"
    AddST "C|F41|n|572718|#,##0|0|0|10.0||||0|1"
    AddST "C|H41|s|1988|General|0|0|10.0|||left|0|1"
    AddST "C|I41|n|35.1435|0.000|0|0|10.0|||right|0|0"
    AddST "C|J41|n|6.152549642669445|0.00|0|0|10.0|||right|0|0"
    AddST "C|K41|n|557277|#,##0|0|0|10.0|||right|0|0"
    AddST "C|B42|s|1989-90|General|0|0|10.0||||0|1"
    AddST "C|C42|n|39.7716|0.000|0|0|10.0||||0|0"
    AddST "C|D42|n|8.085279849985726|0.00|0|0|10.0|||right|1|0"
    AddST "C|E42|n|631607|#,##0|0|0|10.0||||0|0"
    AddST "C|F42|n|631097|#,##0|0|0|10.0||||0|1"
    AddST "C|H42|s|1989|General|0|0|10.0|||left|0|1"
    AddST "C|I42|n|37.9703|0.000|0|0|10.0|||right|0|0"
    AddST "C|J42|n|8.043592698507544|0.00|0|0|10.0|||right|0|0"
    AddST "C|K42|n|616595|#,##0|0|0|10.0|||right|0|0"
    AddST "C|B43|s|1990-91|General|0|0|10.0||||0|1"
    AddST "C|C43|n|43.1083|0.000|0|0|10.0||||0|0"
    AddST "C|D43|n|8.389654929648295|0.00|0|0|10.0|||right|1|0"
    AddST "C|E43|n|681464|#,##0|0|0|10.0||||0|0"
    AddST "C|F43|n|683131|#,##0|0|0|10.0||||0|1"
    AddST "C|H43|s|1990|General|0|0|10.0|||left|0|1"
    AddST "C|I43|n|41.215|0.000|0|0|10.0|||right|0|0"
    AddST "C|J43|n|8.545363086412278|0.00|0|0|10.0|||right|0|0"
    AddST "C|K43|n|672977|#,##0|0|0|10.0|||right|0|0"
    AddST "C|B44|s|1991-92|General|0|0|10.0||||0|1"
    AddST "C|C44|n|45.701|0.000|0|0|10.0||||0|0"
    AddST "C|D44|n|6.014387020596963|0.00|0|0|10.0|||right|1|0"
    AddST "C|E44|n|716584|#,##0|0|0|10.0||||0|0"
    AddST "C|F44|n|716351|#,##0|0|0|10.0||||0|1"
    AddST "C|H44|s|1991|General|0|0|10.0|||left|0|1"
    AddST "C|I44|n|43.9785|0.000|0|0|10.0|||right|0|0"
    AddST "C|J44|n|6.705083100812795|0.00|0|0|10.0|||right|0|0"
    AddST "C|K44|n|708072|#,##0|0|0|10.0|||right|0|0"
    AddST "C|B45|s|1992-93|General|0|0|10.0||||0|1"
    AddST "C|C45|n|46.9538|0.000|0|0|10.0||||0|0"
    AddST "C|D45|n|2.7412966893503437|0.00|0|0|10.0|||right|1|0"
    AddST "C|E45|n|741206|#,##0|0|0|10.0||||0|0"
    AddST "C|F45|n|741414|#,##0|0|0|10.0||||0|1"
    AddST "C|H45|s|1992|General|0|0|10.0|||left|0|1"
    AddST "C|I45|n|45.4142|0.000|0|0|10.0|||right|0|0"
    AddST "C|J45|n|3.26454972316019|0.00|0|0|10.0|||right|0|0"
    AddST "C|K45|n|733221|#,##0|0|0|10.0|||right|0|0"
    AddST "C|B46|s|1993-94|General|0|0|10.0||||0|1"
    AddST "C|C46|n|48.2899|0.000|0|0|10.0||||0|0"
    AddST "C|D46|n|2.845563085415881|0.00|0|0|10.0|||right|1|0"
    AddST "C|E46|n|783220|#,##0|0|0|10.0||||0|0"
    AddST "C|F46|n|782644|#,##0|0|0|10.0||||0|1"
    AddST "C|H46|s|1993|General|0|0|10.0|||left|0|1"
    AddST "C|I46|n|46.7746|0.000|0|0|10.0|||right|0|0"
    AddST "C|J46|n|2.9955388402746244|0.00|0|0|10.0|||right|0|0"
    AddST "C|K46|n|772390|#,##0|0|0|10.0|||right|0|0"
    AddST "C|B47|s|1994-95|General|0|0|10.0||||0|1"
    AddST "C|C47|n|49.091|0.000|0|0|10.0||||0|0"
    AddST "C|D47|n|1.6589390327998155|0.00|0|0|10.0|||right|1|0"
    AddST "C|E47|n|823670|#,##0|0|0|10.0||||0|0"
    AddST "C|F47|n|822442|#,##0|0|0|10.0||||0|1"
    AddST "C|H47|s|1994|General|0|0|10.0|||left|0|1"
    AddST "C|I47|n|47.6384|0.000|0|0|10.0|||right|0|0"
    AddST "C|J47|n|1.8467287801499053|0.00|0|0|10.0|||right|0|0"
    AddST "C|K47|n|813527|#,##0|0|0|10.0|||right|0|0"
    AddST "C|B48|s|1995-96|General|0|0|10.0||||0|1"
    AddST "C|C48|n|50.6447|0.000|0|0|10.0||||0|0"
    AddST "C|D48|n|3.164938583447066|0.00|0|0|10.0|||right|1|0"
    AddST "C|E48|n|866285|#,##0|0|0|10.0||||0|0"
    AddST "C|F48|n|867839|#,##0|0|0|10.0||||0|1"
    AddST "C|H48|s|1995|General|0|0|10.0|||left|0|1"
    AddST "C|I48|n|48.8847|0.000|0|0|10.0|||right|0|0"
    AddST "C|J48|n|2.6161667898166296|0.00|0|0|10.0|||right|0|0"
    AddST "C|K48|n|854921|#,##0|0|0|10.0|||right|0|0"
    AddST "C|B49|s|1996-97|General|0|0|10.0||||0|1"
    AddST "C|C49|n|52.3579|0.000|0|0|10.0||||0|0"
    AddST "C|D49|n|3.3827824036868623|0.00|0|0|10.0|||right|1|0"
    AddST "C|E49|n|926554|#,##0|0|0|10.0||||0|0"
    AddST "C|F49|n|924649|#,##0|0|0|10.0||||0|1"
    AddST "C|H49|s|1996|General|0|0|10.0|||left|0|1"
    AddST "C|I49|n|50.9091|0.000|0|0|10.0|||right|0|0"
    AddST "C|J49|n|4.14117300505066|0.00|0|0|10.0|||right|0|0"
    AddST "C|K49|n|913609|#,##0|0|0|10.0|||right|0|0"
    AddST "C|B50|s|1997-98|General|0|0|10.0||||0|1"
    AddST "C|C50|n|52.4935|0.000|0|0|10.0||||0|0"
    AddST "C|D50|n|0.25898670496715226|0.00|0|0|10.0|||right|1|0"
    AddST "C|E50|n|970388|#,##0|0|0|10.0||||0|0"
    AddST "C|F50|n|972370|#,##0|0|0|10.0||||0|1"
    AddST "C|H50|s|1997|General|0|0|10.0|||left|0|1"
    AddST "C|I50|n|50.929|0.000|0|0|10.0|||right|0|0"
    AddST "C|J50|n|0.03908927873405699|0.00|0|0|10.0|||right|0|0"
    AddST "C|K50|n|958539|#,##0|0|0|10.0|||right|0|0"
    AddST "C|B51|s|1998-99|General|0|0|10.0||||0|1"
    AddST "C|C51|n|53.2311|0.000|0|0|10.0||||0|0"
    AddST "C|D51|n|1.405126348976541|0.00|0|0|10.0|||right|1|0"
    AddST "C|E51|n|1014098|#,##0|0|0|10.0||||0|0"
    AddST "C|F51|n|1012744|#,##0|0|0|10.0||||0|1"
    AddST "C|H51|s|1998|General|0|0|10.0|||left|0|1"
    AddST "C|I51|n|51.606|0.000|0|0|10.0|||right|0|0"
    AddST "C|J51|n|1.3293015767048235|0.00|0|0|10.0|||right|0|0"
    AddST "C|K51|n|1002804|#,##0|0|0|10.0|||right|0|0"
    AddST "C|B52|s|1999-00|General|0|0|10.0||||0|1"
    AddST "C|C52|n|53.8894|0.000|0|0|10.0||||0|0"
    AddST "C|D52|n|1.2366830668537832|0.00|0|0|10.0|||right|1|0"
    AddST "C|E52|n|1060973|#,##0|0|0|10.0||||0|0"
    AddST "C|F52|n|1063163|#,##0|0|0|10.0||||0|1"
    AddST "C|H52|s|1999|General|0|0|10.0|||left|0|1"
    AddST "C|I52|n|52.3258|0.000|0|0|10.0|||right|0|0"
    AddST "C|J52|n|1.3947990543735211|0.00|0|0|10.0|||right|0|0"
    AddST "C|K52|n|1046727|#,##0|0|0|10.0|||right|0|0"
    AddST "C|B53|s|2000-01|General|0|0|10.0||||0|1"
    AddST "C|C53|n|54.3697|0.000|0|0|10.0||||0|0"
    AddST "C|D53|n|0.8912698972339638|0.00|0|0|10.0|||right|1|0"
    AddST "C|E53|n|1117633|#,##0|0|0|10.0||||0|0"
    AddST "C|F53|n|1114932|#,##0|0|0|10.0||||0|1"
    AddST "C|H53|s|2000|General|0|0|10.0|||left|0|1"
    AddST "C|I53|n|52.8377|0.000|0|0|10.0|||right|0|0"
    AddST "C|J53|n|0.9782936906841313|0.00|0|0|10.0|||right|0|0"
    AddST "C|K53|n|1104759|#,##0|0|0|10.0|||right|0|0"
    AddST "C|B54|s|2001-02|General|0|0|10.0||||0|1"
    AddST "C|C54|n|55.4597|0.000|0|0|10.0||||0|0"
    AddST "C|D54|n|2.004793110868731|0.00|0|0|10.0|||right|1|0"
    AddST "C|E54|n|1159585|#,##0|0|0|10.0||||0|0"
    AddST "C|F54|n|1160696|#,##0|0|0|10.0||||0|1"
    AddST "C|H54|s|2001|General|0|0|10.0|||left|0|1"
    AddST "C|I54|n|53.7416|0.000|0|0|10.0|||right|0|0"
    AddST "C|J54|n|1.710710345075581|0.00|0|0|10.0|||right|0|0"
    AddST "C|K54|n|1150542|#,##0|0|0|10.0|||right|0|0"
    AddST "C|B55|s|2002-03|General|0|0|10.0||||0|1"
    AddST "C|C55|n|56.6|0.000|0|0|10.0||||0|0"
    AddST "C|D55|n|2.056087573499322|0.00|0|0|10.0|||right|1|0"
    AddST "C|E55|n|1210146|#,##0|0|0|10.0||||0|0"
    AddST "C|F55|n|1208344|#,##0|0|0|10.0||||0|1"
    AddST "C|H55|s|2002|General|0|0|10.0|||left|0|1"
    AddST "C|I55|n|54.8629|0.000|0|0|10.0|||right|0|0"
    AddST "C|J55|n|2.0864656057877045|0.00|0|0|10.0|||right|0|0"
End Sub

Private Sub DataST4_4()
    Dim s As String
    AddST "C|K55|n|1194646|#,##0|0|0|10.0|||right|0|0"
    AddST "C|B56|s|2003-04|General|0|0|10.0||||0|1"
    AddST "C|C56|n|57.9885|0.000|0|0|10.0||||0|0"
    AddST "C|D56|n|2.453180212014135|0.00|0|0|10.0|||right|1|0"
    AddST "C|E56|n|1277594|#,##0|0|0|10.0||||0|0"
    AddST "C|F56|n|1279226|#,##0|0|0|10.0||||0|1"
    AddST "C|H56|s|2003|General|0|0|10.0|||left|0|1"
    AddST "C|I56|n|56.1691|0.000|0|0|10.0|||right|0|0"
    AddST "C|J56|n|2.3808438853943135|0.00|0|0|10.0|||right|0|0"
    AddST "C|K56|n|1262502|#,##0|0|0|10.0|||right|0|0"
    AddST "C|B57|s|2004-05|General|0|0|10.0||||0|1"
    AddST "C|C57|n|59.7172|0.000|0|0|10.0||||0|0"
    AddST "C|D57|n|2.9811083232020077|0.00|0|0|10.0|||right|1|0"
    AddST "C|E57|n|1346417|#,##0|0|0|10.0||||0|0"
    AddST "C|F57|n|1344682|#,##0|0|0|10.0||||0|1"
    AddST "C|H57|s|2004|General|0|0|10.0|||left|0|1"
    AddST "C|I57|n|57.7042|0.000|0|0|10.0|||right|0|0"
    AddST "C|J57|n|2.732997324151535|0.00|0|0|10.0|||right|0|0"
    AddST "C|K57|n|1327143|#,##0|0|0|10.0|||right|0|0"
    AddST "C|B58|s|2005-06|General|0|0|10.0||||0|1"
    AddST "C|C58|n|61.3576|0.000|0|0|10.0||||0|0"
    AddST "C|D58|n|2.746947278171113|0.00|0|0|10.0|||right|1|0"
    AddST "C|E58|n|1423557|#,##0|0|0|10.0||||0|0"
    AddST "C|F58|n|1424005|#,##0|0|0|10.0||||0|1"
    AddST "C|H58|s|2005|General|0|0|10.0|||left|0|1"
    AddST "C|I58|n|59.3559|0.000|0|0|10.0|||right|0|0"
    AddST "C|J58|n|2.862356639551364|0.00|0|0|10.0|||right|0|0"
    AddST "C|K58|n|1403249|#,##0|0|0|10.0|||right|0|0"
    AddST "C|B59|s|2006-07|General|0|0|10.0||||0|1"
    AddST "C|C59|n|63.2017|0.000|0|0|10.0||||0|0"
    AddST "C|D59|n|3.005495651720414|0.00|0|0|10.0|||right|1|0"
    AddST "C|E59|n|1493448|#,##0|0|0|10.0||||0|0"
    AddST "C|F59|n|1494157|#,##0|0|0|10.0||||0|1"
    AddST "C|H59|s|2006|General|0|0|10.0|||left|0|1"
    AddST "C|I59|n|61.1841|0.000|0|0|10.0|||right|0|0"
    AddST "C|J59|n|3.0800644923251146|0.00|0|0|10.0|||right|0|0"
    AddST "C|K59|n|1478080|#,##0|0|0|10.0|||right|0|0"
    AddST "C|O59|s|  |General|0|0|10.0||||0|0"
    AddST "C|B60|s|2007-08|General|0|0|10.0||||0|1"
    AddST "C|C60|n|64.438|0.000|0|0|10.0||||0|0"
    AddST "C|D60|n|1.956118268970613|0.00|0|0|10.0|||right|1|0"
    AddST "C|E60|n|1571927|#,##0|0|0|10.0||||0|0"
    AddST "C|F60|n|1572773|#,##0|0|0|10.0||||0|1"
    AddST "C|H60|s|2007|General|0|0|10.0|||left|0|1"
    AddST "C|I60|n|62.4026|0.000|0|0|10.0|||right|0|0"
    AddST "C|J60|n|1.9915304793238748|0.00|0|0|10.0|||right|0|0"
    AddST "C|K60|n|1551729|#,##0|0|0|10.0|||right|0|0"
    AddST "C|B61|s|2008-09|General|0|0|10.0||||0|1"
    AddST "C|C61|n|66.8397|0.000|0|0|10.0||||0|0"
    AddST "C|D61|n|3.7271485769266444|0.00|0|0|10.0|||right|1|0"
    AddST "C|E61|n|1593267|#,##0|0|0|10.0||||0|0"
    AddST "C|F61|n|1592276|#,##0|0|0|10.0||||0|1"
    AddST "C|H61|s|2008|General|0|0|10.0|||left|0|1"
    AddST "C|I61|n|64.4827|0.000|0|0|10.0|||right|0|0"
    AddST "C|J61|n|3.3333546999644157|0.00|0|0|10.0|||right|0|0"
    AddST "C|K61|n|1602217|#,##0|0|0|10.0|||right|0|0"
    AddST "C|B62|s|2009-10|General|0|0|10.0||||0|1"
    AddST "C|C62|n|67.7042|0.000|0|0|10.0||||0|0"
    AddST "C|D62|n|1.2933929984724748|0.00|0|0|10.0|||right|1|0"
    AddST "C|E62|n|1567070|#,##0|0|0|10.0||||0|0"
    AddST "C|F62|n|1567756|#,##0|0|0|10.0||||0|1"
    AddST "C|H62|s|2009|General|0|0|10.0|||left|0|1"
    AddST "C|I62|n|65.7619|0.000|0|0|10.0|||right|0|0"
    AddST "C|J62|n|1.9837878997002345|0.00|0|0|10.0|||right|0|0"
    AddST "C|K62|n|1559405|#,##0|0|0|10.0|||right|0|0"
    AddST "C|B63|s|2010-11|General|0|0|10.0||||0|1"
    AddST "C|C63|n|68.8561|0.000|0|0|10.0||||0|0"
    AddST "C|D63|n|1.701371554497354|0.00|0|0|10.0|||right|1|0"
    AddST "C|E63|n|1632926|#,##0|0|0|10.0||||0|0"
    AddST "C|F63|n|1634234|#,##0|0|0|10.0||||0|1"
    AddST "C|H63|s|2010|General|0|0|10.0|||left|0|1"
    AddST "C|I63|n|66.6341|0.000|0|0|10.0|||right|0|0"
    AddST "C|J63|n|1.3262998788052147|0.00|0|0|10.0|||right|0|0"
    AddST "C|K63|n|1615839|#,##0|0|0|10.0|||right|0|0"
    AddST "C|B64|s|2011-12|General|0|0|10.0||||0|1"
    AddST "C|C64|n|70.3261|0.000|0|0|10.0||||0|0"
    AddST "C|D64|n|2.134887105136653|0.00|0|0|10.0|||right|1|0"
    AddST "C|E64|n|1680937|#,##0|0|0|10.0||||0|0"
    AddST "C|F64|n|1680077|#,##0|0|0|10.0||||0|1"
    AddST "C|H64|n|2011|General|0|0|10.0|||left|0|1"
    AddST "C|I64|n|68.2848|0.000|0|0|10.0|||right|0|0"
    AddST "C|J64|n|2.477260141579162|0.00|0|0|10.0|||right|0|0"
    AddST "C|K64|n|1669945|#,##0|0|0|10.0|||right|0|0"
    AddST "C|B65|s|2012-13|General|0|0|10.0||||0|1"
    AddST "C|C65|n|71.5454|0.000|0|0|10.0||||0|0"
    AddST "C|D65|n|1.7337802039356711|0.00|0|0|10.0|||right|1|0"
    AddST "C|E65|n|1733512|#,##0|0|0|10.0||||0|0"
    AddST "C|F65|n|1735093|#,##0|0|0|10.0||||0|1"
    AddST "C|H65|n|2012|General|0|0|10.0|||left|0|1"
    AddST "C|I65|n|69.3408|0.000|0|0|10.0|||right|0|0"
    AddST "C|J65|n|1.5464642204414414|0.00|0|0|10.0|||right|0|0"
    AddST "C|K65|n|1721708|#,##0|0|0|10.0|||right|0|0"
    AddST "C|B66|s|2013-14|General|0|0|10.0||||0|1"
    AddST "C|C66|n|73.0337|0.000|0|0|10.0||||0|0"
    AddST "C|D66|n|2.0802175960998124|0.00|0|0|10.0|||right|1|0"
    AddST "C|E66|n|1812306|#,##0|0|0|10.0||||0|0"
    AddST "C|F66|n|1809716|#,##0|0|0|10.0||||0|1"
    AddST "C|H66|n|2013|General|0|0|10.0|||left|0|1"
    AddST "C|I66|n|70.835|0.000|0|0|10.0|||right|0|0"
    AddST "C|J66|n|2.154864091559359|0.00|0|0|10.0|||right|0|0"
    AddST "C|K66|n|1789072|#,##0|0|0|10.0|||right|0|0"
    AddST "C|B67|s|2014-15|General|0|0|10.0||||0|1"
    AddST "C|C67|n|74.0662|0.000|0|0|10.0||||0|0"
    AddST "C|D67|n|1.413730921478713|0.00|0|0|10.0|||right|1|0"
    AddST "C|E67|n|1888114|#,##0|0|0|10.0||||0|0"
    AddST "C|F67|n|1892400|#,##0|0|0|10.0||||0|1"
    AddST "C|H67|n|2014|General|0|0|10.0|||left|0|1"
    AddST "C|I67|n|71.9658|0.000|0|0|10.0|||right|0|0"
    AddST "C|J67|n|1.596385967389014|0.00|0|0|10.0|||right|0|0"
    AddST "C|K67|n|1875066|#,##0|0|0|10.0|||right|0|0"
    AddST "C|B68|s|2015-16|General|0|0|10.0||||0|1"
    AddST "C|C68|n|74.5804|0.000|0|0|10.0||||0|0"
    AddST "C|D68|n|0.6942437981157431|0.00|0|0|10.0|||right|1|0"
    AddST "C|E68|n|1943837|#,##0|0|0|10.0||||0|0"
    AddST "C|F68|n|1942644|#,##0|0|0|10.0||||0|1"
    AddST "C|H68|n|2015|General|0|0|10.0|||left|0|1"
    AddST "C|I68|n|72.449|0.000|0|0|10.0|||right|0|0"
    AddST "C|J68|n|0.6714300403802869|0.00|0|0|10.0|||right|0|0"
    AddST "C|K68|n|1928016|#,##0|0|0|10.0|||right|0|0"
    AddST "C|B69|s|2016-17|General|0|0|10.0||||0|1"
    AddST "C|C69|n|76.0754|0.000|0|0|10.0||||0|0"
    AddST "C|D69|n|2.00454811183636|0.00|0|0|10.0|||right|0|0"
    AddST "C|E69|n|2027118|#,##0|0|0|10.0||||0|0"
    AddST "C|F69|n|2029417|#,##0|0|0|10.0||||0|1"
    AddST "C|H69|n|2016|General|0|0|10.0|||left|0|1"
    AddST "C|I69|n|73.7039|0.000|0|0|10.0|||right|0|0"
    AddST "C|J69|n|1.7321150050380356|0.00|0|0|10.0|||right|0|0"
    AddST "C|K69|n|2004755|#,##0|0|0|10.0|||right|0|0"
    AddST "C|B70|s|2017-18|General|0|0|10.0||||0|1"
    AddST "C|C70|n|77.0388|0.000|0|0|10.0||||0|0"
    AddST "C|D70|n|1.2663752014448728|0.00|0|0|10.0|||right|0|0"
    AddST "C|E70|n|2114617|#,##0|0|0|10.0||||0|0"
    AddST "C|F70|n|2115079|#,##0|0|0|10.0||||0|1"
    AddST "C|H70|n|2017|General|0|0|10.0|||left|0|1"
    AddST "C|I70|n|74.8427|0.000|0|0|10.0|||right|0|0"
    AddST "C|J70|n|1.5451014125439617|0.00|0|0|10.0|||right|0|0"
    AddST "C|K70|n|2097222|#,##0|0|0|10.0|||right|0|0"
    AddST "C|B71|s|2018-19|General|0|0|10.0||||0|1"
    AddST "C|C71|n|78.7948|0.000|0|0|10.0||||0|0"
    AddST "C|D71|n|2.2793709143963827|0.00|0|0|10.0|||right|0|0"
    AddST "C|E71|n|2189839|#,##0|0|0|10.0||||0|0"
    AddST "C|F71|n|2190380|#,##0|0|0|10.0||||0|1"
    AddST "C|H71|n|2018|General|0|0|10.0|||left|0|1"
    AddST "C|I71|n|76.3117|0.000|0|0|10.0|||right|0|0"
    AddST "C|J71|n|1.9627832774606053|0.00|0|0|10.0|||right|0|0"
    AddST "C|K71|n|2171559|#,##0|0|0|10.0|||right|0|0"
    AddST "C|B72|s|2019-20|General|0|0|10.0||||0|1"
    AddST "C|C72|n|80.8738|0.000|0|0|10.0||||0|0"
    AddST "C|D72|n|2.638498987242823|0.00|0|0|10.0|||right|0|0"
    AddST "C|E72|n|2263357|#,##0|0|0|10.0||||0|0"
End Sub

Private Sub DataST4_5()
    Dim s As String
    AddST "C|F72|n|2258342|#,##0|0|0|10.0||||0|1"
    AddST "C|H72|n|2019|General|0|0|10.0|||left|0|1"
    AddST "C|I72|n|78.1895|0.000|0|0|10.0|||right|0|0"
    AddST "C|J72|n|2.4606973766801072|0.00|0|0|10.0|||right|0|0"
    AddST "C|K72|n|2252961|#,##0|0|0|10.0|||right|0|0"
    AddST "C|B73|s|2020-21|General|0|0|10.0||||0|1"
    AddST "C|C73|n|85.1002|0.000|0|0|10.0||||0|0"
    AddST "C|D73|n|5.225919890001456|0.00|0|0|10.0|||right|0|0"
    AddST "C|E73|n|2109996|#,##0|0|0|10.0||||0|0"
    AddST "C|F73|n|2110471|#,##0|0|0|10.0||||0|1"
    AddST "C|H73|n|2020|General|0|0|10.0|||left|0|1"
    AddST "C|I73|n|82.1469|0.000|0|0|10.0|||right|0|0"
    AddST "C|J73|n|5.061293396172129|0.00|0|0|10.0|||right|0|0"
    AddST "C|K73|n|2124720|#,##0|0|0|10.0|||right|0|0"
    AddST "C|B74|s|2021-22|General|0|0|10.0||||0|1"
    AddST "C|C74|n|85.3036|0.000|0|0|10.0||||0|0"
    AddST "C|D74|n|0.239012364248265|0.00|0|0|10.0|||right|0|0"
    AddST "C|E74|n|2400398|#,##0|0|0|10.0||||0|0"
    AddST "C|F74|n|2400550|#,##0|0|0|10.0||||0|1"
    AddST "C|H74|n|2021|General|0|0|10.0|||left|0|1"
    AddST "C|I74|n|82.5539|0.000|0|0|10.0|||right|0|0"
    AddST "C|J74|n|0.4954538759222764|0.00|0|0|10.0|||right|0|0"
    AddST "C|K74|n|2322652|#,##0|0|0|10.0|||right|0|0"
    AddST "C|B75|s|2022-23|General|0|0|10.0||||0|1"
    AddST "C|C75|n|91.2996|0.000|0|0|10.0||||0|0"
    AddST "C|D75|n|7.029011671254197|0.00|0|0|10.0|||right|0|0"
    AddST "C|E75|n|2634336|#,##0|0|0|10.0||||0|0"
    AddST "C|F75|n|2637129|#,##0|0|0|10.0||||0|1"
    AddST "C|H75|n|2022|General|0|0|10.0|||left|0|1"
    AddST "C|I75|n|87.2435|0.000|0|0|10.0|||right|0|0"
    AddST "C|J75|n|5.680652276876075|0.00|0|0|10.0|||right|0|0"
    AddST "C|K75|n|2580949|#,##0|0|0|10.0|||right|0|0"
    AddST "C|B76|s|2023-24|General|0|0|10.0||||0|1"
    AddST "C|C76|n|96.1117|0.000|0|0|10.0||||0|0"
    AddST "C|D76|n|5.27066931289951|0.00|0|0|10.0|||right|0|0"
    AddST "C|E76|n|2790164|#,##0|0|0|10.0||||0|0"
    AddST "C|F76|n|2778543|#,##0|0|0|10.0||||0|1"
    AddST "C|H76|n|2023|General|0|0|10.0|||left|0|1"
    AddST "C|I76|n|92.7868|0.000|0|0|10.0|||right|0|0"
    AddST "C|J76|n|6.353825786448277|0.00|0|0|10.0|||right|0|0"
    AddST "C|K76|n|2752164|#,##0|0|0|10.0|||right|0|0"
    AddST "C|B77|s|2024-25|General|0|0|10.0||||0|1"
    AddST "C|C77|n|100|0.000|0|0|10.0||||0|0"
    AddST "C|D77|n|4.0456052697018166|0.00|0|0|10.0|||right|0|0"
    AddST "C|E77|n|2935342|#,##0|0|0|10.0||||0|0"
    AddST "C|F77|n|2933871|#,##0|0|0|10.0||||0|1"
    AddST "C|H77|n|2024|General|0|0|10.0|||left|0|1"
    AddST "C|I77|n|96.4342|0.000|0|0|10.0|||right|0|0"
    AddST "C|J77|n|3.930947074368342|0.00|0|0|10.0|||right|0|0"
    AddST "C|K77|n|2891444|#,##0|0|0|10.0|||right|0|0"
    AddST "C|B78|s|2025-26 (1), (2)|General|0|0|10.0||||0|1"
    AddST "C|C78|s|-|0.000|0|0|10.0|||right|0|0"
    AddST "C|D78|n|3.238831156796862|0.00|0|0|10.0|||right|0|0"
    AddST "C|E78|n|3054666|#,##0|0|0|10.0||||0|0"
    AddST "C|F78|n|3062565|#,##0|0|0|10.0||||0|1"
    AddST "C|H78|n|2025|General|0|0|10.0|||left|0|1"
    AddST "C|I78|n|100|0.000|0|0|10.0|||right|0|0"
    AddST "C|J78|n|3.6976508334180154|0.00|0|0|10.0|||right|0|0"
    AddST "C|K78|n|3039861|#,##0|0|0|10.0|||right|0|0"
    AddST "C|B79|s|2026-27 (1), (2)|General|0|0|10.0||||0|1"
    AddST "C|C79|s|-|0.000|0|0|10.0|||right|0|0"
    AddST "C|D79|n|2.0057419249593256|0.00|0|0|10.0|||right|0|0"
    AddST "C|E79|n|3164074|#,##0|0|0|10.0||||0|0"
    AddST "C|F79|n|3164159|#,##0|0|0|10.0||||0|1"
    AddST "C|H79|s|2026 (1), (2)|General|0|0|10.0|||left|0|1"
    AddST "C|I79|s|-|0.000|0|0|10.0|||right|0|0"
    AddST "C|J79|n|2.183088135060407|0.00|0|0|10.0|||right|0|0"
    AddST "C|K79|n|3136973|#,##0|0|0|10.0|||right|0|0"
    AddST "C|B80|s|2027-28 (1), (2)|General|0|0|10.0||||0|1"
    AddST "C|C80|s|-|0.000|0|0|10.0|||right|0|0"
    AddST "C|D80|n|1.9419802624865978|0.00|0|0|10.0|||right|0|0"
    AddST "C|E80|n|3277008|#,##0|0|0|10.0||||0|0"
    AddST "C|F80|n|3277097|#,##0|0|0|10.0||||0|1"
    AddST "C|H80|s|2027 (1), (2)|General|0|0|10.0|||left|0|1"
    AddST "C|I80|s|-|0.000|0|0|10.0|||right|0|0"
    AddST "C|J80|n|1.944412716488264|0.00|0|0|10.0|||right|0|0"
    AddST "C|K80|n|3248777|#,##0|0|0|10.0|||right|0|0"
    AddST "C|B81|s|2028-29 (1), (2)|General|0|0|10.0||||0|1"
    AddST "C|C81|s|-|0.000|0|0|10.0|||right|0|0"
    AddST "C|D81|n|1.8653157522025676|0.00|0|0|10.0|||right|0|0"
    AddST "C|E81|n|3390825|#,##0|0|0|10.0||||0|0"
    AddST "C|F81|n|3390914|#,##0|0|0|10.0||||0|1"
    AddST "C|H81|s|2028 (1), (2)|General|0|0|10.0|||left|0|1"
    AddST "C|I81|s|-|0.000|0|0|10.0|||right|0|0"
    AddST "C|J81|n|1.8875382865722878|0.00|0|0|10.0|||right|0|0"
    AddST "C|K81|n|3362463|#,##0|0|0|10.0|||right|0|0"
    AddST "C|B82|s|2029-30 (1), (2)|General|0|0|10.0||||0|1"
    AddST "C|C82|s|-|0.000|0|0|10.0|||right|0|0"
    AddST "C|D82|n|1.864367236103881|0.00|0|0|10.0|||right|0|0"
    AddST "C|E82|n|3506679|#,##0|0|0|10.0||||0|0"
    AddST "C|F82|n|3506771|#,##0|0|0|10.0||||0|0"
    AddST "C|H82|s|2029 (1), (2)|General|0|0|10.0|||left|0|1"
    AddST "C|I82|s|-|0.000|0|0|10.0|||right|0|0"
    AddST "C|J82|n|1.8438589634765412|0.00|0|0|10.0|||right|0|0"
    AddST "C|K82|n|3477144|#,##0|0|0|10.0|||right|0|0"
    AddST "C|B83|s|2030-31 (1), (2)|General|0|0|10.0||||0|1"
    AddST "C|C83|s|-|0.000|0|0|10.0|||right|0|0"
    AddST "C|D83|n|2.0159623211262856|0.00|0|0|10.0|||right|0|0"
    AddST "C|E83|n|3631831|#,##0|0|0|10.0||||0|0"
    AddST "C|F83|n|3631932|#,##0|0|0|10.0||||0|0"
    AddST "C|H83|s|2030 (1), (2)|General|0|0|10.0|||left|0|1"
    AddST "C|I83|s|-|0.000|0|0|10.0|||right|0|0"
    AddST "C|J83|n|1.9799878099249435|0.00|0|0|10.0|||right|0|0"
    AddST "C|K83|n|3599758|#,##0|0|0|10.0|||right|0|0"
    AddST "C|B84|s|Sources and footnotes:|General|1|0|10.0||||0|0"
    AddST "C|B85|s|GDP Deflator:|General|0|0|10.0||||0|0"
    AddST "C|C85|s|Financial years 1955-56 to 2024-25 taken from ONS series L8GG in data tables: Table N.|General|0|0|10.0|||left|1|0"
    AddST "C|C86|s|https://www.ons.gov.uk/file?uri=/economy/grossdomesticproductgdp/datasets/uksecondestimateofgdpdatatables/quarter4octtodec2025quarterlynationalaccounts/quarterlynationalaccountsdatatables.xlsx|General|0|0|10.0|||left|1|0"
    AddST "C|C87|s|Calendar years 1955 to 2025 taken from ONS series MNF2 in data tables: Table O.|General|0|0|10.0|||left|1|0"
    AddST "C|C88|s|https://www.ons.gov.uk/file?uri=/economy/grossdomesticproductgdp/datasets/uksecondestimateofgdpdatatables/quarter4octtodec2025quarterlynationalaccounts/quarterlynationalaccountsdatatables.xlsx|General|0|0|10.0|||left|1|0"
    AddST "C|C89|s|For years 2025-26 to 2030-31 (2026 to 2030): taken from the Office for Budget Responsibility (OBR) GDP deflator forecasts as of March 2026 EFO detailed forecast tables.|General|0|0|10.0|||left|1|0"
    AddST "C|C90|s|https://obr.uk/efo/economic-and-fiscal-outlook-march-2026/|General|0|0|10.0|||left|1|0"
    AddST "C|B91|s|Money GDP:|General|0|0|10.0||||0|0"
    AddST "C|C91|s|For years 1955-56 to 2024-25 (1955 to 2025): ONS data for money GDP not seasonally adjusted series BKTL in data tables: Table N.|General|0|0|10.0|||left|1|0"
    AddST "C|C92|s|https://www.ons.gov.uk/file?uri=/economy/grossdomesticproductgdp/datasets/uksecondestimateofgdpdatatables/quarter4octtodec2025quarterlynationalaccounts/quarterlynationalaccountsdatatables.xlsx|General|0|0|10.0|||left|1|0"
    AddST "C|C93|s|For years 1955-56 to 2024-25: ONS data for money GDP seasonally adjusted series YBHA in data tables: Table N.|General|0|0|10.0|||left|1|0"
    AddST "C|C94|s|https://www.ons.gov.uk/file?uri=/economy/grossdomesticproductgdp/datasets/uksecondestimateofgdpdatatables/quarter4octtodec2025quarterlynationalaccounts/quarterlynationalaccountsdatatables.xlsx|General|0|0|10.0|||left|1|0"
    AddST "C|C95|s|For years 2025-26 to 2030-31 (2026 to 2030): taken from the Office for Budget Responsibility (OBR) GDP deflator forecasts as of March 2026 EFO detailed forecast tables.|General|0|0|10.0|||left|1|0"
    AddST "C|C96|s|https://obr.uk/efo/economic-and-fiscal-outlook-march-2026/|General|0|0|10.0|||left|1|0"
    AddST "C|B97|s|Footnotes:|General|0|0|10.0||||0|0"
    AddST "C|B98|s|(1)|General|0|0|10.0|||right|0|0"
    AddST "C|C98|s|For years 2025-26 to 2030-31 (2026 to 2030), this presentation only shows percentage changes in line with data taken from the OBR GDP deflator forecasts as of March 2026 EFO detailed forecast tables.|@|0|0|10.0|||left|1|0"
    AddST "C|B99|s|(2)|General|0|0|10.0|||right|0|0"
    AddST "C|C99|s|For years 2025-26 to 2030-31 (2026 to 2030), money GDP forecasts from the OBR as of the Spring Statement, March 2026.|General|0|0|10.0|||left|1|0"
    AddST "C|B100|s|(3)|General|0|0|10.0|||right|0|0"
    AddST "C|C100|s|Non-Seasonally adjusted money GDP (BKTL) from 1955-56 to 2024-25 (1955 to 2025) consistent with ONS Quarterly National Accounts release of 31 March 2026.|General|0|0|10.0|||left|1|0"
    AddST "C|B101|s|(4)|General|0|0|10.0|||right|0|0"
    AddST "C|C101|s|Seasonally adjusted money GDP (YBHA) from 1955-56 to 2024-25 consistent with ONS Quarterly National Accounts release of 31 March 2026.|General|0|0|10.0|||left|1|0"
    AddST "C|B102|s|(5)|General|0|0|10.0|||right|0|0"
    AddST "C|C102|s|For further information and the 'User's Guide' to these series, please visit the following page on the GOV.UK website at: |General|0|0|10.0|||left|1|0"
    AddST "C|C103|s|https://www.gov.uk/government/publications/gross-domestic-product-gdp-deflators-user-guide|General|0|0|10.0|||left|1|0"
    AddST "C|B104|s|(6)|General|0|0|10.0|||right|0|0"
    AddST "C|C104|s|For practical examples of how to use the GDP deflator series, please visit the following page on the GOV.UK website at: |General|0|0|10.0|||left|1|0"
    AddST "C|C105|s|https://www.gov.uk/government/publications/how-to-use-the-gdp-deflator-series-practical-examples|General|0|0|10.0|||left|1|0"
    AddST "W|A|2.83203125"
    AddST "W|B|14.83203125"
    AddST "W|C|16.6640625"
    AddST "W|D|16.6640625"
    AddST "W|E|15.33203125"
    AddST "W|G|2.83203125"
    AddST "W|H|12.6640625"
    AddST "W|I|16.6640625"
    AddST "W|J|16.6640625"
    AddST "W|K|18.6640625"
    AddST "W|L|11.83203125"
    AddST "M|C85:K85"
    AddST "M|C100:K100"
    AddST "M|C94:K94"
    AddST "M|C6:D6"
    AddST "M|I6:J6"
End Sub

Private Sub DataST4_6()
    Dim s As String
    AddST "M|C103:K103"
    AddST "M|C93:K93"
    AddST "M|C99:K99"
    AddST "M|C105:K105"
    AddST "M|H5:K5"
    AddST "M|C90:K90"
    AddST "M|C95:K95"
    AddST "M|B4:K4"
    AddST "M|C91:K91"
    AddST "M|B3:K3"
    AddST "M|C96:K96"
    AddST "M|C102:K102"
    AddST "M|C87:K87"
    AddST "M|C101:K101"
    AddST "M|C86:K86"
    AddST "M|C92:K92"
    AddST "M|C104:K104"
    AddST "M|C98:K98"
    AddST "M|C89:K89"
    AddST "M|B5:F5"
    AddST "M|C88:K88"
End Sub
