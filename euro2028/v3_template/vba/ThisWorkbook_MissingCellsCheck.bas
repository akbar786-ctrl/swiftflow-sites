' Paste into the ThisWorkbook module of the collection template (save as .xlsm).
' Stops Save / Close while any hours or officers cell is showing the missing-cell red.
Option Explicit

' Set BLOCK_SAVE to False to warn but still allow an unfinished draft to be saved.
Private Const BLOCK_SAVE As Boolean = True
Private Const BLOCK_CLOSE As Boolean = True
Private Const RED_FILL As Long = 13551615   ' RGB(255,199,206) = FFC7CE

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    Dim firstCell As Range, n As Long
    n = CountMissing(firstCell)
    If n = 0 Then Exit Sub
    If BLOCK_SAVE Then
        MsgBox n & " required cell(s) are missing and shown in red." & vbCrLf & vbCrLf & _
               "Please complete them before saving. You will be taken to the first one.", _
               vbExclamation, "Missing information"
        Cancel = True
    ElseIf MsgBox(n & " required cell(s) are missing and shown in red." & vbCrLf & vbCrLf & _
                  "Save as an unfinished draft anyway?", vbYesNo + vbExclamation, _
                  "Missing information") = vbYes Then
        Exit Sub
    Else
        Cancel = True
    End If
    GoToCell firstCell
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Dim firstCell As Range, n As Long
    If Not BLOCK_CLOSE Then Exit Sub
    n = CountMissing(firstCell)
    If n = 0 Then Exit Sub
    MsgBox n & " required cell(s) are missing and shown in red." & vbCrLf & vbCrLf & _
           "Please complete them before closing. You will be taken to the first one.", _
           vbExclamation, "Missing information"
    Cancel = True
    GoToCell firstCell
End Sub

Private Function CountMissing(ByRef firstCell As Range) As Long
    CountMissing = CheckSheet("Match Day Resource", "H8:BC371", firstCell) _
                 + CheckSheet("Wider Tournament Resource", "G6:R102", firstCell) _
                 + CheckSheet("Pan London", "F6:Q46", firstCell)
End Function

Private Function CheckSheet(ByVal sheetName As String, ByVal addr As String, _
                            ByRef firstCell As Range) As Long
    Dim ws As Worksheet, rng As Range, grp As Range, c As Range
    Dim r As Long, k As Long, n As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function          ' e.g. no Pan London tab in the General version
    Set rng = ws.Range(addr)
    For r = 1 To rng.Rows.Count
        For k = 1 To rng.Columns.Count Step 3     ' each day / risk block: Officers, Opp Cost, Chargeable
            Set grp = rng.Cells(r, k).Resize(1, 3)
            If Application.WorksheetFunction.CountA(grp) > 0 Then
                For Each c In grp.Cells
                    If IsEmpty(c.Value) Then
                        If c.DisplayFormat.Interior.Color = RED_FILL Then
                            n = n + 1
                            If firstCell Is Nothing Then Set firstCell = c
                        End If
                    End If
                Next c
            End If
        Next k
    Next r
    CheckSheet = n
End Function

Private Sub GoToCell(ByVal c As Range)
    If c Is Nothing Then Exit Sub
    c.Worksheet.Activate
    Application.Goto c, True
End Sub
