Option Explicit

' ============================================================
'  modFixMessyData
'  All cleaning rules live here, in ONE place.
'  - CleanCell     : cleans one cell (used by paste AND by FixMessyData)
'  - FixMessyData  : cleans everything already in the sheet
'  - ResetEvents   : switches auto-clean back on if it stopped working
' ============================================================

Public Sub FixMessyData()
    Dim ws As Worksheet
    Dim lastRow As Long, r As Long, col As Long, n As Long

    Set ws = ActiveSheet

    For col = 1 To 5
        n = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
        If n > lastRow Then lastRow = n
    Next col

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    On Error GoTo Done

    For r = 1 To lastRow
        For col = 1 To 5
            CleanCell ws.Cells(r, col)
        Next col
    Next r

    ws.Columns("A:E").AutoFit

Done:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    If Err.Number <> 0 Then
        MsgBox "Error on row " & r & ": " & Err.Description, vbExclamation
    Else
        MsgBox "Done! " & lastRow & " rows cleaned.", vbInformation
    End If
End Sub

Public Sub CleanCell(ByVal cell As Range)
    Dim v As String, cleaned As String

    If cell.Column > 5 Then Exit Sub
    If IsError(cell.Value) Then Exit Sub
    v = Trim(CStr(cell.Value))
    If v = "" Then Exit Sub

    Select Case cell.Column
        Case 1   ' Sort code: 6 digits
            cell.NumberFormat = "@"
            cell.Value = PadDigits(v, 6)

        Case 2, 5   ' Name / Reference: max 18 chars
            cell.NumberFormat = "@"
            cell.Value = CleanText(v, 18)

        Case 3   ' Account number: 8 digits
            cell.NumberFormat = "@"
            cell.Value = PadDigits(v, 8)

        Case 4   ' Amount: number, 2 decimals
            cleaned = Replace(v, ChrW(163), "")   ' remove pound sign
            cleaned = Replace(cleaned, ",", "")
            cleaned = Replace(cleaned, " ", "")
            If IsNumeric(cleaned) Then
                cell.NumberFormat = "0.00"
                cell.Value = Application.WorksheetFunction.Round(Val(cleaned), 2)
                cell.Interior.Pattern = xlNone
            Else
                cell.Interior.Color = RGB(255, 199, 206)   ' pink: not a number
            End If
    End Select

    cell.HorizontalAlignment = xlLeft
    cell.VerticalAlignment = xlTop
End Sub

Public Sub ResetEvents()
    Application.EnableEvents = True
    MsgBox "Auto-clean is switched back on.", vbInformation
End Sub

' Keep digits only, then pad with leading zeros / keep rightmost digits
Private Function PadDigits(ByVal s As String, ByVal size As Integer) As String
    Dim i As Long, ch As String, out As String
    For i = 1 To Len(s)
        ch = Mid(s, i, 1)
        If ch Like "#" Then out = out & ch
    Next i
    If Len(out) > size Then out = Right(out, size)
    Do While Len(out) < size
        out = "0" & out
    Loop
    PadDigits = out
End Function

' Remove special characters, double spaces, and cut to max length
Private Function CleanText(ByVal s As String, ByVal maxLen As Integer) As String
    Dim specialChars As String, i As Long, ch As String, out As String
    specialChars = "'-]?." & Chr(34) & ",+*/@$#=&:;_()[!%^{}<>|\~`"
    For i = 1 To Len(s)
        ch = Mid(s, i, 1)
        If InStr(specialChars, ch) = 0 Then out = out & ch
    Next i
    Do While InStr(out, "  ") > 0
        out = Replace(out, "  ", " ")
    Loop
    out = Trim(out)
    If Len(out) > maxLen Then out = Trim(Left(out, maxLen))
    CleanText = out
End Function
