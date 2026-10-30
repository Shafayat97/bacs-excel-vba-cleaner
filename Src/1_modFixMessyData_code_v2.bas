Option Explicit

' ============================================================
'  modFixMessyData  (v2 - with exception checking)
'  Column F = Status (OK / CHECK)
'  Column G = Reason the row was flagged
'  Column H = Type Y to approve a flagged row
'  Columns J-N (hidden) = per-column notes used to build column G
' ============================================================

Private Const COL_STATUS As Long = 6
Private Const COL_REASON As Long = 7
Private Const COL_APPROVE As Long = 8
Private Const ISSUE_OFFSET As Long = 9   ' A->J, B->K, C->L, D->M, E->N

Public Sub FixMessyData()
    Dim ws As Worksheet
    Dim lastRow As Long, r As Long, col As Long, flagged As Long

    Set ws = ActiveSheet
    lastRow = LastDataRow(ws)

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    On Error GoTo Done

    For r = 1 To lastRow
        For col = 1 To 5
            CleanCell ws.Cells(r, col), False
        Next col
        RefreshRow ws, r
        If ws.Cells(r, COL_STATUS).Value = "CHECK" Then flagged = flagged + 1
    Next r

    ws.Columns("A:E").AutoFit
    SetupColumns ws

Done:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    If Err.Number <> 0 Then
        MsgBox "Error on row " & r & ": " & Err.Description, vbExclamation
    Else
        MsgBox "Done! " & lastRow & " rows cleaned." & vbCrLf & _
               flagged & " row(s) need review (see columns F and G).", vbInformation
    End If
End Sub

Public Sub CleanCell(ByVal cell As Range, Optional ByVal overwriteIssue As Boolean = True)
    Dim v As String, d As String, cleaned As String, fullText As String
    Dim issue As String, label As String
    Dim amt As Double, rounded As Double

    If cell.Column > 5 Then Exit Sub

    If IsError(cell.Value) Then
        issue = "Cell contains an Excel error"
        GoTo SaveIssue
    End If

    v = Trim(CStr(cell.Value))
    If v = "" Then GoTo SaveIssue

    Select Case cell.Column
        Case 1   ' Sort code: 6 digits
            d = DigitsOnly(v)
            If Len(d) = 0 Then
                issue = "Sort code '" & v & "' has no digits"
            Else
                If Len(d) < 5 Then
                    issue = "Sort code '" & v & "' had only " & Len(d) & " digits, zeros added at front"
                ElseIf Len(d) > 6 Then
                    issue = "Sort code '" & v & "' had " & Len(d) & " digits, cut to last 6"
                End If
                cell.NumberFormat = "@"
                cell.Value = PadDigits(v, 6)
            End If

        Case 2, 5   ' Name / Reference: max 18 chars
            If cell.Column = 2 Then label = "Name" Else label = "Reference"
            fullText = CleanText(v, 32767)
            If fullText = "" Then
                issue = label & " '" & v & "' is empty after removing special characters"
            ElseIf cell.Column = 5 And Len(fullText) > 18 Then
                issue = "Reference cut to 18 characters (full: '" & fullText & "')"
            End If
            cell.NumberFormat = "@"
            cell.Value = CleanText(v, 18)

        Case 3   ' Account number: 8 digits
            d = DigitsOnly(v)
            If Len(d) = 0 Then
                issue = "Account number '" & v & "' has no digits"
            Else
                If Len(d) < 7 Then
                    issue = "Account number '" & v & "' had only " & Len(d) & " digits, zeros added at front"
                ElseIf Len(d) > 8 Then
                    issue = "Account number '" & v & "' had " & Len(d) & " digits, cut to last 8"
                End If
                cell.NumberFormat = "@"
                cell.Value = PadDigits(v, 8)
            End If

        Case 4   ' Amount: number, 2 decimals
            cleaned = Replace(v, ChrW(163), "")   ' remove pound sign
            cleaned = Replace(cleaned, ",", "")
            cleaned = Replace(cleaned, " ", "")
            If IsNumeric(cleaned) Then
                amt = Val(cleaned)
                rounded = Application.WorksheetFunction.Round(amt, 2)
                If Abs(rounded - amt) > 0.0000001 Then
                    issue = "Amount '" & v & "' rounded to " & Format(rounded, "0.00")
                End If
                If rounded <= 0 Then issue = AddText(issue, "Amount is zero or negative")
                cell.NumberFormat = "0.00"
                cell.Value = rounded
                cell.Interior.Pattern = xlNone
            Else
                issue = "Amount '" & v & "' is not a number"
                cell.Interior.Color = RGB(255, 199, 206)   ' pink
            End If
    End Select

    cell.HorizontalAlignment = xlLeft
    cell.VerticalAlignment = xlTop

SaveIssue:
    If overwriteIssue Or issue <> "" Then
        cell.Worksheet.Cells(cell.Row, cell.Column + ISSUE_OFFSET).Value = issue
    End If
End Sub

Public Sub RefreshRow(ByVal ws As Worksheet, ByVal r As Long)
    Dim c As Long, msg As String, note As String, hasData As Boolean
    Dim labels As Variant
    labels = Array("Sort code", "Name", "Account number", "Amount", "Reference")

    For c = 1 To 5
        If Len(ws.Cells(r, c).Text) > 0 Then hasData = True
    Next c

    If Not hasData Then
        ws.Range(ws.Cells(r, COL_STATUS), ws.Cells(r, COL_APPROVE)).ClearContents
        ws.Range(ws.Cells(r, COL_STATUS), ws.Cells(r, COL_REASON)).Interior.Pattern = xlNone
        ws.Range(ws.Cells(r, 1 + ISSUE_OFFSET), ws.Cells(r, 5 + ISSUE_OFFSET)).ClearContents
        Exit Sub
    End If

    For c = 1 To 5
        note = CStr(ws.Cells(r, c + ISSUE_OFFSET).Value)
        If note <> "" Then
            msg = AddText(msg, note)
        ElseIf Len(ws.Cells(r, c).Text) = 0 Then
            msg = AddText(msg, labels(c - 1) & " is missing")
        End If
    Next c

    If msg = "" Then
        ws.Cells(r, COL_STATUS).Value = "OK"
        ws.Cells(r, COL_STATUS).Interior.Color = RGB(198, 239, 206)   ' green
        ws.Cells(r, COL_REASON).ClearContents
        ws.Cells(r, COL_REASON).Interior.Pattern = xlNone
    Else
        ws.Cells(r, COL_STATUS).Value = "CHECK"
        ws.Cells(r, COL_STATUS).Interior.Color = RGB(255, 235, 156)   ' yellow
        ws.Cells(r, COL_REASON).Value = msg
        ws.Cells(r, COL_REASON).Interior.Color = RGB(255, 235, 156)
    End If
End Sub

Public Function ReadyToExport(ByVal ws As Worksheet) As Boolean
    Dim r As Long, c As Long, lastRow As Long
    Dim hasData As Boolean, needReview As Long, rowList As String
    Dim status As String, approved As Boolean

    lastRow = LastDataRow(ws)
    For r = 1 To lastRow
        hasData = False
        For c = 1 To 5
            If Len(ws.Cells(r, c).Text) > 0 Then hasData = True
        Next c

        If hasData Then
            status = CStr(ws.Cells(r, COL_STATUS).Value)
            approved = (UCase(Trim(CStr(ws.Cells(r, COL_APPROVE).Value))) = "Y")
            If Not (status = "OK" Or (status = "CHECK" And approved)) Then
                needReview = needReview + 1
                If needReview <= 15 Then rowList = rowList & r & " "
            End If
        End If
    Next r

    If needReview > 0 Then
        MsgBox needReview & " row(s) need review before export." & vbCrLf & _
               "Rows: " & rowList & vbCrLf & vbCrLf & _
               "For each row: fix the value, or type Y in column H to approve it." & vbCrLf & _
               "(If column F is empty, run FixMessyData first.)" & vbCrLf & vbCrLf & _
               "Export cancelled.", vbExclamation, "Review needed"
        ReadyToExport = False
    Else
        ReadyToExport = True
    End If
End Function

Public Sub SetupColumns(ByVal ws As Worksheet)
    ws.Range(ws.Cells(1, 1 + ISSUE_OFFSET), ws.Cells(1, 5 + ISSUE_OFFSET)).EntireColumn.Hidden = True
    ws.Columns(COL_STATUS).ColumnWidth = 9
    ws.Columns(COL_REASON).ColumnWidth = 60
    ws.Columns(COL_REASON).WrapText = True
    ws.Columns(COL_APPROVE).ColumnWidth = 9
End Sub

Public Sub ResetEvents()
    Application.EnableEvents = True
    MsgBox "Auto-clean is switched back on.", vbInformation
End Sub

Private Function LastDataRow(ByVal ws As Worksheet) As Long
    Dim col As Long, n As Long
    For col = 1 To 5
        n = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
        If n > LastDataRow Then LastDataRow = n
    Next col
End Function

Private Function AddText(ByVal a As String, ByVal b As String) As String
    If a = "" Then AddText = b Else AddText = a & "; " & b
End Function

Private Function DigitsOnly(ByVal s As String) As String
    Dim i As Long, ch As String, out As String
    For i = 1 To Len(s)
        ch = Mid(s, i, 1)
        If ch Like "#" Then out = out & ch
    Next i
    DigitsOnly = out
End Function

' Keep digits only, then pad with leading zeros / keep rightmost digits
Private Function PadDigits(ByVal s As String, ByVal size As Integer) As String
    Dim out As String
    out = DigitsOnly(s)
    If Len(out) > size Then out = Right(out, size)
    Do While Len(out) < size
        out = "0" & out
    Loop
    PadDigits = out
End Function

' Remove special characters, double spaces, and cut to max length
Private Function CleanText(ByVal s As String, ByVal maxLen As Long) As String
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
