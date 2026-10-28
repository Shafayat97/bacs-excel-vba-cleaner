Attribute VB_Name = "modFixMessyData"
' ============================================================
'  FixMessyData
'  Cleans data that is ALREADY in the sheet (columns A-E),
'  using the same rules as the Worksheet_Change code:
'    A = 6 digits, B = max 18 chars, C = 8 digits,
'    D = number with 2 decimals, E = max 18 chars
' ============================================================
Option Explicit

Sub FixMessyData()
    Dim ws As Worksheet
    Dim lastRow As Long, r As Long, col As Long, n As Long
    Dim cell As Range
    Dim v As String, cleaned As String

    Set ws = ActiveSheet

    ' Find the last row used in columns A-E
    For col = 1 To 5
        n = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
        If n > lastRow Then lastRow = n
    Next col

    Application.EnableEvents = False      ' stop Worksheet_Change running twice
    Application.ScreenUpdating = False
    On Error GoTo Done

    For r = 1 To lastRow
        For col = 1 To 5
            Set cell = ws.Cells(r, col)
            If Not IsError(cell.Value) Then
                v = Trim(CStr(cell.Value))
                If v <> "" Then
                    Select Case col
                        Case 1   ' Sort code: 6 digits
                            cleaned = PadDigits(v, 6)
                            cell.NumberFormat = "@"
                            cell.Value = cleaned

                        Case 2, 5   ' Name / Reference: max 18 chars
                            cell.NumberFormat = "@"
                            cell.Value = CleanText(v, 18)

                        Case 3   ' Account number: 8 digits
                            cleaned = PadDigits(v, 8)
                            cell.NumberFormat = "@"
                            cell.Value = cleaned

                        Case 4   ' Amount: number, 2 decimals
                            cleaned = Replace(v, ChrW(163), "")   ' remove pound sign
                            cleaned = Replace(cleaned, ",", "")
                            cleaned = Replace(cleaned, " ", "")
                            If IsNumeric(cleaned) Then
                                cell.Value = Application.WorksheetFunction.Round(Val(cleaned), 2)
                                cell.NumberFormat = "0.00"
                            Else
                                cell.Interior.Color = RGB(255, 199, 206)   ' highlight: not a number
                            End If
                    End Select

                    cell.HorizontalAlignment = xlLeft
                    cell.VerticalAlignment = xlTop
                End If
            End If
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
