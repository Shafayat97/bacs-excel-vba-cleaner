' ============================================================
'  Sheet1_AutoClean
'  Paste this into the SHEET module (Sheet1), not a normal module.
'  Cleans columns A-E automatically as data is typed or pasted.
' ============================================================

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo ExitSub
    Application.EnableEvents = False
    
    Dim cell As Range
    Dim v As String, cleaned As String
    Dim i As Integer, ch As String
    Dim specialChars As String
    specialChars = "'-]?." & Chr(34) & ".,+*/-@$#=&:;_()"
    
    For Each cell In Target
        If cell.row < 1 Then GoTo NextCell
        If Trim(cell.Value) = "" Then GoTo NextCell
        v = CStr(cell.Value)
        
        Select Case cell.Column
            Case 1 ' Col A: exactly 6 digits
                cleaned = ""
                For i = 1 To Len(v)
                    ch = Mid(v, i, 1)
                    If ch >= "0" And ch <= "9" Then cleaned = cleaned & ch
                Next i
                If Len(cleaned) > 6 Then cleaned = Right(cleaned, 6)
                Do While Len(cleaned) < 6: cleaned = "0" & cleaned: Loop
                cell.NumberFormat = "@"
                cell.Value = cleaned
                
            Case 2 ' Col B: max 18 chars, strip special chars, no double spaces
                cleaned = ""
                For i = 1 To Len(v)
                    ch = Mid(v, i, 1)
                    If InStr(specialChars, ch) = 0 Then cleaned = cleaned & ch
                Next i
                Do While InStr(cleaned, "  ") > 0
                    cleaned = Join(Split(cleaned, "  "), " ")
                Loop
                cleaned = Trim(cleaned)
                If Len(cleaned) > 18 Then cleaned = Left(cleaned, 18)
                cell.Value = cleaned
                
            Case 3 ' Col C: exactly 8 digits
                cleaned = ""
                For i = 1 To Len(v)
                    ch = Mid(v, i, 1)
                    If ch >= "0" And ch <= "9" Then cleaned = cleaned & ch
                Next i
                If Len(cleaned) > 8 Then cleaned = Right(cleaned, 8)
                Do While Len(cleaned) < 8: cleaned = "0" & cleaned: Loop
                cell.NumberFormat = "@"
                cell.Value = cleaned
                
            Case 4 ' Col D: number with 2 decimal places
                On Error Resume Next
                Dim numVal As Double
                numVal = CDbl(cell.Value)
                If Err.Number = 0 Then
                    cell.Value = numVal
                    cell.NumberFormat = "0.00"
                End If
                On Error GoTo ExitSub
                
            Case 5 ' Col E: max 18 chars, strip special chars, no double spaces
                cleaned = ""
                For i = 1 To Len(v)
                    ch = Mid(v, i, 1)
                    If InStr(specialChars, ch) = 0 Then cleaned = cleaned & ch
                Next i
                Do While InStr(cleaned, "  ") > 0
                    cleaned = Join(Split(cleaned, "  "), " ")
                Loop
                cleaned = Trim(cleaned)
                If Len(cleaned) > 18 Then cleaned = Left(cleaned, 18)
                cell.Value = cleaned
                
        End Select
        
        cell.HorizontalAlignment = xlLeft
        cell.VerticalAlignment = xlTop
        
NextCell:
    Next cell
    
ExitSub:
    Application.EnableEvents = True
End Sub
