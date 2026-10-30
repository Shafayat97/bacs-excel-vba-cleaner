Option Explicit

' ============================================================
'  Sheet1 code (v2 - with exception checking)
'  Paste this into the SHEET module (Sheet1), not a normal module.
' ============================================================

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim rng As Range, cell As Range, area As Range
    Dim r As Long

    Set rng = Intersect(Target, Me.Range("A:E"), Me.UsedRange)
    If rng Is Nothing Then Exit Sub

    On Error GoTo ExitSub
    Application.EnableEvents = False

    For Each cell In rng.Cells
        CleanCell cell, True
    Next cell

    For Each area In rng.Areas
        For r = area.Row To area.Row + area.Rows.Count - 1
            Me.Cells(r, 8).ClearContents   ' data changed: approval needed again
            RefreshRow Me, r
        Next r
    Next area

    SetupColumns Me

ExitSub:
    Application.EnableEvents = True
End Sub

' ============================================================

Sub SaveAsCSV()
    Dim csvPath As String
    Dim oWS As Worksheet
    Dim fileNum As Integer
    Dim r As Long, c As Long
    Dim cellVal As String, lineStr As String

    Set oWS = ThisWorkbook.ActiveSheet
    If Not ReadyToExport(oWS) Then Exit Sub

    csvPath = ThisWorkbook.Path & "\" & "output.csv"
    fileNum = FreeFile
    Open csvPath For Output As #fileNum

    For r = 1 To oWS.UsedRange.Rows.Count
        lineStr = ""
        For c = 1 To 5
            cellVal = Trim(CStr(oWS.Cells(r, c).Value))
            If c = 4 And cellVal <> "" And IsNumeric(cellVal) Then
                cellVal = Format(CDbl(cellVal), "0.00")
            End If
            If c < 5 Then
                lineStr = lineStr & cellVal & ","
            Else
                lineStr = lineStr & cellVal
            End If
        Next c

        If Trim(Replace(lineStr, ",", "")) <> "" Then
            Print #fileNum, lineStr
        End If
    Next r

    Close #fileNum
    MsgBox "CSV saved at: " & csvPath
End Sub

' ============================================================

Sub SaveAsNotepad()
    Dim csvPath As String, txtPath As String
    Dim fileIn As Integer, fileOut As Integer
    Dim lineStr As String

    csvPath = ThisWorkbook.Path & "\" & "output.csv"
    txtPath = ThisWorkbook.Path & "\" & "output.txt"

    fileIn = FreeFile
    Open csvPath For Input As #fileIn
    fileOut = FreeFile
    Open txtPath For Output As #fileOut

    Do While Not EOF(fileIn)
        Line Input #fileIn, lineStr
        Do While Right(Trim(lineStr), 1) = ","
            lineStr = Left(Trim(lineStr), Len(Trim(lineStr)) - 1)
        Loop
        Do While Left(Trim(lineStr), 1) = ","
            lineStr = Mid(Trim(lineStr), 2)
        Loop
        Do While InStr(lineStr, ",,") > 0
            lineStr = Join(Split(lineStr, ",,"), ",")
        Loop
        Print #fileOut, lineStr
    Loop

    Close #fileIn
    Close #fileOut
    MsgBox "Notepad file saved at: " & txtPath
End Sub

' ============================================================

Sub ExportAll()
    If Not ReadyToExport(ThisWorkbook.ActiveSheet) Then Exit Sub
    Call SaveAsCSV
    Call SaveAsNotepad
    MsgBox "Done! Both CSV and TXT files have been saved."
End Sub
