Option Explicit

' ============================================================
'  Sheet1 code
'  Paste this into the SHEET module (Sheet1), not a normal module.
'  Auto-cleans columns A-E when data is typed or pasted,
'  using the same rules as FixMessyData (CleanCell).
' ============================================================

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim rng As Range, cell As Range

    Set rng = Intersect(Target, Me.Range("A:E"), Me.UsedRange)
    If rng Is Nothing Then Exit Sub

    On Error GoTo ExitSub
    Application.EnableEvents = False

    For Each cell In rng.Cells
        CleanCell cell
    Next cell

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

    csvPath = ThisWorkbook.Path & "\" & "output.csv"
    Set oWS = ThisWorkbook.ActiveSheet

    fileNum = FreeFile
    Open csvPath For Output As #fileNum

    For r = 1 To oWS.UsedRange.Rows.Count
        lineStr = ""
        For c = 1 To 5
            cellVal = Trim(CStr(oWS.Cells(r, c).Value))
            ' Amount always exported with 2 decimals (1250.5 -> 1250.50)
            If c = 4 And cellVal <> "" And IsNumeric(cellVal) Then
                cellVal = Format(CDbl(cellVal), "0.00")
            End If
            If c < 5 Then
                lineStr = lineStr & cellVal & ","
            Else
                lineStr = lineStr & cellVal
            End If
        Next c

        ' Skip completely empty rows
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

        ' Remove trailing commas
        Do While Right(Trim(lineStr), 1) = ","
            lineStr = Left(Trim(lineStr), Len(Trim(lineStr)) - 1)
        Loop

        ' Remove leading commas
        Do While Left(Trim(lineStr), 1) = ","
            lineStr = Mid(Trim(lineStr), 2)
        Loop

        ' Remove double commas
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
    Call SaveAsCSV
    Call SaveAsNotepad
    MsgBox "Done! Both CSV and TXT files have been saved."
End Sub
