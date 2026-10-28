Attribute VB_Name = "Export_Macros"
' ============================================================
'  Export_Macros
'  Import into a normal module (Alt + F11 > File > Import File).
'  SaveAsCSV     -> output.csv
'  SaveAsNotepad -> output.txt (made from output.csv)
'  ExportAll     -> both
' ============================================================

Sub SaveAsCSV()
    Dim csvPath As String
    csvPath = ThisWorkbook.Path & "\" & "output.csv"
    
    Dim oWS As Worksheet
    Set oWS = ThisWorkbook.ActiveSheet
    
    Dim fileNum As Integer
    fileNum = FreeFile
    Open csvPath For Output As #fileNum
    
    Dim row As Long, col As Long
    Dim cellVal As String
    Dim lineStr As String
    
    For row = 1 To oWS.UsedRange.Rows.Count
        lineStr = ""
        For col = 1 To 5
            cellVal = Trim(CStr(oWS.Cells(row, col).Value))
            If col < 5 Then
                lineStr = lineStr & cellVal & ","
            Else
                lineStr = lineStr & cellVal
            End If
        Next col
        
        ' Skip completely empty rows
        If Trim(Replace(lineStr, ",", "")) <> "" Then
            Print #fileNum, lineStr
        End If
    Next row
    
    Close #fileNum
    MsgBox "CSV saved at: " & csvPath
End Sub

' ============================================================

Sub SaveAsNotepad()
    Dim csvPath As String
    Dim txtPath As String
    csvPath = ThisWorkbook.Path & "\" & "output.csv"
    txtPath = ThisWorkbook.Path & "\" & "output.txt"
    
    Dim fileIn As Integer, fileOut As Integer
    fileIn = FreeFile
    Open csvPath For Input As #fileIn
    
    fileOut = FreeFile
    Open txtPath For Output As #fileOut
    
    Dim lineStr As String
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
