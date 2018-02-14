Function StripNonAlpha(TextToReplace As String) As String
    Dim ObjRegex As Object
    Set ObjRegex = CreateObject("vbscript.regexp")
    
    With ObjRegex
    .Global = True
    .Pattern = "[^a-zA-Z\s]+"
    StripNonAlpha = .Replace(Replace(TextToReplace, "-", Chr(32)), vbNullString)
    End With
End Function


Sub Range_End_Method()
'Finds the last non-blank cell in a single row or column

Dim lRow As Long
Dim lCol As Long
    
    'Find the last non-blank cell in column A(1)
    lRow = Range("B:B").Cells(Rows.Count, 1).End(xlUp).Row
    
    'Find the last non-blank cell in row 1
    lCol = Cells(1, Columns.Count).End(xlToLeft).Column
    
    MsgBox "Last Row: " & lRow & vbNewLine & _
            "Last Column: " & lCol
  
End Sub


Function Col_Letter(lngCol As Long) As String
    Dim vArr
    vArr = Split(Cells(1, lngCol).Address(True, False), "$")
    Col_Letter = vArr(0)
End Function


Public Function ColumnLetter(Column As Integer) As String
    If Column < 1 Then Exit Function
    ColumnLetter = ColumnLetter(Int((Column - 1) / 26)) & Chr(((Column - 1) Mod 26) + Asc("A"))
End Function



Function ColName(colNum As Integer) As String
    ColName = Split(Worksheets(1).Cells(1, colNum).Address, "$")(1)
End Function
