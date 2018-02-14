'ctr n
Sub normalizeAlphanumeric()

range1str = InputBox("select 1st range to normalize to alphanumeric + spaces", "Normalize Range1")
range2str = InputBox("select 2nd range to normalize to alphanumeric + spaces", "Normalize Range2")

Set range1 = ActiveWorkbook.ActiveSheet.Range(range1str)
Set range2 = ActiveWorkbook.ActiveSheet.Range(range2str)
Dim mycell1 As Range
Dim mycell2 As Range


For Each mycell1 In range1
    If mycell1 <> "" Then
        mycell1.Offset(0, 1) = AlphaNumericOnly(mycell1.Value)
    End If
Next mycell1

For Each mycell2 In range2
    If mycell2 <> "" Then
        mycell2.Offset(0, 1) = AlphaNumericOnly(mycell2.Value)
    End If
Next mycell2


End Sub


Function AlphaNumericOnly(strSource As String) As String
    Dim I As Integer
    Dim strResult As String

    For I = 1 To Len(strSource)
        Select Case Asc(Mid(strSource, I, 1))
            Case 48 To 57, 65 To 90, 97 To 122, 32: 'include 32 if you want to include space
                strResult = strResult & Mid(strSource, I, 1)
        End Select
    Next
    
    AlphaNumericOnly = UCase(strResult)
End Function
