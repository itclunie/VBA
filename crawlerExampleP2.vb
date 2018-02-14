Public Function newFunc(ByRef html As String, ByVal url As String) As String
'    Application.ScreenUpdating = False

    Dim resultsCollection As New Collection

    Dim regEx As New RegExp 'Need Reference set for "Microsoft VBScript Regular Expression 5.5"
    Dim WB As Workbook
    Dim WS As Worksheet
    Dim FirstRow As Long, LastRow As Long
    Dim url_index As Integer, hostname_index As Integer, HTTP_Response As Integer, counter As Integer, _
        status_index As Integer, retryCounter As Integer
    Dim html_response As String, RetryArr() As String, Button_Caption As String, currentPath As String, _
        attachmentPath As String, tmpstr As String, snipet As String, downloadLink As String, _
		linkAndDescrip(1) As String, tmpCaption As String
    Dim result As Boolean, https_status As Boolean, Debugging As Boolean
    Dim columnNames() As Variant
    Dim headers_col As New Dictionary
    
    Debugging = True                                        ' Set to false to suppress debugging testing code and alerts
    currentPath = ThisWorkbook.Path                         ' Store the path for where this Excel file is running, ex. "C:\User\Me\Documents"
    
    Application.EnableEvents = True                         ' Allow for event to be triggered in Excel
    If Debugging Then
        On Error GoTo 0                                     ' If code breaks, then go to debugger and keep trying until user decides to quit
    End If
    Err.Clear                                               ' Clear any error
    Application.ScreenUpdating = False                      ' Prevent screen from updating til code is finished; speeds up code
    
    Set WB = ActiveWorkbook                                 ' Set variable to point to current workbook
    Set WS = WB.Sheets("Data")                              ' Set variable to point to worksheet which will store results

    FirstRow = 3                                            ' Row 1 is start button, Row 2 is header, This is First Row after Header
    LastRow = WS.Range("A1", WS.Range("A1048576").End(xlUp)).Count + 1 'row to start enter data in Excel spreadsheet (Sheet set by WS)
    
    Source_Year_col = 16
    sitedateID_col = 17
    columnNames = Array("Name", ---ETC--- ")
    For counter = 0 To UBound(columnNames)
        headers_col(columnNames(counter)) = counter + 1
    Next counter
    headers_col("Source URL") = sitedateID_col + 1
    headers_col("Chemicals used/produced") = 0
    headers_col("Plants of this site") = 0

    WS.Cells(LastRow, headers_col("Source URL")) = url

    html_response = html
    
    Dim strPattern As String: strPattern = "<ul>.*?\n.*?<li><a href=""#"" onclick=""doMap\('(\d+\|\d+)"

    Dim allMatches1 As Object
    Dim regexResult As String
    
    With regEx
        .Global = True
        .MultiLine = True
        .IgnoreCase = True
        .Pattern = strPattern
    End With
    
    Set allMatches1 = regEx.Execute(html_response)
    
    If allMatches1.Count <> 0 Then
        regexResult = allMatches1.item(0).SubMatches.item(0)
        WS.Cells(LastRow, sitedateID_col) = regexResult
    End If
  
    With regEx
        .Global = True
        .Pattern = "\n|\r"
    End With
    
    Dim tempStr As String
    tempStr = regEx.Replace(html_response, "#####")  ' Create Temp variable to remove all newline that prevent regex from running multiline

    strPattern = "<th class=""rowhead"">(.+?)</th>.*?<td.*?>(.*?)</tr>"
    regEx.Pattern = strPattern

    Set allMatches1 = regEx.Execute(tempStr)
    Dim category As String: category = ""
    Dim content As String: content = ""
    Dim contentMatches As Object
    Dim printCol As Integer: printCol = 0
    If allMatches1.Count <> 0 Then
        For Each Match In allMatches1
            content = Match.SubMatches.item(1)
            regEx.Global = True
            regEx.Pattern = "</?span.*?>"
            category = regEx.Replace(Match.SubMatches.item(0), "")
            category = WorksheetFunction.Trim(category)
            printCol = headers_col(category)

            Select Case category
                Case "A", "B", "C", "D", "E", "F"
                    regEx.Pattern = ">?\s*(.+?)\s*<"
                Case "G"
                    regEx.Pattern = """>(.+?)</a>"
                Case "H", "I"
                    regEx.Global = True
                    regEx.Pattern = "<br />#+\s*"
                    content = regEx.Replace(content, ", ")
                    regEx.Pattern = "(.+?)#*\s*</td>"
                Case "J"
                    regEx.Global = True
                    regEx.Pattern = "<br />#+\s*"
                    content = regEx.Replace(content, "; ")
                    regEx.Pattern = "#*\s*(.+?)#*\s*</td>"
                Case "K"
                    regEx.Pattern = "title=""(.+?)"">(.+?)<"
                    WS.Cells(LastRow, printCol) = ""
                    regEx.Global = True
                    Set contentMatches = regEx.Execute(content)
                    For Each contentMatch In contentMatches
                        WS.Cells(LastRow, printCol) = WS.Cells(LastRow, printCol) & contentMatch.SubMatches.item(1) & _
                            ":" & contentMatch.SubMatches.item(0) & "; "
                    Next contentMatch
                    GoTo SkipMatch
                Case "L"
                    WS.Cells(LastRow, headers_col("L1")) = ""
                    WS.Cells(LastRow, headers_col("L2")) = ""
                    GoTo MultipleValues
                Case "M"
                    WS.Cells(LastRow, headers_col("M1")) = ""
                    WS.Cells(LastRow, headers_col("M2")) = ""
                    GoTo MultipleValues
                Case Else
                    Debug.Print (category)
                    GoTo SkipMatch
            End Select

            regEx.Global = False
            Set contentMatches = regEx.Execute(content)
            If contentMatches.Count > 0 Then
                WS.Cells(LastRow, printCol) = contentMatches.item(0).SubMatches.item(0)
            End If
            GoTo SkipMatch
            
MultipleValues:

            regEx.Global = False
            regEx.Pattern = "None found"
            Set contentMatches = regEx.Execute(tempStr)
            If contentMatches.Count > 0 Then GoTo SkipMatch
            
            regEx.Pattern = "<th class=""rowhead"">" & category & "</th>.*?<table>.*?<tr>(.*?)</tr>.*?<tr>(.*?)</table>"
            Set contentMatches = regEx.Execute(tempStr)
            regEx.Global = True
            regEx.Pattern = "<th class=""rowhead"">(.+?)</th>"
            Set headerMatches = regEx.Execute(contentMatches.item(0).SubMatches.item(0))
            regEx.Pattern = "<td>(.*?)</td>"
            Set contentMatches = regEx.Execute(contentMatches.item(0).SubMatches.item(1))
            regEx.Pattern = "<.*?>"
            For counter = 0 To contentMatches.Count - 1
                If headerMatches.item(counter Mod headerMatches.Count).SubMatches.item(0) <> "Page" Then
                    WS.Cells(LastRow, headers_col(headerMatches.item(counter Mod headerMatches.Count).SubMatches.item(0))) = WS.Cells _
                    (LastRow, headers_col(headerMatches.item(counter Mod headerMatches.Count).SubMatches.item(0))) & _
                    regEx.Replace(contentMatches.item(counter).SubMatches.item(0), "") & "; "
                End If
            Next counter
'            If contentMatches.count > headerMatches.count Then
'                Debug.Print (category & ": " & contentMatches.count)
'            End If
            
SkipMatch:
        Next Match
    Else
        'newFunc = newFunc & ("Name - No Match")
    End If
    Set contentMatches = Nothing
DoEvents

    regEx.Pattern = "<th class=""subhead"">Year.+?</a></td>.*?<td>(.+?)</td>.*?<td>Current</td>"

    Set allMatches1 = regEx.Execute(tempStr)                                           ' Execute Regex on HTML without newlines
    If allMatches1.Count <> 0 Then                                                     ' If match, do code that follows
        regexResult = Replace(allMatches1.item(0).SubMatches.item(0), "#####", "")     ' Remove 5x octothorph (#) used to replace newlines
        regEx.Pattern = "<.+?>"                                                        ' Set new Regex Pattern to match all HTML tags, ie. <li>, <br>, <div>, etc...
        regexResult = regEx.Replace(regexResult, "")                                   ' Run regex Replace to remove all HTML tags using pattern from previous line
        regexResult = WorksheetFunction.Trim(WorksheetFunction.Clean(regexResult))     ' Use worksheet functions Trim (Remove all leading/trailing spaces but leave one space between words
        WS.Cells(LastRow, Source_Year_col) = regexResult                               ' Print results to sheet
    Else
        'Debug.Print ("Source Year - No Match")
    End If

    Set regEx = Nothing
    newFunc = ""
'    Debug.Print (Timer - startTime)
End Function





