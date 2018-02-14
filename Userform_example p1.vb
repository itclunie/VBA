Option Explicit

Sub showBBForm()
    BB2Portal.Show
End Sub

Public Function saveAsCSV(ByVal worksheetExport As String, ByVal desiredPathandFilename As String)
    Sheets(worksheetExport).Copy
    ActiveWorkbook.SaveAs Filename:=desiredPathandFilename, FileFormat:=xlCSV
    ActiveWorkbook.Close savechanges:=False
    ThisWorkbook.Activate
End Function

Public Function LastRowfunc(ByVal sheetname As String, ByVal wholeCol As String) As Long
    ThisWorkbook.Sheets(sheetname).Select
    Dim lRow As Long
    lRow = Range(wholeCol).Cells(Rows.Count, 1).End(xlUp).Row
    LastRowfunc = lRow
End Function

Public Function nav2file() As String
    Dim fd As Office.FileDialog: Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    With fd
        .AllowMultiSelect = False
        
        ' Set the title of the dialog box.
        .Title = "Select the Black Book export file to be converted to .csv"
        
        ' Clear out the current filters, and add our own.
        .Filters.Clear
        .Filters.Add "Excel xlsx", "*.xlsx"
        .Filters.Add "Excel xlsm", "*.xlsm"
        .Filters.Add "Excel xls", "*.xls"
        .Filters.Add "All Files", "*.*"
        
        ' Show the dialog box. If the .Show method returns True, the user picked at least one file. If the .Show method returns False, the user clicked Cancel.
        If .Show = True Then
            nav2file = .SelectedItems(1) 'replace txtFileName with your textbox
        End If
    End With
End Function

Public Function BrowseForFolder(Optional OpenAt As String) As String
     'Function purpose:  To Browser for a user selected folder.
     'If the "OpenAt" path is provided, open the browser at that directory
     'NOTE:  If invalid, it will open at the Desktop level
    Dim ShellApp As Object
     
     'Create a file browser window at the default folder
    Set ShellApp = CreateObject("Shell.Application").BrowseForFolder(0, "Please choose a folder", 0, OpenAt)
     
    On Error Resume Next 'Set the folder to that selected.  (On error in case cancelled)
    BrowseForFolder = ShellApp.self.path
    On Error GoTo 0
     
    Set ShellApp = Nothing
     
     'Check for invalid or non-entries and send to the Invalid error handler if found
     'Valid selections can begin L: (where L is a letter) or \\ (as in \\servername\sharename.  All others are invalid
    Select Case Mid(BrowseForFolder, 2, 1)
    Case Is = ":"
        If Left(BrowseForFolder, 1) = ":" Then GoTo Invalid
    Case Is = "\"
        If Not Left(BrowseForFolder, 1) = "\" Then GoTo Invalid
    Case Else
        GoTo Invalid
    End Select
     
    Exit Function
     
Invalid:
     'If it was determined that the selection was invalid, set to False
    BrowseForFolder = False
End Function

Public Function RegexCollectionFunc(YourRegexPattern As String, TextToSearch As String, Optional ByVal Sub_match As String) As Object
    Set RegexCollectionFunc = New Collection
    Dim ObjRegex As RegExp
    Dim WinHttpReq As Object
    Dim allMatches As MatchCollection
    Dim holdMatch As String
    Set ObjRegex = CreateObject("vbscript.regexp")
    
    With ObjRegex
        .Global = True
        .MultiLine = True
        .IgnoreCase = True
        .Pattern = YourRegexPattern
    End With
    
    Set allMatches = ObjRegex.Execute(TextToSearch)
    Dim RxMatch As Variant
    
    If allMatches.Count <> 0 Then
        For Each RxMatch In allMatches
            If Sub_match = "" Then
                holdMatch = RxMatch
            Else
                holdMatch = RxMatch.SubMatches.item(Sub_match)
            End If
            RegexCollectionFunc.Add holdMatch
        Next RxMatch
    End If
End Function

Public Function CleanString(strIn As String) As String
        Dim ObjRegex As RegExp
        Set ObjRegex = CreateObject("vbscript.regexp")
    With ObjRegex
     .Global = True
     .Pattern = "[\n]+"
    CleanString = .Replace(strIn, vbNullString)
    End With
End Function


Public Function collection2array(c As Collection) As Variant() 'converts a collection to an array
    Dim a() As Variant: ReDim a(0 To c.Count - 1)
    Dim i As Integer
    For i = 1 To c.Count
        a(i - 1) = c.item(i)
    Next
    collection2array = a
End Function

Public Function htmlUL(inputStr As String) As String 'takes a string & makes an html unordered list.
    Dim item As Variant, word As Variant
    Dim linCol As New Collection
    Dim sentenCol As New Collection
    Dim wordCountr As Integer, lineCountr As Integer
    Dim lineStr As String, unordList As String
    Dim alphanum As String

    Dim brArray() As Variant
    Dim tempArray() As Variant
    Dim lineArray() As String
    
    inputStr = Trim(inputStr)
    inputStr = AlphaNumericOnly(inputStr) 'strip out hidden shit
    
    Dim sentenArray() As String: sentenArray = Split(inputStr, ". ") 'break apart paragraph on periods
    
    sentenCol.Add "<ul>" 'open tag unordered list. add to collection
    
    For Each item In sentenArray
        lineArray = Split(item, " ") 'break into words on spaces
    
        For Each word In lineArray
            wordCountr = wordCountr + 1
            linCol.Add (word & " ")
            
'            If wordCountr = 12 Then  '<br> at 9 words / line
''                linCol.Add ("<br>")
'                lineCountr = lineCountr + 1
'                wordCountr = 0
'            End If
'
'            If lineCountr = 10 Then Exit For 'max lines reached. exit foreach word loop
        Next word
'
'        If lineCountr = 10 Then Exit For  'max lines reached, continue exiting loop
        
        If linCol.Count = 0 Then
            lineStr = ""
        Else
            brArray = collection2array(linCol)
            lineStr = Join(brArray, "")
        End If
        
        sentenCol.Add "<li>" & Trim(lineStr) & "</li>"
    
        Set linCol = Nothing
        wordCountr = 0
    Next item
    
    sentenCol.Add "</ul>"

    tempArray = collection2array(sentenCol)
    unordList = Join(tempArray, "")

    htmlUL = unordList
End Function


Public Function AlphaNumericOnly(strSource As String) As Variant
    Dim i As Integer
    Dim strResult As String

    For i = 1 To Len(strSource)
        Select Case Asc(Mid(strSource, i, 1))
			'list of chr() exceptions. include 32 if you want to include space
            Case 48 To 57, 65 To 90, 97 To 122, 32, 34, 39, 40, 41, 44, 45, 46: 
                strResult = strResult & Mid(strSource, i, 1)
        End Select
    Next
    AlphaNumericOnly = strResult
End Function


