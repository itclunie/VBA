Sub Macro2()

Application.StatusBar = ""
Dim startTime As Double: startTime = Timer
Dim incrementTime As Double: incrementTime = startTime
Application.ScreenUpdating = True
'    Application.ScreenUpdating = False
Application.EnableEvents = False
Application.Calculation = xlCalculationManual

Dim W_B As Workbook: Set W_B = ActiveWorkbook        ' Set variable to point to current workbook
Dim W_S As Worksheet: Set W_S = W_B.Sheets("Sheet2")  'where the top urls are


Dim currentwindow As HTMLWindowProxy
currentPath = ThisWorkbook.Path
Dim AllCountriesAllYears As String
Open currentPath & "\saveTimeHacks" & Int(startTime * 10000) & ".txt" For Output As #3
	Write #3, W_B.Name
Close #3 ' empty contents

'1--------------------------------------------------------------------------------------------------------------------------------1st get

baseUrl = "http://---etc---.do?"
AllCountriesAllYears = ThisWorkbook.Sheets("sheet2").Range("K1")
baseUrl = baseUrl & AllCountriesAllYears 'in sheet2 is the whole query to pull all results since it was too long for here
	
Dim response As String
Dim resultID As String
Dim WinHttpReq As Object

Set WinHttpReq = getUrlHTML(baseUrl) '1st get; 10sec, once

'2--------------------------------------------------------------------------------------------------------------------------------2nd get
'full pull will be 30sec, once
If WinHttpReq.Status = 200 And Not IsEmpty(WinHttpReq.responseBody) Then
	Set WinHttpReq = getUrlHTML("http://---etc---.jsp?resultsPerPage=100000") '**** ADJUST THIS SECTION TO MAKE READY
	'*CHANGE "resultsPerPage=#" to "=100000" to get all results*
End If

'3-------------------------------------------------------------------------------------------------------------------------------1st regex
'1st step making each result url

Dim dict As New Dictionary
Dim front_url As String: front_url = "http://---etc---.do?RequestedSiteID="
Dim myRange As Range: Set myRange = W_S.Range("A1", W_S.Range("A1048576").End(xlUp))
Dim nextAvailableCell As Range: Set nextAvailableCell = W_S.Range("A" & myRange.Count)

If nextAvailableCell <> "" Then Set nextAvailableCell = nextAvailableCell.Offset(1, 0)
For Each link In myRange    '
	dict(link.Value) = link.Row             ' Used to check/update later whether to run link; .offset(0,1) is true for run, false for skip
Next link

Dim strPattern As String: strPattern = "site(\d+)" 'match site id
Dim regEx As New RegExp 'make sure microsoft.vbscript regular expressions 5.5 library is turned on under tools -> references
Dim allMatches As MatchCollection '
   
With regEx 'params
	.Global = True
	.MultiLine = True
	.IgnoreCase = True
	.Pattern = strPattern
End With

Set allMatches = regEx.Execute(WinHttpReq.responseText)  'temp collection that houses text to be matched

If allMatches.Count <> 0 Then
	For Each RxMatch In allMatches
		If Not dict.Exists(front_url & RxMatch.SubMatches.item(0)) Then             ' if links is not on sheet 2, then add it and update myrange
			nextAvailableCell.Value = front_url & RxMatch.SubMatches.item(0)            ' add new link to sheet 2
			nextAvailableCell.Offset(0, 2).Value = "Current"                            ' annotate in column C that it's "Current" year link
			dict(front_url & RxMatch.SubMatches.item(0)) = nextAvailableCell.Row        ' add new link as key in dict
			Set myRange = W_S.Range(myRange, nextAvailableCell)                         ' expand range to include new link
			Set nextAvailableCell = nextAvailableCell.Offset(1, 0)                      ' advanced pointer to next available cell
		End If
	Next RxMatch
Else
	'Debug.Print ("No Match")
End If

'5-----------------------------------------------------------------------------------------------------------------------------gets & regexes for each result
'gets for each top url and regex pulls for current year.
	
regEx.Pattern = "(&RequestedSiteDocumentID=.*?)"">Select</a></td>"
Dim mycounter2 As Integer: mycounter2 = 0
Dim saveIncrement As Integer: saveIncrement = 50
Dim item As Range
Dim i As Long, stopPoint As Long, linkPointer As Long, pulledCountAtStart As Long

linkPointer = 1
stopPoint = myRange.Count
pulledCountAtStart = WorksheetFunction.CountIf(myRange.Offset(0, 1), "Pulled")
'    For Each item In myRange 'run each url to run new_Func and skip searching for date URLs if page isn't a "Current" year link
For i = 1 To stopPoint 'run each url to run new_Func and skip searching for date URLs if page isn't a "Current" year link
	Set item = myRange(linkPointer)
	If item.Offset(0, 1).Value = "Pulled" Or IsEmpty(item) Then GoTo NextItem                   ' skip link if already pulled or link is empty
	
	Set WinHttpReq = getUrlHTML(item.Value)
	If WinHttpReq.Status = 200 And Not IsEmpty(WinHttpReq.responseText) Then
		If item.Offset(0, 2) = "Current" Then
			Set allMatches = regEx.Execute(WinHttpReq.responseText)
			
			If allMatches.Count <> 0 Then
				For Each RxMatch In allMatches
					If Not dict.Exists(item.Value & RxMatch.SubMatches.item(0)) Then             ' if links is not on sheet 2, then add it and update myrange
						nextAvailableCell.Value = item.Value & RxMatch.SubMatches.item(0)            ' add new link to sheet 2
						nextAvailableCell.Offset(0, 2).Value = ""                                    ' annotate in column C that it is not "Current" year link
						dict(item.Value & RxMatch.SubMatches.item(0)) = nextAvailableCell.Row        ' add new link as key in dict
						Set myRange = W_S.Range(myRange, nextAvailableCell)                          ' expand range to include new link
						i = i - 1
						Set nextAvailableCell = nextAvailableCell.Offset(1, 0)                       ' advanced pointer to next available cell
					End If
				Next RxMatch
			End If
		End If
		
		If newFunc(WinHttpReq.responseText, item.Value) <> "" Then '***regex function in new_Func*********
		 '.... do something with error feedback string
		End If
		mycounter2 = mycounter2 + 1
		item.Offset(0, 1).Value = "Pulled"
		
		Application.StatusBar = "Links Remaining: " & myRange.Count - pulledCountAtStart - mycounter2 & _
		"; Estimated time remaining: " & convertSecToTimeString(CLng(((Timer - startTime) / _
		mycounter2) * (myRange.Count - pulledCountAtStart - mycounter2)))
		
		If mycounter2 Mod saveIncrement = 0 Then
			ActiveWorkbook.Save
			Open currentPath & "\saveTimeHacks" & Int(startTime * 10000) & ".txt" For Append As #3
				Write #3, "Links[" & mycounter2 - saveIncrement + 1 & "-" & mycounter2 & "]:" & Timer - incrementTime & "s" & Chr(10)
			Close #3
			incrementTime = Timer
		End If
	End If

NextItem:
	linkPointer = linkPointer + 1
Next i
  
Finish:
Set allMatches = Nothing
Set regEx = Nothing
Set dict = Nothing
Set WinHttpReq = Nothing

Application.ScreenUpdating = True
Application.EnableEvents = True
Application.Calculation = xlCalculationAutomatic
ActiveWorkbook.Save
Debug.Print ("Code completed in: " & Timer - startTime)
Open currentPath & "\saveTimeHacks" & Int(startTime * 10000) & ".txt" For Append As #3
	Write #3, "Code completed in: " & Timer - startTime
Close #3
Application.StatusBar = "Complete"
	
End Sub

Private Function getUrlHTML(ByVal url As String) As Object
'        Dim WinHttpReq As Object
        Set getUrlHTML = CreateObject("Microsoft.XMLHTTP")
        getUrlHTML.Open "GET", url, False
        getUrlHTML.setRequestHeader "Connection", "keep-alive" 'maintian cookie
        getUrlHTML.setRequestHeader "User-Agent", "Mozilla/4.0 (Compatable; MSIE 6.0; Windows NT 5.0)"
        
        If strCookie <> "" Then
            getUrlHTML.setRequestHeader "Cookie", strCookie
        End If
        getUrlHTML.send
End Function

Private Function convertSecToTimeString(ByVal seconds As Long) As String
'    dim secs as
    convertSecToTimeString = ""
    If seconds <= 0 Then
        convertSecToTimeString = "0s"
    ElseIf seconds < 1 And seconds > 0 Then
        convertSecToTimeString = "<1s"
    Else
        While seconds > 0
            convertSecToTimeString = convertSecToTimeString & extractAndDecrementLargestTimeUnit(seconds) & " "
        Wend
        convertSecToTimeString = Left(convertSecToTimeString, Len(convertSecToTimeString) - 1)
    End If
End Function

Private Function extractAndDecrementLargestTimeUnit(ByRef seconds As Long) As String
    Dim denominator As Long
    Dim suffix As String
    Select Case CLng(seconds)
        Case Is >= CLng(CLng(60) * 60 * 24) ' days
            denominator = CLng(CLng(60) * 60 * 24)
            suffix = "d"
        Case Is >= CLng(60 * 60)   ' hours
            denominator = CLng(60 * 60)
            suffix = "h"
        Case Is >= CLng(60)        ' minutes
            denominator = CLng(60)
            suffix = "m"
        Case Else               ' seconds
            denominator = 1
            suffix = "s"
    End Select
    
    extractAndDecrementLargestTimeUnit = Str(CLng(seconds / denominator)) & suffix
    seconds = CLng(seconds Mod denominator)
End Function

