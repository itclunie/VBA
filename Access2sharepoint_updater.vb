Option Compare Database

Public Function callUpdateR()
    Call updateR_v7
End Function


Sub updateR_v7()

Dim updateRdict As New Dictionary, truefalse As New Dictionary
Dim OneUpdate As Integer
Dim LnkdTbl As String, inputStr As String

DoCmd.OpenForm "UpdateForm"

Dim inTable As String: inTable = Forms![UpdateForm]![Combo106] 'these correspond to a blank form in access.

If inTable = "SYZYGY_Upload_all" Then 
	On Error Resume Next
	Dim AFb2 As Boolean: AFb2 = Forms![UpdateForm]![Check7] '-------link to form checkbox
	truefalse.Add "AFRICOM;B;_", AFb2 '-----------------------------split string later by ;
	Dim AFc2 As Boolean: AFc2 = Forms![UpdateForm]![Check9]
	truefalse.Add "AFRICOM;C;_", AFc2
	Dim AFn2 As Boolean: AFn2 = Forms![UpdateForm]![Check11]
	truefalse.Add "AFRICOM;N;_", AFn2
	
	Dim SOb2 As Boolean: SOb2 = Forms![UpdateForm]![Check50]
	truefalse.Add "SOCOM;B;_", SOb2
	Dim SOc2 As Boolean: SOc2 = Forms![UpdateForm]![Check52]
	truefalse.Add "SOCOM;C;_", SOc2
	Dim SOn2 As Boolean: SOn2 = Forms![UpdateForm]![Check54]
	truefalse.Add "SOCOM;N;_", SOn2
	
End If


For Each Key In truefalse.Keys
	If truefalse(Key) = True Then
		COCOM_NBC = Split(Key, ";") 'COCOM_NBC(0)=cocom, COCOM_NBC(1)=N/B/C, COCOM_NBC(2)=CH to signify china

		inputSQL = "SELECT * FROM " & inTable & " as C WHERE C.COCOM='" & COCOM_NBC(0) & "' AND C.NBC_Master='" & COCOM_NBC(1) & "'"
		Out_table = "_" & COCOM_NBC(0) & "_" & COCOM_NBC(1) & "2"
		Debug.Print (inputSQL)
		Debug.Print (Out_table)
		
		Erase COCOM_NBC
		
		runfunction = MirrorTables(inputSQL, Out_table) '*the function below
		
	End If
Next Key
	
End Sub


Public Function MirrorTables(ByVal in_SQL_selection As String, ByVal Out_LnkdTbl As String) As String

Dim inputRS As recordset: Set inputRS = CurrentDb.OpenRecordset(in_SQL_selection) '--------------------fill input record set object with sql query
Dim outputRS As recordset: Set outputRS = CurrentDb.OpenRecordset("SELECT * FROM " & Out_LnkdTbl) '----fill output record set object with sql query
Dim tmpField As Field
Dim keyStr As Variant, itmStr As Variant
Dim loopCount1 As Integer, loopCount2 As Integer: loopCount2 = 1
Dim inputAll As New Dictionary, inputRow As New Dictionary, outputFields As New Dictionary
Dim tempDict As Dictionary

SysCmd acSysCmdSetStatus, "Deleting old records, may take up to 5 minutes" 'status meter
CurrentDb.Execute "DELETE FROM " & Out_LnkdTbl & ";", dbFailOnError 'delete all records in output table
SysCmd acSysCmdClearStatus

inputRS.MoveFirst
Do While Not inputRS.EOF '----------------------------------loop through rows in table
	For i = 0 To inputRS.Fields.Count - 1 '-----------------# of columns in SYZYGY_Upload_all table
		
		keyStr = inputRS.Fields(i).SourceField '------------fields object holds column headers, make headers key in temp dictionary
		itmStr = inputRS.Fields(i) '------------------------the text under that header, make item in temp dictionary
		
		inputRow.Add keyStr, itmStr '-----------------------load temp inner dictionary
	Next i

	inputAll.Add "innerDict" & str(loopCount1), inputRow '----load outer dictionary with inner dict
	Set inputRow = Nothing '----------------------------------empty inner dictionary

	inputRS.MoveNext
	loopCount1 = loopCount1 + 1
Loop

For i = 0 To outputRS.Fields.Count - 1 '--------------------# of columns in output table
	outputFields.Add outputRS.Fields(i).SourceField, "" '---load output column headers into dict to check against later
Next i
   
SysCmd acSysCmdInitMeter, "Writing output to " & Out_LnkdTbl, inputAll.Count 'create progress bar
loopCount1 = 0

For Each Key In inputAll.Keys '----------------------------loop thru each key/value pair from 2d input dictionary
	Set tempDict = inputAll(Key) '-------------------------reference each inner dict in turn, set each to a temp dict
	
	outputRS.AddNew '--------------------------------------add blank row to outputRS
	
	For Each Key2 In tempDict.Keys '-----------------------loop thru each key/value pair from each inner input dictionary, ie each row from input table

		If tempDict(Key2) <> Empty Then '------------------check if tempDict(Key2) is empty value, skip
			If outputFields.Exists(Key2) Then '------------check if the fields match b/t tables, if so update record
				Set tmpField = outputRS.Fields(Key2) '-----set tmpField to a particular column (key2). The "For ... inputAll.Keys" loop adds a new row each iteration, so now we essentially have a row number as well as a col header
				tmpField = tempDict(Key2) '----------------fill with value tempDict(Key2)
			End If
		End If
	Next Key2
	
	outputRS.Update  '-------------------------------------cement changes
	
	loopCount1 = loopCount1 + 1
	loopCount2 = loopCount2 + 1
	
	If loopCount1 = 10 Then '-------------------------------see progress section
		Debug.Print (Out_LnkdTbl & ": " & str(loopCount2) & " out of " & inputAll.Count)
		SysCmd acSysCmdUpdateMeter, loopCount2 '------------update progress bar
		loopCount1 = 0
		DoEvents
	End If
Next Key

Debug.Print ("done updating " & Out_LnkdTbl)
SysCmd acSysCmdRemoveMeter

End Function
