Option Compare Database

Public Function callUpdateR()
    Call updateR_v7
End Function


Sub updateR_v7()
    Dim updateRdict As New Dictionary, truefalse As New Dictionary
    Dim OneUpdate As Integer
    Dim LnkdTbl As String, inputStr As String
    
    DoCmd.OpenForm "UpdateForm"
    
    Dim inTable As String: inTable = Forms![UpdateForm]![Combo106]
    
    If inTable = "SYZYGY_Upload_all" Then
        On Error Resume Next
        Dim AFb2 As Boolean: AFb2 = Forms![UpdateForm]![Check7] '-------link to form checkbox
            truefalse.Add "AFRICOM@B@_", AFb2 '-----------------------------split string later by @
        Dim AFc2 As Boolean: AFc2 = Forms![UpdateForm]![Check9]
            truefalse.Add "AFRICOM@C@_", AFc2
        Dim AFn2 As Boolean: AFn2 = Forms![UpdateForm]![Check11]
            truefalse.Add "AFRICOM@N@_", AFn2
        
		'etc...
                        
    ElseIf inTable = "PATHOGENS_COMBINED" Then
        On Error Resume Next
        Dim AFpath As Boolean: AFpath = Forms![UpdateForm]![Check4]
            truefalse.Add "AFRICOM@@", AFpath
        Dim SOpath As Boolean: SOpath = Forms![UpdateForm]![Check48]
            truefalse.Add "SOUTHCOM@@", SOpath
		
		'etc...
		
    End If
    
    
    For Each Key In truefalse.Keys
        countr = countr + 1
        If truefalse(Key) = True Then
            COCOM_NBC = Split(Key, "@") 'COCOM_NBC(0)=cocom, COCOM_NBC(1)=N/B/C, COCOM_NBC(2)=CH to signify china
            
            COCOM = inTable & ".COCOM='" & COCOM_NBC(0) & "'"
            NBC_Master = inTable & ".NBC_Master='" & COCOM_NBC(1) & "'"
            CountryCH = inTable & ".Country='China'"
            CountryRUS = inTable & ".Country='Russia'"
            
            If COCOM_NBC(2) = "CH" Then
                endFilter = " WHERE " & COCOM & " AND " & NBC_Master & " AND " & CountryCH
                Out_table = "_China_" & COCOM_NBC(1) & "2"
            ElseIf COCOM_NBC(0) = "PACOM" And (COCOM_NBC(1) = "N" Or COCOM_NBC(1) = "C") Then 'exclude China N and C.
                endFilter = " WHERE " & COCOM & " AND " & NBC_Master & " AND NOT " & CountryCH
                Out_table = "_" & COCOM_NBC(0) & "_" & COCOM_NBC(1) & "2"
                
            Else
                endFilter = " WHERE " & COCOM & " AND " & NBC_Master
                Out_table = "_" & COCOM_NBC(0) & "_" & COCOM_NBC(1) & "2"
            End If
            
            Erase COCOM_NBC
            
            Debug.Print ("updating list " & countr & ": " & Out_table & " from " & inTable & " " & endFilter)
            runfunction = MirrorTablesSQL(inTable, Out_table, endFilter)
        End If
        DoEvents
    Next Key
endtest:
End Sub

Public Function MirrorTablesSQL(ByVal in_SQL_selection As String, ByVal Out_LnkdTbl As String, ByVal endFilter As String) As String
    Dim inCols As New Collection
    Dim outCols As New Collection
    Dim matchedCol1 As New Collection
    Dim matchedCol2 As New Collection
    Dim tempArray() 'As Variant
    Dim tmpField2 As Field
    Dim dbs As Database: Set dbs = CurrentDb
    
    'clean out record set & add new place holder field
    CurrentDb.Execute "DELETE * FROM " & Out_LnkdTbl & ";", dbFailOnError
    
    checkfield = fieldexists(Out_LnkdTbl, "QQQ")
    If checkfield = False Then
        dbs.Execute "ALTER TABLE " & Out_LnkdTbl & " ADD COLUMN QQQ TEXT(15);"
    End If
    
    Dim inputRS As Recordset: Set inputRS = CurrentDb.OpenRecordset(in_SQL_selection) '-fill input record set object with sql query
    Dim outputRS As Recordset: Set outputRS = CurrentDb.OpenRecordset(Out_LnkdTbl) '----fill output record set object with sql query

    'this bit adds a dummy line so the append sql query works
    outputRS.AddNew

    Set tmpField2 = outputRS.Fields("QQQ")
    tmpField2 = "Place_Holder"
    outputRS.Update


    'load 2 collections with fields
    For i = 0 To inputRS.Fields.Count - 1 'input columns
        inCols.Add "[" & inputRS.Fields(i).SourceField & "]"
    Next i
    For i = 0 To outputRS.Fields.Count - 1 'output columns
        outCols.Add "[" & outputRS.Fields(i).SourceField & "]"
    Next i

    'see if fields from the 2 collections match
    For i = 1 To inCols.Count
        For j = 1 To outCols.Count
            If inCols(i) = outCols(j) Then
                matchedCol1.Add inCols(i)
                matchedCol2.Add in_SQL_selection & "." & inCols(i)
            End If
        Next j
    Next i

'flatten matched fields into strings
tempArray = collection2array(matchedCol1)
matchedCol1_Str = Join(tempArray, ", ")
Erase tempArray
tempArray = collection2array(matchedCol2)
matchedCol2_Str = Join(tempArray, ", ")

'create/run sql query from flattened strings
dbs.Execute "INSERT INTO " & Out_LnkdTbl & " ( " & matchedCol1_Str & " ) " & _
"SELECT " & matchedCol2_Str & _
" FROM " & in_SQL_selection & ", " & Out_LnkdTbl & endFilter

'delete the place holder
'outputRS.Close
'CurrentDb.Execute "DELETE * FROM " & Out_LnkdTbl & " WHERE " & Out_LnkdTbl & ".[QQQ] = 'Place_Holder';", dbFailOnError
'dbs.Execute "ALTER TABLE " & Out_LnkdTbl & " DROP COLUMN QQQ;"

End Function

Public Function collection2array(c As Collection) As Variant() 'converts a collection to an array h
    Dim a() As Variant: ReDim a(0 To c.Count - 1)
    Dim i As Integer
    For i = 1 To c.Count
        a(i - 1) = c.Item(i)
    Next
    collection2array = a
End Function

Function fieldexists(tablename As String, fieldname As String) As Boolean
Dim exists As Boolean
   exists = False
   On Error Resume Next
   exists = CurrentDb.TableDefs(tablename).Fields(fieldname).Name = fieldname
   fieldexists = exists
End Function
