Attribute VB_Name = "ExcelToAccess"
'@Folder("VBAProject")
Option Explicit

Private Const DATABASE_PATH As String = "C:\Users\User\Documents\excel-table-change-tracking\Database1.accdb"
Private Const SCHEMA_PATH As String = "C:\Users\User\Documents\excel-table-change-tracking\Schema.csv"
Private Const KEY_FK As String = "KeyFK"
Private Const TRACK_FK As String = "TrackFK"

Private Connection As ADODB.Connection
Private Connect As String
Private Recordset As ADODB.Recordset
Private sql As String
Private Schema As Collection
Private KeyTranslation As Collection

Public Sub AAATest()
    Dim coll As Collection
    Set coll = New Collection
    With New FieldChange
        .key = "ABC"
        .ColumnName = "FieldA"
        .After = "XYZ"
        coll.Add .Self
    End With
    
    SaveChangesToDatabase coll
End Sub

Private Function GetConnectionString()
    GetConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & DATABASE_PATH & ";"
End Function

Public Sub LoadChangesFromDatabase()
    Set Connection = New ADODB.Connection
    Connection.Open ConnectionString:=GetConnectionString
    
    Set Recordset = New ADODB.Recordset
    sql = "SELECT * FROM tblDetailA"
    Recordset.Open Source:=sql, ActiveConnection:=Connection

    Dim i As Long
    For i = 0 To Recordset.fields.Count - 1
        If i > 0 Then Debug.Print ",";
        Debug.Print Recordset.fields(i).Name;
    Next i
    Debug.Print vbNullString
    
    If Not Recordset.BOF And Not Recordset.EOF Then
        Do While Not Recordset.EOF
            'Dim i As Long
            For i = 0 To Recordset.fields.Count - 1
                If i > 0 Then Debug.Print ",";
                If Not IsNull(Recordset.fields(i).Value) Then
                    Debug.Print CStr(Recordset.fields(i).Value);
                Else
                    Debug.Print "null";
                End If
            Next i
            Debug.Print vbNullString
            Recordset.MoveNext
        Loop
    End If
    
    Set Recordset = Nothing
    Connection.Close
    Set Connection = Nothing
End Sub

Public Sub SaveChangesToDatabase(ByRef Changes As Collection)
    Set Connection = New ADODB.Connection
    Connect = "Provider=Microsoft.ACE.OLEDB.12.0;"
    Connect = Connect & "Data Source=" & DATABASE_PATH & ";"
    Connection.Open ConnectionString:=Connect
    
    LoadSchema
    LoadKeys
    
    Set Recordset = New ADODB.Recordset
    'SQL = "SELECT * FROM " & Schema("LocalKey")(1) & " WHERE "  & Schema("LocalKey")(2) & " = '" & thisfieldchange.key & "
    'Recordset.Open Source:=SQL, ActiveConnection:=Connection

    Dim KeyFK As Long
    Dim TrackFK As Long
    Dim tableForeign As String
    Dim fieldForeign As String

    Dim ThisFieldChange As FieldChange
    For Each ThisFieldChange In Changes
        KeyFK = KeyTranslation(ThisFieldChange.key)
        TrackFK = 1
        tableForeign = Schema(ThisFieldChange.ColumnName)(1)
        fieldForeign = Schema(ThisFieldChange.ColumnName)(2)
        sql = "INSERT INTO " & tableForeign & " ([" & KEY_FK & "], [" & TRACK_FK & "], [" & fieldForeign & "]) VALUES (" & KeyFK & "," & TrackFK & ",'" & ThisFieldChange.After & "');"
        Connection.Execute sql
    '    Select Case VarType(ThisFieldChange.After)
    '        Case vbString
    '            SQL = "INSERT INTO tblDetailA ([Key], [" & ThisFieldChange.ColumnName & "]) VALUES ('" & ThisFieldChange.Key & "','" & ThisFieldChange.After & "');"
    '        Case vbDate
    '            SQL = "INSERT INTO tblDetailA ([Key], [" & ThisFieldChange.ColumnName & "]) VALUES ('" & ThisFieldChange.Key & "',#" & Format(ThisFieldChange.After, "yyyy/mm/dd") & "#);"
    '        Case vbDouble
    '            SQL = "INSERT INTO tblDetailA ([Key], [" & ThisFieldChange.ColumnName & "]) VALUES ('" & ThisFieldChange.Key & "'," & ThisFieldChange.After & ");"
    '
    '    End Select
    '    Connection.Execute SQL
    Next ThisFieldChange

    Set Recordset = Nothing
    Connection.Close
    Set Connection = Nothing
End Sub



