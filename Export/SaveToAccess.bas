Attribute VB_Name = "SaveToAccess"
'@Folder "Access"
Option Explicit

Public Sub SaveGroupedChangesToAccess(ByVal grpChanges As GroupedDatabaseChanges)
    Dim Connection As ADODB.Connection
    Set Connection = New ADODB.Connection
    Connection.Open ConnectionString:=Access.GetConnectionString
    
    Dim Group As Variant
    For Each Group In grpChanges.Items
        GenerateTrackingRecord Connection, Group
        UpdateOneRecord Connection, Group
    Next Group
    
    Connection.Close
    Set Connection = Nothing
End Sub

Private Sub UpdateOneRecord(ByVal db As ADODB.Connection, ByVal Group As GroupedDatabaseChange)
    Dim origRS As ADODB.Recordset
    Dim copyRS As ADODB.Recordset
    
    Set origRS = New ADODB.Recordset
    Set copyRS = New ADODB.Recordset

    Dim sql As String
    sql = "SELECT " & Group.TableName & ".*, " & Access.TRACKS_TABLE_NAME & ".ValidUntil"
    sql = sql & " FROM " & Group.TableName & " INNER JOIN " & Access.TRACKS_TABLE_NAME
    sql = sql & " ON " & Group.TableName & "." & Access.TRACK_FIELD_NAME & " = " & Access.TRACKS_TABLE_NAME & ".ID"
    sql = sql & " WHERE " & Group.TableName & "." & Access.KEY_FIELD_NAME & " = " & Group.KeyFK
    sql = sql & " AND " & Access.TRACKS_TABLE_NAME & "." & Access.VALID_UNTIL_FIELD_NAME & " = #9999/12/31#;"

    origRS.Open sql, db, adOpenDynamic, adLockOptimistic
    copyRS.Open Group.TableName, db, adOpenDynamic, adLockOptimistic
    
    Dim dbChange As DatabaseChange
    If Not origRS.BOF And Not origRS.EOF Then
        copyRS.AddNew
        
        Dim fld As Field
        For Each fld In origRS.fields
            If fld.Name <> Access.VALID_UNTIL_FIELD_NAME And fld.Name <> "ID" Then
                copyRS.fields(fld.Name).Value = origRS.fields(fld.Name).Value
            End If
        Next fld
        
        For Each dbChange In Group.Items
            copyRS.fields(dbChange.FieldName).Value = dbChange.Value
        Next dbChange
        
        copyRS.fields(Access.TRACK_FIELD_NAME).Value = Group.trackFK

        
        copyRS.Update
        copyRS.Close
        
        origRS.fields(Access.VALID_UNTIL_FIELD_NAME).Value = Now()
        origRS.Update
        origRS.Close
    Else
        ' new record
        copyRS.AddNew

        For Each dbChange In Group.Items
            copyRS.fields(dbChange.FieldName).Value = dbChange.Value
        Next dbChange
        copyRS.fields(Access.KEY_FIELD_NAME).Value = Group.KeyFK
        copyRS.fields(Access.TRACK_FIELD_NAME).Value = Group.trackFK
        
        copyRS.Update
        copyRS.Close
    End If
            
    Set origRS = Nothing
    Set copyRS = Nothing
End Sub

Private Sub GenerateTrackingRecord(ByVal db As ADODB.Connection, ByVal groupChanges As GroupedDatabaseChange)
    Dim sql As String
    sql = "INSERT INTO " & Access.TRACKS_TABLE_NAME & " (ValidFrom, ValidUntil, CommitFK, KeyFK, TableName) VALUES (#" & Format$(Now(), "yyyy/mm/dd") & "#, #9999/12/31#, " & groupChanges.CommitFK & ", " & groupChanges.KeyFK & ", '" & groupChanges.TableName & "');"
    db.Execute sql
    
    Dim trackFK As Long
    Dim rs As ADODB.Recordset
    Set rs = db.Execute("SELECT @@Identity")
    trackFK = rs.fields(0).Value
    
    groupChanges.trackFK = trackFK
End Sub
