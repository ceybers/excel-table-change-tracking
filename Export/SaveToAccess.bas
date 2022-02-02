Attribute VB_Name = "SaveToAccess"
'@Folder("VBAProject")
Option Explicit

Public Sub SaveGroupedChangesToAccess(ByVal grpChanges As GroupedDatabaseChanges)
    Dim Connection As ADODB.Connection
    Set Connection = New ADODB.Connection
    Connection.Open ConnectionString:=Access.GetConnectionString
    
    ' remember we need to validuntil = now existing stuff
    Dim Group As Variant
    Dim dbChg As DatabaseChange
    
    Dim sql As String
    
    Dim dbChange As DatabaseChange
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
    sql = "SELECT " & Group.TableName & ".*, metaTracks.ValidUntil FROM " & Group.TableName & " INNER JOIN metaTracks ON " & Group.TableName & ".TrackFK = metaTracks.ID WHERE KeyFK = " & Group.KeyFK

    origRS.Open sql, db, adOpenDynamic, adLockOptimistic
    copyRS.Open "SELECT * FROM " & Group.TableName, db, adOpenDynamic, adLockOptimistic
    
    Dim dbChange As DatabaseChange
    If Not origRS.BOF And Not origRS.EOF Then
        copyRS.AddNew
        
        Dim fld As Field
        For Each fld In origRS.fields
            If fld.Name <> "ValidUntil" And fld.Name <> "ID" Then
                copyRS.fields(fld.Name).Value = origRS.fields(fld.Name).Value
            End If
        Next fld
        
        For Each dbChange In Group.Items
            copyRS.fields(dbChange.FieldName).Value = dbChange.Value
        Next dbChange
        
        copyRS.fields("TrackFK").Value = Group.TrackFK
        
        copyRS.Update
        copyRS.Close
        
        origRS.fields("ValidUntil").Value = Now()
        origRS.Update
        origRS.Close
    Else
        ' new record
        copyRS.AddNew

        For Each dbChange In Group.Items
            copyRS.fields(dbChange.FieldName).Value = dbChange.Value
        Next dbChange
        copyRS.fields("KeyFK").Value = Group.KeyFK
        copyRS.fields("TrackFK").Value = Group.TrackFK
        
        copyRS.Update
        copyRS.Close
    End If
            
    Set origRS = Nothing
    Set copyRS = Nothing
End Sub

Private Sub GenerateTrackingRecord(ByVal db As ADODB.Connection, ByVal groupChanges As GroupedDatabaseChange)
    Dim sql As String
    sql = "INSERT INTO metaTracks (ValidFrom, ValidUntil, CommitFK) VALUES (#" & Format(Now(), "yyyy/mm/dd") & "#, #9999/12/31#, " & groupChanges.CommitFK & ");"
    db.Execute sql
    
    Dim TrackFK As Long
    Dim rs As ADODB.Recordset
    Set rs = db.Execute("SELECT @@Identity")
    TrackFK = rs.fields(0).Value
    Debug.Print "Generated TrackFK = "; TrackFK
    groupChanges.TrackFK = TrackFK
End Sub
