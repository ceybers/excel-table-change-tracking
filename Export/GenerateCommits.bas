Attribute VB_Name = "GenerateCommits"
'@Folder("VBAProject")
Option Explicit

Private Const TRACKS_TABLE_NAME As String = "metaTracks"
Private Const COMMITS_TABLE_NAME As String = "metaCommits"

Enum CommitStrategy
    PerCommit ' One new Track for everything
    PerKey
    PerKeyAndTable
    'NYIPerColumn ' As below
    'NYIPerField ' This breaks DatabaseChange since we are trying to group INSERTs by Key and not by Field
End Enum

Public Function GenerateCommitsPerKey(ByVal dbChanges As DatabaseChanges) As Scripting.Dictionary
    GetKeysFromAccess dbChanges.DistinctKeys

    Set GenerateCommitsPerKey = dbChanges.DistinctKeys
End Function

Private Sub ZZZ_ChooseStrategy(ByVal strategy As CommitStrategy)
    Select Case strategy
        Case PerCommit ' One new Track for everything
            ' as below, even if only 1 item
        Case PerKey
            ' Pass a dictionary to GetKeysFromAccess with the titles we want; GetKeysFromAccess can place the TracKFK in the .value
        Case PerKeyAndTable
            ' NYI
    End Select
End Sub

Public Sub GetKeysFromAccess(ByVal dict As Scripting.Dictionary)
    Dim Connection As ADODB.Connection
    Set Connection = New ADODB.Connection
    Connection.Open ConnectionString:=Access.GetConnectionString
   
    Dim titlePrefix As String
    Dim title As String
    titlePrefix = "Untitled Commit @ " & Format(Now(), "yyyy/mm/dd hh:MM") & " for "

    Dim key As Variant
    For Each key In dict.Keys
        title = titlePrefix & CStr(key)
        dict(key) = CreateCommitRecord(Connection, title, "PerKey")
    Next key
    
    Connection.Close
    Set Connection = Nothing
End Sub

Private Function CreateCommitRecord(ByVal db As ADODB.Connection, ByVal title As String, ByVal strategy As String) As Long
    Dim sql As String
    sql = "INSERT INTO " & COMMITS_TABLE_NAME & " ([Title], [Strategy]) VALUES ('" & title & "', '" & strategy & "');"
    db.Execute sql
    
    Dim rs2 As ADODB.Recordset
    Set rs2 = db.Execute("SELECT @@Identity")
    
    CreateCommitRecord = rs2.fields(0).Value
End Function
