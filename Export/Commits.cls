VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Commits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Objects")
Option Explicit

Private CommitKeys As Scripting.Dictionary

Public Function Self() As Commits
    Set Self = Me
End Function

Public Sub Load(ByVal dbChanges As DatabaseChanges)
    GetKeysFromAccess dbChanges.DistinctKeys
    Set CommitKeys = dbChanges.DistinctKeys
End Sub

Public Sub Apply(ByVal groupedChanges2 As GroupedDatabaseChanges)
    groupedChanges2.UpdateCommitKeys CommitKeys
End Sub

Private Sub GetKeysFromAccess(ByVal dict As Scripting.Dictionary)
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
