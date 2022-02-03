VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Commits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Commits"
Option Explicit

Private CommitKeys As Scripting.Dictionary

' Public Methods
Public Function Self() As Commits
    Set Self = Me
End Function

Public Sub CreatePerKey(ByVal dbChanges As DatabaseChanges)
    GetKeysFromAccess dbChanges.DistinctKeys
    Set CommitKeys = dbChanges.DistinctKeys
End Sub

Public Sub CreatePerSession(ByVal dbChanges As DatabaseChanges)
    GetKeysFromAccessSingle dbChanges.DistinctKeys
    Set CommitKeys = dbChanges.DistinctKeys
End Sub

Public Sub Apply(ByVal groupedChanges2 As GroupedDatabaseChanges)
    groupedChanges2.UpdateCommitKeys CommitKeys
End Sub

' Private Methods
Private Sub GetKeysFromAccess(ByVal dict As Scripting.Dictionary)
    Dim titlePrefix As String
    Dim title As String
    titlePrefix = "Untitled Commit @ " & Format$(Now(), "yyyy/mm/dd hh:MM") & " for "

    Dim Key As Variant
    For Each Key In dict.Keys
        title = titlePrefix & CStr(Key)
        dict(Key) = CreateCommitRecord(Access.GetConnection, title, "PerKey")
    Next Key
End Sub

Private Sub GetKeysFromAccessSingle(ByVal dict As Scripting.Dictionary)
    Dim title As String
    title = "Untitled Commit @ " & Format$(Now(), "yyyy/mm/dd hh:MM")
    
    Dim trackFK As Long
    trackFK = CreateCommitRecord(Access.GetConnection, title, "PerSession")
    
    Dim Key As Variant
    For Each Key In dict.Keys
        dict(Key) = trackFK
    Next Key
End Sub

Private Function CreateCommitRecord(ByVal db As ADODB.Connection, ByVal title As String, ByVal strategy As String) As Long
    Dim sql As String
    sql = "INSERT INTO " & COMMITS_TABLE_NAME & " ([Title], [Strategy]) VALUES ('" & title & "', '" & strategy & "');"
    db.Execute sql
    
    Dim rs2 As ADODB.Recordset
    Set rs2 = db.Execute("SELECT @@Identity")
    
    CreateCommitRecord = rs2.fields(0).Value
End Function
