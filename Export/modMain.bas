Attribute VB_Name = "modMain"
'@Folder("VBAProject")
Option Explicit

Private Const BEFORE_SHEET_SUFFIX As String = "_CHGTRK"

' Public Methods
Public Sub Start()
    Dim tc As TrackChanges
    Set tc = GetTrackChangesObject
    If Not tc Is Nothing Then
        tc.StartTracking
    End If
    
    Exit Sub
End Sub

Public Sub HighlightChanges()
    Dim tc As TrackChanges
    Set tc = GetTrackChangesObject
    
    tc.CalculateChanges
End Sub

Public Sub Save()
    Dim tc As TrackChanges
    Set tc = GetTrackChangesObject
    
    HighlightChanges
    
    If tc.HasChanges = False Then
        MsgBox "No changes found!", vbInformation
        Exit Sub
    End If
    
    If vbNo = MsgBox("Update Access DB with " & tc.Changes.Items.Count & " change(s)?", vbInformation + vbYesNo + vbDefaultButton1) Then
        Exit Sub
    End If
    
    tc.Changes.Compare
    
    Dim dbChanges As DatabaseChanges
    Set dbChanges = DatabaseChanges.Create(New DatabaseChangeFactory, tc.Changes)
    
    Dim grpDbChanges As GroupedDatabaseChanges
    Set grpDbChanges = GroupedDatabaseChanges.Create(dbChanges)
    
    With New Commits
        If tc.Changes.Items.Count > 1 Then
            If vbYes = MsgBox("Create individual commits for each key?", vbInformation + vbYesNo) Then
                .CreatePerKey dbChanges
            Else
                .CreatePerSession dbChanges
            End If
        Else
            .CreatePerKey dbChanges
        End If
        
        .Apply grpDbChanges
    End With
    
    SaveGroupedChangesToAccess grpDbChanges
    
    Access.GetConnection DoClose:=True
    
    tc.ResetTracking
End Sub

Public Sub Reset()
    GetTrackChangesObject.ResetTracking
End Sub

' Private Methods
Private Function GetTrackChangesObject() As TrackChanges
    Static tc As TrackChanges
    
    If ActiveSheet.ListObjects.Count <> 1 Then
        MsgBox "Cannot find a table to track!", vbCritical + vbOKOnly
        Exit Function
    End If
    
    Dim lo As ListObject
    Set lo = ActiveSheet.ListObjects(1)
    
    Dim ws As Worksheet
    Set ws = AddOrGetWorksheet(lo & BEFORE_SHEET_SUFFIX)
    
    If tc Is Nothing Then
        Set tc = New TrackChanges
    End If
    
    If tc.WorkingTable Is Nothing Then
        Set tc.BeforeWorksheet = ws
        Set tc.WorkingTable = lo
    End If
    
    Set GetTrackChangesObject = tc
End Function
