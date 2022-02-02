Attribute VB_Name = "modMain"
'@Folder("VBAProject")
Option Explicit

Private Const BEFORE_SHEET_NAME As String = "Before"

Private WorkingTable As ListObject
Private BeforeWorksheet As Worksheet
Private BeforeArray As Variant
Private Changes As FieldChanges

Public Sub Start()
    Set WorkingTable = ThisWorkbook.Worksheets(1).ListObjects(1)
    
    SaveToBefore
    
    WorkingTable.Range.Interior.Color = xlNone
    LockFields
End Sub

Public Sub Save()
    Set WorkingTable = ThisWorkbook.Worksheets(1).ListObjects(1)
    BeforeArray = LoadFromBefore
    
    UnlockFields
    
    If CompareHeadings And CompareKeys Then
        WorkingTable.Range.Interior.Color = xlNone
        
        Set Changes = New FieldChanges
        Set Changes.Working = WorkingTable
        Set Changes.Before = ThisWorkbook.Worksheets(BEFORE_SHEET_NAME)
        Changes.Compare
        
        Dim dbChanges As DatabaseChanges
        Set dbChanges = DatabaseChanges.Create(New DatabaseChangeFactory, Changes)
        
        Dim groupedChanges2 As GroupedDatabaseChanges
        Set groupedChanges2 = GroupedDatabaseChanges.Create(dbChanges)
        
        With New Commits
            .Load dbChanges
            .Apply groupedChanges2
        End With
        
        SaveGroupedChangesToAccess groupedChanges2
    End If
End Sub

Private Sub UnlockFields()
    WorkingTable.Parent.Unprotect
End Sub

Private Sub LockFields()
    Dim rng As Range
    Set rng = WorkingTable.DataBodyRange
    If rng.Columns.Count = 1 Then Exit Sub
    
    WorkingTable.Range.Locked = True
    
    Set rng = rng.Offset(0, 1).Resize(rng.Rows.Count, rng.Columns.Count - 1)
    rng.Locked = False
    
    WorkingTable.Parent.Protect AllowFiltering:=True, UserInterfaceOnly:=True
End Sub

Private Function CompareKeys() As Boolean
    Dim After As Variant
    Dim rowCount As Long
    
    CompareKeys = False
    rowCount = UBound(BeforeArray, 1)
    If (WorkingTable.ListRows.Count + 1) <> UBound(BeforeArray, 1) Then Exit Function

    Dim i As Long
    For i = 2 To rowCount
        If WorkingTable.DataBodyRange.Cells(i - 1, 1).Value2 <> BeforeArray(i, 1) Then
            Exit Function
        End If
    Next i
    
    CompareKeys = True
End Function

Private Function CompareHeadings()
    Dim After As Variant
    Dim columnCount As Long
    
    CompareHeadings = False
    columnCount = UBound(BeforeArray, 2)
    If WorkingTable.ListColumns.Count <> UBound(BeforeArray, 2) Then Exit Function

    Dim i As Long
    For i = 1 To columnCount
        If WorkingTable.HeaderRowRange.Cells(1, i).Value2 <> BeforeArray(1, i) Then
            Exit Function
        End If
    Next i
    
    CompareHeadings = True
End Function

Private Function LoadFromBefore() As Variant
    Dim ws As Worksheet
    Set ws = AddOrGetWorksheet(BEFORE_SHEET_NAME)
    If ws.UsedRange.Cells.Count = 0 Then Exit Function
    LoadFromBefore = ws.UsedRange.Value2
End Function

Private Sub SaveToBefore()
    Set BeforeWorksheet = AddOrGetWorksheet(BEFORE_SHEET_NAME)
    
    BeforeWorksheet.Cells.Clear
    
    Dim arr As Variant
    arr = WorkingTable.Range.Value2
    
    Dim rng As Range
    Set rng = BeforeWorksheet.Cells(1, 1)
    Set rng = rng.Resize(UBound(arr, 1), UBound(arr, 2))
    rng.Value2 = arr
End Sub
