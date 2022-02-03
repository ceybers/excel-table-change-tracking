VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TrackChanges"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "TrackChanges"
Option Explicit

Private Const READ_WRITE_COLOR As Long = 33023 'RGB(255, 128, 0)

Private Type TTrackChanges
    WorkingTable As ListObject
    BeforeWorksheet As Worksheet
    BeforeArray As Variant
    Changes As FieldChanges
End Type

Private This As TTrackChanges

' Properties
Public Property Get HasChanges() As Boolean
    If This.Changes Is Nothing Then Exit Property
    HasChanges = This.Changes.Items.Count > 0
End Property

Public Property Get WorkingTable() As ListObject
    Set WorkingTable = This.WorkingTable
End Property

Public Property Set WorkingTable(ByVal RHS As ListObject)
    Set This.WorkingTable = RHS
End Property

Public Property Get BeforeWorksheet() As Worksheet
    Set BeforeWorksheet = This.BeforeWorksheet
End Property

Public Property Set BeforeWorksheet(ByVal RHS As Worksheet)
    Set This.BeforeWorksheet = RHS
    This.BeforeArray = RHS.UsedRange.Value2
End Property

Public Property Get BeforeArray() As Variant
    BeforeArray = This.BeforeArray
End Property

Public Property Get Changes() As FieldChanges
    Set Changes = This.Changes
End Property

' Public Methods
Public Sub StartTracking()
    ResetTracking
    
    Dim arr As Variant
    arr = This.WorkingTable.Range.Value2
    
    Dim rng As Range
    Set rng = This.BeforeWorksheet.Cells(1, 1)
    Set rng = rng.Resize(UBound(arr, 1), UBound(arr, 2))
    rng.Value2 = arr
    
    This.WorkingTable.HeaderRowRange.Interior.Color = READ_WRITE_COLOR
    
    LockFields
End Sub

Public Sub CalculateChanges()
    Debug.Assert Not This.BeforeWorksheet Is Nothing
    Debug.Assert Not This.WorkingTable Is Nothing

    This.BeforeArray = This.BeforeWorksheet.UsedRange.Value2
    
    If CompareHeadings And CompareKeys Then
        Set This.Changes = New FieldChanges
        With This.Changes
            Set .Working = This.WorkingTable
            Set .Before = This.BeforeWorksheet
            .Compare
        End With
    End If
    
    'ResetTracking
End Sub

Public Sub ResetTracking()
    This.BeforeWorksheet.Cells.Clear
    This.WorkingTable.Range.Interior.Color = xlNone
    This.WorkingTable.HeaderRowRange.Interior.Color = xlNone
    
    UnlockFields
End Sub

' Private Methods
Private Function CompareKeys() As Boolean
    CompareKeys = False
    
    If IsEmpty(This.BeforeArray) Then
        'MsgBox "No tracking history found!", vbExclamation + vbOKOnly
        Exit Function
    End If
    
    Dim rowCount As Long
    rowCount = UBound(BeforeArray, 1)
    If (This.WorkingTable.ListRows.Count + 1) <> UBound(This.BeforeArray, 1) Then Exit Function

    Dim i As Long
    For i = 2 To rowCount
        If This.WorkingTable.DataBodyRange.Cells(i - 1, 1).Value2 <> This.BeforeArray(i, 1) Then
            Exit Function
        End If
    Next i
    
    CompareKeys = True
End Function

' Private Methods
Private Function CompareHeadings() As Boolean
    'Debug.Assert IsEmpty(This.BeforeArray) = False
    If IsEmpty(This.BeforeArray) Then
        'MsgBox "No tracking history found!", vbExclamation + vbOKOnly
        Exit Function
    End If
    
    CompareHeadings = False
    
    Dim columnCount As Long
    columnCount = UBound(This.BeforeArray, 2)
    If This.WorkingTable.ListColumns.Count <> UBound(This.BeforeArray, 2) Then Exit Function

    Dim i As Long
    For i = 1 To columnCount
        If WorkingTable.HeaderRowRange.Cells(1, i).Value2 <> This.BeforeArray(1, i) Then
            Exit Function
        End If
    Next i
    
    CompareHeadings = True
End Function

Private Sub UnlockFields()
    Dim ws As Worksheet
    Set ws = This.WorkingTable.Parent
    ws.Unprotect
End Sub

Private Sub LockFields()
    Dim ws As Worksheet
    Set ws = This.WorkingTable.Parent
    ws.Unprotect
    
    Dim rng As Range
    Set rng = This.WorkingTable.DataBodyRange
    If rng.Columns.Count = 1 Then Exit Sub
    
    This.WorkingTable.Range.Locked = True
    
    Set rng = rng.Offset(0, 1).Resize(rng.Rows.Count, rng.Columns.Count - 1)
    rng.Locked = False
    
    ws.Protect AllowFiltering:=True, UserInterfaceOnly:=True
End Sub
