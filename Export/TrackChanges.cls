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

Private this As TTrackChanges

' Properties
Public Property Get HasChanges() As Boolean
    If this.Changes Is Nothing Then Exit Property
    HasChanges = this.Changes.Items.Count > 0
End Property

Public Property Get WorkingTable() As ListObject
    Set WorkingTable = this.WorkingTable
End Property

Public Property Set WorkingTable(ByVal RHS As ListObject)
    Set this.WorkingTable = RHS
End Property

Public Property Get BeforeWorksheet() As Worksheet
    Set BeforeWorksheet = this.BeforeWorksheet
End Property

Public Property Set BeforeWorksheet(ByVal RHS As Worksheet)
    Set this.BeforeWorksheet = RHS
    this.BeforeArray = RHS.UsedRange.Value2
End Property

Public Property Get BeforeArray() As Variant
    BeforeArray = this.BeforeArray
End Property

Public Property Get Changes() As FieldChanges
    Set Changes = this.Changes
End Property

' Public Methods
Public Sub StartTracking()
    Debug.Assert Not this.WorkingTable Is Nothing
    'ResetTracking
    
    Dim arr As Variant
    arr = this.WorkingTable.Range.Value2
    
    Dim rng As Range
    Set rng = this.BeforeWorksheet.Cells(1, 1)
    Set rng = rng.Resize(UBound(arr, 1), UBound(arr, 2))
    
    SetNumberFormatInBeforeWorksheet
    
    rng.Value2 = arr
    
    HighlightHeaders True 'this.WorkingTable.HeaderRowRange.Interior.Color = READ_WRITE_COLOR
    
    LockFields
End Sub

Private Sub SetNumberFormatInBeforeWorksheet()
    Dim i As Long
    For i = 1 To this.WorkingTable.ListColumns.Count
        BeforeWorksheet.Columns(i).NumberFormat = this.WorkingTable.ListColumns(i).DataBodyRange.Cells(1, 1).NumberFormat
    Next i
End Sub

Public Sub CalculateChanges()
    Debug.Assert Not this.BeforeWorksheet Is Nothing
    Debug.Assert Not this.WorkingTable Is Nothing

    this.BeforeArray = this.BeforeWorksheet.UsedRange.Value2
    
    If CompareHeadings And CompareKeys Then
        Set this.Changes = New FieldChanges
        With this.Changes
            Set .Working = this.WorkingTable
            Set .Before = this.BeforeWorksheet
            .Compare
        End With
    End If
    
    'ResetTracking
End Sub

Public Sub ResetTracking()
    this.BeforeWorksheet.Cells.Clear
    'this.WorkingTable.Range.Interior.Color = xlNone
    
    UnlockFields
    
    'this.WorkingTable.HeaderRowRange.Interior.Color = xlNone
    HighlightHeaders False
    
    Set this.BeforeWorksheet = Nothing
    Set this.WorkingTable = Nothing
End Sub

' Private Methods
' TODO Refactor this
Private Sub HighlightHeaders(ByVal Enable As Boolean)
    Dim schema As Collection
    Set schema = New Collection
    LoadSchema schema
    
    Dim headerRange As Range
    Dim unionRange As Range
    Dim intersectRange As Range
    
    Set headerRange = this.WorkingTable.HeaderRowRange
    
    Dim lc As ListColumn
    For Each lc In this.WorkingTable.ListColumns
        Dim si As SchemaItem
        For Each si In schema
            If si.ColumnName = lc.Name Then
                Set intersectRange = Application.Intersect(lc.Range, headerRange)
                If Not intersectRange Is Nothing Then
                    If Not intersectRange.Comment Is Nothing Then
                        intersectRange.Comment.Delete
                    End If
                    intersectRange.AddComment si.ToString
                    AppendRange intersectRange, unionRange
                End If
            End If
        Next si
    Next lc
    
    Dim fc As FormatCondition
    If Enable Then
        Set fc = unionRange.FormatConditions.Add(Type:=xlExpression, Formula1:="=TRUE")
        With fc
            .SetFirstPriority
            .Interior.Color = READ_WRITE_COLOR
            .Font.Color = RGB(255, 255, 255)
        End With
    Else
        ' TODO FIX This will crash Excel as the collection changes when deleting FCs.
        For Each fc In this.WorkingTable.HeaderRowRange.FormatConditions
            If fc.Formula1 = "=TRUE" Then
                fc.Delete
            End If
        Next fc
        For Each lc In this.WorkingTable.ListColumns
            If Not lc.Range.Cells(1, 1).Comment Is Nothing Then
                lc.Range.Cells(1, 1).Comment.Delete
            End If
        Next lc
    End If
End Sub

Private Function CompareKeys() As Boolean
    CompareKeys = False
    
    If IsEmpty(this.BeforeArray) Then
        'MsgBox "No tracking history found!", vbExclamation + vbOKOnly
        Exit Function
    End If
    
    Dim rowCount As Long
    rowCount = UBound(BeforeArray, 1)
    If (this.WorkingTable.ListRows.Count + 1) <> UBound(this.BeforeArray, 1) Then Exit Function

    Dim i As Long
    For i = 2 To rowCount
        If this.WorkingTable.DataBodyRange.Cells(i - 1, 1).Value2 <> this.BeforeArray(i, 1) Then
            Exit Function
        End If
    Next i
    
    CompareKeys = True
End Function

' Private Methods
Private Function CompareHeadings() As Boolean
    'Debug.Assert IsEmpty(This.BeforeArray) = False
    If IsEmpty(this.BeforeArray) Then
        'MsgBox "No tracking history found!", vbExclamation + vbOKOnly
        Exit Function
    End If
    
    CompareHeadings = False
    
    Dim columnCount As Long
    columnCount = UBound(this.BeforeArray, 2)
    If this.WorkingTable.ListColumns.Count <> UBound(this.BeforeArray, 2) Then Exit Function

    Dim i As Long
    For i = 1 To columnCount
        If WorkingTable.HeaderRowRange.Cells(1, i).Value2 <> this.BeforeArray(1, i) Then
            Exit Function
        End If
    Next i
    
    CompareHeadings = True
End Function

Private Sub UnlockFields()
    Dim ws As Worksheet
    Set ws = this.WorkingTable.Parent
    ws.Unprotect
End Sub

Private Sub LockFields()
    Dim ws As Worksheet
    Set ws = this.WorkingTable.Parent
    ws.Unprotect
    
    Dim rng As Range
    Set rng = this.WorkingTable.DataBodyRange
    If rng.Columns.Count = 1 Then Exit Sub
    
    this.WorkingTable.Range.Locked = True
    
    Set rng = rng.Offset(0, 1).Resize(rng.rows.Count, rng.Columns.Count - 1)
    rng.Locked = False
    
    ws.Protect AllowFiltering:=True, UserInterfaceOnly:=True
End Sub
