VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FieldChanges"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Objects"
Option Explicit

Private Type TFieldChanges
    Working As ListObject
    Before As Worksheet
    Items As Collection
    BeforeArray As Variant
End Type

Private this As TFieldChanges

'@Enumerator
Public Property Get Items() As Collection
Attribute Items.VB_UserMemId = -4
    Set Items = this.Items
End Property

Public Property Get Working() As ListObject
    Set Working = this.Working
End Property

Public Property Set Working(ByVal RHS As ListObject)
    Set this.Working = RHS
End Property

Public Property Get Before() As Worksheet
    Set Before = this.Before
End Property

Public Property Set Before(ByVal RHS As Worksheet)
    If RHS Is Nothing Then Exit Property
    Set this.Before = RHS
    this.BeforeArray = RHS.UsedRange.Value2
End Property

Private Sub Class_Initialize()
    Set this.Items = New Collection
End Sub

Public Function IsValid() As Boolean
    If Working Is Nothing Then Exit Function
    If Before Is Nothing Then Exit Function
    IsValid = True
End Function

' ---
Public Sub Compare()
    CollectionClear this.Items
    
    Dim lc As ListColumn
    For Each lc In this.Working.ListColumns
        If lc.DataBodyRange.HasFormula = False And lc.Index <> 1 Then ' TODO Allow key columns in other column index
            CompareFieldsInColumn lc
        End If
    Next lc
End Sub

Private Sub CompareFieldsInColumn(ByVal lc As ListColumn) 'as Variant
    Dim AfterArray As Variant
    
    AfterArray = lc.DataBodyRange.Value2
    
    Dim i As Long
    For i = 1 To UBound(AfterArray, 1)
        CompareField AfterArray, i, lc.Index
    Next i
End Sub

Private Sub CompareField(ByRef AfterArray As Variant, ByVal Row As Long, ByVal Column As Long)
    Dim newColor As Long
    Dim Before As Variant
    Dim After As Variant

    'If Not IsValid Then MsgBox "FieldChanges not valid"
    
    newColor = xlNone
    Before = this.BeforeArray(Row + 1, Column)
    After = AfterArray(Row, 1)
    
    If IsError(Before) Or IsError(After) Then
        newColor = RGB(255, 255, 64)
    ElseIf Before <> After Then
        With New FieldChange
            .key = this.BeforeArray(Row + 1, 1)
            .ColumnName = this.BeforeArray(1, Column)
            .Before = Before
            .After = After
            ' TODO Slow, figure out a better method. Check VarType(After) first, at least
            If InStr(this.Working.DataBodyRange.Cells(Row, Column).NumberFormat, "y") > 0 Then
                .After = CDate(After)
            End If
            this.Items.Add .Self
        End With
        newColor = RGB(64, 255, 128)
    End If
    
    this.Working.DataBodyRange.Cells(Row, Column).Interior.Color = newColor
End Sub

Public Sub PrintChanges()
    Dim chg As FieldChange
    For Each chg In this.Items
        Debug.Print chg.ToString
    Next chg
End Sub

Public Sub GenerateTestFieldChanges()
    CollectionClear this.Items
    With this.Items
        .Add FieldChange.Create("JKL", "FieldA", CStr("sit"), CStr("abc"))
        .Add FieldChange.Create("JKL", "FieldB", CDbl("44652"), CDate("2022/01/31"))
        .Add FieldChange.Create("ABC", "FieldC", CDbl("1"), CDbl("52"))
    End With
End Sub
